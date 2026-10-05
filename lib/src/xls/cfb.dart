// Minimal reader for the OLE2 / Compound File Binary (CFB) container that
// wraps legacy `.xls` workbooks. See [MS-CFB].
//
// Only what is needed to extract a named stream is implemented. All reads are
// bounds checked and every chain walk is cycle-safe, so malformed or hostile
// input ends in a [FormatException] instead of a hang or an out-of-memory.
part of spreadsheet_decoder;

const List<int> _cfbSignature = [
  0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 //
];

/// Returns true if [data] starts with the CFB (OLE2) signature.
bool _hasCfbSignature(List<int> data) {
  if (data.length < _cfbSignature.length) {
    return false;
  }
  for (var i = 0; i < _cfbSignature.length; i++) {
    if (data[i] != _cfbSignature[i]) {
      return false;
    }
  }
  return true;
}

const int _cfbFreeSect = 0xFFFFFFFF;
const int _cfbEndOfChain = 0xFFFFFFFE;
const int _cfbMaxRegSect = 0xFFFFFFFA;
const int _cfbNoStream = 0xFFFFFFFF;
const int _cfbDirEntrySize = 128;
const int _cfbHeaderSize = 512;

const int _cfbObjectStream = 2;
const int _cfbObjectRoot = 5;

class _CfbEntry {
  final String name;
  final int type;
  final int left;
  final int right;
  final int child;
  final int startSector;
  final int size;

  _CfbEntry(this.name, this.type, this.left, this.right, this.child,
      this.startSector, this.size);
}

/// Read-only view of a CFB container held in memory.
class _CfbReader {
  final Uint8List _bytes;
  final ByteData _data;

  late final int _sectorSize;
  late final int _miniSectorSize;
  late final int _miniCutoff;
  late final int _sectorCount;

  late final Uint32List _fat;
  Uint32List? _miniFat;
  late final List<_CfbEntry> _entries;
  Uint8List? _miniStream;

  _CfbReader(List<int> bytes)
      : _bytes = bytes is Uint8List ? bytes : Uint8List.fromList(bytes),
        _data = ByteData.sublistView(
            bytes is Uint8List ? bytes : Uint8List.fromList(bytes)) {
    if (!_hasCfbSignature(_bytes) || _bytes.length < _cfbHeaderSize) {
      throw FormatException('Not a compound file (OLE2) document');
    }

    var sectorShift = _data.getUint16(30, Endian.little);
    var miniShift = _data.getUint16(32, Endian.little);
    // Only 512 byte (v3) and 4096 byte (v4) sectors exist in the wild.
    if ((sectorShift != 9 && sectorShift != 12) || miniShift != 6) {
      throw FormatException('Unsupported compound file sector size');
    }
    _sectorSize = 1 << sectorShift;
    _miniSectorSize = 1 << miniShift;
    _miniCutoff = _data.getUint32(56, Endian.little);
    // The header occupies the first sector (512 or 4096 bytes). A trailing
    // partial sector is tolerated, some writers do not pad the last one.
    _sectorCount =
        (_bytes.length - _sectorSize + _sectorSize - 1) ~/ _sectorSize;

    _readFat();
    _readDirectory();
  }

  /// Names of the streams stored directly in the root storage.
  Iterable<String> get rootStreamNames => _rootStreams().keys;

  /// Returns true when the root storage contains the stream [name].
  bool hasStream(String name) => _rootStreams().containsKey(name);

  /// Returns the content of the root level stream called [name], or `null`
  /// when it does not exist.
  Uint8List? readStream(String name) {
    var entry = _rootStreams()[name];
    if (entry == null) {
      return null;
    }
    if (entry.size < _miniCutoff) {
      return _readMiniChain(entry.startSector, entry.size);
    }
    return _readChain(entry.startSector, entry.size);
  }

  Map<String, _CfbEntry>? _rootStreamCache;

  Map<String, _CfbEntry> _rootStreams() {
    var cached = _rootStreamCache;
    if (cached != null) {
      return cached;
    }
    var streams = <String, _CfbEntry>{};
    var root = _entries.first;
    // Children of a storage form a red-black tree: walk it iteratively, with
    // a visited set, since the tree can be corrupt and contain cycles.
    var visited = <int>{};
    var pending = <int>[root.child];
    while (pending.isNotEmpty) {
      var id = pending.removeLast();
      if (id == _cfbNoStream || id >= _entries.length || !visited.add(id)) {
        continue;
      }
      var entry = _entries[id];
      if (entry.type == _cfbObjectStream) {
        streams.putIfAbsent(entry.name, () => entry);
      }
      pending
        ..add(entry.left)
        ..add(entry.right);
    }
    return _rootStreamCache = streams;
  }

  int _sectorOffset(int sector) => (sector + 1) * _sectorSize;

  bool _validSector(int sector) =>
      sector <= _cfbMaxRegSect && sector < _sectorCount;

  Uint8List _sectorView(int sector) {
    var start = _sectorOffset(sector);
    var end = start + _sectorSize;
    if (end > _bytes.length) {
      end = _bytes.length;
    }
    return Uint8List.sublistView(_bytes, start, end);
  }

  void _readFat() {
    var fatSectorCount = _data.getUint32(44, Endian.little);
    var difatStart = _data.getUint32(68, Endian.little);
    var difatCount = _data.getUint32(72, Endian.little);
    if (fatSectorCount > _sectorCount) {
      throw FormatException('Corrupt compound file: invalid FAT size');
    }

    // Collect the FAT sector numbers: 109 in the header, the rest in the
    // DIFAT chain.
    var fatSectors = <int>[];
    for (var i = 0; i < 109 && fatSectors.length < fatSectorCount; i++) {
      var sector = _data.getUint32(76 + i * 4, Endian.little);
      if (sector == _cfbFreeSect) {
        continue;
      }
      fatSectors.add(sector);
    }
    var difat = difatStart;
    var guard = 0;
    while (fatSectors.length < fatSectorCount &&
        difatCount > 0 &&
        _validSector(difat)) {
      if (++guard > difatCount || guard > _sectorCount) {
        throw FormatException('Corrupt compound file: DIFAT loop');
      }
      var view = ByteData.sublistView(_sectorView(difat));
      var entries = _sectorSize ~/ 4 - 1;
      for (var i = 0;
          i < entries &&
              fatSectors.length < fatSectorCount &&
              (i + 1) * 4 <= view.lengthInBytes;
          i++) {
        var sector = view.getUint32(i * 4, Endian.little);
        if (sector != _cfbFreeSect) {
          fatSectors.add(sector);
        }
      }
      difat = view.lengthInBytes >= _sectorSize
          ? view.getUint32(_sectorSize - 4, Endian.little)
          : _cfbEndOfChain;
    }

    var perSector = _sectorSize ~/ 4;
    var fat = Uint32List(fatSectors.length * perSector)
      ..fillRange(0, fatSectors.length * perSector, _cfbFreeSect);
    var count = 0;
    for (var sector in fatSectors) {
      if (!_validSector(sector)) {
        throw FormatException('Corrupt compound file: invalid FAT sector');
      }
      var view = ByteData.sublistView(_sectorView(sector));
      var n = view.lengthInBytes ~/ 4;
      for (var i = 0; i < n; i++) {
        fat[count + i] = view.getUint32(i * 4, Endian.little);
      }
      count += perSector;
    }
    _fat = fat;
  }

  /// Follows the FAT chain from [start]. Returns the sector numbers.
  List<int> _chain(Uint32List table, int start, int maxLength) {
    var sectors = <int>[];
    var sector = start;
    while (sector != _cfbEndOfChain) {
      if (sector > _cfbMaxRegSect || sector >= table.length) {
        throw FormatException('Corrupt compound file: invalid sector chain');
      }
      // Longer than the number of existing sectors can only mean a loop.
      if (sectors.length >= maxLength) {
        throw FormatException('Corrupt compound file: sector chain loop');
      }
      sectors.add(sector);
      sector = table[sector];
    }
    return sectors;
  }

  Uint8List _readChain(int start, int size) {
    if (size == 0) {
      return Uint8List(0);
    }
    var sectors = _chain(_fat, start, _sectorCount);
    if (sectors.length * _sectorSize < size) {
      throw FormatException('Corrupt compound file: stream is truncated');
    }
    var out = Uint8List(size);
    var written = 0;
    for (var sector in sectors) {
      if (written >= size) {
        break;
      }
      if (!_validSector(sector)) {
        throw FormatException('Corrupt compound file: sector out of range');
      }
      var view = _sectorView(sector);
      var n = size - written < view.length ? size - written : view.length;
      out.setRange(written, written + n, view);
      written += n;
    }
    if (written < size) {
      throw FormatException('Corrupt compound file: stream is truncated');
    }
    return out;
  }

  Uint8List _readMiniChain(int start, int size) {
    if (size == 0) {
      return Uint8List(0);
    }
    var miniFat = _miniFat;
    var miniStream = _miniStream;
    if (miniFat == null || miniStream == null) {
      var root = _entries.first;
      var miniFatStart = _data.getUint32(60, Endian.little);
      var miniFatCount = _data.getUint32(64, Endian.little);
      if (miniFatCount > _sectorCount) {
        throw FormatException('Corrupt compound file: invalid mini FAT size');
      }
      var raw = miniFatStart == _cfbEndOfChain
          ? Uint8List(0)
          : _readChain(miniFatStart, miniFatCount * _sectorSize);
      var table = ByteData.sublistView(raw);
      miniFat = Uint32List(raw.length ~/ 4);
      for (var i = 0; i < miniFat.length; i++) {
        miniFat[i] = table.getUint32(i * 4, Endian.little);
      }
      miniStream = _readChain(root.startSector, root.size);
      _miniFat = miniFat;
      _miniStream = miniStream;
    }

    var maxSectors = miniStream.length ~/ _miniSectorSize;
    var sectors = _chain(miniFat, start, maxSectors);
    if (sectors.length * _miniSectorSize < size) {
      throw FormatException('Corrupt compound file: stream is truncated');
    }
    var out = Uint8List(size);
    var written = 0;
    for (var sector in sectors) {
      if (written >= size) {
        break;
      }
      var offset = sector * _miniSectorSize;
      if (offset + _miniSectorSize > miniStream.length) {
        throw FormatException('Corrupt compound file: mini sector range');
      }
      var n =
          size - written < _miniSectorSize ? size - written : _miniSectorSize;
      out.setRange(written, written + n,
          Uint8List.sublistView(miniStream, offset, offset + n));
      written += n;
    }
    return out;
  }

  void _readDirectory() {
    var dirStart = _data.getUint32(48, Endian.little);
    if (!_validSector(dirStart)) {
      throw FormatException('Corrupt compound file: invalid directory');
    }
    var sectors = _chain(_fat, dirStart, _sectorCount);
    var entries = <_CfbEntry>[];
    for (var sector in sectors) {
      if (!_validSector(sector)) {
        throw FormatException('Corrupt compound file: invalid directory');
      }
      var view = _sectorView(sector);
      if (view.length < _sectorSize) {
        throw FormatException('Corrupt compound file: directory is truncated');
      }
      var data = ByteData.sublistView(view);
      for (var offset = 0;
          offset + _cfbDirEntrySize <= view.length;
          offset += _cfbDirEntrySize) {
        var nameLength = data.getUint16(offset + 64, Endian.little);
        // Includes the UTF-16 terminator; 64 bytes is the field size.
        if (nameLength > 64) {
          nameLength = 64;
        }
        var units = <int>[];
        for (var i = 0; i + 1 < nameLength - 1; i += 2) {
          units.add(data.getUint16(offset + i, Endian.little));
        }
        // The size is 64 bits wide in version 4 files; sizes above 4 GiB
        // cannot be held in memory anyway.
        var sizeLow = data.getUint32(offset + 120, Endian.little);
        var sizeHigh = data.getUint32(offset + 124, Endian.little);
        entries.add(_CfbEntry(
          String.fromCharCodes(units),
          data.getUint8(offset + 66),
          data.getUint32(offset + 68, Endian.little),
          data.getUint32(offset + 72, Endian.little),
          data.getUint32(offset + 76, Endian.little),
          data.getUint32(offset + 116, Endian.little),
          sizeHigh != 0 ? 0x7FFFFFFF : sizeLow,
        ));
      }
    }
    if (entries.isEmpty || entries.first.type != _cfbObjectRoot) {
      throw FormatException('Corrupt compound file: missing root entry');
    }
    _entries = entries;
  }
}
