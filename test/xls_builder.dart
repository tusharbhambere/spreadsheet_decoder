// Helpers to assemble minimal `.xls` files (OLE2 container + BIFF records) so
// tests can exercise record types and corner cases that real Excel/xlwt
// generated fixtures don't contain.
import 'dart:typed_data';

List<int> le16(int v) => [v & 0xFF, (v >> 8) & 0xFF];
List<int> le32(int v) => [...le16(v & 0xFFFF), ...le16((v >> 16) & 0xFFFF)];
List<int> leF64(double v) =>
    (ByteData(8)..setFloat64(0, v, Endian.little)).buffer.asUint8List();

/// BIFF record: id, length, payload.
List<int> record(int id, List<int> data) =>
    [...le16(id), ...le16(data.length), ...data];

/// A record whose payload is split over [parts]: the first part is the record
/// itself, the others are CONTINUE records.
List<int> recordWithContinues(int id, List<List<int>> parts) => [
      ...record(id, parts.first),
      for (var part in parts.skip(1)) ...record(0x003C, part),
    ];

/// BIFF8 string with 8 bit characters when possible, UTF-16 otherwise.
List<int> xlString(String text, {bool short = false, bool? wide}) {
  wide ??= text.codeUnits.any((c) => c > 0xFF);
  return [
    ...(short ? [text.length] : le16(text.length)),
    wide ? 1 : 0,
    for (var c in text.codeUnits) ...(wide ? le16(c) : [c]),
  ];
}

/// BIFF5 string: length followed by single byte characters.
List<int> byteString(String text, {bool short = false}) => [
      ...(short ? [text.length] : le16(text.length)),
      ...text.codeUnits
    ];

// Workbook globals ----------------------------------------------------------

List<int> bof(int type, {bool biff8 = true}) => record(
    0x0809,
    [...le16(biff8 ? 0x0600 : 0x0500), ...le16(type), ...le32(0)] +
        (biff8 ? List.filled(8, 0) : const []));

List<int> eof() => record(0x000A, const []);

List<int> dateMode(int mode) => record(0x0022, le16(mode));

List<int> format(int id, String code, {bool biff8 = true}) => record(0x041E,
    [...le16(id), ...(biff8 ? xlString(code) : byteString(code, short: true))]);

/// Cell XF using number format [formatId].
List<int> xf(int formatId) =>
    record(0x00E0, [...le16(0), ...le16(formatId), ...List.filled(16, 0)]);

List<int> sst(List<String> strings) => record(0x00FC, [
      ...le32(strings.length),
      ...le32(strings.length),
      for (var s in strings) ...xlString(s),
    ]);

// Cells ---------------------------------------------------------------------

List<int> number(int row, int col, double value, {int xfIndex = 0}) => record(
    0x0203, [...le16(row), ...le16(col), ...le16(xfIndex), ...leF64(value)]);

List<int> rk(int row, int col, int rkValue, {int xfIndex = 0}) => record(
    0x027E, [...le16(row), ...le16(col), ...le16(xfIndex), ...le32(rkValue)]);

List<int> mulRk(int row, int firstCol, List<int> rkValues, {int xfIndex = 0}) =>
    record(0x00BD, [
      ...le16(row),
      ...le16(firstCol),
      for (var v in rkValues) ...[...le16(xfIndex), ...le32(v)],
      ...le16(firstCol + rkValues.length - 1),
    ]);

List<int> labelSst(int row, int col, int index, {int xfIndex = 0}) => record(
    0x00FD, [...le16(row), ...le16(col), ...le16(xfIndex), ...le32(index)]);

List<int> label(int row, int col, List<int> encodedText, {int xfIndex = 0}) =>
    record(
        0x0204, [...le16(row), ...le16(col), ...le16(xfIndex), ...encodedText]);

List<int> boolErr(int row, int col, int value, {bool error = false}) => record(
    0x0205, [...le16(row), ...le16(col), ...le16(0), value, error ? 1 : 0]);

List<int> _formula(int row, int col, List<int> result, int xfIndex) =>
    record(0x0006, [
      ...le16(row),
      ...le16(col),
      ...le16(xfIndex),
      ...result,
      ...le16(0), // flags
      ...le32(0), // chn
      ...le16(0), // cce: empty token array
    ]);

List<int> formulaNumber(int row, int col, double value, {int xfIndex = 0}) =>
    _formula(row, col, leF64(value), xfIndex);

/// Formula with a string result: FORMULA followed by its STRING record.
List<int> formulaString(int row, int col, List<int> encodedText,
        {List<int> between = const []}) =>
    [
      ..._formula(row, col, [0, 0, 0, 0, 0, 0, 0xFF, 0xFF], 0),
      ...between,
      ...record(0x0207, encodedText),
    ];

/// Like [formulaString] with a STRING record split over CONTINUE records.
List<int> formulaStringParts(int row, int col, List<List<int>> parts) => [
      ..._formula(row, col, [0, 0, 0, 0, 0, 0, 0xFF, 0xFF], 0),
      ...recordWithContinues(0x0207, parts),
    ];

List<int> formulaBool(int row, int col, bool value) =>
    _formula(row, col, [1, 0, value ? 1 : 0, 0, 0, 0, 0xFF, 0xFF], 0);

List<int> formulaError(int row, int col, int code) =>
    _formula(row, col, [2, 0, code, 0, 0, 0, 0xFF, 0xFF], 0);

List<int> formulaEmptyString(int row, int col) =>
    _formula(row, col, [3, 0, 0, 0, 0, 0, 0xFF, 0xFF], 0);

// Workbook stream -----------------------------------------------------------

class SheetSpec {
  final String name;
  final List<List<int>> records;
  final int type; // BOUNDSHEET type: 0 worksheet, 2 chart, ...

  SheetSpec(this.name, this.records, {this.type = 0});
}

/// Builds the `Workbook` stream: globals, one BOUNDSHEET per sheet, then the
/// sheet substreams.
Uint8List workbookStream(
  List<SheetSpec> sheets, {
  List<List<int>> globals = const [],
  bool biff8 = true,
}) {
  List<int> boundSheet(SheetSpec sheet, int offset) => record(0x0085, [
        ...le32(offset),
        0, // visible
        sheet.type,
        ...(biff8
            ? xlString(sheet.name, short: true)
            : byteString(sheet.name, short: true)),
      ]);

  var head = [
    ...bof(0x0005, biff8: biff8),
    for (var g in globals) ...g,
  ];
  var tail = eof();
  var boundSize = [
    for (var s in sheets) ...boundSheet(s, 0),
  ].length;

  var offset = head.length + boundSize + tail.length;
  var out = <int>[...head];
  var bodies = <List<int>>[];
  for (var sheet in sheets) {
    out.addAll(boundSheet(sheet, offset));
    var body = [
      ...bof(0x0010, biff8: biff8),
      for (var r in sheet.records) ...r,
      ...eof(),
    ];
    bodies.add(body);
    offset += body.length;
  }
  out.addAll(tail);
  for (var body in bodies) {
    out.addAll(body);
  }
  return Uint8List.fromList(out);
}

// Compound file -------------------------------------------------------------

const _free = 0xFFFFFFFF;
const _endOfChain = 0xFFFFFFFE;
const _fatSector = 0xFFFFFFFD;
const _difatSector = 0xFFFFFFFC;

/// Wraps [streams] (name -> content) in a version 3 (512 byte sectors) OLE2
/// compound file. Streams under 4096 bytes go to the mini stream like in real
/// files. The FAT/DIFAT layout grows as needed, so huge payloads exercise the
/// DIFAT chain.
Uint8List compoundFile(Map<String, List<int>> streams, {int sectorSize = 512}) {
  var perSector = sectorSize ~/ 4;
  var small = <String, List<int>>{};
  var large = <String, List<int>>{};
  streams.forEach((name, data) {
    (data.length < 4096 ? small : large)[name] = data;
  });

  // Mini stream: every small stream padded to 64 byte mini sectors.
  var miniStream = <int>[];
  var miniFat = <int>[];
  var miniStarts = <String, int>{};
  small.forEach((name, data) {
    if (data.isEmpty) {
      miniStarts[name] = _endOfChain;
      return;
    }
    var count = (data.length + 63) ~/ 64;
    miniStarts[name] = miniFat.length;
    for (var i = 0; i < count; i++) {
      miniFat.add(i == count - 1 ? _endOfChain : miniFat.length + 1);
    }
    miniStream
      ..addAll(data)
      ..addAll(List.filled(count * 64 - data.length, 0));
  });

  int sectorsFor(int bytes) => (bytes + sectorSize - 1) ~/ sectorSize;

  var names = ['Root Entry', ...streams.keys];
  var dirSectors = sectorsFor(names.length * 128);
  var miniFatSectors = sectorsFor(miniFat.length * 4);
  var miniStreamSectors = sectorsFor(miniStream.length);
  var largeSectors = {
    for (var e in large.entries) e.key: sectorsFor(e.value.length)
  };
  var payload = dirSectors +
      miniFatSectors +
      miniStreamSectors +
      largeSectors.values.fold<int>(0, (a, b) => a + b);

  // Number of FAT sectors f and DIFAT sectors x must cover themselves.
  var fatCount = 1;
  var difatCount = 0;
  while (true) {
    var needed = sectorsFor((payload + fatCount + difatCount) * 4);
    var difat = fatCount > 109
        ? (fatCount - 109 + (perSector - 1) - 1) ~/ (perSector - 1)
        : 0;
    if (needed <= fatCount && difat == difatCount) break;
    fatCount = needed > fatCount ? needed : fatCount;
    difatCount = fatCount > 109
        ? (fatCount - 109 + (perSector - 1) - 1) ~/ (perSector - 1)
        : 0;
  }

  var total = payload + fatCount + difatCount;
  var fat = List<int>.filled(fatCount * perSector, _free);
  var cursor = 0;
  var fatStart = cursor;
  for (var i = 0; i < fatCount; i++) {
    fat[cursor++] = _fatSector;
  }
  var difatStart = cursor;
  for (var i = 0; i < difatCount; i++) {
    fat[cursor++] = _difatSector;
  }

  int chain(int sectors) {
    if (sectors == 0) return _endOfChain;
    var start = cursor;
    for (var i = 0; i < sectors; i++) {
      fat[cursor] = i == sectors - 1 ? _endOfChain : cursor + 1;
      cursor++;
    }
    return start;
  }

  var dirStart = chain(dirSectors);
  var miniFatStart = chain(miniFatSectors);
  var miniStreamStart = chain(miniStreamSectors);
  var largeStarts = {for (var e in largeSectors.entries) e.key: chain(e.value)};
  assert(cursor == total);

  // Directory entries
  List<int> dirEntry(
      String name, int type, int right, int child, int start, int size) {
    var entry = List<int>.filled(128, 0);
    var units = name.codeUnits;
    for (var i = 0; i < units.length; i++) {
      entry.setRange(i * 2, i * 2 + 2, le16(units[i]));
    }
    entry.setRange(64, 66, le16((units.length + 1) * 2));
    entry[66] = type;
    entry[67] = 1; // black
    entry.setRange(68, 72, le32(_free)); // left sibling
    entry.setRange(72, 76, le32(right));
    entry.setRange(76, 80, le32(child));
    entry.setRange(116, 120, le32(start));
    entry.setRange(120, 124, le32(size));
    return entry;
  }

  var dir = <int>[
    ...dirEntry('Root Entry', 5, _free, streams.isEmpty ? _free : 1,
        miniStream.isEmpty ? _endOfChain : miniStreamStart, miniStream.length),
  ];
  var index = 0;
  for (var e in streams.entries) {
    index++;
    var isSmall = small.containsKey(e.key);
    dir.addAll(dirEntry(
        e.key,
        2,
        index == streams.length ? _free : index + 1,
        _free,
        isSmall ? miniStarts[e.key]! : largeStarts[e.key]!,
        e.value.length));
  }
  while (dir.length % sectorSize != 0) {
    dir.addAll(List.filled(128, 0)
      ..setRange(68, 80, le32(_free) + le32(_free) + le32(_free)));
  }

  // Header
  var header = List<int>.filled(sectorSize, 0);
  header.setRange(0, 8, [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
  header.setRange(24, 26, le16(0x003E));
  header.setRange(26, 28, le16(sectorSize == 512 ? 3 : 4));
  header.setRange(28, 30, le16(0xFFFE));
  header.setRange(30, 32, le16(sectorSize == 512 ? 9 : 12));
  header.setRange(32, 34, le16(6));
  header.setRange(44, 48, le32(fatCount));
  header.setRange(48, 52, le32(dirStart));
  header.setRange(56, 60, le32(4096));
  header.setRange(
      60, 64, le32(miniFatSectors == 0 ? _endOfChain : miniFatStart));
  header.setRange(64, 68, le32(miniFatSectors));
  header.setRange(68, 72, le32(difatCount == 0 ? _endOfChain : difatStart));
  header.setRange(72, 76, le32(difatCount));
  for (var i = 0; i < 109; i++) {
    header.setRange(
        76 + i * 4, 80 + i * 4, le32(i < fatCount ? fatStart + i : _free));
  }

  var out = <int>[...header];
  // FAT sectors
  for (var v in fat) {
    out.addAll(le32(v));
  }
  // DIFAT sectors: FAT sector numbers 109.., last entry points to the next
  var rest = [for (var i = 109; i < fatCount; i++) fatStart + i];
  for (var d = 0; d < difatCount; d++) {
    var chunk = rest.skip(d * (perSector - 1)).take(perSector - 1).toList();
    for (var i = 0; i < perSector - 1; i++) {
      out.addAll(le32(i < chunk.length ? chunk[i] : _free));
    }
    out.addAll(le32(d == difatCount - 1 ? _endOfChain : difatStart + d + 1));
  }
  void pad() {
    while (out.length % sectorSize != 0) {
      out.add(0);
    }
  }

  pad();
  out.addAll(dir);
  pad();
  for (var v in miniFat) {
    out.addAll(le32(v));
  }
  pad();
  out.addAll(miniStream);
  pad();
  for (var e in large.entries) {
    out.addAll(e.value);
    pad();
  }
  return Uint8List.fromList(out);
}
