// Low level helpers for the BIFF (Binary Interchange File Format) record
// stream stored inside a legacy `.xls` workbook. See [MS-XLS].
part of spreadsheet_decoder;

// Record identifiers
const int _biffBof = 0x0809;
const int _biffEof = 0x000A;
const int _biffContinue = 0x003C;
const int _biffFilePass = 0x002F;
const int _biffDateMode = 0x0022;
const int _biffBoundSheet = 0x0085;
const int _biffSst = 0x00FC;
const int _biffFormat = 0x041E;
const int _biffXf = 0x00E0;
const int _biffLabelSst = 0x00FD;
const int _biffLabel = 0x0204;
const int _biffRString = 0x00D6;
const int _biffNumber = 0x0203;
const int _biffRk = 0x027E;
const int _biffMulRk = 0x00BD;
const int _biffBoolErr = 0x0205;
const int _biffFormula = 0x0006;
const int _biffString = 0x0207;
// Records that may sit between a FORMULA and its cached STRING result.
const int _biffArray = 0x0221;
const int _biffSharedFormula = 0x04BC;
const int _biffTableOp = 0x0236;
const int _biffTableOp2 = 0x0037;
const int _biffTableOp3 = 0x0036;

// BOF substream types
const int _biffSubstreamGlobals = 0x0005;
const int _biffSubstreamWorksheet = 0x0010;

// BOF versions
const int _biffVersion5 = 0x0500;
const int _biffVersion8 = 0x0600;

/// One BIFF record with the payload of the CONTINUE records that follow it.
class _BiffRecord {
  final int id;
  final Uint8List data;
  final List<Uint8List> continues;

  _BiffRecord(this.id, this.data, this.continues);

  ByteData get view => ByteData.sublistView(data);

  /// Payload split at the record boundaries, as needed to decode strings
  /// that straddle a CONTINUE record.
  List<Uint8List> get chunks => [data, ...continues];
}

/// Sequential reader of BIFF records.
class _BiffReader {
  final Uint8List _bytes;
  final ByteData _view;
  int position;

  _BiffReader(this._bytes, [this.position = 0])
      : _view = ByteData.sublistView(_bytes) {
    if (position < 0 || position > _bytes.length) {
      throw FormatException('Invalid XLS: stream offset out of range');
    }
  }

  /// Returns the next record, or `null` at the end of the stream.
  _BiffRecord? next() {
    var record = _readRaw();
    if (record == null) {
      return null;
    }
    List<Uint8List>? continues;
    while (position + 4 <= _bytes.length &&
        _view.getUint16(position, Endian.little) == _biffContinue) {
      var extra = _readRaw()!;
      (continues ??= <Uint8List>[]).add(extra.data);
    }
    return _BiffRecord(record.id, record.data, continues ?? const []);
  }

  _BiffRecord? _readRaw() {
    if (position + 4 > _bytes.length) {
      return null;
    }
    var id = _view.getUint16(position, Endian.little);
    var length = _view.getUint16(position + 2, Endian.little);
    var start = position + 4;
    if (start + length > _bytes.length) {
      throw FormatException('Invalid XLS: truncated record 0x'
          '${id.toRadixString(16)}');
    }
    position = start + length;
    return _BiffRecord(
        id, Uint8List.sublistView(_bytes, start, start + length), const []);
  }
}

/// Reads fields and strings from a record payload that may be split over
/// several CONTINUE records.
class _BiffChunks {
  final List<Uint8List> _chunks;
  int _chunk = 0;
  int _offset = 0;

  _BiffChunks(this._chunks);

  /// Moves to the next non empty chunk. Returns false at the end of data.
  bool _fill() {
    while (_chunk < _chunks.length && _offset >= _chunks[_chunk].length) {
      _chunk++;
      _offset = 0;
    }
    return _chunk < _chunks.length;
  }

  bool get isAtEnd => !_fill();

  int readU8() {
    if (!_fill()) {
      throw FormatException('Invalid XLS: unexpected end of record');
    }
    return _chunks[_chunk][_offset++];
  }

  int readU16() => readU8() + readU8() * 0x100;

  int readU32() => readU16() + readU16() * 0x10000;

  void skip(int count) {
    while (count > 0) {
      if (!_fill()) {
        throw FormatException('Invalid XLS: unexpected end of record');
      }
      var n = _chunks[_chunk].length - _offset;
      if (n > count) {
        n = count;
      }
      _offset += n;
      count -= n;
    }
  }

  /// Reads [count] characters. When the text straddles a CONTINUE record the
  /// continuation starts with a new flag byte that may switch between 8 and
  /// 16 bit characters.
  String readChars(int count, bool wide) {
    var buffer = StringBuffer();
    while (count > 0) {
      if (_chunk >= _chunks.length) {
        throw FormatException('Invalid XLS: unexpected end of string');
      }
      var chunk = _chunks[_chunk];
      if (_offset >= chunk.length) {
        _chunk++;
        _offset = 0;
        if (_chunk >= _chunks.length) {
          throw FormatException('Invalid XLS: unexpected end of string');
        }
        // Skip empty continuations, then read the new encoding flag.
        if (_chunks[_chunk].isEmpty) {
          continue;
        }
        wide = (_chunks[_chunk][_offset++] & 0x01) != 0;
        continue;
      }
      var width = wide ? 2 : 1;
      var n = (chunk.length - _offset) ~/ width;
      if (n == 0) {
        throw FormatException('Invalid XLS: truncated string');
      }
      if (n > count) {
        n = count;
      }
      if (wide) {
        var view = ByteData.sublistView(chunk);
        for (var i = 0; i < n; i++) {
          buffer.writeCharCode(view.getUint16(_offset + i * 2, Endian.little));
        }
      } else {
        buffer.write(String.fromCharCodes(chunk, _offset, _offset + n));
      }
      _offset += n * width;
      count -= n;
    }
    return buffer.toString();
  }

  /// Reads [count] single byte characters in the given [codePage] flavour.
  String readByteChars(int count) {
    var buffer = StringBuffer();
    while (count > 0) {
      if (!_fill()) {
        throw FormatException('Invalid XLS: unexpected end of string');
      }
      var chunk = _chunks[_chunk];
      var n = chunk.length - _offset;
      if (n > count) {
        n = count;
      }
      for (var i = 0; i < n; i++) {
        buffer.writeCharCode(_decodeAnsi(chunk[_offset + i]));
      }
      _offset += n;
      count -= n;
    }
    return buffer.toString();
  }

  /// BIFF8 `XLUnicodeString` / `ShortXLUnicodeString`: character count,
  /// option flags, optional rich text run and phonetic counts, characters,
  /// then the run and phonetic data which is skipped.
  String readUnicodeString({bool shortLength = false}) {
    var length = shortLength ? readU8() : readU16();
    var flags = readU8();
    var runs = (flags & 0x08) != 0 ? readU16() : 0;
    var phonetic = (flags & 0x04) != 0 ? readU32() : 0;
    var text = length == 0 ? '' : readChars(length, (flags & 0x01) != 0);
    skip(runs * 4);
    skip(phonetic);
    return text;
  }

  /// BIFF5/7 string: a character count followed by single byte characters.
  String readByteString({bool shortLength = false}) {
    var length = shortLength ? readU8() : readU16();
    return readByteChars(length);
  }
}

const List<int> _windows1252High = [
  0x20AC, 0x0081, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, //
  0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008D, 0x017D, 0x008F,
  0x0090, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014,
  0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x009D, 0x017E, 0x0178,
];

/// Decode one byte of a BIFF5/7 string. The BIFF5 code page is not tracked:
/// text is assumed to be Windows-1252, which covers western workbooks.
int _decodeAnsi(int byte) =>
    byte >= 0x80 && byte <= 0x9F ? _windows1252High[byte - 0x80] : byte;

/// Decode an `RkNumber` (30 bit integer or float, optionally divided by 100).
///
/// Written with arithmetic instead of shifts: dart2js evaluates `>>` and `<<`
/// on 32 bit unsigned values, which breaks negative numbers.
num _decodeRk(int rk) {
  var flags = rk & 0x03;
  num value;
  if ((flags & 0x02) != 0) {
    // signed 30 bit integer
    value = (rk.toSigned(32) - (rk.toSigned(32) & 0x03)) ~/ 4;
  } else {
    // the 30 most significant bits of an IEEE double
    var bytes = ByteData(8)..setUint32(4, rk - flags, Endian.little);
    value = bytes.getFloat64(0, Endian.little);
  }
  return (flags & 0x01) != 0 ? value / 100 : value;
}

/// Excel returns whole numbers as integers when parsing XLSX (`<v>1</v>`);
/// mirror that for the binary double so both formats yield the same values.
num _normalizeNumber(num value) {
  if (value is double &&
      value.isFinite &&
      value.abs() < 1e15 &&
      value == value.truncateToDouble()) {
    return value.toInt();
  }
  return value;
}

/// Text of a BIFF error code, as Excel displays it.
String _biffErrorText(int code) {
  switch (code) {
    case 0x00:
      return '#NULL!';
    case 0x07:
      return '#DIV/0!';
    case 0x0F:
      return '#VALUE!';
    case 0x17:
      return '#REF!';
    case 0x1D:
      return '#NAME?';
    case 0x24:
      return '#NUM!';
    case 0x2A:
      return '#N/A';
    case 0x2B:
      return '#GETTING_DATA';
    default:
      return '#ERR$code';
  }
}
