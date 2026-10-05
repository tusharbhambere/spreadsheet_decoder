// Decoder for legacy Excel binary workbooks (`.xls`): Excel 97-2003 (BIFF8)
// and Excel 5.0/95 (BIFF5/7), both wrapped in an OLE2 compound file.
//
// NOTE: This implementation doesn't support following features
//   - password protected (encrypted) workbooks
//   - BIFF2 to BIFF4 workbooks (Excel 4.0 and older, not OLE2 based)
//   - everything listed as unsupported for XLSX (annotations, spans, ...)
part of spreadsheet_decoder;

/// Read and parse legacy XLS spreadsheet.
///
class XlsDecoder extends SpreadsheetDecoder {
  @override
  String get mediaType => 'application/vnd.ms-excel';
  @override
  String get extension => '.xls';

  // Name of the workbook stream in the compound file. Excel 5.0/95 files call
  // it `Book`.
  static const List<String> _workbookStreams = ['Workbook', 'Book'];

  bool _biff8 = true;
  DateTime _epoch = DateTime.utc(1899, 12, 30);
  final List<String> _sharedStrings = <String>[];
  // Number format id of every cell XF record, indexed by XF number.
  final List<int> _numFormats = <int>[];
  final Map<int, String> _customNumFormats = <int, String>{};
  final List<_XlsSheet> _sheetInfos = <_XlsSheet>[];

  XlsDecoder(List<int> data,
      {String dateFormat = SpreadsheetDecoder.defaultDateFormat}) {
    _dateFormat = dateFormat;
    _tables = <String, SpreadsheetTable>{};
    try {
      _parse(data);
    } on RangeError catch (e) {
      // Hostile or truncated input can still step outside a buffer.
      throw FormatException('Invalid XLS: ${e.message}');
    }
  }

  void _parse(List<int> data) {
    var cfb = _CfbReader(data);
    Uint8List? stream;
    for (var name in _workbookStreams) {
      stream = cfb.readStream(name);
      if (stream != null) {
        break;
      }
    }
    if (stream == null) {
      if (cfb.hasStream('EncryptedPackage')) {
        throw UnsupportedError('Encrypted spreadsheets are unsupported');
      }
      // A valid compound file, but not a workbook (.doc, .msg, ...)
      throw UnsupportedError('Spreadsheet format unsupported');
    }

    _parseGlobals(stream);
    for (var info in _sheetInfos) {
      // Chart sheets, macro sheets and VB modules have no cells.
      if (info.type == 0) {
        _parseSheet(stream, info);
      }
    }
  }

  void _parseGlobals(Uint8List stream) {
    var reader = _BiffReader(stream);
    var bof = reader.next();
    if (bof == null || bof.id != _biffBof || bof.data.length < 4) {
      throw FormatException('Invalid XLS: missing workbook BOF record');
    }
    var version = bof.view.getUint16(0, Endian.little);
    if (version != _biffVersion8 && version != _biffVersion5) {
      throw UnsupportedError('Unsupported XLS (BIFF) version 0x'
          '${version.toRadixString(16)}');
    }
    if (bof.view.getUint16(2, Endian.little) != _biffSubstreamGlobals) {
      throw FormatException('Invalid XLS: workbook globals not found');
    }
    _biff8 = version == _biffVersion8;

    var depth = 1;
    while (depth > 0) {
      var record = reader.next();
      if (record == null) {
        throw FormatException('Invalid XLS: workbook globals are truncated');
      }
      switch (record.id) {
        case _biffBof:
          depth++;
          break;
        case _biffEof:
          depth--;
          break;
        case _biffFilePass:
          throw UnsupportedError('Encrypted spreadsheets are unsupported');
        case _biffDateMode:
          if (record.data.length >= 2 &&
              record.view.getUint16(0, Endian.little) == 1) {
            _epoch = DateTime.utc(1904, 1, 1);
          }
          break;
        case _biffBoundSheet:
          _parseBoundSheet(record);
          break;
        case _biffSst:
          if (_biff8) {
            _parseSharedStrings(record);
          }
          break;
        case _biffFormat:
          _parseFormat(record);
          break;
        case _biffXf:
          _numFormats.add(record.view.getUint16(2, Endian.little));
          break;
      }
    }
  }

  void _parseBoundSheet(_BiffRecord record) {
    var view = record.view;
    var offset = view.getUint32(0, Endian.little);
    var type = view.getUint8(5);
    var name = _readShortString(record.data, 6);
    _sheetInfos.add(_XlsSheet(name, offset, type));
  }

  String _readShortString(Uint8List data, int start) {
    var chunks = _BiffChunks([Uint8List.sublistView(data, start)]);
    return _biff8
        ? chunks.readUnicodeString(shortLength: true)
        : chunks.readByteString(shortLength: true);
  }

  void _parseFormat(_BiffRecord record) {
    var id = record.view.getUint16(0, Endian.little);
    var chunks = _BiffChunks([
      Uint8List.sublistView(record.data, 2),
      ...record.continues,
    ]);
    _customNumFormats[id] = _biff8
        ? chunks.readUnicodeString()
        // BIFF5 stores the format length in a single byte
        : chunks.readByteString(shortLength: true);
  }

  void _parseSharedStrings(_BiffRecord record) {
    var chunks = _BiffChunks(record.chunks);
    chunks.skip(4); // total number of references
    var unique = chunks.readU32();
    // Don't trust the declared count: it only bounds the loop, the data
    // decides how many strings there really are.
    for (var i = 0; i < unique && !chunks.isAtEnd; i++) {
      _sharedStrings.add(chunks.readUnicodeString());
    }
  }

  void _parseSheet(Uint8List stream, _XlsSheet info) {
    var reader = _BiffReader(stream, info.offset);
    var bof = reader.next();
    if (bof == null ||
        bof.id != _biffBof ||
        bof.data.length < 4 ||
        bof.view.getUint16(2, Endian.little) != _biffSubstreamWorksheet) {
      throw FormatException("Invalid XLS: sheet '${info.name}' not found");
    }

    var rows = <List>[];
    void put(int row, int col, dynamic value) {
      // BIFF8 grid is 65536 x 256, anything else is corrupt and would let a
      // tiny file allocate gigabytes.
      if (col > 0xFF) {
        throw FormatException('Invalid XLS: cell outside of the sheet');
      }
      while (rows.length <= row) {
        rows.add([]);
      }
      var cells = rows[row];
      while (cells.length < col) {
        cells.add(null);
      }
      if (cells.length == col) {
        cells.add(value);
      } else {
        cells[col] = value;
      }
    }

    int? pendingRow; // formula waiting for its STRING result
    int? pendingCol;
    var depth = 1;
    while (depth > 0) {
      var record = reader.next();
      if (record == null) {
        throw FormatException("Invalid XLS: sheet '${info.name}' is truncated");
      }
      if (pendingRow != null && !_isFormulaCompanion(record.id)) {
        pendingRow = pendingCol = null;
      }

      var data = record.data;
      var view = record.view;
      switch (record.id) {
        case _biffBof:
          depth++;
          break;
        case _biffEof:
          depth--;
          break;
        case _biffLabelSst:
          var index = view.getUint32(6, Endian.little);
          if (index < _sharedStrings.length) {
            put(view.getUint16(0, Endian.little),
                view.getUint16(2, Endian.little), _sharedStrings[index]);
          }
          break;
        case _biffNumber:
          put(
              view.getUint16(0, Endian.little),
              view.getUint16(2, Endian.little),
              _numberValue(view.getFloat64(6, Endian.little),
                  view.getUint16(4, Endian.little)));
          break;
        case _biffRk:
          put(
              view.getUint16(0, Endian.little),
              view.getUint16(2, Endian.little),
              _numberValue(_decodeRk(view.getUint32(6, Endian.little)),
                  view.getUint16(4, Endian.little)));
          break;
        case _biffMulRk:
          var row = view.getUint16(0, Endian.little);
          var first = view.getUint16(2, Endian.little);
          var count = (data.length - 6) ~/ 6;
          for (var i = 0; i < count; i++) {
            put(
                row,
                first + i,
                _numberValue(
                    _decodeRk(view.getUint32(6 + i * 6, Endian.little)),
                    view.getUint16(4 + i * 6, Endian.little)));
          }
          break;
        case _biffBoolErr:
          var isError = view.getUint8(7) != 0;
          put(
              view.getUint16(0, Endian.little),
              view.getUint16(2, Endian.little),
              isError
                  ? _biffErrorText(view.getUint8(6))
                  : view.getUint8(6) != 0);
          break;
        case _biffLabel:
          put(view.getUint16(0, Endian.little),
              view.getUint16(2, Endian.little), _readCellString(record));
          break;
        case _biffRString:
          // BIFF5 only: text followed by formatting runs that we ignore
          put(view.getUint16(0, Endian.little),
              view.getUint16(2, Endian.little), _readCellString(record));
          break;
        case _biffFormula:
          var row = view.getUint16(0, Endian.little);
          var col = view.getUint16(2, Endian.little);
          if (view.getUint16(12, Endian.little) == 0xFFFF) {
            // Not a number: the first byte tells which kind of result
            switch (view.getUint8(6)) {
              case 0: // string, stored in the STRING record that follows
                pendingRow = row;
                pendingCol = col;
                break;
              case 1:
                put(row, col, view.getUint8(8) != 0);
                break;
              case 2:
                put(row, col, _biffErrorText(view.getUint8(8)));
                break;
              case 3: // empty string
                put(row, col, '');
                break;
            }
          } else {
            put(
                row,
                col,
                _numberValue(view.getFloat64(6, Endian.little),
                    view.getUint16(4, Endian.little)));
          }
          break;
        case _biffString:
          if (pendingRow != null) {
            put(pendingRow, pendingCol!, _readStringResult(record));
            pendingRow = pendingCol = null;
          }
          break;
      }
    }

    var table = tables[info.name] = SpreadsheetTable(info.name);
    for (var row in rows) {
      table._rows.add(row);
      if (row.isNotEmpty) {
        table._maxRows = table._rows.length;
        if (table._maxCols < row.length) {
          table._maxCols = row.length;
        }
      }
    }
    _normalizeTable(table);
  }

  static bool _isFormulaCompanion(int id) =>
      id == _biffString ||
      id == _biffArray ||
      id == _biffSharedFormula ||
      id == _biffTableOp ||
      id == _biffTableOp2 ||
      id == _biffTableOp3;

  /// Value of a numeric cell, rendered as a date or time when its number
  /// format says so.
  dynamic _numberValue(num number, int xf) {
    var format = xf < _numFormats.length ? _numFormats[xf] : 0;
    return _formatNumericCell(
        _normalizeNumber(number), format, _customNumFormats, _dateFormat,
        epoch: _epoch);
  }

  /// Text of a LABEL / RSTRING record: cell address (6 bytes) then the text.
  String _readCellString(_BiffRecord record) {
    var chunks = _BiffChunks([
      Uint8List.sublistView(record.data, 6),
      ...record.continues,
    ]);
    return _biff8 ? chunks.readUnicodeString() : chunks.readByteString();
  }

  String _readStringResult(_BiffRecord record) {
    var chunks = _BiffChunks(record.chunks);
    return _biff8 ? chunks.readUnicodeString() : chunks.readByteString();
  }
}

class _XlsSheet {
  final String name;
  // Position of the sheet's BOF record in the workbook stream
  final int offset;
  // 0 worksheet, 1 macro sheet, 2 chart, 6 VB module
  final int type;

  _XlsSheet(this.name, this.offset, this.type);
}
