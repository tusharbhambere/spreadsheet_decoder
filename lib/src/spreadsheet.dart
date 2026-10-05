part of spreadsheet_decoder;

const _spreasheetOds = 'ods';
const _spreasheetXlsx = 'xlsx';
final Map<String, String> _spreasheetExtensionMap = <String, String>{
  _spreasheetOds: 'application/vnd.oasis.opendocument.spreadsheet',
  _spreasheetXlsx:
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

// Normalize new line
String _normalizeNewLine(String text) {
  return text.replaceAll('\r\n', '\n');
}

SpreadsheetDecoder _newSpreadsheetDecoder(Archive archive, String dateFormat) {
  // Lookup at file format
  String? format;

  // Try OpenDocument format
  var mimetype = archive.findFile('mimetype');
  if (mimetype != null) {
    mimetype.decompress();
    var content = utf8.decode(mimetype.content);
    if (content == _spreasheetExtensionMap[_spreasheetOds]) {
      format = _spreasheetOds;
    }

    // Try OpenXml Office format
  } else {
    var xl = archive.findFile('xl/workbook.xml');
    format = xl != null ? _spreasheetXlsx : null;
  }

  switch (format) {
    case _spreasheetOds:
      return OdsDecoder(archive, dateFormat: dateFormat);
    case _spreasheetXlsx:
      return XlsxDecoder(archive, dateFormat: dateFormat);
    default:
      throw UnsupportedError('Spreadsheet format unsupported');
  }
}

/// Decode a spreadsheet file (read-only).
abstract class SpreadsheetDecoder {
  /// Default output format used for date cells when none is provided.
  /// Format tokens follow Excel-style codes (case-insensitive):
  /// `yyyy`/`yy`, `mm`/`m` (month — context-sensitive),
  /// `dd`/`d`/`ddd`/`dddd`, `hh`/`h`, `ss`/`s`, `AM/PM`.
  /// Common intl-style patterns like `dd/MM/yyyy` also work because the
  /// pattern is matched case-insensitively.
  static const String defaultDateFormat = 'yyyy-MM-dd';

  late String _dateFormat;
  late Archive _archive;

  late Map<String, SpreadsheetTable> _tables;

  /// Media type
  String get mediaType;

  /// Filename extension
  String get extension;

  /// Tables contained in spreadsheet file indexed by their names
  Map<String, SpreadsheetTable> get tables => _tables;

  SpreadsheetDecoder();

  /// Decode an XLSX/ODS/XLS spreadsheet from raw [data] bytes.
  ///
  /// Pass [dateFormat] to control how date cells are rendered as strings
  /// (e.g. `'dd/MM/yyyy'`, `'yyyy-MM-dd'`). Defaults to
  /// [defaultDateFormat] (`yyyy-MM-dd`).
  factory SpreadsheetDecoder.decodeBytes(List<int> data,
      {bool verify = false, String dateFormat = defaultDateFormat}) {
    if (_hasCfbSignature(data)) {
      return XlsDecoder(data, dateFormat: dateFormat);
    }
    var archive = ZipDecoder().decodeBytes(data, verify: verify);
    return _newSpreadsheetDecoder(archive, dateFormat);
  }

  /// Decode an XLSX/ODS/XLS spreadsheet from an [input] stream.
  factory SpreadsheetDecoder.decodeBuffer(InputStream input,
      {bool verify = false, String dateFormat = defaultDateFormat}) {
    if (input.length >= _cfbSignature.length &&
        _hasCfbSignature(input.peekBytes(_cfbSignature.length).toUint8List())) {
      return XlsDecoder(input.toUint8List(), dateFormat: dateFormat);
    }
    var archive = ZipDecoder().decodeStream(input, verify: verify);
    return _newSpreadsheetDecoder(archive, dateFormat);
  }

  void _normalizeTable(SpreadsheetTable table) {
    if (table._maxRows == 0) {
      table._rows.clear();
    } else if (table._maxRows < table._rows.length) {
      table._rows.removeRange(table._maxRows, table._rows.length);
    }
    for (var row = 0; row < table._rows.length; row++) {
      if (table._maxCols == 0) {
        table._rows[row].clear();
      } else if (table._maxCols < table._rows[row].length) {
        table._rows[row].removeRange(table._maxCols, table._rows[row].length);
      } else if (table._maxCols > table._rows[row].length) {
        var repeat = table._maxCols - table._rows[row].length;
        for (var index = 0; index < repeat; index++) {
          table._rows[row].add(null);
        }
      }
    }
  }

  bool _isEmptyRow(List row) {
    return row.fold(true, (value, element) => value && (element == null));
  }

  bool _isNotEmptyRow(List row) {
    return !_isEmptyRow(row);
  }

  void _countFilledRow(SpreadsheetTable table, List row) {
    if (_isNotEmptyRow(row)) {
      if (table._maxRows < table._rows.length) {
        table._maxRows = table._rows.length;
      }
    }
  }

  void _countFilledColumn(SpreadsheetTable table, List row, dynamic value) {
    if (value != null) {
      if (table._maxCols < row.length) {
        table._maxCols = row.length;
      }
    }
  }
}

/// Table of a spreadsheet file
class SpreadsheetTable {
  final String name;
  SpreadsheetTable(this.name);

  int _maxRows = 0;
  int _maxCols = 0;

  final List<List> _rows = <List>[];

  /// List of table's rows
  List<List> get rows => _rows;

  /// Get max rows
  int get maxRows => _maxRows;

  /// Get max cols
  int get maxCols => _maxCols;
}
