// Inspired from http://search.cpan.org/~terhechte/Spreadsheet-ReadSXC-0.20/ReadSXC.pm
// table:table format explained here http://books.evc-cit.info/odbook/ch05.html
//
// NOTE: This implementation doesn't support following features
//   - annotations
//   - spanned rows
//   - spanned columns
//   - hidden rows (visible in resulting table)
//   - hidden columns (visible in resulting table)
part of spreadsheet_decoder;

const String contentXML = 'content.xml';

/// Read and parse ODS spreadsheet
class OdsDecoder extends SpreadsheetDecoder {
  @override
  String get mediaType => 'application/vnd.oasis.opendocument.spreadsheet';
  @override
  String get extension => '.ods';

  OdsDecoder(Archive archive,
      {String dateFormat = SpreadsheetDecoder.defaultDateFormat}) {
    _archive = archive;
    _dateFormat = dateFormat;
    _tables = <String, SpreadsheetTable>{};
    _parseContent();
  }

  void _parseContent() {
    var file = _archive.findFile(contentXML);
    if (file == null) {
      throw FormatException('Missing required file: $contentXML');
    }
    file.decompress();
    var content = XmlDocument.parse(utf8.decode(file.content));
    content.findAllElements('table:table').forEach((node) {
      var name = node.getAttribute('table:name')!;
      _parseTable(node, name);
    });
  }

  void _parseTable(XmlElement node, String name) {
    var table = tables[name] = SpreadsheetTable(name);
    var rows = _findRows(node);

    // Remove tailing empty rows
    var filledRows = rows.toList().reversed.skipWhile((row) {
      return _findCells(row).every((cell) => _readCell(cell) == null);
    });

    filledRows.toList().reversed.forEach((child) {
      _parseRow(child, table);
    });

    _normalizeTable(table);
  }

  void _parseRow(XmlElement node, SpreadsheetTable table) {
    var row = [];
    var cells = _findCells(node);

    // Remove tailing empty cells
    var filledCells =
        cells.toList().reversed.skipWhile((cell) => _readCell(cell) == null);

    filledCells.toList().reversed.forEach((child) {
      _parseCell(child, table, row);
    });

    var repeat = _getRowRepeated(node);
    for (var index = 0; index < repeat; index++) {
      table._rows.add(List.from(row));
    }

    _countFilledRow(table, row);
  }

  void _parseCell(XmlElement node, SpreadsheetTable table, List row) {
    var value = _readCell(node);
    var repeat = _getCellRepeated(node);
    for (var index = 0; index < repeat; index++) {
      row.add(value);
    }

    _countFilledColumn(table, row, value);
  }

  dynamic _readCell(XmlElement node) {
    dynamic value;
    var type = node.getAttribute('office:value-type');

    switch (type) {
      case 'float':
      case 'percentage':
      case 'currency':
        value = num.parse(node.getAttribute('office:value')!);
        break;
      case 'boolean':
        value =
            node.getAttribute('office:boolean-value')!.toLowerCase() == 'true';
        break;
      case 'date':
        value = _formatDateTimeWithCode(
            DateTime.parse(node.getAttribute('office:date-value')!),
            _dateFormat);
        break;
      case 'time':
        value = node.getAttribute('office:time-value');
        value = value.substring(2, value.length - 1);
        value = value.replaceAll(RegExp('[H|M]'), ':');
        break;
      case 'string':
      default:
        var list = <String>[];
        node.findElements('text:p').forEach((child) {
          list.add(_readString(child));
        });
        value = (list.isNotEmpty) ? list.join('\n') : null;
    }
    return value;
  }

  String _readString(XmlElement node) {
    var buffer = StringBuffer();

    for (var child in node.children) {
      if (child is XmlElement) {
        buffer.write(_normalizeNewLine(_readString(child)));
      } else if (child is XmlText) {
        buffer.write(_normalizeNewLine(child.text));
      }
    }

    return buffer.toString();
  }

  static Iterable<XmlElement> _findRows(XmlElement table) =>
      table.findElements('table:table-row');

  static Iterable<XmlElement> _findCells(XmlElement row) =>
      row.findElements('table:table-cell');

  static int _getRowRepeated(XmlElement row) {
    var attr = row.getAttribute('table:number-rows-repeated');
    return (attr != null) ? int.parse(attr) : 1;
  }

  static int _getCellRepeated(XmlElement cell) {
    var attr = cell.getAttribute('table:number-columns-repeated');
    return (attr != null) ? int.parse(attr) : 1;
  }
}
