import 'dart:io';
import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';

void main(List<String> args) {
  var file = 'test/files/new_sample_file_bytco.xlsx';
  var bytes = File(file).readAsBytesSync();
  var decoder = SpreadsheetDecoder.decodeBytes(bytes, dateFormat: 'dd/MM/yyyy');
  for (var table in decoder.tables.keys) {
    print(table);
    print(decoder.tables[table]!.maxCols);
    print(decoder.tables[table]!.maxRows);
    for (var row in decoder.tables[table]!.rows) {
      print('$row');
    }
  }
}
