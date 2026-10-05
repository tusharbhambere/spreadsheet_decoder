@TestOn('vm')
library;

import 'dart:convert';
import 'dart:io';

import 'package:archive/archive.dart';
import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';
import 'package:test/test.dart';

const _macro = [0xD0, 0xCF, 0x11, 0xE0, 0x01, 0x02, 0x03]; // fake vbaProject

const _mainType =
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
const _relationshipType =
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/// Rewrites test.xlsx into another SpreadsheetML flavour by changing the
/// workbook content type, the way Excel stores `.xlsm`, `.xltx` and `.xltm`.
List<int> _variant(String contentType,
    {bool macros = false, String? extraSheet}) {
  var archive =
      ZipDecoder().decodeBytes(File('test/files/test.xlsx').readAsBytesSync());
  var out = Archive();
  for (var file in archive.files) {
    if (!file.isFile) continue;
    var content = file.content as List<int>;
    if (file.name == '[Content_Types].xml') {
      var types = utf8.decode(content);
      if (!types.contains(_mainType)) {
        throw StateError('fixture changed: workbook content type not found');
      }
      content = utf8.encode(types.replaceAll(_mainType, contentType));
    } else if (extraSheet != null && file.name == 'xl/workbook.xml') {
      content = utf8.encode(utf8.decode(content).replaceFirst('</sheets>',
          '<sheet name="$extraSheet" sheetId="9" r:id="rIdChart"/></sheets>'));
    } else if (extraSheet != null &&
        file.name == 'xl/_rels/workbook.xml.rels') {
      content = utf8.encode(utf8.decode(content).replaceFirst(
          '</Relationships>',
          '<Relationship Id="rIdChart" Type="$_relationshipType/chartsheet" '
              'Target="chartsheets/sheet1.xml"/></Relationships>'));
    }
    out.addFile(ArchiveFile(file.name, content.length, content));
  }
  if (macros) {
    out.addFile(ArchiveFile('xl/vbaProject.bin', _macro.length, _macro));
  }
  return ZipEncoder().encode(out);
}

void main() {
  var variants = {
    '.xlsx': [
      _mainType,
      false,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    ],
    '.xlsm': [
      'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
      true,
      'application/vnd.ms-excel.sheet.macroEnabled.12'
    ],
    '.xltx': [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml',
      false,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.template'
    ],
    '.xltm': [
      'application/vnd.ms-excel.template.macroEnabled.main+xml',
      true,
      'application/vnd.ms-excel.template.macroEnabled.12'
    ],
  };

  var reference = SpreadsheetDecoder.decodeBytes(
      File('test/files/test.xlsx').readAsBytesSync());

  variants.forEach((extension, info) {
    group('$extension workbook', () {
      var contentType = info[0] as String;
      var macros = info[1] as bool;
      var mediaType = info[2] as String;
      var bytes = _variant(contentType, macros: macros);

      test('is decoded like an xlsx', () {
        var decoder = SpreadsheetDecoder.decodeBytes(bytes);
        expect(decoder, isA<XlsxDecoder>());
        expect(decoder.extension, extension);
        expect(decoder.mediaType, mediaType);
        expect(decoder.tables.keys, reference.tables.keys);
        for (var name in reference.tables.keys) {
          expect(decoder.tables[name]!.rows, reference.tables[name]!.rows);
        }
      });
    });
  });

  test('chart sheets are ignored instead of failing', () {
    var bytes = _variant(_mainType, extraSheet: 'Chart1');
    var decoder = SpreadsheetDecoder.decodeBytes(bytes);
    expect(decoder.tables.keys, reference.tables.keys);
  });

  test('a workbook without content types is treated as xlsx', () {
    var archive = ZipDecoder()
        .decodeBytes(File('test/files/test.xlsx').readAsBytesSync());
    var out = Archive();
    for (var file in archive.files) {
      if (file.isFile && file.name != '[Content_Types].xml') {
        out.addFile(ArchiveFile(file.name, file.size, file.content));
      }
    }
    var decoder = SpreadsheetDecoder.decodeBytes(ZipEncoder().encode(out));
    expect(decoder.extension, '.xlsx');
    expect(decoder.tables.keys, reference.tables.keys);
  });

  test('binary workbooks (xlsb) are reported as unsupported', () {
    var archive = Archive()
      ..addFile(ArchiveFile('[Content_Types].xml', 6, utf8.encode('<Types>')))
      ..addFile(ArchiveFile('xl/workbook.bin', 1, [0]));
    expect(() => SpreadsheetDecoder.decodeBytes(ZipEncoder().encode(archive)),
        throwsUnsupportedError);
  });
}
