# Spreadsheet Decoder

[![Build Status](https://travis-ci.org/sestegra/spreadsheet_decoder.svg)](https://travis-ci.org/sestegra/spreadsheet_decoder?branch=master)
[![Coverage Status](https://coveralls.io/repos/sestegra/spreadsheet_decoder/badge.svg?branch=master)](https://coveralls.io/r/sestegra/spreadsheet_decoder?branch=master)
[![Pub version](https://img.shields.io/pub/v/spreadsheet_decoder.svg)](https://pub.dartlang.org/packages/spreadsheet_decoder)

Spreadsheet Decoder is a library for decoding and updating spreadsheets for ODS and XLSX files.

## Usage

### On server-side

    import 'dart:io';
    import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';

    main() {
      var bytes = File.fromUri(fullUri).readAsBytesSync();
      var decoder = SpreadsheetDecoder.decodeBytes(bytes);
      var table = decoder.tables['Sheet1'];
      var values = table.rows[0];
      ...
      decoder.updateCell('Sheet1', 0, 0, 1337);
      File(join(fullUri).writeAsBytesSync(decoder.encode());
      ...
    }

### On client-side

    import 'dart:html';
    import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';

    main() {
      var reader = FileReader();
      reader.onLoadEnd.listen((event) {
        var decoder = SpreadsheetDecoder.decodeBytes(reader.result);
        var table = decoder.tables['Sheet1'];
        var values = table.rows[0];
        ...
        decoder.updateCell('Sheet1', 0, 0, 1337);
        var bytes = decoder.encode();
        ...
      });
    }

## Features not yet supported
This implementation doesn't support following features:
- annotations
- spanned rows
- spanned columns
- hidden rows (visible in resulting tables)
- hidden columns (visible in resulting tables)

For XLSX format, this implementation supports the native Excel formats for date, time and boolean type conversion, plus custom `<numFmt>` format codes declared in the workbook (e.g. `dd/mm/yyyy`).

Important: Excel often stores date cells as numeric serial values and only formats them for display. The decoder applies the workbook's format code to render the date as a string. If you need a fixed text representation independent of the workbook formatting, the source cell must be stored as text in the spreadsheet.

## License

The MIT License, see [LICENSE](https://github.com/sestegra/spreadsheet_decoder/raw/master/LICENSE).
