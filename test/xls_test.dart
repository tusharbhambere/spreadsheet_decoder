@TestOn('vm')
library;

import 'dart:io';
import 'dart:math';
import 'dart:typed_data';

import 'package:archive/archive.dart';
import 'package:spreadsheet_decoder/spreadsheet_decoder.dart';
import 'package:test/test.dart';

import 'xls_builder.dart';

Uint8List _readFile(String name) => File('test/files/$name').readAsBytesSync();

SpreadsheetDecoder _decodeFile(String name, {String? dateFormat}) =>
    SpreadsheetDecoder.decodeBytes(_readFile(name),
        dateFormat: dateFormat ?? SpreadsheetDecoder.defaultDateFormat);

Map<String, List<List>> _rows(SpreadsheetDecoder decoder) =>
    {for (var e in decoder.tables.entries) e.key: e.value.rows};

/// One worksheet called `S`, wrapped in a compound file.
Uint8List _sheetFile(
  List<List<int>> cells, {
  List<List<int>> globals = const [],
  bool biff8 = true,
  int sectorSize = 512,
}) =>
    compoundFile({
      biff8 ? 'Workbook' : 'Book': workbookStream([SheetSpec('S', cells)],
          globals: globals, biff8: biff8),
    }, sectorSize: sectorSize);

List<List> _decodeSheet(
  List<List<int>> cells, {
  List<List<int>> globals = const [],
  String dateFormat = SpreadsheetDecoder.defaultDateFormat,
}) {
  var decoder = SpreadsheetDecoder.decodeBytes(
      _sheetFile(cells, globals: globals),
      dateFormat: dateFormat);
  return decoder.tables['S']!.rows;
}

int _rkInt(int value) => ((value << 2) | 0x02) & 0xFFFFFFFF;

/// RK encoding of a double whose low 32 bits are zero, like 1.5.
int _rkFloat(double value) {
  var high = ByteData(8)..setFloat64(0, value, Endian.little);
  return high.getUint32(4, Endian.little) & 0xFFFFFFFC;
}

void main() {
  group('xls fixtures', () {
    test('same content as the xlsx twin', () {
      expect(_rows(_decodeFile('test.xls')), _rows(_decodeFile('test.xlsx')));
    });

    test('sheet and table geometry', () {
      var tables = _decodeFile('test.xls').tables;
      expect(tables.keys, ['ONE', 'TWO', 'THREE', 'EMPTY']);
      expect(tables['ONE']!.maxRows, 12);
      expect(tables['ONE']!.maxCols, 3);
      expect(tables['EMPTY']!.rows, isEmpty);
      expect(tables['EMPTY']!.maxRows, 0);
    });

    test('exposes media type and extension', () {
      var decoder = _decodeFile('test.xls');
      expect(decoder, isA<XlsDecoder>());
      expect(decoder.mediaType, 'application/vnd.ms-excel');
      expect(decoder.extension, '.xls');
    });

    test('sheet names with spaces and accents', () {
      var tables = _decodeFile('sheets.xls').tables;
      expect(tables.keys, ['First sheet', 'Éte']);
      expect(tables['Éte']!.rows, [
        ['b', null],
        [null, 2]
      ]);
    });

    test('value types, unicode and sparse cells', () {
      var rows = _decodeFile('types.xls').tables['Types']!.rows;
      expect(rows[0].sublist(0, 9), [
        'text',
        42,
        -7,
        3.14159,
        -1500.99,
        0.01,
        1234567890123,
        true,
        false,
      ]);
      expect(rows[1].sublist(0, 4),
          ['日本語と漢字', 'Ünïcödé àçcents', 'line1\nline2', '😀 emoji']);
      expect(rows[4][2], 'C5');
      expect(rows[7][5], 'F8');
      expect(rows[2].every((v) => v == null), isTrue);
      // rows are rectangular, like for xlsx
      expect(rows.every((r) => r.length == 9), isTrue);
    });

    test('dates and times use the workbook format codes', () {
      var rows = _decodeFile('types.xls').tables['Types']!.rows;
      expect(rows[9].sublist(0, 4),
          ['2008-07-21', '2008-07-21', '21/07/2008', '2008-07-21 13:45']);
      expect(rows[10].sublist(0, 4), ['13:45:30', 0.5, 1234.5, 1]);
    });

    test('dateFormat overrides built in and custom date formats', () {
      var rows = _decodeFile('types.xls', dateFormat: 'dd.MM.yyyy')
          .tables['Types']!
          .rows;
      expect(rows[9].sublist(0, 4),
          ['21.07.2008', '21.07.2008', '21.07.2008', '21.07.2008']);
      // not dates: untouched
      expect(rows[10].sublist(1, 4), [0.5, 1234.5, 1]);
    });

    test('1904 date system', () {
      var rows = _decodeFile('date1904.xls').tables['Dates1904']!.rows;
      expect(rows[0], ['2008-07-21', '1904-01-01']);
    });

    test('shared strings spanning CONTINUE records', () {
      var table = _decodeFile('strings.xls').tables['Strings']!;
      expect(table.maxRows, 1000);
      expect(table.maxCols, 4);
      for (var i = 0; i < 2500; i++) {
        var text = 'row $i ${'é' * (i % 7)}';
        if (i % 11 == 0) text += ' 日本語';
        var row = i % 1000, col = i ~/ 1000;
        if (row == 0 && col == 3 || row == 1 && col == 3) continue;
        expect(table.rows[row][col], text, reason: 'cell $row,$col');
      }
      expect(table.rows[0][3], 'x' * 9000);
      expect(table.rows[1][3], '日' * 5000);
    });

    test('decodeBuffer', () {
      var decoder = SpreadsheetDecoder.decodeBuffer(
          InputMemoryStream(_readFile('test.xls')));
      expect(_rows(decoder), _rows(_decodeFile('test.xlsx')));
    });
  });

  group('xls cells', () {
    test('RK and MULRK numbers', () {
      var rows = _decodeSheet([
        rk(0, 0, _rkInt(5)),
        rk(0, 1, _rkInt(-3)),
        rk(0, 2, _rkInt(123) | 0x01), // 123 / 100
        rk(0, 3, _rkFloat(1.5)),
        rk(0, 4, _rkFloat(150.0) | 0x01), // 150 / 100
        mulRk(1, 1, [_rkInt(10), _rkInt(20), _rkInt(-30)]),
      ]);
      expect(rows[0], [5, -3, 1.23, 1.5, 1.5]);
      expect(rows[1], [null, 10, 20, -30, null]);
    });

    test('whole doubles become integers, NaN and infinity survive', () {
      var rows = _decodeSheet([
        number(0, 0, 7),
        number(0, 1, 7.25),
        number(0, 2, double.nan),
        number(0, 3, double.infinity),
        number(0, 4, 1e20),
      ]);
      expect(rows[0][0], isA<int>());
      expect(rows[0][0], 7);
      expect(rows[0][1], 7.25);
      expect((rows[0][2] as double).isNaN, isTrue);
      expect(rows[0][3], double.infinity);
      expect(rows[0][4], 1e20);
    });

    test('NaN and out of range serials in date cells do not throw', () {
      var rows = _decodeSheet([
        number(0, 0, double.nan, xfIndex: 0),
        number(0, 1, 1e300, xfIndex: 0),
      ], globals: [
        xf(14)
      ]);
      expect((rows[0][0] as double).isNaN, isTrue);
      expect(rows[0][1], 1e300);
    });

    test('labels in the shared string table', () {
      var rows = _decodeSheet([
        labelSst(0, 0, 1),
        labelSst(0, 1, 0),
        labelSst(1, 0, 99)
      ], globals: [
        sst(['alpha', '日本'])
      ]);
      expect(rows[0], ['日本', 'alpha']);
      // an index outside of the table leaves the cell empty
      expect(rows.length, 1);
    });

    test('inline LABEL records', () {
      var rows = _decodeSheet([
        label(0, 0, xlString('plain')),
        label(0, 1, xlString('wide ☃')),
      ]);
      expect(rows[0], ['plain', 'wide ☃']);
    });

    test('booleans and errors', () {
      var rows = _decodeSheet([
        boolErr(0, 0, 1),
        boolErr(0, 1, 0),
        boolErr(0, 2, 0x07, error: true),
        boolErr(0, 3, 0x2A, error: true),
        boolErr(0, 4, 0x0F, error: true),
      ]);
      expect(rows[0], [true, false, '#DIV/0!', '#N/A', '#VALUE!']);
    });

    test('cached formula results', () {
      var rows = _decodeSheet([
        formulaNumber(0, 0, 3.5),
        formulaString(0, 1, xlString('computed')),
        formulaBool(0, 2, true),
        formulaError(0, 3, 0x17),
        formulaEmptyString(0, 4),
        // a shared formula record sits between FORMULA and its STRING
        formulaString(1, 0, xlString('shared'),
            between: record(0x04BC, List.filled(10, 0))),
        // a string formula without its STRING record must not leak into the
        // next, unrelated, record
        ...[
          record(0x0006, [
            ...le16(2),
            ...le16(0),
            ...le16(0),
            0, 0, 0, 0, 0, 0, 0xFF, 0xFF, //
            ...le16(0), ...le32(0), ...le16(0),
          ]),
          number(3, 0, 1),
          record(0x0207, xlString('stray')),
        ],
      ]);
      expect(rows[0], [3.5, 'computed', true, '#REF!', '']);
      expect(rows[1][0], 'shared');
      expect(rows[2].every((v) => v == null), isTrue);
      expect(rows[3][0], 1);
    });

    test('long STRING result continued in another record', () {
      var text = 'z' * 300;
      var encoded = xlString(text);
      var rows = _decodeSheet([
        formulaStringParts(0, 0, [
          encoded.sublist(0, 103), // length, flags and 100 characters
          [0, ...encoded.sublist(103)], // CONTINUE: flag byte, 200 characters
        ])
      ]);
      expect(rows[0][0], text);
    });

    test('dates through XF and FORMAT records', () {
      var globals = [
        format(164, 'dd/mm/yyyy'),
        format(165, 'hh:mm'),
        xf(0),
        xf(14),
        xf(164),
        xf(21),
        xf(165),
      ];
      var rows = _decodeSheet([
        number(0, 0, 39650, xfIndex: 0), // General
        number(0, 1, 39650, xfIndex: 1), // built in date
        number(0, 2, 39650, xfIndex: 2), // custom date
        number(0, 3, 0.5, xfIndex: 3), // built in time
        number(0, 4, 0.75, xfIndex: 4), // custom time
        number(0, 5, 39650, xfIndex: 99), // unknown XF: plain number
        mulRk(1, 0, [_rkInt(39650), _rkInt(39650)], xfIndex: 2),
        formulaNumber(2, 0, 39650, xfIndex: 1),
      ], globals: globals);
      expect(rows[0],
          [39650, '2008-07-21', '21/07/2008', '12:00:00', '18:00', 39650]);
      expect(rows[1].sublist(0, 2), ['21/07/2008', '21/07/2008']);
      expect(rows[2][0], '2008-07-21');
    });

    test('1904 date system shifts the epoch', () {
      var rows = _decodeSheet([
        number(0, 0, 38188, xfIndex: 0),
      ], globals: [
        dateMode(1),
        xf(14)
      ]);
      expect(rows[0][0], '2008-07-21');
    });

    test('chart sheets are skipped, hidden order is kept', () {
      var stream = workbookStream([
        SheetSpec('A', [number(0, 0, 1)]),
        SheetSpec('Chart', const [], type: 2),
        SheetSpec('B', [number(0, 0, 2)]),
      ]);
      var decoder =
          SpreadsheetDecoder.decodeBytes(compoundFile({'Workbook': stream}));
      expect(decoder.tables.keys, ['A', 'B']);
      expect(decoder.tables['B']!.rows, [
        [2]
      ]);
    });

    test('rows and columns are normalised like xlsx', () {
      var rows = _decodeSheet([
        number(2, 1, 5),
        number(0, 3, 6),
      ]);
      expect(rows, [
        [null, null, null, 6],
        [null, null, null, null],
        [null, 5, null, null],
      ]);
    });

    test('a cell outside of the BIFF8 grid is rejected', () {
      expect(() => _decodeSheet([number(0, 300, 1)]), throwsFormatException);
    });
  });

  group('xls shared strings', () {
    List<int> splitSst(List<List<int>> parts) =>
        recordWithContinues(0x00FC, parts);

    List<List> decodeWith(List<int> sstRecord, int count) {
      var cells = [for (var i = 0; i < count; i++) labelSst(i, 0, i)];
      return _decodeSheet(cells, globals: [sstRecord]);
    }

    test('split inside characters, 8 bit to 16 bit', () {
      // "abcdef": "abc" is stored 8 bit, the CONTINUE restarts with a flag
      // byte announcing 16 bit characters for "def"
      var rows = decodeWith(
          splitSst([
            [...le32(1), ...le32(1), ...le16(6), 0, 0x61, 0x62, 0x63],
            [1, 0x64, 0, 0x65, 0, 0x66, 0],
          ]),
          1);
      expect(rows[0][0], 'abcdef');
    });

    test('split inside characters, 16 bit to 8 bit', () {
      var rows = decodeWith(
          splitSst([
            [...le32(1), ...le32(1), ...le16(4), 1, 0x61, 0, 0x62, 0],
            [0, 0x63, 0x64],
          ]),
          1);
      expect(rows[0][0], 'abcd');
    });

    test('split between two strings', () {
      var rows = decodeWith(
          splitSst([
            [...le32(2), ...le32(2), ...xlString('one')],
            xlString('two'),
          ]),
          2);
      expect([rows[0][0], rows[1][0]], ['one', 'two']);
    });

    test('split between header fields and characters', () {
      var rows = decodeWith(
          splitSst([
            [...le32(1), ...le32(1), ...le16(3), 0],
            [0, 0x61, 0x62, 0x63],
          ]),
          1);
      expect(rows[0][0], 'abc');
    });

    test('rich text runs and phonetic data are skipped', () {
      var rich = [
        ...le16(2),
        0x08,
        ...le16(2),
        0x61,
        0x62,
        ...List.filled(8, 9)
      ];
      var phonetic = [
        ...le16(2),
        0x04,
        ...le32(5),
        0x63,
        0x64,
        ...List.filled(5, 7),
      ];
      var rows = decodeWith(
          splitSst([
            [...le32(3), ...le32(3), ...rich, ...phonetic, ...xlString('end')],
          ]),
          3);
      expect([rows[0][0], rows[1][0], rows[2][0]], ['ab', 'cd', 'end']);
    });

    test('rich text runs split over a CONTINUE record', () {
      var rows = decodeWith(
          splitSst([
            [
              ...le32(2),
              ...le32(2),
              ...le16(2),
              0x08,
              ...le16(2),
              0x61,
              0x62,
              1,
              2,
              3
            ],
            // rest of the runs, no flag byte: not inside characters
            [4, 5, 6, 7, 8, ...xlString('next')],
          ]),
          2);
      expect([rows[0][0], rows[1][0]], ['ab', 'next']);
    });

    test('a count larger than the data is tolerated', () {
      var rows = decodeWith(
          record(0x00FC, [...le32(900), ...le32(900), ...xlString('only')]), 1);
      expect(rows[0][0], 'only');
    });
  });

  group('xls other versions and containers', () {
    test('BIFF5 workbook in a Book stream', () {
      var rows = SpreadsheetDecoder.decodeBytes(_sheetFile(
          [
            label(0, 0, [...le16(5), ...'hello'.codeUnits]),
            label(0, 1, [...le16(2), 0x80, 0xE9]), // € é in Windows-1252
            number(0, 2, 12, xfIndex: 1),
            number(1, 0, 39650, xfIndex: 0),
          ],
          biff8: false,
          globals: [
            format(164, 'dd/mm/yyyy', biff8: false),
            xf(164),
            xf(0),
          ])).tables['S']!.rows;
      expect(rows[0], ['hello', '€é', 12]);
      expect(rows[1][0], '21/07/2008');
    });

    test('compound file with 4096 byte sectors', () {
      var rows = SpreadsheetDecoder.decodeBytes(
              _sheetFile([number(0, 0, 1), number(0, 1, 2)], sectorSize: 4096))
          .tables['S']!
          .rows;
      expect(rows, [
        [1, 2]
      ]);
    });

    test('workbook stored in regular sectors, not in the mini stream', () {
      var cells = [for (var i = 0; i < 400; i++) number(i, 0, i.toDouble())];
      var file = _sheetFile(cells);
      var rows = SpreadsheetDecoder.decodeBytes(file).tables['S']!.rows;
      expect(rows.length, 400);
      expect(rows[399][0], 399);
    });

    test('FAT spread over the DIFAT chain', () {
      var stream = workbookStream([
        SheetSpec('S', [number(0, 0, 1)])
      ]);
      // trailing padding after the last EOF is ignored by the reader
      var padded = Uint8List(7500000)..setRange(0, stream.length, stream);
      var rows =
          SpreadsheetDecoder.decodeBytes(compoundFile({'Workbook': padded}))
              .tables['S']!
              .rows;
      expect(rows, [
        [1]
      ]);
    });

    test('root level stream found among other streams', () {
      var stream = workbookStream([
        SheetSpec('S', [number(0, 0, 1)])
      ]);
      var rows = SpreadsheetDecoder.decodeBytes(compoundFile({
        '\u0005SummaryInformation': List.filled(100, 1),
        'Workbook': stream,
        'Other': List.filled(5000, 2),
      })).tables['S']!.rows;
      expect(rows, [
        [1]
      ]);
    });
  });

  group('xls invalid input', () {
    test('compound file that is not a workbook', () {
      var file = compoundFile({'WordDocument': List.filled(200, 0)});
      expect(
          () => SpreadsheetDecoder.decodeBytes(file), throwsUnsupportedError);
    });

    test('encrypted xlsx container', () {
      var file = compoundFile({
        'EncryptionInfo': List.filled(100, 0),
        'EncryptedPackage': List.filled(100, 0),
      });
      expect(
          () => SpreadsheetDecoder.decodeBytes(file),
          throwsA(isA<UnsupportedError>()
              .having((e) => e.message, 'message', contains('Encrypted'))));
    });

    test('password protected xls (FILEPASS)', () {
      var file = _sheetFile([number(0, 0, 1)],
          globals: [record(0x002F, List.filled(54, 0))]);
      expect(
          () => SpreadsheetDecoder.decodeBytes(file),
          throwsA(isA<UnsupportedError>()
              .having((e) => e.message, 'message', contains('Encrypted'))));
    });

    test('old BIFF versions', () {
      var file = compoundFile({
        'Workbook': record(0x0809, [...le16(0x0400), ...le16(5), ...le32(0)])
      });
      expect(
          () => SpreadsheetDecoder.decodeBytes(file), throwsUnsupportedError);
    });

    test('workbook without BOF', () {
      var file = compoundFile({'Workbook': List.filled(64, 0)});
      expect(() => SpreadsheetDecoder.decodeBytes(file), throwsFormatException);
    });

    test('truncated sheet', () {
      var stream = workbookStream([
        SheetSpec('S', [number(0, 0, 1)])
      ]);
      // chop the trailing EOF record of the sheet
      var file =
          compoundFile({'Workbook': stream.sublist(0, stream.length - 4)});
      expect(() => SpreadsheetDecoder.decodeBytes(file), throwsFormatException);
    });

    test('sheet offset not pointing to a BOF', () {
      var stream = workbookStream([
        SheetSpec('S', [number(0, 0, 1)])
      ]);
      // BOUNDSHEET payload starts right after the globals BOF (20 bytes) and
      // the 4 byte record header
      stream.setRange(24, 28, le32(0));
      var file = compoundFile({'Workbook': stream});
      expect(() => SpreadsheetDecoder.decodeBytes(file), throwsFormatException);
    });

    test('sector chain loop', () {
      var stream = workbookStream([
        SheetSpec('S', [for (var i = 0; i < 400; i++) number(i, 0, 1)])
      ]);
      var file = compoundFile({'Workbook': stream});
      // layout: header, FAT, directory, then the stream from sector 2: make
      // sector 3 point back to sector 2
      ByteData.sublistView(file).setUint32(512 + 3 * 4, 2, Endian.little);
      expect(() => SpreadsheetDecoder.decodeBytes(file), throwsFormatException);
    });

    test('stream larger than its sector chain', () {
      var file = _sheetFile([number(0, 0, 1)]);
      var bad = Uint8List.fromList(file);
      // directory is sector 1: patch the size of the Workbook entry (128 +
      // 120 relative to the directory start) to 100000 bytes
      ByteData.sublistView(bad)
          .setUint32(512 * 2 + 128 + 120, 100000, Endian.little);
      expect(() => SpreadsheetDecoder.decodeBytes(bad), throwsFormatException);
    });

    test('garbage after the compound file signature', () {
      var bytes = Uint8List(1024)
        ..setRange(0, 8, [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
      expect(
          () => SpreadsheetDecoder.decodeBytes(bytes), throwsFormatException);
      var tiny =
          Uint8List.fromList([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
      expect(() => SpreadsheetDecoder.decodeBytes(tiny), throwsFormatException);
    });

    test('truncated files never fail with anything but FormatException', () {
      var original = _readFile('types.xls');
      for (var length = 8; length < original.length; length += 331) {
        try {
          SpreadsheetDecoder.decodeBytes(original.sublist(0, length));
        } on FormatException {
          // expected
        }
      }
    });

    test('random corruption only fails with FormatException/UnsupportedError',
        () {
      var random = Random(20240607);
      for (var name in ['types.xls', 'test.xls']) {
        var original = _readFile(name);
        for (var round = 0; round < 400; round++) {
          var bytes = Uint8List.fromList(original);
          for (var i = 0; i < 1 + random.nextInt(8); i++) {
            bytes[random.nextInt(bytes.length)] = random.nextInt(256);
          }
          try {
            SpreadsheetDecoder.decodeBytes(bytes);
          } on FormatException {
            // expected
          } on UnsupportedError {
            // expected
          }
        }
      }
    });

    test('random corruption of record data only', () {
      // Corrupt inside the workbook stream rather than the container, so the
      // BIFF parser itself is fuzzed.
      var random = Random(7);
      var stream = workbookStream([
        SheetSpec('S', [
          number(0, 0, 1),
          rk(0, 1, _rkInt(4)),
          mulRk(1, 0, [_rkInt(1), _rkInt(2)]),
          labelSst(2, 0, 0),
          label(2, 1, xlString('x')),
          formulaString(3, 0, xlString('f')),
          boolErr(4, 0, 1),
        ])
      ], globals: [
        sst(['a', 'b']),
        xf(14),
        format(164, 'yyyy'),
      ]);
      for (var round = 0; round < 3000; round++) {
        var bytes = Uint8List.fromList(stream);
        for (var i = 0; i < 1 + random.nextInt(4); i++) {
          bytes[random.nextInt(bytes.length)] = random.nextInt(256);
        }
        try {
          SpreadsheetDecoder.decodeBytes(compoundFile({'Workbook': bytes}));
        } on FormatException {
          // expected
        } on UnsupportedError {
          // expected
        }
      }
    });
  });
}
