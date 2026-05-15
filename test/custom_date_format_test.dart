@TestOn('vm')
library spreadsheet_custom_date_format_test;

import 'package:test/test.dart';

import 'common_io.dart';

/// Tests that custom <numFmt> date format codes (numFmtId outside the
/// built-in 14-22/45-47 range) are recognised and decoded as formatted
/// date strings instead of leaking through as raw numbers.
void main() {
  group('Custom date number format (xlsx):', () {
    test('sample_date_format.xlsx — dd/mm/yyyy custom numFmt', () {
      var decoder = decode('sample_date_format.xlsx');
      var sheet = decoder.tables.values.first;

      // Cells L2, P2, R2 use the custom numFmtId=59 with
      // formatCode="dd/mm/yyyy" and raw serial values 39650, 45782, 46144.
      // Column L = index 11, P = 15, R = 17 (0-based).
      final row2 = sheet.rows[1];

      final ddmmyyyy = RegExp(r'^\d{2}/\d{2}/\d{4}$');

      expect(row2[11], isA<String>(),
          reason: 'L2 must decode as a date string, not a number');
      expect(row2[15], isA<String>(),
          reason: 'P2 must decode as a date string, not a number');
      expect(row2[17], isA<String>(),
          reason: 'R2 must decode as a date string, not a number');

      expect(row2[11] as String, matches(ddmmyyyy),
          reason: 'L2 should match the dd/mm/yyyy format code');
      expect(row2[15] as String, matches(ddmmyyyy));
      expect(row2[17] as String, matches(ddmmyyyy));

      // The exact day-of-month follows the existing decoder's serial-to-date
      // math (1899-12-30 baseline, the same convention used for the
      // built-in numFmtId 14-22 path). These are the deterministic outputs
      // for the three serials in this fixture.
      expect(row2[11], equals('21/07/2008'));
      expect(row2[15], equals('05/05/2025'));
      expect(row2[17], equals('02/05/2026'));

      // Sanity: none of the date cells should leak through as numbers.
      expect(row2[11], isNot(isA<num>()));
      expect(row2[15], isNot(isA<num>()));
      expect(row2[17], isNot(isA<num>()));
    });
  });
}
