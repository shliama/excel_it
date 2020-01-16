# ExcelIt


ExcelIt is a library for decoding and updating spreadsheets for XLSX files.

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
- hidden rows (visible in resulting tables)
- hidden columns (visible in resulting tables)

For XLSX format, this implementation only supports native Excel format for date, time and boolean type conversion.
In other words, custom format for date, time, boolean aren't supported and then file exported from LibreOffice as well.

