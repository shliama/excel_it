import 'dart:io';
import 'package:path/path.dart';
import 'package:excel_it/excel_it.dart';

main(List<String> args) {
  var file = "/Users/kawal/Desktop/test.xlsx";
  var bytes = File(file).readAsBytesSync();
  var decoder = ExcelIt.decodeBytes(bytes, update: true);
  for (var table in decoder.tables.keys) {
    print(table);
    print(decoder.tables[table].maxCols);
    print(decoder.tables[table].maxRows);
    for (var row in decoder.tables[table].rows) {
      print("$row");
    }
  }

  var sheet = 'Sheet4';
  decoder
    ..updateCell(sheet, 0, 0, "A1")
    ..updateCell(sheet, 2, 0, "C1")
    ..updateCell(sheet, 0, 1, "A2")
    ..updateCell(sheet, 4, 4, "A5");

  File(join("/Users/kawal/Desktop/${basename(file)}"))
    ..createSync(recursive: true)
    ..writeAsBytesSync(decoder.encode());

  print("************************************************************");
  for (var table in decoder.tables.keys) {
    print("printint");
    print(table);
    print(decoder.tables[table].maxCols);
    print(decoder.tables[table].maxRows);
    for (var row in decoder.tables[table].rows) {
      print("$row");
    }
  }
}
