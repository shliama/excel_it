import 'dart:io';
import 'package:path/path.dart';
import 'package:excel_it/excel_it.dart';

main(List<String> args) {
  var file = "/Users/kawal/Desktop/excel_it/test/files/test.xlsx";
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

  var sheet = decoder.tables.keys.first;
  decoder
    ..updateCell(sheet, 0, 0, "Ashumendra")
    ..updateCell(sheet, 1, 0, null)
    ..updateCell(sheet, 2, 0, "C")
    ..updateCell(sheet, 1, 1, 42.3)
    
    ..updateCell(sheet, 4, 3, "A14");

  File(join("/Users/kawal/Desktop/output/${basename(file)}"))
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
