import 'dart:io';
import 'package:path/path.dart';
import 'package:excel_it/excel_it.dart';

//excel_it/test/files
main(List<String> args) {
  var file = "/Users/kawal/Desktop/test.xlsx";
  //var bytes = File(file).readAsBytesSync();
  var decoder = ExcelIt.createExcel();  //.decodeBytes(bytes, update: true);

  print(decoder.toString());
  /* for (var table in decoder.tables.keys) {
    print(table);
    print(decoder.tables[table].maxCols);
    print(decoder.tables[table].maxRows);
    for (var row in decoder.tables[table].rows) {
      print("$row");
    }
  }

  var sheet = 'Sheet1';
  decoder
    ..updateCell(sheet, 0, 0, "Font RED")
    ..updateCell(sheet, 2, 0, "Font BLUE")
    ..updateCell(sheet, 0, 1, "Font GREEN")
    ..updateCell(sheet, 4, 4, "Font Orange");

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
  } */
}
