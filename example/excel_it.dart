import 'dart:io';
import 'package:path/path.dart';
import 'package:excel_it/excel_it.dart';

main(List<String> args) {

  var decoder = ExcelIt.createExcel();

  /**
   * ------------ Or ------------
   * For Editing Pre-Existing Excel File
   * var file = "/Users/kawal/Desktop/new_excel/New.xlsx";
   * var bytes = File(file).readAsBytesSync();
   * var decoder = ExcelIt.decodeBytes(bytes, update: true);
   **/

  print(decoder.toString());
  for (var table in decoder.tables.keys) {
    print(table);
    print(decoder.tables[table].maxCols);
    print(decoder.tables[table].maxRows);
    for (var row in decoder.tables[table].rows) {
      print("$row");
    }
  }

  /*
    Define Your own sheet name:
    var sheet = 'SheetName'
    
    ---------- Or ----------

   Iterate throught the [existing sheets] by

   var sheet;
   for (var tableName in decoder.tables.keys) {
    if( desiredSheetName.toString() == tableName.toString() ){
      sheet = tableName.toString();
      break;
    }
  }
  */

  var sheet = 'Sheet';

  decoder
    ..updateCell(sheet, 0, 0, "Font RED")
    ..updateCell(sheet, 2, 0, "Font BLUE")
    ..updateCell(sheet, 0, 1, "Font GREEN")
    ..updateCell(sheet, 4, 4, "Font Orange");

  File(join("/Users/kawal/Desktop/New.xlsx"))
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
