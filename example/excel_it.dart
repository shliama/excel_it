import 'dart:io';
import 'package:path/path.dart';
import 'package:excel_it/excel_it.dart';

void main(List<String> args) {
  var decoder = ExcelIt.createExcel();

  /**
   * Create new Excel Sheet
   * var decoder = ExcelIt.createExcel();
   * 
   * ------------ Or ------------
   * For Editing Pre-Existing Excel File
   * 
   * var file = "Path_to_pre_existing_Excel_File/NewExcel.xlsx";
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

  /**
   * Define Your own sheet name:
   * var sheet = 'SheetName'
   * 
   * ---------- Or ----------
   * 
   * Find the desired sheet by iterating throught the [existing sheets]:
   * var sheet;
   * for (var tableName in decoder.tables.keys) {
   *    if( desiredSheetName.toString() == tableName.toString() ){
   *      sheet = tableName.toString();
   *      break;
   *    }
   * }
   */

  // if [MySheetName] does not exist then it will be automatically created.
  var sheet = 'MySheetName';

  decoder
    ..updateCell(sheet, 0, 0, "Font RED")
    ..updateCell(sheet, 2, 0, "Font BLUE")
    ..updateCell(sheet, 0, 1, "Font GREEN")
    ..updateCell(sheet, 4, 4, "Font Orange");

  File(join("Path_to_Excel_File/ExcelFileName.xlsx"))
    ..createSync(recursive: true)
    ..writeAsBytesSync(decoder.encode());

  print(
      "\n****************************** Printing Updated Data ******************************\n");
  for (var table in decoder.tables.keys) {
    print("Table Name:-" + table);
    print("Max Columns:-" + decoder.tables[table].maxCols.toString());
    print("Max Rows:-" + decoder.tables[table].maxRows.toString());
    print("Data in Table:\n");
    for (var row in decoder.tables[table].rows) {
      print("$row");
    }
    print("\n******************************\n");
  }
}
