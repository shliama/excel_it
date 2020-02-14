import 'dart:io';
import 'package:path/path.dart';
import 'package:excel_it/excel_it.dart';

void main(List<String> args) {
  var decoder = ExcelIt.createExcel();

  for (var table in decoder.tables.keys) {
    print(table);
    print(decoder.tables[table].maxCols);
    print(decoder.tables[table].maxRows);
    for (var row in decoder.tables[table].rows) {
      print("$row");
    }
  }

  // if [Sheet2] does not exist then it will be automatically created.
  var sheet = 'Sheet2';

  decoder
    ..updateCell(sheet, 0, 0, "It is Verically top aligned",
        fontColorHex: "#1AFF1A", verticalAlign: VerticalAlign.Top)
    ..updateCell(
        sheet, 2, 0, "It is a Wrapped Text jdfhgdfjhgdkfjhdkfgjhfdgkjh",
        wrap: TextWrapping.WrapText)
    ..updateCell(sheet, 0, 1, "It is Clip dfhgldhlfdflh  fdjhdfkjgdfkjgh",
        wrap: TextWrapping.Clip, backgroundColorHex: "#1AFF1A")
    ..updateCell(sheet, 4, 4, "It is Aligned Right ldfhgldfhgh",
        horizontalAlign: HorizontalAlign.Right, backgroundColorHex: "#112E9C");

  decoder.encode().then((onValue) {
    File(join("/Users/kawal/Desktop/excel.xlsx"))
      ..createSync(recursive: true)
      ..writeAsBytesSync(onValue);
  });

  print(
      "\n****************************** Printing Updated Data ******************************\n");
  for (var table in decoder.tables.keys) {
    print("Table Name   :" + table);
    print("Max Columns  :" + decoder.tables[table].maxCols.toString());
    print("Max Rows     :" + decoder.tables[table].maxRows.toString());
    print("Data in Table:\n");
    for (var row in decoder.tables[table].rows) {
      print("$row");
    }
    print("\n******************************\n");
  }
}
