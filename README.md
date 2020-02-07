# Excelit

Excelit is a library for decoding and updating spreadsheets for XLSX files.

## Usage

### In Flutter App

    import 'dart:io';
    import 'package:path/path.dart';
    import 'package:excel_it/excel_it.dart';

    ...
    
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
     
    for (var table in decoder.tables.keys) {
      print(table);
      print(decoder.tables[table].maxCols);
      print(decoder.tables[table].maxRows);
      for (var row in decoder.tables[table].rows) {
        print("$row");
      }
    }

    /**
     * Define Your own sheet name
     * var sheet = 'SheetName'
     * 
     * ---------- Or ----------
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
    
    // Save the file

    File(join("Path_to_Excel_File/ExcelFileName.xlsx"))
      ..createSync(recursive: true)
      ..writeAsBytesSync(decoder.encode());
    
    ...

## Features coming in next version
On-going implementation for future:
- annotations
- spanned rows (Comming Soon in future updates)
- spanned columns (Comming Soon in future updates)
- font colour (Comming Soon in future updates)

## Important:
For XLSX format, this implementation only supports native Excel format for date, time and boolean type conversion.
In other words, custom format for date, time, boolean aren't supported and also the files exported from LibreOffice as well.