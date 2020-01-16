import 'dart:io';
import 'package:excel_it/excel_it.dart';

void main(List<String> args) {
  var path = args[0];
  var sheet = args.length > 1 ? args[1] : null;
  var data = File(path).readAsBytesSync();
  var decoder = ExcelIt.decodeBytes(data, update: true);
  print(decoder.dumpXmlContent(sheet));
}
