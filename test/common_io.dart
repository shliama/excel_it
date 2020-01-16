import 'dart:convert';
import 'dart:io';
import 'package:excel_it/excel_it.dart';

List<int> _readBytes(String filename) {
  var fullUri = Uri.file('test/files/$filename');
  return File.fromUri(fullUri).readAsBytesSync();
}

String readBase64(String filename) {
  return base64Encode(_readBytes(filename));
}

ExcelIt decode(String filename, {bool update = false}) {
  return ExcelIt.decodeBytes(_readBytes(filename), update: update, verify: true);
}

void save(String file, List<int> data) {
  File(file)
    ..createSync(recursive: true)
    ..writeAsBytesSync(data);
}
