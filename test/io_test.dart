@TestOn('vm')

library spreadsheet_test;

import 'dart:convert';
import 'package:excel_it/excel_it.dart';
import 'package:test/test.dart';

import 'common_io.dart';
part 'common.dart';

void main() {
  testUnsupported();
  testXlsx();
  testUpdateOds();
  testUpdateXlsx();
}
