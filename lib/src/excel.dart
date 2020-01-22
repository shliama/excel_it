part of excel_it;

const String _relationships =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const _spreasheetXlsx = 'xlsx';
final intRegex = RegExp(r'\s+(\d+)\s+', multiLine: true);
/* final Map<String, String> _spreasheetExtensionMap = <String, String>{
  _spreasheetXlsx:
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
}; */

// Normalize new line
String _normalizeNewLine(String text) {
  return text.replaceAll("\r\n", "\n");
}

ExcelIt _newExcelIt(Archive archive, bool update) {
  // Lookup at file format
  var format;

  // Try OpenDocument format
  var mimetype = archive.findFile('mimetype');
  if (mimetype == null) {
    var xl = archive.findFile('xl/workbook.xml');
    format = xl != null ? _spreasheetXlsx : null;
  }

  switch (format) {
    case _spreasheetXlsx:
      return XlsxDecoder(archive, update: update);
    default:
      throw UnsupportedError("Excel format unsupported");
  }
}

const List<String> _noCompression = const <String>[
  'mimetype',
  'Thumbnails/thumbnail.png',
];

/**
 * Decode a excel file.
 */
abstract class ExcelIt {
  bool _update;
  Archive _archive;
  Map<String, XmlNode> _sheets;
  Map<String, XmlDocument> _xmlFiles;
  Map<String, ArchiveFile> _archiveFiles;
  Map<String, String> _worksheetTargets;
  List<String> _sharedStrings, _rId;
  List<int> _numFormats;

  Map<String, SpreadsheetTable> _tables;

  /// Media type
  String get mediaType;

  /// Filename extension
  String get extension;

  /// Tables contained in spreadshet file indexed by their names
  Map<String, SpreadsheetTable> get tables => _tables;

  ExcelIt();

  factory ExcelIt.createExcel() {
    print("Directory:\n" + Directory.current.toString());
    List<int> data = File(Directory.current.toString()).readAsBytesSync();
    ;
    var archive = ZipDecoder().decodeBytes(data, verify: false);
    return _newExcelIt(archive, true);
  }
  factory ExcelIt.decodeBytes(List<int> data,
      {bool update: false, bool verify: false}) {
    var archive = ZipDecoder().decodeBytes(data, verify: verify);
    return _newExcelIt(archive, update);
  }

  factory ExcelIt.decodeBuffer(InputStream input,
      {bool update: false, bool verify: false}) {
    var archive = ZipDecoder().decodeBuffer(input, verify: verify);
    return _newExcelIt(archive, update);
  }

  /**
   * Uses the [newSheet] as the name of the sheet and also adds it to the [ xl/worksheets/ ] directory
   * Add the sheet details in the workbook.xml. as well as in the workbook.xml.rels
   * Then add the sheet physically into the [_xmlFiles] so as to get it into the archieve.
   * Also add it into the [_sheets] and [_tables] map so as to allow the editing.
   */
  void _createSheet(String newSheet) {
    XmlElement lastSheet = _xmlFiles["xl/workbook.xml"]
        .findAllElements('sheets')
        .first
        .children
        .last;
    int sheetNumber = int.parse(lastSheet.getAttribute('sheetId'));
    _rId.sort((a, b) =>
        int.parse(a.substring(3)).compareTo(int.parse(b.substring(3))));
    List<String> got = new List<String>.from(_rId.last.split(''));
    got.removeWhere((item) => !'0123456789'.split('').contains(item));
    int ridNumber = int.parse(got.join().toString()) + 1;

    _xmlFiles["xl/_rels/workbook.xml.rels"]
        .findAllElements('Relationships')
        .first
        .children
        .add(XmlElement(XmlName('Relationship'), <XmlAttribute>[
          XmlAttribute(XmlName('Id'), 'rId${ridNumber}'),
          XmlAttribute(XmlName('Type'), "${_relationships}/worksheet"),
          XmlAttribute(
              XmlName('Target'), 'worksheets/sheet${sheetNumber + 1}.xml'),
        ]));

    _xmlFiles["xl/workbook.xml"]
        .findAllElements('sheets')
        .first
        .children
        .add(XmlElement(
          XmlName('sheet'),
          <XmlAttribute>[
            XmlAttribute(XmlName('state'), 'visible'),
            XmlAttribute(XmlName('name'), newSheet),
            XmlAttribute(XmlName('sheetId'), '${sheetNumber + 1}'),
            XmlAttribute(XmlName('r:id'), 'rId${ridNumber}')
          ],
        ));

    _worksheetTargets['rId${ridNumber}'] =
        "worksheets/sheet${sheetNumber + 1}.xml";

    _xmlFiles["xl/worksheets/sheet${sheetNumber + 1}.xml"] =
        _xmlFiles["xl/worksheets/sheet${sheetNumber}.xml"];

    _xmlFiles["xl/worksheets/sheet${sheetNumber + 1}.xml"]
        .findElements('worksheet')
        .first
        .children
      ..clear()
      ..addAll(_getNodeValue());

    var content = utf8.encode(
        _xmlFiles["xl/worksheets/sheet${sheetNumber + 1}.xml"].toString());

    _archive.addFile(ArchiveFile(
        'xl/worksheets/sheet${sheetNumber + 1}.xml', content.length, content));

    var file = _archive.findFile("[Content_Types].xml");
    file.decompress();

    _xmlFiles["[Content_Types].xml"] = parse(utf8.decode(file.content));

    _xmlFiles["[Content_Types].xml"]
        .findAllElements('Types')
        .first
        .children
        .add(XmlElement(
          XmlName('Override'),
          <XmlAttribute>[
            XmlAttribute(XmlName('ContentType'),
                'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
            XmlAttribute(XmlName('PartName'),
                '/xl/worksheets/sheet${sheetNumber + 1}.xml'),
          ],
        ));
    _parseTable(_xmlFiles["xl/workbook.xml"].findAllElements('sheet').last);
  }

  /// Dump XML content (for debug purpose)
  String dumpXmlContent([String sheet]);

  void _checkSheetArguments(String sheet) {
    if (_update != true)
      throw ArgumentError("'update' should be set to 'true' on constructor");
    if (_sheets.containsKey(sheet) == false) _createSheet(sheet);
  }

  /// Insert column in [sheet] at position [columnIndex]
  void insertColumn(String sheet, int columnIndex) {
    _checkSheetArguments(sheet);
    if (columnIndex < 0 /* || columnIndex > _tables[sheet]._maxCols  */) {
      throw RangeError.range(columnIndex, 0, _tables[sheet]._maxCols);
    }

    var table = _tables[sheet];
    int columnLength = _tables[sheet]._maxCols;
    if (columnIndex >= columnLength) {
      table.rows.forEach((row) {
        int len = columnLength;
        while (len <= columnIndex) {
          row.insert(len, null);
          len++;
        }
      });
      table._maxCols += columnIndex - columnLength + 1;
    } else {
      table.rows.forEach((row) => row.insert(columnIndex, null));
      table._maxCols++;
    }
  }

  /// Remove column in [sheet] at position [columnIndex]
  void removeColumn(String sheet, int columnIndex) {
    _checkSheetArguments(sheet);
    if (columnIndex < 0 || columnIndex >= _tables[sheet]._maxCols) {
      throw RangeError.range(columnIndex, 0, _tables[sheet]._maxCols - 1);
    }

    var table = _tables[sheet];
    table.rows.forEach((row) => row.removeAt(columnIndex));
    table._maxCols--;
  }

  /// Insert row in [sheet] at position [rowIndex]
  void insertRow(String sheet, int rowIndex) {
    _checkSheetArguments(sheet);
    if (rowIndex < 0 /* || rowIndex > _tables[sheet]._maxRows */) {
      throw RangeError.range(rowIndex, 0, _tables[sheet]._maxRows);
    }

    var table = _tables[sheet];
    if (rowIndex >= _tables[sheet]._maxRows) {
      while (_tables[sheet]._maxRows <= rowIndex) {
        table.rows.insert(_tables[sheet]._maxRows,
            List.generate(table._maxCols, (_) => null));
        table._maxRows++;
      }
    } else {
      table.rows.insert(rowIndex, List.generate(table._maxCols, (_) => null));
      table._maxRows++;
    }
  }

  /// Remove row in [sheet] at position [rowIndex]
  void removeRow(String sheet, int rowIndex) {
    _checkSheetArguments(sheet);
    if (rowIndex < 0 || rowIndex >= _tables[sheet]._maxRows) {
      throw RangeError.range(rowIndex, 0, _tables[sheet]._maxRows - 1);
    }

    var table = _tables[sheet];
    table.rows.removeAt(rowIndex);
    table._maxRows--;
  }

  /// Update the contents from [sheet] of the cell [columnIndex]x[rowIndex] with indexes start from 0
  void updateCell(String sheet, int columnIndex, int rowIndex, dynamic value) {
    _checkSheetArguments(sheet);

    if (columnIndex >= _tables[sheet]._maxCols)
      insertColumn(sheet, columnIndex);

    if (rowIndex >= _tables[sheet]._maxRows) insertRow(sheet, rowIndex);

    _tables[sheet].rows[rowIndex][columnIndex] = value.toString();
  }

  /// Encode bytes after update
  List<int> encode() {
    if (_update != true) {
      throw ArgumentError("'update' should be set to 'true' on constructor");
    }

    for (var xmlFile in _xmlFiles.keys) {
      var xml = _xmlFiles[xmlFile].toString();
      var content = utf8.encode(xml);
      _archiveFiles[xmlFile] = ArchiveFile(xmlFile, content.length, content);
    }
    return ZipEncoder().encode(_cloneArchive(_archive));
  }

  /// Encode data url
  String dataUrl() {
    var buffer = StringBuffer();
    buffer.write("data:${mediaType};base64,");
    buffer.write(base64Encode(encode()));
    return buffer.toString();
  }

  Archive _cloneArchive(Archive archive) {
    var clone = Archive();
    archive.files.forEach((file) {
      if (file.isFile) {
        ArchiveFile copy;
        if (_archiveFiles.containsKey(file.name)) {
          copy = _archiveFiles[file.name];
        } else {
          var content = (file.content as Uint8List).toList();
          //var compress = file.compress;
          var compress = _noCompression.contains(file.name) ? false : true;
          copy = ArchiveFile(file.name, content.length, content)
            ..compress = compress;
        }
        clone.addFile(copy);
      }
    });
    return clone;
  }

  _normalizeTable(SpreadsheetTable table) {
    if (table._maxRows == 0) {
      table._rows.clear();
    } else if (table._maxRows < table._rows.length) {
      table._rows.removeRange(table._maxRows, table._rows.length);
    }
    for (var row = 0; row < table._rows.length; row++) {
      if (table._maxCols == 0) {
        table._rows[row].clear();
      } else if (table._maxCols < table._rows[row].length) {
        table._rows[row].removeRange(table._maxCols, table._rows[row].length);
      } else if (table._maxCols > table._rows[row].length) {
        var repeat = table._maxCols - table._rows[row].length;
        for (var index = 0; index < repeat; index++) {
          table._rows[row].add(null);
        }
      }
    }
  }

  bool _isEmptyRow(List row) {
    return row.fold(true, (value, element) => value && (element == null));
  }

  bool _isNotEmptyRow(List row) {
    return !_isEmptyRow(row);
  }

  _countFilledRow(SpreadsheetTable table, List row) {
    if (_isNotEmptyRow(row)) {
      if (table._maxRows < table._rows.length) {
        table._maxRows = table._rows.length;
      }
    }
  }

  _countFilledColumn(SpreadsheetTable table, List row, dynamic value) {
    if (value != null) {
      if (table._maxCols < row.length) {
        table._maxCols = row.length;
      }
    }
  }

  _parseTable(XmlElement node) {
    var name = node.getAttribute('name');
    var target = _worksheetTargets[node.getAttribute('r:id')];

    tables[name] = SpreadsheetTable(name);
    var table = tables[name];

    var file = _archive.findFile("xl/$target");
    file.decompress();

    var content = parse(utf8.decode(file.content));
    var worksheet = content.findElements('worksheet').first;
    var sheet = worksheet.findElements('sheetData').first;

    _findRows(sheet).forEach((child) {
      _parseRow(child, table);
    });
    if (_update == true) {
      _sheets[name] = sheet;
      _xmlFiles["xl/$target"] = content;
    }

    _normalizeTable(table);
  }

  _parseRow(XmlElement node, SpreadsheetTable table) {
    var row = List();

    _findCells(node).forEach((child) {
      _parseCell(child, table, row);
    });

    var rowIndex = _getRowNumber(node) - 1;
    if (_isNotEmptyRow(row) && rowIndex > table._rows.length) {
      var repeat = rowIndex - table._rows.length;
      for (var index = 0; index < repeat; index++) {
        table._rows.add(List());
      }
    }

    if (_isNotEmptyRow(row)) {
      table._rows.add(row);
    } else {
      table._rows.add(List());
    }

    _countFilledRow(table, row);
  }

  _parseCell(XmlElement node, SpreadsheetTable table, List row) {
    var colIndex = _getCellNumber(node) - 1;
    if (colIndex > row.length) {
      var repeat = colIndex - row.length;
      for (var index = 0; index < repeat; index++) {
        row.add(null);
      }
    }

    if (node.children.isEmpty) {
      return;
    }

    var value;
    var type = node.getAttribute('t');

    switch (type) {
      // sharedString
      case 's':
        value = _sharedStrings[
            int.parse(_parseValue(node.findElements('v').first))];
        break;
      // boolean
      case 'b':
        value = _parseValue(node.findElements('v').first) == '1';
        break;
      // error
      case 'e':
      // formula
      case 'str':
        // <c r="C6" s="1" vm="15" t="str">
        //  <f>CUBEVALUE("xlextdat9 Adventure Works",C$5,$A6)</f>
        //  <v>2838512.355</v>
        // </c>
        value = _parseValue(node.findElements('v').first);
        break;
      // inline string
      case 'inlineStr':
        // <c r="B2" t="inlineStr">
        // <is><t>Hello world</t></is>
        // </c>
        value = _parseValue(node.findAllElements('t').first);
        break;
      // number
      case 'n':
      default:
        var s = node.getAttribute('s');
        var valueNode = node.findElements('v');
        var content = valueNode.first;
        if (s != null) {
          var fmtId = _numFormats[int.parse(s)];
          // date
          if (((fmtId >= 14) && (fmtId <= 17)) || (fmtId == 22)) {
            var delta = num.parse(_parseValue(content)) * 24 * 3600 * 1000;
            var date = DateTime(1899, 12, 30);
            value = date
                .add(Duration(milliseconds: delta.toInt()))
                .toIso8601String();
            // time
          } else if (((fmtId >= 18) && (fmtId <= 21)) ||
              ((fmtId >= 45) && (fmtId <= 47))) {
            var delta = num.parse(_parseValue(content)) * 24 * 3600 * 1000;
            var date = DateTime(0);
            date = date.add(Duration(milliseconds: delta.toInt()));
            value =
                "${_twoDigits(date.hour)}:${_twoDigits(date.minute)}:${_twoDigits(date.second)}";
            // number
          } else {
            value = num.parse(_parseValue(content));
          }
        } else {
          value = num.parse(_parseValue(content));
        }
    }
    row.add(value);

    _countFilledColumn(table, row, value);
  }

  _parseValue(XmlElement node) {
    var buffer = StringBuffer();

    node.children.forEach((child) {
      if (child is XmlText) {
        buffer.write(_normalizeNewLine(child.text));
      }
    });

    return buffer.toString();
  }

  Iterable<XmlElement> _findRows(XmlElement table) => table.findElements('row');

  Iterable<XmlElement> _findCells(XmlElement row) => row.findElements('c');

  int _getCellNumber(XmlElement cell) =>
      cellCoordsFromCellId(cell.getAttribute('r'))[0];

  int _getRowNumber(XmlElement row) => int.parse(row.getAttribute('r'));

  Iterable<XmlNode> _getNodeValue() => <XmlElement>[
        XmlElement(XmlName('sheetPr'), [], <XmlElement>[
          XmlElement(XmlName('outlinePr'), <XmlAttribute>[
            XmlAttribute(XmlName('summaryBelow'), '0'),
            XmlAttribute(XmlName('summaryRight'), '0')
          ])
        ]),
        XmlElement(XmlName('sheetViews'), <XmlAttribute>[], <XmlElement>[
          XmlElement(XmlName('sheetView'),
              <XmlAttribute>[XmlAttribute(XmlName('workbookViewId'), '0')])
        ]),
        XmlElement(XmlName('sheetFormatPr'), <XmlAttribute>[
          XmlAttribute(XmlName('customHeight'), '1'),
          XmlAttribute(XmlName('defaultColWidth'), '14.43'),
          XmlAttribute(XmlName('defaultRowHeight'), '15.0')
        ]),
        XmlElement(XmlName('sheetData')),
        XmlElement(XmlName('drawing'),
            <XmlAttribute>[XmlAttribute(XmlName('r:id'), 'rId1')])
      ];
}

/// Table of a spreadsheet file
class SpreadsheetTable {
  final String name;
  SpreadsheetTable(this.name);

  int _maxRows = 0;
  int _maxCols = 0;

  List<List> _rows = List<List>();

  /// List of table's rows
  List<List> get rows => _rows;

  /// Get max rows
  int get maxRows => _maxRows;

  /// Get max cols
  int get maxCols => _maxCols;
}