part of excel_it;

const String _relationships =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const _spreasheetXlsx = 'xlsx';

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

/// Convert a number to character based column
String numericToLetters(int number) {
  var letters = '';

  while (number != 0) {
    // Set remainder from 1..26
    var remainder = number % 26;

    if (remainder == 0) {
      remainder = 26;
    }

    // Convert the remainder to a character.
    var letter = String.fromCharCode(65 + remainder - 1);

    // Accumulate the column letters, right to left.
    letters = letter + letters;

    // Get the next order of magnitude.
    number = (number - 1) ~/ 26;
  }
  return letters;
}

/// Decode a excel file.
abstract class ExcelIt {
  bool _update;
  Archive _archive;
  Map<String, XmlNode> _sheets;
  Map<String, XmlDocument> _xmlFiles;
  Map<String, ArchiveFile> _archiveFiles;
  Map<String, String> _worksheetTargets;
  Map<String, Map<String, List<String>>> _colorMap;
  Map<String, List<String>> _patternFill;
  Map<String, List<int>> _cellXfs;
  List<String> _sharedStrings, _rId, _fontColorHex;
  List<int> _numFormats;

  Map<String, SpreadsheetTable> _tables;
  String _stylesTarget, _sharedStringsTarget;

  /// Media type
  String get mediaType;

  /// Filename extension
  String get extension;

  /// Tables contained in spreadsheet file indexed by their names
  Map<String, SpreadsheetTable> get tables => _tables;

  ExcelIt();

  factory ExcelIt.createExcel() {
    String newSheet =
        "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";
    return ExcelIt.decodeBytes(Base64Decoder().convert(newSheet), update: true);
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

  void _damagedExcel() => throw ArgumentError("\nDamaged Excel file\n");

  /// Uses the [newSheet] as the name of the sheet and also adds it to the [ xl/worksheets/ ] directory
  /// Add the sheet details in the workbook.xml. as well as in the workbook.xml.rels
  /// Then add the sheet physically into the [_xmlFiles] so as to get it into the archieve.
  /// Also add it into the [_sheets] and [_tables] map so as to allow the editing.
  void _createSheet(String newSheet) {
    List<XmlNode> list =
        _xmlFiles["xl/workbook.xml"].findAllElements('sheets').first.children;
    if (list.length < 1) {
      throw ArgumentError("");
    }

    XmlElement lastSheet = list.last;

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
          XmlAttribute(XmlName('Id'), 'rId$ridNumber'),
          XmlAttribute(XmlName('Type'), "$_relationships/worksheet"),
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
            XmlAttribute(XmlName('r:id'), 'rId$ridNumber')
          ],
        ));

    _worksheetTargets['rId$ridNumber'] =
        "worksheets/sheet${sheetNumber + 1}.xml";

    _xmlFiles["xl/worksheets/sheet${sheetNumber + 1}.xml"] =
        _xmlFiles["xl/worksheets/sheet$sheetNumber.xml"];

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

  _parseStyles(String _stylesTarget) {
    var styles = _archive.findFile('xl/$_stylesTarget');
    if (styles != null) {
      styles.decompress();
      var document = parse(utf8.decode(styles.content));
      if (_xmlFiles != null) _xmlFiles["xl/$_stylesTarget"] = document;
      document
          .findAllElements('cellXfs')
          .first
          .findElements('xf')
          .forEach((node) {
        var numFmtId = node.getAttribute('numFmtId');
        if (numFmtId != null)
          _numFormats.add(int.parse(numFmtId));
        else
          _numFormats.add(0);
      });
    } else
      _damagedExcel();
  }

  /// Sets/Updates the Font Color in [xl/styles.xml] from the Cells of the sheets
  _setFontColors() {
    _colorMap.forEach((key, innerMap) => innerMap.forEach((keyIn, color) {
          if (color[0] != null &&
              color[0] != "FF000000" &&
              !_fontColorHex.contains(color[0])) _fontColorHex.add(color[0]);

          color[1] ??= color[2];
          color[2] ??= color[1];

          if (color[1] != null && color[2] != null) {
            String c = '${color[1]}-${color[2]}';
            if (!_patternFill.containsKey(c))
              _patternFill[c] = [color[1], color[2]];
          }
        }));

    _fontColorHex.removeWhere((value) => value == "FF000000");

    XmlElement fonts =
        _xmlFiles["xl/styles.xml"].findAllElements('fonts').first;
    fonts.getAttributeNode("count").value = "${_fontColorHex.length + 1}";

    fonts.children
      ..clear()
      ..addAll([
        XmlElement(XmlName("font"), [], [
          XmlElement(
              XmlName("color"), [XmlAttribute(XmlName("rgb"), "FF000000")], [])
        ])
      ]);

    _fontColorHex.forEach((colorValue) =>
        fonts.children.add(XmlElement(XmlName("font"), [], [
          XmlElement(
              XmlName("color"), [XmlAttribute(XmlName("rgb"), colorValue)], [])
        ])));
  }

  _setPatternFillColor() {
    XmlElement fills =
        _xmlFiles["xl/styles.xml"].findAllElements('fills').first;
    if (fills.getAttributeNode("count") != null)
      fills.getAttributeNode("count").value = "${_patternFill.keys.length + 1}";
    else
      fills.attributes.add(
          XmlAttribute(XmlName('count'), '${_patternFill.keys.length + 1}'));

    fills.children
      ..clear()
      ..addAll([
        XmlElement(XmlName("fill"), [], [
          XmlElement(XmlName("patternFill"),
              [XmlAttribute(XmlName("patternType"), "none")], [])
        ])
      ]);

    _patternFill.forEach((key, index) {
      fills.children.add(XmlElement(XmlName("fill"), [], [
        XmlElement(XmlName("patternFill"), [
          XmlAttribute(XmlName("patternType"), "solid")
        ], [
          XmlElement(
              XmlName("fgColor"), [XmlAttribute(XmlName("rgb"), index[0])], []),
          XmlElement(
              XmlName("bgColor"), [XmlAttribute(XmlName("rgb"), index[1])], [])
        ])
      ]));
    });
  }

  _setCellXfs() {
    _colorMap.forEach((key, innerMap) {
      innerMap.forEach((keyIn, color) {
        if (!_cellXfs.containsKey(color.toString())) {
          String c = '${color[1]}-${color[2]}';
          _cellXfs[color.toString()] = [
            _fontColorHex.indexOf(color[0].toString()) + 1,
            _patternFill.containsKey(c)
                ? _patternFill.keys.toList().indexOf(c) + 1
                : 0
          ];
        }
      });
    });

    XmlElement celx =
        _xmlFiles["xl/styles.xml"].findAllElements('cellXfs').first;
    celx.getAttributeNode("count").value = "${_cellXfs.keys.length + 1}";

    celx.children
      ..clear()
      ..addAll([
        XmlElement(XmlName("xf"), [
          XmlAttribute(XmlName("borderId"), "0"),
          XmlAttribute(XmlName("fillId"), "0"),
          XmlAttribute(XmlName("fontId"), "0"),
          XmlAttribute(XmlName("numFmtId"), "0"),
          XmlAttribute(XmlName("xfId"), "0"),
          XmlAttribute(XmlName("applyFill"), "1"),
          XmlAttribute(XmlName("applyFont"), "1")
        ], [])
      ]);

    _cellXfs.values.forEach((value) {
      celx.children.add(XmlElement(XmlName("xf"), [
        XmlAttribute(XmlName("borderId"), "0"),
        XmlAttribute(XmlName("fillId"), "${value[1]}"),
        XmlAttribute(XmlName("fontId"), "${value[0]}"),
        XmlAttribute(XmlName("numFmtId"), "0"),
        XmlAttribute(XmlName("xfId"), "0"),
        XmlAttribute(XmlName("applyFill"), "${value[1] == 0 ? 0 : 1}"),
        XmlAttribute(XmlName("applyFont"), "${value[0] == 0 ? 0 : 1}")
      ], []));
    });
  }

  _setSharedStrings() {
    String count = _sharedStrings.length.toString();
    List uniqueList = _sharedStrings.toSet().toList();
    String uniqueCount = uniqueList.length.toString();

    XmlElement shareString =
        _xmlFiles["xl/$_sharedStringsTarget"].findAllElements("sst").first;

    if (shareString.getAttributeNode("count") == null)
      shareString.attributes.add(XmlAttribute(XmlName("count"), count));
    else
      shareString.getAttributeNode("count").value = count;

    if (shareString.getAttributeNode("uniqueCount") == null)
      shareString.attributes
          .add(XmlAttribute(XmlName("uniqueCount"), uniqueCount));
    else
      shareString.getAttributeNode("uniqueCount").value = uniqueCount;

    shareString.children.clear();

    _sharedStrings.forEach((string) {
      shareString.children.add(XmlElement(XmlName("si"), [], [
        XmlElement(XmlName("t"), [], [XmlText(string)])
      ]));
    });
  }

  _updateSheetElements() {
    _tables.forEach((sheet, table) {
      for (int rowIndex = 0; rowIndex < table.rows.length; rowIndex++) {
        for (int columnIndex = 0;
            columnIndex < table.rows[rowIndex].length;
            columnIndex++) {
          if (table.rows[rowIndex][columnIndex] != null) {
            var foundRow = _findRowByIndex(_sheets[sheet], rowIndex);
            _updateCell(sheet, foundRow, columnIndex, rowIndex,
                table.rows[rowIndex][columnIndex]);
          }
        }
      }
    });
  }

  /// Dump XML content (for debug purpose)
  String dumpXmlContent([String sheet]);

  void _checkSheetArguments(String sheet) {
    if (!_update)
      throw ArgumentError("'update' should be set to 'true' on constructor");
    if (!_sheets.containsKey(sheet)) _createSheet(sheet);
  }

  /// Insert column in [sheet] at position [columnIndex]
  void insertColumn(String sheet, int columnIndex) {
    _checkSheetArguments(sheet);
    if (columnIndex < 0 /* || columnIndex > _tables[sheet]._maxCols  */)
      throw RangeError.range(columnIndex, 0, _tables[sheet]._maxCols);

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
    if (!_update)
      throw ArgumentError("'update' should be set to 'true' on constructor");
    if (!_sheets.containsKey(sheet) || columnIndex >= _tables[sheet]._maxCols)
      return;

    if (columnIndex < 0)
      throw RangeError.range(columnIndex, 0, _tables[sheet]._maxCols - 1);

    var table = _tables[sheet];
    table.rows.forEach((row) => row.removeAt(columnIndex));
    table._maxCols--;
  }

  /// Insert row in [sheet] at position [rowIndex]
  void insertRow(String sheet, int rowIndex) {
    _checkSheetArguments(sheet);
    if (rowIndex < 0)
      throw RangeError.range(rowIndex, 0, _tables[sheet]._maxRows);

    var table = _tables[sheet];
    if (rowIndex >= _tables[sheet]._maxRows)
      while (_tables[sheet]._maxRows <= rowIndex) {
        table.rows.insert(_tables[sheet]._maxRows,
            List.generate(table._maxCols, (_) => null));
        table._maxRows++;
      }
    else {
      table.rows.insert(rowIndex, List.generate(table._maxCols, (_) => null));
      table._maxRows++;
    }
  }

  /// Remove row in [sheet] at position [rowIndex]
  void removeRow(String sheet, int rowIndex) {
    if (!_update)
      throw ArgumentError("'update' should be set to 'true' on constructor");
    if (!_sheets.containsKey(sheet) || rowIndex >= _tables[sheet]._maxRows)
      return;
    if (rowIndex < 0)
      throw RangeError.range(rowIndex, 0, _tables[sheet]._maxRows - 1);

    var table = _tables[sheet];
    table.rows.removeAt(rowIndex);
    table._maxRows--;
  }

  /// Update the contents from [sheet] of the cell [columnIndex]x[rowIndex] with indexes start from 0
  /// [fontColorHex] or [backgroundColorHex] as example [#FF0000]
  void updateCell(String sheet, int columnIndex, int rowIndex, dynamic value,
      {String fontColorHex,
      String foregroundColorHex,
      String backgroundColorHex}) {
    _checkSheetArguments(sheet);

    if (columnIndex >= _tables[sheet]._maxCols)
      insertColumn(sheet, columnIndex);

    if (rowIndex >= _tables[sheet]._maxRows) insertRow(sheet, rowIndex);

    _tables[sheet].rows[rowIndex][columnIndex] = value.toString();
    _sharedStrings.add(value.toString());
  }

  /// Encode bytes after update
  List<int> encode() {
    if (!_update)
      throw ArgumentError("'update' should be set to 'true' on constructor");

    _setFontColors();
    _setPatternFillColor();
    _setSharedStrings();
    _setCellXfs();
    _updateSheetElements();

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
    buffer.write("data:$mediaType;base64,");
    buffer.write(base64Encode(encode()));
    return buffer.toString();
  }

  Archive _cloneArchive(Archive archive) {
    var clone = Archive();
    archive.files.forEach((file) {
      if (file.isFile) {
        ArchiveFile copy;
        if (_archiveFiles.containsKey(file.name))
          copy = _archiveFiles[file.name];
        else {
          var content = (file.content as Uint8List).toList();
          var compress = !_noCompression.contains(file.name);
          copy = ArchiveFile(file.name, content.length, content)
            ..compress = compress;
        }
        clone.addFile(copy);
      }
    });
    return clone;
  }

  _normalizeTable(SpreadsheetTable table) {
    if (table._maxRows == 0)
      table._rows.clear();
    else if (table._maxRows < table._rows.length)
      table._rows.removeRange(table._maxRows, table._rows.length);

    for (var row = 0; row < table._rows.length; row++) {
      if (table._maxCols == 0)
        table._rows[row].clear();
      else if (table._maxCols < table._rows[row].length)
        table._rows[row].removeRange(table._maxCols, table._rows[row].length);
      else if (table._maxCols > table._rows[row].length) {
        var repeat = table._maxCols - table._rows[row].length;
        for (var index = 0; index < repeat; index++) table._rows[row].add(null);
      }
    }
  }

  bool _isEmptyRow(List row) =>
      row.fold(true, (value, element) => value && (element == null));

  bool _isNotEmptyRow(List row) => !_isEmptyRow(row);

  _countFilledRow(SpreadsheetTable table, List row) {
    if (_isNotEmptyRow(row) && table._maxRows < table._rows.length)
      table._maxRows = table._rows.length;
  }

  _countFilledColumn(SpreadsheetTable table, List row, dynamic value) {
    if (value != null && table._maxCols < row.length)
      table._maxCols = row.length;
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

    _findRows(sheet).forEach((child) => _parseRow(child, table));

    if (_update) {
      _sheets[name] = sheet;
      _xmlFiles["xl/$target"] = content;
    }

    _normalizeTable(table);
  }

  _parseRow(XmlElement node, SpreadsheetTable table) {
    var row = List();

    _findCells(node).forEach((child) => _parseCell(child, table, row));

    var rowIndex = _getRowNumber(node) - 1;
    if (_isNotEmptyRow(row) && rowIndex > table._rows.length) {
      var repeat = rowIndex - table._rows.length;
      for (var index = 0; index < repeat; index++) table._rows.add(List());
    }

    if (_isNotEmptyRow(row))
      table._rows.add(row);
    else
      table._rows.add(List());

    _countFilledRow(table, row);
  }

  _parseCell(XmlElement node, SpreadsheetTable table, List row) {
    var colIndex = _getCellNumber(node) - 1;
    if (colIndex > row.length) {
      var repeat = colIndex - row.length;
      for (var index = 0; index < repeat; index++) row.add(null);
    }

    if (node.children.isEmpty) return;

    var value, type = node.getAttribute('t');

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
          } else
            value = num.parse(_parseValue(content));
        } else
          value = num.parse(_parseValue(content));
    }
    row.add(value);

    _countFilledColumn(table, row, value);
  }

  _parseValue(XmlElement node) {
    var buffer = StringBuffer();

    node.children.forEach((child) {
      if (child is XmlText) buffer.write(_normalizeNewLine(child.text));
    });

    return buffer.toString();
  }

  Iterable<XmlElement> _findRows(XmlElement table) => table.findElements('row');

  Iterable<XmlElement> _findCells(XmlElement row) => row.findElements('c');

  int _getCellNumber(XmlElement cell) =>
      cellCoordsFromCellId(cell.getAttribute('r'))[0];

  int _getRowNumber(XmlElement row) => int.parse(row.getAttribute('r'));

  Iterable<XmlNode> _getNodeValue() => <XmlElement>[
        XmlElement(XmlName('sheetData')),
        XmlElement(XmlName('drawing'),
            <XmlAttribute>[XmlAttribute(XmlName('r:id'), 'rId1')])
      ];

  XmlElement _findRowByIndex(XmlElement table, int rowIndex) {
    XmlElement row;
    var rows = _findRows(table);

    var currentIndex = 0;
    for (var currentRow in rows) {
      currentIndex = _getRowNumber(currentRow) - 1;
      if (currentIndex >= rowIndex) {
        row = currentRow;
        break;
      }
    }

    // Create row if required
    if (row == null || currentIndex != rowIndex)
      row = __insertRow(table, row, rowIndex);

    return row;
  }

  XmlElement _updateCell(String sheet, XmlElement node, int columnIndex,
      int rowIndex, dynamic value) {
    XmlElement cell;
    var cells = _findCells(node);

    var currentIndex = 0; // cells could be empty
    for (var currentCell in cells) {
      currentIndex = _getCellNumber(currentCell) - 1;
      if (currentIndex >= columnIndex) {
        cell = currentCell;
        break;
      }
    }

    if (cell == null || currentIndex != columnIndex)
      cell = _insertCell(sheet, node, cell, columnIndex, rowIndex, value);
    else
      cell = _replaceCell(sheet, node, cell, columnIndex, rowIndex, value);

    return cell;
  }

  XmlElement _createRow(int rowIndex) => XmlElement(XmlName('row'),
      [XmlAttribute(XmlName('r'), (rowIndex + 1).toString())], []);

  XmlElement __insertRow(XmlElement table, XmlElement lastRow, int rowIndex) {
    var row = _createRow(rowIndex);
    if (lastRow == null)
      table.children.add(row);
    else {
      var index = table.children.indexOf(lastRow);
      table.children.insert(index, row);
    }
    return row;
  }

  XmlElement _insertCell(String sheet, XmlElement row, XmlElement lastCell,
      int columnIndex, int rowIndex, dynamic value) {
    var cell = _createCell(sheet, columnIndex, rowIndex, value);
    if (lastCell == null)
      row.children.add(cell);
    else {
      var index = row.children.indexOf(lastCell);
      row.children.insert(index, cell);
    }
    return cell;
  }

  XmlElement _replaceCell(String sheet, XmlElement row, XmlElement lastCell,
      int columnIndex, int rowIndex, dynamic value) {
    var index = lastCell == null ? 0 : row.children.indexOf(lastCell);
    var cell = _createCell(sheet, columnIndex, rowIndex, value);
    row.children
      ..removeAt(index)
      ..insert(index, cell);
    return cell;
  }

  // Manage value's type
  XmlElement _createCell(
      String sheet, int columnIndex, int rowIndex, dynamic value) {
    if (!_sharedStrings.contains(value)) _sharedStrings.add(value);

    String rC = "${numericToLetters(columnIndex + 1)}${rowIndex + 1}";

    var attributes = <XmlAttribute>[
      XmlAttribute(XmlName('r'), rC),
      XmlAttribute(XmlName('t'), 's'),
    ];

    if (_colorMap.containsKey(sheet) && _colorMap[sheet].containsKey(rC)) {
      String color = _colorMap[sheet][rC].toString();

      attributes.insert(
        1,
        XmlAttribute(XmlName('s'),
            '${_cellXfs.containsKey(color) ? _cellXfs.keys.toList().indexOf(color) + 1 : 0}'),
      );
    }
    var children = value == null
        ? <XmlElement>[]
        : <XmlElement>[
            XmlElement(XmlName('v'), [],
                [XmlText(_sharedStrings.indexOf(value).toString())]),
          ];
    return XmlElement(XmlName('c'), attributes, children);
  }
}

/// Table of a spreadsheet file
class SpreadsheetTable {
  final String name;
  SpreadsheetTable(this.name);

  int _maxRows = 0, _maxCols = 0;

  List<List> _rows = List<List>();

  /// List of table's rows
  List<List> get rows => _rows;

  /// Get max rows
  int get maxRows => _maxRows;

  /// Get max cols
  int get maxCols => _maxCols;
}
