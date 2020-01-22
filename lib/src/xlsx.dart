part of excel_it;

const String _relationshipsStyles =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const String _relationshipsWorksheet =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const String _relationshipsSharedStrings =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

/// Convert a character based column
int lettersToNumeric(String letters) {
  var sum = 0;
  var mul = 1;
  var n;
  for (var index = letters.length - 1; index >= 0; index--) {
    var c = letters[index].codeUnitAt(0);
    n = 1;
    if (65 <= c && c <= 90) {
      n += c - 65;
    } else if (97 <= c && c <= 122) {
      n += c - 97;
    }
    sum += n * mul;
    mul = mul * 26;
  }
  return sum;
}

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

int _letterOnly(int rune) {
  if (65 <= rune && rune <= 90) {
    return rune;
  } else if (97 <= rune && rune <= 122) {
    return rune - 32;
  }
  return 0;
}

// Not used
//int _intOnly(int rune) {
//  if (rune >= 48 && rune < 58) {
//    return rune;
//  }
//  return 0;
//}

String _twoDigits(int n) {
  if (n >= 10) {
    return "${n}";
  }
  return "0${n}";
}

/// Returns the coordinates from a cell name.
/// "A1" returns [1, 1] and the "B3" return [2, 3].
List cellCoordsFromCellId(String cellId) {
  var letters = cellId.runes.map(_letterOnly);
  var lettersPart =
      utf8.decode(letters.where((rune) => rune > 0).toList(growable: false));
  var numericsPart = cellId.substring(lettersPart.length);
  var x = lettersToNumeric(lettersPart);
  var y = int.parse(numericsPart);
  return [x, y];
}

/// Read and parse XSLX spreadsheet
class XlsxDecoder extends ExcelIt {
  String get mediaType =>
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  String get extension => ".xlsx";

  String _stylesTarget, _sharedStringsTarget;
  List<String> _rId;

  XlsxDecoder(Archive archive, {bool update = false}) {
    this._archive = archive;
    this._update = update;
    if (_update == true) {
      _archiveFiles = <String, ArchiveFile>{};
      _sheets = <String, XmlNode>{};
      _xmlFiles = <String, XmlDocument>{};
    }
    _worksheetTargets = Map<String, String>();
    _tables = Map<String, SpreadsheetTable>();
    _sharedStrings = List<String>();
    _numFormats = List<int>();
    _rId = new List<String>();
    _parseRelations();
    _parseStyles();
    _parseSharedStrings();
    _parseContent();
  }

  String dumpXmlContent([String sheet]) {
    if (sheet == null) {
      var buffer = StringBuffer();
      _sheets.forEach((name, document) {
        buffer.writeln(name);
        buffer.writeln(document.toXmlString(pretty: true));
      });
      return buffer.toString();
    } else {
      return _sheets[sheet].toXmlString(pretty: true);
    }
  }

  void insertColumn(String sheet, int columnIndex) {
    super.insertColumn(sheet, columnIndex);

    for (var row in _findRows(_sheets[sheet])) {
      XmlElement cell;
      var cells = _findCells(row);

      var currentIndex = 0; // cells could be empty
      for (var currentCell in cells) {
        currentIndex = _getCellNumber(currentCell) - 1;
        if (currentIndex >= columnIndex) {
          cell = currentCell;
          break;
        }
      }

      if (cell != null) {
        cells
            .skipWhile((c) => c != cell)
            .forEach((c) => _setCellColNumber(c, _getCellNumber(c) + 1));
      }
      // Nothing to do if cell == null
    }
  }

  void removeColumn(String sheet, int columnIndex) {
    super.removeColumn(sheet, columnIndex);

    for (var row in _findRows(_sheets[sheet])) {
      XmlElement cell;
      var cells = _findCells(row);

      var currentIndex = 0; // cells could be empty
      for (var currentCell in cells) {
        currentIndex = _getCellNumber(currentCell) - 1;
        if (currentIndex >= columnIndex) {
          cell = currentCell;
          break;
        }
      }

      if (cell != null) {
        cells
            .skipWhile((c) => c != cell)
            .forEach((c) => _setCellColNumber(c, _getCellNumber(c) - 1));
        cell.parent.children.remove(cell);
      }
    }
  }

  void insertRow(String sheet, int rowIndex) {
    super.insertRow(sheet, rowIndex);

    var parent = _sheets[sheet];
    if (rowIndex < _tables[sheet]._maxRows - 1) {
      var foundRow = _findRowByIndex(_sheets[sheet], rowIndex);
      _insertRow(parent, foundRow, rowIndex);
      parent.children.skipWhile((row) => row != foundRow).forEach((row) {
        var rIndex = _getRowNumber(row) + 1;
        _setRowNumber(row, rIndex);
        _findCells(row).forEach((cell) {
          _setCellRowNumber(cell, rIndex);
        });
      });
    } else {
      _insertRow(parent, null, rowIndex);
    }
  }

  void removeRow(String sheet, int rowIndex) {
    super.removeRow(sheet, rowIndex);

    var parent = _sheets[sheet];
    var foundRow = _findRowByIndex(parent, rowIndex);
    parent.children.skipWhile((row) => row != foundRow).forEach((row) {
      var rIndex = _getRowNumber(row) - 1;
      _setRowNumber(row, rIndex);
      _findCells(row).forEach((cell) {
        _setCellRowNumber(cell, rIndex);
      });
    });
    parent.children.remove(foundRow);
  }

  void updateCell(String sheet, int columnIndex, int rowIndex, dynamic value) {
    super.updateCell(sheet, columnIndex, rowIndex, value);

    var foundRow = _findRowByIndex(_sheets[sheet], rowIndex);
    _updateCell(foundRow, columnIndex, rowIndex, value);
  }

  _parseRelations() {
    var relations = _archive.findFile('xl/_rels/workbook.xml.rels');
    if (relations != null) {
      relations.decompress();
      var document = parse(utf8.decode(relations.content));
      if (_xmlFiles != null) _xmlFiles["xl/_rels/workbook.xml.rels"] = document;
      document.findAllElements('Relationship').forEach((node) {
        String id = node.getAttribute('Id');
        switch (node.getAttribute('Type')) {
          case _relationshipsStyles:
            _stylesTarget = node.getAttribute('Target');
            break;
          case _relationshipsWorksheet:
            _worksheetTargets[id] = node.getAttribute('Target');
            break;
          case _relationshipsSharedStrings:
            _sharedStringsTarget = node.getAttribute('Target');
            break;
        }
        if (!_rId.contains(id)) _rId.add(id);
      });
    }
  }

  _parseStyles() {
    var styles = _archive.findFile('xl/$_stylesTarget');
    if (styles != null) {
      styles.decompress();
      var document = parse(utf8.decode(styles.content));
      document
          .findAllElements('cellXfs')
          .first
          .findElements('xf')
          .forEach((node) {
        var numFmtId = node.getAttribute('numFmtId');
        if (numFmtId != null) {
          _numFormats.add(int.parse(numFmtId));
        } else {
          _numFormats.add(0);
        }
      });
    }
  }

  _parseSharedStrings() {
    var sharedStrings = _archive.findFile('xl/$_sharedStringsTarget');
    if (sharedStrings != null) {
      sharedStrings.decompress();
      var document = parse(utf8.decode(sharedStrings.content));
      document.findAllElements('si').forEach((node) {
        _parseSharedString(node);
      });
    }
  }

  _parseSharedString(XmlElement node) {
    var list = List();
    node.findAllElements('t').forEach((child) {
      list.add(_parseValue(child));
    });
    _sharedStrings.add(list.join(''));
  }

  _parseContent() {
    var workbook = _archive.findFile('xl/workbook.xml');
    workbook.decompress();
    var document = parse(utf8.decode(workbook.content));
    if (_xmlFiles != null) _xmlFiles["xl/workbook.xml"] = document;
    document.findAllElements('sheet').forEach((node) {
      _parseTable(node);
    });
  }

  static void _setRowNumber(XmlElement row, int index) =>
      row.getAttributeNode('r').value = index.toString();

  static void _setCellColNumber(XmlElement cell, int colIndex) {
    var attr = cell.getAttributeNode('r');
    var coords = cellCoordsFromCellId(attr.value);
    attr.value = '${numericToLetters(colIndex)}${coords[1]}';
  }

  static void _setCellRowNumber(XmlElement cell, int rowIndex) {
    var attr = cell.getAttributeNode('r');
    var coords = cellCoordsFromCellId(attr.value);
    attr.value = '${numericToLetters(coords[0])}${rowIndex}';
  }

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
    if (row == null || currentIndex != rowIndex) {
      row = _insertRow(table, row, rowIndex);
    }

    return row;
  }

  XmlElement _updateCell(
      XmlElement node, int columnIndex, int rowIndex, dynamic value) {
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

    if (cell == null || currentIndex != columnIndex) {
      cell = _insertCell(node, cell, columnIndex, rowIndex, value);
    } else {
      cell = _replaceCell(node, cell, columnIndex, rowIndex, value);
    }

    return cell;
  }

  static XmlElement _createRow(int rowIndex) {
    var attributes = <XmlAttribute>[
      XmlAttribute(XmlName('r'), (rowIndex + 1).toString()),
    ];
    return XmlElement(XmlName('row'), attributes, []);
  }

  static XmlElement _insertRow(
      XmlElement table, XmlElement lastRow, int rowIndex) {
    var row = _createRow(rowIndex);
    if (lastRow == null) {
      table.children.add(row);
    } else {
      var index = table.children.indexOf(lastRow);
      table.children.insert(index, row);
    }
    return row;
  }

  static XmlElement _insertCell(XmlElement row, XmlElement lastCell,
      int columnIndex, int rowIndex, dynamic value) {
    var cell = _createCell(columnIndex, rowIndex, value);
    if (lastCell == null) {
      row.children.add(cell);
    } else {
      var index = row.children.indexOf(lastCell);
      row.children.insert(index, cell);
    }
    return cell;
  }

  static XmlElement _replaceCell(XmlElement row, XmlElement lastCell,
      int columnIndex, int rowIndex, dynamic value) {
    var index = lastCell == null ? 0 : row.children.indexOf(lastCell);
    var cell = _createCell(columnIndex, rowIndex, value);
    row.children
      ..removeAt(index)
      ..insert(index, cell);
    return cell;
  }

  // TODO Manage value's type
  static XmlElement _createCell(int columnIndex, int rowIndex, dynamic value) {
    var attributes = <XmlAttribute>[
      XmlAttribute(
          XmlName('r'), '${numericToLetters(columnIndex + 1)}${rowIndex + 1}'),
      XmlAttribute(XmlName('t'), 'inlineStr'),
    ];
    var children = value == null
        ? <XmlElement>[]
        : <XmlElement>[
            XmlElement(XmlName('is'), [], [
              XmlElement(XmlName('t'), [], [XmlText(value.toString())])
            ]),
          ];
    return XmlElement(XmlName('c'), attributes, children);
  }
}
