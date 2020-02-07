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

  List<String> _rId;

  XlsxDecoder(Archive archive, {bool update = false}) {
    this._archive = archive;
    this._update = update;
    if (_update) {
      _archiveFiles = <String, ArchiveFile>{};
      _sheets = <String, XmlNode>{};
      _xmlFiles = <String, XmlDocument>{};
    }
    _worksheetTargets = Map<String, String>();
    _colorMap = Map<String, Map<String, List<String>>>();
    _fontColorHex = List<String>();
    _foregroundColorHex = List<String>();
    _backgroundColorHex = List<String>();
    _tables = Map<String, SpreadsheetTable>();
    _sharedStrings = List<String>();
    _rId = new List<String>();
    _numFormats = List<int>();
    _parseRelations();
    _parseStyles(_stylesTarget);
    _parseSharedStrings();
    _parseContent();
    _extractColors();
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

  void updateCell(String sheet, int columnIndex, int rowIndex, dynamic value,
      {String fontColorHex,
      String foregroundColorHex,
      String backgroundColorHex}) {
    super.updateCell(sheet, columnIndex, rowIndex, value);

    String rC = '${numericToLetters(columnIndex + 1)}${rowIndex + 1}';

    if (fontColorHex != null) {
      _addColor(sheet, fontColorHex, _fontColorHex, rC, 0);
    }

    if (backgroundColorHex != null) {
      _addColor(sheet, foregroundColorHex, _foregroundColorHex, rC, 1);
    }

    if (backgroundColorHex != null) {
      _addColor(sheet, backgroundColorHex, _backgroundColorHex, rC, 2);
    }

    /* var foundRow = _findRowByIndex(_sheets[sheet], rowIndex);
    _updateCell(foundRow, columnIndex, rowIndex, value); */
  }

  _addColor(
      String sheet, String color, List<String> list, String rowCol, int index) {
    if (color.length != 7)
      throw ArgumentError(
          "\nIn-appropriate Color provided. Use colorHex as example of: #FF0000\n");

    String hex = color.replaceAll(new RegExp(r'#'), 'FF');
    if (!list.contains(hex)) list.add(hex);

    if (_colorMap.containsKey(sheet) && _colorMap[sheet].containsKey(rowCol))
      _colorMap[sheet][rowCol][index] = hex;
    else {
      List l = new List<String>(2);
      l[index] = hex;
      _colorMap[sheet] = new Map<String, List<String>>.from({rowCol: l});
    }
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

  _parseSharedStrings() {
    var sharedStrings = _archive.findFile('xl/$_sharedStringsTarget');
    if (sharedStrings != null) {
      sharedStrings.decompress();
      var document = parse(utf8.decode(sharedStrings.content));
      if (_xmlFiles != null) _xmlFiles["xl/$_sharedStringsTarget"] = document;
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
}

/* 
  void _insertColumn(String sheet, int columnIndex) {
    super._insertColumn(sheet, columnIndex);

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
  } */

/* void removeColumn(String sheet, int columnIndex) {
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
  } */
/* 
  void _insertRow(String sheet, int rowIndex) {
    super._insertRow(sheet, rowIndex);

    var parent = _sheets[sheet];
    if (rowIndex < _tables[sheet]._maxRows - 1) {
      var foundRow = _findRowByIndex(_sheets[sheet], rowIndex);
      __insertRow(parent, foundRow, rowIndex);
      parent.children.skipWhile((row) => row != foundRow).forEach((row) {
        var rIndex = _getRowNumber(row) + 1;
        _setRowNumber(row, rIndex);
        _findCells(row).forEach((cell) {
          _setCellRowNumber(cell, rIndex);
        });
      });
    } else {
      __insertRow(parent, null, rowIndex);
    }
  }
 */
/* void removeRow(String sheet, int rowIndex) {
    super.removeRow(sheet, rowIndex);

    var parent = _sheets[sheet];
    var foundRow = _findRowByIndex(parent, rowIndex);
    parent.children.skipWhile((row) => row != foundRow).forEach((row) {
      var rIndex = _getRowNumber(row) - 1;
      _setRowNumber(row, rIndex);
      _findCells(row).forEach((cell) => _setCellRowNumber(cell, rIndex));
    });
    parent.children.remove(foundRow);
  } */
