part of excel_it;

const String _relationshipsStyles =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const String _relationshipsWorksheet =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const String _relationshipsSharedStrings =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

/// Convert a character based column
int lettersToNumeric(String letters) {
  var sum = 0, mul = 1, n;
  for (var index = letters.length - 1; index >= 0; index--) {
    var c = letters[index].codeUnitAt(0);
    n = 1;
    if (65 <= c && c <= 90)
      n += c - 65;
    else if (97 <= c && c <= 122) n += c - 97;

    sum += n * mul;
    mul = mul * 26;
  }
  return sum;
}

int _letterOnly(int rune) {
  if (65 <= rune && rune <= 90)
    return rune;
  else if (97 <= rune && rune <= 122) return rune - 32;

  return 0;
}

String _twoDigits(int n) {
  if (n > 9) return "$n";
  return "0$n";
}

/// Returns the coordinates from a cell name.
/// "A1" returns [1, 1] and the "B3" return [2, 3].
List cellCoordsFromCellId(String cellId) {
  var letters = cellId.runes.map(_letterOnly);
  var lettersPart =
      utf8.decode(letters.where((rune) => rune > 0).toList(growable: false));
  var numericsPart = cellId.substring(lettersPart.length);

  return [lettersToNumeric(lettersPart), int.parse(numericsPart)]; // [x , y]
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
    _worksheetTargets = new Map<String, String>();
    _colorMap = new Map<String, Map<String, List<String>>>();
    _fontColorHex = new List<String>();
    _patternFill = new Map<String, List<String>>();
    _cellXfs = new Map<String, List<int>>();
    _tables = new Map<String, SpreadsheetTable>();
    _sharedStrings = new List<String>();
    _rId = new List<String>();
    _numFormats = new List<int>();
    _putContentXml();
    _parseRelations();
    _parseStyles(_stylesTarget);
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
    } else
      return _sheets[sheet].toXmlString(pretty: true);
  }

  void updateCell(String sheet, int columnIndex, int rowIndex, dynamic value,
      {String fontColorHex,
      String foregroundColorHex,
      String backgroundColorHex}) {
    super.updateCell(sheet, columnIndex, rowIndex, value);

    String rC = '${numericToLetters(columnIndex + 1)}${rowIndex + 1}';

    if (fontColorHex != null) {
      _addColor(sheet, fontColorHex, rC, 0);
    }

    if (foregroundColorHex != null) {
      _addColor(sheet, foregroundColorHex, rC, 1);
    }

    if (backgroundColorHex != null) {
      _addColor(sheet, backgroundColorHex, rC, 2);
    }
  }

  _addColor(String sheet, String color, String rowCol, int index) {
    if (color != null && color.length != 7)
      throw ArgumentError(
          "InAppropriate Color provided. Use colorHex as example of: #FF0000");

    String hex = color.replaceAll(new RegExp(r'#'), 'FF');

    if (_colorMap.containsKey(sheet)) {
      if (_colorMap[sheet].containsKey(rowCol))
        _colorMap[sheet][rowCol][index] = hex;
      else {
        List l = new List<String>(3);
        l[index] = hex;
        Map temp = new Map<String, List<String>>.from(_colorMap[sheet]);
        temp[rowCol] = l;
        _colorMap[sheet] = new Map<String, List<String>>.from(temp);
      }
    } else {
      List l = new List<String>(3);
      l[index] = hex;
      _colorMap[sheet] = new Map<String, List<String>>.from({rowCol: l});
    }
  }

  _putContentXml() {
    var file = _archive.findFile("[Content_Types].xml");

    if (_xmlFiles != null) {
      if (file == null) _damagedExcel();
      file.decompress();
      _xmlFiles["[Content_Types].xml"] = parse(utf8.decode(file.content));
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
    } else
      _damagedExcel();
  }

  _parseSharedStrings() {
    var sharedStrings = _archive.findFile('xl/$_sharedStringsTarget');
    if (sharedStrings == null) {
      var content = utf8.encode(
          "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\">");
      _archive.addFile(
          ArchiveFile('xl/$_sharedStringsTarget', content.length, content));
      sharedStrings = _archive.findFile('xl/$_sharedStringsTarget');
    }
    sharedStrings.decompress();
    var document = parse(utf8.decode(sharedStrings.content));
    if (_xmlFiles != null) _xmlFiles["xl/$_sharedStringsTarget"] = document;
    document.findAllElements('si').forEach((node) {
      _parseSharedString(node);
    });
  }

  _parseSharedString(XmlElement node) {
    var list = List();
    node.findAllElements('t').forEach((child) => list.add(_parseValue(child)));
    _sharedStrings.add(list.join(''));
  }

  _parseContent() {
    var workbook = _archive.findFile('xl/workbook.xml');
    if (workbook == null) _damagedExcel();
    workbook.decompress();
    var document = parse(utf8.decode(workbook.content));
    if (_xmlFiles != null) _xmlFiles["xl/workbook.xml"] = document;
    document.findAllElements('sheet').forEach((node) => _parseTable(node));
  }
}
