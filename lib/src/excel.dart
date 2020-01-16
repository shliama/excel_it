part of excel_it;

const _spreasheetXlsx = 'xlsx';
final Map<String, String> _spreasheetExtensionMap = <String, String>{
  _spreasheetXlsx:
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

// Normalize new line
String _normalizeNewLine(String text) {
  return text.replaceAll("\r\n", "\n");
}

ExcelIt _newSpreadsheetDecoder(Archive archive, bool update) {
  // Lookup at file format
  var format;

  // Try OpenDocument format
  var mimetype = archive.findFile('mimetype');
  if (mimetype != null) {
    mimetype.decompress();
    var content = utf8.decode(mimetype.content);

    // Try OpenXml Office format
  } else {
    var xl = archive.findFile('xl/workbook.xml');
    format = xl != null ? _spreasheetXlsx : null;
  }

  switch (format) {
    case _spreasheetXlsx:
      return XlsxDecoder(archive, update: update);
    default:
      throw UnsupportedError("Spreadsheet format unsupported");
  }
}

// Issue on archive
// https://github.com/brendan-duncan/archive/pull/43
const List<String> _noCompression = const <String>[
  'mimetype',
  'Thumbnails/thumbnail.png',
];

/**
 * Decode a spreadsheet file.
 */
abstract class ExcelIt {
  bool _update;
  Archive _archive;
  Map<String, XmlNode> _sheets;
  Map<String, XmlDocument> _xmlFiles;
  Map<String, ArchiveFile> _archiveFiles;

  Map<String, SpreadsheetTable> _tables;

  /// Media type
  String get mediaType;

  /// Filename extension
  String get extension;

  /// Tables contained in spreadshet file indexed by their names
  Map<String, SpreadsheetTable> get tables => _tables;

  ExcelIt();

  factory ExcelIt.decodeBytes(List<int> data,
      {bool update: false, bool verify: false}) {
    var archive = ZipDecoder().decodeBytes(data, verify: verify);
    return _newSpreadsheetDecoder(archive, update);
  }

  factory ExcelIt.decodeBuffer(InputStream input,
      {bool update: false, bool verify: false}) {
    var archive = ZipDecoder().decodeBuffer(input, verify: verify);
    return _newSpreadsheetDecoder(archive, update);
  }

  /// Dump XML content (for debug purpose)
  String dumpXmlContent([String sheet]);

  void _checkSheetArguments(String sheet) {
    if (_update != true) {
      throw ArgumentError("'update' should be set to 'true' on constructor");
    }
    if (_sheets.containsKey(sheet) == false) {
      throw ArgumentError("'$sheet' not found");
    }
  }

  /// Insert column in [sheet] at position [columnIndex]
  void insertColumn(String sheet, int columnIndex) {
    _checkSheetArguments(sheet);
    if (columnIndex < 0 /* || columnIndex > _tables[sheet]._maxCols */) {
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
