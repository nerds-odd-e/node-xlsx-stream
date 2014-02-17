var esc, utils, xml;

utils = require('./utils');

xml = utils.compress;

esc = utils.escapeXML;

module.exports = {
  worksheet: {
    header: function(opts) {
      var colIdx, cols, colsTag, decoded, frozenCell, i, minMax, options, rangePattern, sheetView, width, _i, _len, _ref, _ref1;
      options = opts || {};
      if (options.frozenCell) {
        frozenCell = options.frozenCell;
        decoded = utils.cellDecode(frozenCell);
        sheetView = "<sheetView workbookViewId=\"0\">\n  <pane topLeftCell=\"" + frozenCell + "\" ySplit=\"" + (decoded.row - 1) + ".0\" xSplit=\"" + (decoded.col - 1) + ".0\" activePane=\"bottomRight\" state=\"frozen\" />\n</sheetView>";
      } else {
        sheetView = '<sheetView workbookViewId="0"/>';
      }
      colsTag = "";
      if (options.hiddenColumns) {
        _ref = options.hiddenColumns;
        for (i = _i = 0, _len = _ref.length; _i < _len; i = ++_i) {
          rangePattern = _ref[i];
          minMax = utils.rangePatternToMinMax(rangePattern, 1025);
          colsTag += "<col hidden=\"true\" min=\"" + minMax[0] + "\" max=\"" + minMax[1] + "\" width=\"0\" />";
        }
      }
      if (options.columnsWidth) {
        _ref1 = options.columnsWidth;
        for (colIdx in _ref1) {
          width = _ref1[colIdx];
          colsTag += "<col min=\"" + colIdx + "\" max=\"" + colIdx + "\" width=\"" + width + "\" />";
        }
      }
      cols = colsTag ? "  <cols>\n    " + colsTag + "\n  </cols>" : "";
      return xml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">\n  <sheetViews>\n    " + sheetView + "\n  </sheetViews>\n  <sheetFormatPr customHeight=\"1\" defaultColWidth=\"14.43\" defaultRowHeight=\"15\"/>\n    " + cols);
    },
    footer: function(sheet, nRow) {
      var buf;
      buf = "";
      buf += nRow > 0 ? '</sheetData>' : '<sheetData/>';
      if (sheet.comments.length) {
        buf += '<legacyDrawing r:id="rId1" />';
      }
      buf += '</worksheet>';
      return xml(buf);
    }
  },
  sheet_related: {
    "[Content_Types].xml": {
      header: function(opts) {
        var customProps;
        if (opts.custom) {
          customProps = '\n<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>';
        } else {
          customProps = '';
        }
        return xml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n  <Default ContentType=\"image/jpeg\" Extension=\"jpeg\"/>\n  <Default ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\" Extension=\"vml\"/>\n  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n  <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n  <Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>\n  <Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\n  <Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>" + customProps);
      },
      sheet: function(sheet) {
        var buf;
        buf = "<Override PartName=\"/" + (esc(sheet.path)) + "\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";
        if (sheet.comments.length) {
          buf += "<Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\" PartName=\"/xl/comments" + sheet.index + ".xml\"/>";
        }
        return buf;
      },
      footer: xml("</Types>")
    },
    "xl/_rels/workbook.xml.rels": {
      header: function(opts) {
        return xml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
      },
      sheet: function(sheet) {
        return "<Relationship Id=\"rSheet" + (esc(sheet.index)) + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"" + (esc(sheet.rel)) + "\"/>";
      },
      footer: xml("  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>\n  <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\n</Relationships>")
    },
    "xl/workbook.xml": {
      header: function(opts) {
        return xml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n  <fileVersion appName=\"xl\" lastEdited=\"5\" lowestEdited=\"5\" rupBuild=\"9303\"/>\n  <workbookPr defaultThemeVersion=\"124226\"/>\n  <bookViews>\n  <workbookView xWindow=\"480\" yWindow=\"60\" windowWidth=\"18195\" windowHeight=\"8505\"/>\n  </bookViews>\n  <sheets>");
      },
      sheet: function(sheet) {
        return xml("<sheet name=\"" + (esc(sheet.name)) + "\" sheetId=\"" + (esc(sheet.index)) + "\" r:id=\"rSheet" + (esc(sheet.index)) + "\"/>");
      },
      footer: xml("  </sheets>\n  <calcPr calcId=\"145621\"/>\n</workbook>")
    }
  },
  styles: function(styl) {
    var cellXfItems, cellXfs, item, numFmtItems, numFmts, _i, _j, _len, _len1, _ref, _ref1;
    numFmtItems = "";
    _ref = styl.numFmts;
    for (_i = 0, _len = _ref.length; _i < _len; _i++) {
      item = _ref[_i];
      numFmtItems += "  <numFmt numFmtId=\"" + item.numFmtId + "\" formatCode=\"" + (esc(item.formatCode)) + "\" />\n";
    }
    numFmts = numFmtItems ? "<numFmts count=\"" + styl.numFmts.length + "\">\n  " + numFmtItems + "</numFmts>" : "";
    cellXfItems = "";
    _ref1 = styl.cellStyleXfs;
    for (_j = 0, _len1 = _ref1.length; _j < _len1; _j++) {
      item = _ref1[_j];
      cellXfItems += "  <xf xfId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" numFmtId=\"" + item.numFmtId + "\" applyNumberFormat=\"1\"/>\n";
    }
    cellXfs = cellXfItems ? "<cellXfs count=\"" + (Object.keys(styl.cellStyleXfs).length) + "\">\n  " + cellXfItems + "\n</cellXfs>" : "";
    return xml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">\n  " + numFmts + "\n  <fonts count=\"1\" x14ac:knownFonts=\"1\">\n    <font>\n      <sz val=\"11\"/>\n      <color theme=\"1\"/>\n      <name val=\"Calibri\"/>\n      <family val=\"2\"/>\n      <scheme val=\"minor\"/>\n    </font>\n  </fonts>\n  <fills count=\"2\">\n    <fill>\n      <patternFill patternType=\"none\"/>\n    </fill>\n    <fill>\n      <patternFill patternType=\"gray125\"/>\n    </fill>\n  </fills>\n  <borders count=\"1\">\n    <border>\n      <left/>\n      <right/>\n      <top/>\n      <bottom/>\n      <diagonal/>\n    </border>\n  </borders>\n  " + cellXfs + "\n  <cellStyles count=\"1\">\n    <cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>\n  </cellStyles>\n  <dxfs count=\"0\"/>\n  <tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>\n  <extLst>\n    <ext uri=\"{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">\n      <x14:slicerStyles defaultSlicerStyle=\"SlicerStyleLight1\"/>\n    </ext>\n  </extLst>\n</styleSheet>");
  },
  statics: {
    "xl/sharedStrings.xml": xml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"/>"),
    "docProps/app.xml": xml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\n  <Application>node-xlsx-stream</Application>\n  <DocSecurity>0</DocSecurity>\n  <ScaleCrop>false</ScaleCrop>\n  <Company></Company>\n  <LinksUpToDate>false</LinksUpToDate>\n  <SharedDoc>false</SharedDoc>\n  <HyperlinksChanged>false</HyperlinksChanged>\n  <AppVersion>" + (require('../package.json').version) + "</AppVersion>\n</Properties>")
  },
  semiStatics: {
    "_rels/.rels": function(opts) {
      var customProps;
      if (opts.custom) {
        customProps = "<Relationship Id=\"rId4\" Target=\"docProps/custom.xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties\"/>";
      } else {
        customProps = "";
      }
      return xml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n  <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\n  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>\n  <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>\n  " + customProps + "\n</Relationships>");
    },
    "docProps/core.xml": function(opts) {
      var coreProps, extra, today;
      if (!opts) {
        opts = {};
      }
      coreProps = opts.core || {};
      today = new Date().toISOString();
      extra = coreProps.title ? "<dc:title>" + (esc(coreProps.title)) + "</dc:title>\n  " : "";
      if (coreProps.subject) {
        extra += "<dc:subject>" + (esc(coreProps.subject)) + "</dc:subject>\n  ";
      }
      extra += "<dc:creator>" + (coreProps.creator ? esc(coreProps.creator) : 'node-xlsx-stream') + "</dc:creator>\n  ";
      extra += "<cp:lastModifiedBy>" + (coreProps.lastModifiedBy ? esc(coreProps.lastModifiedBy) : 'node-xlsx-stream') + "</cp:lastModifiedBy>\n  ";
      if (coreProps.description) {
        extra += "<dc:description>" + (esc(coreProps.description)) + "</dc:description>\n  ";
      }
      return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n  " + extra + "<dcterms:created xsi:type=\"dcterms:W3CDTF\">" + today + "</dcterms:created>\n  <dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + today + "</dcterms:modified>\n</cp:coreProperties>";
    },
    "docProps/custom.xml": function(opts) {
      var extra, getVTypeProperty, index, key, prop, value, _i, _len, _ref, _ref1;
      if (!(opts && opts.custom)) {
        return;
      }
      extra = "";
      index = 1;
      getVTypeProperty = function(name, value, index) {
        var type;
        if (typeof value === 'string') {
          type = 'lpwstr';
          value = esc(value);
        } else if (typeof value === 'boolean') {
          type = 'bool';
        } else if (typeof value === 'number') {
          if (parseInt(value, 10) === value) {
            type = 'i4';
          } else {
            type = 'r8';
          }
        } else if (value instanceof Date) {
          type = 'filetime';
          value = value.toISOString();
        } else if (value === null) {
          type = 'null';
        } else if (value !== void 0) {
          type = 'lpwstr';
          value = value.toString();
        }
        return "<property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" name=\"" + name + "\" pid=\"" + index + "\">\n  <vt:" + type + ">" + value + "</vt:" + type + ">\n</property>";
      };
      if (Array.isArray(opts.custom)) {
        _ref = opts.custom;
        for (_i = 0, _len = _ref.length; _i < _len; _i++) {
          prop = _ref[_i];
          index++;
          extra += getVTypeProperty(prop.name, prop.value, index);
        }
      } else {
        _ref1 = opts.custom;
        for (key in _ref1) {
          value = _ref1[key];
          index++;
          extra += getVTypeProperty(key, value, index);
        }
      }
      return xml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\n  " + extra + "\n</Properties>");
    }
  },
  sheetRels: function(sheet) {
    var relBuf;
    relBuf = sheet.comments.length ? "<Relationship Id=\"rId1\" Target=\"../drawings/vmlDrawing" + sheet.index + ".vml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\"/>\n<Relationship Id=\"rId2\" Target=\"../comments" + sheet.index + ".xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\"/>" : "";
    return xml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" + relBuf + "</Relationships>");
  },
  vmlDrawing: function(sheet) {
    var footer, header, shape, shapes;
    header = xml("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:mv=\"http://macVmlSchemaUri\">\n  <o:shapelayout v:ext=\"edit\">\n   <o:idmap v:ext=\"edit\" data=\"1\"/>\n  </o:shapelayout>\n  <v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m0,0l0,21600,21600,21600,21600,0xe\">\n    <v:stroke joinstyle=\"miter\"/>\n    <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\n  </v:shapetype>");
    shape = function(a1Notation) {
      var colNumber, decoded, point_from_left, point_from_top, rowNumber, unique_id;
      decoded = utils.cellDecode(a1Notation);
      rowNumber = decoded.row;
      colNumber = decoded.col;
      sheet.shapeCounter++;
      unique_id = "_x" + sheet.index + "_s" + sheet.shapeCounter;
      point_from_left = colNumber * 100 + 30;
      point_from_top = rowNumber * 20 + 5;
      return xml("<v:shape id=\"" + unique_id + "\" type=\"#_x0000_t202\" style='position:absolute;margin-left:\"" + point_from_left + "\"pt;margin-top:\"" + point_from_top + "\"pt;width:104pt;height:64pt;z-index:" + sheet.shapeCounter + ";visibility:hidden;mso-wrap-style:tight' fillcolor=\"#fbf6d6\" strokecolor=\"#edeaa1\">\n  <v:fill color2=\"#fbfe82\" angle=\"-180\" type=\"gradient\">\n   <o:fill v:ext=\"view\" type=\"gradientUnscaled\"/>\n  </v:fill>\n  <v:shadow on=\"t\" obscured=\"t\"/>\n  <v:path o:connecttype=\"none\"/>\n  <v:textbox>\n   <div style='text-align:left'></div>\n  </v:textbox>\n  <x:ClientData ObjectType=\"Note\">\n   <x:MoveWithCells/>\n   <x:SizeWithCells/>\n   <x:Anchor>\n    " + colNumber + ", 15, " + rowNumber + ", 2, " + (colNumber + 3) + ", 54, " + (rowNumber + 3) + ", 4</x:Anchor>\n   <x:AutoFill>False</x:AutoFill>\n   <x:Row>" + (rowNumber - 1) + "</x:Row>\n   <x:Column>" + (colNumber - 1) + "</x:Column>\n  </x:ClientData>\n </v:shape>");
    };
    shapes = function() {
      var buffer, comment, _i, _len, _ref, _results;
      buffer = "";
      _ref = sheet.comments;
      _results = [];
      for (_i = 0, _len = _ref.length; _i < _len; _i++) {
        comment = _ref[_i];
        _results.push(buffer += shape(comment.ref));
      }
      return _results;
    };
    footer = xml("</xml>");
    return xml(header + shapes() + footer);
  },
  comments: function(sheet) {
    var author, authors, body, footer, header, _i, _len, _ref;
    authors = "";
    if (sheet.authors.length === 0) {
      sheet.authors.push("");
    }
    _ref = sheet.authors;
    for (_i = 0, _len = _ref.length; _i < _len; _i++) {
      author = _ref[_i];
      authors += "<author>" + author + "</author>";
    }
    header = xml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n<authors>\n  " + authors + "\n</authors>\n<commentList>");
    body = function(sheet) {
      var all, comment, generateOneComment, generateOneLine, _j, _len1, _ref1;
      generateOneLine = function(line) {
        return "<r>\n  <rPr>" + (line.b ? '\n    <b/>' : '') + "\n    <sz val=\"9\"/>\n    <color indexed=\"81\"/>\n    <rFont val=\"Calibri\"/>\n    <family val=\"2\"/>\n  </rPr>\n  <t xml:space=\"preserve\">" + (line.t ? line.t : line) + "</t>\n</r>";
      };
      generateOneComment = function(comment) {
        var authorId, line, lines, _j, _len1, _ref1;
        authorId = typeof comment.authorId === 'number' ? comment.authorId : 0;
        lines = "";
        _ref1 = comment.lines;
        for (_j = 0, _len1 = _ref1.length; _j < _len1; _j++) {
          line = _ref1[_j];
          lines += generateOneLine(line);
        }
        return "<comment authorId=\"" + authorId + "\" ref=\"" + comment.ref + "\">\n  <text>" + lines + "</text>\n</comment>";
      };
      all = "";
      _ref1 = sheet.comments;
      for (_j = 0, _len1 = _ref1.length; _j < _len1; _j++) {
        comment = _ref1[_j];
        all += generateOneComment(comment);
      }
      return all;
    };
    footer = "</commentList>\n</comments>";
    return header + '\n' + body(sheet) + '\n' + footer;
  }
};
