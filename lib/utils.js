var compress, computeColIndex, escapeXML, _;

_ = require("lodash");

computeColIndex = function(colName) {
  var i, multiplier, result, value, _i, _ref;
  result = 0;
  multiplier = 1;
  for (i = _i = _ref = colName.length - 1; _i >= 0; i = _i += -1) {
    value = (colName.charCodeAt(i) - "A".charCodeAt(0)) + 1;
    result = result + value * multiplier;
    multiplier = multiplier * 26;
  }
  return result;
};

module.exports = {
  colChar: function(input) {
    var a, colIndex;
    input = input.toString(26);
    colIndex = '';
    while (input.length) {
      a = input.charCodeAt(input.length - 1);
      colIndex = String.fromCharCode(a + (a >= 48 && a <= 57 ? 17 : -22)) + colIndex;
      input = input.length > 1 ? (parseInt(input.substr(0, input.length - 1), 26) - 1).toString(26) : "";
    }
    return colIndex;
  },
  cellDecode: function(a1Notation) {
    var col, m;
    m = a1Notation.match(/^([A-Z]+)([\d]+)$/);
    if (!(m && m.length === 3)) {
      throw 'Invalid a1Notation';
    }
    col = computeColIndex(m[1]);
    return {
      col: col,
      row: parseInt(m[2], 10)
    };
  },
  escapeXML: escapeXML = function(str) {
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  },
  compress: compress = function(str) {
    return String(str).replace(/\n\s*/g, '');
  },
  buildCell: function(ref, val, styles) {
    var f, getStyle, r, s, t, v;
    getStyle = function(nf) {
      var getBuiltinNumFmtId, numFmtId, r, s;
      if (!nf) {
        return;
      }
      r = styles.formatCodesToStyleIndex[nf];
      if (r) {
        return r;
      }
      getBuiltinNumFmtId = function(nf) {
        var builtin_nfs;
        builtin_nfs = {
          'General': 0,
          '': 0,
          '0': 1,
          '0.00': 2,
          '#,##0': 3,
          '#,##0.00': 4,
          '0%': 9,
          '0.00%': 10,
          '0.00E+00': 11,
          '# ?/?': 12,
          '# ??/??': 13,
          'm/d/yy': 14,
          'd-mmm-yy': 15,
          'd-mmm': 16,
          'mmm-yy': 17,
          'h:mm AM/PM': 18,
          'h:mm:ss AM/PM': 19,
          'h:mm': 20,
          'h:mm:ss': 21,
          'm/d/yy h:mm': 22,
          '[$-404]e/m/d': 27,
          '#,##0 ;(#,##0)': 37,
          '#,##0 ;[Red](#,##0)': 38,
          '#,##0.00;(#,##0.00)': 39,
          '#,##0.00;[Red](#,##0.00)': 40,
          '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)': 44,
          'mm:ss': 45,
          '[h]:mm:ss': 46,
          'mmss.0': 47,
          '##0.0E+0': 48,
          '@': 49,
          't0': 59,
          't0.00': 60,
          't#,##0': 61,
          't#,##0.00': 62,
          't0%': 67,
          't0.00%': 68,
          't# ?/?': 69,
          't# ??/??': 70
        };
        r = builtin_nfs[nf];
        return r;
      };
      numFmtId = getBuiltinNumFmtId(nf);
      if (!numFmtId) {
        styles.customFormatsCount++;
        numFmtId = 164 + styles.customFormatsCount;
        styles.numFmts.push({
          numFmtId: numFmtId,
          formatCode: nf
        });
      }
      s = styles.cellStyleXfs.length;
      styles.cellStyleXfs.push({
        numFmtId: numFmtId,
        formatCode: nf
      });
      styles.formatCodesToStyleIndex[nf] = s;
      return s;
    };
    if (val == null) {
      return '';
    }
    if (typeof val === 'object' && !_.isDate(val)) {
      v = val.v;
      t = val.t;
      s = val.s;
      f = val.f;
      if (!s && val.nf) {
        s = getStyle(val.nf);
      }
    } else {
      v = val;
    }
    if (_.isNumber(v) && _.isFinite(v)) {
      v = '<v>' + v + '</v>';
      if (val.nf && !t) {
        t = 'n';
      }
    } else if (_.isDate(v)) {
      t = 'd';
      if (s == null) {
        s = '2';
      }
      v = '<v>' + v.toISOString() + '</v>';
    } else if (_.isBoolean(v)) {
      t = 'b';
      v = '<v>' + (v === true ? '1' : '0') + '</v>';
    } else if (v) {
      v = '<is><t>' + escapeXML(v) + '</t></is>';
      t = 'inlineStr';
    }
    if (!(v || f)) {
      return '';
    }
    r = '<c r="' + ref + '"';
    if (t) {
      r += ' t="' + t + '"';
    }
    if (s) {
      r += ' s="' + s + '"';
    }
    r += '>';
    if (f) {
      r += '<f>' + escapeXML(f) + '</f>';
    }
    if (v) {
      r += v;
    }
    r += '</c>';
    return r;
  },
  buildComment: function(ref, val, comments, authors) {
    var authorId, comment;
    if (!(val && (val.c != null))) {
      return;
    }
    comment = val.c;
    if (typeof comment === 'string') {
      comment = {
        lines: [comment]
      };
    }
    if (comment.author) {
      authorId = authors.indexOf(comment.author);
      if (authorId === -1) {
        authorId = authors.length;
        authors.push(comment.author);
      }
      comment.authorId = authorId;
    }
    comment.ref = ref;
    return comments.push(comment);
  }
};
