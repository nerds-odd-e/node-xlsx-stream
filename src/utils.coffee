_ = require "lodash"

computeColIndex = (colName)->
  result = 0
  multiplier = 1
  for i in [colName.length - 1..0] by -1
    value = (colName.charCodeAt(i) - "A".charCodeAt(0)) + 1
    result = result + value * multiplier
    multiplier = multiplier * 26
  return result

msPerDay = 8.64e7
zeroLocal = new Date(1899,11,30); # dont utc, it displays the wrong date in excel
toOADate = (d)->
  v = (d - zeroLocal)/msPerDay
  # Deal with dates prior to 1899-12-30 00:00:00
  if v < 0
    dec = v - Math.floor(v);
    if dec
      v = Math.floor(v) - dec;
  return v;

module.exports =
  colChar: (input)->
    input = input.toString(26)
    colIndex = ''
    while input.length
      a = input.charCodeAt(input.length - 1)
      colIndex = String.fromCharCode(a + if a >= 48 and a <= 57 then 17 else -22) + colIndex
      input = if input.length > 1 then (parseInt(input.substr(0, input.length - 1), 26) - 1).toString(26) else ""
    return colIndex

  cellDecode: (a1Notation)->
    m = a1Notation.match(/^([A-Z]+)([\d]+)$/)
    throw 'Invalid a1Notation' unless m && m.length == 3
    col = computeColIndex m[1]
    return {col: col, row: parseInt(m[2], 10)}

  escapeXML: escapeXML = (str)->
    String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;')
  compress: compress = (str)->
    String(str).replace(/\n\s*/g, '')

  buildCell: (ref, val, styles, opts)->

    getStyle = (nf)->
      return unless nf
      r = styles.formatCodesToStyleIndex[nf]
      return r if r

      getBuiltinNumFmtId = (nf)->
        # ECMA-376 18.8.30
        builtin_nfs =
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
          'm/d/yy': 14, # also 30
          'd-mmm-yy': 15,
          'd-mmm': 16,
          'mmm-yy': 17,
          'h:mm AM/PM': 18,
          'h:mm:ss AM/PM': 19,
          'h:mm': 20,
          'h:mm:ss': 21,
          'm/d/yy h:mm': 22,
          '[$-404]e/m/d': 27, # also 36, 50, 57
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

        r = builtin_nfs[nf]
        return r

      numFmtId = getBuiltinNumFmtId(nf)
      unless numFmtId
        styles.customFormatsCount++
        numFmtId = 164 + styles.customFormatsCount
        styles.numFmts.push(
          numFmtId: numFmtId,
          formatCode: nf)

      s = styles.cellStyleXfs.length
      styles.cellStyleXfs.push({ numFmtId: numFmtId, formatCode: nf })
      styles.formatCodesToStyleIndex[nf] = s
      return s

    return '' unless val?
    if typeof val == 'object' and !_.isDate(val)
      v = val.v
      t = val.t
      s = val.s
      f = val.f
      s = getStyle(val.nf) if not s and val.nf
    else
      v = val

    if _.isNumber(v) and _.isFinite(v)
      v = '<v>' + v + '</v>'
      t = 'n' if val.nf and not t
    else if _.isDate(v)
      if opts.ddates
        t = 'd'
        v = '<v>' + v.toISOString() + '</v>'
      else
        t = 'n'
        v = '<v>' + toOADate(v) + '</v>'
      s = '2' unless s?
    else if _.isBoolean(v)
      t = 'b'
      v = '<v>' + (if v is true then '1' else '0') + '</v>'
    else if v
      v = '<is><t>' + escapeXML(v) + '</t></is>'
      t = 'inlineStr'

    return '' unless v or f
    r = '<c r="' + ref + '"'
    r += ' t="' + t + '"' if t
    r += ' s="' + s + '"' if s
    r += '>'
    r += '<f>' + escapeXML(f) + '</f>' if f
    r += v if v
    r += '</c>'
    return r

  buildComment: (ref, val, comments, authors)->
    return unless val && val.c?
    comment = val.c
    comment = { lines: [ comment ] } if typeof comment == 'string'
    if comment.author
      authorId = authors.indexOf comment.author
      if authorId == -1
        authorId = authors.length
        authors.push comment.author
      comment.authorId = authorId
    comment.ref = ref
    comments.push comment

  rangePatternToMinMax: (rangePattern, max) ->
    splitPattern = rangePattern.split "-"

    if splitPattern.length != 1 && splitPattern.length != 2
      throw new Error "Illegal range pattern : '#{rangePattern}'"

    if splitPattern[1]
      return splitPattern

    if rangePattern[rangePattern.length - 1] == '-'
      splitPattern[1] = max
    else
      splitPattern[1] = splitPattern[0]

    return splitPattern
