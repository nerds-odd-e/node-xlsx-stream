_ = require "lodash"
through = require('through')

utils = require('./utils')
template = require('./templates').worksheet

module.exports = sheetStream = (zip, sheet, opts={})->
  # 列番号の26進表記(A, B, .., Z, AA, AB, ..)
  # 一度計算したらキャッシュしておく。
  colChar = _.memoize utils.colChar

  # 行ごとに変換してxl/worksheets/sheet1.xml に追加
  nRow = 0
  onData = (row)->
    buf = if nRow == 0 then "<sheetData>" else ""
    nRow++
    buf += "<row r='#{nRow}'>"
    if opts.columns?
      for col, i in opts.columns
        ref = "#{colChar(i)}#{nRow}"
        utils.buildComment(ref, row[col], sheet.comments, sheet.authors)
        buf += utils.buildCell(ref, row[col], sheet.styles, opts)
    else
      for val, i in row
        ref = "#{colChar(i)}#{nRow}"
        utils.buildComment(ref, val, sheet.comments, sheet.authors)
        buf += utils.buildCell(ref, val, sheet.styles, opts)
    buf += '</row>'
    @queue buf
  onEnd = ->
    # フッタ部分を追加
    @queue template.footer(sheet, nRow)
    @queue null
    converter = colChar = zip = null

  converter = through(onData, onEnd)
  zip.append converter, name: sheet.path, store: opts.store

  # ヘッダ部分を追加
  converter.queue template.header(sheet.opts)

  return converter
