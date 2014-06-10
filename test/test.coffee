xlsx_stream = require "../"
vows = require "vows"
assert = require "assert"
office = require "office"

fs = require "fs"
path = require "path"

tmp = (filename)-> path.resolve(__dirname, '../tmp', filename)

vows.describe('xlsx-stream').addBatch(
  "archiver":
    topic: ->
      zip = require("archiver").create('zip')
      zip.pipe require('concat-stream')(@callback)

      stream = require('through')()

      process.nextTick ->
        zip.append stream, name: "0.txt"
        zip.append "aaa", name: "a.txt"
        zip.append "bbb", name: "b.txt"
        zip.finalize()

        process.nextTick ->
          stream.write("ccc")
          stream.end("ddd")
      return
    "bbb": (d)->
      fs.writeFileSync("./tmp/test.zip", d)

).addBatch(
  "Array input":
    "create":
      topic: ->
        x = xlsx_stream()
        output = fs.createWriteStream(tmp('array.xlsx'))
        output.on 'close', @callback
        x.on 'finalize', -> console.log "FINALIZE:", arguments
        x.pipe output
        x.write ["String", "てすと", "&'\";<>", "&amp;"]
        x.write ["Integer", 1,2,-3]
        x.write ["Float", 1.5, 0.3, 0.123456789e+23]
        x.write ["Boolean", true, false]
        x.write ["Date", new Date]
        x.write ["2 Decimals Built-in format #2", { v: 1.5, nf: '0.00' }]
        x.write ["Time Built-in format #18", { v: 1.5, nf: 'h:mm AM/PM' }]
        x.write ["Percentage Built-in format #9", { v: 0.5, nf: '0.00%' }]
        x.write ["Percentage Custom format", { v: 0.5, nf: '00.000%' }]
        x.write ["Duration 36 hours format #46", {v: 1.5, t: 'n', nf: '[h]:mm:ss' }]
        x.write ["Formula", {v: "ok", f: "CONCATENATE(A1,B2)"}]
        x.write ["A simple comment", {v: "cell with comment", c: 'very simple comment' }]
        x.write ["A full comment", {v: "cell with comment", c: { author: 'Joe', lines: [ { t: "bold text\n", b: true }, "plain text" ] } }]
        x.end()
        return
      "Parse xlsx":
        #topic: ->
          #office.parse(tmp('b.zip'), @callback)
          #return

        "log": (rows)->
          console.log arguments # Seems like node-xlsx does not support Inline String.
          #assert.ok()

  "Multiple sheets":
    "create":
      topic: ->
        x = xlsx_stream({ core: { creator: 'Pony Foo' }, custom: { url: 'http://localhost/doc/foo', aboolean: true, adate: new Date(), afloat: -3.14, aninteger: 42 } })
        output = fs.createWriteStream(tmp('multi.xlsx'))
        output.on 'close', @callback
        x.on 'finalize', -> console.log "FINALIZE:", arguments
        x.pipe output

        sheet1 = x.sheet("1st sheet", {
                                        frozenCell: 'C5',
                                        hiddenColumns: [ "12", "15-18", "20-"],
                                        columnsWidth:  {
                                          '1' : 26, # ~2"
                                          '2' : 30,
                                          '5' : 10
                                        }
                                      })
        # "{hidden:true  }, {}, {hidden:true  }"
        sheet1.write ["This", "is", "my", "first", "worksheet"]
        sheet1.end()

        sheet2 = x.sheet("２枚目のシート")
        sheet2.write ["これが", "２枚目の", "ワークシート", "です"]
        sheet2.end()

        sheet3 = x.sheet("Hidden Sheet", { hideSheet: true});
        sheet3.write ["This", "sheet", "is", "hidden"]
        sheet3.end()

        x.finalize()
        return

      "Parse xlsx":
        #topic: ->
          #office.parse(tmp('b.zip'), @callback)
          #return

        "log": (rows)->
          console.log arguments # Seems like node-xlsx does not support Inline String.
          # assert.ok()

  "Empty Spreadsheet":
    "create":
      topic: ->
        x = xlsx_stream()
        output = fs.createWriteStream(tmp('empty.xlsx'))
        output.on 'close', @callback
        x.pipe output
        x.end()
        return

      "log": (rows)->
        console.log arguments # Seems like node-xlsx does not support Inline String.
        #assert.ok()

).export(module, error: false)
