node-xlsx-stream
================

Creates SpreadsheetML (.xlsx) files in sequence with streaming interface.

* Installation

        npm install xlsx-stream

* Features

        Multiple sheets, String, Number, Date, Duration, Cell Formats, Frozen panes, Comments, Custom properties

* Usage

        # coffee-script
        xlsx = require "xlsx-stream"
        fs = require "fs"

        x = xlsx()
        x.pipe fs.createWriteStream("./out.xlsx")

        x.write ["foo", "bar", "buz"]
        x.write [1,2,3]
        x.write ["Date", new Date]
        x.write ["Duration", { v: 1.5, t: 'n', nf: '[h]:mm:ss' }]
        x.write ["Formula", {v: "ok", f: "CONCATENATE(A1,B2)"}]
        x.write ["Percentage Built-in format #9", { v: 0.5, nf: '0.00%' }]
        x.write ["Percentage Custom format", { v: 0.5, nf: '00.000%' }]
        x.write ["A simple comment", {v: "cell with comment", c: 'very simple comment' }]
        x.write ["A full comment", {v: "cell with comment", c: { author: 'Joe', lines: [ { t: "bold text\n", b: true }, "plain text" ] } }]

        x.end()

* Multiple sheets support

        # coffee-script

        x = xlsx()
        x.pipe fs.createWriteStream("./out.xlsx")

        sheet1 = x.sheet('first sheet', { frozenCell: "C2" })
        sheet1.write ["first", "sheet"]
        sheet1.end()

        sheet2 = x.sheet('another')
        sheet2.write ["second", "sheet"]
        sheet2.end()

        sheet3 = x.sheet("Hidden Sheet", { hideSheet: true});
        sheet3.write ["This", "sheet", "is", "hidden"]
        sheet3.end()

        x.finalize()

* Custom properties

        # coffee-script

        x = xlsx_stream({ core: { creator: 'Pony Foo' }, custom: [ name: 'url', value: 'http://localhost/doc/foo' ] })
