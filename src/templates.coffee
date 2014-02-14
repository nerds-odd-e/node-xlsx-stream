utils = require('./utils')

xml = utils.compress
esc = utils.escapeXML

module.exports =
  # worksheet
  worksheet:
    header: (opts)->
      options = opts || {}
      if options.frozenCell
        frozenCell = options.frozenCell
        decoded = utils.cellDecode frozenCell
        sheetView = """
          <sheetView workbookViewId="0">
            <pane topLeftCell="#{frozenCell}" ySplit="#{decoded.row - 1}.0" xSplit="#{decoded.col - 1}.0" activePane="bottomRight" state="frozen" />
          </sheetView>
        """
      else
        sheetView = '<sheetView workbookViewId="0"/>'

      colsTag = ""
      if options.hiddenColumns
        for rangePattern, i in options.hiddenColumns
          # 1025 is the maximum column count according to the spec
          minMax = utils.rangePatternToMinMax(rangePattern, 1025)
          colsTag += """<col hidden="true" min="#{minMax[0]}" max="#{minMax[1]}" width="0" />"""

      if options.columnsWidth
        for colIdx, width of options.columnsWidth
          colsTag += """<col min="#{colIdx}" max="#{colIdx}" width="#{width}" />""" # 1" = 12.959

      xml """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
          <sheetViews>
            #{sheetView}
          </sheetViews>
          <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
          <cols>
            #{colsTag}
          </cols>
          <sheetData>
      """
    footer: (sheet)-> xml """
        </sheetData>#{if sheet.comments.length then '\n<legacyDrawing r:id="rId1" />' else ''}
      </worksheet>
    """

  # Static files
  sheet_related:
    "[Content_Types].xml":
      header: (opts)->
        if opts.custom
          customProps = '\n<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>'
        else
          customProps = ''

        xml """
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
            <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
            <Default Extension="xml" ContentType="application/xml"/>
            <Default ContentType="image/jpeg" Extension="jpeg"/>
            <Default ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing" Extension="vml"/>
            <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
            <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
            <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
            <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
            <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>#{customProps}
        """
      sheet: (sheet)->
        buf = """
          <Override PartName="/#{esc sheet.path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        """
        buf += """
          <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml" PartName="/xl/comments#{sheet.index}.xml"/>
        """ if sheet.comments.length
        return buf
      footer: xml """
        </Types>
      """

    "xl/_rels/workbook.xml.rels":
      header: (opts)->
        xml """
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        """
      sheet: (sheet)-> """
          <Relationship Id="rSheet#{esc sheet.index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="#{esc sheet.rel}"/>
          """
      footer: xml """
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        </Relationships>
      """

    "xl/workbook.xml":
      header: (opts)->
        xml """
          <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9303"/>
            <workbookPr defaultThemeVersion="124226"/>
            <bookViews>
            <workbookView xWindow="480" yWindow="60" windowWidth="18195" windowHeight="8505"/>
            </bookViews>
            <sheets>
        """
      sheet: (sheet)-> xml """
            <sheet name="#{esc sheet.name}" sheetId="#{esc sheet.index}" r:id="rSheet#{esc sheet.index}"/>
      """
      footer: xml """
          </sheets>
          <calcPr calcId="145621"/>
        </workbook>
      """

  # Styles file
  styles: (styl)->
    numFmtItems = ""
    for item in styl.numFmts
      numFmtItems += "  <numFmt numFmtId=\"#{item.numFmtId}\" formatCode=\"#{esc(item.formatCode)}\" />\n"
    numFmts = if numFmtItems then """
      <numFmts count="#{styl.numFmts.length}">
        #{numFmtItems}</numFmts>
    """ else ""

    cellXfItems = ""
    for item in styl.cellStyleXfs
      cellXfItems += "  <xf xfId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" numFmtId=\"#{item.numFmtId}\" applyNumberFormat=\"1\"/>\n"
    cellXfs = if cellXfItems then """
      <cellXfs count="#{Object.keys(styl.cellStyleXfs).length}">
        #{cellXfItems}
      </cellXfs>
    """ else ""

    xml """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
        #{numFmts}
        <fonts count="1" x14ac:knownFonts="1">
          <font>
            <sz val="11"/>
            <color theme="1"/>
            <name val="Calibri"/>
            <family val="2"/>
            <scheme val="minor"/>
          </font>
        </fonts>
        <fills count="2">
          <fill>
            <patternFill patternType="none"/>
          </fill>
          <fill>
            <patternFill patternType="gray125"/>
          </fill>
        </fills>
        <borders count="1">
          <border>
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal/>
          </border>
        </borders>
        #{cellXfs}
        <cellStyles count="1">
          <cellStyle name="Normal" xfId="0" builtinId="0"/>
        </cellStyles>
        <dxfs count="0"/>
        <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
        <extLst>
          <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/>
          </ext>
        </extLst>
      </styleSheet>
    """

  # Static files
  statics:
    "xl/sharedStrings.xml": xml """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>
    """

    "docProps/app.xml": xml """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
        <Application>node-xlsx-stream</Application>
        <DocSecurity>0</DocSecurity>
        <ScaleCrop>false</ScaleCrop>
        <Company></Company>
        <LinksUpToDate>false</LinksUpToDate>
        <SharedDoc>false</SharedDoc>
        <HyperlinksChanged>false</HyperlinksChanged>
        <AppVersion>#{require('../package.json').version}</AppVersion>
      </Properties>
      """

  semiStatics:
    "_rels/.rels": (opts)->
      if opts.custom
        customProps = """
            <Relationship Id="rId4" Target="docProps/custom.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"/>
        """
      else
        customProps = ""

      xml """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
          #{customProps}
        </Relationships>
      """

    "docProps/core.xml": (opts)->
      opts = {} unless opts
      coreProps = opts.core || {}
      today = new Date().toISOString()
      extra = if coreProps.title then "<dc:title>#{esc(coreProps.title)}</dc:title>\n  " else ""
      extra += "<dc:subject>#{esc(coreProps.subject)}</dc:subject>\n  " if coreProps.subject
      extra += "<dc:creator>#{if coreProps.creator then esc(coreProps.creator) else 'node-xlsx-stream'}</dc:creator>\n  "
      extra += "<cp:lastModifiedBy>#{if coreProps.lastModifiedBy then esc(coreProps.lastModifiedBy) else 'node-xlsx-stream'}</cp:lastModifiedBy>\n  "
      extra += "<dc:description>#{esc(coreProps.description)}</dc:description>\n  " if coreProps.description

      """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
        #{extra}<dcterms:created xsi:type="dcterms:W3CDTF">#{today}</dcterms:created>
        <dcterms:modified xsi:type="dcterms:W3CDTF">#{today}</dcterms:modified>
      </cp:coreProperties>
      """

    "docProps/custom.xml": (opts)->
      return unless opts and opts.custom
      extra = ""
      index = 1
      getVTypeProperty = (name, value, index)->
        if typeof value == 'string'
          type = 'lpwstr'
          value = esc(value)
        else if typeof value == 'boolean'
          type = 'bool'
        else if typeof value == 'number'
          if parseInt(value,10) == value
            type = 'i4'
          else
            type = 'r8'
        else if value instanceof Date
          type = 'filetime'
          value = value.toISOString()
        else if value == null
          type = 'null'
        else
          type = 'lpwstr'
          value = value.toString()
        """
          <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" name="#{name}" pid="#{index}">
            <vt:#{type}>#{value}</vt:#{type}>
          </property>
        """

      if Array.isArray(opts.custom)
        for prop in opts.custom
          index++
          extra += getVTypeProperty(prop.name, prop.value, index)
      else
        for key, value of opts.custom
          index++
          extra += getVTypeProperty(key, value, index)

      xml """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
        #{extra}
      </Properties>
      """

  sheetRels: (sheet)->
    relBuf = if sheet.comments.length then """
      <Relationship Id="rId1" Target="../drawings/vmlDrawing#{sheet.index}.vml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"/>
      <Relationship Id="rId2" Target="../comments#{sheet.index}.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"/>
    """ else ""
    xml """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">#{relBuf}</Relationships>
    """

  vmlDrawing: (sheet)->
    header = xml """
      <xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:mv="http://macVmlSchemaUri">
        <o:shapelayout v:ext="edit">
         <o:idmap v:ext="edit" data="1"/>
        </o:shapelayout>
        <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m0,0l0,21600,21600,21600,21600,0xe">
          <v:stroke joinstyle="miter"/>
          <v:path gradientshapeok="t" o:connecttype="rect"/>
        </v:shapetype>
      """
    shape = (a1Notation)->
      decoded = utils.cellDecode a1Notation
      rowNumber = decoded.row
      colNumber = decoded.col
      sheet.shapeCounter++;
      unique_id = "_x#{sheet.index}_s#{sheet.shapeCounter}"
      point_from_left = colNumber * 100 + 30 # wild guess
      point_from_top = rowNumber * 20 + 5 # wild guess
      xml """
        <v:shape id="#{unique_id}" type="#_x0000_t202" style='position:absolute;margin-left:"#{point_from_left}"pt;margin-top:"#{point_from_top}"pt;width:104pt;height:64pt;z-index:#{sheet.shapeCounter};visibility:hidden;mso-wrap-style:tight' fillcolor="#fbf6d6" strokecolor="#edeaa1">
          <v:fill color2="#fbfe82" angle="-180" type="gradient">
           <o:fill v:ext="view" type="gradientUnscaled"/>
          </v:fill>
          <v:shadow on="t" obscured="t"/>
          <v:path o:connecttype="none"/>
          <v:textbox>
           <div style='text-align:left'></div>
          </v:textbox>
          <x:ClientData ObjectType="Note">
           <x:MoveWithCells/>
           <x:SizeWithCells/>
           <x:Anchor>
            #{colNumber}, 15, #{rowNumber}, 2, #{colNumber+3}, 54, #{rowNumber+3}, 4</x:Anchor>
           <x:AutoFill>False</x:AutoFill>
           <x:Row>#{rowNumber-1}</x:Row>
           <x:Column>#{colNumber-1}</x:Column>
          </x:ClientData>
         </v:shape>
      """
    shapes = ()->
      buffer = ""
      for comment in sheet.comments
        buffer += shape(comment.ref)
    footer = xml """
      </xml>
      """
    return xml(header + shapes() + footer)

  comments: (sheet)->
    authors = ""
    sheet.authors.push "" if sheet.authors.length == 0
    authors += "<author>#{author}</author>" for author in sheet.authors
    header = xml """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <authors>
        #{authors}
      </authors>
      <commentList>
    """
    body = (sheet)->
      generateOneLine = (line)->
        """
          <r>
            <rPr>#{if line.b then '\n    <b/>' else ''}
              <sz val="9"/>
              <color indexed="81"/>
              <rFont val="Calibri"/>
              <family val="2"/>
            </rPr>
            <t xml:space="preserve">#{ if line.t then line.t else line }</t>
          </r>
        """
      generateOneComment= (comment)->
        authorId = if typeof comment.authorId == 'number' then comment.authorId else 0
        lines = ""
        for line in comment.lines
          lines += generateOneLine line
        """
          <comment authorId="#{authorId}" ref="#{comment.ref}">
            <text>#{lines}</text>
          </comment>
        """

      all = ""
      for comment in sheet.comments
        all += generateOneComment comment

      all

    footer = """
      </commentList>
      </comments>
    """
    return header + '\n' + body(sheet) + '\n' + footer
