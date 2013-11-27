fs = require('fs')
path = require('path')
Zip = require('node-zip')

blobs = require('./blobs')

numberRegex = /^[1-9\.][\d\.]+$/

class XlsxWriter
    @write = (out, data, cb) ->
        rows = data.length
        columns = 0
        columns += 1 for key of data[0]

        writer = new XlsxWriter(out)
        writer.prepare(rows, columns)

        for row in data
            writer.addRow(row)
        writer.pack(cb)

    constructor: (@out) ->
        @strings = []
        @stringMap = {}
        @stringIndex = 0
        @currentRow = 0

        @haveHeader = false
        @prepared = false

        @cellMap = []
        @cellLabelMap = {}

    addRow: (obj) ->
        throw Error('Should call prepare() first!') if !@prepared

        if !@haveHeader
            @_startRow()
            col = 1
            for key of obj
                @_addCell(key, col)
                @cellMap.push(key)
                col += 1
            @_endRow()

            @haveHeader = true

        @_startRow()
        for key, col in @cellMap
            @_addCell(obj[key] || "", col + 1)
        @_endRow()

    prepare: (rows, columns) ->
        # Add one extra row for the header
        dimensions = @dimensions(rows + 1, columns)

        # Create header and mark this sheet as ready
        @sheetData = blobs.sheetHeader(dimensions)
        @prepared = true

    pack: (cb) ->
        throw Error('Should call prepare() first!') if !@prepared

        # Create Zip (JSZip port, no native deps)
        zipFile = new Zip()

        # Add static supporting files
        zipFile.file('[Content_Types].xml', blobs.contentTypes)
        zipFile.file('_rels/.rels', blobs.rels)
        zipFile.file('xl/workbook.xml', blobs.workbook)
        zipFile.file('xl/styles.xml', blobs.styles)
        zipFile.file('xl/_rels/workbook.xml.rels', blobs.workbookRels)

        # Add shared strings
        stringTable = ''
        for string in @strings
            stringTable += blobs.string(@escapeXml(string))
        stringsData = blobs.stringsHeader(@strings.length) + stringTable + blobs.stringsFooter
        zipFile.file('xl/sharedStrings.xml', stringsData)

        # Append footer to sheet
        @sheetData += blobs.sheetFooter

        # Add sheet
        zipFile.file('xl/worksheets/sheet1.xml', @sheetData)

        # Pack it up
        results = zipFile.generate({
            base64: false
            compression: 'DEFLATE'    
        })

        # Write to output location
        fs.writeFile(@out, results, 'binary', cb)

    dimensions: (rows, columns) ->
        return "A1:" + @cell(rows, columns)

    cell: (row, col) ->
        colIndex = ''
        if @cellLabelMap[col]
            colIndex = @cellLabelMap[col]
        else
            if col == 0
                # Provide a fallback for empty spreadsheets
                row = 1
                col = 1

            input = (+col - 1).toString(26)
            while input.length
                a = input.charCodeAt(input.length - 1)
                colIndex = String.fromCharCode(a + if a >= 48 and a <= 57 then 17 else -22) + colIndex
                input = if input.length > 1 then (parseInt(input.substr(0, input.length - 1), 26) - 1).toString(26) else ""
            @cellLabelMap[col] = colIndex

        return colIndex + row

    _startRow: () ->
        @rowBuffer = blobs.startRow(@currentRow)
        @currentRow += 1

    _lookupString: (value) ->
        if !@stringMap[value]
            @stringMap[value] = @stringIndex
            @strings.push(value)
            @stringIndex += 1
        return @stringMap[value]

    _addCell: (value = '', col) ->
        row = @currentRow
        cell = @cell(row, col)

        if numberRegex.test(value)
            @rowBuffer += blobs.numberCell(value, cell)
        else
            index = @_lookupString(value)
            @rowBuffer += blobs.cell(index, cell)

    _endRow: () ->
        @sheetData += @rowBuffer + blobs.endRow

    escapeXml: (str = '') ->
        return str.replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')

module.exports = XlsxWriter
