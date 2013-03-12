fs = require('fs')
temp = require('temp')
path = require('path')
async = require('async')
zipper = require('zipper')

blobs = require('./blobs')

numberRegex = /^[1-9\.][\d\.]+$/

class XlsxWriter
    @write = (out, data, cb) ->
        rows = data.length
        columns = 0
        columns += 1 for key of data[0]

        writer = new XlsxWriter(out)
        writer.prepare rows, columns, (err) ->
            return cb(err) if err
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

        @tempPath = ''

        @sheetStream = null

        @cellMap = []

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

    prepare: (rows, columns, cb) ->
        # Add one extra row for the header
        dimensions = @dimensions(rows + 1, columns)

        async.series [
            (cb) => temp.mkdir 'xlsx', (err, p) =>
                @tempPath = p
                cb(err)
            (cb) => fs.mkdir(@_filename('_rels'), cb)
            (cb) => fs.mkdir(@_filename('xl'), cb)
            (cb) => fs.mkdir(@_filename('xl', '_rels'), cb)
            (cb) => fs.mkdir(@_filename('xl', 'worksheets'), cb)
            (cb) => fs.writeFile(@_filename('[Content_Types].xml'), blobs.contentTypes, cb)
            (cb) => fs.writeFile(@_filename('_rels', '.rels'), blobs.rels, cb)
            (cb) => fs.writeFile(@_filename('xl', 'workbook.xml'), blobs.workbook, cb)
            (cb) => fs.writeFile(@_filename('xl', 'styles.xml'), blobs.styles, cb)
            (cb) => fs.writeFile(@_filename('xl', '_rels', 'workbook.xml.rels'), blobs.workbookRels, cb)
            (cb) =>
                @sheetStream = fs.createWriteStream(@_filename('xl', 'worksheets', 'sheet1.xml'))
                @sheetStream.write(blobs.sheetHeader(dimensions))
                cb()
        ], (err) =>
            @prepared = true
            cb(err)

    pack: (cb) ->
        throw Error('Should call prepare() first!') if !@prepared

        zipfile = new zipper.Zipper(@out)

        async.series [
            (cb) =>
                @sheetStream.write(blobs.sheetFooter)
                @sheetStream.end(cb)
            (cb) =>
                stream = fs.createWriteStream(@_filename('xl', 'sharedStrings.xml'))
                stream.write(blobs.stringsHeader(@strings.length))
                for string in @strings
                    stream.write(blobs.string(@escapeXml(string)))
                stream.write(blobs.stringsFooter)
                stream.end(cb)
            (cb) => zipfile.addFile(@_filename('[Content_Types].xml'), '[Content_Types].xml', cb)
            (cb) => zipfile.addFile(@_filename('_rels', '.rels'), '_rels/.rels', cb)
            (cb) => zipfile.addFile(@_filename('xl', 'workbook.xml'), 'xl/workbook.xml', cb)
            (cb) => zipfile.addFile(@_filename('xl', 'styles.xml'), 'xl/styles.xml', cb)
            (cb) => zipfile.addFile(@_filename('xl', 'sharedStrings.xml'), 'xl/sharedStrings.xml', cb)
            (cb) => zipfile.addFile(@_filename('xl', '_rels', 'workbook.xml.rels'), 'xl/_rels/workbook.xml.rels', cb)
            (cb) => zipfile.addFile(@_filename('xl', 'worksheets', 'sheet1.xml'), 'xl/worksheets/sheet1.xml', cb)
        ], cb

    dimensions: (rows, columns) ->
        return "A1:#{@cell(rows, columns)}"

    cell: (row, col) ->
        colIndex = ''
        input = (+col - 1).toString(26)
        while input.length
            a = input.charCodeAt(input.length - 1)
            colIndex = String.fromCharCode(a + if a >= 48 and a <= 57 then 17 else -22) + colIndex
            input = if input.length > 1 then (parseInt(input.substr(0, input.length - 1), 26) - 1).toString(26) else ""
        return "#{colIndex}#{row}"

    _filename: (folder, name) ->
        parts = Array::slice.call(arguments)
        parts.unshift(@tempPath)
        return path.join.apply(@, parts)

    _startRow: () ->
        @sheetStream.write(blobs.startRow(@currentRow))
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
            @sheetStream.write(blobs.numberCell(value, cell))
        else
            index = @_lookupString(value)
            @sheetStream.write(blobs.cell(index, cell))

    _endRow: () ->
        @sheetStream.write(blobs.endRow)

    escapeXml: (str = '') ->
        return str.replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')

module.exports = XlsxWriter
