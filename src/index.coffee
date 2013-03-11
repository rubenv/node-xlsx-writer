fs = require('fs')
temp = require('temp')
path = require('path')
async = require('async')
zipper = require('zipper')

blobs = require('./blobs')

module.exports =
    write: (out, data, cb) ->
        tempPath = ''

        filename = (folder, name) ->
            parts = Array::slice.call(arguments)
            parts.unshift(tempPath)
            return path.join.apply(@, parts)

        zipfile = new zipper.Zipper(out)

        async.series [
            (cb) -> temp.mkdir 'xlsx', (err, p) ->
                tempPath = p
                # Debug:
                #tempPath = 'tmp'
                cb(err)
            (cb) -> fs.mkdir(filename('_rels'), cb)
            (cb) -> fs.mkdir(filename('xl'), cb)
            (cb) -> fs.mkdir(filename('xl', '_rels'), cb)
            (cb) -> fs.mkdir(filename('xl', 'worksheets'), cb)
            (cb) -> fs.writeFile(filename('[Content_Types].xml'), blobs.contentTypes, cb)
            (cb) -> fs.writeFile(filename('_rels', '.rels'), blobs.rels, cb)
            (cb) -> fs.writeFile(filename('xl', 'workbook.xml'), blobs.workbook, cb)
            (cb) -> fs.writeFile(filename('xl', 'styles.xml'), blobs.styles, cb)
            (cb) -> fs.writeFile(filename('xl', 'sharedStrings.xml'), blobs.strings, cb)
            (cb) -> fs.writeFile(filename('xl', '_rels', 'workbook.xml.rels'), blobs.workbookRels, cb)
            (cb) -> fs.writeFile(filename('xl', 'worksheets', 'sheet1.xml'), blobs.sheet, cb)
            (cb) -> zipfile.addFile(filename('[Content_Types].xml'), '[Content_Types].xml', cb)
            (cb) -> zipfile.addFile(filename('_rels', '.rels'), '_rels/.rels', cb)
            (cb) -> zipfile.addFile(filename('xl', 'workbook.xml'), 'xl/workbook.xml', cb)
            (cb) -> zipfile.addFile(filename('xl', 'styles.xml'), 'xl/styles.xml', cb)
            (cb) -> zipfile.addFile(filename('xl', 'sharedStrings.xml'), 'xl/sharedStrings.xml', cb)
            (cb) -> zipfile.addFile(filename('xl', '_rels', 'workbook.xml.rels'), 'xl/_rels/workbook.xml.rels', cb)
            (cb) -> zipfile.addFile(filename('xl', 'worksheets', 'sheet1.xml'), 'xl/worksheets/sheet1.xml', cb)
        ], cb
