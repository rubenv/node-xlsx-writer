xlsx = require('..')

_ = require('underscore')
assert = require('assert')
fs = require('fs')
parser = require('excel')

module.exports = (name, data) ->
    describe name, ->
        filename = "tmp/#{name}.xlsx"
        result = []

        before (done) ->
            xlsx.write filename, data, (err) ->
                return done(err) if err

                parser filename, (workbook) ->
                    assert.notEqual(workbook, null)

                    result = workbook
                    done()

        it 'Should create XLSX file', ->
            assert(fs.existsSync(filename), 'file needs to exist')

        it 'Should have header row', ->
            assert(result.length >= 1, "Should have header row")

            for key, index in _.keys(data[0])
                assert.equal(result[0][index], key)

        it 'Should contain right values', ->
            assert.equal(result.length, data.length + 1)

            for row, rowNr in result
                continue if rowNr == 0 # Header
                for key, index in _.keys(data[0])
                    assert.equal(row[index], data[rowNr - 1][key])
