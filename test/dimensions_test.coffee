XlsxWriter = require('..')

assert = require('assert')

describe 'Dimensions', ->
    writer = new XlsxWriter('tmp.xslx')

    it 'Calculates 1x1 dimensions', ->
        assert.equal(writer.dimensions(1, 1), 'A1:A1')

    it 'Calculates 2x2 dimensions', ->
        assert.equal(writer.dimensions(2, 2), 'A1:B2')

    it 'Calculates 20x26 dimensions', ->
        assert.equal(writer.dimensions(20, 26), 'A1:Z20')

    it 'Calculates 20x27 dimensions', ->
        assert.equal(writer.dimensions(20, 27), 'A1:AA20')

    it 'Calculates 20x52 dimensions', ->
        assert.equal(writer.dimensions(20, 52), 'A1:AZ20')

    it 'Calculates 20x132 dimensions', ->
        assert.equal(writer.dimensions(20, 132), 'A1:EB20')

    it 'Calculates 20x148 dimensions', ->
        assert.equal(writer.dimensions(20, 148), 'A1:ER20')
