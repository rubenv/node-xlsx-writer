# XLSX writer

  Simple XLSX writer. Reverse-engineered from sample XLSX files.
  
  [![Build Status](https://travis-ci.org/rubenv/node-xlsx-writer.png?branch=master)](https://travis-ci.org/rubenv/node-xlsx-writer)

## Dependency
  xlsx-writer is dependent libzip-devel (Fedora etc.) or libzip-dev (Ubuntu etc.).
  If you got any error when you try npm install, try the following:

```bash
# Ubuntu/Debian:
apt-get install libzip-dev

# Fedora/CentOS/RHEL:
yum install libzip-devel
```


## Usage
  
  You can install the latest version via npm:
  
    $ npm install --save xlsx-writer

  Require the module:

    var xlsx = require('xlsx-writer');

  Write a spreadsheet:

    var data = [
        {
            "Name": "Bob",
            "Location": "Sweden"
        },
        {
            "Name": "Alice",
            "Location": "France"
        }
    ];

    xlsx.write('mySpreadsheet.xlsx', data, function (err) {
        // Error handling here
    });

  This will write a spreadsheet like this:

    Name    | Location
    --------+---------
    Bob     | Sweden
    Alice   | France

  In other words: The key names are used for the first row (headers),
  The values are used for the columns. All field names should be present
  in the first row.

## Advanced usage

  You can also use the full API manually. This allows you to build the
  spreadsheet incrementally:

    var XlsxWriter = require('xlsx-writer');

    var writer = new XlsxWriter('mySpreadsheet.xlsx');

    // Pass the number of rows and columns:
    writer.prepare(2, 2, function (err) {
        if (err) {
            throw err;
        }

        // Add some rows
        writer.addRow({
            "Name": "Bob",
            "Location": "Sweden"
        });
        writer.addRow({
            "Name": "Alice",
            "Location": "France"
        });

        // Finalize the spreadsheet
        writer.pack(function (err) {
            if (err) {
                console.log('Something went wrong!');
            } else {
                console.log('All done!');
            }
        });
    });

## Contributing

  In lieu of a formal styleguide, take care to maintain the existing coding
  style. Add unit tests for any new or changed functionality.

  All source-code is written in CoffeeScript and is located in the `src`
  folder. Do not edit the generated files in `lib`, they will get overwritten
  (and aren't included in git anyway).

  You can build and test your code using [Grunt](http://gruntjs.com/). The
  default task will clean the source, compiled it and run the tests.

## License 

    (The MIT License)

    Copyright (C) 2013 by Ruben Vermeersch <ruben@savanne.be>

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in
    all copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
    THE SOFTWARE.
