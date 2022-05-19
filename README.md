# SimpleXlsxReader [![Build Status](https://travis-ci.org/woahdae/simple_xlsx_reader.svg?branch=master)](https://travis-ci.org/woahdae/simple_xlsx_reader)

A **fast** xlsx reader for Ruby that parses xlsx cell values into plain ruby
primitives and dates/times.

This is *not* a rewrite of excel in Ruby. Font styles, for
example, are parsed to determine whether a cell is a number or a date,
then forgotten. We just want to get the data, and get out!

## Usage

### Summary (now with stream parsing):

    doc = SimpleXlsxReader.open('/path/to/workbook.xlsx')
    doc.sheets # => [<#SXR::Sheet>, ...]
    doc.sheets.first.name # 'Sheet1'
    doc.sheets.first.rows # <SXR::Document::RowsProxy>
    doc.sheets.first.rows.each {} # Streams the rows to your block
    doc.sheets.first.rows.each(headers: true) {} # Streams row-hashes
    doc.sheets.first.rows.slurp # Slurps the rows into memory

That's the gist of it.

For all the options, see the [Document](https://github.com/woahdae/simple_xlsx_reader/blob/2.0.0-pre/lib/simple_xlsx_reader/document.rb)
object (and occompanying documentation), which is the entirety of the public
API.

## Why?

### Accurate

This was written primarily because other Ruby xlsx parsers didn't
import data with the correct types. Numbers as strings, dates as numbers,
hyperlinks with inaccessible URLs, or the worst, simple dates as DateTime
objects. If your app uses a timezone offset, then depending on what timezone and
what time of day you load the xlsx file, your data might end up a day off!
SimpleXlsxReader understands all these correctly (although this experience with
other XLSX libraries was a long time ago, maybe they have improved).

### Idiomatic

Also though, other Ruby xlsx parsers were very un-Ruby-like. Maybe it's because
they're supporting all of excel's quirky features? In any case,
SimpleXlsxReader strives to be fairly idiomatic Ruby.

### Now faster

Finally, as of v2.0, SimpleXlsxReader might be the fastest and most
memory-efficient parser. Previously this project couldn't load anything over
around 10k rows. Other parsers could load 100k+ rows, but were still taking
~1gb RSS to do so, even "streaming," which seemed excessive. So a SAX
implementation was born, bringing that particular file down to 250mb RSS (which
is mostly just holding shared strings in memory - if your sheet doesn't use
that xlsx feature, the overhead will be almost nothing, although
Excel-generated xlsx files use shared strings).

## Installation

Add this line to your application's Gemfile:

    gem 'simple_xlsx_reader'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install simple_xlsx_reader

## Versioning

This project follows [semantic versioning 1.0](http://semver.org/spec/v1.0.0.html)

## Contributing

Remember to write tests, think about edge cases, and run the existing
suite.

Note that as of commit 665cbafdde, the most extreme end of the
linear-time performance test, which is 10,000 rows (12 columns), runs in
~4 seconds on Ruby 2.1 on a 2012 MBP. If the linear time assertion fails
or you're way off that, there is probably a performance regression in
your code.

Then, the standard stuff:

1. Fork this project
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
