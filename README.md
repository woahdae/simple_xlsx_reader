# SimpleXlsxReader [![Build Status](https://travis-ci.org/woahdae/simple_xlsx_reader.svg?branch=master)](https://travis-ci.org/woahdae/simple_xlsx_reader)

A **fast** xlsx reader for Ruby that parses xlsx cell values into plain ruby
primitives and dates/times.

This is *not* a rewrite of excel in Ruby. Font styles, for
example, are parsed to determine whether a cell is a number or a date,
then forgotten. We just want to get the data, and get out!

## Summary (now with stream parsing):

    doc = SimpleXlsxReader.open('/path/to/workbook.xlsx')
    doc.sheets # => [<#SXR::Sheet>, ...]
    doc.sheets.first.name # 'Sheet1'
    doc.sheets.first.rows # <SXR::Document::RowsProxy>
    doc.sheets.first.rows.each # an <Enumerator> ready to chain or stream
    doc.sheets.first.rows.each {} # Streams the rows to your block
    doc.sheets.first.rows.each(headers: true) {} # Streams row-hashes
    doc.sheets.first.rows.each(headers: {id: /ID/}) {} # finds & maps headers, streams
    doc.sheets.first.rows.slurp # Slurps rows into memory as a 2D array

That's the gist of it!

See also the [Document](https://github.com/woahdae/simple_xlsx_reader/blob/2.0.0-pre/lib/simple_xlsx_reader/document.rb) object.

## Why?

### Accurate

This project was started years ago, primarily because other Ruby xlsx parsers
didn't import data with the correct types. Numbers as strings, dates as numbers,
hyperlinks with inaccessible URLs, or - subtly buggy - simple dates as DateTime
objects. If your app uses a timezone offset, depending on what timezone and
what time of day you load the xlsx file, your dates might end up a day off!
SimpleXlsxReader understands all these correctly.

### Idiomatic

Many Ruby xlsx parsers seem to be inspired more by Excel than Ruby, frankly.
SimpleXlsxReader strives to be fairly idiomatic Ruby:

    # quick example having fun w/ ruby
    doc = SimpleXlsxReader.open(path_or_io)
    doc.sheets.first.rows.each(headers: {id: /ID/})
      .with_index.with_object({}) do |(row, index), acc|
        acc[row[:id]] = index
      end

### Now faster

Finally, as of v2.0, SimpleXlsxReader might be the fastest and most
memory-efficient parser. Previously this project couldn't reasonably load
anything over ~10k rows. Other parsers could load 100k+ rows, but were still
taking ~1gb RSS to do so, even "streaming," which seemed excessive. So a SAX
implementation was born, bringing some 100k-row files down to ~200mb RSS (which
is mostly just holding shared strings in memory - if your sheet doesn't use
that xlsx feature, the overhead will be almost nothing, although
Excel-generated xlsx files do use shared strings).

## Usage

### Streaming

SimpleXlsxReader is performant by default - If you use
`rows.each {|row| ...}` it will stream the XLSX rows to your block without
loading either the sheet XML or the row data into memory.*

You can also chain `rows.each` with other Enumerable functions without
triggering a slurp, and you have lots of ways to find and map headers while
streaming.

If you had an excel sheet representing this data:

```
| Hero ID | Hero Name  | Location     |
| 13576   | Samus Aran | Planet Zebes |
| 117     | John Halo  | Ring World   |
| 9704133 | Iron Man   | Planet Earth |
```

Get a handle on the rows proxy:

`rows = SimpleXlsxReader.open('suited_heroes.xlsx').sheets.first.rows`

Simple streaming (kinda boring):

`rows.each { |row| ... }`

Streaming with headers, and how about a little enumerable chaining:

```
# Map of hero names by ID: { 117 => 'John Halo', ... }

rows.each(headers: true).with_object({}) do |row, acc|
  acc[row['Hero ID']] = row['Hero Name']
end
```

Sometimes though you have some junk at the top of your spreadsheet:

```
| Unofficial Report  |                        |              |
| Dont tell Nintendo | Yes "John Halo" I know |              |
|                    |                        |              |
| Hero ID            | Hero Name              | Location     |
| 13576              | Samus Aran             | Planet Zebes |
| 117                | John Halo              | Ring World   |
| 9704133            | Iron Man               | Planet Earth |
```

For this, `headers` can be a hash whose keys replace headers and whose values
help find the correct header row:

```
# Same map of hero names by ID: { 117 => 'John Halo', ... }

rows.each(headers: {id: /ID/, name: /Name/}).with_object({}) do |row, acc|
  acc[row[:id]] = row[:name]
end
```

If your header-to-attribute mapping is more complicated than key/value, you
can do the mapping elsewhere, but use a block to find the header row:

```
# Example roughly analogous to some production code mapping a single spreadsheet
# across many objects. Might be a simpler way now that we have the headers-hash
# feature.

object_map = { Hero => { id: 'Hero ID', name: 'Hero Name', location: 'Location' } }

HEADERS = ['Hero ID', 'Hero Name', 'Location']

rows.each(headers: ->(row) { (HEADERS & row).any? }) do |row|
  object_map.each_pair do |klass, attribute_map|
    attributes =
      attribute_map.each_pair.with_object({}) do |(key, header), attrs|
        attrs[key] = row[header]
      end

    klass.new(attributes)
  end
end
```

### Slurping

To make SimpleXlsxReader rows act like an array, for use with legacy
SimpleXlsxReader apps or otherwise, we still support slurping the whole array
into memory. The good news is even when doing this, the xlsx worksheet & shared
string files are never slurped into Nokogiri, so that's nice.

By default, to prevent accidental slurping, `<RowsProxy>` will throw an exception
if you try to access it with array methods like `[]` and `shift` without
explicitly slurping first. You can slurp either by calling `rows.slurp` or
globally by setting `SimpleXlsxReader.configuration.auto_slurp = true`.

Once slurped, enumerable methods on `rows` will use the slurped data
(i.e. not re-parse the sheet), and those Array-like methods will work.

We don't support all Array methods, just the few we have used in real projects,
as we transition towards streaming instead.

### Load Errors

By default, cell load errors (ex. if a date cell contains the string
'hello') result in a SimpleXlsxReader::CellLoadError.

If you would like to provide better error feedback to your users, you
can set `SimpleXlsxReader.configuration.catch_cell_load_errors =
true`, and load errors will instead be inserted into Sheet#load_errors keyed
by [rownum, colnum]:

    {
      [rownum, colnum] => '[error]'
    }

### * Streaming loads "shared strings" into memory

SpreadsheetML, which Excel uses, has an optional feature where it will store
string-type cell values in a separate, workbook-wide XML sheet, and the
sheet XML files will reference the shared strings instead of storing the value
directly.

Excel seems to *always* use this feature, and while it potentially makes
the xlsx files themselves smaller, it makes stream parsing the files more
memory-intensive because we have to load the whole shared strings reference
table before parsing the main sheets. At least now it does so without slurping
the Nokogiri representation into memory.

For large files, say 100k rows and 20 columns, the shared strings array can be a
million strings and ~200mb. If someone has a clever idea about making this
more memory efficient, speak up!

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
