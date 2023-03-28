# SimpleXlsxReader

A [fast](#performance) xlsx reader for Ruby that parses xlsx cell values into
plain ruby primitives and dates/times.

This is *not* a rewrite of excel in Ruby. Font styles, for
example, are parsed to determine whether a cell is a number or a date,
then forgotten. We just want to get the data, and get out!

## Summary (now with stream parsing):

    doc = SimpleXlsxReader.open('/path/to/workbook.xlsx')
    doc.sheets # => [<#SXR::Sheet>, ...]
    doc.sheets.first.name # 'Sheet1'
    rows = doc.sheet.first.rows # <SXR::Document::RowsProxy>
    rows.each # an <Enumerator> ready to chain or stream
    rows.each {} # Streams the rows to your block
    rows.each(headers: true) {} # Streams row-hashes
    rows.each(headers: {id: /ID/}) {} # finds & maps headers, streams
    rows.slurp # Slurps rows into memory as a 2D array

That's the gist of it!

See also the [Document](https://github.com/woahdae/simple_xlsx_reader/blob/2.0.0-pre/lib/simple_xlsx_reader/document.rb) object.

## Why?

### Accurate

This project was started years ago, primarily because other Ruby xlsx parsers
didn't import data with the correct types. Numbers as strings, dates as numbers,
[hyperlinks](https://github.com/woahdae/simple_xlsx_reader/blob/master/lib/simple_xlsx_reader/hyperlink.rb)
with inaccessible URLs, or - subtly buggy - simple dates as DateTime
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

Finally, as of v2.0, SimpleXlsxReader is the fastest and most
memory-efficient parser. Previously this project couldn't reasonably load
anything over ~10k rows. Other parsers could load 100k+ rows, but were still
taking ~1gb RSS to do so, even "streaming," which seemed excessive. So a SAX
implementation was born. See [performance](#performance) for details.

## Usage

### Streaming

SimpleXlsxReader is performant by default - If you use
`rows.each {|row| ...}` it will stream the XLSX rows to your block without
loading either the sheet XML or the full sheet data into memory.

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
string files are never loaded as a (big) Nokogiri doc, so that's nice.

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

### Performance

SimpleXlsxReader is (as of this writing) the fastest and most memory efficient
Ruby xlsx parser.

Recent updates here have focused on large spreadsheets with especially
non-unique strings in sheets using xlsx' shared strings feature
(Excel-generated spreadsheets always use this). Other projects have implemented
streaming parsers for the sheet data, but currently none stream while loading
the shared strings file, which is the second-largest file in an xlsx archive
and can represent millions of strings in large files.

For more details, see [my fork of @shkm's excel benchmark project](https://github.com/woahdae/excel-parsing-benchmarks), but here's the summary:

1mb excel file, 10,000 rows of sample "sales records" with a fair amount of
non-unique strings (ran on an M1 Macbook Pro):

| Gem                | Parses/second | RSS Increase | Allocated Mem | Retained Mem | Allocated Objects | Retained Objects |
|--------------------|---------------|--------------|---------------|--------------|-------------------|------------------|
| simple_xlsx_reader | 1.13          | 36.94mb      | 614.51mb      | 1.13kb       | 8796275           | 3                |
| roo                | 0.75          | 74.0mb       | 164.47mb      | 2.18kb       | 2128396           | 4                |
| creek              | 0.65          | 107.55mb     | 581.38mb      | 3.3kb        | 7240760           | 16               |
| xsv                | 0.61          | 75.66mb      | 2127.42mb     | 3.66kb       | 5922563           | 10               |
| rubyxl             | 0.27          | 373.52mb     | 716.7mb       | 2.18kb       | 10612577          | 4                |

Here is a benchmark for the "worst" file I've seen, a 26mb file whose shared
strings represent 10% of the archive (note, MemoryProfiler has too much
overhead to reasonably measure allocations so that analysis was left off, and
we just measure total time for one parse):

| Gem                | Time    | RSS Increase |
|--------------------|---------|--------------|
| simple_xlsx_reader | 28.71s  | 148.77mb     |
| roo                | 40.25s  | 1322.08mb    |
| xsv                | 45.82s  | 391.27mb     |
| creek              | 60.63s  | 886.81mb     |
| rubyxl             | 238.68s | 9136.3mb     |

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

The full suite contains a performance test that on an M1 MBP runs the final
large file in about five seconds. Check out that test before & after your
change to check for performance changes.

Then, the standard stuff:

1. Fork this project
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
