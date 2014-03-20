# SimpleXlsxReader

An xlsx reader for Ruby that parses xlsx cell values into plain ruby
primitives and dates/times.

This is *not* a rewrite of excel in Ruby. Font styles, for
example, are parsed to determine whether a cell is a number or a date,
then forgotten. We just want to get the data, and get out!

## Usage

### Summary:

    doc = SimpleXlsxReader.open('/path/to/workbook.xlsx')
    doc.sheets # => [<#SXR::Sheet>, ...]
    doc.sheets.first.name # 'Sheet1'
    doc.sheets.first.rows # [['Header 1', 'Header 2', ...]
                             ['foo', 2, ...]]

That's it!

### Load Errors

By default, cell load errors (ex. if a date cell contains the string
'hello') result in a SimpleXlsxReader::CellLoadError.

If you would like to provide better error feedback to your users, you
can set `SimpleXlsxReader.configuration.catch_cell_load_errors =
true`, and load errors will instead be inserted into Sheet#load_errors keyed
by [rownum, colnum].

### More

Here's the totality of the public api, in code:

    module SimpleXlsxReader
      def self.open(file_path)
        Document.new(file_path).tap(&:sheets)
      end

      class Document
        attr_reader :file_path

        def initialize(file_path)
          @file_path = file_path
        end

        def sheets
          @sheets ||= Mapper.new(xml).load_sheets
        end

        def to_hash
          sheets.inject({}) {|acc, sheet| acc[sheet.name] = sheet.rows; acc}
        end

        def xml
          Xml.load(file_path)
        end

        class Sheet < Struct.new(:name, :rows)
          def headers
            rows[0]
          end

          def data
            rows[1..-1]
          end

          # Load errors will be a hash of the form:
          # {
          #   [rownum, colnum] => '[error]'
          # }
          def load_errors
            @load_errors ||= {}
          end
        end
      end
    end

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
