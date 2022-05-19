# frozen_string_literal: true

module SimpleXlsxReader

  ##
  # Everything in Document is the public API.
  #
  # ### Basic Usage
  #
  #     doc = SimpleXlsxReader.open('/path/to/workbook.xlsx')
  #     doc.sheets # => [<#SXR::Sheet>, ...]
  #     doc.sheets.first.name # 'Sheet1'
  #     doc.sheets.first.rows # <SXR::Document::RowsProxy>
  #     doc.sheets.first.rows.each {} # Streams the rows to your block
  #     doc.sheets.first.rows.each(headers: true) {} # Streams row-hashes
  #     doc.sheets.first.rows.slurp # Slurps the rows into memory
  #
  # ### On Streaming vs Slurping
  #
  # SimpleXlsxReader is performant by default - If you use
  # `rows.each {|row| ...}` it will stream the XLSX rows to your block without
  # loading either the sheet XML or the row data into memory.*
  #
  # By default, to prevent accidental slurping, it will also throw an exception
  # if you try to access `rows` like an array, as you used to be able to do
  # pre-2.0. You can change this to pre-2.0 behavior by either calling
  # `rows.slurp` or globally by setting
  # `SimpleXlsxReader.configuration.auto_slurp = true`.
  #
  # Once slurped, methods on `rows` (`#each`, `#map`) will use the slurped data
  # and not re-parse the sheet. Additionally, other methods will no longer
  # raise, such as `#shift` and `#[]`. This is mostly for legacy support,
  # though.
  #
  # A common reason to want to slurp the rows into memory is to find the headers,
  # which might not be the first row.
  #
  # To reduce the need to slurp rows, `rows.each` accepts a `headers` parameter
  # with a couple options to help get the header row:
  #
  # * `true` - simply takes the first row as the header row
  # * block - calls the block with successive rows until the block returns true,
  #   which it then uses that row for the headers.
  #
  # If any header option is given, the block will yield a hash of header -> value
  # instead of an array of values.
  #
  # ### * It loads "shared strings" into memory
  #
  # SpreadsheetML, which Excel uses, has an optional feature where it will store
  # string-type cell values in a separate, workbook-wide XML sheet, and the
  # sheet XML files reference the shared strings instead of storing the value
  # directly.
  #
  # Excel seems to *always* use this feature, and while it potentially makes
  # the xlsx files themselves smaller, it makes stream parsing the files more
  # memory-intensive because we have to load the whole reference table before
  # parsing the main sheets. SimpleXlsxReader loads them into memory, although
  # it does so without slurping the *XML* into memory, so that's nice.
  #
  # For large files, say 100k rows and 20 columns, it can be a million strings
  # and ~200mb. If someone has a clever idea about making this string
  # dictionary more memory efficient, speak up!
  #
  # ### Load Errors
  #
  # By default, cell load errors (ex. if a date cell contains the string
  # 'hello') result in a SimpleXlsxReader::CellLoadError.
  #
  # If you would like to provide better error feedback to your users, you
  # can set `SimpleXlsxReader.configuration.catch_cell_load_errors =
  # true`, and load errors will instead be inserted into Sheet#load_errors keyed
  # by [rownum, colnum]:
  #
  #     {
  #       [rownum, colnum] => '[error]'
  #     }
  #
  class Document
    attr_reader :file_path

    def initialize(file_path)
      @file_path = file_path
    end

    def sheets
      @sheets ||= Loader.new(file_path).init_sheets
    end

    # Expensive because it slurps all the sheets into memory,
    # probably only appropriate for testing
    def to_hash
      sheets.each_with_object({}) { |sheet, acc| acc[sheet.name] = sheet.rows.to_a; }
    end

    # `rows` is a RowsProxy that responds to #each
    class Sheet
      extend Forwardable

      attr_reader :name, :rows

      def_delegators :rows, :load_errors, :slurp

      def initialize(name:, sheet_parser:)
        @name = name
        @rows = RowsProxy.new(sheet_parser: sheet_parser)
      end

      # Legacy - consider `rows.each(headers: true)` for better performance
      def headers
        rows.slurped![0]
      end

      # Legacy - consider `rows` or `rows.each(headers: true)` for better
      # performance
      def data
        rows.slurped![1..-1]
      end
    end

    # Waits until we call #each with a block to parse the rows
    class RowsProxy
      include Enumerable

      attr_reader :slurped, :load_errors

      def initialize(sheet_parser:)
        @sheet_parser = sheet_parser
        @slurped = nil
        @load_errors = {}
      end

      # By default, #each streams the rows to the provided block, either as
      # arrays, or as header => cell value pairs if provided a `headers:`
      # argument.
      #
      # `headers` can be:
      #
      # * `true` - simply takes the first row as the header row
      # * block - calls the block with successive rows until the block returns
      #   true, which it then uses that row for the headers. All data prior to
      #   finding the headers is ignored.
      #
      # If rows have been slurped, #each will iterate the slurped rows instead.
      #
      # Note, calls to this after slurping will raise if given the `headers:`
      # argument, as that's handled by the sheet parser. If this is important
      # to someone, speak up and we could support it.
      def each(headers: false, &block)
        if slurped?
          raise '#each does not support headers with slurped rows' if headers

          slurped.each(&block)
        elsif block_given?
          @sheet_parser.parse(headers: headers, &block).tap do
            @load_errors = @sheet_parser.load_errors
          end
        else
          to_enum(:each)
        end
      end

      # Mostly for legacy support, I'm not aware of a use case for doing this
      # when you don't have to.
      #
      # Note that #each will use slurped results if available, and since we're
      # leveraging Enumerable, all the other Enumerable methods will too.
      def slurp
        # possibly release sheet parser from memory on next GC run;
        # untested, but it can hold a lot of stuff, so worth a try
        @slurped ||= to_a.tap { @sheet_parser = nil }
      end

      def slurped?
        !!@slurped
      end

      def slurped!
        check_slurped

        slurped
      end

      def [](*args)
        check_slurped

        slurped[*args]
      end

      def shift(*args)
        check_slurped

        slurped.shift(*args)
      end

      private

      def check_slurped
        slurp if SimpleXlsxReader.configuration.auto_slurp
        return if slurped?

        raise 'Called a slurp-y method without explicitly slurping;'\
          ' use #each or call rows.slurp first'
      end
    end
  end
end
