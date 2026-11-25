# frozen_string_literal: true

require 'forwardable'

module SimpleXlsxReader

  ##
  # Main class for the public API. See the README for usage examples,
  # or read the code, it's pretty friendly.
  class Document
    attr_reader :string_or_io

    def initialize(legacy_file_path = nil, file_path: nil, string_or_io: nil)
      fail(ArgumentError, 'either file_path or string_or_io must be provided') if legacy_file_path.nil? && file_path.nil? && string_or_io.nil?
  
      @string_or_io = string_or_io || File.new(legacy_file_path || file_path)
    end

    def sheets
      @sheets ||= Loader.new(string_or_io).init_sheets
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
      # * hash - transforms the header row by replacing cells with keys matched
      #   by value, ex. `{id: /ID|Identity/, name: /Name/i, date: 'Date'}` would
      #   potentially yield the row `{id: 5, name: 'Jane', date: [Date object]}`
      #   instead of the headers from the sheet. It would also search for the
      #   row that matches at least one header, in case the header row isn't the
      #   first.
      #
      # `column_styles` overrides the style of a column's number cells, which
      # otherwise comes from their number format and decides how they are
      # cast. Keys are headers as yielded, or column indexes when not using
      # `headers:`; a key that matches no column is ignored. Styles:
      #
      # * `:text` - the number's digits as a string: "99770064", never
      #   99770064.0 or "9.9770064E7"
      # * `:float` - a Float
      # * `:fixnum` - an Integer
      # * `:date` - a Date, or a Time if the number has a time part
      #
      # Text cells are not changed. Streaming only, like `headers:`.
      #
      # If rows have been slurped, #each will iterate the slurped rows instead.
      #
      # Note, calls to this after slurping will raise if given the `headers:`
      # argument, as that's handled by the sheet parser. If this is important
      # to someone, speak up and we could potentially support it.
      def each(headers: false, column_styles: nil, &block)
        if slurped?
          raise '#each does not support headers with slurped rows' if headers
          raise '#each does not support column_styles with slurped rows' if column_styles&.any?

          slurped.each(&block)
        elsif block_given?
          # It's possible to slurp while yielding to the block, which would
          # null out @sheet_parser, so let's just keep track of it here too
          sheet_parser = @sheet_parser
          @sheet_parser.parse(headers: headers, column_styles: column_styles, &block).tap do
            @load_errors = sheet_parser.load_errors
          end
        else
          to_enum(:each, headers: headers, column_styles: column_styles)
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
