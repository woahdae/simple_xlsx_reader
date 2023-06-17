# frozen_string_literal: true

require 'forwardable'

module SimpleXlsxReader
  class Loader
    class SheetParser < Nokogiri::XML::SAX::Document
      extend Forwardable

      attr_accessor :xrels_file
      attr_accessor :hyperlinks_by_cell

      attr_reader :load_errors

      def_delegators :@loader, :style_types, :shared_strings, :base_date

      def initialize(file_io:, loader:)
        @file_io = file_io
        @loader = loader
      end

      def parse(headers: false, &block)
        raise 'parse called without a block; what should this do?'\
          unless block_given?

        @headers = headers
        @each_callback = block
        @load_errors = {}
        @current_row_num = nil
        @last_seen_row_idx = 0
        @url = nil # silence warnings
        @function = nil # silence warnings
        @capture = nil # silence warnings
        @captured = nil # silence warnings
        @dimension = nil # silence warnings

        # In this project this is only used for GUI-made hyperlinks (as opposed
        # to FUNCTION-based hyperlinks). Unfortunately the're needed to parse
        # the spreadsheet, and they come AFTER the sheet data. So, solution is
        # to just stream-parse the file twice, first for the hyperlinks at the
        # bottom of the file, then for the file itself. In the future it would
        # be clever to use grep to extract the xml into its own smaller file.
        if xrels_file&.grep(/hyperlink/)&.any?
          xrels_file.rewind
          load_gui_hyperlinks # represented as hyperlinks_by_cell
        end

        @file_io.rewind # in case we've already parsed this once

        Nokogiri::XML::SAX::Parser.new(self).parse(@file_io)
      end

      ###
      # SAX document hooks

      def start_element(name, attrs = [])
        case name
        when 'dimension' then @dimension = attrs.last.last
        when 'row'
          @current_row_num = attrs.find {|(k, v)| k == 'r'}&.last&.to_i
          @current_row = Array.new(column_length)
        when 'c'
          attrs = attrs.inject({}) {|acc, (k, v)| acc[k] = v; acc}
          @cell_name = attrs['r']
          @type = attrs['t']
          @style = attrs['s'] && style_types[attrs['s'].to_i]
        when 'f' then @function = true
        when 'v', 't' then @capture = true
        end
      end

      def characters(string)
        if @function
          # the only "function" we support is a hyperlink
          @url = string.slice(/HYPERLINK\("(.*?)"/, 1)
        end

        return unless @capture

        captured =
          begin
            SimpleXlsxReader::Loader.cast(
              string, @type, @style,
              url: @url || hyperlinks_by_cell&.[](@cell_name),
              shared_strings: shared_strings,
              base_date: base_date
            )
          rescue StandardError => e
            column, row = @cell_name.match(/([A-Z]+)([0-9]+)/).captures
            col_idx = column_letter_to_number(column) - 1
            row_idx = row.to_i - 1

            if !SimpleXlsxReader.configuration.catch_cell_load_errors
              error = CellLoadError.new(
                "Row #{row_idx}, Col #{col_idx}: #{e.message}"
              )
              error.set_backtrace(e.backtrace)
              raise error
            else
              @load_errors[[row_idx, col_idx]] = e.message

              string
            end
          end

        # For some reason I can't figure out in a reasonable timeframe,
        # SAX parsing some workbooks captures separate strings in the same cell
        # when we encounter UTF-8, although I can't get workbooks made in my
        # own version of excel to repro it. Our fix is just to keep building
        # the string in this case, although maybe there's a setting in Nokogiri
        # to make it not do this (looked, couldn't find it).
        #
        # Loading the workbook test/chunky_utf8.xlsx repros the issue.
        @captured = @captured ? @captured + (captured || '') : captured
      end

      def end_element(name)
        case name
        when 'row'
          if @headers == true # ya a little funky
            @headers = @current_row
          elsif @headers.is_a?(Hash)
            test_headers_hash_against_current_row
            # in case there were empty rows before finding the header
            @last_seen_row_idx = @current_row_num - 1
          elsif @headers.respond_to?(:call)
            @headers = @current_row if @headers.call(@current_row)
            # in case there were empty rows before finding the header
            @last_seen_row_idx = @current_row_num - 1
          elsif @headers
            possibly_yield_empty_rows(headers: true)
            yield_row(@current_row, headers: true)
          else
            possibly_yield_empty_rows(headers: false)
            yield_row(@current_row, headers: false)
          end

          @last_seen_row_idx += 1

          # Note that excel writes a '/worksheet/dimension' node we can get
          # this from, but some libs (ex. simple_xlsx_writer) don't record it.
          # In that case, we assume the data is of uniform column length and
          # store the column name of the last header row we see. Obviously this
          # isn't the most robust strategy, but it likely fits 99% of use cases
          # considering it's not a problem with actual excel docs.
          @dimension = "A1:#{@cell_name}" if @dimension.nil?
        when 'v', 't'
          @current_row[cell_idx] = @captured
          @capture = false
          @captured = nil
        when 'f' then @function = false
        when 'c' then @url = nil
        end
      end

      ###
      # /End SAX hooks

      def test_headers_hash_against_current_row
        found = false

        @current_row.each_with_index do |cell, cell_idx|
          @headers.each_pair do |key, search|
            if search.is_a?(String) ? cell == search : cell&.match?(search)
              found = true
              @current_row[cell_idx] = key
            end
          end
        end

        @headers = @current_row if found
      end

      def possibly_yield_empty_rows(headers:)
        while @current_row_num && @current_row_num > @last_seen_row_idx + 1
          @last_seen_row_idx += 1
          yield_row(Array.new(column_length), headers: headers)
        end
      end

      def yield_row(row, headers:)
        if headers
          @each_callback.call(Hash[@headers.zip(row)])
        else
          @each_callback.call(row)
        end
      end

      # This sax-parses the whole sheet, just to extract hyperlink refs at the end.
      def load_gui_hyperlinks
        self.hyperlinks_by_cell =
          HyperlinksParser.parse(@file_io, xrels: xrels)
      end

      class HyperlinksParser < Nokogiri::XML::SAX::Document
        def initialize(file_io, xrels:)
          @file_io = file_io
          @xrels = xrels
        end

        def self.parse(file_io, xrels:)
          new(file_io, xrels: xrels).parse
        end

        def parse
          @hyperlinks_by_cell = {}
          Nokogiri::XML::SAX::Parser.new(self).parse(@file_io)
          @hyperlinks_by_cell
        end

        def start_element(name, attrs)
          case name
          when 'hyperlink'
            attrs = attrs.inject({}) {|acc, (k, v)| acc[k] = v; acc}
            id = attrs['id'] || attrs['r:id']

            @hyperlinks_by_cell[attrs['ref']] =
              @xrels.at_xpath(%(//*[@Id="#{id}"])).attr('Target')
          end
        end
      end

      def xrels
        @xrels ||= Nokogiri::XML(xrels_file.read) if xrels_file
      end

      def column_length
        return 0 unless @dimension

        @column_length ||= column_letter_to_number(last_cell_letter)
      end

      def cell_idx
        column_letter_to_number(@cell_name.scan(/[A-Z]+/).first) - 1
      end

      ##
      # Returns the last column name, ex. 'E'
      def last_cell_letter
        return unless @dimension

        @dimension.scan(/:([A-Z]+)/)&.first&.first || 'A'
      end

      # formula fits an exponential factorial function of the form:
      # 'A'   = 1
      # 'B'   = 2
      # 'Z'   = 26
      # 'AA'  = 26 * 1  + 1
      # 'AZ'  = 26 * 1  + 26
      # 'BA'  = 26 * 2  + 1
      # 'ZA'  = 26 * 26 + 1
      # 'ZZ'  = 26 * 26 + 26
      # 'AAA' = 26 * 26 * 1 + 26 * 1  + 1
      # 'AAZ' = 26 * 26 * 1 + 26 * 1  + 26
      # 'ABA' = 26 * 26 * 1 + 26 * 2  + 1
      # 'BZA' = 26 * 26 * 2 + 26 * 26 + 1
      def column_letter_to_number(column_letter)
        pow = column_letter.length - 1
        result = 0
        column_letter.each_byte do |b|
          result += 26**pow * (b - 64)
          pow -= 1
        end
        result
      end
    end
  end
end
