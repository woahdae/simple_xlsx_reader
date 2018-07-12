require "simple_xlsx_reader/version"
require 'nokogiri'
require 'date'

# Rubyzip 1.0 only has different naming, everything else is the same, so let's
# be flexible so we don't force people into a dependency hell w/ other gems.
begin
  # Try loading rubyzip < 1.0
  require 'zip/zip'
  require 'zip/zipfilesystem'
  SimpleXlsxReader::Zip = Zip::ZipFile
rescue LoadError
  # Try loading rubyzip >= 1.0
  require 'zip'
  require 'zip/filesystem'
  SimpleXlsxReader::Zip = Zip::File
end

module SimpleXlsxReader
  class CellLoadError < StandardError; end

  # We support hyperlinks as a "type" even though they're technically
  # represented either as a function or an external reference in the xlsx spec.
  #
  # Since having hyperlink data in our sheet usually means we might want to do
  # something primarily with the URL (store it in the database, download it, etc),
  # we go through extra effort to parse the function or follow the reference
  # to represent the hyperlink primarily as a URL. However, maybe we do want
  # the hyperlink "friendly name" part (as MS calls it), so here we've subclassed
  # string to tack on the friendly name. This means 80% of us that just want
  # the URL value will have to do nothing extra, but the 20% that might want the
  # friendly name can access it.
  #
  # Note, by default, the value we would get by just asking the cell would
  # be the "friendly name" and *not* the URL, which is tucked away in the
  # function definition or a separate "relationships" meta-document.
  #
  # See MS documentation on the HYPERLINK function for some background:
  # https://support.office.com/en-us/article/HYPERLINK-function-333c7ce6-c5ae-4164-9c47-7de9b76f577f
  class Hyperlink < String
    attr_reader :friendly_name

    def initialize(url, friendly_name = nil)
      @friendly_name = friendly_name
      super(url)
    end
  end

  def self.configuration
    @configuration ||= Struct.new(:catch_cell_load_errors).new.tap do |c|
      c.catch_cell_load_errors = false
    end
  end

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

    ##
    # For internal use; stores source xml in nokogiri documents
    class Xml
      attr_accessor :workbook, :shared_strings, :sheets, :sheet_rels, :styles

      def self.load(file_path)
        self.new.tap do |xml|
          SimpleXlsxReader::Zip.open(file_path) do |zip|
            xml.sheets = []
            xml.sheet_rels = []

            # This weird style of enumerating over the entries lets us
            # concisely assign entries in a case insensitive and
            # slash insensitive ('/' vs '\') manner.
            #
            # RubyZip used to normalize the slashes, but doesn't now:
            # https://github.com/rubyzip/rubyzip/issues/324
            zip.entries.each do |entry|
              if entry.name.match(/^xl.workbook\.xml$/) # xl/workbook.xml
                xml.workbook = Nokogiri::XML(zip.read(entry)).remove_namespaces!
              elsif entry.name.match(/^xl.styles\.xml$/) # xl/styles.xml
                xml.styles   = Nokogiri::XML(zip.read(entry)).remove_namespaces!
              elsif entry.name.match(/^xl.sharedStrings\.xml$/i) # xl/sharedStrings.xml
                # optional feature used by excel, but not often used by xlsx
                # generation libraries. Path name is sometimes lowercase, too.
                xml.shared_strings = Nokogiri::XML(zip.read(entry)).remove_namespaces!
              elsif match = entry.name.match(/^xl.worksheets.sheet([0-9]*)\.xml$/)
                sheet_number = match.captures.first.to_i
                xml.sheets[sheet_number] =
                  Nokogiri::XML(zip.read(entry)).remove_namespaces!
              elsif match = entry.name.match(/^xl.worksheets._rels.sheet([0-9]*)\.xml\.rels$/)
                sheet_number = match.captures.first.to_i
                xml.sheet_rels[sheet_number] =
                  Nokogiri::XML(zip.read(entry)).remove_namespaces!
              end
            end

            # Sometimes there's a zero-index sheet.xml, ex.
            # Google Docs creates:
            #
            # xl/worksheets/sheet.xml
            # xl/worksheets/sheet1.xml
            # xl/worksheets/sheet2.xml
            # While Excel creates:
            # xl/worksheets/sheet1.xml
            # xl/worksheets/sheet2.xml
            #
            # So, for the latter case, let's shift [null, <Sheet 1>, <Sheet 2>]
            if !xml.sheets[0]
              xml.sheets.shift
              xml.sheet_rels.shift
            end
          end
        end
      end
    end

    ##
    # For internal use; translates source xml to Sheet objects.
    class Mapper < Struct.new(:xml)
      DATE_SYSTEM_1900 = Date.new(1899, 12, 30)
      DATE_SYSTEM_1904 = Date.new(1904, 1, 1)

      def load_sheets
        sheet_toc.each_with_index.map do |(sheet_name, _sheet_number), i|
          parse_sheet(sheet_name, xml.sheets[i], xml.sheet_rels[i])  # sheet_number is *not* the index into xml.sheets
        end
      end

      # Table of contents for the sheets, ex. {'Authors' => 0, ...}
      def sheet_toc
        xml.workbook.xpath('/workbook/sheets/sheet').
          inject({}) do |acc, sheet|

          acc[sheet.attributes['name'].value] =
            sheet.attributes['sheetId'].value.to_i - 1 # keep things 0-indexed

          acc
        end
      end

      def parse_sheet(sheet_name, xsheet, xrels)
        sheet = Sheet.new(sheet_name)
        sheet_width, sheet_height = *sheet_dimensions(xsheet)
        cells_w_links = xsheet.xpath('//hyperlinks/hyperlink').inject({}) {|acc, e| acc[e.attr(:ref)] = e.attr(:id); acc}

        sheet.rows = Array.new(sheet_height) { Array.new(sheet_width) }
        xsheet.xpath("/worksheet/sheetData/row/c").each do |xcell|
          column, row = *xcell.attr('r').match(/([A-Z]+)([0-9]+)/).captures
          col_idx = column_letter_to_number(column) - 1
          row_idx = row.to_i - 1

          type  = xcell.attributes['t'] &&
                  xcell.attributes['t'].value
          style = xcell.attributes['s'] &&
                  style_types[xcell.attributes['s'].value.to_i]

          # This is the main performance bottleneck. Using just 'xcell.text'
          # would be ideal, and makes parsing super-fast. However, there's
          # other junk in the cell, formula references in particular,
          # so we really do have to look for specific value nodes.
          # Maybe there is a really clever way to use xcell.text and parse out
          # the correct value, but I can't think of one, or an alternative
          # strategy.
          #
          # And yes, this really is faster than using xcell.at_xpath(...),
          # by about 60%. Odd.
          xvalue = type == 'inlineStr' ?
            (xis = xcell.children.find {|c| c.name == 'is'}) && xis.children.find {|c| c.name == 't'} :
            xcell.children.find {|c| c.name == 'f' && c.text.start_with?('HYPERLINK(') || c.name == 'v'}

          if xvalue
            value = xvalue.text.strip

            if rel_id = cells_w_links[xcell.attr('r')] # a hyperlink made via GUI
              url = xrels.at_xpath(%(//*[@Id="#{rel_id}"])).attr('Target')
            elsif xvalue.name == 'f' # only time we have a function is if it's a hyperlink
              url = value.slice(/HYPERLINK\("(.*?)"/, 1)
            end
          end

          cell = begin
            self.class.cast(value, type, style,
                            :url => url,
                            :shared_strings => shared_strings,
                            :base_date => base_date)
          rescue => e
            if !SimpleXlsxReader.configuration.catch_cell_load_errors
              error = CellLoadError.new(
                "Row #{row_idx}, Col #{col_idx}: #{e.message}")
              error.set_backtrace(e.backtrace)
              raise error
            else
              sheet.load_errors[[row_idx, col_idx]] = e.message

              xcell.text.strip
            end
          end

          # This shouldn't be necessary, but just in case, we'll create
          # the row so we don't blow up. This means any null rows in between
          # will be null instead of [null, null, ...]
          sheet.rows[row_idx] ||= Array.new(sheet_width)

          sheet.rows[row_idx][col_idx] = cell
        end

        sheet
      end

      ##
      # Returns the last column name, ex. 'E'
      #
      # Note that excel writes a '/worksheet/dimension' node we can get the
      # last cell from, but some libs (ex. simple_xlsx_writer) don't record
      # this. In that case, we assume the data is of uniform column length
      # and check the column name of the last header row. Obviously this isn't
      # the most robust strategy, but it likely fits 99% of use cases
      # considering it's not a problem with actual excel docs.
      def last_cell_label(xsheet)
        dimension = xsheet.at_xpath('/worksheet/dimension')
        if dimension
          col = dimension.attributes['ref'].value.match(/:([A-Z]+[0-9]+)/)
          col ? col.captures.first : 'A1'
        else
          last = xsheet.at_xpath("/worksheet/sheetData/row[last()]/c[last()]")
          last ? last.attributes['r'].value.match(/([A-Z]+[0-9]+)/).captures.first : 'A1'
        end
      end

      # Returns dimensions (1-indexed)
      def sheet_dimensions(xsheet)
        column, row = *last_cell_label(xsheet).match(/([A-Z]+)([0-9]+)/).captures
        [column_letter_to_number(column), row.to_i]
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

      # Excel doesn't record types for some cells, only its display style, so
      # we have to back out the type from that style.
      #
      # Some of these styles can be determined from a known set (see NumFmtMap),
      # while others are 'custom' and we have to make a best guess.
      #
      # This is the array of types corresponding to the styles a spreadsheet
      # uses, and includes both the known style types and the custom styles.
      #
      # Note that the xml sheet cells that use this don't reference the
      # numFmtId, but instead the array index of a style in the stored list of
      # only the styles used in the spreadsheet (which can be either known or
      # custom). Hence this style types array, rather than a map of numFmtId to
      # type.
      def style_types
        @style_types ||=
            xml.styles.xpath('/styleSheet/cellXfs/xf').map {|xstyle|
              style_type_by_num_fmt_id(num_fmt_id(xstyle))}
      end

      #returns the numFmtId value if it's available
      def num_fmt_id(xstyle)
        if xstyle.attributes['numFmtId']
          xstyle.attributes['numFmtId'].value
        else
          nil
        end
      end

      # Finds the type we think a style is; For example, fmtId 14 is a date
      # style, so this would return :date.
      #
      # Note, custom styles usually (are supposed to?) have a numFmtId >= 164,
      # but in practice can sometimes be simply out of the usual "Any Language"
      # id range that goes up to 49. For example, I have seen a numFmtId of
      # 59 specified as a date. In Thai, 59 is a number format, so this seems
      # like a bad idea, but we try to be flexible and just go with it.
      def style_type_by_num_fmt_id(id)
        return nil if id.nil?

        id = id.to_i
        NumFmtMap[id] || custom_style_types[id]
      end

      # Map of (numFmtId >= 164) (custom styles) to our best guess at the type
      # ex. {164 => :date_time}
      def custom_style_types
        @custom_style_types ||=
          xml.styles.xpath('/styleSheet/numFmts/numFmt').
          inject({}) do |acc, xstyle|

          acc[xstyle.attributes['numFmtId'].value.to_i] =
            determine_custom_style_type(xstyle.attributes['formatCode'].value)

          acc
        end
      end

      # This is the least deterministic part of reading xlsx files. Due to
      # custom styles, you can't know for sure when a date is a date other than
      # looking at its format and gessing. It's not impossible to guess right,
      # though.
      #
      # http://stackoverflow.com/questions/4948998/determining-if-an-xlsx-cell-is-date-formatted-for-excel-2007-spreadsheets
      def determine_custom_style_type(string)
        return :float if string[0] == '_'
        return :float if string[0] == ' 0'

        # Looks for one of ymdhis outside of meta-stuff like [Red]
        return :date_time if string =~ /(^|\])[^\[]*[ymdhis]/i

        return :unsupported
      end

      ##
      # The heart of typecasting. The ruby type is determined either explicitly
      # from the cell xml or implicitly from the cell style, and this
      # method expects that work to have been done already. This, then,
      # takes the type we determined it to be and casts the cell value
      # to that type.
      #
      # types:
      # - s: shared string (see #shared_string)
      # - n: number (cast to a float)
      # - b: boolean
      # - str: string
      # - inlineStr: string
      # - ruby symbol: for when type has been determined by style
      #
      # options:
      # - shared_strings: needed for 's' (shared string) type
      def self.cast(value, type, style, options = {})
        return nil if value.nil? || value.empty?

        # Sometimes the type is dictated by the style alone
        if type.nil? ||
          (type == 'n' && [:date, :time, :date_time].include?(style))
          type = style
        end

        casted = case type

        ##
        # There are few built-in types
        ##

        when 's' # shared string
          options[:shared_strings][value.to_i]
        when 'n' # number
          value.to_f
        when 'b'
          value.to_i == 1
        when 'str'
          value
        when 'inlineStr'
          value

        ##
        # Type can also be determined by a style,
        # detected earlier and cast here by its standardized symbol
        ##

        when :string, :unsupported
          value
        when :fixnum
          value.to_i
        when :float
          value.to_f
        when :percentage
          value.to_f / 100
        # the trickiest. note that  all these formats can vary on
        # whether they actually contain a date, time, or datetime.
        when :date, :time, :date_time
          value = Float(value)
          days_since_date_system_start = value.to_i
          fraction_of_24 = value - days_since_date_system_start

          # http://stackoverflow.com/questions/10559767/how-to-convert-ms-excel-date-from-float-to-date-format-in-ruby
          date = options.fetch(:base_date, DATE_SYSTEM_1900) + days_since_date_system_start

          if fraction_of_24 > 0 # there is a time associated
            seconds = (fraction_of_24 * 86400).round
            return Time.utc(date.year, date.month, date.day) + seconds
          else
            return date
          end
        when :bignum
          if defined?(BigDecimal)
            BigDecimal.new(value)
          else
            value.to_f
          end

        ##
        # Beats me
        ##

        else
          value
        end

        if options[:url]
          Hyperlink.new(options[:url], casted)
        else
          casted
        end
      end

      ## Returns the base_date from which to calculate dates.
      # Defaults to 1900 (minus two days due to excel quirk), but use 1904 if
      # it's set in the Workbook's workbookPr.
      # http://msdn.microsoft.com/en-us/library/ff530155(v=office.12).aspx
      def base_date
        @base_date ||=
          begin
            return DATE_SYSTEM_1900 if xml.workbook == nil
            xml.workbook.xpath("//workbook/workbookPr[@date1904]").each do |workbookPr|
              return DATE_SYSTEM_1904 if workbookPr["date1904"] =~ /true|1/i
            end
            DATE_SYSTEM_1900
          end
      end

      # Map of non-custom numFmtId to casting symbol
      NumFmtMap = {
        0  => :string,         # General
        1  => :fixnum,         # 0
        2  => :float,          # 0.00
        3  => :fixnum,         # #,##0
        4  => :float,          # #,##0.00
        5  => :unsupported,    # $#,##0_);($#,##0)
        6  => :unsupported,    # $#,##0_);[Red]($#,##0)
        7  => :unsupported,    # $#,##0.00_);($#,##0.00)
        8  => :unsupported,    # $#,##0.00_);[Red]($#,##0.00)
        9  => :percentage,     # 0%
        10 => :percentage,     # 0.00%
        11 => :bignum,         # 0.00E+00
        12 => :unsupported,    # # ?/?
        13 => :unsupported,    # # ??/??
        14 => :date,           # mm-dd-yy
        15 => :date,           # d-mmm-yy
        16 => :date,           # d-mmm
        17 => :date,           # mmm-yy
        18 => :time,           # h:mm AM/PM
        19 => :time,           # h:mm:ss AM/PM
        20 => :time,           # h:mm
        21 => :time,           # h:mm:ss
        22 => :date_time,      # m/d/yy h:mm
        37 => :unsupported,    # #,##0 ;(#,##0)
        38 => :unsupported,    # #,##0 ;[Red](#,##0)
        39 => :unsupported,    # #,##0.00;(#,##0.00)
        40 => :unsupported,    # #,##0.00;[Red](#,##0.00)
        45 => :time,           # mm:ss
        46 => :time,           # [h]:mm:ss
        47 => :time,           # mmss.0
        48 => :bignum,         # ##0.0E+0
        49 => :unsupported     # @
      }

      # For performance reasons, excel uses an optional SpreadsheetML feature
      # that puts all strings in a separate xml file, and then references
      # them by their index in that file.
      #
      # http://msdn.microsoft.com/en-us/library/office/gg278314.aspx
      def shared_strings
        @shared_strings ||= begin
          if xml.shared_strings
            xml.shared_strings.xpath('/sst/si').map do |xsst|
              # a shared string can be a single value...
              sst = xsst.at_xpath('t/text()')
              sst = sst.text if sst
              # ... or a composite of seperately styled words/characters
              sst ||= xsst.xpath('r/t/text()').map(&:text).join
            end
          else
            []
          end
        end
      end

    end

  end
end
