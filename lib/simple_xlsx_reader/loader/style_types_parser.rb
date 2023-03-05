# frozen_string_literal: true

module SimpleXlsxReader
  class Loader
    StyleTypesParser = Struct.new(:file_io) do
      def self.parse(file_io)
        new(file_io).tap(&:parse).style_types
      end

      # Map of non-custom numFmtId to casting symbol
      NumFmtMap = {
        0 => :string,        # General
        1 => :fixnum,        # 0
        2 => :float,         # 0.00
        3 => :fixnum,        # #,##0
        4 => :float,         # #,##0.00
        5 => :unsupported,   # $#,##0_);($#,##0)
        6 => :unsupported,   # $#,##0_);[Red]($#,##0)
        7 => :unsupported,   # $#,##0.00_);($#,##0.00)
        8 => :unsupported,   # $#,##0.00_);[Red]($#,##0.00)
        9 => :percentage,    # 0%
        10 => :percentage,   # 0.00%
        11 => :bignum,       # 0.00E+00
        12 => :unsupported,  # # ?/?
        13 => :unsupported,  # # ??/??
        14 => :date,         # mm-dd-yy
        15 => :date,         # d-mmm-yy
        16 => :date,         # d-mmm
        17 => :date,         # mmm-yy
        18 => :time,         # h:mm AM/PM
        19 => :time,         # h:mm:ss AM/PM
        20 => :time,         # h:mm
        21 => :time,         # h:mm:ss
        22 => :date_time,    # m/d/yy h:mm
        37 => :unsupported,  # #,##0 ;(#,##0)
        38 => :unsupported,  # #,##0 ;[Red](#,##0)
        39 => :unsupported,  # #,##0.00;(#,##0.00)
        40 => :unsupported,  # #,##0.00;[Red](#,##0.00)
        44 => :float,        # some odd currency format ?from Office 2007?
        45 => :time,         # mm:ss
        46 => :time,         # [h]:mm:ss
        47 => :time,         # mmss.0
        48 => :bignum,       # ##0.0E+0
        49 => :unsupported   # @
      }.freeze

      def parse
        @xml = Nokogiri::XML(file_io.read).remove_namespaces!
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
        @xml.xpath('/styleSheet/cellXfs/xf').map do |xstyle|
          style_type_by_num_fmt_id(
            xstyle.attributes['numFmtId']&.value
          )
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
          @xml.xpath('/styleSheet/numFmts/numFmt')
            .each_with_object({}) do |xstyle, acc|
              acc[xstyle.attributes['numFmtId'].value.to_i] =
                determine_custom_style_type(xstyle.attributes['formatCode'].value)
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

        :unsupported
      end
    end
  end
end
