# frozen_string_literal: true

module SimpleXlsxReader
  class Loader
    WorkbookParser = Struct.new(:file_io) do
      def self.parse(file_io)
        parser = new(file_io).tap(&:parse)
        [parser.sheet_toc, parser.base_date]
      end

      def parse
        @xml = Nokogiri::XML(file_io.read).remove_namespaces!
      end

      # Table of contents for the sheets, ex. {'Authors' => 0, ...}
      def sheet_toc
        @xml.xpath('/workbook/sheets/sheet')
          .each_with_object({}) do |sheet, acc|
            acc[sheet.attributes['name'].value] =
              sheet.attributes['sheetId'].value.to_i - 1 # keep things 0-indexed
          end
      end

      ## Returns the base_date from which to calculate dates.
      # Defaults to 1900 (minus two days due to excel quirk), but use 1904 if
      # it's set in the Workbook's workbookPr.
      # http://msdn.microsoft.com/en-us/library/ff530155(v=office.12).aspx
      def base_date
        return DATE_SYSTEM_1900 if @xml.nil?

        @xml.xpath('//workbook/workbookPr[@date1904]').each do |workbookPr|
          return DATE_SYSTEM_1904 if workbookPr['date1904'] =~ /true|1/i
        end

        DATE_SYSTEM_1900
      end
    end
  end
end
