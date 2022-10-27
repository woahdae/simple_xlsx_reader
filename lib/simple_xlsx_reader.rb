# frozen_string_literal: true

require 'nokogiri'
require 'date'

require 'simple_xlsx_reader/version'
require 'simple_xlsx_reader/hyperlink'
require 'simple_xlsx_reader/document'
require 'simple_xlsx_reader/loader'
require 'simple_xlsx_reader/loader/workbook_parser'
require 'simple_xlsx_reader/loader/shared_strings_parser'
require 'simple_xlsx_reader/loader/sheet_parser'
require 'simple_xlsx_reader/loader/style_types_parser'


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
  DATE_SYSTEM_1900 = Date.new(1899, 12, 30)
  DATE_SYSTEM_1904 = Date.new(1904, 1, 1)

  class CellLoadError < StandardError; end

  class << self
    def configuration
      @configuration ||= Struct.new(:catch_cell_load_errors, :auto_slurp).new.tap do |c|
        c.catch_cell_load_errors = false
        c.auto_slurp = false
      end
    end

    def open(file_path)
      Document.new(file_path: file_path).tap(&:sheets)
    end
    
    def parse(string_or_io)
      Document.new(string_or_io: string_or_io).tap(&:sheets)
    end
  end
end
