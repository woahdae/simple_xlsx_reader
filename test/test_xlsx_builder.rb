# frozen_string_literal: true

require 'nokogiri'

TestXlsxBuilder = Struct.new(:shared_strings, :styles, :sheets, :workbook, :rels, keyword_init: true) do

  DEFAULTS = {
    workbook:
      Nokogiri::XML(
        <<-XML
          <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheets>
              <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
            </sheets>
          </styleSheet>
        XML
      ).remove_namespaces!,

    styles:
      Nokogiri::XML(
        <<-XML
          <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <cellXfs count="1">
              <xf numFmtId="0" />
            </cellXfs>
          </styleSheet>
        XML
      ).remove_namespaces!,

    sheet:
      Nokogiri::XML(
        <<-XML
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <dimension ref="A1:C1" />
          <sheetData>
            <row>
              <c r='A1' s='0'>
                <v>Cell A</v>
              </c>
              <c r='B1' s='0'>
                <v>Cell B</v>
              </c>
              <c r='C1' s='0'>
                <v>Cell C</v>
              </c>
            </row>
          </sheetData>
        </worksheet>
        XML
      ).remove_namespaces!
  }

  def initialize(*args)
    super

    self.workbook ||= DEFAULTS[:workbook]
    self.styles ||= DEFAULTS[:styles]
    self.sheets ||= [DEFAULTS[:sheet]]
    self.rels ||= []
  end

  def archive
    tmpfile = Tempfile.new(['workbook', '.xlsx'])
    tmpfile.binmode
    tmpfile.rewind

    Zip::File.open(tmpfile.path, create: true) do |zip|
      zip.mkdir('xl')

      zip.get_output_stream('xl/workbook.xml') do |wb_file|
        wb_file.write(workbook)
      end

      zip.get_output_stream('xl/styles.xml') do |styles_file|
        styles_file.write(styles)
      end

      if shared_strings
        zip.get_output_stream('xl/sharedStrings.xml') do |ss_file|
          ss_file.write(shared_strings)
        end
      end

      zip.mkdir('xl/worksheets')

      sheets.each_with_index do |sheet, i|
        zip.get_output_stream("xl/worksheets/sheet#{i + 1}.xml") do |sf|
          sf.write(sheet)
        end

        if rels[i]
          zip.mkdir('xl/worksheets/_rels')
          zip.get_output_stream("xl/worksheets/_rels/sheet#{i + 1}.xml.rels") do |rf|
            rf.write(rels[i])
          end
        end
      end
    end

    tmpfile
  end
end

