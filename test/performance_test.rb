require_relative 'test_helper'
require 'minitest/benchmark'

describe 'SimpleXlsxReader Benchmark' do

  # n is 0-indexed for us, then converted to 1-indexed for excel
  def build_row(n)
    n += 1
    <<-XML
      <row>
        <c r='A#{n}' s='0'>
          <v>Cell A#{n}</v>
        </c>
        <c r='B#{n}' s='1'>
          <v>2.4</v>
        </c>
        <c r='C#{n}' s='2'>
          <v>30687</v>
        </c>
        <c r='D#{n}' t='inlineStr' s='0'>
          <is><t>Cell D#{n}</t></is>
        </c>

        <c r='E#{n}' s='0'>
          <v>Cell E#{n}</v>
        </c>
        <c r='F#{n}' s='1'>
          <v>2.4</v>
        </c>
        <c r='G#{n}' s='2'>
          <v>30687</v>
        </c>
        <c r='H#{n}' t='inlineStr' s='0'>
          <is><t>Cell H#{n}</t></is>
        </c>

        <c r='I#{n}' s='0'>
          <v>Cell I#{n}</v>
        </c>
        <c r='J#{n}' s='1'>
          <v>2.4</v>
        </c>
        <c r='K#{n}' s='2'>
          <v>30687</v>
        </c>
        <c r='L#{n}' t='inlineStr' s='0'>
          <is><t>Cell L#{n}</t></is>
        </c>
      </row>
    XML
  end

  before do
    base = Nokogiri::XML(
      <<-XML
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
          </sheetData>
        </worksheet>
      XML
    ).remove_namespaces!
    base.at_xpath("/worksheet/sheetData").add_child(build_row(0))

    @xml = SimpleXlsxReader::Document::Xml.new.tap do |xml|
      xml.sheets = [base]

      # s='0' above refers to the value of numFmtId at cellXfs index 0,
      # which is in this case 'General' type
      xml.styles = Nokogiri::XML(
        <<-XML
          <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <cellXfs count="1">
              <xf numFmtId="0" />
              <xf numFmtId="2" />
              <xf numFmtId="14" />
            </cellXfs>
          </styleSheet>
        XML
      ).remove_namespaces!
    end

    # Every new sheet has one more row
    self.class.bench_range.each do |range|
      sheet = base.clone

      range.times do |n|
        sheet.xpath("/worksheet/sheetData/row").last.
          add_next_sibling(build_row(n+1))
      end

      @xml.sheets[range] = sheet
    end
  end

  def self.bench_range
    bench_exp(1,10000)
  end

  bench_performance_linear 'parses sheets in linear time', 0.9999 do |n|

    raise "not enough sample data; asked for #{n}, only have #{@xml.sheets.size}"\
      if @xml.sheets[n].nil?

    sheet = SimpleXlsxReader::Document::Mapper.new(@xml).
      parse_sheet('test', @xml.sheets[n])

    raise "sheet didn't parse correctly; expected #{n + 1} rows, got #{sheet.rows.size}"\
      if sheet.rows.size != n + 1
  end

end
