require_relative 'test_helper'
require 'time'

describe SimpleXlsxReader do
  let(:sesame_street_blog_file) { File.join(File.dirname(__FILE__),
                                            'sesame_street_blog.xlsx') }

  let(:subject) { SimpleXlsxReader::Document.new(sesame_street_blog_file) }

  describe '#to_hash' do
    it 'reads an xlsx file into a hash of {[sheet name] => [data]}' do
      subject.to_hash.must_equal({
        "Authors"=>
          [["Name", "Occupation"],
           ["Big Bird", "Teacher"]],

        "Posts"=>
          [["Author Name", "Title", "Body", "Created At", "Comment Count"],
           ["Big Bird", "The Number 1", "The Greatest", Time.parse("2002-01-01 11:00:00 UTC"), 1],
           ["Big Bird", "The Number 2", "Second Best", Time.parse("2002-01-02 14:00:00 UTC"), 2],
           ["Big Bird", "Formula Dates", "Tricky tricky", Time.parse("2002-01-03 14:00:00 UTC"), 0],
           ["Empty Eagress", nil, "The title, date, and comment have types, but no values", nil, nil]]
      })
    end
  end

  describe SimpleXlsxReader::Document::Mapper do
    let(:described_class) { SimpleXlsxReader::Document::Mapper }

    describe '::cast' do
      it 'reads type s as a shared string' do
        described_class.cast('1', 's', nil, :shared_strings => ['a', 'b', 'c']).
          must_equal 'b'
      end

      it 'reads type inlineStr as a string' do
        described_class.cast('the value', nil, 'inlineStr').
          must_equal 'the value'
      end

      it 'reads date styles' do
        described_class.cast('41505', nil, :date).
          must_equal Date.parse('2013-08-19')
      end

      it 'reads time styles' do
        described_class.cast('41505.77083', nil, :time).
          must_equal Time.parse('2013-08-19 18:30 UTC')
      end

      it 'reads date_time styles' do
        described_class.cast('41505.77083', nil, :date_time).
          must_equal Time.parse('2013-08-19 18:30 UTC')
      end

      it 'reads number types styled as dates' do
        described_class.cast('41505', 'n', :date).
          must_equal Date.parse('2013-08-19')
      end

      it 'reads number types styled as times' do
        described_class.cast('41505.77083', 'n', :time).
          must_equal Time.parse('2013-08-19 18:30 UTC')
      end

      it 'reads number types styled as date_times' do
        described_class.cast('41505.77083', 'n', :date_time).
          must_equal Time.parse('2013-08-19 18:30 UTC')
      end
    end

    describe '#shared_strings' do
      let(:xml) do
        SimpleXlsxReader::Document::Xml.new.tap do |xml|
          xml.shared_strings = Nokogiri::XML(File.read(
            File.join(File.dirname(__FILE__), 'shared_strings.xml') )).remove_namespaces!
        end
      end

      subject { described_class.new(xml) }

      it 'parses strings formatted at the cell level' do
        subject.shared_strings[0..2].must_equal ['Cell A1', 'Cell B1', 'My Cell']
      end

      it 'parses strings formatted at the character level' do
        subject.shared_strings[3..5].must_equal ['Cell A2', 'Cell B2', 'Cell Fmt']
      end
    end

    describe '#style_types' do
      let(:xml) do
        SimpleXlsxReader::Document::Xml.new.tap do |xml|
          xml.styles = Nokogiri::XML(File.read(
            File.join(File.dirname(__FILE__), 'styles.xml') )).remove_namespaces!
        end
      end

      let(:mapper) do
        SimpleXlsxReader::Document::Mapper.new(xml)
      end

      it 'reads custom formatted styles (numFmtId >= 164)' do
        mapper.style_types[1].must_equal :date_time
        mapper.custom_style_types[164].must_equal :date_time
      end

      # something I've seen in the wild; don't think it's correct, but let's be flexible.
      it 'reads custom formatted styles given an id < 164, but not explicitly defined in the SpreadsheetML spec' do
        mapper.style_types[2].must_equal :date_time
        mapper.custom_style_types[59].must_equal :date_time
      end
    end

    describe '#last_cell_label' do

      let(:generic_style) do
          Nokogiri::XML(
            <<-XML
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <cellXfs count="1">
                <xf numFmtId="0" />
              </cellXfs>
            </styleSheet>
            XML
          ).remove_namespaces!
      end

      # Note, this is not a valid sheet, since the last cell is actually D1 but
      # the dimension specifies C1. This is just for testing.
      let(:sheet) do
        Nokogiri::XML(
          <<-XML
          <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <dimension ref="A1:C1" />
            <sheetData>
              <row>
                <c r='A1' s='0'>
                  <v>Cell A</v>
                </c>
                <c r='C1' s='0'>
                  <v>Cell C</v>
                </c>
                <c r='D1' s='0'>
                  <v>Cell D</v>
                </c>
              </row>
            </sheetData>
          </worksheet>
          XML
        ).remove_namespaces!
      end

      let(:empty_sheet) do
        Nokogiri::XML(
          <<-XML
          <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <dimension ref="A1" />
            <sheetData>
            </sheetData>
          </worksheet>
          XML
        ).remove_namespaces!
      end

      let(:xml) do
        SimpleXlsxReader::Document::Xml.new.tap do |xml|
          xml.sheets = [sheet]
          xml.styles = generic_style
        end
      end

      subject { described_class.new(xml) }

      it 'uses /worksheet/dimension if available' do
        subject.last_cell_label(sheet).must_equal 'C1'
      end

      it 'uses the last header cell if /worksheet/dimension is missing' do
        sheet.xpath('/worksheet/dimension').remove
        subject.last_cell_label(sheet).must_equal 'D1'
      end

      it 'returns "A1" if the dimension is just one cell' do
        subject.last_cell_label(empty_sheet).must_equal 'A1'
      end

      it 'returns "A1" if the sheet is just one cell, but /worksheet/dimension is missing' do
        sheet.at_xpath('/worksheet/dimension').remove
        subject.last_cell_label(empty_sheet).must_equal 'A1'
      end
    end

    describe '#column_letter_to_number' do
      let(:subject) { described_class.new }

      [ ['A',   1    ],
        ['B',   2    ],
        ['Z',   26   ],
        ['AA',  27   ],
        ['AB',  28   ],
        ['AZ',  52   ],
        ['BA',  53   ],
        ['BZ',  78   ],
        ['ZZ',  702  ],
        ['AAA', 703  ],
        ['AAZ', 728  ],
        ['ABA', 729  ],
        ['ABZ', 754  ],
        ['AZZ', 1378 ],
        ['ZZZ', 18278] ].each do |(letter, number)|
        it "converts #{letter} to #{number}" do
          subject.column_letter_to_number(letter).must_equal number
        end
      end
    end

    describe "parse errors" do
      after do
        SimpleXlsxReader.configuration.catch_cell_load_errors = false
      end

      let(:xml) do
        SimpleXlsxReader::Document::Xml.new.tap do |xml|
          xml.sheets = [Nokogiri::XML(
            <<-XML
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <dimension ref="A1:A1" />
              <sheetData>
                <row>
                  <c r='A1' s='0'>
                    <v>14 is a date style; this is not a date</v>
                  </c>
                </row>
              </sheetData>
            </worksheet>
            XML
          ).remove_namespaces!]

          # s='0' above refers to the value of numFmtId at cellXfs index 0
          xml.styles = Nokogiri::XML(
            <<-XML
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <cellXfs count="1">
                <xf numFmtId="14" />
              </cellXfs>
            </styleSheet>
            XML
          ).remove_namespaces!
        end
      end

      it 'raises if configuration.catch_cell_load_errors' do
        SimpleXlsxReader.configuration.catch_cell_load_errors = false

        lambda { described_class.new(xml).parse_sheet('test', xml.sheets.first) }.
          must_raise(SimpleXlsxReader::CellLoadError)
      end

      it 'records a load error if not configuration.catch_cell_load_errors' do
        SimpleXlsxReader.configuration.catch_cell_load_errors = true

        sheet = described_class.new(xml).parse_sheet('test', xml.sheets.first)
        sheet.load_errors[[0,0]].must_include 'invalid value for Integer'
      end
    end

    describe "missing numFmtId attributes" do

      let(:xml) do
        SimpleXlsxReader::Document::Xml.new.tap do |xml|
          xml.sheets = [Nokogiri::XML(
                            <<-XML
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <dimension ref="A1:A1" />
              <sheetData>
                <row>
                  <c r='A1' s='s'>
                    <v>some content</v>
                  </c>
                </row>
              </sheetData>
            </worksheet>
                        XML
                        ).remove_namespaces!]

          xml.styles = Nokogiri::XML(
              <<-XML
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">

            </styleSheet>
          XML
          ).remove_namespaces!
        end
      end

      before do
        @row = described_class.new(xml).parse_sheet('test', xml.sheets.first).rows[0]
      end

      it 'continues even when cells are missing numFmtId attributes ' do
        @row[0].must_equal 'some content'
      end

    end

    describe 'parsing types' do
      let(:xml) do
        SimpleXlsxReader::Document::Xml.new.tap do |xml|
          xml.sheets = [Nokogiri::XML(
            <<-XML
              <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <dimension ref="A1:G1" />
                <sheetData>
                  <row>
                    <c r='A1' s='0'>
                      <v>Cell A1</v>
                    </c>

                    <c r='C1' s='1'>
                      <v>2.4</v>
                    </c>
                    <c r='D1' s='1' />

                    <c r='E1' s='2'>
                      <v>30687</v>
                    </c>
                    <c r='F1' s='2' />

                    <c r='G1' t='inlineStr' s='0'>
                      <is><t>Cell G1</t></is>
                    </c>
                  </row>
                </sheetData>
              </worksheet>
            XML
          ).remove_namespaces!]

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
      end

      before do
        @row = described_class.new(xml).parse_sheet('test', xml.sheets.first).rows[0]
      end

      it "reads 'Generic' cells as strings" do
        @row[0].must_equal "Cell A1"
      end

      it "reads empty 'Generic' cells as nil" do
        @row[1].must_equal nil
      end

      # We could expand on these type tests, but really just a couple
      # demonstrate that it's wired together. Type-specific tests should go
      # on #cast

      it "reads floats" do
        @row[2].must_equal 2.4
      end

      it "reads empty floats as nil" do
        @row[3].must_equal nil
      end

      it "reads dates" do
        @row[4].must_equal Date.parse('Jan 6, 1984')
      end

      it "reads empty date cells as nil" do
        @row[5].must_equal nil
      end

      it "reads strings formatted as inlineStr" do
        @row[6].must_equal 'Cell G1'
      end
    end

    describe 'parsing documents with blank rows' do
      let(:xml) do
        SimpleXlsxReader::Document::Xml.new.tap do |xml|
          xml.sheets = [Nokogiri::XML(
            <<-XML
              <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <dimension ref="A1:D7" />
                <sheetData>
                <row r="2" spans="1:1">
                  <c r="A2" s="0">
                    <v>0</v>
                  </c>
                </row>
                <row r="4" spans="1:1">
                  <c r="B4" s="0">
                    <v>1</v>
                  </c>
                </row>
                <row r="5" spans="1:1">
                  <c r="C5" s="0">
                    <v>2</v>
                  </c>
                </row>
                <row r="7" spans="1:1">
                  <c r="D7" s="0">
                    <v>3</v>
                  </c>
                </row>
                </sheetData>
              </worksheet>
            XML
          ).remove_namespaces!]

          xml.styles = Nokogiri::XML(
            <<-XML
              <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
                <cellXfs count="1">
                  <xf numFmtId="0" />
                </cellXfs>
              </styleSheet>
            XML
          ).remove_namespaces!
        end
      end

      before do
        @rows = described_class.new(xml).parse_sheet('test', xml.sheets.first).rows
      end

      it "reads row data despite gaps in row numbering" do
        @rows.must_equal [
          [nil,nil,nil,nil],
          ["0",nil,nil,nil],
          [nil,nil,nil,nil],
          [nil,"1",nil,nil],
          [nil,nil,"2",nil],
          [nil,nil,nil,nil],
          [nil,nil,nil,"3"]
        ]
      end
    end

  end
end
