# frozen_string_literal: true

require_relative 'test_helper'
require 'time'

SXR = SimpleXlsxReader

describe SimpleXlsxReader do
  let(:sesame_street_blog_file) do
    File.join(File.dirname(__FILE__), 'sesame_street_blog.xlsx')
  end

  let(:document) { SimpleXlsxReader.open(sesame_street_blog_file) }

  ##
  # A high-level acceptance test testing misc features such as date/time parsing,
  # hyperlinks (both function and ref kinds), formula dates, emty rows, etc.

  let(:sesame_street_blog_file_path) { File.join(File.dirname(__FILE__), 'sesame_street_blog.xlsx') }
  let(:sesame_street_blog_io) { File.new(sesame_street_blog_file_path) }
  let(:sesame_street_blog_string) { IO.read(sesame_street_blog_file_path) }

  let(:expected_result) do
    {
      'Authors' =>
      [
        ['Name', 'Occupation'],
        ['Big Bird', 'Teacher']
      ],
      'Posts' =>
      [
        ['Author Name', 'Title', 'Body', 'Created At', 'Comment Count', 'URL'],
        ['Big Bird', 'The Number 1', 'The Greatest', Time.parse('2002-01-01 11:00:00 UTC'), 1, SXR::Hyperlink.new('http://www.example.com/hyperlink-function', 'This uses the HYPERLINK() function')],
        ['Big Bird', 'The Number 2', 'Second Best', Time.parse('2002-01-02 14:00:00 UTC'), 2, SXR::Hyperlink.new('http://www.example.com/hyperlink-gui', 'This uses the hyperlink GUI option')],
        ['Big Bird', 'Formula Dates', 'Tricky tricky', Time.parse('2002-01-03 14:00:00 UTC'), 0, nil],
        ['Empty Eagress', nil, 'The title, date, and comment have types, but no values', nil, nil, nil]
      ]
    }
  end

  describe SimpleXlsxReader do
    describe 'load from file path' do
      let(:subject) { SimpleXlsxReader.open(sesame_street_blog_file_path) }

      it 'reads an xlsx file into a hash of {[sheet name] => [data]}' do
        _(subject.to_hash).must_equal(expected_result)
      end
    end

    describe 'load from buffer' do
      let(:subject) { SimpleXlsxReader.parse(sesame_street_blog_io) }

      it 'reads an xlsx buffer into a hash of {[sheet name] => [data]}' do
        _(subject.to_hash).must_equal(expected_result)
      end
    end

    describe 'load from string' do
      let(:subject) { SimpleXlsxReader.parse(sesame_street_blog_io) }

      it 'reads an xlsx string into a hash of {[sheet name] => [data]}' do
        _(subject.to_hash).must_equal(expected_result)
      end
    end

    it 'outputs strings in UTF-8 encoding' do
      document = SimpleXlsxReader.parse(sesame_street_blog_io)
      _(document.sheets[0].rows.to_a.flatten.map(&:encoding).uniq)
        .must_equal [Encoding::UTF_8]
    end

    it 'can use all our enumerable nicities without slurping' do
      document = SimpleXlsxReader.parse(sesame_street_blog_io)

      headers = {
        name: 'Author Name',
        title: 'Title',
        body: 'Body',
        created_at: 'Created At',
        count: /Count/
      }

      rows = document.sheets[1].rows
      result =
        rows.each(headers: headers).with_index.with_object({}) do |(row, i), acc|
          acc[i] = row
        end

      _(result[0]).must_equal(
        name: 'Big Bird',
        title: 'The Number 1',
        body: 'The Greatest',
        created_at: Time.parse('2002-01-01 11:00:00 UTC'),
        count: 1,
        "URL" => 'This uses the HYPERLINK() function'
      )

      _(rows.slurped?).must_equal false
    end
  end

  ##
  # For more fine-grained unit tests, we sometimes build our own workbook via
  # Nokogiri. TestXlsxBuilder has some defaults, and this let-style lets us
  # concisely override them in nested describe blocks.

  let(:shared_strings) { nil }
  let(:styles) { nil }
  let(:sheet) { nil }
  let(:workbook) { nil }
  let(:rels) { nil }

  let(:xlsx) do
    TestXlsxBuilder.new(
      shared_strings: shared_strings,
      styles: styles,
      sheets: sheet && [sheet],
      workbook: workbook,
      rels: rels
    )
  end

  let(:reader) { SimpleXlsxReader.open(xlsx.archive.path) }

  describe 'when parsing escaped characters' do
    let(:escaped_content) do
      '&lt;a href="https://www.example.com"&gt;Link A&lt;/a&gt; &amp;bull; &lt;a href="https://www.example.com"&gt;Link B&lt;/a&gt;'
    end

    let(:unescaped_content) do
      '<a href="https://www.example.com">Link A</a> &bull; <a href="https://www.example.com">Link B</a>'
    end

    let(:sheet) do
      <<~XML
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <dimension ref="A1:B1" />
          <sheetData>
            <row r="1">
              <c r="A1" s="1" t="s">
                <v>0</v>
              </c>
              <c r='B1' s='0'>
                <v>#{escaped_content}</v>
              </c>
            </row>
          </sheetData>
        </worksheet>
      XML
    end

    let(:shared_strings) do
      <<~XML
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
          <si>
            <t>#{escaped_content}</t>
          </si>
        </sst>
      XML
    end

    it 'loads correctly using inline strings' do
      _(reader.sheets[0].rows.slurp[0][0]).must_equal(unescaped_content)
    end

    it 'loads correctly using shared strings' do
      _(reader.sheets[0].rows.slurp[0][1]).must_equal(unescaped_content)
    end
  end

  describe 'Sheet#rows#each(headers: true)' do
    let(:sheet) do
      <<~XML
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <dimension ref="A1:B3" />
          <sheetData>
            <row r="1">
              <c r="A1" s="0">
                <v>Header 1</v>
              </c>
              <c r="B1" s="0">
                <v>Header 2</v>
              </c>
            </row>
            <row r="2">
              <c r="A2" s="0">
                <v>Data 1-A</v>
              </c>
              <c r="B2" s="0">
                <v>Data 1-B</v>
              </c>
            </row>
            <row r="4">
              <c r="A4" s="0">
                <v>Data 2-A</v>
              </c>
              <c r="B4" s="0">
                <v>Data 2-B</v>
              </c>
            </row>
          </sheetData>
        </worksheet>
      XML
    end

    it 'yields rows as hashes' do
      acc = []

      reader.sheets[0].rows.each(headers: true) do |row|
        acc << row
      end

      _(acc).must_equal(
        [
          { 'Header 1' => 'Data 1-A', 'Header 2' => 'Data 1-B' },
          { 'Header 1' => nil, 'Header 2' => nil },
          { 'Header 1' => 'Data 2-A', 'Header 2' => 'Data 2-B' }
        ]
      )
    end
  end

  describe 'Sheet#rows#each(headers: ->(row) {...})' do
    let(:sheet) do
      <<~XML
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <dimension ref="A1:B7" />
          <sheetData>
            <row r="1">
              <c r="A1" s="0">
                <v>a chart or something</v>
              </c>
              <c r="B1" s="0">
                <v>Rabble rabble</v>
              </c>
            </row>
            <row r="2">
              <c r="A2" s="0">
                <v>Chatty junk</v>
              </c>
              <c r="B2" s="0">
                <v></v>
              </c>
            </row>
            <row r="4">
              <c r="A4" s="0">
                <v>Header 1</v>
              </c>
              <c r="B4" s="0">
                <v>Header 2</v>
              </c>
            </row>
            <row r="5">
              <c r="A5" s="0">
                <v>Data 1-A</v>
              </c>
              <c r="B5" s="0">
                <v>Data 1-B</v>
              </c>
            </row>
            <row r="7">
              <c r="A7" s="0">
                <v>Data 2-A</v>
              </c>
              <c r="B7" s="0">
                <v>Data 2-B</v>
              </c>
            </row>
          </sheetData>
        </worksheet>
      XML
    end

    it 'yields rows as hashes' do
      acc = []

      finder = ->(row) { row.find {|c| c&.match(/Header/)} }
      reader.sheets[0].rows.each(headers: finder) do |row|
        acc << row
      end

      _(acc).must_equal(
        [
          { 'Header 1' => 'Data 1-A', 'Header 2' => 'Data 1-B' },
          { 'Header 1' => nil, 'Header 2' => nil },
          { 'Header 1' => 'Data 2-A', 'Header 2' => 'Data 2-B' }
        ]
      )
    end
  end

  describe "Sheet#rows#each(headers: a_hash)" do
    let(:sheet) do
      Nokogiri::XML(
        <<~XML
          <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <dimension ref="A1:C7" />
            <sheetData>
              <row r="1">
                <c r="A1" s="0">
                  <v>a chart or something</v>
                </c>
                <c r="B1" s="0">
                  <v>Rabble rabble</v>
                </c>
                <c r="C1" s="0">
                  <v>Rabble rabble</v>
                </c>
              </row>
              <row r="2">
                <c r="A2" s="0">
                  <v>Chatty junk</v>
                </c>
                <c r="B2" s="0">
                  <v></v>
                </c>
                <c r="C2" s="0">
                  <v></v>
                </c>
              </row>
              <row r="4">
                <c r="A4" s="0">
                  <v>ID Number</v>
                </c>
                <c r="B4" s="0">
                  <v>ExacT</v>
                </c>
                <c r="C4" s="0">
                  <v>FOO Name</v>
                </c>

              </row>
              <row r="5">
                <c r="A5" s="0">
                  <v>ID 1-A</v>
                </c>
                <c r="B5" s="0">
                  <v>Exact 1-B</v>
                </c>
                <c r="C5" s="0">
                  <v>Name 1-C</v>
                </c>
              </row>
              <row r="7">
                <c r="A7" s="0">
                  <v>ID 2-A</v>
                </c>
                <c r="B7" s="0">
                  <v>Exact 2-B</v>
                </c>
                <c r="C7" s="0">
                  <v>Name 2-C</v>
                </c>
              </row>
            </sheetData>
          </worksheet>
        XML
      )
    end

    it 'transforms headers into symbols based on the header map' do
      header_map = {id: /ID/, name: /foo/i, exact: 'ExacT'}
      result = reader.sheets[0].rows.each(headers: header_map).to_a

      _(result).must_equal(
        [
          { id: 'ID 1-A', exact: 'Exact 1-B', name: 'Name 1-C' },
          { id: nil, exact: nil, name: nil },
          { id: 'ID 2-A', exact: 'Exact 2-B', name: 'Name 2-C' },
        ]
      )
    end

    it 'if a match isnt found, uses un-matched header name' do
      sheet.xpath("//*[text() = 'ExacT']")
        .first.children.first.content = 'not ExacT'

      header_map = {id: /ID/, name: /foo/i, exact: 'ExacT'}
      result = reader.sheets[0].rows.each(headers: header_map).to_a

      _(result).must_equal(
        [
          { id: 'ID 1-A', 'not ExacT' => 'Exact 1-B', name: 'Name 1-C' },
          { id: nil, 'not ExacT' => nil, name: nil },
          { id: 'ID 2-A', 'not ExacT' => 'Exact 2-B', name: 'Name 2-C' },
        ]
      )
    end
  end

  describe 'Sheet#rows[]' do
    it 'raises a RuntimeError if rows not slurped yet' do
      _(-> { reader.sheets[0].rows[1] }).must_raise(RuntimeError)
    end

    it 'works if the rows have been slurped' do
      _(reader.sheets[0].rows.tap(&:slurp)[0]).must_equal(
        ['Cell A', 'Cell B', 'Cell C']
      )
    end

    it 'works if the config allows auto slurping' do
      SimpleXlsxReader.configuration.auto_slurp = true

      _(reader.sheets[0].rows[0]).must_equal(
        ['Cell A', 'Cell B', 'Cell C']
      )

      SimpleXlsxReader.configuration.auto_slurp = false
    end
  end

  describe 'Sheet#rows#slurp' do
    let(:rows) { reader.sheets[0].rows.tap(&:slurp) }

    it 'loads the sheet parser results into memory' do
      _(rows.slurped).must_equal(
        [['Cell A', 'Cell B', 'Cell C']]
      )
    end

    it '#each and #map use slurped results' do
      _(rows.map(&:reverse)).must_equal(
        [['Cell C', 'Cell B', 'Cell A']]
      )
    end
  end

  describe 'Sheet#rows#each' do
    let(:sheet) do
      <<~XML
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <dimension ref="A1:B3" />
          <sheetData>
            <row r="1">
              <c r="A1" s="0">
                <v>Header 1</v>
              </c>
              <c r="B1" s="0">
                <v>Header 2</v>
              </c>
            </row>
            <row r="2">
              <c r="A2" s="0">
                <v>Data 1-A</v>
              </c>
              <c r="B2" s="0">
                <v>Data 1-B</v>
              </c>
            </row>
            <row r="4">
              <c r="A4" s="0">
                <v>Data 2-A</v>
              </c>
              <c r="B4" s="0">
                <v>Data 2-B</v>
              </c>
            </row>
          </sheetData>
        </worksheet>
      XML
    end

    let(:rows) { reader.sheets[0].rows }

    it 'with no block, returns an enumerator when not slurped' do
      _(rows.each.class).must_equal Enumerator
    end

    it 'with no block, passes on header argument in enumerator' do
      _(rows.each(headers: true).inspect).must_match 'headers: true'
    end

    it 'returns an enumerator when slurped' do
      rows.slurp
      _(rows.each.class).must_equal Enumerator
    end
  end

  describe 'Sheet#rows#map' do
    let(:sheet) do
      <<~XML
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <dimension ref="A1:B3" />
          <sheetData>
            <row r="1">
              <c r="A1" s="0">
                <v>Header 1</v>
              </c>
              <c r="B1" s="0">
                <v>Header 2</v>
              </c>
            </row>
            <row r="2">
              <c r="A2" s="0">
                <v>Data 1-A</v>
              </c>
              <c r="B2" s="0">
                <v>Data 1-B</v>
              </c>
            </row>
            <row r="4">
              <c r="A4" s="0">
                <v>Data 2-A</v>
              </c>
              <c r="B4" s="0">
                <v>Data 2-B</v>
              </c>
            </row>
          </sheetData>
        </worksheet>
      XML
    end

    let(:rows) { reader.sheets[0].rows }

    it 'does not slurp' do
      _(rows.map(&:first)).must_equal(
        ["Header 1", "Data 1-A", nil, "Data 2-A"]
      )
      _(rows.slurped?).must_equal false
    end
  end

  describe 'Sheet#headers' do
    let(:doc_sheet) { reader.sheets[0] }

    it 'raises a RuntimeError if rows not slurped yet' do
      _(-> { doc_sheet.headers }).must_raise(RuntimeError)
    end

    it 'returns first row if slurped' do
      _(doc_sheet.tap(&:slurp).headers).must_equal(
        ['Cell A', 'Cell B', 'Cell C']
      )
    end

    it 'returns first row if auto_slurp' do
      SimpleXlsxReader.configuration.auto_slurp = true

      _(doc_sheet.headers).must_equal(
        ['Cell A', 'Cell B', 'Cell C']
      )

      SimpleXlsxReader.configuration.auto_slurp = false
    end
  end

  describe SimpleXlsxReader::Loader do
    let(:described_class) { SimpleXlsxReader::Loader }

    describe '::cast' do
      it 'reads type s as a shared string' do
        _(described_class.cast('1', 's', nil, shared_strings: %w[a b c]))
          .must_equal 'b'
      end

      it 'reads type inlineStr as a string' do
        _(described_class.cast('the value', nil, 'inlineStr'))
          .must_equal 'the value'
      end

      it 'reads date styles' do
        _(described_class.cast('41505', nil, :date))
          .must_equal Date.parse('2013-08-19')
      end

      it 'reads time styles' do
        _(described_class.cast('41505.77083', nil, :time))
          .must_equal Time.parse('2013-08-19 18:30 UTC')
      end

      it 'reads date_time styles' do
        _(described_class.cast('41505.77083', nil, :date_time))
          .must_equal Time.parse('2013-08-19 18:30 UTC')
      end

      it 'reads number types styled as dates' do
        _(described_class.cast('41505', 'n', :date))
          .must_equal Date.parse('2013-08-19')
      end

      it 'reads number types styled as times' do
        _(described_class.cast('41505.77083', 'n', :time))
          .must_equal Time.parse('2013-08-19 18:30 UTC')
      end

      it 'reads less-than-zero complex number types styled as times' do
        _(described_class.cast('6.25E-2', 'n', :time))
          .must_equal Time.parse('1899-12-30 01:30:00 UTC')
      end

      it 'reads number types styled as date_times' do
        _(described_class.cast('41505.77083', 'n', :date_time))
          .must_equal Time.parse('2013-08-19 18:30 UTC')
      end

      it 'raises when date-styled values are not numerical' do
        _(-> { described_class.cast('14 is not a valid date', nil, :date) })
          .must_raise(ArgumentError)
      end

      describe 'with the url option' do
        let(:url) { 'http://www.example.com/hyperlink' }
        it 'creates a hyperlink with a string type' do
          _(described_class.cast('A link', 'str', :string, url: url))
            .must_equal SXR::Hyperlink.new(url, 'A link')
        end

        it 'creates a hyperlink with a shared string type' do
          _(described_class.cast('2', 's', nil, shared_strings: %w[a b c], url: url))
            .must_equal SXR::Hyperlink.new(url, 'c')
        end
      end
    end

    describe 'shared_strings' do
      let(:xml) do
        File.open(File.join(File.dirname(__FILE__), 'shared_strings.xml'))
      end

      let(:ss) { SimpleXlsxReader::Loader::SharedStringsParser.parse(xml) }

      it 'parses strings formatted at the cell level' do
        _(ss[0..2]).must_equal ['Cell A1', 'Cell B1', 'My Cell']
      end

      it 'parses strings formatted at the character level' do
        _(ss[3..5]).must_equal ['Cell A2', 'Cell B2', 'Cell Fmt']
      end

      it 'parses looong strings containing unicode' do
        _(ss[6]).must_include 'It only happens with both unicode *and* really long text.'
      end
    end

    describe 'style_types' do
      let(:xml_file) do
        File.open(File.join(File.dirname(__FILE__), 'styles.xml'))
      end

      let(:parser) do
        SimpleXlsxReader::Loader::StyleTypesParser.new(xml_file).tap(&:parse)
      end

      it 'reads custom formatted styles (numFmtId >= 164)' do
        _(parser.style_types[1]).must_equal :date_time
        _(parser.custom_style_types[164]).must_equal :date_time
      end

      # something I've seen in the wild; don't think it's correct, but let's be flexible.
      it 'reads custom formatted styles given an id < 164, but not explicitly defined in the SpreadsheetML spec' do
        _(parser.style_types[2]).must_equal :date_time
        _(parser.custom_style_types[59]).must_equal :date_time
      end
    end

    describe '#last_cell_label' do
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

      let(:loader) do
        SimpleXlsxReader::Loader.new(nil).tap do |l|
          l.shared_strings = []
          l.sheet_toc = { 'Sheet1': 0 }
          l.style_types = []
          l.base_date = SimpleXlsxReader::DATE_SYSTEM_1900
        end
      end

      let(:sheet_parser) do
        tempfile = Tempfile.new(['sheet', '.xml'])
        tempfile.write(sheet)
        tempfile.rewind

        SimpleXlsxReader::Loader::SheetParser.new(
          file_io: tempfile,
          loader: loader
        ).tap { |parser| parser.parse {} }
      end

      it 'uses /worksheet/dimension if available' do
        _(sheet_parser.last_cell_letter).must_equal 'C'
      end

      it 'uses the last header cell if /worksheet/dimension is missing' do
        sheet.at_xpath('/worksheet/dimension').remove
        _(sheet_parser.last_cell_letter).must_equal 'D'
      end

      it 'returns "A1" if the dimension is just one cell' do
        sheet.xpath('/worksheet/sheetData/row').remove
        sheet.xpath('/worksheet/dimension').attr('ref', 'A1')
        _(sheet_parser.last_cell_letter).must_equal 'A'
      end

      it 'returns nil if the sheet is just one cell, but /worksheet/dimension is missing' do
        sheet.xpath('/worksheet/sheetData/row').remove
        sheet.xpath('/worksheet/dimension').remove
        _(sheet_parser.last_cell_letter).must_be_nil
      end
    end

    describe '#column_letter_to_number' do
      let(:subject) { SXR::Loader::SheetParser.new(file_io: nil, loader: nil) }

      [
        ['A', 1],
        ['B',   2],
        ['Z',   26],
        ['AA',  27],
        ['AB',  28],
        ['AZ',  52],
        ['BA',  53],
        ['BZ',  78],
        ['ZZ',  702],
        ['AAA', 703],
        ['AAZ', 728],
        ['ABA', 729],
        ['ABZ', 754],
        ['AZZ', 1378],
        ['ZZZ', 18_278]
      ].each do |(letter, number)|
        it "converts #{letter} to #{number}" do
          _(subject.column_letter_to_number(letter)).must_equal number
        end
      end
    end
  end

  describe 'parse errors' do
    after do
      SimpleXlsxReader.configuration.catch_cell_load_errors = false
    end

    let(:sheet) do
      Nokogiri::XML(
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
      ).remove_namespaces!
    end

    let(:styles) do
      # s='0' above refers to the value of numFmtId at cellXfs index 0
      Nokogiri::XML(
        <<-XML
          <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <cellXfs count="1">
              <xf numFmtId="14" />
            </cellXfs>
          </styleSheet>
        XML
      ).remove_namespaces!
    end

    it 'raises if configuration.catch_cell_load_errors' do
      SimpleXlsxReader.configuration.catch_cell_load_errors = false

      _(-> { SimpleXlsxReader.open(xlsx.archive.path).to_hash })
        .must_raise(SimpleXlsxReader::CellLoadError)
    end

    it 'records a load error if not configuration.catch_cell_load_errors' do
      SimpleXlsxReader.configuration.catch_cell_load_errors = true

      sheet = SimpleXlsxReader.open(xlsx.archive.path).sheets[0].tap(&:slurp)
      _(sheet.load_errors).must_equal(
        [0, 0] => 'invalid value for Float(): "14 is a date style; this is not a date"'
      )
    end
  end

  describe 'missing numFmtId attributes' do
    let(:sheet) do
      Nokogiri::XML(
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
      ).remove_namespaces!
    end

    let(:styles) do
      Nokogiri::XML(
        <<-XML
          <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">

          </styleSheet>
        XML
      ).remove_namespaces!
    end

    before do
      @row = SimpleXlsxReader.open(xlsx.archive.path).sheets[0].rows.to_a[0]
    end

    it 'continues even when cells are missing numFmtId attributes ' do
      _(@row[0]).must_equal 'some content'
    end
  end

  describe 'parsing types' do
    let(:sheet) do
      Nokogiri::XML(
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

                <c r='H1' s='0'>
                  <f>HYPERLINK("http://www.example.com/hyperlink-function", "HYPERLINK function")</f>
                  <v>HYPERLINK function</v>
                </c>

                <c r='I1' s='0'>
                  <v>GUI-made hyperlink</v>
                </c>

                <c r='J1' s='0'>
                  <v>1</v>
                </c>
              </row>
            </sheetData>

            <hyperlinks>
              <hyperlink ref="I1" id="rId1"/>
            </hyperlinks>
          </worksheet>
        XML
      ).remove_namespaces!
    end

    let(:styles) do
      # s='0' above refers to the value of numFmtId at cellXfs index 0,
      # which is in this case 'General' type
      Nokogiri::XML(
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

    # Although not a "type" or "style" according to xlsx spec,
    # it sure could/should be, so let's test it with the rest of our
    # typecasting code.
    let(:rels) do
      [
        Nokogiri::XML(
          <<-XML
            <Relationships>
              <Relationship
                Id="rId1"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                Target="http://www.example.com/hyperlink-gui"
                TargetMode="External"
              />
            </Relationships>
          XML
        ).remove_namespaces!
      ]
    end

    before do
      @row = SimpleXlsxReader.open(xlsx.archive.path).sheets[0].rows.to_a[0]
    end

    it "reads 'Generic' cells as strings" do
      _(@row[0]).must_equal 'Cell A1'
    end

    it "reads empty 'Generic' cells as nil" do
      _(@row[1]).must_be_nil
    end

    # We could expand on these type tests, but really just a couple
    # demonstrate that it's wired together. Type-specific tests should go
    # on #cast

    it 'reads floats' do
      _(@row[2]).must_equal 2.4
    end

    it 'reads empty floats as nil' do
      _(@row[3]).must_be_nil
    end

    it 'reads dates' do
      _(@row[4]).must_equal Date.parse('Jan 6, 1984')
    end

    it 'reads empty date cells as nil' do
      _(@row[5]).must_be_nil
    end

    it 'reads strings formatted as inlineStr' do
      _(@row[6]).must_equal 'Cell G1'
    end

    it 'reads hyperlinks created via HYPERLINK()' do
      _(@row[7]).must_equal(
        SXR::Hyperlink.new(
          'http://www.example.com/hyperlink-function', 'HYPERLINK function'
        )
      )
    end

    it 'reads hyperlinks created via the GUI' do
      _(@row[8]).must_equal(
        SXR::Hyperlink.new(
          'http://www.example.com/hyperlink-gui', 'GUI-made hyperlink'
        )
      )
    end

    it "reads 'Generic' cells with numbers as numbers" do
      _(@row[9]).must_equal 1
    end
  end

  describe 'parsing documents with blank rows' do
    let(:sheet) do
      Nokogiri::XML(
        <<-XML
          <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <dimension ref="A1:D7" />
            <sheetData>
            <row r="2" spans="1:1">
              <c r="A2" s="0">
                <v>a</v>
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
      ).remove_namespaces!
    end

    before do
      @rows = SimpleXlsxReader.open(xlsx.archive.path).sheets[0].rows.to_a
    end

    it 'reads row data despite gaps in row numbering' do
      _(@rows).must_equal [
        [nil, nil, nil, nil],
        ['a', nil, nil, nil],
        [nil, nil, nil, nil],
        [nil, 1, nil, nil],
        [nil, nil, 2, nil],
        [nil, nil, nil, nil],
        [nil, nil, nil, 3]
      ]
    end
  end

  # https://support.microsoft.com/en-us/office/available-number-formats-in-excel-0afe8f52-97db-41f1-b972-4b46e9f1e8d2
  describe 'numeric fields styled as "General"' do
    let(:misc_numbers_path) do
      File.join(File.dirname(__FILE__), 'misc_numbers.xlsx')
    end

    let(:sheet) { SimpleXlsxReader.open(misc_numbers_path).sheets[0] }

    it 'reads medium sized integers as integers' do
      _(sheet.rows.slurp[1][0]).must_equal 98070
    end

    it 'reads large (>12 char) integers as integers' do
      _(sheet.rows.slurp[1][1]).must_equal 1234567890123
    end
  end

  describe 'with mysteriously chunky UTF-8 text' do
    let(:chunky_utf8_path) do
      File.join(File.dirname(__FILE__), 'chunky_utf8.xlsx')
    end

    let(:sheet) { SimpleXlsxReader.open(chunky_utf8_path).sheets[0] }

    it 'reads the whole cell text' do
      _(sheet.rows.slurp[1]).must_equal(
        ["sample-company-1", "Korntal-Münchingen", "Bronholmer straße"]
      )
    end
  end

  describe 'when using percentages & currencies' do
    let(:pnc_path) do
      # This file provided by a GitHub user having parse errors in these fields
      File.join(File.dirname(__FILE__), 'percentages_n_currencies.xlsx')
    end

    let(:sheet) { SimpleXlsxReader.open(pnc_path).sheets[0] }

    it 'reads percentages as floats of the form 0.XX' do
      _(sheet.rows.slurp[1][2]).must_equal(0.87)
    end

    it 'reads currencies as floats' do
      _(sheet.rows.slurp[1][4]).must_equal(300.0)
    end
  end
end
