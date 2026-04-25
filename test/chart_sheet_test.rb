# frozen_string_literal: true

require_relative 'test_helper'

describe 'chart sheets' do
  # Chart-only sheets live in xl/chartsheets/, not xl/worksheets/, so they
  # have no SheetParser. The workbook.xml sheet list includes them though,
  # which causes init_sheets to mis-map every sheet that follows a chart sheet
  # to the wrong parser (off by one per chart sheet encountered).
  let(:xlsx) do
    workbook = <<~XML
      <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets>
          <sheet name="DataSheet1" sheetId="1" r:id="rId1"/>
          <sheet name="Chart1"     sheetId="2" r:id="rId2"/>
          <sheet name="DataSheet2" sheetId="3" r:id="rId3"/>
        </sheets>
      </workbook>
    XML

    sheet1 = <<~XML
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row><c r="A1" s="0"><v>from sheet1</v></c></row>
        </sheetData>
      </worksheet>
    XML

    sheet2 = <<~XML
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row><c r="A1" s="0"><v>from sheet2</v></c></row>
        </sheetData>
      </worksheet>
    XML

    TestXlsxBuilder.new(workbook: workbook, sheets: [sheet1, sheet2]).archive
  end

  it 'does not shift sheet index when a chart sheet is present' do
    doc = SimpleXlsxReader.open(xlsx.path)

    sheet1 = doc.sheets.detect {|s| s.name == 'DataSheet1'}
    sheet2 = doc.sheets.detect {|s| s.name == 'DataSheet2'}

    _(sheet1.rows.to_a).must_equal [['from sheet1']]
    _(sheet2.rows.to_a).must_equal [['from sheet2']]
  end
end
