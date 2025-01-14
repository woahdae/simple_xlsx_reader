# frozen_string_literal: true

require_relative 'test_helper'

describe SimpleXlsxReader do
  let(:sheet) do
    <<~XML
      <?xml version="1.0" encoding="utf-8"?>
      <x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <x:sheetData>
          <x:row>
            <x:c s="2" t="inlineStr">
              <x:is>
                <x:t>Salmon</x:t>
              </x:is>
            </x:c>
            <x:c s="2" t="inlineStr">
              <x:is>
                <x:t>Trout</x:t>
              </x:is>
            </x:c>
          </x:row>
          <x:row>
            <x:c s="2" t="inlineStr">
              <x:is>
                <x:t>Cat</x:t>
              </x:is>
            </x:c>
            <x:c s="2" t="inlineStr">
              <x:is>
                <x:t>Dog</x:t>
              </x:is>
            </x:c>
          </x:row>
        </x:sheetData>
      </x:worksheet>
    XML
  end

  let(:styles) do
    <<~XML
      <?xml version="1.0" encoding="utf-8"?><x:styleSheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:numFmts><x:numFmt numFmtId="181" formatCode="0" /><x:numFmt numFmtId="182" formatCode="m/d/yyyy h:mm:ss AM/PM" /><x:numFmt numFmtId="183" formatCode="dd MMMM yyyy" /></x:numFmts><x:fonts><x:font /><x:font><x:b /></x:font></x:fonts><x:fills><x:fill><x:patternFill patternType="none" /></x:fill><x:fill><x:patternFill patternType="gray125" /></x:fill></x:fills><x:borders><x:border /><x:border><x:bottom style="thin" /></x:border><x:border><x:right style="thin" /></x:border></x:borders><x:cellXfs><x:xf /><x:xf fontId="1" /><x:xf borderId="1" /><x:xf fontId="1" borderId="1" /><x:xf borderId="2" /><x:xf fontId="1" borderId="2" /><x:xf><x:alignment vertical="top" /></x:xf><x:xf fontId="1"><x:alignment vertical="top" /></x:xf><x:xf numFmtId="181" /><x:xf numFmtId="182" /><x:xf numFmtId="183" /><x:xf numFmtId="182" fontId="1" /><x:xf numFmtId="181" fontId="1" /><x:xf numFmtId="183" fontId="1" /></x:cellXfs></x:styleSheet>
    XML
  end

  let(:wonky_file) do
    TestXlsxBuilder.new(
      sheets: [sheet],
      styles: styles
    )
  end

  let(:subject) { SimpleXlsxReader::Document.new(wonky_file.archive.path) }

  describe '#to_hash' do
    it 'should contain Salmon and a Dog' do
      _(subject.sheets.first.rows.to_a[0]).must_include('Salmon')
      _(subject.sheets.first.rows.to_a[1]).must_include('Dog')
    end
  end
end
