require_relative 'test_helper'

describe SimpleXlsxReader do
  let(:lower_case_shared_strings) { File.join(File.dirname(__FILE__),
                                                'lower_case_sharedstrings.xlsx') }

  let(:subject) { SimpleXlsxReader::Document.new(lower_case_shared_strings) }


  describe '#to_hash' do
    it 'should have the word Well in the first row' do
      subject.sheets.first.rows[0].must_include('Well')
    end
  end
end
