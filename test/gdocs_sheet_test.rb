require_relative 'test_helper'
require 'time'

describe SimpleXlsxReader do
  let(:one_sheet_file) { File.join(File.dirname(__FILE__), 'gdocs_sheet.xlsx') }
  let(:subject) { SimpleXlsxReader::Document.new(one_sheet_file) }

  it 'able to load file from google docs' do
    subject.to_hash.must_equal({
      "List 1" => [["Empty gdocs list 1"]],
      "List 2" => [["Empty gdocs list 2"]]
    })
  end

end
