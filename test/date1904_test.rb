require_relative 'test_helper'

describe SimpleXlsxReader do
  let(:date1904_file) { File.join(File.dirname(__FILE__), 'date1904.xlsx') }
  let(:subject) { SimpleXlsxReader::Document.new(date1904_file) }

  it 'supports converting dates with the 1904 date system' do
    subject.to_hash.must_equal({
      "date1904" => [[Date.parse("2014-05-01")]]
    })
  end

end
