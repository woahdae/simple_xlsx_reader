# frozen_string_literal: true

require_relative 'test_helper'

describe SimpleXlsxReader do
  let(:file_path) do
    File.join(
      File.dirname(__FILE__),
      'inconsistent_references.xlsx'
    )
  end

  let(:subject) { SimpleXlsxReader::Document.new(file_path) }

  it 'reads file correctly' do
    _(subject.to_hash).must_equal(
      'Sheet1' => [['Text'], [123.0], ['Text with hyperlink']]
    )
  end
end
