# frozen_string_literal: true

require_relative 'test_helper'

describe SimpleXlsxReader do
  let(:lower_case_shared_strings) do
    File.join(
      File.dirname(__FILE__),
      'lower_case_sharedstrings.xlsx'
    )
  end

  let(:subject) { SimpleXlsxReader::Document.new(lower_case_shared_strings) }

  describe '#to_hash' do
    it 'should have the word Well in the first row' do
      _(subject.sheets.first.rows.to_a[0]).must_include('Well')
    end
  end
end
