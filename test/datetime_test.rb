# frozen_string_literal: true

require_relative 'test_helper'

describe SimpleXlsxReader do
  let(:datetimes_file) do
    File.join(
      File.dirname(__FILE__),
      'datetimes.xlsx'
    )
  end

  let(:subject) { SimpleXlsxReader::Document.new(datetimes_file) }

  it 'converts date_times with the correct precision' do
    _(subject.to_hash).must_equal(
      'Datetimes' =>
        [
          [Time.parse('2013-08-19 18:29:59 UTC')],
          [Time.parse('2013-08-19 18:30:00 UTC')],
          [Time.parse('2013-08-19 18:30:01 UTC')],
          [Time.parse('1899-12-30 00:30:00 UTC')]
        ]
    )
  end
end
