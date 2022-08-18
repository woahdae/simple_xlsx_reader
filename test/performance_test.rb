# frozen_string_literal: true

require_relative 'test_helper'
require 'minitest/benchmark'

describe 'SimpleXlsxReader Benchmark' do
  # n is 0-indexed for us, then converted to 1-indexed for excel
  def sheet_with_n_rows(row_count)
    acc = +""
    acc <<
      <<~XML
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
      XML

    row_count.times.each do |n|
      n += 1
      acc <<
        <<~XML
          <row>
            <c r='A#{n}' s='0'>
              <v>Cell A#{n}</v>
            </c>
            <c r='B#{n}' s='1'>
              <v>2.4</v>
            </c>
            <c r='C#{n}' s='2'>
              <v>30687</v>
            </c>
            <c r='D#{n}' t='inlineStr' s='0'>
              <is><t>Cell D#{n}</t></is>
            </c>

            <c r='E#{n}' s='0'>
              <v>Cell E#{n}</v>
            </c>
            <c r='F#{n}' s='1'>
              <v>2.4</v>
            </c>
            <c r='G#{n}' s='2'>
              <v>30687</v>
            </c>
            <c r='H#{n}' t='inlineStr' s='0'>
              <is><t>Cell H#{n}</t></is>
            </c>

            <c r='I#{n}' s='0'>
              <v>Cell I#{n}</v>
            </c>
            <c r='J#{n}' s='1'>
              <v>2.4</v>
            </c>
            <c r='K#{n}' s='2'>
              <v>30687</v>
            </c>
            <c r='L#{n}' t='inlineStr' s='0'>
              <is><t>Cell L#{n}</t></is>
            </c>
          </row>
        XML
    end

    acc <<
      <<~XML
          </sheetData>
        </worksheet>
      XML
  end

  let(:styles) do
    # s='0' above refers to the value of numFmtId at cellXfs index 0,
    # which is in this case 'General' type
    styles =
      <<-XML
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <cellXfs count="1">
            <xf numFmtId="0" />
            <xf numFmtId="2" />
            <xf numFmtId="14" />
          </cellXfs>
        </styleSheet>
      XML
  end

  before do
    @xlsxs = {}

    # Every new sheet has one more row
    self.class.bench_range.each do |num_rows|
      @xlsxs[num_rows] =
        TestXlsxBuilder.new(
          sheets: [sheet_with_n_rows(num_rows)],
          styles: styles
        ).archive
    end
  end

  def self.bench_range
    # Works out to a max just shy of 265k rows, which takes ~20s on my M1 Mac.
    # Second-largest is ~65k rows @ ~5s.
    max = ENV['BIG_PERF_TEST'] ? 265_000 : 66_000
    bench_exp(100, max, 4)
  end

  bench_performance_linear 'parses sheets in linear time', 0.999 do |n|
    SimpleXlsxReader.open(@xlsxs[n].path).sheets[0].rows.each(headers: true) {|_row| }
  end
end
