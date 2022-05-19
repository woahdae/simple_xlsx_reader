# frozen_string_literal: true

gem 'minitest'
require 'minitest/autorun'
require 'minitest/spec'
require 'pry'
require 'time'
require 'test_xlsx_builder'

$LOAD_PATH.unshift File.expand_path('lib')
require 'simple_xlsx_reader'
