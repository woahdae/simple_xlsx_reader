# frozen_string_literal: true

require "bundler/gem_tasks"

require 'rake/testtask'
Rake::TestTask.new do |t|
  t.pattern = "test/**/*_test.rb"
  t.libs << 'test'
end

task default: [:test]
