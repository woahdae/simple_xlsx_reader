# -*- encoding: utf-8 -*-
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'simple_xlsx_reader/version'

Gem::Specification.new do |gem|
  gem.name          = "simple_xlsx_reader"
  gem.version       = SimpleXlsxReader::VERSION
  gem.authors       = ["Woody Peterson"]
  gem.email         = ["woody@sigby.com"]
  gem.description   = %q{Read xlsx data the Ruby way}
  gem.summary       = %q{Read xlsx data the Ruby way}
  gem.homepage      = ""

  gem.add_dependency 'nokogiri'
  gem.add_dependency 'rubyzip'

  gem.add_development_dependency 'minitest', '>= 5.0'
  gem.add_development_dependency 'pry'

  gem.files         = `git ls-files`.split($/)
  gem.executables   = gem.files.grep(%r{^bin/}).map{ |f| File.basename(f) }
  gem.test_files    = gem.files.grep(%r{^(test|spec|features)/})
  gem.require_paths = ["lib"]
end
