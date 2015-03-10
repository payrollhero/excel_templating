# -*- encoding: utf-8 -*-

require File.expand_path('../lib/excel_templating/version', __FILE__)

Gem::Specification.new do |gem|
  gem.name          = "excel_templating"
  gem.version       = ExcelTemplating::VERSION
  gem.summary       = %q{A library which allows you to slam data into excel files using mustaching.}
  gem.description   = %q{.}
  gem.license       = "MIT"
  gem.authors       = ["bramski"]
  gem.email         = "bramski@gmail.com"
  gem.homepage      = "https://github.com/payrollhero/excel_templating"

  gem.files         = `git ls-files`.split($/)
  gem.executables   = gem.files.grep(%r{^bin/}).map{ |f| File.basename(f) }
  gem.test_files    = gem.files.grep(%r{^(test|spec|features)/})
  gem.require_paths = ['lib']

  gem.add_dependency "mustache"
  gem.add_dependency "roo"
  gem.add_dependency "write_xlsx"
  gem.add_dependency "writeexcel"

  gem.add_development_dependency 'bundler', '~> 1.0'
  gem.add_development_dependency 'rake', '~> 0.8'
  gem.add_development_dependency 'rspec', '~> 2.4'
  gem.add_development_dependency 'rubygems-tasks', '~> 0.2'
  gem.add_development_dependency 'yard', '~> 0.8'
  gem.add_development_dependency 'rspec-its'
end
