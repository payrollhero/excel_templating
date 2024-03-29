require File.expand_path('../lib/excel_templating/version', __FILE__)

Gem::Specification.new do |gem|
  gem.name          = "excel_templating"
  gem.version       = ExcelTemplating::VERSION
  gem.summary       = "A library which allows you to slam data into excel files using mustaching."
  gem.description   = "."
  gem.license       = "MIT"
  gem.authors       = ["bramski", "Mykola Kyryk"]
  gem.email         = "bramski@gmail.com"
  gem.homepage      = "https://github.com/payrollhero/excel_templating"

  gem.files         = `git ls-files`.split($/)
  gem.executables   = gem.files.grep(%r{^bin/}).map { |f| File.basename(f) }
  gem.test_files    = gem.files.grep(%r{^(test|spec|features)/})
  gem.require_paths = ['lib']
  gem.required_ruby_version = "> 2.6", "< 3.4"

  gem.add_dependency "mustache"
  gem.add_dependency "roo", ">= 2.0.0beta1", "< 3"
  gem.add_dependency "roo-xls"
  gem.add_dependency "write_xlsx"
  gem.add_dependency "writeexcel"

  gem.add_development_dependency 'bundler', '> 1.17', '< 3'
  gem.add_development_dependency 'rake', '> 0.8', '< 20'
  gem.add_development_dependency 'rspec', '~> 3.3'
  gem.add_development_dependency 'rubocop_challenger'
  gem.add_development_dependency 'rubygems-tasks', '~> 0.2'
  gem.add_development_dependency "simplecov"
  gem.add_development_dependency 'yard', '~> 0.8'
  gem.add_development_dependency 'rspec-its'
end
