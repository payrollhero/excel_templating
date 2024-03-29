require 'simplecov'
SimpleCov.minimum_coverage 97
SimpleCov.start 'rails'

require 'rspec'
require 'excel_templating/version'
require 'excel_templating'
require 'excel_helper'
require 'excel_templating/rspec_excel_matcher'
require 'rspec/its'

RSpec.configure do |config|
  config.include ExcelHelper
end
