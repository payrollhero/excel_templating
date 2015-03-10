require 'rspec'
require 'excel_templating/version'
require 'excel_templating'
require 'excel_helper'
require 'excel_templating/rspec_excel_matcher'
require 'rspec/its'
require 'byebug'

RSpec.configure do |config|
  config.include ExcelHelper
end
