require 'spec_helper'
require 'roo'

describe ExcelAbstraction::WorkBook do
  let(:spreadsheet) { ExcelAbstraction::SpreadSheet.new }

  subject { spreadsheet.workbook }

  describe "#title" do
    it "sets the title property" do
      subject.should_receive(:set_properties).with(title: 'Foo')
      subject.title('Foo')
    end
  end

  describe "#organization" do
    it "sets the company property" do
      subject.should_receive(:set_properties).with(company: 'Foo')
      subject.organization('Foo')
    end
  end
end
