require 'spec_helper'

describe 'cell validation' do
  class SimpleCellValidatedDocument < ExcelTemplating::Document
    template "spec/assets/valid_cell.mustache.xlsx"
    title "Valid cell test"
    organization "Unimportant"
    default_styling(
      text_wrap: 0,
      font: "Calibri",
      size: 10,
      align: :left,
    )
    data_sources :valid_foos
    sheet 1 do
      validate_cell row: 2, column: 1, with: :valid_foos
    end
  end

  subject { SimpleCellValidatedDocument.new(data) }

  let(:data) {
    {
      all_sheets:
        {
          valid_value: "foo",
          valid_foos: {
            title: "Foos",
            list: ["foo", "bar"]
          }
        }
    }
  }

  describe "#render" do
    it do
      expect do
        subject.render do |path|
          expect(path).to match_excel_content('spec/assets/valid_cell_expected.xlsx')
        end
      end.not_to raise_error
    end
  end
end
