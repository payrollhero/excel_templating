require 'spec_helper'

describe 'column validation' do
  context "rendered to the data sheet" do
    class SimpleColumnValidatedDocument < ExcelTemplating::Document
      template "spec/assets/valid_cell.mustache.xlsx"
      title "Valid cell test"
      organization "Unimportant"
      default_styling(
        text_wrap: 0,
        font: "Calibri",
        size: 10,
        align: :left,
      )
      list_source :valid_foos, title: "Foos", list: ["foo", "bar"]
      sheet 1 do
        repeat_row 2, with: :foo_data do
          validate_column 1, with: :valid_foos
        end
      end
    end

    subject { SimpleColumnValidatedDocument.new(data) }

    let(:data) {
      {
        all_sheets: {},
        1 => {
          foo_data: [
            { valid_value: 'foo'},
            { valid_value: 'bar'}
          ]
        }
      }
    }

    describe "#render" do
      it do
        expect do
          subject.render do |path|
            expect(path).to match_excel_content('spec/assets/valid_column_expected.xlsx')
          end
        end.not_to raise_error
      end
    end
  end
end
