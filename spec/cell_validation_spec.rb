require 'spec_helper'

describe 'cell validation' do
  context "rendered to the data sheet" do
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
      list_source :valid_foos, title: "Foos", list: ["foo", "bar"]
      sheet 1 do
        validate_cell row: 2, column: 1, with: :valid_foos
      end
    end

    subject { SimpleCellValidatedDocument.new(data) }

    let(:data) do
      {
        all_sheets:
          {
            valid_value: "foo",
          }
      }
    end

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

  context "list taken from the data" do
    class SimpleFromDataValidatedDocument < ExcelTemplating::Document
      template "spec/assets/valid_cell.mustache.xlsx"
      title "Valid cell test"
      organization "Unimportant"
      default_styling(
        text_wrap: 0,
        font: "Calibri",
        size: 10,
        align: :left,
      )
      list_source :valid_foos, title: "Foos"
      sheet 1 do
        validate_cell row: 2, column: 1, with: :valid_foos
      end
    end

    subject { SimpleFromDataValidatedDocument.new(data) }

    let(:data) do
      {
        all_sheets:
          {
            valid_foos: ["foo", "bar"],
            valid_value: "foo",
          }
      }
    end

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

  context "rendered inline" do
    class InlineCellValidatedDocument < ExcelTemplating::Document
      template "spec/assets/valid_cell.mustache.xlsx"
      title "Valid cell test"
      list_source :valid_foos, title: "Foos", list: ["foo", "bar"], inline: true
      sheet 1 do
        validate_cell row: 2, column: 1, with: :valid_foos
      end
    end

    subject { InlineCellValidatedDocument.new(data) }

    let(:data) do
      {
        all_sheets:
          {
            valid_value: "foo",
          }
      }
    end

    describe "#render" do
      it do
        expect do
          subject.render do |path|
            expect(path).to match_excel_content('spec/assets/valid_cell_expected_inline.xlsx')
          end
        end.not_to raise_error
      end
    end
  end
end
