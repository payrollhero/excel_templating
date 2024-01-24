require 'spec_helper'

describe ExcelAbstraction::Sheet do
  let(:spreadsheet) { ExcelAbstraction::SpreadSheet.new }

  subject { spreadsheet.workbook.active_sheet }

  describe "#header" do
    it "sets the header" do
      test_excel_file do |file|
        excel = create_excel(file, spreadsheet) do
          subject.header("Foo")
        end

        expect(excel.cell("A", 1)).to eq "Foo"
        expect(excel.font("A", 1)).to be_bold
      end
    end
  end

  describe "#headers" do
    it "sets the headers" do
      test_excel_file do |file|
        excel = create_excel(file, spreadsheet) do
          subject.header(%w[Foo Bar])
        end

        expect(excel.cell("A", 1)).to eq "Foo"
        expect(excel.font("A", 1)).to be_bold
        expect(excel.cell("B", 1)).to eq "Bar"
        expect(excel.font("B", 1)).to be_bold
      end
    end
  end

  describe "#cell" do
    it "sets the cell value" do
      test_excel_file do |file|
        excel = create_excel(file, spreadsheet) do
          subject.cell("Foo")
        end

        expect(excel.cell("A", 1)).to eq "Foo"
      end
    end
  end

  describe "#cells" do
    it "sets the cell values" do
      test_excel_file do |file|
        excel = create_excel(file, spreadsheet) do
          subject.cells(%w[Foo Bar])
        end

        expect(excel.cell("A", 1)).to eq "Foo"
        expect(excel.cell("B", 1)).to eq "Bar"
      end
    end
  end

  describe "#merge" do
    it "merges the cells" do
      test_excel_file do |file|
        excel = create_excel(file, spreadsheet) do
          subject.merge(1, "Foo")
          subject.cell("Bar")
        end

        expect(excel.cell("A", 1)).to eq "Foo"
        expect(excel.cell("B", 1)).to be_nil
        expect(excel.cell("C", 1)).to eq "Bar"
      end
    end
  end

  describe "#style_row" do
    # hard to test; already tested manually
  end

  describe "#style_col" do
    # hard to test; already tested manually
  end
end
