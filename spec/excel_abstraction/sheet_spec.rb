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

        excel.cell("A", 1).should == "Foo"
        excel.font("A", 1).should be_bold
      end
    end
  end

  describe "#headers" do
    it "sets the headers" do
      test_excel_file do |file|
        excel = create_excel(file, spreadsheet) do
          subject.header(["Foo", "Bar"])
        end

        excel.cell("A", 1).should == "Foo"
        excel.font("A", 1).should be_bold
        excel.cell("B", 1).should == "Bar"
        excel.font("B", 1).should be_bold
      end
    end
  end

  describe "#cell" do
    it "sets the cell value" do
      test_excel_file do |file|
        excel = create_excel(file, spreadsheet) do
          subject.cell("Foo")
        end

        excel.cell("A", 1).should == "Foo"
      end
    end
  end

  describe "#cells" do
    it "sets the cell values" do
      test_excel_file do |file|
        excel = create_excel(file, spreadsheet) do
          subject.cells(["Foo", "Bar"])
        end

        excel.cell("A", 1).should == "Foo"
        excel.cell("B", 1).should == "Bar"
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

        excel.cell("A", 1).should == "Foo"
        excel.cell("B", 1).should be_nil
        excel.cell("C", 1).should == "Bar"
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
