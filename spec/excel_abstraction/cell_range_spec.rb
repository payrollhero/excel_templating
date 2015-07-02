require 'spec_helper'

describe ExcelAbstraction::CellRange do
  subject { described_class.new }

  describe "enumeration" do
    before :each do
      subject << {row: 1, col: 0}
      subject << {row: 1, col: 1}
      subject << {row: 1, col: 2}
    end

    it "enumerates over the cell references in the range" do
      expect(subject.inject(0) { |sum, cell| sum + (cell.row * cell.col) }).to eq 3
    end
  end

  describe "#<<" do
    before :each do
      subject << {row: 1, col: 0}
    end

    context "when the cell to be inserted is not in the same row" do
      it "raises an exception" do
        expect { subject << {row: 2, col: 1} }.to raise_exception(ArgumentError)
      end
    end

    context "when the cell to be inserted is in the same row" do
      it "inserts the cell" do
        subject << {row: 1, col: 1}
        expect(subject.last).to eq(ExcelAbstraction::CellReference.new(row: 1, col: 1))
      end
    end
  end
end
