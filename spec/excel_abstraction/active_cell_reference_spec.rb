require 'spec_helper'

describe ExcelAbstraction::ActiveCellReference do
  subject { described_class.new(row: 2, col: 3) }

  describe "#up" do
    it "moves to the cell above" do
      expect(subject.up).to eq(ExcelAbstraction::CellReference.new(row: 1, col: 3))
    end
  end

  describe "#down" do
    it "moves to the cell above" do
      expect(subject.down).to eq(ExcelAbstraction::CellReference.new(row: 3, col: 3))
    end
  end

  describe "#left" do
    it "moves to the cell above" do
      expect(subject.left).to eq(ExcelAbstraction::CellReference.new(row: 2, col: 2))
    end
  end

  describe "#right" do
    it "moves to the cell above" do
      expect(subject.right).to eq(ExcelAbstraction::CellReference.new(row: 2, col: 4))
    end
  end

  describe "#move" do
    context "when the movement commands are supported" do
      it "moves to the cell position specified" do
        expect(subject.move(right: 2)).to eq(ExcelAbstraction::CellReference.new(row: 2, col: 5))
      end
    end

    context "when the movement commands are not supported" do
      it "moves to the cell above" do
        expect { subject.move(diagonal: 2) }.to raise_exception(ArgumentError)
      end
    end
  end

  describe "#carriage_return" do
    it "moves to the cell at the beginning of the row" do
      expect(subject.carriage_return).to eq(ExcelAbstraction::CellReference.new(row: 2, col: 0))
    end
  end

  describe "#linefeed" do
    it "moves to the cell in the next row" do
      expect(subject.linefeed).to eq(ExcelAbstraction::CellReference.new(row: 3, col: 3))
    end
  end

  describe "#newline" do
    it "moves to the cell above" do
      expect(subject.newline).to eq(ExcelAbstraction::CellReference.new(row: 3, col: 0))
    end
  end

  describe "#goto" do
    it "moves to the cell above" do
      expect(subject.goto(5, 5)).to eq(ExcelAbstraction::CellReference.new(row: 5, col: 5))
    end
  end

  describe "#reset" do
    it "moves to the cell above" do
      expect(subject.reset).to eq(ExcelAbstraction::CellReference.new(row: 0, col: 0))
    end
  end
end
