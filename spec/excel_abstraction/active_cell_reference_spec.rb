require 'spec_helper'

describe ExcelAbstraction::ActiveCellReference do
  subject { described_class.new(row: 2, col: 3) }

  describe "#up" do
    it "moves to the cell above" do
      subject.up.should eq(ExcelAbstraction::CellReference.new(row: 1, col: 3))
    end
  end

  describe "#down" do
    it "moves to the cell above" do
      subject.down.should eq(ExcelAbstraction::CellReference.new(row: 3, col: 3))
    end
  end

  describe "#left" do
    it "moves to the cell above" do
      subject.left.should eq(ExcelAbstraction::CellReference.new(row: 2, col: 2))
    end
  end

  describe "#right" do
    it "moves to the cell above" do
      subject.right.should eq(ExcelAbstraction::CellReference.new(row: 2, col: 4))
    end
  end

  describe "#move" do
    context "when the movement commands are supported" do
      it "moves to the cell position specified" do
        subject.move(right: 2).should eq(ExcelAbstraction::CellReference.new(row: 2, col: 5))
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
      subject.carriage_return.should eq(ExcelAbstraction::CellReference.new(row: 2, col: 0))
    end
  end

  describe "#linefeed" do
    it "moves to the cell in the next row" do
      subject.linefeed.should eq(ExcelAbstraction::CellReference.new(row: 3, col: 3))
    end
  end

  describe "#newline" do
    it "moves to the cell above" do
      subject.newline.should eq(ExcelAbstraction::CellReference.new(row: 3, col: 0))
    end
  end

  describe "#goto" do
    it "moves to the cell above" do
      subject.goto(5, 5).should eq(ExcelAbstraction::CellReference.new(row: 5, col: 5))
    end
  end

  describe "#reset" do
    it "moves to the cell above" do
      subject.reset.should eq(ExcelAbstraction::CellReference.new(row: 0, col: 0))
    end
  end
end
