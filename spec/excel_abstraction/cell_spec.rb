require 'spec_helper'

describe ExcelAbstraction::Cell do
  subject { described_class.new(position: 1, val: "CellValue", style: {bold: true}) }

  describe "#<=>" do
    context "when the other position is before" do
      let(:other){ described_class.new(position: 0, val: nil) }

      it "return 1" do
        (subject <=> other).should == 1
      end
    end

    context "when the other position is same" do
      let(:other){ described_class.new(position: 1, val: nil) }

      it "return 0" do
        (subject <=> other).should == 0
      end
    end

    context "when the other position is after" do
      let(:other){ described_class.new(position: 2, val: nil) }

      it "return -1" do
        (subject <=> other).should == -1
      end
    end
  end

  describe "#==" do
    context "when the cells are equal" do
      let(:other){ described_class.new(position: 1, val: "CellValue", style: {bold: true}) }

      it "returns true" do
        (subject == other).should be true
      end
    end

    context "when the cells are not equal" do
      let(:other){ described_class.new(position: 1, val: "OtherCellValue", style: {bold: true}) }

      it "returns false" do
        (subject == other).should be false
      end
    end
  end

  describe "#to_cell" do
    its(:to_cell){ should == subject }
  end

end
