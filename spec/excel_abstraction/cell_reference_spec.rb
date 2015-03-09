require 'spec_helper'

describe ExcelAbstraction::CellReference do
  subject { described_class.new(row: 1, col: 1) }

  describe "#<=>" do
    context "when the row of the other cell reference is same" do
      context "when the other cell reference is before" do
        let(:other){ described_class.new(row: 1, col: 0) }

        it "returns 1" do
          (subject <=> other).should == 1
        end
      end

      context "when the other cell reference is same" do
        let(:other){ described_class.new(row: 1, col: 1) }

        it "returns 0" do
          (subject <=> other).should == 0
        end
      end

      context "when the other cell reference is after" do
        let(:other){ described_class.new(row: 1, col: 2) }

        it "returns -1" do
          (subject <=> other).should == -1
        end
      end
    end

    context "when the row of the other cell reference is before" do
      let(:other){ described_class.new(row: 0, col: 0) }

      it "returns 1" do
        (subject <=> other).should == 1
      end
    end

    context "when the row of the other cell reference is after" do
      let(:other){ described_class.new(row: 2, col: 0) }

      it "returns -1" do
        (subject <=> other).should == -1
      end
    end
  end

  describe "#succ" do
    its(:succ){ should == described_class.new(row: 1, col: 2) }
  end

  describe "#to_s" do
    its(:to_s){ should == "B2"}
  end

  describe "#to_cell_reference" do
    its(:to_cell_reference){ should == subject }
  end

  describe "#to_ary" do
    its(:to_ary){ should == [1, 1] }
  end

  describe "#to_a" do
    its(:to_a){ should == [1, 1] }
  end
end
