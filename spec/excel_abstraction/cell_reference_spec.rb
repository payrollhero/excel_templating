require 'spec_helper'

describe ExcelAbstraction::CellReference do
  subject { described_class.new(row: 1, col: 1) }

  describe "#<=>" do
    context "when the row of the other cell reference is same" do
      context "when the other cell reference is before" do
        let(:other){ described_class.new(row: 1, col: 0) }

        it "returns 1" do
          expect(subject <=> other).to eq 1
        end
      end

      context "when the other cell reference is same" do
        let(:other){ described_class.new(row: 1, col: 1) }

        it "returns 0" do
          expect(subject <=> other).to eq 0
        end
      end

      context "when the other cell reference is after" do
        let(:other){ described_class.new(row: 1, col: 2) }

        it "returns -1" do
          expect(subject <=> other).to eq -1
        end
      end
    end

    context "when the row of the other cell reference is before" do
      let(:other){ described_class.new(row: 0, col: 0) }

      it "returns 1" do
        expect(subject <=> other).to eq 1
      end
    end

    context "when the row of the other cell reference is after" do
      let(:other){ described_class.new(row: 2, col: 0) }

      it "returns -1" do
        expect(subject <=> other).to eq -1
      end
    end
  end

  describe "#succ" do
    its(:succ){ is_expected.to eq described_class.new(row: 1, col: 2) }
  end

  describe "#to_s" do
    its(:to_s){ is_expected.to eq "B2"}
  end

  describe "#to_cell_reference" do
    its(:to_cell_reference){ is_expected.to eq subject }
  end

  describe "#to_ary" do
    its(:to_ary){ is_expected.to eq [1, 1] }
  end

  describe "#to_a" do
    its(:to_a){ is_expected.to eq [1, 1] }
  end
end
