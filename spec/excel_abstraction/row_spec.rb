require 'spec_helper'

describe ExcelAbstraction::Row do
  subject { described_class.new }

  describe "enumeration" do
    before :each do
      subject << {position: 0, val: "foo"}
      subject << {position: 1, val: "bar"}
      subject << {position: 2, val: "baz"}
    end

    it "enumerates over the cell references in the range" do
      expect(subject.reduce(""){ |str, cell| str += cell.val }).to eq "foobarbaz"
    end
  end

  describe "#[]" do
    before :each do
      subject << {position: 0, val: "foo"}
      subject << {position: 1, val: "bar"}
      subject << {position: 2, val: "baz"}
    end

    it "returns the cell with the given position" do
      expect(subject[1].val).to eq("bar")
    end
  end

  describe "#<<" do
    before :each do
      subject << {position: 0, val: "foo"}
    end

    context "when the cell to be inserted is in the same row" do
      it "inserts the cell" do
        subject << {position: 1, val: "bar"}
        expect(subject.count).to eq 2
      end
    end
  end
end
