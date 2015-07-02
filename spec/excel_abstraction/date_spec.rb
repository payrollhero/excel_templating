require "spec_helper"

describe ExcelAbstraction::Date do
  context "when date is before REFERENCE date" do
    subject { described_class.new(Date.parse("1900-01-01")) }

    it "returns 1.0" do
      expect(subject).to eq 1.0
    end
  end

  context "when date is after REFERENCE date" do
    subject { described_class.new(Date.parse("2000-01-19")) }

    it "should return 36544.50 for Jan 19, 2000 12:00" do
      expect(subject).to eq 36544.0
    end
  end

  describe "#to_excel_date" do
    subject { described_class.new(Date.parse("2012-01-23")) }

    it "returns the same object" do
      expect(subject.to_excel_date).to eq(subject)
    end
  end
end
