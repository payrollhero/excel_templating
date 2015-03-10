require "spec_helper"

describe ExcelAbstraction::Time do
  context "when date is before REFERENCE date" do
    subject { described_class.new(Time.parse("1900-01-01 00:00 +00:00")) }

    it "returns 1.0" do
      subject.should == 1.0
    end
  end

  context "when date is after REFERENCE date" do
    subject { described_class.new(Time.parse("2000-01-19 12:00")) }

    it "should return 36544.50 for Jan 19, 2000 12:00" do
      subject.should == 36544.833333333336
    end
  end

  describe "#to_excel_time" do
    subject { described_class.new(Time.parse("2012-01-23 14:00")) }

    it "returns the same object" do
      subject.to_excel_time.should eq(subject)
    end
  end
end
