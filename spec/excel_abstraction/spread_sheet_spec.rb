require 'spec_helper'
require 'roo'

describe ExcelAbstraction::SpreadSheet do
  describe "#close" do
    it "closes the excel file" do
      subject.close
      expect(subject).to be_closed
    end

    context "when a block is passed" do
      it "executes the block before closing the excel file" do
        test = 0
        subject.close { test += 1 }
        expect(test).to eq(1)
        expect(subject).to be_closed
      end
    end
  end

  describe "#to_s" do
    it "closes the file and returns it's data" do
      file = Tempfile.new('xls')
      file.write subject.to_s
      file.close

      expect do
        Roo::Excel.new(file.path, packed: nil, file_warning: :ignore)
      end.to_not raise_exception

      File.unlink(file.path)
      expect(subject).to be_closed
    end
  end
end
