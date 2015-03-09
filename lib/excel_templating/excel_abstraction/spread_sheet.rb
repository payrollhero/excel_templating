require 'tempfile'

module ExcelAbstraction
  class SpreadSheet < DelegateClass(Tempfile)
    attr_reader :workbook

    def initialize(format: :xls, skip_default_sheet: false)
      extension = format == :xls ? ".xls" : ".xlsx"
      tmp_file = Tempfile.new(["temp_spreadsheet_#{::Time.now.to_i}", extension])
      super(tmp_file)
      @workbook = ExcelAbstraction::WorkBook.new(tmp_file.path, format: format, skip_default_sheet: skip_default_sheet)
    end

    def close
      workbook.close
      yield if block_given?
      super
    end

    def to_s
      data = nil
      close do
        data = read
      end
      data
    end
  end
end
