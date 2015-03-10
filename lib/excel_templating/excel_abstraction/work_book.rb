module ExcelAbstraction
  class WorkBook < SimpleDelegator
    attr_accessor :active_sheet

    def initialize(file, format: :xls, skip_default_sheet: false)
      @format = format
      @file = file
      super(workbook)
      unless skip_default_sheet
        @active_sheet = ExcelAbstraction::Sheet.new(workbook.add_worksheet, workbook)
      end
    end


    def title(text)
      set_properties(title: text)
      self
    end

    def organization(name)
      set_properties(company: name)
      self
    end

    private

    attr_reader :format, :file

    def default_options
      {
        :font => 'Calibri',
        :size => 12,
        :align => 'center',
        :text_wrap => 1
      }
    end

    def workbook
      @workbook ||= if format == :xlsx
        WriteXLSX.new(file, default_options)
      else
        WriteExcel.new(file, default_options)
      end
    end
  end

end
