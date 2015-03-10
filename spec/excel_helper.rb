require 'tempfile'

module ExcelHelper
  def create_excel(file, spreadsheet)
    yield if block_given?
    file.puts spreadsheet.to_s
    file.close
    Roo::Excel.new(file.path, packed: nil, file_warning: :ignore)
  end

  def test_excel_file
    file = Tempfile.new('test_xls')
    yield(file) if block_given?
    file.unlink
  end
end
