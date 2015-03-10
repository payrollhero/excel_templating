module ExcelTemplating
  # Simple class for representing a repeated row on a sheet.
  class Document::Sheet::RepeatedRow
    # @param [Integer] row_number
    # @param [Symol] data_attribute
    def initialize(row_number, data_attribute)
      @row_number = row_number
      @data_attribute = data_attribute
    end

    attr_reader :row_number, :data_attribute
  end
end
