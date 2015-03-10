module ExcelTemplating
  # Define a sheet on a document
  # @example
  #   sheet 1 do
  #     repeat_row 17, with: :people
  #   end
  class Document::Sheet
    # @param [Integer] sheet_number
    def initialize(sheet_number)
      @sheet_number = sheet_number
      @repeated_rows = {}
    end

    # @param [Float] decimal_inches
    # @return [Float] inches converted to excel integer size.
    def inches(decimal_inches)
      # empirically determined number. 30.0 seems to be the measurement for 2.6 inches
      # in open office.
      (30.0 / 2.6) * decimal_inches
    end

    # @param [Hash] default default styling for all columns.
    # @param [Hash] columns specific styling for numbered columns.
    def style_columns(default:, columns: nil)
      @default_column_style = default
      @column_styles = columns
    end

    # Repeat a numbered row in the template using an array from the data
    # will result in expanding the produced excel document by a number of rows.
    # it is expected that the sheet specific data will contain :with as an Array.
    # @example
    #   repeat_row 17, with: :employee_data
    # @param [Integer] row_number
    # @param [Symbol] with
    def repeat_row(row_number, with:)
      repeated_rows[row_number] = RepeatedRow.new(row_number, with)
    end

    def default_column_style
      @default_column_style
    end

    def column_styles
      @column_styles
    end

    def sheet_data(data)
      data[sheet_number] || {}
    end

    def repeated_row?(row_number)
      repeated_rows.has_key?(row_number)
    end

    def each_row_at(row_number, sheet_data)
      if repeated_row?(row_number)
        repeater = repeated_rows[row_number]
        verify_array!(sheet_data, repeater.data_attribute)
        sheet_data[repeater.data_attribute].each_with_index do |row_data, index|
          yield({ index: index }.merge(row_data).merge(sheet_data))
        end
      else
        yield sheet_data
      end
    end

    private

    attr_reader :sheet_number, :repeated_rows

    def verify_array!(sheet_data, attribute)
      unless sheet_data[attribute].is_a?(Array)
        raise ArgumentError, "Data for sheet #{sheet_number} did not contain #{attribute} array as expected!"
      end
    end
  end
end

require_relative 'sheet/repeated_row'
