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
      @validated_cells = {}
    end

    ### Sheet Dsl Methods ####

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
    def repeat_row(row_number, with:, &block)
      repeated_rows[row_number] = RepeatedRow.new(row_number, with)
      repeated_rows[row_number].instance_eval(&block) if block_given?
    end

    # Validate a particular cell using a declared data source
    # @example
    #   validate_cell row: 1, column :5, with: :valid_foos
    # @param [Integer] row
    # @param [Integer] column
    # @param [Symbol] with
    def validate_cell(row:, column:, with:)
      validated_cells["#{row}:#{column}"] = with
    end


    #### Non DSL Methods ###

    def default_column_style
      @default_column_style
    end

    def column_styles
      @column_styles
    end

    def sheet_data(data)
      data[sheet_number] || {}
    end

    # @param [Integer] row_number
    def repeated_row?(row_number)
      repeated_rows.has_key?(row_number)
    end

    # @param [Integer] row_number
    # @param [Integer] column_number
    def validated_cell?(row_number, column_number)
      (repeated_row?(row_number) && repeated_rows[row_number].validated_column?(column_number)) ||
        validated_cells.has_key?("#{row_number}:#{column_number}")
    end

    # @param [Integer] row_number
    # @param [Integer] column_number
    # @return [Symbol] The registered symbol for that row & column or Nil
    def validation_source_name(row_number, column_number)
      if repeated_row?(row_number)
        repeated_rows[row_number].validated_column_source(column_number)
      else
        validated_cells["#{row_number}:#{column_number}"]
      end
    end

    # Repeat each row of the data if it is repeated, yielding each item in succession.
    # @param [Integer] row_number
    # @param [Hash] sheet_data Data for this sheet
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

    attr_reader :sheet_number, :repeated_rows, :validated_cells

    def verify_array!(sheet_data, attribute)
      unless sheet_data[attribute].is_a?(Array)
        raise ArgumentError, "Data for sheet #{sheet_number} did not contain #{attribute} array as expected!"
      end
    end
  end
end

require_relative 'sheet/repeated_row'
