module ExcelTemplating
  # Simple class for representing a repeated row on a sheet.
  class Document::Sheet::RepeatedRow
    # @param [Integer] row_number
    # @param [Symol] data_attribute
    def initialize(row_number, data_attribute)
      @row_number = row_number
      @data_attribute = data_attribute
      @column_validations = {}
    end

    attr_reader :row_number, :data_attribute

    # Validate a particular row in a repeeated set as being part of a declared data source
    # @example
    #   validate_column 5, with: :valid_foos
    # @param [Integer] column_number
    # @param [Symbol] with
    def validate_column(column_number, with:)
      @column_validations[column_number] = with
    end

    def validated_column_source(column_number)
      @column_validations[column_number]
    end

    def validated_column?(column_number)
      @column_validations.has_key?(column_number)
    end
  end
end
