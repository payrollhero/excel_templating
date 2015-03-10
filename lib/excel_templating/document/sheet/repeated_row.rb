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

    ### Dsl Methods ###

    # Validate a particular row in a repeated set as being part of a declared data source
    # @example
    #   validate_column 5, with: :valid_foos
    # @param [Integer] column_number
    # @param [Symbol] with
    def validate_column(column_number, with:)
      @column_validations[column_number] = with
    end

    ### Non Dsl Methods ###

    attr_reader :row_number, :data_attribute

    # @param [Integer] column_number
    # @return [Symbol] Registered source at that column
    def validated_column_source(column_number)
      @column_validations[column_number]
    end

    # @param [Integer] column_number
    # @return [TrueClass|FalseClass]
    def validated_column?(column_number)
      @column_validations.has_key?(column_number)
    end
  end
end
