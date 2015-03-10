module ExcelTemplating
  class Document::DataSourceRegistry::RegistryRenderer
    def initialize(source_symbols, data)
      @source_symbols = source_symbols
      @data = data
    end

    def absolute_reference_for(source_symbol)
      unless source_symbols.include?(source_symbol)
        raise ArgumentError, "#{source_symbol} is not a defined data_source.  Defined data sources are" +
              "#{source_symbols}"
      end
      validation_options_for(source_symbol)
    end

    def write_sheet(workbook)
      return unless source_symbols.length > 0

      data_sheet = workbook.add_worksheet(sheet_name)
      source_symbols.each_with_index do |source_symbol, index|
        write_data_source_to_sheet(data_sheet, source_symbol, index)
      end
    end

    private
    attr_reader :data, :source_symbols

    def write_data_source_to_sheet data_sheet, source_symbol, index
      validate_data_options(source_symbol)
      valid_data = data[source_symbol]
      column_letter = RenderHelper.letter_for(index + 1)
      data_sheet.write "#{column_letter}1", valid_data[:title]
      valid_data[:list].each_with_index do |item, item_index|
        row_offset = item_index + 2
        data_sheet.write("#{column_letter}#{row_offset}", item)
      end
    end

    def validation_options_for(source_symbol)
      validate_data_options(source_symbol)
      valid_data = data[source_symbol]
      source_index = source_symbols.index(source_symbol)
      row_letter = RenderHelper.letter_for(source_index + 1)
      start_column = 2
      end_column = valid_data.length + 1
      list_validation(ref: "#{sheet_name}!$#{row_letter}$#{start_column}:$#{row_letter}$#{end_column}")
    end

    def validate_data_options(source_symbol)
      valid_data = data[source_symbol]
      unless valid_data.kind_of?(Hash) && valid_data.has_key?(:list) && valid_data.has_key?(:title) &&
        valid_data[:list].kind_of?(Array)
        raise ArgumentError, "#{source_symbol} must have :title (string) and :list (array)"
      end
    end

    def list_validation(ref:)
      {
        validate: 'list',
        source: ref
      }
    end

    def sheet_name
      "DataSource"
    end
  end
end
