module ExcelTemplating
  # In charge of rendering the data source registry to the excel document
  class Document::DataSourceRegistry::RegistryRenderer
    def initialize(registry, data)
      @registry = registry
      @data = data
    end

    # @param [Symbol] source_symbol
    # @return [Hash] Gives back a hash of options which adds the validation options for the symbol
    def absolute_reference_for(source_symbol)
      unless registry.has_registry?(source_symbol)
        raise ArgumentError, "#{source_symbol} is not a defined data_source.  Defined data sources are " \
                             "#{registry.supported_registries}"
      end
      registry_info = registry[source_symbol]
      validation_options_for(registry_info)
    end

    # Wrote this registry to the specified workbook.  Uses the sheet name 'DataSource'
    # @param [ExcelAbstraction::Workbook] workbook
    def write_sheet(workbook)
      return unless registry.any_data_sheet_symbols?

      data_sheet = workbook.add_worksheet(sheet_name)
      registry.each do |registry_info|
        write_data_source_to_sheet(data_sheet, registry_info)
      end
    end

    private

    attr_reader :data, :registry

    def write_data_source_to_sheet(data_sheet, registry_info)
      column_letter = RenderHelper.letter_for(registry_info.order)
      data_sheet.write "#{column_letter}1", registry_info.title
      registry_info.items(data).each_with_index do |item, item_index|
        row_offset = item_index + 2
        data_sheet.write("#{column_letter}#{row_offset}", item)
      end
    end

    def validation_options_for(registry_info)
      if registry_info.inline?
        inline_validation_options(registry_info)
      else
        data_sheet_validation_options(registry_info)
      end
    end

    def data_sheet_validation_options(registry_info)
      row_letter = RenderHelper.letter_for(registry_info.order)
      start_column = 2
      end_column = registry_info.items(data).length + 1
      list_validation(source: "#{sheet_name}!$#{row_letter}$#{start_column}:$#{row_letter}$#{end_column}")
    end

    def inline_validation_options(registry_info)
      list_validation(source: registry_info.items(data).map(&:to_s))
    end

    def list_validation(source:)
      {
        validate: 'list',
        source: source
      }
    end

    def sheet_name
      "DataSource"
    end
  end
end
