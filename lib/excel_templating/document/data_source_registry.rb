module ExcelTemplating
  class Document::DataSourceRegistry

    def initialize(source_symbols)
      @source_symbols = source_symbols
    end

    def renderer(data:)
      RegistryRenderer.new(@source_symbols, data)
    end

  end
end

require_relative 'data_source_registry/registry_renderer'
