module ExcelTemplating
  class Document::DataSourceRegistry
    include Enumerable
    extend Forwardable

    def initialize
      @source_symbols = {}
    end

    def renderer(data:)
      RegistryRenderer.new(self, data)
    end

    def add_list(source_symbol, title, list, inline)
      source_symbols[source_symbol] = RegistryList.new(source_symbols.size + 1, source_symbol, title, list, inline)
    end

    def [](source_symbol)
      source_symbols[source_symbol]
    end

    def has_registry?(source_symbol)
      source_symbols.has_key?(source_symbol)
    end

    def any_data_sheet_symbols?
      self.select {|info|
        info.data_sheet?
      }.any?
    end

    def supported_registries
      source_symbols.keys
    end

    delegate [:each] => :source_symbols_ordered

    private
    attr_reader :source_symbols

    def source_symbols_ordered
      source_symbols.values.sort_by(&:order)
    end

  end
end

require_relative 'data_source_registry/registry_renderer'
require_relative 'data_source_registry/registry_list'
