module ExcelTemplating
  # A registry for validation data sources within the excel spreadsheet DSL
  # Supports Enumerable#each for iterating the registry entries.
  class Document::DataSourceRegistry
    include Enumerable
    extend Forwardable

    # Create an empty DataSourceRegistry
    def initialize
      @source_symbols = {}
    end

    # @param [Hash] data
    # @return [RegistryRenderer]
    def renderer(data:)
      RegistryRenderer.new(self, data)
    end

    # @param [Symbol] source_symbol
    # @param [String] title
    # @param [Array<String>|Symbol] list
    # @param [TrueClass|FalseClass] inline
    def add_list(source_symbol, title, list, inline)
      source_symbols[source_symbol] = RegistryList.new(source_symbols.size + 1, source_symbol, title, list, inline)
    end

    # @param [Symbol] source_symbol
    # @return [RegistryList]
    def [](source_symbol)
      source_symbols[source_symbol]
    end

    # @param [Symbol] source_symbol
    def has_registry?(source_symbol)
      source_symbols.has_key?(source_symbol)
    end

    # @return [TrueClass|FalseClass]
    def any_data_sheet_symbols?
      select do |info|
        info.data_sheet?
      end.any?
    end

    # @return [Array<Symbol>]
    def supported_registries
      source_symbols.keys
    end

    delegate [:each] => :ordered_registries

    private

    attr_reader :source_symbols

    def ordered_registries
      source_symbols.values.sort_by(&:order)
    end

  end
end

require_relative 'data_source_registry/registry_renderer'
require_relative 'data_source_registry/registry_list'
