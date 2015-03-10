module ExcelTemplating
  # Represents a data source list used for validation of cell information
  class Document::DataSourceRegistry::RegistryList
    # @param [Integer] order
    # @param [Symbol] symbol
    # @param [String] title
    # @param [Array<String>|Symbol] list
    # @param [TrueClass|FalseClass] inline
    def initialize(order, symbol, title, list, inline)
      @title = title
      @list = list
      @inline = inline
      @order = order
      @symbol = symbol
      pre_validate!
    end

    attr_reader :title, :order, :symbol, :list

    # @return [Boolean] Is this to be rendered inline?
    def inline?
      @inline
    end

    # @return [Boolean] Is this to be rendered on the data sheet?
    def data_sheet?
      !inline?
    end

    # @param [Hash] data The data object from which the document is being rendered
    # @return [Array<String>] The validation objects
    def items(data)
      if list == :from_data
        data[symbol]
      else
        list
      end
    end

    private

    def pre_validate!
      unless list.is_a?(Array) || list == :from_data
        raise ArgumentError, "List must be an array or :from_data"
      end
    end
  end
end
