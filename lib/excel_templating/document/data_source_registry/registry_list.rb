module ExcelTemplating
  class Document::DataSourceRegistry::RegistryList
    def initialize(order, symbol, title, list, inline)
      @title = title
      @list = list
      @inline = inline
      @order = order
      @symbol = symbol
      pre_validate!
    end

    attr_reader :title, :order, :symbol, :list

    def inline?
      @inline
    end

    def data_sheet?
      !inline?
    end

    alias_method :items, :list

    private

    def pre_validate!
      unless list.kind_of?(Array)
        raise ArgumentError, "List must be an array."
      end
    end
  end
end
