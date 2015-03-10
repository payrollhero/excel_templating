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

    def items(data)
      if list == :from_data
        data[symbol]
      else
        list
      end
    end

    private

    def pre_validate!
      unless list.kind_of?(Array) || list == :from_data
        raise ArgumentError, "List must be an array or :from_data"
      end
    end
  end
end
