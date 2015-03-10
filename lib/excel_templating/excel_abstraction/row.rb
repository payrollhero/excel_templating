module ExcelAbstraction
  class Row
    include Enumerable
    extend Forwardable

    attr_accessor :styles
    delegate [:each] => :cells

    alias :first :min
    alias :last :max

    def initialize
      @cells = []
      @styles = {}
    end

    def [](index)
      find { |cell| cell.position == index }
    end

    def <<(attrs)
      @cells << ExcelAbstraction::Cell.new(attrs)
    end

    protected

    attr_accessor :cells
  end
end
