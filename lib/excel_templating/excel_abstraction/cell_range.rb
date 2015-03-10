module ExcelAbstraction
  class CellRange
    include Enumerable

    alias_method :first, :min
    alias_method :last, :max

    def initialize
      @cell_references = []
    end

    def each(&block)
      cell_references.each { |cell_reference| yield(cell_reference) }
    end

    def <<(attrs)
      cell_reference = ExcelAbstraction::CellReference.new(attrs)
      raise(ArgumentError, "Must be a CellReference belonging to the same row") if last && last.row != cell_reference.row
      self.cell_references << cell_reference
    end

    protected

    attr_reader :cell_references
  end
end
