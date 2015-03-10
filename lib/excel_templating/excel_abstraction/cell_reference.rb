module ExcelAbstraction
  class CellReference
    include Comparable

    COLS = ('A'..'ZZ').to_a

    attr_accessor :row, :col

    def initialize(attrs = {})
      @row = attrs.fetch(:row) { 0 }
      @col = attrs.fetch(:col) { 0 }
    end

    def <=>(other)
      other = other.to_cell_reference
      (self.row == other.row) ? (self.col <=> other.col) : (self.row <=> other.row)
    end

    def succ
      self.class.new(row: self.row, col: self.col + 1)
    end

    def to_s
      COLS[col] + (row + 1).to_s
    end

    def to_cell_reference
      self
    end

    def to_ary
      [row, col]
    end

    def to_a
      to_ary
    end
  end
end
