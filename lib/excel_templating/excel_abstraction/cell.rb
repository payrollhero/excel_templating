module ExcelAbstraction
  class Cell
    attr_reader :position, :val, :styles

    def initialize(attrs = {})
      @position = Integer(attrs.fetch(:position) { raise ArgumentError.new("Position absent for ExcelAbstraction cell") })
      @val = attrs.fetch(:val) { raise ArgumentError.new("Value absent for ExcelAbstraction cell") }
      @styles = attrs.fetch(:styles) { {} }
    end

    def <=>(other)
      position <=> other.position
    end

    def ==(other)
      position == other.position && val == other.val && styles == other.styles
    end

    def to_cell
      self
    end
  end
end
