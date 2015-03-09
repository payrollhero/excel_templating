module ExcelAbstraction
  class ActiveCellReference
    extend Forwardable

    def_delegators :position, :row, :col

    def initialize(attrs = {})
      @position = ExcelAbstraction::CellReference.new(attrs)
    end

    def move(directions = {})
      directions.each do |type, times|
        self.respond_to?(type) ? self.__send__(type, times) : raise(ArgumentError.new("Movement direction is not valid."))
      end
      position
    end

    def up(times = 1)
      goto(row - times, col)
    end

    def down(times = 1)
      goto(row + times, col)
    end

    def left(times = 1)
      goto(row, col - times)
    end

    def right(times = 1)
      goto(row, col + times)
    end

    def carriage_return
      goto(row, 0)
    end

    def linefeed
      down
    end

    def newline
      carriage_return
      linefeed
    end

    def goto(row, col)
      self.position = ExcelAbstraction::CellReference.new(row: row, col: col)
    end

    def reset
      self.position = ExcelAbstraction::CellReference.new(row: 0, col: 0)
    end

    protected

    attr_accessor :position
  end
end
