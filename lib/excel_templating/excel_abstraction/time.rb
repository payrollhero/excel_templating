module ExcelAbstraction
  class Time < SimpleDelegator
    ADJUSTMENT = "1900-03-01 00:00 +00:00"
    REFERENCE = "1900-01-01 00:00 +00:00"

    attr_reader :value

    def initialize(raw_value)
      super(convert(raw_value))
    end

    def to_excel_time
      self
    end

    private

    def adjustment
      ::Time.parse(ADJUSTMENT)
    end

    def reference
      ::Time.parse(REFERENCE)
    end

    def adjust(raw_value)
      adjustment < raw_value ? two_days : one_day
    end

    def two_days
      one_day * 2.0
    end

    def one_day
      60 * 60 * 24.to_f
    end

    def convert(raw_value)
      (raw_value - reference + adjust(raw_value)).to_f / one_day
    end
  end
end
