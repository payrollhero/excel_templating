module ExcelAbstraction
  class Time < DelegateClass(Float)
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
      adjustment < raw_value ? 2.days : 1.day
    end

    def convert(raw_value)
      (raw_value - reference + adjust(raw_value)).to_f / 1.day
    end
  end
end
