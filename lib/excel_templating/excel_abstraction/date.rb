require 'delegate'

module ExcelAbstraction
  class Date < DelegateClass(Float)
    ADJUSTMENT = "1900-03-01"
    REFERENCE = "1900-01-01"

    attr_reader :value

    def initialize(raw_value)
      super(convert(raw_value))
    end

    def to_excel_date
      self
    end

    private

    def reference
      REFERENCE.to_date
    end

    def adjustment
      ADJUSTMENT.to_date
    end

    def adjust(raw_value)
      adjustment < raw_value ? 2 : 1
    end

    def convert(raw_value)
      (raw_value - reference + adjust(raw_value)).to_f
    end
  end
end
