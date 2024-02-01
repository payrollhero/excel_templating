module ExcelAbstraction
  class Sheet < SimpleDelegator
    attr_reader :active_cell_reference

    def initialize(sheet, workbook)
      super(sheet)
      @workbook = workbook
      @active_cell_reference = ExcelAbstraction::ActiveCellReference.new
    end

    def header(value, **options)
      cell(value, bold: 1, **options)
      self
    end

    def headers(array, options = {})
      Array(array).each { |element| header(element, options) }
      self
    end

    # Fills the cell at the current pointer with the value and format or options specified
    # @param [Object] value Value to place in the cell.
    # @param [Object] options (optional) options to create a format object
    # @param [Object] format (optional) The format to use when creating this cell
    def cell(value, format: nil, type: :auto, **options)
      format = format || _format(options)
      value = Float(value) if value.is_a?(ExcelAbstraction::Time) || value.is_a?(ExcelAbstraction::Date)

      writer = case type
               when :string then method(:write_string)
               else method(:write)
               end

      writer.call(active_cell_reference.row, active_cell_reference.col, value, format)

      if options[:new_row]
        next_row
      else
        active_cell_reference.right
      end
      self
    end

    # Advance the pointer to the next row.
    # @return nil
    def next_row
      active_cell_reference.newline
    end

    def cells(array, **options)
      Array(array).each { |value| cell(value, **options) }
      self
    end

    def merge(length, value, options = {})
      format = _format(options)
      merge_range(active_cell_reference.row, active_cell_reference.col, active_cell_reference.row, active_cell_reference.col + length, value, format)
      active_cell_reference.right(length + 1)
      self
    end

    def style_row(row, properties = {})
      height = properties[:height]
      format = _format(properties[:format] || {})
      hidden = properties[:hidden] || 0
      outline_level = properties[:level] || 0
      if outline_level
        raise "Outline level can only be between 0 and 7" unless (0..7) === outline_level
      end
      collapse = properties[:collapse] || 0
      set_row(row, height, format, hidden, outline_level, collapse)
      self
    end

    # Style the numbered column
    # @param [Integer] col 0 based index of the column.
    # @param [Integer] width numeric width of the column (30 = 2.6 inches)
    # @param [Hash] format Properties to set for the format of the column
    # @param [Integer] level Outline level
    # @param [Integer] collapse 1 = collapse, 0 = do not collapse
    def style_col(col, width: nil, format: {}, level: 0, collapse: 0, hidden: nil)
      format = _format(format || {})
      if level
        raise "Outline level can only be between 0 and 7" unless (0..7) === level
      end
      set_column(col, col, width, format, hidden, level, collapse)
      self
    end

    private

    def _format(options)
      format = @workbook.add_format
      format.set_align('valign')
      options.each do |key, value|
        method = "set_#{key.to_s}".to_sym
        format.send(method, value) if format.respond_to?(method)
      end
      format
    end
  end
end
