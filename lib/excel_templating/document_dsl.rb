module ExcelTemplating
  # The descriptor module for how to define your template class
  module DocumentDsl
    # @return [String] The template path.
    def template_path
      @template_path
    end

    # @param [String] path Set the path to the template for this document class.
    def template(path)
      raise ArgumentError, "Template path must be a string." unless path.is_a?(String)
      @template_path = path
    end

    # @return [Array<ExcelTemplating::Document::Sheet>] The sheets defined for this document class.
    def sheets
      @sheets ||= []
    end

    # Define a sheet on this document
    # @example
    #   sheet 1 do
    #     repeat_row 17, with: :people
    #   end
    # @param [Integer] sheet_number
    # @param [Proc] block
    def sheet(sheet_number, &block)
      sheet = ExcelTemplating::Document::Sheet.new(sheet_number)
      sheets << sheet
      sheet.instance_eval(&block)
      nil
    end

    # Add a list validator to the excel document
    # @example
    #   list_source :valid_foos, title: "Valid Foos", list: ['foo','bar'], inline: false
    # @param [Symbol] source_symbol symbol to registry for the validator
    # @param [String] title Title to show when displaying this validator
    # @param [Array<String>] list items to use for validation
    # @param [TrueClass|FalseClass] inline If true then the validator will be written to the document inline.
    #                                       Otherwise it will be written to a 'DataSheet'
    def list_source(source_symbol, title:, list:, inline: false)
      data_source_registry.add_list(source_symbol, title, list, inline)
    end

    def data_source_registry
      @data_source_registry ||= ExcelTemplating::Document::DataSourceRegistry.new
    end

    # Define a title for this workbook.  You may use mustaching here.
    # @param [String] string
    def title(string)
      @document_title = string
    end

    # Define the default styling to use when writing
    # a column to the worksheet.  See Writeexcel::Format
    # @param [String] font Set the font name
    # @param [Integer] size font size
    # @param [Symbol] align :left, :right, or :center
    # @param [Hash] options Additional options to pass to Format
    def default_styling(font: "Calibri", size: 10, align: :left, ** options)
      @default_styling = {
        name: font,
        size: size,
        align: align
      }.merge(options)
    end

    # Define an organization for the workbook.  May use mustaching.
    # @param [String] string
    def organization(string)
      @document_organization = string
    end

    # @return [String] The document title
    def document_title
      @document_title
    end

    # @return [String] The document organization
    def document_organization
      @document_organization
    end

    # @return [Hash] The default styling for the document
    def document_default_styling
      @default_styling || default_styling
    end
  end
end
