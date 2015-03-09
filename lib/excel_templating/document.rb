module ExcelTemplating
  # The base document class for an ExcelTemplating.
  # Inherit from document to create your own ExcelTemplating that you may then use to generate
  # Excel Spreadsheets from the template you supply.
  # @example
  #   class MyTemplate < ExcelTemplating::Document
  #     template "my_template.mustache.xlsx"
  #     title "My title: {{my_company}}"
  #     organization "{{organization_name}}"
  #     default_styling(
  #       text_wrap: 0,
  #       font: "Calibri",
  #       size: 10,
  #       align: :left,
  #     )
  #     sheet 1 do
  #       repeat_row 17, with: :people
  #     end
  #   end
  class Document
    class << self
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
        sheet = Document::Sheet.new(sheet_number)
        sheets << sheet
        sheet.instance_eval(&block)
        nil
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
      def default_styling(font: "Calibri", size: 10, align: :left, **options)
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

    # Create a new document with given data.  'all_sheets' is available to the template on each sheet.
    # otherwise each numeric key in 'sheet_data' provides the data for that specific sheet.
    # For example {all_sheets: {foo: 'bar'}, 1 => {var1: "foo"}}
    # @param [Hash] data Hash with variables for rendering.
    def initialize(data)
      @data = data
    end

    # Render this template.
    # @example
    #   instance = MyTemplate.new(all_sheets: {foo: 1},1 => {bar: "foo"})
    #   instance.render do |file_path|
    #     FileUtils.cp(file_path, somewhere_else)
    #   end
    def render(&block)
      new_renderer.render(&block)
    end

    attr_reader :data

    private

    def new_renderer
      ExcelTemplating::Renderer.new(self)
    end

  end
end

require_relative 'document/sheet'
