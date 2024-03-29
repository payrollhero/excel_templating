require_relative 'document_dsl'

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
    extend DocumentDsl
    class << self
      ## The non Dsl Methods, not expected to be used as part of the document description
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
require_relative 'document/data_source_registry'
