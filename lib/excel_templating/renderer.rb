require_relative 'excel_abstraction'
require 'mustache'

module ExcelTemplating
  # Render class for ExcelTemplating Documents.  Used by the Document to render the defined document
  # with the data to a new file.  Responsible for reading the template and applying the data to it
  class Renderer

    # @param [ExcelTemplating::Document] document Document to render with
    def initialize(document)
      @template_document = document
    end

    # Render the document provided.  Yields the path to the tempfile created.
    def render
      @spreadsheet = ExcelAbstraction::SpreadSheet.new(format: :xlsx)
      @template = Roo::Spreadsheet.open(template_path)
      apply_document_level_items
      apply_data_to_sheets

      @spreadsheet.close
      yield(spreadsheet.path)
    end

    private

    attr_reader :template_document, :spreadsheet, :template
    delegate :workbook, to: :spreadsheet
    delegate :data, to: :template_document
    delegate :active_sheet, to: :workbook

    def template_path
      template_document.class.template_path
    end

    def default_format_styling
      template_document.class.document_default_styling
    end

    def sheets
      template_document.class.sheets
    end

    def apply_document_level_items
      locals = data[:all_sheets].stringify_keys
      workbook.title mustachify(template_document.class.document_title, locals: locals)
      workbook.organization mustachify(template_document.class.document_organization, locals: locals)
    end

    def style_columns(sheet, template_sheet)
      default_style = sheet.default_column_style
      column_styles = sheet.column_styles
      if column_styles || default_style
        roo_columns(template_sheet).each do |column_number|
          style = column_styles[column_number] || default_style
          active_sheet.style_col(column_number - 1, style) # Note: Styling columns is zero indexed
        end
      end
    end

    def roo_columns(roo_sheet)
      (roo_sheet.first_column .. roo_sheet.last_column)
    end

    def roo_rows(roo_sheet)
      (roo_sheet.first_row .. roo_sheet.last_row)
    end

    def apply_data_to_sheets
      sheets.each_with_index do |sheet, sheet_number|
        sheet_data = sheet.sheet_data(data)
        template_sheet = template.sheet(sheet_number)

        roo_rows(template_sheet).each do |row_number|
          sheet.each_row_at(row_number, sheet_data) do |row_data|
            local_data = data[:all_sheets].merge(row_data).stringify_keys
            roo_columns(template_sheet).each do |column_letter|
              apply_data_to_cell(local_data, template_sheet, row_number, column_letter)
            end
            active_sheet.next_row
          end
        end
        style_columns(sheet, template_sheet)
      end
    end

    def apply_data_to_cell(local_data, template_sheet, row_number, column_letter)
      template_cell = template_sheet.cell(row_number, column_letter)
      font = template_sheet.font(row_number, column_letter)
      format = format_for(font)
      value = mustachify(template_cell, locals: local_data)

      active_sheet.cell(
        value,
        type: type_for_value(value),
        format: format
      )
    end

    def type_for_value(value)
      looks_like_a_separator?(value) ? :string : :auto
    end

    def looks_like_a_separator?(value)
      value.is_a?(String) && value =~ /^==/
    end

    def format_for(font)
      @font_formats ||= {}
      unless @font_formats.has_key?(font)
        format = workbook.add_format(
          {
            bold: font.bold? ? 1 : 0,
            italic: font.italic? ? 1 : 0,
            underline: font.underline? ? 1 : 0
          }.merge(default_format_styling))
        @font_formats[font] = format
      end
      @font_formats[font]
    end

    def mustachify(inline_template, locals: {})
      if whole_cell_template?(inline_template)
        locals.fetch(extract_key_from_template(inline_template))
      elsif no_mustache_found?(inline_template)
        inline_template
      else
        MustacheRenderer.render(inline_template, locals)
      end
    end

    def no_mustache_found?(value)
      !value.is_a?(String) || !(value.match(/{{.+}}/))
    end

    def extract_key_from_template(template)
      template[whole_cell_template_matcher, 1]
    end

    def whole_cell_template_matcher
      /^{{([^{}]+)}}$/
    end

    def whole_cell_template?(template)
      template =~ whole_cell_template_matcher
    end

    class MustacheRenderer < Mustache

      def self.render(template, locals)
        renderer = new
        renderer.template = template
        renderer.raise_on_context_miss = true
        renderer.render(locals)
      end

    end
  end
end
