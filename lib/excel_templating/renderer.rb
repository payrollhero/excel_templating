require_relative 'excel_abstraction'
require 'mustache'

module ExcelTemplating
  # Render class for ExcelTemplating Documents.  Used by the Document to render the defined document
  # with the data to a new file.  Responsible for reading the template and applying the data to it
  class Renderer
    extend Forwardable

    # @param [ExcelTemplating::Document] document Document to render with
    def initialize(document)
      @template_document = document
      @data_source_registry = document.class.data_source_registry
    end

    # Render the document provided.  Yields the path to the tempfile created.
    def render
      @spreadsheet = ExcelAbstraction::SpreadSheet.new(format: :xlsx)
      @template = Roo::Spreadsheet.open(template_path)
      @registry_renderer = data_source_registry.renderer(data: data[:all_sheets])
      apply_document_level_items
      apply_data_to_sheets
      protect_spreadsheet
      registry_renderer.write_sheet(@spreadsheet.workbook)

      @spreadsheet.close
      yield(spreadsheet.path)
    end

    private

    attr_reader :template_document, :spreadsheet, :template, :registry_renderer, :data_source_registry, :current_sheet
    delegate [:workbook] => :spreadsheet
    delegate [:data] => :template_document
    delegate [:active_sheet] => :workbook
    delegate [:active_cell_reference] => :active_sheet

    def current_row
      active_cell_reference.row
    end

    def current_col
      active_cell_reference.col
    end

    def template_path
      template_document.class.template_path
    end

    def default_format_styling
      template_document.class.document_default_styling
    end

    def protected?
      template_document.class.protected?
    end

    def sheets
      template_document.class.sheets
    end

    def apply_document_level_items
      workbook.title mustachify(template_document.class.document_title, locals: common_data_variables)
      workbook.organization mustachify(template_document.class.document_organization, locals: common_data_variables)
    end

    def common_data_variables
      stringify_keys(data[:all_sheets])
    end

    def stringify_keys(hash)
      Hash[hash.map { |k, v| [k.to_s, v] } ]
    end

    def style_columns(sheet, template_sheet)
      default_style = sheet.default_column_style
      column_styles = sheet.column_styles
      if column_styles || default_style
        roo_columns(template_sheet).each do |column_number|
          style = column_styles[column_number] || default_style
          active_sheet.style_col(column_number - 1, **style) # Note: Styling columns is zero indexed
        end
      end
    end

    def style_rows(sheet, template_sheet)
      default_style = sheet.default_row_style
      row_styles = sheet.row_styles
      if row_styles || default_style
        roo_rows(template_sheet).each do |row_number|
          style = row_styles[row_number] || default_style
          active_sheet.style_row(row_number - 1, style) # Note: Styling rows is zero indexed
        end
      end
    end

    def roo_columns(roo_sheet)
      (roo_sheet.first_column .. roo_sheet.last_column)
    end

    def roo_rows(roo_sheet)
      (roo_sheet.first_row .. roo_sheet.last_row)
    end

    def protect_spreadsheet
      active_sheet.protect if protected?
    end

    def apply_data_to_sheets
      sheets.each_with_index do |sheet, sheet_number|
        @current_sheet = sheet
        sheet_data = sheet.sheet_data(data)
        template_sheet = template.sheet(sheet_number)

        # column and row styles should be applied before writing any data to the sheet
        style_columns(sheet, template_sheet)
        # row styles have priority over column styles
        style_rows(sheet, template_sheet)

        roo_rows(template_sheet).each do |row_number|
          sheet.each_row_at(row_number, sheet_data) do |row_data|
            local_data = stringify_keys(data[:all_sheets]).merge(stringify_keys(row_data))
            roo_columns(template_sheet).each do |column_number|
              apply_data_to_cell(local_data, template_sheet, row_number, column_number)
              if sheet.validated_cell?(row_number, column_number)
                add_validation(sheet, row_number, column_number)
              end
            end
            active_sheet.next_row
          end
        end
      end
    end

    def apply_data_to_cell(local_data, template_sheet, row_number, column_number)
      template_cell = template_sheet.cell(row_number, column_number)
      font = template_sheet.font(row_number, column_number)
      format = format_for(font, row_number, column_number)
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

    def format_for(font, row_number, column_number)
      format = font_formats(font)

      set_column_lock(format, column_number) if column_is_locked?(column_number)
      set_row_lock(format, row_number)  if row_is_locked?(row_number)

      format
    end

    def column_is_locked?(column_number)
      row_or_column_is_locked? current_sheet.column_styles[column_number]
    end

    def row_is_locked?(row_number)
      row_or_column_is_locked? current_sheet.row_styles[row_number]
    end

    def row_or_column_is_locked?(row_or_column)
      row_or_column && row_or_column[:format] && row_or_column[:format].has_key?(:locked)
    end

    def row_or_column_lock_attribute(row_or_column)
      row_or_column && row_or_column[:format] && row_or_column[:format][:locked]
    end

    def set_row_lock(format, row_number)
      format.set_locked row_or_column_lock_attribute(current_sheet.row_styles[row_number])
    end

    def set_column_lock(format, column_number)
      format.set_locked row_or_column_lock_attribute(current_sheet.column_styles[column_number])
    end

    def font_formats(font=nil)
      template_font = font || Roo::Font.new

      format_details = {
        bold: template_font.bold? ? 1 : 0,
        italic: template_font.italic? ? 1 : 0,
        underline: template_font.underline? ? 1 : 0
      }.merge(default_format_styling)

      workbook.add_format format_details
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

    def add_validation(sheet, row_number, column_number)
      raise ArgumentError, "No :data_sources defined for validation!" unless data_source_registry
      source = sheet.validation_source_name(row_number, column_number)
      #Use current_row and current_col here because row_number and column_number refer to the template
      #sheet and we want to write a reference to the cell we just wrote
      active_sheet.data_validation absolute_reference(current_row + 1, current_col),
                                   registry_renderer.absolute_reference_for(source)
    end

    def absolute_reference(row_number, column_number)
      "$#{RenderHelper.letter_for(column_number)}$#{row_number}"
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
