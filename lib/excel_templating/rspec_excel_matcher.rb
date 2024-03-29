require 'roo'
require "roo-xls"

# Matcher for rspec 'match_excel_content'
# @example
#   expect do
#     subject.render do |path|
#       expect(path).to match_excel_content('spec/assets/spreadsheets/seven_three_expected.xlsx')
#     end
#   end
RSpec::Matchers.define :match_excel_content do |expected_excel_path|
  match do |excel_path|
    @excel_matcher = RSpec::Matchers::ExcelMatcher.new
    @excel_matcher.expected_roo = Roo::Spreadsheet.open(expected_excel_path)
    @excel_matcher.actual_roo = Roo::Spreadsheet.open(excel_path)
    @excel_matcher.match?
  end

  failure_message do |excel_path|
    messages = ["expected excel:#{expected_excel_path} to exactly match the content of #{excel_path} but did not."]
    @excel_matcher.errors.group_by(&:first).each do |sheet_name, errors|
      messages << "#{sheet_name}:"
      errors.each do |_, error|
        messages << "  #{error}:"
      end
    end
    messages.join("\n")
  end

  failure_message_when_negated do |excel_path|
    "expected excel:#{expected_excel_path} to not exactly match the content of #{excel_path} but it did."
  end
end

# Specific Matcher class helper for comparing two excel documents.
class RSpec::Matchers::ExcelMatcher

  def initialize
    @errors = []
  end

  def match?
    sheet_count_equal? && all_sheets_equal?
  end

  attr_accessor :errors, :expected_roo, :actual_roo

  private

  def check(sheet_name, check_result, fail_msg)
    if check_result
      true
    else
      errors << [sheet_name, fail_msg]
      false
    end
  end

  def sheet_count_equal?
    check(:base, expected_roo.sheets.count == actual_roo.sheets.count, "Number of sheets do not match.")
  end

  def all_sheets_equal?
    sheets(expected_roo).zip(sheets(actual_roo)).all? do |expected_sheet, actual_sheet|
      sheets_equal?(expected_sheet, actual_sheet)
    end
  end

  def sheets(roo)
    (0 .. (roo.sheets.count - 1)).map do |i|
      roo.sheet(i)
    end
  end

  def sheets_equal?(expected_workbook, actual_workbook)
    expected_workbook.sheets.all? do |sheet_name|
      expected = expected_workbook.sheet(sheet_name)
      actual = begin
        actual_workbook.sheet(sheet_name)
      rescue
        nil
      end

      check(sheet_name, actual, "Sheet names do not match.") if actual != nil
      # check(sheet_name, expected.first_row == actual.first_row, "Number of rows do not match.")
      # check(sheet_name, expected.last_row == actual.last_row, "Number of rows do not match")
      # check(sheet_name, expected.first_column == actual.first_column, "Number of columns do not match")
      # check(sheet_name, expected.last_column == actual.last_column, "Number of columns do not match")
      check_cells(sheet_name, expected, actual)

      errors.empty?
    end
  end

  def check_cells(sheet_name, sheet1, sheet2)
    (sheet1.first_row .. sheet1.last_row).each do |row|
      (sheet1.first_column .. sheet1.last_column).each do |col|
        check(sheet_name, compare_value(col, row, sheet1, sheet2), discrepancy_message(col, row, sheet1, sheet2))
      end
    end
  end

  def compare_value(col, row, sheet1, sheet2)
    value2 = sheet2.cell(row, col)
    value1 = sheet1.cell(row, col)
    if value2.class == Float
      value2 == value1.to_f
    elsif value1.class == Float
      value1 == value2.to_f
    else
      value2 == value1
    end
  end

  def discrepancy_message(col, row, sheet1, sheet2)
    cell_id = "#{column_letter(col)}#{row}"
    actual_value = sheet2.cell(row, col).inspect
    expected_value = sheet1.cell(row, col).inspect
    "Cell #{cell_id} actual:#{actual_value} expected:#{expected_value}"
  end

  def column_letter(col_number)
    column_letters[col_number - 1]
  end

  def column_letters
    @column_letters ||= ('A' .. 'ZZ').to_a
  end
end
