require 'spec_helper'

describe ExcelTemplating do
  context "alphalist 7.4 document" do
    class AlphalistClass < ExcelTemplating::Document
      template "spec/assets/alphalist_7_4.mustache.xlsx"
      title "7.4 Alphalist of Employees as of Dec 31 with previous employer within the year"
      organization "{{organization_name}}"
      default_styling(
        text_wrap: 0,
        font: "Calibri",
        size: 10,
        align: :left,
      )
      sheet 1 do
        repeat_row 17, with: :employee_data

        style_columns(
          default: {
            width: inches(1.98)
          },
          columns: {
            3 => { width: inches(1.98) },
            4 => { width: inches(1.98) },
            5 => { width: inches(0.39) }
          }
        )
      end
    end

    subject { AlphalistClass.new(data) }
    let(:organization_name) { "The Coffee Bean and Tea Leaf" }
    let(:year) { 1971 }

    let(:data) {
      {
        all_sheets:
          {
            organization_name: organization_name,
            company_tin: '220406705-0000',
            full_company_name: 'THE COFFEE BEAN AND TEA LEAF PHILIPPINES INC.',
            year: year
          },
        1 => {
          employee_data: [
            {
              tin_number: 12345,
              first_name: "JOHNSON JOHNSONVILLE",
              middle_initial: "Q",
              last_name: "JACKSON",
              enrollment_start_date: "1/1/2013",
              enrollment_end_date: "4/24/2013",
              gross_compensation_income: 123.45,
              non_taxable_13th_and_other: 124.56,
              non_taxable_de_minimis: 234.56,
              non_taxable_government_deductions: 345.67,
              non_taxable_other_compensation: 456.78,
              non_taxable_total: 567.89,
              taxable_basic_salary: 678.90,
              taxable_13th_month_and_other: 789.01,
              taxable_other_compensation: 890.12,
              taxable_total: 901.23,
              exemption_code: 'S',
              exemption_amount: 12.34,
              health_premium: 123.45,
              net_taxable_compensation_income: 234.56,
              tax_due: 345.67,
              tax_withheld: 456.78,
              tax_withheld_in_december: 567.89,
              tax_over_withheld: 678.90,
              tax_withheld_adjusted: 789.01,
              substituted_filing: 'Y',

              previous_non_taxable_13th_and_other: 0,
              previous_non_taxable_de_minimis: 0,
              previous_non_taxable_government_deductions: 0,
              previous_non_taxable_other_compensation: 0,
              previous_non_taxable_total: 0,
              previous_taxable_basic_salary: 0,
              previous_taxable_13th_month_and_other: 0,
              previous_taxable_other_compensation: 0,
              previous_taxable_total: 0,
              combined_taxable_total: 901.23,
            }
          ],
          total_comp_income: 8000,
          total_non_tax_13th_other: 9000,
          total_non_tax_de_minimis: 9000,
          total_non_tax_gov_deduct: 9000,
          total_non_tax_other: 9000,
          total_non_tax_total: 15000,
          total_tax_basic: 10000,
          total_tax_13th_other: 120000,
          total_tax_other: 999,
          total_tax_total: 99999,
          total_exempt: 999,
          total_premium: 4567.8,
          total_net_tax_comp: 987.98,
          total_tax_due: 123.56,
          total_tax_withheld: 234.56,
          total_tax_withheld_in_december: 975246.78,
          total_tax_over_withheld: 987987.88,
          total_tax_withheld_adjusted: 98787.89,

          previous_total_non_tax_13th_other: 0,
          previous_total_non_tax_de_minimis: 0,
          previous_total_non_tax_gov_deduct: 0,
          previous_total_non_tax_other: 0,
          previous_total_non_tax_total: 0,
          previous_total_tax_basic: 0,
          previous_total_tax_13th_other: 0,
          previous_total_tax_other: 0,
          previous_total_tax_total: 0,
          combined_total_tax_total: 99998,

        }
      }
    }

    describe "#render" do
      it do
        expect do
          subject.render do |path|
            expect(path).to match_excel_content('spec/assets/alphalist_seven_four_expected.xlsx')
          end
        end.not_to raise_error
      end
    end
  end
end
