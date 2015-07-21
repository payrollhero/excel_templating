# excel_templating

* [Homepage](https://github.com/payrollhero/excel_templating)
* [Documentation](http://rubydoc.info/gems/excel_templating/frames)
* [Email](mailto:bramski at gmail.com)

## Description

A library that does excel templating using mustaching.

## Features

## Examples
```ruby
    require 'excel_templating'

    class MyTemplate < ExcelTemplating::Document
      template "my_template.mustache.xlsx")
      title "My fancy report {{year}}"
      organization "{{organization_name}}"
      default_styling(
        text_wrap: 0,
        font: "Calibri",
        size: 10,
        align: :left,
      )
      sheet 1 do
        repeat_row 17, with: :repeating_data

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
```

## Cell Validation
You may validate that cells belong to a particular set of values 'dropdown'
this is done by specifying data sources for your sheet and then referencing
them in your template.
``` ruby
    class MyTemplate < ExcelTemplating::Document
      template 'my_template.mustache.xlsx'

      list_source :valid_foos, title: "Foos", list: ["foo", "bar"]
      sheet 1 do
        validate_cell row: 5, column: 1, with: :valid_foos
        repeat_row 17, with :repeating_data do
          validate_column 1, with: :valid_foos
        end
      end
    end
```

The 'list' item may be an Array or :from_data, if it says :from_data, the list
will be sourced from the same key in the 'all_sheets' data portion.

The excel templater will add an additional sheet to your generated xls
file called 'Data Sources' and 'foo' and 'bar' will be written to that sheet.
If you don't want a sheet to be created, use inline: true to write the validation
directly to the cell, NOTE there are limits on the size of the list
you may write inline.

## Protecting document
You can specify locking for any row or column. By default all cells are marked as locked.
Locking is applied when you call `protect_document` method.
``` ruby
    class MyTemplate < ExcelTemplating::Document
      template 'my_template.mustache.xlsx'
      default_styling locked: 0 # default set to not locked

      list_source :valid_foos, title: "Foos", list: ["foo", "bar"]
      sheet 1 do
        validate_cell row: 5, column: 1, with: :valid_foos
        repeat_row 17, with :repeating_data do
          validate_column 1, with: :valid_foos
        end

        # Lets lock the first row
        style_rows(
          default: { },
          rows: {
            1 => { format: { locked: 1 } }
          }
        )
      end

      # add call `protect_document` to lock specified row
      protect_document
    end
```

## Rspec Excel Matching
The library also adds an excel rspec matcher.
```ruby
    require 'excel_templating/rspec_excel_matcher'

    describe MyTemplate do
      subject { described_class.new }
      it do
        expect do
          subject.render do |path|
            expect(path).to match_excel_content('my_expected_file.xlsx')
          end
        end.not_to raise_error
      end
    end
```

## Deploying

1. Update lib/excel_templating/version.rb
2. Update ChangeLog.md
3. Commit the 2 changed files with the version number
5. Push this to git
6. Run `rake release`

## Install

    $ gem install excel_templating

## Copyright

Copyright (c) 2015 payrollhero

See {file:LICENSE.txt} for details.
