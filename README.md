# excel_templating

* [Homepage](https://rubygems.org/gems/excel_templating)
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

## Requirements

## Install

    $ gem install excel_templating

## Copyright

Copyright (c) 2015 payrollhero

See {file:LICENSE.txt} for details.
