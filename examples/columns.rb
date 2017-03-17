# frozen_string_literal: true
require_relative '../lib/lynxlsx'

default_font_options = { name: 'Arial', sz: 10 }.freeze

Lynxlsx::Workbook.new('columns.xlsx') do |wb|
  default_style = wb.styles.add_cell_xfs(font: default_font_options)
  date_style = wb.styles.add_cell_xfs(
    format_code: 'DD.MM.YYYY',
    font: default_font_options.merge(color: '0000CC')
  )

  wb.write_worksheet('Sheet1') do |ws|
    ws.columns = [
      { style: default_style, width: 5 },
      { style: default_style, width: 15 },
      { style: date_style, width: 10 },
      { style: default_style }
    ]

    10.times.each { |i| ws.add_row [rand(5), rand, Time.now, "String #{i}"] }
  end
end
