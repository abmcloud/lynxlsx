# frozen_string_literal: true
require_relative '../lib/lynxlsx'

default_font_options = { name: 'Arial', sz: 10 }.freeze

Lynxlsx::Workbook.new('styles.xlsx') do |wb|
  default_style = wb.styles.add_cell_xfs(font: default_font_options)
  header_style = wb.styles.add_cell_xfs(
    font: default_font_options.merge(b: true)
  )
  date_style = wb.styles.add_cell_xfs(
    format_code: 'DD.MM.YYYY',
    font: default_font_options.merge(color: '0000CC')
  )

  wb.write_worksheet('Sheet1') do |ws|
    ws.add_row(['Int', 'Float', 'Date', 'String'], default_style: header_style)

    10.times.each do |i|
      values = [rand(5), rand, Time.now, "String #{i}"]
      styles = [default_style, default_style, date_style, default_style]
      ws.add_row(values, styles: styles)
    end
  end
end
