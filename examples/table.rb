# frozen_string_literal: true
require_relative '../lib/lynxlsx'

columns = ['Int', 'Float', 'String'].freeze

Lynxlsx::Workbook.new('table.xlsx') do |wb|
  wb.write_worksheet('Sheet1') do |ws|
    table_first_cell = 'A1'

    ws.add_row columns
    ws.add_row [1, 10.5, 'foo']
    ws.add_row [-1, 0.1345, 'bar']

    table_last_cell = ws.last_cell

    ws.add_table do |table|
      table.columns = columns
      table.first_cell = table_first_cell
      table.last_cell = table_last_cell
    end
  end
end
