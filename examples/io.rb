# frozen_string_literal: true
require_relative '../lib/lynxlsx'

File.open('io.xlsx', 'wb') do |f|
  Lynxlsx::Workbook.new(f) do |wb|
    wb.write_worksheet('Sheet1') do |ws|
      ws.add_row ['Int', 'Float', 'String']
      ws.add_row [1, 10.5, 'foo']
      ws.add_row [-1, 0.1345, 'bar']
    end
  end
end
