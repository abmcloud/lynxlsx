# frozen_string_literal: true
$LOAD_PATH.unshift File.dirname(__FILE__)

require 'lynxlsx/version'

module Lynxlsx
  autoload :ColumnIndexes, 'lynxlsx/column_indexes'
  autoload :ContentTypes, 'lynxlsx/content_types'
  autoload :Relationships, 'lynxlsx/relationships'
  autoload :Styles, 'lynxlsx/styles'
  autoload :Workbook, 'lynxlsx/workbook'
  autoload :Worksheet, 'lynxlsx/worksheet'
  autoload :XML, 'lynxlsx/xml'

  module Handler
    autoload :Base, 'lynxlsx/handler/base'
    autoload :Test, 'lynxlsx/handler/test'
    autoload :Zip, 'lynxlsx/handler/zip'
  end
end
