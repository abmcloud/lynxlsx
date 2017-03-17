# frozen_string_literal: true
module Lynxlsx
  module XML
    VERSION = %(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n)
    NS_PREFIX = 'http://schemas.openxmlformats.org/'
    NS_SSML = "#{NS_PREFIX}spreadsheetml/2006/main"
    NS_REL = "#{NS_PREFIX}officeDocument/2006/relationships"
  end
end
