# frozen_string_literal: true
require 'hescape'

module Lynxlsx
  class Worksheet
    attr_reader :id, :name

    def initialize(id, name)
      @id = id
      @name = name
      @columns = []
    end

    # Writes a new row to the worksheet
    #
    # ==== Options
    #
    # * <tt>:styles</tt> - Array of style indexes for each cell value
    #
    # ==== Examples
    #
    # worksheet.add_row([1, 'foo_bar', Time.now])
    # worksheet.add_row([1, 123.45], styles: [qty_cell_style, price_cell_style])
    def add_row(row_values, options = {})
      styles = options[:styles] || []
      default_style = options[:default_style]

      @current_buf << '<row>'
      row_values.each_with_index do |value, i|
        @current_buf << self.class.build_cell_xml(value, styles[i] || default_style)
      end
      @current_buf << '</row>'
    end

    def columns=(columns)
      @columns = columns.map.with_index do |options, i|
        options.merge(min: i + 1, max: i + 1)
      end
    end

    def write_to_buf(buf)
      buf << XML::VERSION
      buf << %(<worksheet xmlns="#{XML::NS_SSML}" xmlns:r="#{XML::NS_REL}" xml:space="preserve">)

      buf << '<sheetData>'
      @current_buf = buf
      yield self
      @current_buf = nil
      buf << '</sheetData>'

      if @columns.any?
        buf << '<cols>'
        @columns.each { |column| buf << self.class.build_col_xml(column) }
        buf << '</cols>'
      end

      buf << '</worksheet>'
    end

    def part_name
      "xl/worksheets/sheet#{@id}.xml"
    end

    def content_type
      'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
    end

    def relationship_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
    end

    def self.build_col_xml(column)
      attrs = column.map { |name, value| %(#{name}="#{value}") }
      "<col #{attrs.join(' ')}/>"
    end

    def self.build_cell_xml(value, style = nil)
      attrs = %( s="#{style}") unless style.nil?

      case value
      when Numeric then %(<c#{attrs}><v>#{value}</v></c>)
      when Time then %(<c#{attrs}><v>#{time_to_oa_date(value)}</v></c>)
      else %(<c t="inlineStr"#{attrs}><is><t>#{Hescape.escape_html(value)}</t></is></c>)
      end
    end

    def self.time_to_oa_date(value)
      (value + value.utc_offset).to_f / 86400 + 25569
    end
  end
end