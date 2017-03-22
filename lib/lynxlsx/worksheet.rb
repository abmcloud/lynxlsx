# frozen_string_literal: true
require 'cgi/escape'

module Lynxlsx
  class Worksheet
    attr_reader :id, :name, :relationships, :tables, :last_col, :last_row

    def initialize(id, name, options = {})
      @id = id
      @name = name
      @columns = []
      @relationships = Relationships.new(part_name)
      @tables = []
      @last_col = 0
      @last_row = 0
      @column_indexes = ColumnIndexes.new

      if options.key?(:columns)
        @columns = options[:columns].map.with_index do |column_options, i|
          column = { min: i + 1, max: i + 1 }
          column[:customWidth] = 1 if column_options.key?(:width)
          column_options.merge(column)
        end
      end
    end

    def last_cell
      "#{@column_indexes[@last_col - 1]}#{@last_row}" if @last_col > 0 && @last_row > 0
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
      @last_col = row_values.size
      @last_row += 1

      styles = options[:styles] || []
      default_style = options[:default_style]

      @current_buf << '<row>'
      row_values.each_with_index do |value, i|
        @current_buf << self.class.build_cell_xml(value, styles[i] || default_style)
      end
      @current_buf << '</row>'
    end

    def add_table
      table = Table.new
      yield table
      @tables << table
      @relationships.add("../../#{table.part_name}", table.relationship_type)
    end

    def write_to_buf(buf)
      buf << XML::VERSION
      buf << %(<worksheet xmlns="#{XML::NS_SSML}" xmlns:r="#{XML::NS_REL}" xml:space="preserve">)

      if @columns.any?
        buf << '<cols>'
        @columns.each { |column| buf << self.class.build_col_xml(column) }
        buf << '</cols>'
      end

      buf << '<sheetData>'
      @current_buf = buf
      yield self
      @current_buf = nil
      buf << '</sheetData>'

      if @tables.any?
        buf << %(<tableParts count="#{@tables.size}">)
        @tables.each do |table|
          rid = @relationships.rid("../../#{table.part_name}")
          buf << %(<tablePart r:id="#{rid}"/>)
        end
        buf << '</tableParts>'
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
      when nil then %(<c t="inlineStr"#{attrs}><is><t></t></is></c>)
      else %(<c t="inlineStr"#{attrs}><is><t>#{CGI.escapeHTML(value.to_s)}</t></is></c>)
      end
    end

    def self.time_to_oa_date(value)
      (value + value.utc_offset).to_f / 86400 + 25569
    end
  end
end
