# frozen_string_literal: true
module Lynxlsx
  class Workbook
    DEFAULT_HANDLER = :Zip

    attr_reader :styles

    def initialize(name, options = {})
      @options = options
      @worksheets = []

      @root_rels = Relationships.new
      @root_rels.add_part(self)

      @workbook_rels = Relationships.new(part_name)

      @content_types = ContentTypes.new
      @content_types.add_part(self)
      @content_types.add_part(@root_rels)
      @content_types.add_part(@workbook_rels)

      @styles = Styles.new

      @handler = handler_class.new(name)
      @handler.run do
        yield self
        @handler.write_entry(part_name, method(:write_to_buf))
        write_styles
        write_rels
        write_content_types
      end
    end

    def write_worksheet(name, &block)
      ws = Worksheet.new(@worksheets.count + 1, name)
      @worksheets << ws
      @content_types.add_part(ws)
      @workbook_rels.add_part(ws)
      @handler.write_entry(ws.part_name) { |buf| ws.write_to_buf(buf, &block) }
      ws.tables.each do |table|
        @content_types.add_part(table)
        @handler.write_entry(table.part_name, table.method(:write_to_buf))
      end
      if ws.relationships.any?
        @handler.write_entry(ws.relationships.part_name, ws.relationships.method(:write_to_buf))
      end
    end

    def part_name
      'xl/workbook.xml'
    end

    def content_type
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
    end

    def relationship_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
    end

    def write_to_buf(buf)
      buf << XML::VERSION
      buf << %(<workbook xmlns="#{XML::NS_SSML}" xmlns:r="#{XML::NS_REL}"><workbookPr date1904="false"/><sheets>)
      @worksheets.each do |ws|
        rid = @workbook_rels.rid(ws.part_name)
        buf << %(<sheet name="#{ws.name}" sheetId="#{ws.id}" r:id="#{rid}"/>)
      end
      buf << '</sheets></workbook>'
    end

    private

    def write_styles
      @content_types.add_part(@styles)
      @workbook_rels.add_part(@styles)
      @handler.write_entry(@styles.part_name, @styles.method(:write_to_buf))
    end

    def write_content_types
      @handler.write_entry(@content_types.part_name, @content_types.method(:write_to_buf))
    end

    def write_rels
      @handler.write_entry(@root_rels.part_name, @root_rels.method(:write_to_buf))
      @handler.write_entry(@workbook_rels.part_name, @workbook_rels.method(:write_to_buf))
    end

    def handler_class
      Handler.const_get(@options[:handler] || DEFAULT_HANDLER)
    end
  end
end
