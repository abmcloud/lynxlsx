# frozen_string_literal: true
module Lynxlsx
  class Table
    attr_reader :id, :name
    attr_accessor :autofilter, :columns, :first_cell, :last_cell

    def self.next_id
      @counter ||= 0
      @counter += 1
    end

    def initialize(name = nil)
      @id = self.class.next_id
      @name = name || "Table#{@id}"
      @autofilter = true
      @columns = []
    end

    def write_to_buf(buf)
      buf << XML::VERSION
      buf << %(<table xmlns="#{XML::NS_SSML}" id="#{@id}" name="#{@name}" displayName="#{@name}" ref="#{@first_cell}:#{@last_cell}" totalsRowShown="0">)
      buf << %(<autoFilter ref="#{@first_cell}:#{@last_cell}"/>) if @autofilter
      buf << %(<tableColumns count="#{@columns.size}">)
      @columns.each_with_index do |name, index|
        buf << %(<tableColumn id="#{index + 1}" name="#{name}"/>)
      end
      buf << '</tableColumns></table>'
    end

    def part_name
      "xl/tables/table#{@id}.xml"
    end

    def content_type
      'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'
    end

    def relationship_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'
    end
  end
end
