# frozen_string_literal: true
module Lynxlsx
  class Relationships
    attr_reader :part_name

    def self.next_rid
      @counter ||= 0
      @counter += 1
      "rId#{@counter}"
    end

    def initialize(name = nil)
      @name = name
      @entries = {}
      @prefix = ''
      @part_name =
        if @name.nil?
          '_rels/.rels'
        else
          m = %r{(.+/)([^/]+)$}.match(name)
          @prefix = m[1]
          "#{@prefix}_rels/#{m[2]}.rels"
        end
    end

    def add(target, type)
      @entries[target] = {
        target: remove_prefix(target),
        type: type,
        id: self.class.next_rid
      }
    end

    def rid(target)
      @entries.fetch(target)[:id]
    end

    def add_part(part)
      add(part.part_name, part.relationship_type)
    end

    def write_to_buf(buf)
      buf << %(#{XML::VERSION}<Relationships xmlns="#{XML::NS_PREFIX}package/2006/relationships">)
      @entries.each do |_, attrs|
        buf << %(<Relationship Target="#{attrs[:target]}" Type="#{attrs[:type]}" Id="#{attrs[:id]}"/>)
      end
      buf << '</Relationships>'
    end

    def content_type
      'application/vnd.openxmlformats-package.relationships+xml'
    end

    private

    def remove_prefix(target)
      @prefix.empty? ? target : target.sub(@prefix, '')
    end
  end
end
