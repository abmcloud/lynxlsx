# frozen_string_literal: true
module Lynxlsx
  class ContentTypes
    def initialize
      @entries = {}
    end

    def add(part_name, content_type)
      @entries[part_name] = content_type
    end

    def add_part(part)
      add(part.part_name, part.content_type)
    end

    def write_to_buf(buf)
      buf << XML::VERSION
      buf << %(<Types xmlns="#{XML::NS_PREFIX}package/2006/content-types">)
      @entries.each do |part_name, content_type|
        buf << %(<Override PartName="/#{part_name}" ContentType="#{content_type}"/>)
      end
      buf << '</Types>'
    end

    def part_name
      '[Content_Types].xml'
    end
  end
end
