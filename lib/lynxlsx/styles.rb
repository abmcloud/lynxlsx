# frozen_string_literal: true
module Lynxlsx
  class Styles
    def initialize
      @num_formats = []
      @fonts = []
      @fills = []
      @borders = []
      @cell_xfs = []
    end

    def add_num_format(format_code)
      @num_formats << format_code
      @num_formats.size - 1
    end

    def add_font(options = {})
      @fonts << options
      @fonts.size - 1
    end

    def add_cell_xfs(options = {})
      xf = {}

      if options.key?(:format_code)
        xf[:numFmtId] = add_num_format(options[:format_code])
        xf[:applyNumberFormat] = true
      end

      if options.key?(:font)
        xf[:fontId] = add_font(options[:font])
        xf[:applyFont] = true
      end

      @cell_xfs << xf
      @cell_xfs.size - 1
    end

    def write_to_buf(buf)
      buf << XML::VERSION
      buf << %(<styleSheet xmlns="#{XML::NS_SSML}">)

      if @num_formats.any?
        buf << %(<numFmts count="#{@num_formats.size}">)
        @num_formats.each_with_index do |format_code, i|
          buf << %(<numFmt numFmtId="#{i}" formatCode="#{format_code}"/>)
        end
        buf << '</numFmts>'
      end

      if @fonts.any?
        buf << %(<fonts count="#{@fonts.size}">)
        @fonts.each { |font| buf << self.class.build_font_xml(font) }
        buf << '</fonts>'
      end

      if @cell_xfs.any?
        buf << %(<cellXfs count="#{@cell_xfs.size}">)
        @cell_xfs.each do |cell_xfs|
          buf << '<xf'
          cell_xfs.each { |k, v| buf << %( #{k}="#{v}") }
          buf << '></xf>'
        end
        buf << '</cellXfs>'
      end

      buf << '</styleSheet>'
    end

    def part_name
      'xl/styles.xml'
    end

    def content_type
      'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
    end

    def relationship_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
    end

    def self.build_font_xml(font)
      attrs = font.map do |name, value|
        value_attr = name == :color ? :rgb : :val
        %(<#{name} #{value_attr}="#{value}"/>)
      end

      "<font>#{attrs.join(' ')}</font>"
    end
  end
end
