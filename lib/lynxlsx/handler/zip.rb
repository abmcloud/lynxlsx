# frozen_string_literal: true
require 'pathname'
require 'zip'

module Lynxlsx
  module Handler
    class Zip < Base
      def initialize(output)
        @output = output
      end

      def run(&block)
        case @output
        when String
          path = Pathname.new(@output)
          path.unlink if path.exist?
          path.open('wb') { |io| write_to_io(io, &block) }
        when IO, StringIO
          write_to_io(@output, &block)
        else
          raise ArgumentError
        end
      end

      def write_entry(name, data = nil)
        @buf.put_next_entry(name)
        if block_given?
          yield @buf
        elsif data.respond_to? :call
          data.call(@buf)
        else
          @buf.write(data)
        end
      end

      private

      def write_to_io(io)
        ::Zip::OutputStream.write_buffer(io) do |buf|
          @buf = buf
          yield self
          @buf = nil
        end
      end
    end
  end
end
