# frozen_string_literal: true
require 'pathname'
require 'zip'

module Lynxlsx
  module Handler
    class Zip < Base
      def initialize(name)
        super
        @path = Pathname.new(@name)
      end

      def run
        @path.unlink if @path.exist?
        @path.open('wb') do |f|
          ::Zip::OutputStream.write_buffer(f) do |buf|
            @buf = buf
            yield self
          end
        end
        @buf = nil
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
    end
  end
end
