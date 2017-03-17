# frozen_string_literal: true
module Lynxlsx
  module Handler
    class Base
      def self.open(name, &block)
        handler = new(name)
        handler.run(&block)
      end

      def initialize(name)
        @name = name
      end

      def run
        raise 'Not implemented'
      end
    end
  end
end
