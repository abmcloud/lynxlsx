# frozen_string_literal: true
module Lynxlsx
  module Handler
    class Test < Base
      def run
        yield self if block_given?
      end

      # Stub
      def write_entry(*)
      end
    end
  end
end
