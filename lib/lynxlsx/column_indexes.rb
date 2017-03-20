# frozen_string_literal: true
module Lynxlsx
  class ColumnIndexes
    def initialize
      @indexes = {}
    end

    def [](index)
      @indexes[index] || (@indexes[index] = self.class.evaluate_index(index))
    end

    class << self
      def evaluate_index(index)
        list = []
        while index >= 26
          list.unshift num2chr(index)
          index /= 26
        end
        list.unshift num2chr(index)
        list.join
      end

      alias [] evaluate_index

      private

      def num2chr(num)
        (65 + num % 26).chr
      end
    end
  end
end
