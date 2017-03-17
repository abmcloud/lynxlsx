# frozen_string_literal: true
module Lynxlsx
  class ColumnIndexes
    def initialize
      @indexes = {}
    end

    def [](index)
      @indexes[index] || (@indexes[index] = evaluate_index(index))
    end

    private

    def num2chr(num)
      (65 + num % 26).chr
    end

    def evaluate_index(index)
      list = []
      while index >= 26
        list.unshift num2chr(index)
        index /= 26
      end
      list.unshift num2chr(index)
      list.join
    end
  end
end
