# frozen_string_literal: true
RSpec.describe Lynxlsx::ColumnIndexes do
  describe '#[]' do
    it 'returns alpha column index by number' do
      expect(subject[0]).to eq 'A'
      expect(subject[25]).to eq 'Z'
      expect(subject[26]).to eq 'BA'
      expect(subject[999]).to eq 'BML'
    end
  end
end
