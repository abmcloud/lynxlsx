# frozen_string_literal: true
RSpec.describe Lynxlsx::Worksheet do
  subject { described_class.new(1, 'Sheet1') }

  it_behaves_like 'document part'
end
