# frozen_string_literal: true
require 'time'

RSpec.describe Lynxlsx::Worksheet do
  subject { described_class.new(1, 'Sheet1') }

  it_behaves_like 'document part'

  describe '.build_cell_xml' do
    it 'should escape special chars' do
      expect(described_class.build_cell_xml(%(<value"'&>)))
        .to match /&lt;value&quot;&#39;&amp;&gt;/
    end
  end

  describe '.time_to_oa_date' do
    let(:time) { Time.parse('2017-03-16 12:00:00 +0200') }

    it { expect(described_class.time_to_oa_date(time)).to eq 42810.5 }
  end
end
