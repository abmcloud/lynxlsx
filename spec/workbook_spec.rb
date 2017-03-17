# frozen_string_literal: true
RSpec.describe Lynxlsx::Workbook do
  subject { described_class.new('dummy.xlsx', handler: :Test) {} }

  it_behaves_like 'document part'

  describe '#part_name' do
    it { expect(subject.part_name).to eq 'xl/workbook.xml' }
  end

  describe '#content_type' do
    it { expect(subject.content_type).to eq 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' }
  end

  describe '#relationship_type' do
    it { expect(subject.relationship_type).to eq 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' }
  end
end
