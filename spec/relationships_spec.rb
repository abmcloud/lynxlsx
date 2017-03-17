# frozen_string_literal: true
RSpec.describe Lynxlsx::Relationships do
  it_behaves_like 'document part'

  describe '#content_type' do
    it { expect(subject.content_type).to eq 'application/vnd.openxmlformats-package.relationships+xml' }
  end

  context 'root rels' do
    describe '#part_name' do
      it { expect(subject.part_name).to eq '_rels/.rels' }
    end
  end

  context 'workbook rels' do
    subject { described_class.new('xl/workbook.xml') }

    describe '#part_name' do
      it { expect(subject.part_name).to eq 'xl/_rels/workbook.xml.rels' }
    end
  end
end
