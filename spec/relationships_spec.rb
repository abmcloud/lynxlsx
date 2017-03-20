# frozen_string_literal: true
RSpec.describe Lynxlsx::Relationships do
  it_behaves_like 'document part'

  describe '#any?' do
    it { expect(subject.any?).to eq false }
  end

  describe '#empty?' do
    it { expect(subject.empty?).to eq true }
  end

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

  context 'worksheet rels' do
    subject { described_class.new('xl/worksheets/sheet1.xml') }

    describe '#part_name' do
      it { expect(subject.part_name).to eq 'xl/worksheets/_rels/sheet1.xml.rels' }
    end
  end
end
