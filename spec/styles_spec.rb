# frozen_string_literal: true
RSpec.describe Lynxlsx::Styles do
  it_behaves_like 'document part'

  describe '.build_font_xml' do
    it 'returns a font tag' do
      expect(described_class.build_font_xml(name: 'Arial'))
        .to eq '<font><name val="Arial"/></font>'
      expect(described_class.build_font_xml(color: 'FF'))
        .to eq '<font><color rgb="FF"/></font>'
    end
  end

  describe '#part_name' do
    it { expect(subject.part_name).to eq 'xl/styles.xml' }
  end

  describe '#content_type' do
    it { expect(subject.content_type).to eq 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml' }
  end

  describe '#relationship_type' do
    it { expect(subject.relationship_type).to eq 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' }
  end
end
