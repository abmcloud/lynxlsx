RSpec.shared_examples 'document part' do
  it { is_expected.to respond_to :content_type }
  it { is_expected.to respond_to :part_name }
  it { is_expected.to respond_to :write_to_buf }
end
