# frozen_string_literal: true
Gem::Specification.new do |s|
  s.name        = 'lynxlsx'
  s.version     = '0.1.0'
  s.summary     = 'Fast XLSX writer'
  s.authors     = ['Eugene Sheverdin']
  s.email       = ['scaint88@gmail.com']
  s.homepage    = 'https://github.com/abmcloud/lynxlsx'
  s.license     = 'Nonstandard'
  s.files       = Dir['lib/**/*.rb']
  s.test_files  = Dir['spec/**/*_spec.rb']

  s.add_runtime_dependency 'rubyzip', '~> 1.2'

  s.add_development_dependency 'rspec', '~> 3.0'
end
