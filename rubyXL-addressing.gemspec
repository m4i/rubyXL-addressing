# frozen_string_literal: true

lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'rubyXL/addressing/version'

Gem::Specification.new do |spec|
  spec.name          = 'rubyXL-addressing'
  spec.version       = RubyXL::Addressing::VERSION
  spec.authors       = ['Masaki Takeuchi']
  spec.email         = ['m.ishihara@gmail.com']

  spec.summary       = 'Addressing for rubyXL.'
  spec.description   = 'rubyXL-addressing provides addressing for rubyXL.'
  spec.homepage      = 'https://github.com/m4i/rubyXL-addressing'
  spec.license       = 'MIT'

  spec.files         = `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test|spec|features)/}) }
  spec.bindir        = 'exe'
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ['lib']

  spec.required_ruby_version = '>= 2.1.0'

  spec.add_dependency 'rubyXL', '>= 2.1.1'

  spec.add_development_dependency 'bundler',  '~> 1.11'
  spec.add_development_dependency 'rake',     '~> 10.0'
  spec.add_development_dependency 'minitest', '~> 5.0'
  spec.add_development_dependency 'appraisal'
  spec.add_development_dependency 'codeclimate-test-reporter'
end
