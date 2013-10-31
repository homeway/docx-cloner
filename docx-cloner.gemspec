# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'docx/cloner/version'

Gem::Specification.new do |spec|
  spec.name          = "docx-cloner"
  spec.version       = Docx::Cloner::VERSION
  spec.authors       = ["homeway"]
  spec.email         = ["homeway.xue@gmail.com"]
  spec.description   = %q{This is a tool to generate docx file by replacing tags}
  spec.summary       = %q{generate docx file with tags}
  spec.homepage      = ""
  spec.license       = "MIT"

  spec.files         = `git ls-files`.split($/)
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.3"
  spec.add_development_dependency "rake"
end
