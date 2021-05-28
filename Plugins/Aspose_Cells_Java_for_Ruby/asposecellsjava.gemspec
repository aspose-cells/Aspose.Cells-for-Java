# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'asposecellsjava/version'

Gem::Specification.new do |spec|
  spec.name          = 'asposecellsjava'
  spec.version       = Asposecellsjava::VERSION
  spec.authors       = ['Aspose Marketplace']
  spec.email         = ['marketplace@aspose.com']
  spec.summary       = %q{A Ruby gem to work with aspose.cells libraries}
  spec.description   = %q{AsposeCellsJava is a Ruby gem that can help working with Aspose.Cells libraries}
  spec.homepage      = 'https://github.com/asposecells/Aspose_Cells_Java/tree/master/Plugins/Aspose_Cells_Java_for_Ruby'
  spec.license       = 'MIT'

spec.files         = `git ls-files`.split($/)
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
spec.require_paths = ['lib']

  spec.add_development_dependency 'bundler', '>= 2.2.10'
  spec.add_development_dependency 'rake', '>= 12.3.3'
  spec.add_development_dependency 'rspec'

  spec.add_dependency 'rjb', '~> 1.5.2'

end
