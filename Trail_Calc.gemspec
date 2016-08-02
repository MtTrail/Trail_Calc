# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'Trail_Calc/version'

Gem::Specification.new do |spec|
  spec.name          = "Trail_Calc"
  spec.version       = TrailCalc::VERSION
  spec.authors       = ["Mt.Trail"]
  spec.email         = ["trail@trail4you.com"]

  spec.summary       = %q{handling OpenOffice.org's CALC from Ruby.}
  spec.description   = %q{handling OpenOffice.org's CALC from Ruby.}
  spec.homepage      = "https://github.com/MtTrail/Trail_Calc"

  spec.files         = `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test|spec|features)/}) }
  spec.bindir        = "exe"
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.12"
  spec.add_development_dependency "rake", "~> 10.0"
end
