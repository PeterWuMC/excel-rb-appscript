# -*- encoding: utf-8 -*-

Gem::Specification.new do |gem|
  gem.authors       = ["Peter Wu"]
  gem.email         = ["peter.wu@xbridge.com"]
  gem.description   = %q{rb-appscript excel extension}
  gem.homepage      = ""

  gem.files         = `git ls-files`.split($\)
  gem.name          = "excel-rb-appscript"
  gem.require_paths = ["lib"]
  gem.version       = 0.0.0
  
  
  gem.add_runtime_dependency "rb-appscript"
  
  
end
