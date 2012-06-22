# -*- encoding: utf-8 -*-

Gem::Specification.new do |gem|
  gem.authors       = ["Peter Wu"]
  gem.email         = ["peter.wu@xbridge.com"]
  gem.description   = %q{simplified rb-appscript for excel}
  gem.summary       = %q{with basic excel dictionary}
  gem.homepage      = ""

  gem.files         = `git ls-files`.split($\)
  # gem.executables   = gem.files.grep(%r{^bin/}).map{ |f| File.basename(f) }
  gem.name          = "excel-rb-appscript"
  gem.require_paths = ["lib"]
  gem.version       = "0.0.7"
  
  
  gem.add_runtime_dependency "rb-appscript"
  
  
end
