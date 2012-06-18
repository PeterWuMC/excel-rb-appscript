# -*- encoding: utf-8 -*-
# require File.expand_path('../lib/remodel/version', __FILE__)

Gem::Specification.new do |gem|
  gem.authors       = ["Peter Wu"]
  gem.email         = ["peter.wu@xbridge.com"]
  gem.description   = %q{rb-appscript excel extension}
  gem.summary       = %q{with basic excel dictionary}
  gem.homepage      = ""

  gem.files         = `git ls-files`.split($\)
  # gem.executables   = gem.files.grep(%r{^bin/}).map{ |f| File.basename(f) }
  gem.name          = "sb_remodel"
  gem.require_paths = ["lib"]
  gem.version       = Remodel::VERSION
  
  
  gem.add_runtime_dependency "rb-appscript"
  
  
end
