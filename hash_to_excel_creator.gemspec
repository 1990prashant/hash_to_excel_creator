Gem::Specification.new do |s|
  s.name        = 'hash_to_excel_creator'
  s.version     = '0.0.0'
  s.date        = '2017-12-29'
  s.summary     = "Generate excel from different kind of hashesh"
  s.description = "Generate excel from different kind of hashesh"
  s.authors     = "Prashant Kumar Mishra"
  s.email       = 'pmmishra78@gmail.com'
  s.files       = ["hash_to_excel_creator.gemspec".freeze, "lib/hash_to_excel_creator.rb".freeze]
  # s.homepage    = 'https://github.com/1990prashant/excel_from_hash'
  s.license     = 'MIT'
  s.require_path = 'lib'
  
  if RUBY_PLATFORM =~ /java/
    s.add_dependency 'java'
  end
  s.add_dependency(%q<bundler>.freeze, ["~> 1.15"])
  s.add_dependency(%q<rake>.freeze, ["~> 12.0"])
  s.add_dependency 'axlsx', ">= 2.1.0.pre"
  
  # s.add_dependency 'zip', '>= 2.0.2'
end