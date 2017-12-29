Gem::Specification.new do |s|
  s.name        = 'h_to_e'
  s.version     = '0.0.0'
  s.date        = '2017-12-29'
  s.summary     = "Generate excel from different kind of hashesh"
  s.description = "Generate excel from different kind of hashesh"
  s.authors     = "Prashant Kumar Mishra"
  s.email       = 'pmmishra78@gmail.com'
  s.files       = ["hash_to_excel_creator.gemspec".freeze, "lib/hash_to_excel_creator.rb".freeze]
  # s.homepage    = 'https://github.com/1990prashant/excel_from_hash'
  s.license     = 'MIT'
  
  s.add_dependency(%q<bundler>.freeze, ["~> 1.8"])
  s.add_dependency(%q<rake>.freeze, ["~> 10.0"])
  s.add_dependency(%q<axlsx>.freeze)
end