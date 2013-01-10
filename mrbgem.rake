MRuby::Gem::Specification.new('mruby-msagent') do |spec|
  spec.license = 'MIT'
  spec.authors = 'mattn'
 
  if ENV['OS'] == 'Windows_NT'
    spec.mruby_libs = '-lole32 -loleaut32 -luuid'
  else
    raise "mruby-msagent does not support on non-Windows OSs."
  end
end
