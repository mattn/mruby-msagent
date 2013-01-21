MRuby::Gem::Specification.new('mruby-msagent') do |spec|
  spec.license = 'MIT'
  spec.authors = 'mattn'
 
  if ENV['OS'] == 'Windows_NT'
    spec.linker.libraries << ['ole32', 'oleaut32', 'uuid']
  else
    raise "mruby-msagent does not support on non-Windows OSs."
  end
end
