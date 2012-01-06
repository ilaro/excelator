require 'rubygems'
require 'rake'
require 'echoe'

Echoe.new('excelator', '0.1.0') do |p|
  p.description    = "Generate an Excel readable XML file"
  p.url            = "http://github.com/ilaro/uniquify"
  p.author         = "Ilaro"
  p.email          = "info@ilaro.com.ar"
  p.ignore_pattern = ["tmp/*", "script/*"]
  p.development_dependencies = []
end

Dir["#{File.dirname(__FILE__)}/tasks/*.rake"].sort.each { |ext| load ext }