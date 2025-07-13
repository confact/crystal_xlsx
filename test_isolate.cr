require "xml"

# Load files one by one to isolate the issue
require "./src/crystal_xlsx/cell.cr"
puts "cell.cr loaded"

require "./src/crystal_xlsx/cols.cr"
puts "cols.cr loaded"

require "./src/crystal_xlsx/doc_props_app.cr"
puts "doc_props_app.cr loaded"

require "./src/crystal_xlsx/doc_props_core.cr"
puts "doc_props_core.cr loaded"

require "./src/crystal_xlsx/format.cr"
puts "format.cr loaded"

puts "Test successful"