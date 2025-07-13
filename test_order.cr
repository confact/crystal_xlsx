require "xml"

# Load the main module first
require "./src/crystal_xlsx.cr"

# Then try to load format separately
require "./src/crystal_xlsx/format.cr"

puts "Test successful"