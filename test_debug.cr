require "xml"

puts "Starting debug test..."

# Try to load files one by one
begin
  puts "Loading format.cr..."
  require "./src/crystal_xlsx/format.cr"
  puts "format.cr loaded successfully"
rescue ex
  puts "Error loading format.cr: #{ex.message}"
end

puts "Debug test completed"