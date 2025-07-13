require "xml"

# Load format file first
require "./src/crystal_xlsx/format.cr"

# Define the module and macros
module CrystalXlsx
  VERSION = "0.1.0"

  # Macro for XML element with attributes
  macro xml_element(name, **attrs, &block)
    xml.element({{name.id.stringify}}) do
      {% for key, value in attrs %}
        xml.attribute({{key.stringify}}, {{value}})
      {% end %}
      {{block}}
    end
  end
end

puts "Test successful"