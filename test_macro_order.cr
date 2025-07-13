require "xml"

# Define macros first
module CrystalXlsx
  VERSION = "0.1.0"

  # Macro for common cell types
  macro cell_type_for(value_type)
    case {{value_type}}
    when String then "s"
    when Int32, Int64, Float32, Float64 then "n"
    when Bool then "b"
    when Time then nil
    else nil
    end
  end

  # Macro for XML element with attributes
  macro xml_element(name, **attrs, &block)
    xml.element({{name.id.stringify}}) do
      {% for key, value in attrs %}
        xml.attribute({{key.stringify}}, {{value}})
      {% end %}
      {{block}}
    end
  end

  # Macro for cell XML generation
  macro cell_xml(cell_ref, cell_type, format_index, &block)
    xml.element("c") do
      xml.attribute("r", {{cell_ref}})
      xml.attribute("t", {{cell_type}}) if {{cell_type}}
      xml.attribute("s", {{format_index}}) if {{format_index}}
      {{block}}
    end
  end
end

# Now load the format file
require "./src/crystal_xlsx/format_new.cr"

puts "Test successful"