require "xml"

# Define the module and macros first
module CrystalXlsx
  VERSION = "0.1.0"

  # Excel 2007+ limits
  MAX_COLUMNS = 16_384
  MAX_ROWS = 1_048_576

  # Excel epoch for date conversion
  EPOCH = Time.utc(1899, 12, 30).to_unix_f
  DAY_IN_SECONDS = 86_400

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
      {% if block %}
        {{block}}
      {% end %}
    end
  end

  # Macro for simple XML element with attributes (no block)
  macro xml_element_simple(name, **attrs)
    xml.element({{name.id.stringify}}) do
      {% for key, value in attrs %}
        xml.attribute({{key.stringify}}, {{value}})
      {% end %}
    end
  end

  # Macro for cell XML generation
  macro cell_xml(cell_ref, cell_type, format_index, &block)
    xml.element("c") do
      xml.attribute("r", {{cell_ref}})
      {% if cell_type %}
        xml.attribute("t", {{cell_type}})
      {% end %}
      {% if format_index %}
        xml.attribute("s", {{format_index}})
      {% end %}
      {% if block %}
        {{block}}
      {% end %}
    end
  end

  # Macro for row XML generation
  macro row_xml(row_number, spans, &block)
    xml.element("row") do
      xml.attribute("r", {{row_number}})
      xml.attribute("spans", {{spans}})
      xml.attribute("ht", 15)
      xml.attribute("x14ac:dyDescent", 0.2)
      {{block}}
    end
  end

  # Macro for worksheet XML generation
  macro worksheet_xml(&block)
    xml.element("worksheet") do
      xml.attribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
      xml.attribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
      xml.attribute("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
      xml.attribute("xmlns:x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
      xml.attribute("mc:Ignorable", "x14ac")
      xml.attribute("xmlns:xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
      xml.attribute("xr:uid", "00000000-0001-0000-0000-000000000000")
      xml.attribute("xmlns:xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
      {{block}}
    end
  end

  # Macro for workbook XML generation
  macro workbook_xml(&block)
    xml.element("workbook") do
      xml.attribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
      xml.attribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
      xml.attribute("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
      xml.attribute("mc:Ignorable", "x15")
      xml.attribute("xmlns:x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main")
      {{block}}
    end
  end

  # Macro for style sheet XML generation
  macro stylesheet_xml(&block)
    xml.element("styleSheet") do
      xml.attribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
      {{block}}
    end
  end

  # Convenience method to create a workbook
  def self.create(&block)
    workbook = Workbook.new
    yield workbook
    workbook
  end

  # Convenience method to create a workbook and save it
  def self.create(filename : String, &block)
    workbook = create(&block)
    workbook.save(filename)
    workbook
  end
end

# Now load the files in dependency order
require "./crystal_xlsx/format.cr"
require "./crystal_xlsx/row.cr"
require "./crystal_xlsx/cell.cr"
require "./crystal_xlsx/worksheet.cr"
require "./crystal_xlsx/workbook.cr"
require "./crystal_xlsx/style.cr"
require "./crystal_xlsx/theme.cr"
require "./crystal_xlsx/shared_string.cr"
require "./crystal_xlsx/hyperlink.cr"
require "./crystal_xlsx/cols.cr"
require "./crystal_xlsx/sheetview.cr"
require "./crystal_xlsx/pane.cr"
require "./crystal_xlsx/rels.cr"
require "./crystal_xlsx/doc_props_core.cr"
require "./crystal_xlsx/doc_props_app.cr"

# Load formula files last
require "./crystal_xlsx/formula/formula.cr"
require "./crystal_xlsx/formula/link.cr"
require "./crystal_xlsx/formula/sum.cr"