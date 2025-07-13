require "xml"
require "./crystal_xlsx/formula/*"
require "./crystal_xlsx/*"

module CrystalXlsx
  VERSION = "0.1.0"

  # Excel 2007+ limits
  MAX_COLUMNS = 16_384
  MAX_ROWS = 1_048_576

  # Excel epoch for date conversion
  EPOCH = Time.utc(1899, 12, 30).to_unix_f
  DAY_IN_SECONDS = 86_400

  # Macro for common XML attributes
  macro xml_attrs(**attrs)
    {% for key, value in attrs %}
      xml.attribute({{key}}, {{value}})
    {% end %}
  end

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
