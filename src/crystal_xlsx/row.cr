class CrystalXlsx::Row
  EPOCH = Time.utc(1900, 1, 1).to_unix_f
  DAY_IN_SECONDS = 86_400
  alias ValuesTypes = Array(Bool | Float32 | Int32 | String | Time) | Array(Float32 | Int32 | String | Time) | Array(String | Time | Int32 | Float64 | Bool) | Array(Float32 | Time | Int32 | String) | Array(Int32 | String) | Array(String) | Array(Int32) | Array(Float64) | Array(Bool)
  property number : Int32
  property values : ValuesTypes
  property format : CrystalXlsx::Format | Nil

  def initialize(data : ValuesTypes, format : CrystalXlsx::Format? = nil)
    @values = data
    @format = format
    @number = 0
  end

  def to_a
    @values
  end

  def to_xml(xml)
    format_id = @format.try(&.index)
    xml.element("row", r: @number) do
      @values.each_with_index do |value, index|
        value_format = format_id || cell_format_for_value(value)
        xml.element("c", r: column_index(index), s: value_format) do
          cell_type = cell_type(value)
          xml.attribute("t", cell_type) if cell_type
          if value.is_a?(String)
            xml.element("is") do
              xml.element("t") { xml.text(value) }
            end
          else
            xml.element("v") { xml.text(render_value(value)) }
          end
        end
      end
    end
  end

  private def cell_type(value)
    case value
    when Int32, Int64, Float64, Float32, Time
      "n"
    when Bool
      "b"
    else
      "s"
    end
  end

  private def cell_format_for_value(value)
    case value
    when Time
      22
    else
      0
    end
  end

  private def render_value(value)
    case value
    when String
      value
    when Int32
      value.to_s
    when Float64
      value.to_s
    when Bool
      value ? "1" : "0"
    when Time
      excel_serial_date(value).to_s
    else
      value.to_s
    end
  end

  private def column_index(index)
    String.build do |str|
      str << ('A'.ord + index).chr
      str << @number
    end
  end

  private def excel_serial_date(date : Time)
    (date.to_unix_f - EPOCH) / DAY_IN_SECONDS + 2
  end
end