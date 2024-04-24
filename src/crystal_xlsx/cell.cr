class CrystalXlsx::Cell
  alias ValueTypes = String | Int32 | Int64 | Float32 | Float64 | Bool | Time

  property value : ValueTypes
  property index : Int32 = 0
  property string_index : Int32?
  property row : CrystalXlsx::Row
  property format : CrystalXlsx::Format?

  def initialize(@value : ValueTypes, @row, @index = 0, @format = nil)
    need_format_of_value
    add_string_to_shared_if_necessary
  end

  private def add_string_to_shared_if_necessary
    return unless value.is_a?(String) && shared_strings_enabled?

    @string_index = row.worksheet.workbook.try(&.shared_strings.add(@value.to_s)) || 0
  end

  def to_xml(xml)
    type = cell_type_char
    xml.element("c") do
      xml.attribute("r", column_index(row.number))
      xml.attribute("t", type) if type
      xml.attribute("s", @format.try(&.index)) if @format
      if shared_strings_enabled?
        shared_string_xml(xml)
      else
        non_shared_string_xml(xml)
      end
    end
  end

  private def shared_string_xml(xml)
    xml.element("v") do
      if value.is_a?(String)
        xml.text(@string_index.to_s)
      else
        xml.text(render_value(@value))
      end
    end
  end

  private def non_shared_string_xml(xml)
    if !value.is_a?(String)
      xml.element("v") do
        xml.text(render_value(@value))
      end
    else
      xml.element("is") do
        xml.element("t") do
          xml.text(@value.to_s)
        end
      end
    end
  end

  private def shared_strings_enabled? : Bool
    row.worksheet.workbook.try(&.enable_shared_strings?) || false
  end

  private def cell_type_char
    case value
    when String
      shared_strings_enabled? ? 's' : "inlineStr"
    when Int32, Int64, Float32, Float64 then 'n'
    when Time                           then nil # XML does not require type attribute for dates.
    when Bool                           then 'b'
    else                                     nil
    end
  end

  private def column_index(row_index : Int32)
    String.build do |str|
      str << ('A'.ord + index).chr
      str << row_index
    end
  end

  private def render_value(value) : String
    case value
    when Bool then value ? "1" : "0"
    when Time then excel_serial_date(value).to_s
    else           value.to_s
    end
  end

  private def need_format_of_value
    return unless value.is_a?(Time) || value.is_a?(Float32) || value.is_a?(Float64)

    num_id = if value.is_a?(Time)
               22
             else
               2
             end

    new_format = @format.try(&.merge(num_form_id: num_id)) || CrystalXlsx::Format.new(num_form_id: num_id)

    # add  to all formats
    @format = row.worksheet.workbook.try(&.add_format(new_format))
  end

  private def excel_serial_date(date : Time)
    (date.to_unix_f - EPOCH) / DAY_IN_SECONDS
  end
end
