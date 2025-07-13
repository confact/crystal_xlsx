class CrystalXlsx::Cols
  property column_widths : Hash(Int32, Float64) = {} of Int32 => Float64
  property? have_column_widths : Bool = false

  def add_column_width(column : Int32, width : Float64)
    @column_widths[column] = width
    @have_column_widths = true
  end

  def to_xml(xml)
    return unless have_column_widths?

    xml.element("cols") do
      width_groups.each do |group|
        xml.element("col", min: group[:min], max: group[:max], width: group[:width], customWidth: 1)
      end
    end
  end

  private def width_groups
    groups = [] of {width: Float64, min: Int32, max: Int32}
    
    # Create a separate group for each column width
    sorted_widths = @column_widths.to_a.sort_by(&.first)
    sorted_widths.each do |(index, width)|
      groups << {width: width, min: index, max: index}
    end

    groups
  end
end
