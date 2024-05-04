class CrystalXlsx::Cols
  property column_widths : Hash(Int32, Float64) = {} of Int32 => Float64
  property? have_column_widths : Bool = false

  def add_column_width(column, width : Float64)
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
    width_groups = [] of {width: Float64, min: Int32, max: Int32}
    current_width = nil
    start_column = 0
    column = 0

    @column_widths.each do |index, width|
      column = index + 1 # Column numbering starts at 1

      if width == current_width
        # Continue the current group
        next
      else
        # Finish the current group if there was one
        if current_width
          width_groups << {width: current_width, min: start_column, max: column - 1}
        end

        # Start a new group
        current_width = width
        start_column = column
      end
    end

    if current_width
      width_groups << {width: current_width, min: start_column, max: column}
    end
    width_groups
  end
end
