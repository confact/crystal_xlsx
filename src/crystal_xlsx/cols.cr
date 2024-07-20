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
    current_width = 0.0
    start_column = 0

    # Ensure the column widths are processed in order
    sorted_widths = @column_widths.to_a.sort_by(&.first)
    sorted_widths.each_with_index do |(index, width), i|
      column = index

      if width != current_width
        # Close the current group and start a new one
        if current_width
          groups << {width: current_width, min: start_column, max: column - 1}
        end
        current_width = width
        start_column = column
      elsif i == sorted_widths.size - 1
        # If last column, close the last group
        groups << {width: current_width, min: start_column, max: column}
      end
    end

    groups
  end
end
