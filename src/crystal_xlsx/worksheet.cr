class CrystalXlsx::Worksheet
  property name : String
  property rows : Array(Row)
  property workbook : CrystalXlsx::Workbook?
  property column_widths : Hash(Int32, Float64) = {} of Int32 => Float64
  property? have_column_widths : Bool = false

  def initialize(name : String, workbook : CrystalXlsx::Workbook? = nil)
    @name = name
    @workbook = workbook
    @rows = [] of Row
  end

  def add_row(data : CrystalXlsx::Row::ValuesTypes, format : Format? = nil)
    row = Row.new(rows.size + 1, self, format)
    row.add(data)

    rows << row
  end

  def column_width(column : Int32, width : Float64 | Int32)
    width = width.to_f if width.is_a?(Int32)
    @column_widths[column] = width
    @have_column_widths = true
  end

  def columns_width=(columns : Array(Int32))
    columns.each_with_index do |column, index|
      column_width(index, column.to_f)
    end
  end

  def columns_width=(columns : Array(Float64))
    columns.each_with_index do |column, index|
      column_width(index, column)
    end
  end

  # generate the xml for the worksheet
  def to_xml(io : IO)
    XML.build(io, indent: "  ", encoding: "UTF-8") do |xml|
      xml.element("worksheet", xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006", "xmlns:x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac", "mc:Ignorable": "x14ac", "xmlns:xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "xr:uid": "00000000-0001-0000-0000-000000000000", "xmlns:xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2") do
        xml.element("dimension", ref: reference) if rows.size > 0
        xml.element("sheetViews") do
          xml.element("sheetView", tabSelected: 1, workbookViewId: 0)
        end
        xml.element("sheetFormatPr", baseColWidth: 10, defaultRowHeight: 16, "x14ac:dyDescent": 0.2)
        column_widths(xml)
        xml.element("sheetData") do
          rows.each(&.to_xml(xml))
        end
        xml.element("pageMargins", left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3)
      end
    end
  end

  private def column_widths(xml)
    return unless have_column_widths?

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

    xml.element("cols") do
      width_groups.each do |group|
        xml.element("col", min: group[:min], max: group[:max], width: group[:width], customWidth: 1)
      end
    end
  end

  private def reference
    String.build do |str|
      str << "A1:"
      str << ('A'.ord + longest_row.size - 1).chr
      str << rows.size
    end
  end

  private def longest_row
    rows.max_by(&.size)
  end
end
