class CrystalXlsx::Worksheet
  property name : String
  property rows : Array(Row)
  property workbook : CrystalXlsx::Workbook? = nil
  property column_widths : Hash(Int32, Float64) = {} of Int32 => Float64
  property have_column_widths : Bool = false

  def initialize(name : String)
    @name = name
    @rows = [] of Row
  end

  def add_row(data : CrystalXlsx::Row::ValuesTypes, format : Format? = nil)
    row = Row.new(data, format)
    row.number = rows.size + 1

    # Add string values to the workbook's shared strings, if the workbook exists
    workbook.try &.shared_strings.try &.add(row.values)

    rows << row
  end

  def set_column_width(column : Int32, width : Float64 | Int32)
    width = width.to_f if width.is_a?(Int32)
    @column_widths[column] = width
    @have_column_widths = true
  end

  def set_columns_width(columns : Array(Int32))
    columns.each_with_index do |column, index|
      set_column_width(index, column.to_f)
    end
  end

  def set_columns_width(columns : Array(Float64))
    columns.each_with_index do |column, index|
      set_column_width(index, column)
    end
  end

  # generate the xml for the worksheet
  def to_xml(io : IO)
    XML.build(io, indent: "  ", encoding: "UTF-8") do |xml|
      xml.element("worksheet", xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006", "xmlns:x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac", "mc:Ignorable": "x14ac", "xmlns:xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "xr:uid": "00000000-0001-0000-0000-000000000000", "xmlns:xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2") do
        xml.element("dimension", ref: "A1:#{('A'.ord + rows.first.values.size - 1).chr}#{rows.size}") if rows.size > 0
        xml.element("sheetViews") do
          xml.element("sheetView", tabSelected: 1, workbookViewId: 0)
        end
        xml.element("sheetFormatPr", defaultRowHeight: 15, baseColWidth: 10, defaultColWidth: 8.8, "x14ac:dyDescent": 0.2)
        if @have_column_widths
          xml.element("cols") do
            @column_widths.each { |column, width| xml.element("col", min: column, max: column, width: width, customWidth: 1) }
          end
        end
        xml.element("sheetData") do
          rows.each { |row| row.to_xml(xml) }
        end
        xml.element("pageMargins", left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3)
      end
    end
  end 
end