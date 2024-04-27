class CrystalXlsx::Worksheet
  property name : String
  property rows : Array(Row)
  property workbook : CrystalXlsx::Workbook?
  @cols : CrystalXlsx::Cols = CrystalXlsx::Cols.new
  @sheetviews : CrystalXlsx::Sheetview = CrystalXlsx::Sheetview.new

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

  def <<(data : CrystalXlsx::Row::ValuesTypes)
    add_row(data)
  end

  def column_width(column : Int32, width : Float64 | Int32)
    width = width.to_f if width.is_a?(Int32)
    @cols.add_column_width(column, width)
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

  def set_pane(xSplit : Int32, ySplit : Int32, topLeftCell : String = "A2", activePane : String = "bottomLeft", state : String = "frozen")
    @sheetviews.add_pane(xSplit, ySplit, topLeftCell, activePane, state)
  end

  def freeze_pane(xSplit : Int32, ySplit : Int32, topLeftCell : String = "A2")
    set_pane(xSplit, ySplit, topLeftCell)
  end

  def freeze_row(row : Int32)
    freeze_pane(0, row, "A#{row + 1}")
  end

  # generate the xml for the worksheet
  def to_xml(io : IO)
    XML.build(io, indent: "  ", encoding: "UTF-8") do |xml|
      xml.element("worksheet", xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006", "xmlns:x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac", "mc:Ignorable": "x14ac", "xmlns:xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "xr:uid": "00000000-0001-0000-0000-000000000000", "xmlns:xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2") do
        xml.element("dimension", ref: reference) if rows.size > 0
        @sheetviews.to_xml(xml)
        xml.element("sheetFormatPr", baseColWidth: 10, defaultRowHeight: 16, "x14ac:dyDescent": 0.2)
        @cols.to_xml(xml)
        xml.element("sheetData") do
          rows.each(&.to_xml(xml))
        end
        xml.element("pageMargins", left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3)
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
