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

  # Add a row to the worksheet
  def add_row(data : CrystalXlsx::Row::ValuesTypes, format : Format? = nil) : Row
    raise "max columns exceeded" if data.size > CrystalXlsx::MAX_COLUMNS

    row = Row.new(rows.size + 1, self, format)
    row << data
    rows << row
    row
  end

  # Add a row to the worksheet - alias for add_row
  def <<(data : CrystalXlsx::Row::ValuesTypes)
    add_row(data)
  end

  # Get a row by index
  def row(index : Int32) : Row
    rows[index]
  end

  # Get a cell by row and column
  def cell(row : Int32, column : Int32) : Cell
    rows[row][column] || raise "Cell not found"
  end

  # Get a cell by sheet index
  def cell(index : String) : Cell
    row, column = parse_cell_index(index)
    cell(row, column)
  end

  def add_formula(row : Int32, column : Int32, formula : CrystalXlsx::Formula)
    cell(row, column).formula = formula.to_s
  end

  # Set the width of a column
  def column_width(column : Int32, width : Float64 | Int32)
    @cols.add_column_width(column, width.to_f)
  end

  # Set the width of multiple columns
  def columns_width=(columns : Array(Float64 | Int32))
    columns.each_with_index do |width, index|
      column_width(index, width.to_f)
    end
  end

  # Set pane of the worksheet
  def set_pane(xSplit : Int32, ySplit : Int32, topLeftCell : String = "A2", activePane : String = "bottomLeft", state : String = "frozen")
    @sheetviews.add_pane(xSplit, ySplit, topLeftCell, activePane, state)
  end

  # Freeze the pane of the worksheet
  def freeze_pane(xSplit : Int32, ySplit : Int32, topLeftCell : String = "A2")
    set_pane(xSplit, ySplit, topLeftCell, "bottomLeft", "frozen")
  end

  # Freeze the first row of the worksheet
  def freeze_row(row : Int32)
    freeze_pane(0, row, "A#{row + 1}")
  end

  # Generate the XML for the worksheet
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
    longest_row = rows.max_by(&.size) || return "A1"
    "A1:#{('A'.ord + longest_row.size - 1).chr}#{rows.size}"
  end

  private def parse_cell_index(index : String) : {Int32, Int32}
    row, column = index.split(/(?=[A-Z])/)
    {row.to_i - 1, column.ord - 'A'.ord}
  end
end
