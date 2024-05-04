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
  # @param data [Array] The data to add to the row
  # @param format [Format] The format to apply to the row
  # @return [Row] The row that was added
  def add_row(data : CrystalXlsx::Row::ValuesTypes, format : Format? = nil)
    raise "max columns exceeded" if data.size > CrystalXlsx::MAX_COLUMNS

    row = Row.new(rows.size + 1, self, format)
    row << data

    rows << row
    row
  end

  # Add a row to the worksheet - alias for add_row
  # @param data [Array] The data to add to the row
  def <<(data : CrystalXlsx::Row::ValuesTypes)
    add_row(data)
  end

  # get a row by index
  def row(index : Int32) : Row
    rows[index]
  end

  # get a cell by row and column
  def cell(row : Int32, column : Int32) : Cell
    rows[row].try &.[column] || raise "Cell not found"
  end

  # get a cell by sheet index
  def cell(index : String) : Cell
    # split the index into row and column by first letter
    row, column = index.split(/(?=[A-Z])/)

    cell(row.to_i - 1, column.ord - 'A'.ord)
  end

  def add_formula(row : Int32, column : Int32, formula : String)
    cell = cell(row, column)
    cell.formula = formula
  end

  def add_formula(row : Int32, column : Int32, formula : CrystalXlsx::Formula)
    cell = cell(index)
    cell.formula = formula.to_s
  end

  # set the width of a column
  # @param column [Int] The column to set the width of
  # @param width [Float | Int] The width to set the column to
  def column_width(column : Int32, width : Float64 | Int32)
    width = width.to_f if width.is_a?(Int32)
    @cols.add_column_width(column, width)
  end

  # set the width of multiple columns
  # @param columns [Array] The widths to set the columns to (in order)
  def columns_width=(columns : Array(Int32))
    columns.each_with_index do |column, index|
      column_width(index, column.to_f)
    end
  end

  # set the width of multiple columns
  # @param columns [Array] The widths to set the columns to (in order)
  def columns_width=(columns : Array(Float64))
    columns.each_with_index do |column, index|
      column_width(index, column)
    end
  end

  # set pane of the worksheet
  # @param xSplit [Int] The number of columns to split by
  # @param ySplit [Int] The number of rows to split by
  # @param topLeftCell [String] The top left cell of the bottom right pane
  # @param activePane [String] The active pane
  # @param state [String] The state of the pane
  # @return [Sheetview] The sheetview object
  def set_pane(xSplit : Int32, ySplit : Int32, topLeftCell : String = "A2", activePane : String = "bottomLeft", state : String = "frozen")
    @sheetviews.add_pane(xSplit, ySplit, topLeftCell, activePane, state)
  end

  # freeze the pane of the worksheet
  # @param xSplit [Int] The number of columns to split by
  # @param ySplit [Int] The number of rows to split by
  # @param topLeftCell [String] The top left cell of the bottom right pane
  # @return [Sheetview] The sheetview object
  def freeze_pane(xSplit : Int32, ySplit : Int32, topLeftCell : String = "A2")
    set_pane(xSplit, ySplit, topLeftCell)
  end

  # freeze the first row of the worksheet
  # @param row [Int] The row to freeze
  # @return [Sheetview] The sheetview object
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
