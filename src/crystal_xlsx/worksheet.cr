# Represents a worksheet in an Excel workbook.
# 
# Example:
# ```
# sheet = workbook.sheet("Data")
# sheet.add(["Name", "Age"])
# sheet.add(["Alice", 25])
# sheet.formula(1, 2, "SUM(B2:B10)")
# sheet.link(0, 0, "https://example.com", "More info")
# sheet.merge("A1:B1")
# ```
class CrystalXlsx::Worksheet
  # Properties
  getter name : String
  getter rows = [] of Row
  getter workbook : Workbook?
  getter hyperlinks = [] of Hyperlink
  getter merged_cells = [] of String
  
  # Private properties
  @cols = Cols.new
  @sheetviews = Sheetview.new

  def initialize(@name : String, @workbook : Workbook? = nil)
  end

  # Adds data to the worksheet as a new row.
  # 
  # Example:
  # ```
  # sheet.add(["Product", "Price", "Quantity"])
  # sheet.add(["Widget", 10.99, 5])
  # ```
  def add(data : Row::ValuesTypes, format : Format? = nil) : Row
    raise "max columns exceeded" if data.size > CrystalXlsx::MAX_COLUMNS

    row = Row.new(rows.size + 1, self, format)
    row << data
    rows << row
    row
  end

  # Adds data to the worksheet (alias for add).
  def <<(data : Row::ValuesTypes)
    add(data)
  end

  # Adds a row with a block for building the row data.
  # 
  # Example:
  # ```
  # sheet.row do |row_data|
  #   row_data << "Product"
  #   row_data << "Price"
  #   row_data << "Quantity"
  # end
  # ```
  def row(&block)
    row_data = [] of Cell::ValueTypes
    yield row_data
    add(row_data)
  end

  # Gets a row by index.
  def row(index : Int32) : Row
    rows[index]
  end

  # Gets a cell by row and column indices.
  def cell(row : Int32, column : Int32) : Cell
    rows[row][column] || raise "Cell not found"
  end

  # Gets a cell by cell reference (e.g., "A1").
  def cell(index : String) : Cell
    row, column = parse_cell_index(index)
    cell(row, column)
  end

  # Adds a formula to a cell.
  # 
  # Example:
  # ```
  # sheet.formula(1, 2, Formula::Sum.new("B2:B10"))
  # ```
  def formula(row : Int32, column : Int32, formula : Formula)
    cell(row, column).formula = formula
  end

  # Adds a formula string to a cell.
  # 
  # Example:
  # ```
  # sheet.formula(1, 2, "SUM(B2:B10)")
  # ```
  def formula(row : Int32, column : Int32, formula_str : String)
    cell(row, column).set_formula_string(formula_str)
  end

  # Sets the width of a column.
  # 
  # Example:
  # ```
  # sheet.column_width(0, 20)  # Set column A width to 20
  # ```
  def column_width(column : Int32, width : Float64 | Int32)
    @cols.add_column_width(column, width.to_f)
  end

  # Sets multiple column widths at once.
  # 
  # Example:
  # ```
  # sheet.column_widths = [20, 15, 25, 10]  # Set widths for columns A, B, C, D
  # ```
  def column_widths=(widths : Array(Float64 | Int32))
    widths.each_with_index do |width, index|
      column_width(index, width.to_f)
    end
  end

  # Freezes panes at the specified position.
  # 
  # Example:
  # ```
  # sheet.freeze_pane(1, 1, "B2")  # Freeze at row 1, column 1
  # ```
  def freeze_pane(x_split : Int32, y_split : Int32, top_left_cell : String = "A2")
    @sheetviews.add_pane(x_split, y_split, top_left_cell, "bottomLeft", "frozen")
  end

  # Freezes the first N rows.
  # 
  # Example:
  # ```
  # sheet.freeze_row(1)  # Freeze the first row
  # ```
  def freeze_row(row : Int32)
    freeze_pane(0, row, "A#{row + 1}")
  end

  # Adds a hyperlink to a cell.
  # 
  # Example:
  # ```
  # sheet.link(0, 0, "https://example.com", "Click here")
  # ```
  def link(row : Int32, column : Int32, url : String, text : String? = nil)
    cell_ref = "#{('A'.ord + column).chr}#{row + 1}"
    hyperlink = Hyperlink.new(cell_ref, url, text)
    @hyperlinks << hyperlink
  end

  # Adds a hyperlink by cell reference.
  # 
  # Example:
  # ```
  # sheet.link("A1", "https://example.com", "Click here")
  # ```
  def link(cell_ref : String, url : String, text : String? = nil)
    hyperlink = Hyperlink.new(cell_ref, url, text)
    @hyperlinks << hyperlink
  end

  # Adds a link to another worksheet cell.
  # 
  # Example:
  # ```
  # sheet.link_to_sheet(0, 0, other_sheet, "A1", "Go to other sheet")
  # ```
  def link_to_sheet(row : Int32, column : Int32, target_worksheet : Worksheet, target_cell : String, text : String? = nil)
    cell_ref = "#{('A'.ord + column).chr}#{row + 1}"
    hyperlink = Hyperlink.new_worksheet_link(cell_ref, target_worksheet, target_cell, text)
    @hyperlinks << hyperlink
  end

  # Merges a range of cells.
  # 
  # Example:
  # ```
  # sheet.merge("A1:B2")  # Merge cells A1 through B2
  # ```
  def merge(range : String)
    @merged_cells << range
  end

  # Merges cells from one reference to another.
  # 
  # Example:
  # ```
  # sheet.merge("A1", "C3")  # Merge cells A1 through C3
  # ```
  def merge(from : String, to : String)
    @merged_cells << "#{from}:#{to}"
  end

  # Generates the XML representation of the worksheet.
  def to_xml(io : IO)
    XML.build(io, indent: "  ", encoding: "UTF-8") do |xml|
      # CrystalXlsx.worksheet_xml do
      xml.element("worksheet", 
        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "xmlns:x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
        "mc:Ignorable": "x14ac",
        "xmlns:xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
        "xr:uid": "00000000-0001-0000-0000-000000000000",
        "xmlns:xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
      ) do
        xml.element("dimension", ref: reference) if rows.size > 0
        @sheetviews.to_xml(xml)
        xml.element("sheetFormatPr", baseColWidth: 10, defaultRowHeight: 16, "x14ac:dyDescent": 0.2)
        @cols.to_xml(xml)
        xml.element("sheetData") do
          rows.each(&.to_xml(xml))
        end
        
        # Merged cells
        if @merged_cells.size > 0
          xml.element("mergeCells", count: @merged_cells.size) do
            @merged_cells.each do |range|
              xml.element("mergeCell", ref: range)
            end
          end
        end
        
        # Hyperlinks
        if @hyperlinks.size > 0
          xml.element("hyperlinks") do
            @hyperlinks.each_with_index do |hyperlink, index|
              hyperlink.to_xml(xml, index + 1)
            end
          end
        end
        
        xml.element("pageMargins", left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3)
      end
      # end
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
