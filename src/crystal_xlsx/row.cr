# Represents a row in an Excel worksheet.
class CrystalXlsx::Row
  alias ValuesTypes = Array(Bool | Int32 | String | Time) |
                      Array(Bool | Float32 | Int32 | String | Time) |
                      Array(Float32 | Int32 | String | Time) |
                      Array(String | Time | Int32 | Float64 | Bool) |
                      Array(Float32 | Time | Int32 | String) |
                      Array(Int32 | String) |
                      Array(String) |
                      Array(Int32) |
                      Array(Float64) |
                      Array(Bool)

  # Properties
  property number : Int32
  property cells : Array(Cell) = [] of Cell
  property format : CrystalXlsx::Format | Nil
  property worksheet : CrystalXlsx::Worksheet

  def initialize(number, worksheet, format : CrystalXlsx::Format? = nil)
    @format = format
    @number = number
    @worksheet = worksheet
  end

  # Adds values to this row.
  def add(values : ValuesTypes)
    values.each_with_index do |value, index|
      if index < @cells.size
        @cells[index].value = value
        @cells[index].format = @format
      else
        @cells << Cell.new(value, self, index, @format)
      end
    end
  end

  # Adds values to this row (alias).
  def <<(values)
    add(values)
  end

  # Gets a cell by index with bounds checking.
  def [](index)
    if index < 0 || index >= @cells.size
      raise "Cell not found"
    end
    @cells[index]
  end

  # Returns the number of cells in this row.
  def size
    @cells.size
  end

  # Returns the row data as an array.
  def to_a
    @cells.map(&.value)
  end

  # Generates the XML representation of the row.
  def to_xml(xml)
    CrystalXlsx.xml_element(:row, r: @number, spans: "1:#{size}", ht: 15, "x14ac:dyDescent": 0.2) do
      @cells.each do |cell|
        cell.to_xml(xml)
      end
    end
  end
end
