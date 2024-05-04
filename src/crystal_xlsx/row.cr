class CrystalXlsx::Row
  alias ValuesTypes = Array(Bool | Int32 | String | Time) | Array(Bool | Float32 | Int32 | String | Time) | Array(Float32 | Int32 | String | Time) | Array(String | Time | Int32 | Float64 | Bool) | Array(Float32 | Time | Int32 | String) | Array(Int32 | String) | Array(String) | Array(Int32) | Array(Float64) | Array(Bool)
  property number : Int32
  property cells : Array(Cell) = [] of Cell
  property format : CrystalXlsx::Format | Nil
  property worksheet : CrystalXlsx::Worksheet

  def initialize(number, worksheet, format : CrystalXlsx::Format? = nil)
    @format = format
    @number = number
    @worksheet = worksheet
  end

  def add(values : ValuesTypes)
    values.each_with_index do |value, index|
      cell = Cell.new(value, self, index, @format)
      @cells << cell
    end
  end

  def [](index)
    @cells[index]
  end

  def <<(values)
    add(values)
  end

  def size
    @cells.size
  end

  def to_a
    @cells.map(&.value)
  end

  def to_xml(xml)
    xml.element("row", r: @number, spans: "1:#{size}", ht: 15, "x14ac:dyDescent": 0.2) do
      @cells.each do |cell|
        cell.to_xml(xml)
      end
    end
  end
end
