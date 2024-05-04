class CrystalXlsx::Formula::Sum < CrystalXlsx::Formula
  property range : String

  def initialize(range)
    @range = range
  end

  def initialize(from, to)
    @range = "#{from}:#{to}"
  end

  def value(sheet) : String
    # sheet range (a1:b2) to Range object
    matches = range.match(/[A-Za-z](\d+):[A-Za-z](\d+)/)
    puts "matches: #{matches}"
    puts "range: #{range}"
    if captures = matches.try(&.captures)
      start_row = captures[0]?.try(&.to_i) || 1
      end_row = captures[1].try(&.to_i) || 1
    else
      matches = range.match(/[A-Za-z](\d+)/)
      start_row = end_row = matches.try(&.captures[0]?.try(&.to_i)) || 1
    end

    puts "start_row: #{start_row}"
    puts "end_row: #{end_row}"

    end_row = start_row if end_row < start_row
    end_row = (sheet.rows.size - 1) if end_row > (sheet.rows.size - 1)

    number_range = start_row..end_row
    puts "number_range: #{number_range}"
    cell_number = char_to_index(range.chars.first)

    number_range.reduce(0) do |sum, row|
      cell = sheet.cell(row - 1, cell_number - 1)
      case cell.value
      when Bool
        sum += cell.value ? 1 : 0
      when Float32, Float64
        sum += cell.value.as(Float64 | Float32).to_i
      when Int
        sum += cell.value.as(Int32 | Int64)
      else
        sum
      end
    end.to_s
  end

  private def char_to_index(char : Char) : Int32
    index = char.ord - 'A'.ord + 1
    return index
  end

  def to_s(io : IO)
    io << "SUM(#{range})"
  end
end