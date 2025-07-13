abstract class CrystalXlsx::Formula
  abstract def to_s(io : IO)
  # Return the Excel formula string (e.g., SUM(A1:A10))
  def excel_formula : String
    io = IO::Memory.new
    to_s(io)
    io.to_s
  end
end

# Convenience formula for string formulas
class CrystalXlsx::StringFormula < CrystalXlsx::Formula
  property formula_str : String
  def initialize(@formula_str : String)
  end
  def to_s(io : IO)
    io << formula_str
  end
  def excel_formula : String
    formula_str
  end
end