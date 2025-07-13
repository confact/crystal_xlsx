require "xml"

module CrystalXlsx
  VERSION = "0.1.0"
end

class CrystalXlsx::Format
  property index : Int32 = 0
  property font_size : Int32 = 11
  property font_name : String = "Calibri"
  property bold : Bool = false
  property text_color : String? = nil
  property bg_color : String? = nil
  property border : Bool = false
  property num_form_id : Int32 = 0
  property horizontal_alignment : String? = nil
  property vertical_alignment : String? = nil

  def initialize(@num_form_id = 0, @font_size : Int32 = 11, @font_name : String = "Calibri", @bold : Bool = false, @bg_color : String? = nil, @text_color : String? = nil, @border : Bool = false, @horizontal_alignment : String? = nil, @vertical_alignment : String? = nil)
  end

  def merge(**args)
    CrystalXlsx::Format.new(
      num_form_id: args[:num_form_id]? || num_form_id,
      font_size: args[:font_size]? || font_size,
      font_name: args[:font_name]? || font_name,
      bold: args[:bold]? || bold,
      bg_color: args[:bg_color]? || bg_color,
      text_color: args[:text_color]? || text_color,
      border: args[:border]? || border,
      horizontal_alignment: args[:horizontal_alignment]? || horizontal_alignment,
      vertical_alignment: args[:vertical_alignment]? || vertical_alignment
    )
  end

  def bold? : Bool
    bold
  end

  def border? : Bool
    border
  end

  def font_id
    index
  end
end

puts "Test successful"