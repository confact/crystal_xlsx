# Represents cell formatting in Excel.
class CrystalXlsx::Format
  # Properties
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

  # Merges this format with new options.
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

  # Generates the font XML.
  def to_font_xml(xml)
    xml.element("font") do
      xml.element("sz", val: font_size)
      xml.element("name", val: font_name)
      xml.element("b", val: bold)
      xml.element("color", rgb: "FF#{text_color}") if text_color
    end
  end

  # Generates the fill XML.
  def to_fill_xml(xml)
    unless text_color || bg_color
      xml.element("fill") do
        xml.element("patternFill", patternType: "none")
      end
      return
    end

    xml.element("fill") do
      xml.element("patternFill", patternType: "solid") do
        xml.element("fgColor", rgb: "FF#{bg_color}") if bg_color
        xml.element("bgColor", indexed: "0")
      end
    end
  end

  # Generates the border XML.
  def to_border_xml(xml)
    xml.element("border") do
      xml.element("left")
      xml.element("right")
      xml.element("top")
      xml.element("bottom")
    end
  end

  # Generates the format XML.
  def to_xml(xml)
    has_fill = text_color || bg_color
    has_border = border
    has_alignment = horizontal_alignment || vertical_alignment
    
    xml.element("xf") do
      xml.attribute("numFmtId", num_form_id)
      xml.attribute("fontId", font_id)
      xml.attribute("fillId", has_fill ? font_id : 0)
      xml.attribute("borderId", has_border ? font_id : 0)
      xml.attribute("applyFont", "1")
      xml.attribute("applyFill", has_fill ? 1 : 0)
      xml.attribute("applyBorder", has_border ? 1 : 0)
      xml.attribute("applyAlignment", has_alignment ? 1 : 0)
      
      if has_alignment
        xml.element("alignment") do
          xml.attribute("horizontal", horizontal_alignment) if horizontal_alignment
          xml.attribute("vertical", vertical_alignment) if vertical_alignment
        end
      end
    end
  end

  # Boolean property accessors
  def bold? : Bool
    bold
  end

  def border? : Bool
    border
  end

  # Returns the font ID (same as format index).
  def font_id
    index
  end
end