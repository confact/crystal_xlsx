class CrystalXlsx::Format
  property index : Int32 = 0
  property font_size : Int32 = 11
  property font_name : String = "Calibri"
  property? bold : Bool = false
  property text_color : String? = nil # Optional RGB color
  property bg_color : String? = nil   # Optional RGB color
  property? border : Bool = false
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
      bold: args[:bold]? || bold?,
      bg_color: args[:bg_color]? || bg_color,
      text_color: args[:text_color]? || text_color,
      border: args[:border]? || border?,
      horizontal_alignment: args[:horizontal_alignment]? || horizontal_alignment,
      vertical_alignment: args[:vertical_alignment]? || vertical_alignment
    )
  end

  def to_font_xml(xml)
    xml.element("font") do
      xml.element("sz", val: font_size)
      xml.element("name", val: font_name)
      xml.element("b", val: bold?)
      xml.element("color", rgb: "FF#{text_color}") if text_color
    end
  end

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
        xml.element("bgColor", indexed: "0")  # This sets bg color to auto
      end
    end
  end

  def to_border_xml(xml)
    # return nil unless border

    xml.element("border") do
      xml.element("left")
      xml.element("right")
      xml.element("top")
      xml.element("bottom")
    end
  end

  def to_xml(xml)
    xml.element("xf", numFmtId: num_form_id, fontId: font_id, fillId: (text_color || bg_color) ? font_id : 0, borderId: border? ? font_id : 0, applyFont: "1", applyFill: (text_color || bg_color) ? 1 : 0, applyBorder: border? ? 1 : 0, applyAlignment: (horizontal_alignment || vertical_alignment) ? 1 : 0) do
      if horizontal_alignment || vertical_alignment
        xml.element("alignment", horizontal: horizontal_alignment, vertical: vertical_alignment)
      end
    end
  end

  def font_id
    index
  end
end
