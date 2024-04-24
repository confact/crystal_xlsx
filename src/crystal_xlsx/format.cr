class CrystalXlsx::Format
  property index : Int32 = 0
  property font_size : Int32 = 11
  property font_name : String = "Calibri"
  property? bold : Bool = false
  property fill_color : String? = nil # Optional RGB color
  property bg_color : String? = nil   # Optional RGB color
  property? border : Bool = false
  property num_form_id : Int32 = 0

  def initialize(@num_form_id = 0, @font_size : Int32 = 11, @font_name : String = "Calibri", @bold : Bool = false, @bg_color : String? = nil, @fill_color : String? = nil, @border : Bool = false)
  end

  def merge(**args)
    CrystalXlsx::Format.new(
      num_form_id: args[:num_form_id]? || num_form_id,
      font_size: args[:font_size]? || font_size,
      font_name: args[:font_name]? || font_name,
      bold: args[:bold]? || bold?,
      bg_color: args[:bg_color]? || bg_color,
      fill_color: args[:fill_color]? || fill_color,
      border: args[:border]? || border?
    )
  end

  def to_font_xml(xml)
    xml.element("font") do
      xml.element("sz", val: font_size)
      xml.element("name", val: font_name)
      xml.element("b", val: bold?)
      xml.element("color", rgb: fill_color) if fill_color
    end
  end

  def to_fill_xml(xml)
    unless fill_color || bg_color
      xml.element("fill") do
        xml.element("patternFill", patternType: "none")
      end
      return
    end

    xml.element("fill") do
      xml.element("patternFill", patternType: "solid") do
        xml.element("fgColor", rgb: fill_color) if fill_color
        xml.element("bgColor", rgb: bg_color) if bg_color
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
    xml.element("xf", numFmtId: num_form_id, fontId: font_id, fillId: (fill_color || bg_color) ? font_id : 0, borderId: border? ? font_id : 0, applyFont: "1", applyFill: (fill_color || bg_color) ? 1 : 0, applyBorder: border? ? 1 : 0)
  end

  def font_id
    index
  end
end
