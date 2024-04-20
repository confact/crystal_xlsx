class CrystalXlsx::Format
  property index : Int32 = 0
  property font_size : Int32 = 11
  property font_name : String = "Calibri"
  property bold : Bool = false
  property fill_color : String? = nil  # Optional RGB color
  property bg_color : String? = nil  # Optional RGB color
  property border : Bool = false
  property num_form_id : Int32 = 165

  def initialize(@num_form_id = 165, @font_size : Int32 = 11, @font_name : String = "Calibri", @bold : Bool = false, @bg_color : String? = nil, @fill_color : String? = nil, @border : Bool = false)
  end

  def to_font_xml(xml)
    xml.element("font") do
      xml.element("sz", val: font_size)
      xml.element("name", val: font_name)
      xml.element("b", val: bold ? 1 : 0)
      xml.element("color", rgb: fill_color) if fill_color
    end
  end

  def to_fill_xml(xml)
    return nil unless fill_color || bg_color

    xml.element("fill") do
      xml.element("patternFill", patternType: "solid") do
        xml.element("fgColor", rgb: fill_color) if fill_color
        xml.element("bgColor", rgb: bg_color) if bg_color
      end
    end
  end

  def to_border_xml(xml)
    return nil unless border

    xml.element("border") do
      xml.element("left")
      xml.element("right")
      xml.element("top")
      xml.element("bottom")
    end
  end

  def to_xml(xml)
    xml.element("xf", numFmtId: num_form_id, fontId: index, fillId: (fill_color || bg_color) ? index : 0, borderId: "#{border ? index : 0}", applyFont: "1", applyFill: (fill_color || bg_color) ? 1 : 0, applyBorder: border ? 1 : 0)
  end
end