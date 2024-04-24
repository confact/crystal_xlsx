class CrystalXlsx::Theme
  def to_xml(io : IO)
    XML.build(io, indent: "  ", encoding: "UTF-8") do |xml|
      xml.element("a:theme", "xmlns:a": "http://schemas.openxmlformats.org/drawingml/2006/main", "name": "Office Theme") do
        build_theme_elements(xml)
      end
    end
  end

  private def build_theme_elements(xml)
    xml.element("a:themeElements") do
      build_color_scheme(xml, "Office")
      build_font_scheme(xml, "Office")
    end
  end

  private def build_color_scheme(xml, name : String)
    xml.element("a:clrScheme", "name": name) do
      define_color(xml, "a:dk1", "a:sysClr", "windowText", "000000")
      define_color(xml, "a:lt1", "a:sysClr", "window", "FFFFFF")
      define_simple_colors(xml)
    end
  end

  private def define_color(xml, element_name : String, color_type : String, value : String, last_color : String)
    xml.element(element_name) do
      xml.element(color_type, "val": value, "lastClr": last_color)
    end
  end

  private def define_simple_colors(xml)
    simple_colors = {
      "a:dk2"      => "44546A",
      "a:lt2"      => "E7E6E6",
      "a:accent1"  => "4472C4",
      "a:accent2"  => "ED7D31",
      "a:accent3"  => "A5A5A5",
      "a:accent4"  => "FFC000",
      "a:accent5"  => "5B9BD5",
      "a:accent6"  => "70AD47",
      "a:hlink"    => "0563C1",
      "a:folHlink" => "954F72",
    }
    simple_colors.each do |tag, color|
      xml.element(tag) do
        xml.element("a:srgbClr", "val": color)
      end
    end
  end

  private def build_font_scheme(xml, name : String)
    xml.element("a:fontScheme", "name": name) do
      build_major_font(xml)
    end
  end

  private def build_major_font(xml)
    xml.element("a:majorFont") do
      # Latin font and other regional fonts
      xml.element("a:latin", "typeface": "Calibri Light")
      xml.element("a:ea", "typeface": "")
      xml.element("a:cs", "typeface": "")
      regional_fonts.each do |script, typeface|
        xml.element("a:font", "script": script, "typeface": typeface)
      end
    end
  end

  private def regional_fonts
    {
      "Jpan" => "游ゴシック Light", "Hang" => "맑은 고딕", "Hans" => "等线 Light",
      "Hant" => "新細明體", "Arab" => "Times New Roman", "Hebr" => "Times New Roman",
      "Thai" => "Tahoma", "Ethi" => "Nyala", "Beng" => "Vrinda", "Gujr" => "Shruti",
      "Khmr" => "MoolBoran", "Knda" => "Tunga", "Guru" => "Raavi", "Cans" => "Euphemia",
      "Cher" => "Plantagenet Cherokee", "Yiii" => "Microsoft Yi Baiti", "Tibt" => "Microsoft Himalaya",
      "Thaa" => "MV Boli", "Deva" => "Mangal", "Telu" => "Gautami", "Taml" => "Latha",
      "Syrc" => "Estrangelo Edessa", "Orya" => "Kalinga", "Mlym" => "Kartika",
      "Laoo" => "DokChampa", "Sinh" => "Iskoola Pota", "Mong" => "Mongolian Baiti",
      "Viet" => "Times New Roman", "Uigh" => "Microsoft Uighur", "Geor" => "Sylfaen",
    }
  end
end
