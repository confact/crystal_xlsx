class CrystalXlsx::Style
  property format_hashes : Set(String) = Set(String).new
  property formats : Array(CrystalXlsx::Format) = [] of CrystalXlsx::Format

  def add_format(**args) : CrystalXlsx::Format
    format = CrystalXlsx::Format.new(**args)
    add_format(format)
  end

  def add_format(format : CrystalXlsx::Format) : CrystalXlsx::Format
    format_hash = generate_format_hash(format)
    unless format_hashes.includes?(format_hash)
      format_hashes.add(format_hash)
      format.index = format_hashes.size
      formats << format
      return format
    end
    formats.find! { |fmt| generate_format_hash(fmt) == format_hash }
  end

  private def generate_format_hash(format : CrystalXlsx::Format) : String
    String.build do |str|
      str << format.font_name
      str << format.font_size
      str << format.fill_color
      str << format.bg_color
      str << format.bold?
      str << format.border?
      str << format.num_form_id
    end
  end

  def to_xml(io : IO)
    XML.build(io, indent: "  ", encoding: "UTF-8") do |xml|
      xml.element("styleSheet", xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main") do
        # Fonts
        xml.element("fonts", count: formats.size + 1) do
          # Default font
          xml.element("font") do
            xml.element("sz", val: "11")
            xml.element("color", theme: "1")
            xml.element("name", val: "Calibri")
            xml.element("family", val: "2")
            xml.element("scheme", val: "minor")
          end

          # Custom fonts
          formats.each do |format|
            format.to_font_xml(xml)
          end
        end

        # Fills
        xml.element("fills", count: formats.size + 1) do
          # Default fills
          xml.element("fill") do
            xml.element("patternFill", patternType: "none")
          end

          # Custom fills
          formats.each do |format|
            format.to_fill_xml(xml)
          end
        end

        # Borders
        xml.element("borders", count: formats.size + 1) do
          # Default border
          xml.element("border") do
            xml.element("left")
            xml.element("right")
            xml.element("top")
            xml.element("bottom")
            xml.element("diagonal")
          end

          # Custom borders
          formats.each do |format|
            format.to_border_xml(xml)
          end
        end

        # Cell style XFs
        xml.element("cellStyleXfs", count: formats.size + 1) do
          xml.element("xf", numFmtId: "0", fontId: "0", fillId: "0", borderId: "0")

          # Custom styles
          formats.each do |format|
            format.to_xml(xml)
          end
        end

        # Cell XFs
        xml.element("cellXfs", count: formats.size + 1) do
          # Default style
          xml.element("xf", numFmtId: "0", fontId: "0", fillId: "0", borderId: "0")

          # Custom styles
          formats.each do |format|
            format.to_xml(xml)
          end
        end
      end
    end
  end
end
