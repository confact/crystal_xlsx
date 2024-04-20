class CrystalXlsx::Style
  property formats : Array(CrystalXlsx::Format) = [] of CrystalXlsx::Format

  def to_xml(io : IO)
    XML.build(io, indent: "  ", encoding: "UTF-8") do |xml|
      xml.element("styleSheet", xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main") do
        # Number formats
        xml.element("numFmts", count: "1") do
          xml.element("numFmt", numFmtId: "164", formatCode: "yyyy-mm-dd hh:mm:ss")
        end
        
        # Fonts
        xml.element("fonts", count: "#{formats.size + 1}") do
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
        xml.element("fills", count: "#{formats.size + 1}") do
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
        xml.element("borders", count: "#{formats.size + 1}") do
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
        xml.element("cellStyleXfs", count: "#{formats.size + 2}") do
          xml.element("xf", numFmtId: "0", fontId: "0", fillId: "0", borderId: "0", applyNumberFormat: "1")
          xml.element("xf", numFmtId: "22", fontId: "0", fillId: "0", borderId: "0", applyNumberFormat: "1")

          # Custom styles
          formats.each do |format|
            format.to_xml(xml)
          end
        end
        
        # Cell XFs
        xml.element("cellXfs", count: "#{formats.size + 2}") do
          # Default style
          xml.element("xf", numFmtId: "0", fontId: "0", fillId: "0", borderId: "0", applyNumberFormat: "1")
          xml.element("xf", numFmtId: "22", fontId: "0", fillId: "0", borderId: "0", applyNumberFormat: "1")
          
          # Custom styles
          formats.each do |format|
            format.to_xml(xml)
          end
        end
      end
    end
  end 
end