class CrystalXlsx::DocPropsApp
  def self.to_xml(worksheets, io : IO)
    XML.build(io, indent: "  ", encoding: "UTF-8") do |xml|
      xml.element("Properties", xmlns: "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", "xmlns:vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes") do
        xml.element("Application") do
          xml.text("Microsoft Excel")
        end
        xml.element("AppVersion") do
          xml.text("16.0300")
        end
        xml.element("DocSecurity") do
          xml.text("0")
        end
        xml.element("ScaleCrop") do
          xml.text("false")
        end
        xml.element("HeadingPairs") do
          xml.element("vt:vector", size: "2", "baseType": "variant") do
            xml.element("vt:variant") do
              xml.element("vt:lpstr") do
                xml.text("Worksheets")
              end
            end
            xml.element("vt:variant") do
              xml.element("vt:i4") do
                xml.text(worksheets.size.to_s)
              end
            end
          end
        end
        xml.element("TitlesOfParts") do
          xml.element("vt:vector", size: worksheets.size, "baseType": "lpstr") do
            worksheets.each do |worksheet|
              xml.element("vt:lpstr") do
                xml.text(worksheet.name)
              end
            end
          end
        end
        xml.element("Manager")
        xml.element("Company")
        xml.element("LinksUpToDate") do
          xml.text("false")
        end
        xml.element("SharedDoc") do
          xml.text("false")
        end
        xml.element("HyperlinksChanged") do
          xml.text("false")
        end
      end
    end
  end
end
