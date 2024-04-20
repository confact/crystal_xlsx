class CrystalXlsx::DocPropsCore
  def self.to_xml(io : IO)
    time = Time.utc
    XML.build(io, indent: "  ", encoding: "UTF-8") do |xml|
      xml.element("cp:coreProperties", "xmlns:cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties", "xmlns:dc": "http://purl.org/dc/elements/1.1/", "xmlns:dcterms": "http://purl.org/dc/terms/", "xmlns:dcmitype": "http://purl.org/dc/dcmitype/", "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance") do
        xml.element("dc:title")
        xml.element("dc:subject")
        xml.element("dc:creator") do
          xml.text("#{CrystalXlsx::NAME} #{CrystalXlsx::VERSION}")
        end
        xml.element("cp:keywords")
        xml.element("dc:description")
        xml.element("cp:lastModifiedBy") do
          xml.text("#{CrystalXlsx::NAME} #{CrystalXlsx::VERSION}")
        end
        xml.element("dcterms:created", "xsi:type": "dcterms:W3CDTF") do
          xml.text(Time::Format::ISO_8601_DATE_TIME.format(time))
        end
        xml.element("dcterms:modified", "xsi:type": "dcterms:W3CDTF") do
          xml.text(Time::Format::ISO_8601_DATE_TIME.format(time))
        end
        xml.element("cp:category")
      end
    end
  end
end