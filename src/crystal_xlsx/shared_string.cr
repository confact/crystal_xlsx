class CrystalXlsx::SharedStrings
  getter strings : Set(String) = Set(String).new

  def add(string : String) : Int32
    strings << string
    strings.size - 1
  end

  def add(strings)
    strings.each do |string|
      add(string) if string.is_a?(String)
    end
  end

  def to_xml(io : IO)
    XML.build(io, indent: "  ", encoding: "UTF-8") do |xml|
      xml.element("sst", xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main", count: "#{strings.size}", uniqueCount: "#{strings.size}") do
        strings.each do |string|
          xml.element("si") do
            xml.element("t") do
              xml.text(string)
            end
          end
        end
      end
    end
  end
end