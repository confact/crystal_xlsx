class CrystalXlsx::SharedStrings
  property string_indices : Hash(String, Int32) = {} of String => Int32
  property strings : Array(String) = [] of String
  property string_set : Set(String) = Set(String).new
  property count : Int32 = 0

  def add(string : String) : Int32
    unless string_set.includes?(string)
      string_indices[string] = strings.size
      strings << string
      string_set << string
    end
    @count += 1
    string_indices[string]
  end

  def add_data(new_strings : Array(String))
    new_strings.each { |string| add(string) }
  end

  def index(string : String) : Int32
    string_indices[string]? || raise("String not found: #{string}")
  end

  def size
    strings.size
  end

  def to_xml(io : IO)
    XML.build(io, indent: "  ", encoding: "UTF-8") do |xml|
      xml.element("sst", xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main", count: count, uniqueCount: size) do
        strings.each do |string|
          xml.element("si") do
            xml.element("t") { xml.text(string) }
          end
        end
      end
    end
  end
end