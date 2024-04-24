class CrystalXlsx::SharedStrings
  property strings : Hash(String, Int32) = Hash(String, Int32).new
  property count : Int32 = 0

  def add(string : String) : Int32
    if @strings.has_key?(string)
      @strings[string] += 1
    else
      @strings[string] = 1
    end
    @count += 1
    index(string)
  end

  def add_data(new_strings)
    new_strings.each do |string|
      add(string) if string.is_a?(String)
    end
  end

  def index(string : String) : Int32
    @strings.keys.index!(string)
  end

  def size
    @strings.size
  end

  def to_xml(io : IO)
    XML.build(io, indent: "  ", encoding: "UTF-8") do |xml|
      xml.element("sst", xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main", count: count, uniqueCount: size) do
        strings.each_key do |string|
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
