class CrystalXlsx::Rels
  property rels : Array(NamedTuple(id: String, type: String, target: String))

  def initialize
    @rels = [] of NamedTuple(id: String, type: String, target: String)
  end

  def add_rel(id, type, target)
    @rels << {id: id, type: type, target: target}
  end

  def <<(rel)
    @rels << rel
  end

  def to_xml(io : IO)
    XML.build(io, indent: "  ", encoding: "UTF-8") do |xml|
      xml.element("Relationships", xmlns: "http://schemas.openxmlformats.org/package/2006/relationships") do
        @rels.each do |rel|
          xml.element("Relationship", Id: rel[:id], Type: rel[:type], Target: rel[:target])
        end
      end
    end
  end
end
