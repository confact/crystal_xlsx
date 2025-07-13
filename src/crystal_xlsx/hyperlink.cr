class CrystalXlsx::Hyperlink
  property cell_ref : String
  property url : String?
  property display_text : String?
  property target_worksheet : CrystalXlsx::Worksheet?
  property target_cell : String?

  def initialize(@cell_ref : String, @url : String, @display_text : String? = nil)
  end

  def initialize(@cell_ref : String, @target_worksheet : CrystalXlsx::Worksheet, @target_cell : String, @display_text : String? = nil)
  end

  # Create a hyperlink to another worksheet
  def self.new_worksheet_link(cell_ref : String, target_worksheet : CrystalXlsx::Worksheet, target_cell : String, display_text : String? = nil)
    new(cell_ref, target_worksheet, target_cell, display_text)
  end

  def to_xml(xml, relationship_id : Int32)
    if target_worksheet
      # Internal worksheet link
      xml.element("hyperlink", ref: cell_ref, "r:id": "rId#{relationship_id}") do
        xml.element("display") do
          xml.text(display_text || "#{target_worksheet.try(&.name) || "Sheet"}!#{target_cell}")
        end
      end
    else
      # External URL link
      xml.element("hyperlink", ref: cell_ref, "r:id": "rId#{relationship_id}") do
        xml.element("display") do
          xml.text(display_text || url)
        end
      end
    end
  end

  def relationship_type : String
    if target_worksheet
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    else
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    end
  end

  def relationship_target : String
    if target_worksheet
      "worksheets/#{target_worksheet.name}.xml"
    else
      url || ""
    end
  end
end