class CrystalXlsx::Formula::Link < CrystalXlsx::Formula
  property link : String? = nil
  property display : String? = nil
  property worksheet : CrystalXlsx::Worksheet? = nil
  property cell_ref : String? = nil

  def initialize(@link : String, @display : String)
  end

  def initialize(@link : String, @display : String, @worksheet : CrystalXlsx::Worksheet, @cell_ref : String)
  end

  # create a link to a worksheet's cell
  # @param worksheet [CrystalXlsx::Worksheet] the worksheet to link to
  # @param cell_ref [String] the cell reference to link to
  # @param display [String] the text to display for the link
  # @return [CrystalXlsx::Link] the link object
  def self.sheet_link(worksheet : CrystalXlsx::Worksheet, cell_ref : String, display : String? = nil) : CrystalXlsx::Link
    link = "##{worksheet.name}!#{cell_ref}"
    display = display || worksheet.name
    new link, display, worksheet, cell_ref
  end

  def value(sheet : CrystalXlsx::Worksheet) : String
    if links_to_sheet?
      sheet_name = worksheet.try(&.name) || display
      cell_ref = self.cell_ref
      return "'#{sheet_name}'!#{cell_ref}"
    else
      display || link || ""
    end
  end

  def to_s(io : IO)
    io << "HYPERLINK("
    io << link
    io << ","
    io << display
    io << ")"
  end

  def links_to_sheet? : Bool
    worksheet != nil
  end

  def links_to_url? : Bool
    link.start_with?("http")
  end
end

