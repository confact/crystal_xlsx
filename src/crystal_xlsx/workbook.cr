require "compress/zip"

# Represents an Excel workbook with multiple worksheets.
# 
# Example:
# ```
# workbook = CrystalXlsx::Workbook.new
# workbook.sheet("Data") do |sheet|
#   sheet.add(["Name", "Age"])
#   sheet.add(["Alice", 25])
# end
# workbook.save("output.xlsx")
# ```
class CrystalXlsx::Workbook
  # Properties
  getter worksheets = [] of Worksheet
  getter shared_strings = SharedStrings.new
  getter theme = Theme.new
  getter style = Style.new
  getter workbook_rels = Rels.new
  getter rels = Rels.new
  
  # Configuration
  property? enable_shared_strings = true

  # Creates a new worksheet with the given name and yields it to the block.
  # 
  # Example:
  # ```
  # workbook.sheet("Sales") do |sheet|
  #   sheet.add(["Product", "Revenue"])
  #   sheet.add(["Widget", 1000])
  # end
  # ```
  def sheet(name : String, &block)
    worksheet = Worksheet.new(name, self)
    @worksheets << worksheet
    yield worksheet
    worksheet
  end

  # Creates a new worksheet with the given name.
  # 
  # Example:
  # ```
  # sheet = workbook.sheet("Data")
  # sheet.add(["A", "B", "C"])
  # ```
  def sheet(name : String) : Worksheet
    worksheet = Worksheet.new(name, self)
    @worksheets << worksheet
    worksheet
  end

  # Creates a new style format with the given options.
  # 
  # Example:
  # ```
  # format = workbook.style(font_size: 12, bold: true, text_color: "FF0000")
  # sheet.add(["Header"], format)
  # ```
  def style(**options) : Format
    style.add_format(**options)
  end

  # Creates a new style format from an existing format.
  def style(format : Format) : Format
    style.add_format(format)
  end

  # Saves the workbook to a file.
  # 
  # Example:
  # ```
  # workbook.save("output.xlsx")
  # ```
  def save(filename : String)
    File.open(filename, "w") do |file|
      write_to(file)
    end
  end

  # Writes the workbook to an IO stream.
  def write_to(io : IO)
    Compress::Zip::Writer.open(io) do |zip|
      build_zip_contents(zip)
    end
  end

  # Returns the workbook as a string.
  def to_s : String
    io = IO::Memory.new
    write_to(io)
    io.to_s
  end

  # Returns the workbook as an IO stream.
  def to_io : IO
    stream = IO::Memory.new
    write_to(stream)
    stream.rewind
    stream
  end

  private def build_zip_contents(zip)
    # Content types
    zip.add("[Content_Types].xml") { |io| generate_content_types(io) }
    
    # Document properties
    zip.add("docProps/app.xml") { |io| DocPropsApp.to_xml(worksheets, io) }
    zip.add("docProps/core.xml") { |io| DocPropsCore.to_xml(io) }
    
    # Main workbook
    zip.add("xl/workbook.xml") { |io| create_workbook_xml(io) }
    zip.add("xl/styles.xml") { |io| style.to_xml(io) }
    zip.add("xl/theme/theme1.xml") { |io| theme.to_xml(io) }
    
    # Relationships
    zip.add("_rels/.rels") { |io| generate_root_rels(io) }
    zip.add("xl/_rels/workbook.xml.rels") { |io| generate_workbook_rels(io) }

    # Shared strings
    if enable_shared_strings?
      zip.add("xl/sharedStrings.xml") { |io| shared_strings.to_xml(io) }
    end

    # Worksheets
    worksheets.each_with_index do |worksheet, index|
      sheet_num = index + 1
      zip.add("xl/worksheets/sheet#{sheet_num}.xml") { |io| worksheet.to_xml(io) }
      
      # Worksheet relationships (for hyperlinks)
      if worksheet.hyperlinks.size > 0
        zip.add("xl/worksheets/_rels/sheet#{sheet_num}.xml.rels") { |io| generate_worksheet_rels(index, io) }
      end
    end
  end

  private def create_workbook_xml(io : IO)
    XML.build(io, indent: "  ", encoding: "UTF-8") do |xml|
      # CrystalXlsx.workbook_xml do
      xml.element("workbook", 
        xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
        "mc:Ignorable": "x15",
        "xmlns:x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
      ) do
        xml.element("fileVersion", appName: "xl", lastEdited: "4", lowestEdited: "4", rupBuild: "9302")
        xml.element("workbookPr", defaultThemeVersion: "202300")
        xml.element("bookViews") do
          xml.element("workbookView", xWindow: "0", yWindow: "0", windowWidth: "25600", windowHeight: "19020", "xr2:uid": "{00000000-000D-0000-FFFF-FFFF00000000}", "xmlns:xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
        end
        xml.element("sheets") do
          worksheets.each_with_index do |sheet, index|
            xml.element("sheet", name: sheet.name, sheetId: index + 1, "r:id": "rId#{index + 1}")
          end
        end
        xml.element("calcPr", calcId: "0")
      end
      # end
    end
  end

  private def generate_content_types(io)
    XML.build(io, indent: "  ", encoding: "UTF-8") do |xml|
      xml.element("Types", xmlns: "http://schemas.openxmlformats.org/package/2006/content-types") do
        # Defaults
        xml.element("Default", "Extension": "rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml")
        xml.element("Default", "Extension": "xml", "ContentType": "application/xml")
        
        # Overrides
        xml.element("Override", "PartName": "/xl/workbook.xml", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
        
        worksheets.each_with_index do |_, index|
          xml.element("Override", "PartName": "/xl/worksheets/sheet#{index + 1}.xml", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
        end

        xml.element("Override", "PartName": "/xl/theme/theme1.xml", "ContentType": "application/vnd.openxmlformats-officedocument.theme+xml")
        xml.element("Override", "PartName": "/xl/styles.xml", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
        xml.element("Override", "PartName": "/xl/sharedStrings.xml", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml") if enable_shared_strings?
        xml.element("Override", "PartName": "/docProps/core.xml", "ContentType": "application/vnd.openxmlformats-package.core-properties+xml")
        xml.element("Override", "PartName": "/docProps/app.xml", "ContentType": "application/vnd.openxmlformats-officedocument.extended-properties+xml")
        
        # Hyperlink content types
        worksheets.each_with_index do |worksheet, index|
          if worksheet.hyperlinks.size > 0
            xml.element("Override", "PartName": "/xl/worksheets/_rels/sheet#{index + 1}.xml.rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml")
          end
        end
      end
    end
  end

  private def generate_root_rels(io : IO)
    rels << {id: "rId1", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", target: "xl/workbook.xml"}
    rels << {id: "rId2", type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", target: "docProps/core.xml"}
    rels << {id: "rId3", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", target: "docProps/app.xml"}
    rels.to_xml(io)
  end

  private def generate_workbook_rels(io : IO)
    worksheets.each_with_index do |_, index|
      workbook_rels << {id: "rId#{index + 1}", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", target: "worksheets/sheet#{index + 1}.xml"}
    end
    workbook_rels << {id: "rId#{worksheets.size + 1}", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", target: "sharedStrings.xml"} if enable_shared_strings?
    workbook_rels << {id: "rId#{worksheets.size + 2}", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", target: "styles.xml"}
    workbook_rels << {id: "rId#{worksheets.size + 3}", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", target: "theme/theme1.xml"}
    workbook_rels.to_xml(io)
  end

  private def generate_worksheet_rels(worksheet_index : Int32, io : IO)
    worksheet = worksheets[worksheet_index]
    worksheet_rels = Rels.new
    
    worksheet.hyperlinks.each_with_index do |hyperlink, index|
      rel_id = index + 1
      worksheet_rels << {
        id: "rId#{rel_id}",
        type: hyperlink.relationship_type,
        target: hyperlink.relationship_target
      }
    end
    
    worksheet_rels.to_xml(io)
  end
end
