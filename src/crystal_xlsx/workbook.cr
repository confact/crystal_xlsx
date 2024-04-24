require "compress/zip"

class CrystalXlsx::Workbook
  property worksheets : Array(CrystalXlsx::Worksheet) = [] of CrystalXlsx::Worksheet
  property shared_strings : CrystalXlsx::SharedStrings = SharedStrings.new
  property theme : CrystalXlsx::Theme = CrystalXlsx::Theme.new
  property style : CrystalXlsx::Style = CrystalXlsx::Style.new
  property? enable_shared_strings : Bool = true

  def add_worksheet(name)
    worksheet = CrystalXlsx::Worksheet.new(name, workbook: self)
    @worksheets << worksheet
    worksheet
  end

  def add_worksheet(name, &)
    worksheet = CrystalXlsx::Worksheet.new(name, workbook: self)
    @worksheets << worksheet
    yield worksheet
  end

  def add_format(**options)
    style.add_format(**options)
  end

  def add_format(format : CrystalXlsx::Format)
    style.add_format(format)
  end

  def close(filepath = "./temp.xlsx")
    File.open(filepath, "w") do |file|
      to_io(file)
    end
  end

  def save(filepath = "./temp.xlsx")
    close(filepath)
  end

  def read : IO
    stream = IO::Memory.new
    to_io(stream)
    stream.rewind
    stream
  end

  def read_string : String
    read.to_s
  end

  def to_io(io)
    Compress::Zip::Writer.open(io) do |zip|
      build_zip_contents(zip)
    end
  end

  private def build_zip_contents(zip)
    zip.add("[Content_Types].xml") do |io|
      generate_content_type_xml(io)
    end
    zip.add("docProps/app.xml") do |io|
      CrystalXlsx::DocPropsApp.to_xml(worksheets, io)
    end
    zip.add("docProps/core.xml") do |io|
      CrystalXlsx::DocPropsCore.to_xml(io)
    end
    zip.add("xl/workbook.xml") do |io|
      create_workbook_xml(io)
    end
    zip.add("xl/styles.xml") do |io|
      style.to_xml(io)
    end
    zip.add("xl/theme/theme1.xml") do |io|
      theme.to_xml(io)
    end
    zip.add("_rels/.rels") do |io|
      generate_root_rels_xml(io)
    end
    zip.add("xl/_rels/workbook.xml.rels") do |io|
      generate_workbook_rels_xml(io)
    end

    zip.add("xl/sharedStrings.xml") do |io|
      shared_strings.to_xml(io)
    end if enable_shared_strings?

    worksheets.each_with_index do |worksheet, index|
      zip.add("xl/worksheets/sheet#{index + 1}.xml") do |io|
        worksheet.to_xml(io)
      end
    end
  end

  private def create_workbook_xml(io : IO)
    XML.build(io, indent: "  ", encoding: "UTF-8") do |xml|
      xml.element("workbook", xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "xmlns:r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "xmlns:mc": "http://schemas.openxmlformats.org/markup-compatibility/2006", "mc:Ignorable": "x15", "xmlns:x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main") do
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
    end
  end

  private def generate_content_type_xml(io)
    XML.build(io, indent: "  ", encoding: "UTF-8") do |xml|
      xml.element("Types", xmlns: "http://schemas.openxmlformats.org/package/2006/content-types") do
        xml.element("Default", "Extension": "rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml")
        xml.element("Default", "Extension": "xml", "ContentType": "application/xml")
        xml.element("Override", "PartName": "/xl/workbook.xml", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")

        worksheets.each_with_index do |_, index|
          xml.element("Override", "PartName": "/xl/worksheets/sheet#{index + 1}.xml", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
        end

        xml.element("Override", "PartName": "/xl/theme/theme1.xml", "ContentType": "application/vnd.openxmlformats-officedocument.theme+xml")
        xml.element("Override", "PartName": "/xl/styles.xml", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
        xml.element("Override", "PartName": "/xl/sharedStrings.xml", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml") if enable_shared_strings?
        xml.element("Override", "PartName": "/docProps/core.xml", "ContentType": "application/vnd.openxmlformats-package.core-properties+xml")
        xml.element("Override", "PartName": "/docProps/app.xml", "ContentType": "application/vnd.openxmlformats-officedocument.extended-properties+xml")
      end
    end
  end

  private def generate_root_rels_xml(io : IO)
    XML.build(io, indent: "  ", encoding: "UTF-8") do |xml|
      xml.element("Relationships", "xmlns": "http://schemas.openxmlformats.org/package/2006/relationships") do
        xml.element("Relationship", Id: "rId1", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", Target: "xl/workbook.xml")
        xml.element("Relationship", Id: "rId2", Type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", Target: "docProps/core.xml")
        xml.element("Relationship", Id: "rId3", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", Target: "docProps/app.xml")
      end
    end
  end

  private def generate_workbook_rels_xml(io : IO)
    XML.build(io, indent: "  ", encoding: "UTF-8") do |xml|
      xml.element("Relationships", "xmlns": "http://schemas.openxmlformats.org/package/2006/relationships") do
        # Relationships to sheets
        worksheets.each_with_index do |_, index|
          xml.element("Relationship", Id: "rId#{index + 1}", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", Target: "worksheets/sheet#{index + 1}.xml")
        end

        # Relationship to shared strings (if used)
        xml.element("Relationship", Id: "rId#{worksheets.size + 1}", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings", Target: "sharedStrings.xml") if enable_shared_strings?

        # Relationship to styles (if used)
        xml.element("Relationship", Id: "rId#{worksheets.size + 2}", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", Target: "styles.xml")
        xml.element("Relationship", Id: "rId#{worksheets.size + 3}", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", Target: "theme/theme1.xml")
      end
    end
  end
end
