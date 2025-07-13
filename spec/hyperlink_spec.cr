require "./spec_helper"

describe CrystalXlsx::Hyperlink do
  describe "initialize" do
    it "should create a URL hyperlink" do
      hyperlink = CrystalXlsx::Hyperlink.new("A1", "https://example.com", "Click here")
      hyperlink.cell_ref.should eq("A1")
      hyperlink.url.should eq("https://example.com")
      hyperlink.display_text.should eq("Click here")
    end

    it "should create a worksheet hyperlink" do
      worksheet = CrystalXlsx::Worksheet.new("Sheet2")
      hyperlink = CrystalXlsx::Hyperlink.new("A1", worksheet, "B5", "Go to Sheet2")
      hyperlink.cell_ref.should eq("A1")
      hyperlink.target_worksheet.should eq(worksheet)
      hyperlink.target_cell.should eq("B5")
      hyperlink.display_text.should eq("Go to Sheet2")
    end
  end

  describe "to_xml" do
    it "should generate XML for URL hyperlink" do
      hyperlink = CrystalXlsx::Hyperlink.new("A1", "https://example.com", "Click here")
      xml = XML.build_fragment do |xml|
        hyperlink.to_xml(xml, 1)
      end
      xml.to_s.should match(/<hyperlink ref="A1" r:id="rId1"><display>Click here<\/display><\/hyperlink>/)
    end

    it "should generate XML for worksheet hyperlink" do
      worksheet = CrystalXlsx::Worksheet.new("Sheet2")
      hyperlink = CrystalXlsx::Hyperlink.new("A1", worksheet, "B5", "Go to Sheet2")
      xml = XML.build_fragment do |xml|
        hyperlink.to_xml(xml, 1)
      end
      xml.to_s.should match(/<hyperlink ref="A1" r:id="rId1"><display>Go to Sheet2<\/display><\/hyperlink>/)
    end
  end

  describe "relationship_type" do
    it "should return hyperlink type for URL" do
      hyperlink = CrystalXlsx::Hyperlink.new("A1", "https://example.com")
      hyperlink.relationship_type.should eq("http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink")
    end

    it "should return worksheet type for worksheet link" do
      worksheet = CrystalXlsx::Worksheet.new("Sheet2")
      hyperlink = CrystalXlsx::Hyperlink.new("A1", worksheet, "B5")
      hyperlink.relationship_type.should eq("http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
    end
  end
end

describe CrystalXlsx::Worksheet do
  describe "hyperlinks" do
    it "should add URL hyperlink" do
      worksheet = CrystalXlsx::Worksheet.new("Sheet1")
      worksheet.add_hyperlink(0, 0, "https://example.com", "Click here")
      worksheet.hyperlinks.size.should eq(1)
      worksheet.hyperlinks.first.cell_ref.should eq("A1")
      worksheet.hyperlinks.first.url.should eq("https://example.com")
    end

    it "should add worksheet hyperlink" do
      worksheet1 = CrystalXlsx::Worksheet.new("Sheet1")
      worksheet2 = CrystalXlsx::Worksheet.new("Sheet2")
      worksheet1.add_worksheet_hyperlink(0, 0, worksheet2, "B5", "Go to Sheet2")
      worksheet1.hyperlinks.size.should eq(1)
      worksheet1.hyperlinks.first.target_worksheet.should eq(worksheet2)
      worksheet1.hyperlinks.first.target_cell.should eq("B5")
    end
  end
end