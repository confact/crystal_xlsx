require "./spec_helper" # Adjust the path as necessary

describe CrystalXlsx::Cell do
  describe "initialize" do
    it "should create a new Cell object" do
      worksheet = CrystalXlsx::Worksheet.new("Sheet1")
      row = CrystalXlsx::Row.new(1, worksheet)
      cell = CrystalXlsx::Cell.new("value", row)
      cell.should be_a CrystalXlsx::Cell
      cell.value.should eq "value"
      cell.row.should eq row
      cell.format.should eq nil
    end
  end

  describe "to_xml" do
    it "should return the xml representation of the cell" do
      worksheet = CrystalXlsx::Worksheet.new("Sheet1")
      row = CrystalXlsx::Row.new(1, worksheet)
      cell = CrystalXlsx::Cell.new("value", row)
      xml = XML.build_fragment do |xml|
        cell.to_xml(xml)
      end
      xml.to_s.should match(/<c r="A1" t="inlineStr"><is><t>value<\/t><\/is><\/c>/)
    end

    it "should return the xml representation of the cell with a number format" do
      worksheet = CrystalXlsx::Worksheet.new("Sheet1")
      row = CrystalXlsx::Row.new(1, worksheet)
      cell = CrystalXlsx::Cell.new(2, row)
      xml = XML.build_fragment do |xml|
        cell.to_xml(xml)
      end
      xml.to_s.should match(/<c r="A1" t="n"><v>2<\/v><\/c>/)
    end

    it "should return the xml representation of the cell with a boolean format" do
      worksheet = CrystalXlsx::Worksheet.new("Sheet1")
      row = CrystalXlsx::Row.new(1, worksheet)
      cell = CrystalXlsx::Cell.new(true, row)
      xml = XML.build_fragment do |xml|
        cell.to_xml(xml)
      end
      xml.to_s.should match(/<c r="A1" t="b"><v>1<\/v><\/c>/)
    end
  end
end
