require "./spec_helper"

describe CrystalXlsx::Worksheet do
  describe "merged cells" do
    it "should add a merged cell range by string" do
      ws = CrystalXlsx::Worksheet.new("Sheet1")
      ws.merge_cells("A1:B2")
      ws.merged_cells.should eq(["A1:B2"])
    end

    it "should add a merged cell range by from/to" do
      ws = CrystalXlsx::Worksheet.new("Sheet1")
      ws.merge_cells("A1", "C3")
      ws.merged_cells.should eq(["A1:C3"])
    end

    it "should output merged cells in XML" do
      ws = CrystalXlsx::Worksheet.new("Sheet1")
      ws.merge_cells("A1:B2")
      io = IO::Memory.new
      ws.to_xml(io)
      xml = io.to_s
      xml.should match(/<mergeCells count="1">\s*<mergeCell ref="A1:B2"\/>\s*<\/mergeCells>/)
    end
  end
end