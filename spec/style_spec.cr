require "./spec_helper" # Adjust path as necessary to include necessary configurations and libraries

describe CrystalXlsx::Style do
  it "should be able to create a new style" do
    style = CrystalXlsx::Style.new
    style.should be_a(CrystalXlsx::Style)
  end

  it "add format" do
    style = CrystalXlsx::Style.new
    style.add_format(font_size: 12, bold: true)
    style.formats.first.font_size.should eq(12)
    style.formats.first.bold?.should eq(true)
  end

  describe "#to_xml" do
    it "should return a string" do
      style = CrystalXlsx::Style.new
      style.add_format(font_size: 12, bold: true)
      io = IO::Memory.new
      style.to_xml(io)
      xml = io.to_s
      xml.should be_a(String)
    end
  end
end
