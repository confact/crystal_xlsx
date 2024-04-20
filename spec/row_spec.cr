require "./spec_helper"  # Adjust the path as necessary

describe CrystalXlsx::Row do
  describe "#initialize" do
    it "initializes with values and optional format" do
      format = CrystalXlsx::Format.new(font_size: 14)
      row = CrystalXlsx::Row.new([1, "Test", true, 3.14], format)
      row.values.size.should eq 4
      row.format.should_not be_nil
      row.format.not_nil!.font_size.should eq 14
    end

    it "initializes without a format" do
      row = CrystalXlsx::Row.new([2.5_f32, 100, "Value"])
      row.values.size.should eq 3
      row.format.should be_nil
    end
  end

  describe "#to_xml" do
    it "generates correct XML output for mixed types" do
      format = CrystalXlsx::Format.new(font_size: 12)
      format.index = 1
      row = CrystalXlsx::Row.new([123, "Hello", false, 78.9], format)
      row.number = 5
      xml = XML.build_fragment do |xml|
        row.to_xml(xml)
      end
      xml.to_s.should match(/<row r="5"><c r="A5" s="1"><v>123<\/v><\/c><c r="B5" s="1"><is><t>Hello<\/t><\/is><\/c><c r="C5" s="1"><v>0<\/v><\/c><c r="D5" s="1"><v>78.9<\/v><\/c><\/row>/)
    end
  end
end
