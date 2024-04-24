require "./spec_helper" # Adjust the path as necessary

describe CrystalXlsx::Row do
  describe "#initialize" do
    it "initializes with values and optional format" do
      format = CrystalXlsx::Format.new(font_size: 14)
      worksheet = CrystalXlsx::Worksheet.new("Sheet1")
      row = CrystalXlsx::Row.new(1, worksheet, format)
      row.add([1, "Test", true, 3.14])
      row.size.should eq 4
      row.format.should_not be_nil
      row.format.not_nil!.font_size.should eq 14
    end

    it "initializes without a format" do
      worksheet = CrystalXlsx::Worksheet.new("Sheet1")
      row = CrystalXlsx::Row.new(1, worksheet)
      row.add([2.5_f32, 100, "Value"])
      row.size.should eq 3
      row.format.should be_nil
    end
  end

  describe "#to_xml" do
    it "generates correct XML output for mixed types" do
      worksheet = CrystalXlsx::Worksheet.new("Sheet1")
      format = CrystalXlsx::Format.new(font_size: 12)
      format.index = 1
      row = CrystalXlsx::Row.new(5, worksheet, format)
      row.add([123, "Hello", false, 78.9])
      xml = XML.build_fragment do |xml|
        row.to_xml(xml)
      end
      xml.to_s.should match(/<row r="5" spans="1:4" ht="15" x14ac:dyDescent="0.2"><c r="A5" t="n" s="1"><v>123<\/v><\/c><c r="B5" t="inlineStr" s="1"><is><t>Hello<\/t><\/is><\/c><c r="C5" t="b" s="1"><v>0<\/v><\/c><c r="D5" t="n"><v>78.9<\/v><\/c><\/row>/)
    end
  end
end
