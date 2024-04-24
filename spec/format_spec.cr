require "./spec_helper" # Adjust path as necessary to include necessary configurations and libraries

describe CrystalXlsx::Format do
  describe "#initialize" do
    it "initializes with default values" do
      format = CrystalXlsx::Format.new
      format.font_size.should eq 11
      format.font_name.should eq "Calibri"
      format.bold.should be_false
      format.fill_color.should be_nil
      format.bg_color.should be_nil
      format.border.should be_false
    end

    it "initializes with custom values" do
      format = CrystalXlsx::Format.new(font_size: 14, font_name: "Arial", bold: true, fill_color: "FFFFFF", bg_color: "000000", border: true)
      format.font_size.should eq 14
      format.font_name.should eq "Arial"
      format.bold.should be_true
      format.fill_color.should eq "FFFFFF"
      format.bg_color.should eq "000000"
      format.border.should be_true
    end
  end

  describe "#to_font_xml" do
    it "generates correct font XML" do
      format = CrystalXlsx::Format.new(bold: true, fill_color: "FF0000")
      xml = XML.build_fragment do |xml|
        format.to_font_xml(xml)
      end
      xml.should match(/<font><sz val="11"\/><name val="Calibri"\/><b val="true"\/><color rgb="FF0000"\/><\/font>/)
    end
  end

  describe "#to_fill_xml" do
    it "generates correct fill XML when colors are specified" do
      format = CrystalXlsx::Format.new(fill_color: "FF0000", bg_color: "00FF00")
      xml = XML.build_fragment do |xml|
        format.to_fill_xml(xml)
      end
      xml.should match(/<fill><patternFill patternType="solid"><fgColor rgb="FF0000"\/><bgColor rgb="00FF00"\/><\/patternFill><\/fill>/)
    end

    it "returns nil when no colors are specified" do
      format = CrystalXlsx::Format.new
      io = IO::Memory.new
      xml = XML::Builder.new(io)
      (format.to_fill_xml(xml) == nil).should be_true
    end
  end

  describe "#to_border_xml" do
    it "generates border XML if border is true" do
      format = CrystalXlsx::Format.new(border: true)
      xml = XML.build_fragment do |xml|
        format.to_border_xml(xml)
      end
      xml.should match(/<border><left\/><right\/><top\/><bottom\/><\/border>/)
    end

    it "returns nil if border is false" do
      io = IO::Memory.new
      xml = XML::Builder.new(io)
      format = CrystalXlsx::Format.new
      (format.to_border_xml(xml) == nil).should be_true
    end
  end

  describe "#to_xml" do
    it "generates the correct overall XML" do
      format = CrystalXlsx::Format.new(fill_color: "FFFFFF", border: true)
      format.index = 2
      xml = XML.build_fragment do |xml|
        format.to_xml(xml)
      end
      xml.should match(/<xf numFmtId="0" fontId="2" fillId="2" borderId="2" applyFont="1" applyFill="1" applyBorder="1"\/>/)
    end
  end
end
