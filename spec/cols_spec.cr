require "./spec_helper" # Adjust the path as necessary

describe CrystalXlsx::Cols do
  it "should have a cols attribute" do
    cols = CrystalXlsx::Cols.new
    cols.column_widths.should be_a(Hash(Int32, Float64))
  end

  it "should have a add method" do
    cols = CrystalXlsx::Cols.new
    cols.add_column_width(0, 10)
    cols.column_widths[0].should eq(10)
  end

  describe "#to_xml" do
    it "should return an empty string if there are no columns" do
      cols = CrystalXlsx::Cols.new
      xml = XML.build_fragment do |xml|
        cols.to_xml(xml)
      end
      xml.should eq("\n")
    end

    it "should return the correct XML" do
      cols = CrystalXlsx::Cols.new
      cols.add_column_width(0, 10)
      cols.add_column_width(1, 20)
      xml = XML.build_fragment do |xml|
        cols.to_xml(xml)
      end
      xml.should match(/<cols><col min="1" max="1" width="10.0" customWidth="1"\/><col min="2" max="2" width="20.0" customWidth="1"\/><\/cols>/)
    end
  end
end
