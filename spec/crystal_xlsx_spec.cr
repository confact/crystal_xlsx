require "./spec_helper"

describe CrystalXlsx do
  describe "VERSION" do
    it "should return the version" do
      CrystalXlsx::VERSION.should eq "0.1.0"
    end
  end

  describe "MAX_COLUMNS" do
    it "should return the maximum columns" do
      CrystalXlsx::MAX_COLUMNS.should eq 16_384
    end
  end

  describe "MAX_ROWS" do
    it "should return the maximum rows" do
      CrystalXlsx::MAX_ROWS.should eq 1_048_576
    end
  end

  describe "EPOCH" do
    it "should return the epoch" do
      CrystalXlsx::EPOCH.should eq Time.utc(1899, 12, 30).to_unix_f
    end
  end

  describe "DAY_IN_SECONDS" do
    it "should return the number of seconds in a day" do
      CrystalXlsx::DAY_IN_SECONDS.should eq 86400
    end
  end

  describe "create" do
    it "should create a workbook with a block" do
      workbook = CrystalXlsx.create do |wb|
        wb.sheet("Test") do |sheet|
          sheet.add(["Hello", "World"])
        end
      end
      workbook.should be_a(CrystalXlsx::Workbook)
      workbook.worksheets.size.should eq(1)
    end
  end
end
