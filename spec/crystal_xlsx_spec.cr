require "./spec_helper"

describe CrystalXlsx do
  describe "NAME" do
    it "should return the name of the gem" do
      CrystalXlsx::NAME.should eq "Crystal.xlsx"
    end
  end

  describe "EPOCH" do
    it "should return the epoch of the gem" do
      CrystalXlsx::EPOCH.should eq Time.utc(1899, 12, 30).to_unix_f
    end
  end

  describe "DAY_IN_SECONDS" do
    it "should return the number of seconds in a day" do
      CrystalXlsx::DAY_IN_SECONDS.should eq 86400
    end
  end
end
