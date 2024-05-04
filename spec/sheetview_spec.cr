require "./spec_helper" # Adjust the path as necessary

describe CrystalXlsx::Sheetview do
  it "should create a sheetview" do
    sheetview = CrystalXlsx::Sheetview.new
    sheetview.tab_selected?.should eq(false)
    sheetview.workbook_view_id.should eq(0)
  end

  it "should add pane" do
    sheetview = CrystalXlsx::Sheetview.new
    sheetview.add_pane(0, 0, "A1", "bottomLeft", "frozen")
    pane = sheetview.pane

    pane.should_not be_nil

    pane.try(&.active_pane).should eq("bottomLeft")
    pane.try(&.state).should eq("frozen")
    pane.try(&.top_left_cell).should eq("A1")
    pane.try(&.x_split).should eq(0)
    pane.try(&.y_split).should eq(0)
  end
end