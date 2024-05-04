class CrystalXlsx::Pane
  property x_split : Int32 = 0
  property y_split : Int32 = 0
  property top_left_cell : String = "A1"
  property active_pane : String = "bottomLeft"
  property state : String = "frozen"

  def initialize(x_split, y_split, top_left_cell, active_pane, state)
    @x_split = x_split
    @y_split = y_split
    @top_left_cell = top_left_cell
    @active_pane = active_pane
    @state = state
  end

  def to_xml(xml)
    xml.element("pane") do
      xml.attribute("xSplit", x_split)
      xml.attribute("ySplit", y_split)
      xml.attribute("topLeftCell", top_left_cell)
      xml.attribute("activePane", active_pane)
      xml.attribute("state", state)
    end
  end
end
