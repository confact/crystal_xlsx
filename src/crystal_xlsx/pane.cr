class CrystalXlsx::Pane
  property xSplit : Int32 = 0
  property ySplit : Int32 = 0
  property topLeftCell : String = "A1"
  property activePane : String = "bottomLeft"
  property state : String = "frozen"

  def initialize(xSplit, ySplit, topLeftCell, activePane, state)
    @xSplit = xSplit
    @ySplit = ySplit
    @topLeftCell = topLeftCell
    @activePane = activePane
    @state = state
  end

  def to_xml(xml)
    xml.element("pane") do
      xml.attribute("xSplit", @xSplit)
      xml.attribute("ySplit", @ySplit)
      xml.attribute("topLeftCell", @topLeftCell)
      xml.attribute("activePane", @activePane)
      xml.attribute("state", @state)
    end
  end
end