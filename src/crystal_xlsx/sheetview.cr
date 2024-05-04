class CrystalXlsx::Sheetview
  property? tab_selected : Bool = false
  property workbook_view_id : Int32 = 0
  property pane : CrystalXlsx::Pane?

  def initialize(@tab_selected = false, @workbook_view_id = 0)
  end

  # Add a pane to the sheetview
  # @param xsplit [Integer] The number of columns in the top pane
  # @param ysplit [Integer] The number of rows in the left pane
  # @param top_left_cell [String] The cell in the top left of the bottom right pane
  # @param active_pane [String] The pane that is active
  # @param state [String] The state of the pane
  # @return [void]
  def add_pane(xsplit, ysplit, top_left_cell, active_pane, state)
    @pane = CrystalXlsx::Pane.new(xsplit, ysplit, top_left_cell, active_pane, state)
  end

  # Convert the sheetview to an XML representation
  def to_xml(xml)
    xml.element("sheetViews") do
      xml.element("sheetView") do
        xml.attribute("tabSelected", @tab_selected ? 1 : 0)
        xml.attribute("workbookViewId", @workbook_view_id)
        @pane.try(&.to_xml(xml))
      end
    end
  end
end
