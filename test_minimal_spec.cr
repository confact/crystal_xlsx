require "spec"

# Define the module and macros first
module CrystalXlsx
  VERSION = "0.1.0"

  # Macro for XML element with attributes
  macro xml_element(name, **attrs, &block)
    xml.element({{name.id.stringify}}) do
      {% for key, value in attrs %}
        xml.attribute({{key.stringify}}, {{value}})
      {% end %}
      {{block}}
    end
  end
end

# Define the format class
class CrystalXlsx::Format
  property index : Int32 = 0
  property font_size : Int32 = 11
  property font_name : String = "Calibri"
  property bold : Bool = false
  property text_color : String? = nil
  property bg_color : String? = nil
  property border : Bool = false
  property num_form_id : Int32 = 0
  property horizontal_alignment : String? = nil
  property vertical_alignment : String? = nil

  def initialize(@num_form_id = 0, @font_size : Int32 = 11, @font_name : String = "Calibri", @bold : Bool = false, @bg_color : String? = nil, @text_color : String? = nil, @border : Bool = false, @horizontal_alignment : String? = nil, @vertical_alignment : String? = nil)
  end

  def bold? : Bool
    bold
  end

  def border? : Bool
    border
  end
end

# Simple test
describe CrystalXlsx::Format do
  it "initializes with default values" do
    format = CrystalXlsx::Format.new
    format.font_size.should eq 11
    format.font_name.should eq "Calibri"
    format.bold?.should be_false
    format.text_color.should be_nil
    format.bg_color.should be_nil
    format.border?.should be_false
  end
end