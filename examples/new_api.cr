require "../src/crystal_xlsx"

# Create a workbook with the new API
CrystalXlsx.create("example.xlsx") do
  # Create a worksheet with a block
  workbook = CrystalXlsx::Workbook.new
  workbook.sheet("Sales Data") do |sheet|
    # Add data
    sheet.add(["Product", "Q1", "Q2", "Q3", "Q4"])
    sheet.add(["Widgets", 100, 150, 200, 250])
    sheet.add(["Gadgets", 75, 125, 175, 225])
    
    # Add a formula
    sheet.formula(1, 5, "SUM(B2:E2)")
    sheet.formula(2, 5, "SUM(B3:E3)")
    
    # Add a hyperlink
    sheet.link(0, 0, "https://example.com", "Product Info")
    
    # Merge cells for a title
    sheet.merge("A1:E1")
    
    # Set column widths
    sheet.column_widths = [20, 10, 10, 10, 10, 15]
    
    # Freeze the header row
    sheet.freeze_row(1)
  end
  
  # Create another worksheet
  workbook.sheet("Summary") do |sheet|
    sheet.add(["Total Sales", "=Sales Data!F2+F3"])
    sheet.link_to_sheet(0, 0, workbook.sheet("Sales Data"), "A1", "Back to Sales")
  end
end

puts "Excel file created successfully!"