require "../src/crystal_xlsx"

# Create a workbook with the new API
workbook = CrystalXlsx::Workbook.new

# Create a worksheet
workbook.sheet("Sales Data") do |sheet|
  # Add data
  sheet.add(["Product", "Q1", "Q2", "Q3", "Q4"])
  sheet.add(["Widgets", 100, 150, 200, 250])
  sheet.add(["Gadgets", 75, 125, 175, 225])
  
  # Add a formula (row 1, column 4 = E2)
  sheet.formula(1, 4, "SUM(B2:E2)")
  sheet.formula(2, 4, "SUM(B3:E3)")
  
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
  sheet.add(["Total Sales", "=Sales Data!E2+E3"])
  sheet.link_to_sheet(0, 0, workbook.sheet("Sales Data"), "A1", "Back to Sales")
end

# Save the workbook
workbook.save("example.xlsx")

puts "Excel file created successfully!"