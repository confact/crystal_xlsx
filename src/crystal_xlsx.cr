require "xml"
require "./crystal_xlsx/formula/*"
require "./crystal_xlsx/*"

module CrystalXlsx
  VERSION = "0.1.0"

  # max columns in a worksheet for Excel 2007
  MAX_COLUMNS = 16_384

  NAME = "Crystal.xlsx"

  EPOCH          = Time.utc(1899, 12, 30).to_unix_f
  DAY_IN_SECONDS = 86_400
end
