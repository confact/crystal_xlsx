# crystal_xlsx

A library to create xlsx (Excel 2006) file with styles, column width and freeze rows.

## Installation

1. Add the dependency to your `shard.yml`:

   ```yaml
   dependencies:
     crystal_xlsx:
       github: confact/crystal_xlsx
   ```

2. Run `shards install`

## Usage

```crystal
require "crystal_xlsx"
```

And then execute:

```crystal
workbook = CrystalXlsx::Workbook.new

worksheet = workbook.add_worksheet("A worksheet")

worksheet.add_row(["Hello", "World"])
# or
# worksheet << ["Hello", "World"]

workbook.close("./hello_world.xlsx")
```

You can set the column width and freeze rows:

```crystal
workbook = CrystalXlsx::Workbook.new

worksheet = workbook.add_worksheet("A worksheet")

worksheet.add_row(["Hello", "World"])

worksheet.freeze_row(1) # Freeze the first row

worksheet.columns_width = [10, 20] # Set the width of the first column to 10 and the second to 20

workbook.close("./hello_world.xlsx")
```

You can also set the style of the cells:

```crystal
workbook = CrystalXlsx::Workbook.new

worksheet = workbook.add_worksheet("A worksheet")

style = workbook.add_format(
  font_name: "Arial",
  font_size: 12,
  bold: true,
  text_color: "FF0000",
  bg_color: "FFFF00",
  horizontal_align: "center",
  vertical_align: "center"
)

worksheet.add_row(["Hello", "World"], style)

workbook.close("./hello_world.xlsx")
```

You can get the result of the workbook as a string:

```crystal
workbook = CrystalXlsx::Workbook.new

worksheet = workbook.add_worksheet("A worksheet")

worksheet.add_row(["Hello", "World"])

puts workbook.to_s

# or send it to a IO
workbook.to_io(STDOUT)

```

## Known issues
- The library does not support formulas
- The library does not support images
- The library does not support charts


## Benchmark
As I made this library due to issues of ruby libraries being slow, I have made a benchmark to compare them. I compare it with the benchmarks at [fast_excel](https://github.com/Paxa/fast_excel) that does a benchmark of 20k write rows and compare it with caxls and write_xlsx.

```
           FastExcel:        1.4 i/s
               Axlsx:        0.4 i/s - 3.46x  slower
          write_xlsx:        0.1 i/s - 17.04x  slower
```
And crystal_xlsx (Using Crystal's Benchmark module):
```
        write 20k rows   3.52  (284.13ms) (± 2.35%)  271MB/op
```

## Development

Feel free to contribute to the project. You can run the tests with `crystal spec`.


## Contributing

1. Fork it (<https://github.com/confact/crystal_xlsx/fork>)
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request

## Contributors

- [Håkan Nylén](https://github.com/confact) - creator and maintainer
