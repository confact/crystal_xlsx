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

workbook.close("./hello_world.xlsx")
```

## Development

TODO: Write development instructions here

## Contributing

1. Fork it (<https://github.com/confact/crystal_xlsx/fork>)
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request

## Contributors

- [Håkan Nylén](https://github.com/confact) - creator and maintainer
