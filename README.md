# xlsxwriter

TODO: Write a description here

## Installation

1. Add the dependency to your `shard.yml`:

   ```yaml
   dependencies:
     xlsxwriter:
       github: your-github-user/xlsxwriter
   ```

2. Run `shards install`

## Usage

```crystal
require "xlsxwriter"
```

And then execute:

```crystal
workbook = Xlsxwriter::Workbook.new

worksheet = workbook.add_worksheet("A worksheet")

worksheet.add_row(["Hello", "World"])

workbook.close("./hello_world.xlsx")
```

## Development

TODO: Write development instructions here

## Contributing

1. Fork it (<https://github.com/your-github-user/xlsxwriter/fork>)
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create a new Pull Request

## Contributors

- [Håkan Nylén](https://github.com/your-github-user) - creator and maintainer
