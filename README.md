# Docx::Cloner

TODO: Write a gem description

## Installation

Add this line to your application's Gemfile:

    gem 'docx-cloner'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install docx-cloner

## Usage
    require 'docx-cloner'

    sourc_file = 'source.docx'
    dest_file = 'dest.docx'
    docx = DocxTool.new @sourc_file

    docx.set_single_tag '{Name}', '周大福'

    table_title = ["{名称1}", "{00.01}"]
    table_data = [["自行车1", "125.00"], ["大卡车1", "256500.00"]]
    docx.set_row_tags table_title, table_data, 'tr'

    docx.save dest_file
    docx.release


## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
