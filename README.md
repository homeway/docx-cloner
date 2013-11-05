# Docx::Cloner

This is a tool to clone docx file with tag field. And it supported utf-8 well.

## Installation

Add this line to your application's Gemfile:

    gem 'docx-cloner'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install docx-cloner

## Usage
* docx_cloner
* set_text
* set_row_tr

it's so easy to use it with DSL:

    #encoding: utf-8
    require 'docx/dsl'

    class C
      include Docx::DSL

      def my_method
        #template is 'source.docx' and save to 'dest.docx'
        docx_cloner 'source.docx', 'dest.docx' do
          #replace the text of '{Name}' in 'source.docx'
          set_text '{Name}', '周大福'

          #multi lines replace
          tags = ["{名称1}", "{00.01}"]
          data = [["自行车1", "125.00"], ["大卡车1", "256500.00"]]
          set_row_tr tags, data
        end
      end
    end

## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
