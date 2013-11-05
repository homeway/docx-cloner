#encoding: utf-8
require 'docx/cloner'

module Docx
  module DSL
    def set_text tag, value
      puts "#set_text: #{tag}, #{value}"
      @docx.set_text_tag tag, value
    end

    def set_row tag
      puts "#set_row_#{tag}"
      md = {}
      yield md
      puts "tags: #{md[:tags]}"
      puts "data: #{md[:data]}"
      @docx.set_row_tags md[:tags], md[:data], tag
    end

    def docx_cloner source, dest
      @docx = Docx::Cloner::DocxTool.new source
      return unless block_given?
      puts 'start dsl'
      yield self
      puts 'end dsl'

      @docx.save dest
      @docx.release
    end

  end
end