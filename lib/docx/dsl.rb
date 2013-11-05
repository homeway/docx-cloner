#encoding: utf-8
require 'docx/cloner'

module Docx
  module DSL
    lambda{
      docx = nil
      define_method :set_text do |tag, value|
        puts "#set_text: #{tag}, #{value}"
        docx.set_text_tag tag, value
      end

      define_method :set_row do |tags, data, type|
        puts "#set_row_#{type}"
        puts "tags: #{tags}"
        puts "data: #{data}"
        docx.set_row_tags tags, data, type
      end

      define_method :method_missing do |name, *args|
        puts "name: #{name}"
        puts "args: #{args}"
        super
      end

      define_method :docx_cloner do |source, dest, &b|
        docx = Docx::Cloner::DocxTool.new source
        #return unless block_given?
        puts 'start dsl'
        b.call self
        puts 'end dsl'

        docx.save dest
        docx.release
      end
    }.call
  end
end