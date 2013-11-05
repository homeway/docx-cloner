#encoding: utf-8
require 'docx/cloner'

module Docx
  module DSL
    lambda{
      docx = nil
      define_method :set_text do |tag, value|
        #puts "#set_text: #{tag}, #{value}"
        docx.set_text_tag tag, value
      end

      define_method :set_row do |tags, data, type|
        #puts "#set_row_#{type}"
        #puts "tags: #{tags}"
        #puts "data: #{data}"
        docx.set_row_tags tags, data, type
      end

      define_method :docx_cloner do |source, dest, &b|
        docx = Docx::Cloner::DocxTool.new source
        #return unless block_given?
        #puts 'start dsl'
        b.call self
        #puts 'end dsl'

        docx.save dest
        docx.release
      end
    }.call

    def method_missing name, *args
      puts "name: #{name}"
      puts "args: #{args}"
      /set_row_(.+)/.match name do
        return set_row *(args << $1) if $1
      end
      super
    end

  end
end