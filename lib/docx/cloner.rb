#encoding: utf-8
require "docx/cloner/version"
require 'zip/zip'  #rubyzip gem
require 'nokogiri'

module Docx
  module Cloner
    class WordXmlFile
      def self.open(path, &block)
        self.new(path, &block)
      end

      def initialize(path, &block)
        @replace = {}
        if block_given?
          @zip = Zip::ZipFile.open(path)
          yield self
          @zip.close
        else
          @zip = Zip::ZipFile.open(path)
        end
      end

      def merge(rec)
        _xml = @zip.read("word/document.xml")
        doc = Nokogiri::XML(_xml)
        tags = doc.root.xpath("//w:t[contains(., '_Name')]")
        tags.each do |field|
          new_field = field
          if field.content == 'First_Name'
            field.inner_html = 'Adi'
            new_field.inner_html = 'My Adi'
            field.add_next_sibling(new_field.to_html)
          elsif field.content == 'Last_Name'
            field.inner_html = 'Zhou'          
          end
        end
        @replace["word/document.xml"] = doc.serialize :save_with => 0
      end

      def save(path)
        Zip::ZipFile.open(path, Zip::ZipFile::CREATE) do |out|
          @zip.each do |entry|
            out.get_output_stream(entry.name) do |o|
              if @replace[entry.name]
                o.write(@replace[entry.name])
              else
                o.write(@zip.read(entry.name))
              end
            end
          end
        end
        @zip.close
      end
    end

    class DocxReader

      '加载docx文件，将段落存储到@paragraph，用@paragraph[:text_content]检索，再从段落内检索xml标签位置'
      def initialize(file)
        @zip = Zip::ZipFile.open(file)
        @paragraph = []
        @replace = {}
        _xml = @zip.read("word/document.xml")
        @doc = Nokogiri::XML(_xml)
        wp_set = @doc.xpath(".//w:p")
        #puts "#{wp_set.size}'s wp"
        wp_set.each do |wp|
          p = {text_content: '', text_run: []}
          wp.xpath(".//w:t").each do |t|
            p[:text_content] << t.content
            p[:text_run] << t
          end
          @paragraph << p
          #puts p[:text_content].include? '$名字$'
        end
        #puts @paragraph
      end

      def release
        @zip.close
      end
      
      def include_single_tag?(tag)
        @paragraph.each do |p|
          if p[:text_content].include? tag
            return true
          end
        end
        return false
      end

      def read_single_tag_xml(tag)
        findit = false
        @paragraph.each do |p|
          if p[:text_content].include? tag
            findit = true
            return tag
          end
        end
        return ''
      end

    end
  end
end
