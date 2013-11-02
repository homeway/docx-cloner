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

    class DocxTool

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

      def save(path)
        @replace["word/document.xml"] = @doc.serialize :save_with => 0

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
        @paragraph.each do |p|
          if p[:text_content].include? tag
            from = p[:text_content].index tag
            to = from + tag.size - 1
            #puts "from:#{from}, to:#{to}"
            pos = 0
            dest = ""
            p[:text_run].each do |wt|
              #puts "pos:#{pos}"
              if pos >= from && pos < to
                dest << wt.parent.to_xml << "\n"
              end
              if pos >= to
                return dest
              end
              pos += wt.content.size
            end
            return dest
          end

        end
        return ''
      end

      def set_single_tag(tag, value)
        @paragraph.each do |p|
          #puts p[:text_content]
          if p[:text_content].include? tag
            from = p[:text_content].index tag
            to = from + tag.size - 1
            #puts "from:#{from}, to:#{to}"
            pos = 0
            dest = []
            #puts p[:text_run]
            p[:text_run].each do |wt|
              #puts "pos:#{pos}"
              #通常情况下，msword会把标签拆分成多个xml标签，如'{name}'被拆分成'<wt>{</wt>'和'<wt>name}</wt>'
              #这可能跟编辑器有关，在处理中文时，这是一种常见的情形
              if pos >= from && pos < to
                dest << wt
              end
              if pos >= to
                break if pos >= to
              end
              pos += wt.content.size

              #这里要处理一下标签没有被拆分的情形，而是作为纯文本被包含在某个标签中
              #例如'{name}'包含在'<wt>my {name}</wt>'中
              #puts "pos:#{pos}, to:#{to}, dest.size:#{dest.size}"
              #puts wt
              if pos >= to && dest.size == 0
                wt.inner_html = wt.content.sub(tag, value)
                return true #如果是这种简单情形，就不再需要后续处理了
              end
            end

            if dest.size > 0
              #puts dest.first
              dest.first.inner_html = value
              dest[1..-1].each {|node| node.remove }
              return true
            else
              return false
            end
          end

        end
        return ''
      end

    end
  end
end
