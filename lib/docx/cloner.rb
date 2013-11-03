#encoding: utf-8
require "docx/cloner/version"
require 'zip/zip'  #rubyzip gem
require 'nokogiri'

module Docx
  module Cloner
    class DocxTool

      '加载docx文件，将段落存储到@paragraph，用@paragraph[:text_content]检索，再从段落内检索xml标签位置'
      def initialize(file)
        @zip = Zip::ZipFile.open(file)
        _xml = @zip.read("word/document.xml")
        @doc = Nokogiri::XML(_xml)
        @global_paragraph = generate_paragraph @doc

        @replace = {}

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
        @global_paragraph.each do |p|
          if p[:text_content].include? tag
            return true
          end
        end
        return false
      end

      def read_single_tag_xml(tag)
        @global_paragraph.each do |p|
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

      #替换单个标签为指定值
      def set_single_tag tag, value
        replace_tag tag, value
      end

      #获取标签所在的范围，例如表格的行
      #简单的考虑，则tags中第一个标签位置即可确定为scope位置
      #复杂的考虑，则可根据tags中所有标签的共同根（如<w:tr>）确定scope位置，这种情况将允许标签名拥有自己的作用域
      #这里仅做简单的考虑
      def get_tag_scope tag, type
        @global_paragraph.each do |p|
          if p[:text_content].include? tag #这里是简单的考虑，即使行内标签也必须全局唯一
            node = p[:text_run].first
            while true
              return unless node                      #查找父节点失败
              return node if node.node_name == type   #查找到匹配的父节点
              node = node.parent
            end
          end
        end
        return false
      end

      def generate_paragraph node
        paragraphs = []
        puts "查找范围：#{node.path}"
        wp_set = node.xpath(".//w:p")
        #puts "#{wp_set.size}'s wp"
        wp_set.each do |wp|
          p = {text_content: '', text_run: []}
          wp.xpath(".//w:t").each do |t|
            p[:text_content] << t.content
            p[:text_run] << t
            #puts "node name: #{t.node_name}" if t.content.size > 0
            #puts t.path
          end
          paragraphs << p
          #puts p[:text_content].include? '$名字$'
        end
        return paragraphs
      end

      #在指定的范围内替换标签
      def replace_tag tag, value, node=nil
        paragraphs = node ? generate_paragraph(node) : @global_paragraph 
        #puts paragraphs
        paragraphs.each do |p|
          #puts p[:text_content]
          if p[:text_content].include? tag
            from = p[:text_content].index tag
            to = from + tag.size - 1
            #puts "tag:#{tag} | from:#{from}, to:#{to} >> #{p[:text_content]}"
            pos = 0
            dest = []
            #puts p[:text_run]
            p[:text_run].each do |wt|
              #puts "pos:#{pos}"
              #通常情况下，msword会把标签拆分成多个xml标签，如'{name}'被拆分成'<wt>{</wt>'和'<wt>name}</wt>'
              #这可能跟编辑器有关，在处理中文时，这是一种常见的情形
              if pos+1 >= from && pos <= to #通过pos+1修正临界点问题
                dest << wt
              end
              if pos > to
                break
              end
              pos += wt.content.size

              #这里要处理一下标签没有被拆分的情形，而是作为纯文本被包含在某个标签中
              #例如'{name}'包含在'<wt>my {name}</wt>'中
              #puts "pos:#{pos}, to:#{to}, dest.size:#{dest.size}"
              #puts wt
              if pos >= to && dest.size == 0
                #puts "simple_type | pos:#{pos}, to:#{to} >> #{wt.content}"
                wt.inner_html = wt.content.sub(tag, value)
                return true #如果是这种简单情形，就不再需要后续处理了
              end
            end

            if dest.size > 0
              puts "被替换节点：#{dest.first.path}"
              dest.first.content = value
              dest[1..-1].each do |node|
                #puts node
                node.remove
              end
              #puts "\n"
              return true
            else
              return false
            end
          end

        end
        return false

      end

      #clone标签所在的范围，例如表格的行
      #返回一组新的行对象集合
      def clone_tag_scope node, times
        #puts "clone #{node.node_name} #{times} times"
        nodes = Array.new times
        puts "被克隆节点：#{node.path}"
        times.downto(1).each do |_i|
          i = _i.to_i - 1
          nodes[i] = node.dup
          node.add_next_sibling nodes[i]
          puts "第#{i+1}个节点克隆：#{nodes[i].path}"
        end
        return nodes
      end

      #根据行标签设置，替换成多行数据，这里考虑表格的一般情况
      def set_row_tags tags, values, type
        puts "tags:#{tags}, values:#{values}, type:#{type}"
        #找到标签所在行的父节点
        tag_scope_node = get_tag_scope tags.first, type
        value_scope_nodes = clone_tag_scope tag_scope_node, values.size
        value_scope_nodes.each_with_index do |node, r|
          puts "查找范围：#{node.path}"
          tags.each_with_index do |tag, c|
            replace_tag tag, values[r][c], node
          end
        end
        #清除标签
        tag_scope_node.remove
        return true
      end

    end
  end
end
