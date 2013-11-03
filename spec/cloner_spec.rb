#encoding: utf-8
require 'spec_helper'

module Docx
  module Cloner
    describe DocxTool do
      context "以下适用于读单个标签" do
        before :all do
          @docx = DocxTool.new 'docx-examples/read-single-tags.docx'
        end
        context "#include_single_tag?" do
          it "读取'$名字$'标签" do
            result = @docx.include_single_tag? "$名字$"
            result.should be_true
          end
          it "读取'{Name}'标签" do
            result = @docx.include_single_tag? "{Name}"
            result.should be_true
          end
        end

        context "#read_single_tags_xml" do
          it "读取'{name}'的xml标签'<w:r>'，应该是" do
            result = @docx.read_single_tag_xml "{name}"
            result.should  == <<-HERE
<w:r w:rsidR="000F595B">
  <w:rPr>
    <w:rFonts w:hint="eastAsia"/>
    <w:lang w:eastAsia="zh-CN"/>
  </w:rPr>
  <w:t>{</w:t>
</w:r>
<w:r>
  <w:rPr>
    <w:rFonts w:hint="eastAsia"/>
    <w:lang w:eastAsia="zh-CN"/>
  </w:rPr>
  <w:t>n</w:t>
</w:r>
<w:r w:rsidR="000F595B">
  <w:rPr>
    <w:rFonts w:hint="eastAsia"/>
    <w:lang w:eastAsia="zh-CN"/>
  </w:rPr>
  <w:t>ame}</w:t>
</w:r>
            HERE
          end
        end

        after :all do
          @docx.release
        end
      end

      context "以下测试是针对标签替换，即写操作" do
        before :all do
          @sourc_file = 'docx-examples/source.docx'
          @dest_file = 'docx-examples/dest.docx'
          File.delete @dest_file if File.exist?(@dest_file)

          @source_docx = DocxTool.new @sourc_file
        end

        context "#set_single_tag" do
          it "设置单个标签{Name}" do
            value = '周大福'
            @source_docx.set_single_tag '{Name}', value
            @source_docx.save @dest_file

            dest = DocxTool.new @dest_file
            result = dest.include_single_tag? value
            dest.release
            result.should be_true
          end
        end

        context "#set_row_tags" do
          it "找到标签所在行的父节点" do
            tags = ["{名称1}", "{00.01}"]
            tags.each do |tag|
              node = @source_docx.get_tag_scope tag, 'tr'
              node.node_name.should == 'tr'
            end
          end
          it "设置表格中的行标签", wip: true do
            data = [["{名称1}", "{00.01}"], ["自行车1", "125.00"], ["大卡车1", "256500.00"], ["自行车2", "125.00"], ["大卡车2", "256500.00"],]
            result = @source_docx.set_row_tags data.first, data[1..-1], 'tr'
            @source_docx.save @dest_file
            result.should be_true            
          end
        end

        after :all do
          @source_docx.release
        end
      end

    end
  end
end
