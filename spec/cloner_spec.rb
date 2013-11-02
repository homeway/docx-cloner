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
          sourc_file = 'docx-examples/source.docx'
          dest_file = 'docx-examples/dest.docx'
          File.delete dest_file if File.exist?(dest_file)

          FileUtils.copy sourc_file, dest_file
          @dest_docx = DocxTool.new dest_file
        end

        context "#set_single_tag" do
          it "设置单个标签{name}" do
            result = @dest_docx.set_single_tag '{name}', '周大福'
            result.should be_true
          end

        end

        after :all do
          @dest_docx.release
          File.delete dest_file if File.exist?(dest_file)
        end
      end

    end
  end
end
