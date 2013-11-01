#encoding: utf-8
lib = File.expand_path('../../../lib', __FILE__)
require "#{lib}/docx/cloner"

假如(/^"(.*?)"示例文件夹中存在一个"(.*?)"的文件$/) do |folder, file|
  @filename = File.expand_path "#{folder}/#{file}"
  s = File.stat @filename
  s.file?.should be_true
end


那么(/^程序应该能读到"(.*?)"这个标签词$/) do |tag_name|
  docx = Docx::Cloner::DocxReader.new @filename
  result = docx.read_single_tag tag_name
  result.should == tag_name
end