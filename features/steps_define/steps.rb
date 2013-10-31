#encoding: utf-8
lib = File.expand_path('../../../lib', __FILE__)
require "#{lib}/docx/cloner"

假如(/^"(.*?)"示例文件夹中存在一个"(.*?)"的文件$/) do |folder, file|
  s = File.stat "#{folder}/#{file}"
  @filename = file
  s.file?.should be_true
  @docx = Docx::Cloner::DocxReader.new @filename
end

假如(/^我在DSL中设定一个正则表达式"(.*?)"作为提取标签的算法$/) do |regx|
  @docx.set_regx regx
end

那么(/^程序应该能读到"(.*?)"这个标签词$/) do |tag_name|
  result = @docx.read_single_tag tag_name
  result.should == tag_name
end