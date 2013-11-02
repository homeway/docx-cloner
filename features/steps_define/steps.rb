#encoding: utf-8
lib = File.expand_path('../../../lib', __FILE__)
require "#{lib}/docx/cloner"
require 'fileutils'

假如(/^"(.*?)"示例文件夹中存在一个"(.*?)"的文件$/) do |folder, file|
  @source_filename = File.expand_path "#{folder}/#{file}"
  File.exists?(@source_filename).should be_true
end


那么(/^程序应该能读到"(.*?)"这个标签词$/) do |tag_name|
  docx = Docx::Cloner::DocxTool.new @source_filename
  result = docx.include_single_tag? tag_name
  docx.release
  result.should be_true
end


假如(/^"(.*?)"这个目标文件已经被清除$/) do |dest|
  @dest_filename = dest
  File.delete @dest_filename if File.exist?(dest)
  File.exist?(dest).should be_false
end

假如(/^程序将目标文件中的"(.*?)"替换为"(.*?)"$/) do |tag, value|
  docx = Docx::Cloner::DocxTool.new @source_filename
  result = docx.set_single_tag tag, value
  docx.save @dest_filename
  docx.release
  result.should be_true
end

那么(/^应该生成目标文件$/) do
  File.exist?(@dest_filename).should be_true
end

而且(/^被目标文件中应该包含"(.*?)"这个标签词$/) do |value|
  docx = Docx::Cloner::DocxTool.new @dest_filename
  result = docx.include_single_tag? value
  docx.release
  result.should be_true
end

假如(/^程序同时将目标文件中的"(.*?)"替换为"(.*?)"，"(.*?)"替换为"(.*?)"，"(.*?)"替换为"(.*?)"，"(.*?)"替换为"(.*?)"，"(.*?)"替换为"(.*?)"$/) do |t1, v1, t2, v2, t3, v3, t4, v4, t5, v5|
  docx = Docx::Cloner::DocxTool.new @source_filename
  result = docx.set_single_tag t1, v1
  result &= docx.set_single_tag t2, v2
  result &= docx.set_single_tag t3, v3
  result &= docx.set_single_tag t4, v4
  result &= docx.set_single_tag t5, v5
  docx.save @dest_filename
  docx.release
  result.should be_true
end

那么(/^被目标文件中应该包含"(.*?)"、"(.*?)"、"(.*?)"、"(.*?)"、"(.*?)"、这些字符串$/) do |v1, v2, v3, v4, v5|
  docx = Docx::Cloner::DocxTool.new @dest_filename
  result = docx.include_single_tag? v1
  result = docx.include_single_tag? v2
  result = docx.include_single_tag? v3
  result = docx.include_single_tag? v4
  result = docx.include_single_tag? v5
  docx.release
  result.should be_true
end
