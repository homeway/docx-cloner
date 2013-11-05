#encoding: utf-8
require 'docx/dsl'

<<-ABC
sourc_file = 'source.docx'
dest_file = 'dest.docx'
docx = Docx::Cloner::DocxTool.new sourc_file

docx.set_text_tag '{Name}', '周大福'

table_title = ["{名称1}", "{00.01}"]
table_data = [["自行车1", "125.00"], ["大卡车1", "256500.00"]]
docx.set_row_tags table_title, table_data, 'tr'

docx.save dest_file
docx.release

ABC
##################

extend Docx::DSL
docx_cloner 'source.docx', 'dest.docx' do
  set_text '{Name}', '周大福'
  set_row 'tr' do |t|
    t[:tags] = ["{名称1}", "{00.01}"]
    t[:data] = [["自行车1", "125.00"], ["大卡车1", "256500.00"]]
  end
end
