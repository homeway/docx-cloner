#encoding: utf-8
require 'spec_helper'

module Docx
  module Cloner
    describe DocxReader do
      before :all do
        @docx = DocxReader.new 'docx-examples/read-single-tags.docx'
      end

      context "#read_single_tag" do
        it "读取'$名字$'标签" do
          result = @docx.read_single_tag "名字"
          result.should == "名字"
        end
      end

      after :all do
        @docx.release
      end
    end
  end
end
