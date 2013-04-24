# -*- coding: utf-8 -*-
require 'spec_helper'

DOCX_SAMPLE_PATH = TestDataFile.expand_path("simple_with_table.docx")

describe POI::Document do
  it "should open a docx format document and allow access to its content" do
    doc = POI::Document.open(DOCX_SAMPLE_PATH)
    doc.filename.should == DOCX_SAMPLE_PATH
  end
end