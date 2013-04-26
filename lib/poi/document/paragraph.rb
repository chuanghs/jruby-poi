# -*- coding: utf-8 -*-
module POI
  class Paragraphs
    def initialize(document)
      @document = document
      @poi_document = document
    end
    
    def size
      @poi_document.get_paragraphs.size
    end
    
    def [](index)
      paragraph = @poi_document.paragraph_array(index)
      Paragraph.new(paragraph, @document)
    end
  end
  
  class Paragraph < Facade(:poi_paragraph, org.apache.poi.xwpf.usermodel.XWPFParagraph)
    def initialize(paragraph, document)
      @paragraph = paragraph
      @document = document
    end
    
    def poi_paragraph
      @paragraph
    end
    
    def text
      poi_paragraph ? poi_paragraph.getText() : "" 
    end
  end
end