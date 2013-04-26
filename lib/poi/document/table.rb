# -*- coding: utf-8 -*-
module POI
  class Tables
    def initialize(document)
      @document = document
      @poi_document = document.poi_document
      @tables = document._tables
    end
    
    def [](index)
      poi_table = @poi_document.tables[index]
      Table.new(@document, poi_table)
    end
    
    def size
      @tables.size
    end
  end
  
  class Table < Facade(:poi_table, org.apache.poi.xwpf.usermodel.XWPFTable)
    def initialize(document, table)
      @document = document
      @table = table
    end
    
    alias :_rows :rows
    def rows
      @rows ||= TableRows.new(@document, self)
    end
    
    def poi_table
      @table
    end
  end
end