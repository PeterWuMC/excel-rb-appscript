require 'worksheet'

module Excel
  class WorkBook
    @app_object = nil
    @workbook_name = nil
  
    def self.new(app, wb_name)
      @app_object = app
      @workbook_name = wb_name
      self
    end
  
    def self.workbook_name
      @workbook_name
    end
  
    def self.app_object
      @app_object.app_object
    end
  
    def self.worksheets
      app_object.workbooks[workbook_name].worksheets.get.collect {|ws| ws.name.get }
    end
  
    def self.worksheet ws
      # raise "Worksheet not found" if !(worksheets.include? ws)
      WorkSheet.new(self, ws)
    end
  end
end