require 'appscript'
require 'excel/workbook'

module Excel
  class App
    @file_path          = ""
    @current_worksheet  = nil
  
    # attr_accessor :trades_hash
    @app_object = Appscript.app("Microsoft Excel")
    @work_book  = nil
    @work_sheet = nil
  
    def self.app_object
      @app_object
    end
  
    def self.open(params)
      file_path = MacTypes::Alias.path(params[:path])
      app_object.run
      app_object.activate
    
      app_object.open file_path
    
      File.basename(file_path.path)
    end
  
    def self.workbooks
      return [] if app_object.workbooks.get == :missing_value
      app_object.workbooks.get.collect {|wb| wb.name.get }
    end
    
    def self.workbooks_paths
      return [] if app_object.workbooks.get == :missing_value
      app_object.workbooks.get.collect {|wb| "#{MacTypes::FileURL.hfs_path(wb.path.get).path}/#{wb.name.get}" }
    end
  
    def self.workbook wb
      # raise "Workbook not found" if !(workbooks.include? wb)
      WorkBook.new(self, wb)
    end
  
    def self.close
      app_object.run
      app_object.activate
      app_object.workbooks.saved.set(true)
      app_object.workbooks.close
      app_object.quit
    end
    
    def self.save_all
      app_object.save
    end
  end
end