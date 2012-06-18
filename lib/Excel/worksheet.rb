module Excel
  class WorkSheet
    @worksheet_name = nil
    @workbook = nil

    def self.new(wb, ws_name)
      @workbook = wb
      @worksheet_name = ws_name
      self
    end

    def self.set_cell_value(cell, value)
      @workbook.app_object.set(@workbook.app_object.workbooks[@workbook.workbook_name].worksheets[@worksheet_name].cells[cell].value, {:to => value})
    end

    def self.get_cell_value(cell)
      @workbook.app_object.workbooks[@workbook.workbook_name].worksheets[@worksheet_name].cells[cell].value.get
    end
  end
end