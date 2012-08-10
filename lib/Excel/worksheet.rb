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

    def self.get_values_in_range(from, to)
      from = from.match(/^([a-zA-Z]+)([0-9]+)$/)
      from = {column: from[1], row: from[2].to_i}

      to = to.match(/^([a-zA-Z]+)([0-9]+)$/)
      to = {column: to[1], row: to[2].to_i}

      return_hash = Hash.new{|h,k| h[k] = Hash.new}
      (from[:column]..to[:column]).each do |column|
        (from[:row]..to[:row]).each do |row|
          return_hash[column][row] = get_cell_value("#{column}#{row}")
        end
      end
      return_hash
    end
  end
end