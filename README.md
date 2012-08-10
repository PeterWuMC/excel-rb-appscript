# Excel rb-appscript

This is just an rb-appscript API for Excel.
Because of the lack of documentations on how to use Ruby - Apple Script on Excel, the idea is of this is to have some of most frequently used excel dictionary/actions here.


## Excel Structure

### Excel Module

	Application
	|-- Workbooks
		|-- Worksheets

### Application

	Excel::App.open(path_of_file)
	This will start the Excel and open the file
	Return the workbook name which it opened

	Excel::App.close
	Close the file and shut down Excel without saving

	Excel::App.workbooks
	Return an array of all the opened workbooks

	Excel::App.workbooks_paths
	Return an array of all the paths for opened workbooks

	Excel::App.workbook(workbook_name)
	Return the Workbook object

	Excel::App.save_all
	Save all opened workbook

### Workbook

	Excel::App.workbook("workbook.xls").worksheets
	Return all the worksheets in workbook "workbook.xls"

	Excel::App.workbook("workbook.xls").worksheet(worksheet_name)
	Return the Worksheet object

	Excel::App.workbook("workbook.xls").close
	Close "workbook.xls" without saving

	Excel::App.workbook("workbook.xls").save
	Save "workbook.xls"

### Worksheet

	Excel::App.workbook("workbook.xls").worksheet("sheet1").get_cell_value(cell_name)
	e.g. cell_name: "A1" will return the value of cell "A1" in "sheet1"

	Excel::App.workbook("workbook.xls").worksheet("sheet1").set_cell_value(cell_name, "Hello")
	e.g. cell_name: "A1" will update the value of cell "A1" in "sheet1" to "Hello"

	Excel::App.workbook("workbook.xls").worksheet("sheet1").get_values_in_range(from_cell, to_cell)
	e.g. from_cell "A1", to_cell "B2", it will return a hash like this
	{ A: { 1 => "hello", 2 => "world" }, B: { 1 => "lifes", 2 => "awesome" } }