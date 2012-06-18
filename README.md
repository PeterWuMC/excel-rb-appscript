# Excel rb-appscript

	This is just an extension on rb-appscript, for Excel.
	Because of the lack of documentation how to use Ruby - Apple Script on Excel, the idea is to have some of most frequently used excel dictionary/actions.


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
	Close the file and shut down Excel
	
	Excel::App.workbooks
	Return an array of all the opened workbooks
	
	Excel::App.workbook(workbook_name)
	Return the Workbook object
	
### Workbook

	Excel::App.workbook("workbook.xls").worksheets
	Return all the worksheets in workbook "workbook.xls"
	
	Excel::App.workbook("workbook.xls").worksheet(worksheet_name)
	Return the Worksheet object

### Worksheet

	Excel::App.workbook("workbook.xls").worksheet("sheet1").get_cell_value(cell_name)
	e.g. cell_name: "A1" will return the value of cell "A1" in "sheet1"
	
	Excel::App.workbook("workbook.xls").worksheet("sheet1").set_cell_value(cell_name, "Hello")
	e.g. cell_name: "A1" will update the value of cell "A1" in "sheet1" to "Hello"