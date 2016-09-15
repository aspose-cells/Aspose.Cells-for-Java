package com.aspose.cells.examples.articles;

public class WorkingWithColdFusion {

	<html>
	<head><title>Hello World!</title></head>
	<body>
	    <b>This example shows how to create a simple MS Excel Workbook using Aspose.Cells</b>
	    <cfset workbook=CreateObject("java", "com.aspose.cells.Workbook").init()>
	    <cfset worksheets = workbook.getWorksheets()>
	    <cfset sheet= worksheets.get("Sheet1")>
	    <cfset cells= sheet.getCells()>
	    <cfset cell= cells.getCell(0,0)>
	    <cfset cell.setValue("Hello World!")>
	    <cfset workbook.save("C:\output.xls")>
	</body>
	</html>

}
