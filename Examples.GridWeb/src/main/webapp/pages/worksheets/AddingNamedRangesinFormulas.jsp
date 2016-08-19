<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"  pageEncoding="UTF-8"%>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<title>Copying Worksheet Usin Sheet Name</title>
<%
ExtPage BeanManager=ExtPage.getInstance();
GridWebBean gridweb=BeanManager.getBean(request);
out.println(gridweb.getHTMLHead());
%>
</head>
<body>
<%

String filePath = application.getRealPath("/Sample.xlsx");

gridweb.setReqRes(request, response);

//Clear the Worksheets first
gridweb.getWorkSheets().clear();

//Load an Excel file to the GridWeb which contains a named range i.e. Range1
gridweb.importExcelFile(filePath);

//Apply a formula to a cell that refers to a named range "Range1"
gridweb.getWorkSheets().get(0).getCells().get("G6").setFormula("=SUM(Range1)");

//Add a new named range "MyRange" with based area A2:B5
int index = gridweb.getWorkSheets().getNames().add("MyRange", "Sheet1!A2:B5");
//Get the named range
GridName name = gridweb.getWorkSheets().getNames().get(index);

//Apply a formula to G7 cell
gridweb.getWorkSheets().get(0).getCells().get("G7").setFormula("=Average(MyRange)");
//Calculate the results of the formulas
gridweb.getWorkSheets().calculateFormula();
		
//Save the Excel file
gridweb.saveToExcelFile("AddingNameRangeInFormula.xlsx");

gridweb.prepareRender();

out.print(gridweb.getHTMLBody());

%>
</body>
</html>