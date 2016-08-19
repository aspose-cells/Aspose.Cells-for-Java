<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"  pageEncoding="UTF-8"%>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<title>Adding Formulas to Cells</title>
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
gridweb.importExcelFile(filePath);

//Accessing the worksheet of the Grid that is currently active
GridWorksheet sheet = gridweb.getWorkSheets().get(gridweb.getActiveSheetIndex());

//Putting some values to cells
sheet.getCells().get("A1").putValue("1st Value");
sheet.getCells().get("A2").putValue("2nd Value");
sheet.getCells().get("A3").putValue("Sum");
sheet.getCells().get("B1").putValue(125.56);
sheet.getCells().get("B2").putValue(23.93);

//Calculating all formulas added in worksheets
gridweb.getWorkSheets().calculateFormula();

//Adding a simple formula to "B3" cell
sheet.getCells().get("B3").setFormula("=SUM(B1:B2)");

gridweb.prepareRender();

out.print(gridweb.getHTMLBody());

%>
</body>
</html>