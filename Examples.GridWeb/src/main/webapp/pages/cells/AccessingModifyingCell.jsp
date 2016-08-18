<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"  pageEncoding="UTF-8"%>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<title>Accessing & Modifying a Cell's Value</title>
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

//Accessing "B1" cell of the worksheet
GridCell cell = sheet.getCells().get("B1");

//Putting a value in "B1" cell
cell.putValue(Calendar.getInstance());

gridweb.prepareRender();

out.print(gridweb.getHTMLBody());

%>
</body>
</html>