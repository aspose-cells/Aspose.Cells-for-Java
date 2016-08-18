<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"  pageEncoding="UTF-8"%>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<title>Adding a Comments</title>
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

//Accessing the reference of the worksheet that is currently active
GridWorksheet sheet = gridweb.getWorkSheets().get(gridweb.getActiveSheetIndex());

//Adding comment to "A1" cell of the worksheet
GridComment comment = sheet.getComments().get(sheet.getComments().add("A1"));

//Setting the comment note
comment.setNote("These are my comments for the cell");

gridweb.prepareRender();

out.print(gridweb.getHTMLBody());

%>
</body>
</html>