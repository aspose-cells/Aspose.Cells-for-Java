<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"  pageEncoding="UTF-8"%>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<title>Accessing a Comments</title>
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
GridWorksheet sheet =  gridweb.getWorkSheets().get(gridweb.getActiveSheetIndex());

//Accessing a specific cell that contains comment
GridCell cell = sheet.getCells().get("A1");

//Accessing the comment from the specific cell
GridComment comment = sheet.getComments().get(cell.getName());

//Modifying the comment note
comment.setNote("I have modified the comment note.");

gridweb.prepareRender();

out.print(gridweb.getHTMLBody());

%>
</body>
</html>