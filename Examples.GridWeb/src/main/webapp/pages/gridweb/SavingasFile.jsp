<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"  pageEncoding="UTF-8"%>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<title>Protecting Cells in Rows & Columns</title>
<%
ExtPage BeanManager=ExtPage.getInstance();
GridWebBean gridweb=BeanManager.getBean(request);
out.println(gridweb.getHTMLHead());
%>
</head>
<body>
<%

gridweb.setReqRes(request, response);

//Saving Grid content to an Excel file
gridweb.saveToExcelFile("/new.xlsx");

gridweb.prepareRender();

out.print(gridweb.getHTMLBody());

%>
</body>
</html>