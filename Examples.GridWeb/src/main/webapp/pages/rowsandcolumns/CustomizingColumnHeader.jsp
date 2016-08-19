<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"  pageEncoding="UTF-8"%>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<title>Customizing Column Header</title>
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

//Accessing the worksheet that is currently active
GridWorksheet worksheet = gridweb.getWorkSheets().get(gridweb.getActiveSheetIndex());

//Setting the header of 1st column to "ID"
worksheet.SetColumnCaption(0, "ID");

//Setting the header of 2nd column to "Name"
worksheet.SetColumnCaption(1, "Name");

gridweb.prepareRender();

out.print(gridweb.getHTMLBody());

%>
</body>
</html>