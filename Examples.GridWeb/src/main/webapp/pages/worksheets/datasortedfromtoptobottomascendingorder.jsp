<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"  pageEncoding="UTF-8"%>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<title>data sorted from top to bottom in ascending order</title>
<%
ExtPage BeanManager=ExtPage.getInstance();
GridWebBean gridweb=BeanManager.getBean(request);
out.println(gridweb.getHTMLHead());
%>
</head>
<body>
<%

String filePath = application.getRealPath("/data.xls");

gridweb.setReqRes(request, response);
gridweb.importExcelFile(filePath);

//Accessing the reference of the worksheet that is currently active
GridWorksheet sheet = gridweb.getWorkSheets().get(gridweb.getActiveSheetIndex());

int startRow = 1;
int startColumn = 0;
int rows = 12;
int columns = 2;
int index = 0; //This is the index of the column or row which you need to sort
boolean isAsending = true;
boolean isCaseSensitive = false;
boolean islefttoright = false;

//Sorting data in ascending order
sheet.getCells().sort(startRow, startColumn, rows, columns, index, isAsending, isCaseSensitive, islefttoright);

gridweb.prepareRender();

out.print(gridweb.getHTMLBody());

%>
</body>
</html>