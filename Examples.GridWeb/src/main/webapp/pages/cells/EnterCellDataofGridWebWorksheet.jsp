<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"  pageEncoding="UTF-8"%>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<title>Accessing the cell by name</title>
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

//Access cell A1 of first gridweb worksheet
GridCell cellA1 = gridweb.getWorkSheets().get(0).getCells().get("A1");

//Access cell style and set its number format to 10 which is a Percentage 0.00% format
GridTableItemStyle st = cellA1.getStyle();
st.setNumberType(10);
cellA1.setStyle(st);

gridweb.prepareRender();

out.print(gridweb.getHTMLBody());

%>
</body>
</html>