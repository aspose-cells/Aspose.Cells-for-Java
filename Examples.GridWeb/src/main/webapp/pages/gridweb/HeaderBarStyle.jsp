<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"  pageEncoding="UTF-8"%>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<title>Header Bar Style</title>
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

//Setting the background color of the header bars
gridweb.getHeaderBarStyle().setBackColor(Color.getBrown());

//Setting the foreground color of the header bars
gridweb.getHeaderBarStyle().setForeColor(Color.getYellow());

//Setting the font of the header bars to bold
gridweb.getHeaderBarStyle().getFont().setBold(true);

//Setting the font name to "Century Gothic"
gridweb.getHeaderBarStyle().getFont().setName("Century Gothic");

//Setting the border width to 2 points
gridweb.getHeaderBarStyle().setBorderWidth(Unit.Point(2));

gridweb.prepareRender();

out.print(gridweb.getHTMLBody());

%>
</body>
</html>