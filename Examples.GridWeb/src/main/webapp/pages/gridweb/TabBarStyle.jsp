<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"  pageEncoding="UTF-8"%>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<title>Tab Bar Style</title>
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

//Setting the background color of tabs to Yellow
gridweb.getTabStyle().setBackColor(Color.getYellow());

//Setting the foreground color of tabs to Blue
gridweb.getTabStyle().setForeColor(Color.getBlue());

//Setting the background color of active tab to Blue
gridweb.getActiveTabStyle().setBackColor(Color.getBlue());

//Setting the foreground color of active tab to Yellow
gridweb.getActiveTabStyle().setForeColor(Color.getYellow());

gridweb.prepareRender();

out.print(gridweb.getHTMLBody());

%>
</body>
</html>