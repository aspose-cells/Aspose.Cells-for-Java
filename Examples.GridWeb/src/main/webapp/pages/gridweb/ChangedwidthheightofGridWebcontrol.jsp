<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"  pageEncoding="UTF-8"%>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<title>Changed width and height of the GridWeb control</title>
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

//Setting the height of GridWeb control
gridweb.setHeight(Unit.Point(200));

//Setting the width of GridWeb control
gridweb.setWidth(Unit.Point(520));

gridweb.prepareRender();

out.print(gridweb.getHTMLBody());

%>
</body>
</html>