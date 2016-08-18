<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"  pageEncoding="UTF-8"%>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<title>Double Click Row Header</title>
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

//Event Handler for RowDoubleClick event
RowColumnEventHandler re = new RowColumnEventHandler() {

	public void handleCellEvent(Object sender, RowColumnEventArgs e) {
		System.out.println("Row header:" + (e.getNum() + 1) + " is double-clicked.");
	}
};

gridweb.RowDoubleClick = re;

gridweb.prepareRender();

out.print(gridweb.getHTMLBody());

%>
</body>
</html>