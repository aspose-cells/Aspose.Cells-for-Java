<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"  pageEncoding="UTF-8"%>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<title>Handling Grid Events</title>
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

CellEventHandler ce = new CellEventHandler() {
	public void handleCellEvent(Object sender, CellEventArgs e) {
		//Your event handler code goes here
		if (e.getArgument().toString().equals("A1")) {
			//Your rest of the code
		}
	}
};
gridweb.CellCommand = ce;

gridweb.prepareRender();

out.print(gridweb.getHTMLBody());

%>
</body>
</html>