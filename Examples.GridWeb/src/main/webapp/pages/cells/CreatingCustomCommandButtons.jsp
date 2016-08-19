<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"  pageEncoding="UTF-8"%>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<title>Creating Custom Command Buttons</title>
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
//gridweb.importExcelFile(filePath);

//Instantiating a CustomCommandButton object
CustomCommandButton button = new CustomCommandButton();

//Setting the command for button
button.setCommand("MyButton");

//Setting text of the button
button.setText("MyButton");

//Setting tooltip of the button
button.setToolTip("My Custom Command Button");

//Setting image URL of the button
button.setImageUrl("icon.png");

//Adding button to CustomCommandButtons collection of GridWeb
gridweb.getCustomCommandButtons().add(button);

gridweb.prepareRender();

out.print(gridweb.getHTMLBody());

%>
</body>
</html>