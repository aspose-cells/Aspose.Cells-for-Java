<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"  pageEncoding="UTF-8"%>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<title>Text URL Hyperlinks</title>
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

//Accessing the reference of the worksheet that is currently active
GridWorksheet sheet = gridweb.getWorkSheets().get(gridweb.getActiveSheetIndex());

//Adding hyperlink to the worksheet
String address = "http://www.aspose.com";
int lnkIdx = sheet.getHyperlinks().add("B1", address);
GridHyperlink lnk = sheet.getHyperlinks().get(lnkIdx);

//Setting text of the hyperlink
lnk.setTextToDisplay("Aspose");

//Setting target of the hyperlink
lnk.setTarget("_blank");

//Setting tool tip of the hyperlink
lnk.setScreenTip("Open Aspose Web Site in new window");

//Adding hyperlink to the worksheet
address = "http://www.aspose.com/community/forums/aspose.cells-product-family/19/showforum.aspx";
lnkIdx = sheet.getHyperlinks().add("B2", address);
lnk = sheet.getHyperlinks().get(lnkIdx);

//Setting text of the hyperlink
lnk.setTextToDisplay("Aspose.Grid Technical Support");

//Setting target of the hyperlink
lnk.setTarget("_parent");

//Setting tool tip of the hyperlink
lnk.setScreenTip("Open Aspose.Grid Technical Support in parent window");

gridweb.prepareRender();

out.print(gridweb.getHTMLBody());

%>
</body>
</html>