<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"  pageEncoding="UTF-8"%>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<title>Create Data Validation in a GridCell of GridWeb</title>
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

//Access first worksheet
GridWorksheet sheet = gridweb.getWorkSheets().get(0);

//Access cell B3
GridCell cell = sheet.getCells().get("B3");

//Add validation inside the gridcell
//Any value which is not between 20 and 40 will cause error in a gridcell
GridValidation val = cell.createValidation(GridValidationType.WHOLE_NUMBER, true);
val.setFormula1("=20");
val.setFormula2("=40");
val.setOperator(OperatorType.BETWEEN);
val.setShowError(true);
val.setShowInput(true);

gridweb.prepareRender();

out.print(gridweb.getHTMLBody());

%>
</body>
</html>