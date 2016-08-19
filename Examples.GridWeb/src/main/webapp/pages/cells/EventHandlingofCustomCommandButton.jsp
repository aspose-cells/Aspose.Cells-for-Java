<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"  pageEncoding="UTF-8"%>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<title>Event Handling of Custom Command Button</title>
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

//Create custom command event handler to handle the click event
CustomCommandEventHandler cceh=new CustomCommandEventHandler(){
	public void handleCellEvent(Object sender, String command){

	    //Identifying a specific button by checking its command
	    if (command.equals("MyButton"))
	    {
	        //Accessing the cells collection of the worksheet that is currently active
	        GridWorksheet sheet = gridweb.getWorkSheets().get(gridweb.getActiveSheetIndex());

	        //Putting value to "A1" cell
	        sheet.getCells().get("A1").putValue("My Custom Command Button is Clicked.");
	        sheet.getCells().setColumnWidth(0, 50);
	    }
	}
};

//Assign the custom command event handler created above to gridweb
gridweb.CustomCommand = cceh;

gridweb.prepareRender();

out.print(gridweb.getHTMLBody());

%>
</body>
</html>