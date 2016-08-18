<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"  pageEncoding="UTF-8"%>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<title>Insert title here</title>
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

// ExStart:SamplePage

WorkbookEventHandler we=new WorkbookEventHandler(){
	public void handleCellEvent(Object sender, CellEventArgs e){
		System.out.println("----------Save Command----------");
	}

};
CellEventHandler ceh=new CellEventHandler(){
	public void handleCellEvent(Object sender, CellEventArgs e){
		System.out.println("---------Cell Double Click---------");
	}

};
RowColumnEventHandler reh=new RowColumnEventHandler(){
	public void handleCellEvent(Object sender, RowColumnEventArgs e){
		System.out.println("----------Row Double Click---------------");
	}

};

RowColumnEventHandler cdbclick=new RowColumnEventHandler(){
	public void handleCellEvent(Object sender, RowColumnEventArgs e){
		System.out.println("----------Column Double Click-------------");
	}

};


gridweb.setEnableDoubleClickEvent(true);
gridweb.SaveCommand=we;
gridweb.CellDoubleClick=ceh;
gridweb.RowDoubleClick=reh;
gridweb.ColumnDoubleClick=cdbclick;

// ExEnd:SamplePage
gridweb.prepareRender();

out.print(gridweb.getHTMLBody());

%>
</body>
</html>