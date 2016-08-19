<%@ page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"
    pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript">
function myvalidate(){
	//alert(this);
}
function dealwithcellselectcallback(cell, ret)
{
	alert(cell + ":" + ret+"***"+this);
   // console.log(cell + ":" + ret+"***"+this);
    //cell.styleStr = "|||||#ff0000|||||||";
                     
   // cell.style.color = "red";             
   // var instance = gridwebinstance.get("_ctl0_MainContent_GridWeb1");
   // instance.updateCellFontColor(cell, 'olive');
    // to do CELLSNET-41831 investigate servisde action is the set take affect on serverside????,want to set red color 
}
</script>
<title>Insert title here</title>
<%
ExtPage BeanManager=ExtPage.getInstance();
GridWebBean gridweb=BeanManager.getBean(request);
//gridweb.setACWClientPath("../grid/acw_client/");
 
%>
</head>
<body>
hello  world
<% 
gridweb.setReqRes(request, response);
gridweb.ImportExcelFile(application.getRealPath("/")+"/file/list.xls");
//gridweb.setOnValidationErrorClientFunction("myvalidate");
 // page=page
//	final	 HttpServletResponse response_it=response;
WorkbookEventHandler we=new WorkbookEventHandler(){
	public void handleCellEvent(Object sender, CellEventArgs e){
		System.out.println("hSaveCommand");
	}
	
};
CellEventHandler ceh=new CellEventHandler(){
	public void handleCellEvent(Object sender, CellEventArgs e){
		 System.out.println("hello cell double click");
	}
	
};
RowColumnEventHandler reh=new RowColumnEventHandler(){
	public void handleCellEvent(Object sender, RowColumnEventArgs e){
		 System.out.println("hello row.... RowColumnEventArgs");
	}
	
};

RowColumnEventHandler cdbclick=new RowColumnEventHandler(){
	public void handleCellEvent(Object sender, RowColumnEventArgs e){
		 System.out.println("hello column double click");
	}
	
};

CellEventStringHandler cesh=new CellEventStringHandler(){
	public String handleCellEvent(Object sender, CellEventArgs e){
		return e.getCell()+"$$$$hello_CellEventStringHandler";
	}
	
};

CellEventHandler cellcommand=new CellEventHandler(){
	public void handleCellEvent(Object sender, CellEventArgs e){
		 System.out.println("hello cellcommand"+e.getCell());
	}
	
};

gridweb.setEnableDoubleClickEvent(true);
gridweb.SaveCommand=we;
gridweb.CellDoubleClick=ceh;
gridweb.RowDoubleClick=reh;
gridweb.ColumnDoubleClick=cdbclick;
 gridweb.CellClickOnAjax=cesh;
 gridweb.setOnCellSelectedAjaxCallBackClientFunction("dealwithcellselectcallback");
gridweb.CellCommand=cellcommand;

 
 
gridweb.prepareRender();

out.print(gridweb.getHTMLBody());
//System.out.println(gridweb.getPresetStyle()+",has set default"+",get enable page expected false:"+gridweb.getEnablePaging());
%>
</body>
</html>