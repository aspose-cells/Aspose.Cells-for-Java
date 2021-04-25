<%@page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*" pageEncoding="UTF-8"%> 

<!DOCTYPE html> 
<html xmlns="http://www.w3.org/1999/xhtml"> 
<head> 
<% 
String path = request.getContextPath(); 
String basePath = request.getScheme()+"://"+request.getServerName()+":"+request.getServerPort()+path+"/"; 
%> 

<base href="<%=basePath%>"> 
<script type="text/javascript" language="javascript" src="grid/acw_client/acwmain.js"></script> 
<script type="text/javascript" language="javascript" src="grid/acw_client/lang_en.js"></script> 
<script type="text/javascript" language="javascript" src="Scripts/jquery-2.1.1.js"></script> 


<title>Aspose.Cells.GridWeb for Java - Sample JSP Page</title> 
<% 
//Print GridWeb version on Console 
System.out.println("Aspose.Cells.GridWeb for Java Version: " + GridWebBean.getVersion()); 

System.out.println(path); 
System.out.println(basePath); 


ExtPage BeanManager=ExtPage.getInstance(); 

//GridWebBean gridweb=BeanManager.getBean(request); 
BeanManager.setServlet(request,response);
GridWebBean gridweb=BeanManager.getBean();
out.println(gridweb.getHTMLHead()); 


//gridweb.setReqRes(request, response); 
gridweb.importExcelFile(application.getRealPath("/")+"/file/comments2.xls"); 
  
GridWorksheet sheet=gridweb.getWorkSheets().get(0); 
GridCommentCollection gcc=sheet.getComments(); 
int id=gcc.add("C6"); 
GridComment gc=gcc.get(id); 
gc.setNote("hello comment in c6"); 

  
  
gridweb.prepareRender(); 

out.print(gridweb.getHTMLBody()); 
%> 
<script type="text/javascript"> 
function ReadGridWebCells() { 
// Access GridWeb instance and cells array 
var gridwebins = gridwebinstance.get("<%=gridweb.get_ClientID()%>"); 
var cells = gridwebins.getCellsArray(); 
// Log cell name, values, row & column indexes in console 
for (var j = 0; j < cells.length; j++) 
{ 
var cellInfo = j + ":" + gridwebins.getCellName(cells[j]) + ","; 
cellInfo += "value is:" + gridwebins.getCellValueByCell(cells[j]) + " ,"; 
cellInfo += "row:" + gridwebins.getCellRow(cells[j]) + ","; 
cellInfo += "col:" + gridwebins.getCellColumn(cells[j]); 
console.log(cellInfo); 
document.write(cellInfo); 

} 
} 
</script> 

</head> 
<body> 
<% 
//gridweb.setReqRes(request, response); 
//gridweb.setEnableAJAX(true); 
//gridweb.setWidth(Unit.Pixel(400)); 
//gridweb.setHeight(Unit.Pixel(400)); 
//gridweb.prepareRender(); 
//out.print(gridweb.getHTMLBody()); 
%> 

<button type="button" onclick="ReadGridWebCells()">Click me</button> 
</body> 
</html> 