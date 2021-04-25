<%@ page language="java" contentType="text/html; charset=UTF-8"
 import="com.aspose.gridweb.*" pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE9"/>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>Insert title here</title>
<%
License al=new License();
//al.setLicense("D:\\grid_project\\ZZZZZZ_release_version\\Licenses\\Aspose.Total.Java20151205.lic");

ExtPage BeanManager=ExtPage.getInstance();
//BeanManager.setPageaction("http://localhost:18080/gridwebdemoV8.6.3//GridWebServlet"); 
//BeanManager.setPageajaxcallpath("http://localhost:18080/gridwebdemoV8.6.3//GridWebServlet?acw_ajax_call=true"); 
//GridWebBean gridweb=BeanManager.getBean(request);
BeanManager.setServlet(request,response);
GridWebBean gridweb=BeanManager.getBean();
//gridweb.setACWClientPath("../grid/acw_client/");
out.println(gridweb.getHTMLHead());
%>
</head>
<BODY>
 

<%
//gridweb.setReqRes(request, response);
gridweb.setEnableAsync(false);
//String test= (Class.forName("com.aspose.gridweb.GridWebBean").getProtectionDomain().getCodeSource().getLocation()).getPath() ; 
gridweb.importExcelFile(application.getRealPath("/")+"/file/charttest.xls");
 gridweb.getWorkSheets().get(0).setColumnCaption(1, "Price"); 
 gridweb.getWorkSheets().get(0).setColumnHeaderToolTip(1, "Unit Price of Products");
 gridweb.getWorkSheets().get(0).getCells().get("A1").putValue(300); 

 gridweb.getWorkSheets().get(0).getCells().get("C1").putValue(gridweb.getVersion()); 
  gridweb.getWorkSheets().get(0).getCells().get("D1").putValue("hello"); 
  gridweb.prepareRender();
out.println(gridweb.getHTMLBody());
  %>
  <br>
 
</BODY>
</html>