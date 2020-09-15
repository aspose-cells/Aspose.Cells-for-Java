<%@ page language="java" contentType="text/html; charset=UTF-8"
 import="com.aspose.gridweb.*,java.io.*" pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
 
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>Insert title here</title>

<%
License al=new License();
//al.setLicense("D:\\grid_project\\ZZZZZZ_release_version\\Licenses\\Aspose.Total.Java20151205.lic");

ExtPage BeanManager=ExtPage.getInstance();
BeanManager.setServlet(request, response);
BeanManager.setPageaction( request.getContextPath()+"/GridWebServlet"); 
BeanManager.setPageajaxcallpath(request.getContextPath()+"/GridWebServlet?acw_ajax_call=true"); 
GridWebBean gridweb=BeanManager.getBean();
//gridweb.setACWClientPath("../grid/acw_client/");
//out.println(gridweb.getHTMLHead());
%>
<%@include file="/head.jsp" %>
</head>
<BODY>
 

<%
gridweb.setEnableAsync(true);
gridweb.setRenderHiddenRow(true);
gridweb.setWidth(Unit.Pixel(1200));
gridweb.setHeight(Unit.Pixel(600));
//String test= (Class.forName("com.aspose.gridweb.GridWebBean").getProtectionDomain().getCodeSource().getLocation()).getPath() ; 
//InputStream f = new FileInputStream(application.getRealPath("/")+"/file/Math.xls");
 //File afile=new File(application.getRealPath("/")+"/file/Math.xls");
 //String file="C:\\Users\\peter\\Desktop\\ctest\\test17canbe\\works.xlsx";
 String file="C:\\Users\\peter\\Desktop\\ctest\\test17canbe\\2.xlsx";
// file="C:\\Users\\peter\\Desktop\\ctest\\imagenotshow_41.xlsm";
gridweb.importExcelFile(file);
 gridweb.getWorkSheets().get(0).setColumnCaption(1, "Price"); 
 gridweb.getWorkSheets().get(0).setColumnHeaderToolTip(1, "Unit Price of Products");
gridweb.prepareRender();
out.println(gridweb.getHTMLBody());
  %>
  <br>
 
</BODY>
</html>