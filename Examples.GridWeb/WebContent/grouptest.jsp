<%@ page language="java" contentType="text/html; charset=UTF-8"
 import="com.aspose.gridweb.*" pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
 
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>Insert title here</title>
<%
License al=new License();
//al.setLicense("D:\\grid_project\\ZZZZZZ_release_version\\Licenses\\Aspose.Total.Java20151205.lic");

ExtPage BeanManager=ExtPage.getInstance();
//GridWebBean gridweb=BeanManager.getBean(request);
BeanManager.setServlet(request,response);
GridWebBean gridweb=BeanManager.getBean();
//gridweb.setACWClientPath("../grid/acw_client/");
out.println(gridweb.getHTMLHead());
%>
</head>
<BODY>
<%
//response.getWriter().write("This is response<BR>");
    // This scriptlet declares and initializes "date"
  //  System.out.println( "Evaluating date now" );
  //  java.util.Date date = new java.util.Date();
%>
Hello!  The time is now
<%
  //  out.println( date );
  //  out.println( "<BR>Your machine's address is " );
   // out.println( request.getRemoteHost());
    
    
%>

<%
gridweb.setID("activities45150");
//gridweb.setReqRes(request, response);
//String test= (Class.forName("com.aspose.gridweb.GridWebBean").getProtectionDomain().getCodeSource().getLocation()).getPath() ; 
String absPath=new java.io.File(application.getRealPath(request.getRequestURI())).getParent(); 
//报java.lang.NullPointerException com.aspose.gridweb.zagn.c(Unknown Source)
//gridweb.importExcelFile(absPath+"/../file/0 Web - W010 - Liasse Locale 99-02_0.xls");
// gridweb.importExcelFile(absPath+"/../file/W030 - Liasse imps_119.xls");
//gridweb.importExcelFile(absPath+"/../file/W030 - Liasse imps_119.xls");
gridweb.importExcelFile(absPath+"/../file/0 Web - W010 - Liasse Locale 99-02_0.xls");
//gridweb.importExcelFile("D:\\grid_project\\temp\\gridweb(java)0822\\sample\\WebRoot\\file\\list.xls");
//Setting the background color of the header bars 
//GridTableItemStyle headstyle=gridweb.getHeaderBarStyle();
//headstyle.setBackColor(Color.getAliceBlue());
//headstyle.setForeColor(Color.getRed());
//gridweb.setHeaderBarStyle(headstyle);
//Setting the background color of tabs to Yellow 
//GridTableItemStyle tabstyle=gridweb.getTabStyle();
//tabstyle.setBackColor(Color.getCornsilk()); 
//tabstyle.setForeColor(Color.getLimeGreen());
//gridweb.setTabStyle(tabstyle);
//gridweb.getActiveSheet().getCells().get("A1").putValue("A1");
gridweb.setWidth(Unit.Pixel(1200));
gridweb.setHeight(Unit.Pixel(600));
gridweb.setEnableAsync(true);
gridweb.prepareRender();
out.println(gridweb.getHTMLBody());
 out.println(application.getRealPath(request.getContextPath()));
  out.println("<br>这路径不对头？？？怎么两个gridwebdemoV8.6.3 路径<br>");
  out.println(request.getContextPath());
   out.println("<br>");
  out.println(request.getRequestURI());
  %>
 
</BODY>
</html>