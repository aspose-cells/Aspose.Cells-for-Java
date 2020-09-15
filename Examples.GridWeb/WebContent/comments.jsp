<%@ page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"
    pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript">
 
</script>
<title> comments demo</title>
<%
ExtPage BeanManager=ExtPage.getInstance();
//GridWebBean gridweb=BeanManager.getBean(request);
BeanManager.setServlet(request,response);
GridWebBean gridweb=BeanManager.getBean();
 
%>
</head>
<body>
 move mouse to C6 and B5 ,B7 you can see there is tool tips for the cell comment.
 <%out.print(request.getParameter("Name"));%>
 <%out.print(request.getParameter("actaaa"));%>
<% 
//gridweb.setReqRes(request, response);
gridweb.importExcelFile(application.getRealPath("/")+"/file/comments.xls");
 
GridWorksheet sheet=gridweb.getWorkSheets().get(0);
GridCommentCollection gcc=sheet.getComments();
int id=gcc.add("C6");
GridComment gc=gcc.get(id);
gc.setNote("hello comment in c6");

 
 
gridweb.prepareRender();

out.print(gridweb.getHTMLBody());
//System.out.println(gridweb.getPresetStyle()+",has set default"+",get enable page expected false:"+gridweb.getEnablePaging());
%>

</body>
</html>