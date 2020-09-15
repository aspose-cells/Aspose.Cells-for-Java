<%@ page language="java" contentType="text/html; charset=UTF-8" import="com.aspose.gridweb.*"
    pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<script type="text/javascript">
 function buttonclick(a)
 {
	  document.getElementById("opid").value=a;
	   
	  
	  document.getElementById("myform").submit();
 }
 function showcommentindiv(s1,s2)
 {document.getElementById("CommentDiv").innerHTML= s1;
 document.getElementById("CommentSpan").innerHTML= s2;
 }
</script>
<title>gridcomment server api usage</title>
<%
ExtPage BeanManager=ExtPage.getInstance();
BeanManager.setServlet(request, response);
GridWebBean gridweb=BeanManager.getBean();
 
%>
</head>
<body>
 move mouse to C7 you can see there is tool tips for the cell comment.click create will add comment to C6,click delete will delte comment in C7
<br>
here is the comment after click get button:
<br>
<span id="CommentSpan">
</span>
<div id="CommentDiv">
</div>
 <br>

 

<form  id="myform"  >
 
<input type="hidden" name="op"  id="opid"  value="">
<input type="button" value="insert" onclick="buttonclick(this.value)">
<input type="button" value="delete" onclick="buttonclick(this.value)">
<input type="button" value="get" onclick="buttonclick(this.value)">

</form>
<br>
<!-- current operation is:<%=request.getParameter("op") %> -->
<br>
<% 
gridweb.importExcelFile(application.getRealPath("/")+"/file/commentscrud.xls");
 
String op=request.getParameter("op");
String commentnote="";
String commenthtml="";
if(op!=null)
{

if(op.equals("insert"))
{
GridWorksheet sheet=gridweb.getWorkSheets().get(0);
GridCommentCollection gcc=sheet.getComments();
int id=gcc.add("C6");
GridComment gc=gcc.get(id);
gc.setNote("hello comment in c6");
}else if(op.equals("delete"))
{
GridWorksheet sheet=gridweb.getWorkSheets().get(0);
sheet.getCells().get("C7").removeComment();
}else{
GridWorksheet sheet=gridweb.getWorkSheets().get(0);
GridComment gc=sheet.getCells().get("C7").getComment();
if(gc!=null)
	{//notece escape \n ,other wise below html string will be invalid
	commentnote=gc.getNote().replaceAll("\n","");
	commenthtml=gc.getHtmlNote().replaceAll("\n","");
	}
}

}
 
 
gridweb.prepareRender();

out.print(gridweb.getHTMLBody());
 
//System.out.println(gridweb.getPresetStyle()+",has set default"+",get enable page expected false:"+gridweb.getEnablePaging());
%>

 <script type="text/javascript">
 var note='<%=commentnote %>';
  var notehtml='<%=commenthtml %>';
 
 showcommentindiv(note,notehtml);
</script>
 
 
<br>
</body>
</html>