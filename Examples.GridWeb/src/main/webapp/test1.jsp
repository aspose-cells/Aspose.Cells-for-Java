
<%@ page language="java" contentType="text/html; charset=UTF-8"
	import="com.aspose.gridweb.*" pageEncoding="UTF-8"%>
<jsp:useBean id="gridweb" scope="page"
	class="com.aspose.gridweb.GridWebBean" />
<%
String path = request.getContextPath();
String basePath = request.getScheme()+"://"+request.getServerName()+":"+request.getServerPort()+path+"/";
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<base href="<%=basePath%>">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<script type="text/javascript" src="grid/acw_client/jquery-1.7.2.min.js"></script>
<title>Insert title here</title>
<%
	//gridweb.setACWClientPath("grid/acw_client/");
	gridweb.setACWLanguageFileUrl("grid/acw_client/lang_cn.js");
	out.print(gridweb.getHTMLHead());
%>
</head>
<body>
	hello world
	<%
	gridweb.init(request, response);
	gridweb.setWidth(Unit.Pixel(1000));
	gridweb.setHeight(Unit.Pixel(400));
	gridweb.ImportExcelFile("E:\\project\\cells\\Aspose.Cells.GridWeb.Demos\\file\\list.xls");
	/*
	 //Color c=new Aspose.
	 gridweb.setActiveCellColor(Color.getBurlyWood());
	 Cell ac=gridweb.getWorkSheets().get(0).getCells().get("D4");
	 //gridweb.getWorkSheets()[0].
	 gridweb.setActiveCell(ac);
	 gridweb.setActiveCellBgColor(Color.getAliceBlue());
	 gridweb.setActiveHeaderBgColor(Color.getBrown());
	 gridweb.setActiveHeaderColor(Color.getCornflowerBlue());
	 gridweb.setActiveSheetIndex(2);
	 GridTableItemStyle itemstyle=new GridTableItemStyle();
	 Color cl=Color.getChocolate();
	 Color cb=Color.getAliceBlue();

	 itemstyle.set_BackColor(cl);
	 gridweb.setActiveTabStyle(itemstyle);

	 GridTableStyle itemstylebottom=new GridTableStyle();
	 itemstylebottom.set_BackColor(Color.getCyan());
	 itemstylebottom.set_BorderColor(Color.getDarkBlue());
	 //gridweb.setBottomTableStyle(itemstylebottom);

	 //gridweb.setDefaultFontSize(23); 
	 //gridweb.setDefaultFontName('Arail2'); 
	 //gridweb.setDefaultGridLineColor(Color.get_DarkBlue());

	 //gridweb.setDisplayCellTip(true);
	 //gridweb.setEditMode(true);

	
	 gridweb.set_BackColor(cb);
	 //gridweb.setBackColor(cb);

	 //no effect
	 gridweb.set_BorderColor(Color.getGold());
	 */

	gridweb.prepareRender();
	out.print(gridweb.getHTMLBody());
	//request.getParameter(name);
	//out.print("<br>"+gridweb.get_ClientID()+"<br>"+gridweb.get_CssClass());
%>
</body>
</html>