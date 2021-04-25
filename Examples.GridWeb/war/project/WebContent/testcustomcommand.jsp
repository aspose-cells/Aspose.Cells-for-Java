
<%@ page language="java" contentType="text/html; charset=UTF-8"
	import="com.aspose.gridweb.*" pageEncoding="UTF-8"%>
 
<%
String path = request.getContextPath();
String basePath = request.getScheme()+"://"+request.getServerName()+":"+request.getServerPort()+path+"/";
//GridWebBean gridweb=ExtPage.getInstance().getBean(request);
ExtPage BeanManager=ExtPage.getInstance();
BeanManager.setServlet(request,response);
GridWebBean gridweb=BeanManager.getBean();
//gridweb.setReqRes(request, response);
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<base href="<%=basePath%>">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<script type="text/javascript" src="grid/acw_client/jquery-1.7.2.min.js"></script>
<title>Insert title here</title>
<style type="text/css">
	#mycomponent_CCMD_0 {
    
    width: 15px;
}
</style>
<%
gridweb.importExcelFile("D:/codebase/local/grid/csharp_version/gridweb_old/Aspose.Cells.GridWeb.Demos/file/List.xls");

final GridWorksheet sheet=gridweb.getWorkSheets().get(0);
CustomCommandEventHandler myact=new CustomCommandEventHandler(){
	 public void handleCellEvent(Object sender, String command){
    	 //Identifying a specific button by checking its command
        if (command.equals("MyButton"))
        {
            //Accessing the cells collection of the worksheet that is currently active
           
            //Putting value to "A1" cell
            sheet.getCells().get("A1").putValue("My Custom Command Button is Clicked.");
            System.out.println("My Custom Command Button is Clicked:"+sheet.getName()+" "+sheet.getCells().get("A1").getValue());
        }	
    }
	
};
 
	//gridweb.setACWClientPath("grid/acw_client/");
	gridweb.setACWLanguageFileUrl("grid/acw_client/lang_cn.js");
	out.print(gridweb.getHTMLHead());
%>
</head>
<body>
	hello world
	<%
	 
	gridweb.setWidth(Unit.Pixel(1000));
	gridweb.setHeight(Unit.Pixel(400));
	 GridCell ac=gridweb.getWorkSheets().get(0).getCells().get("D4");
	 GridValidation v=ac.createValidation(GridValidationType.CUSTOM_EXPRESSION, true);  
	  v.setRegEx("/d{4}"); 
	// v.setRegEx("/d{4}-d{2}-d{2}"); 
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
	 CustomCommandButton button = new CustomCommandButton();

	//Setting the command for button
	button.setCommand("MyButton") ;

	//Setting text of the button
	button.setText("MyButton");

	//Setting image URL of the button
	button.setImageUrl("http://webmap2.map.bdstatic.com/wolfman/static/common/images/new/newlogo_bb40ad7.png") ;

	//Adding button to CustomCommandButtons collection of GridWeb
	gridweb.getCustomCommandButtons().add(button);
	
	 CustomCommandButton button2 = new CustomCommandButton();

		//Setting the command for button
		button2.setCommand("MyButton2") ;

		//Setting text of the button
		button2.setText("MyButton2");

		//Setting image URL of the button
		button2.setImageUrl("http://s1.bdstatic.com/r/www/cache/static/soutu/img/camera_new_679c15cc.png") ;
		button2.setWidth(Unit.Pixel(18));
		gridweb.getCustomCommandButtons().add(button2);
		
	gridweb.CustomCommand=myact;
	 
	gridweb.prepareRender();
	out.print(gridweb.getHTMLBody());
	 System.out.println("2222222222222222:"+sheet.getCells().get("A1").getValue());
	//request.getParameter(name);
	//out.print("<br>"+gridweb.get_ClientID()+"<br>"+gridweb.get_CssClass());
%>
</body>
</html>