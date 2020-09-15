<%@ page language="java" contentType="text/html;charset=UTF-8" pageEncoding="UTF-8" isELIgnored="false"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<title>Update Font Settings From Client Side - Aspose.Cells Grid Suite Demos</title>
 <script type="text/javascript">
 function makeFontItalic(){
     gridwebinstance.getByIndex(0).rangeupdate(updateCellFontStyle, "i");
 }
 function makeFontBold(){
     gridwebinstance.getByIndex(0).rangeupdate(updateCellFontStyle, "b");
 }
 function changeFontSize(){
     gridwebinstance.getByIndex(0).rangeupdate(updateCellFontSize, "5pt");
 }
 function changeFontFamily(){
     gridwebinstance.getByIndex(0).rangeupdate(updateCellFontName, "Corbel Light");
 }
 function changeFontColor(){
     gridwebinstance.getByIndex(0).rangeupdate(updateCellFontColor, "green");
 }
 function changeFontBackgroundColor(){
     gridwebinstance.getByIndex(0).rangeupdate(updateCellBackGroundColor, "yellow");
 }
 function renderData(data)
	{    //the default stype part for  gridweb component  is Stylemycomponent
		$("#Stylemycomponent").remove(); 
		//need to render gridweb ,this will trigger reinit of gridweb component
		 //the default   name for  gridweb component is mycomponent
		gridwebinstance.remove("mycomponent") ;
		$("#gridweb").html(data);
	}
	function doClick(method) {
		$.post("FeatureServlet", {
			flag : method.id,
			gridwebuniqueid : $("#mycomponent").attr("webuniqueid")
		}, function(data) {
           renderData(data);
		}, "html");
	}
	
	//页面加载
	$(document).ready(function(){
		
		//loadHead();//
		
		var method = {id:"updateFontFromClientSide"};
		doClick(method);
	});
</script>
</head>
<body>
	<div>
		<p>
			This demo shows updating font settings from client side..
		</p>
	</div>

	<div>
		<table>
			<tr>
				<th>Reload Data:</th>
				<td><input id=loadSampleFile type="button" value="Reload" onClick="doClick(this);"></td>
			</tr>
		</table>
	</div>
	<div>
        <input type="button" value="Make Font Italic" onclick="makeFontItalic()"/><br/><br/>
        <input type="button" value="Make Font Bold" onclick="makeFontBold()"/><br/><br/>
        <input type="button" value="Change Font Size" onclick="changeFontSize()"/><br/><br/>
        <input type="button" value="Change Font Family" onclick="changeFontFamily()"/><br/><br/>
        <input type="button" value="Change Font Color" onclick="changeFontColor()"/><br/><br/>
        <input type="button" value="Change Font Background Color" onclick="changeFontBackgroundColor()"/><br/><br/><br/>

    </div>
	
	<div id="gridweb"></div>
</body>
</html>