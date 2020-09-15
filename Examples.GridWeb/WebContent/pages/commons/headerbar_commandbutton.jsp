<%@ page language="java" contentType="text/html;charset=UTF-8" pageEncoding="UTF-8" isELIgnored="false"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<title>Worksheets - Aspose.Cells Grid Suite Demos-header bar </title>
 
<script type="text/javascript">
function renderData(data)
	{    //the default stype part for  gridweb component  is Stylemycomponent
		$("#Stylemycomponent").remove(); 
		//need to render gridweb ,this will trigger reinit of gridweb component
		 //the default   name for  gridweb component is mycomponent
		gridwebinstance.remove("mycomponent") ;
		$("#gridweb").html(data);
	}
	function doClick(method) {
		//alert( $("#showHeaderBar").attr("checked"));
		$.post("FunctionServlet", {
			flag : "headerBarAndCommandButton",
			gridwebuniqueid : $("#mycomponent").attr("webuniqueid"),
			showHeaderBar : $("#showHeaderBar").is(':checked'),
			showSubmitButton : $("#showSubmitButton").is(':checked'),
			showSaveButton : $("#showSaveButton").is(':checked'),
			showUndoButton : $("#showUndoButton").is(':checked'),
			noScrollBars : $("#noScrollBars").is(':checked')
		}, function(data) {
           renderData(data);
		}, "html");
	}
	
	//页面加载
	$(document).ready(function(){
		
		//loadHead();//
		 $(":checkbox").bind('change', function(){ doClick(this); }); 
		var method = {id:"headerBarAndCommandButton"};
		doClick(method);
	});
</script>
</head>
<body>
	<div>
		<p>
			Click <b>Reload</b> to see how demo demonstrates how to hyperlink table cells so
            that browser windows would be opened when clicked and displays data in the GridWeb
            Control.
		</p>
	</div>

	<div>
		<table>
			<tr>
				<td>
					<input type="checkbox" id="showHeaderBar"  checked="checked">Show Header Bar
					<input type="checkbox" id="showSubmitButton"  checked="checked">Show Submit Button
				</td>
			</tr>
			<tr>
				<td>
					<input type="checkbox" id="showSaveButton"  checked="checked">Show Save Button
					<input type="checkbox" id="showUndoButton"  checked="checked">Show Undo Button
					<input type="checkbox" id="noScrollBars" >No Scroll Bars
				</td>
			</tr>
		</table>
	</div>
	
	<div id="gridweb"></div>
</body>
</html>