<%@ page language="java" contentType="text/html;charset=UTF-8" pageEncoding="UTF-8" isELIgnored="false"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<title>Worksheets - Aspose.Cells Grid Suite Demos</title>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<script type="text/javascript">
	function doClick(method) {
		debugger;
		$.post("FunctionServlet", {
			flag : "headerBarAndCommandButton",
			gridwebuniqueid : $("#mycomponent").attr("webuniqueid"),
			showHeaderBar : $("#showHeaderBar").attr("checked")?true:false,
			showSubmitButton : $("#showSubmitButton").attr("checked")?true:false,
			showSaveButton : $("#showSaveButton").attr("checked")?true:false,
			showUndoButton : $("#showUndoButton").attr("checked")?true:false,
			noScrollBars : $("#noScrollBars").attr("checked")?true:false
		}, function(data) {
			$("#gridweb").html(data);
		}, "html");
	}
	
	$(document).ready(function(){
		
		//loadHead();//
		
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
					<input type="checkbox" id="showHeaderBar" onClick="doClick(this);" checked="checked">Show Header Bar
					<input type="checkbox" id="showSubmitButton" onClick="doClick(this);" checked="checked">Show Submit Button
				</td>
			</tr>
			<tr>
				<td>
					<input type="checkbox" id="showSaveButton" onClick="doClick(this);" checked="checked">Show Save Button
					<input type="checkbox" id="showUndoButton" onClick="doClick(this);" checked="checked">Show Undo Button
					<input type="checkbox" id="noScrollBars" onClick="doClick(this);">No Scroll Bars
				</td>
			</tr>
		</table>
	</div>
	
	<div id="gridweb"></div>
</body>
</html>