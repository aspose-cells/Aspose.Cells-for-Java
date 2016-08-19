<%@ page language="java" contentType="text/html;charset=UTF-8" pageEncoding="UTF-8" isELIgnored="false"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<title>Worksheets - Aspose.Cells Grid Suite Demos</title>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<script type="text/javascript">
	function doClick(method) {
		$.post("FormatServlet", {
			flag : method.id,
			gridwebuniqueid : $("#mycomponent").attr("webuniqueid"),
			value:$("#value").val(),
			format:$("#format").val()
		}, function(data) {
			$("#gridweb").html(data);
		}, "html");
	}
	
	$(document).ready(function(){
		
		//loadHead();//
		
		var method = {id:"loadCustomFormatFile"};
		doClick(method);
	});
</script>
</head>
<body>
	<div>
		<p>
			Pick a date format from the list, enter a value (text) and click <b>Submit</b> to
            see how demo applies custom date format to a grid cell and displays your value in
            it.
		</p>
	</div>

	<div>
		<table>
			<tr>
				<th>Reload Data:</th>
				<td><input id="loadCustomFormatFile" type="button" value="Reload" onClick="doClick(this);"></td>
			</tr>
			<tr>
				<th>Custom Format:</th>
				<td><input id="format" type="text"></td>
				<th>Input Value:</th>
				<td>
					<input id="value" type="text">
					<input id="customFormat" type="button" value="Submit" onClick="doClick(this);">
				</td>
			</tr>
		</table>
	</div>
	
	<div id="gridweb"></div>
</body>
</html>