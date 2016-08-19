<%@ page language="java" contentType="text/html;charset=UTF-8"
	pageEncoding="UTF-8" isELIgnored="false"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<title>Worksheets - Aspose.Cells Grid Suite Demos</title>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>

<script type="text/javascript">
	function doClick(method) {
		$.post("FeatureServlet", {
			flag : method.id,
			gridwebuniqueid : $("#mycomponent").attr("webuniqueid")
		}, function(data) {
			$("#gridweb").html(data);
		}, "html");
	}

	$(document).ready(function() {
		
		//loadHead();//
		
		var method = {
			id : "events"
		};
		doClick(method);
	});
	
	function showMsg(message){
		alert(message);
	}
</script>
</head>
<body>
	<div>
		<p>
			Click <b>Create Caption</b> to see how demo demonstrates handling events (change the page index/click sheet tab/click submit button/click undo button) in the GridWeb Control,event will be show in A1.
		</p>
	</div>

	<div>
		<table>
			<tr>
				<td><input id="events" type="button" value="Reload Data" 
					onClick="doClick(this);"></td>
			</tr>
		</table>
	</div>

	<div id="gridweb"></div>
</body>
</html>