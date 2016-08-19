<%@ page language="java" contentType="text/html;charset=UTF-8"
	pageEncoding="UTF-8"%>
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
			gridwebuniqueid : $("#mycomponent").attr("webuniqueid"),
			
			row : $("#row").val(),
			column : $("#column").val(),
			rowNumber : $("#rowNumber").val(),
			columnNumber : $("#columnNumber").val()
		}, function(data) {
			$("#gridweb").html(data);
		}, "html");
	}

	$(document).ready(function() {
		
		//loadHead();//
		
		var method = {
			id : "reload"
		};
		doClick(method);
	});
</script>
</head>
<body>

	<div>
		<p>
			Click <b>Create Caption</b> to see how demo freezes/unfreezes panes and displays
        	data in the GridWeb Control.
		</p>
	</div>

	<div>
		<table>
			<tr>
				<th>Row:<input type="text" id="row" value="3" style="width: 20px;">
				Column:<input type="text" id="column" value="3" style="width: 20px;">
				Row Number:<input type="text" id="rowNumber" value="3" style="width: 20px;">
				Column Number:<input type="text" id="columnNumber" value="3" style="width: 20px;">
				</th>
				<td>
					<input id="freezePane" type="button" value="Freeze Pane" onClick="doClick(this);">
					<input id="unfreezePane" type="button" value="Unfreeze Pane" onClick="doClick(this);">
				</td>
			</tr>
		</table>
	</div>

	<div id="gridweb"></div>
</body>
</html>