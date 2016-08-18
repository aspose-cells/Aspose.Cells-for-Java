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
		$.post("WebCellsServlet", {
			flag : method.id,
			gridwebuniqueid : $("#mycomponent").attr("webuniqueid"),
			columnIndex : $("#columnIndex").val(),
			rowIndex : $("#rowIndex").val(),
			startRow : $("#startRow").val(),
			startColumn : $("#startColumn").val(),
			rowNumber : $("#rowNumber").val(),
			columnNumber : $("#columnNumber").val(),
			startRow_c : $("#startRow_c").val(),
			startColumn_c : $("#startColumn_c").val(),
			comment : $("#comment").val()
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
				Click <b>Reload</b> to reload data from data source. Click
			<ul>
				<li><b>Insert Column</b> to see how demo inserts a column</li>
				<li><b>Insert Row</b> to see how demo inserts a row</li>
				<li><b>Delete Row</b> to see how demo deletes a row</li>
				<li><b>Delete Column</b> to see how demo deletes a column</li>
			</ul>
		
	</div>

	<div>
		<table>
			<tr>
				<th>Reload Data:</th>
				<td><input id="reload" type="button" value="Reload" onClick="doClick(this);"></td>
			</tr>
			<tr>
				<th>Insert/Delete Column:</th>
				<td>ColumnIndex:<input type="text" id="columnIndex" value="2" style="width: 20px;">
					<input id="inserColumn" type="button" value="Insert Column" onClick="doClick(this);">
					<input id="deleteColumn" type="button" value="Delete Column" onClick="doClick(this);">
				</td>
			</tr>
			<tr>
				<th>Insert/Delete Row:</th>
				<td>RowIndex:<input type="text" id="rowIndex" value="2" style="width: 20px;">
					<input id="insertRow" type="button" value="Insert Row" onClick="doClick(this);">
					<input id="deleteRow" type="button" value="Delete Row" onClick="doClick(this);">
				</td>
			</tr>
			<tr>
				<th>Merge Cells:</th>
				<td>StartRow:<input type="text" id="startRow" value="0" style="width: 20px;">
				StartColumn:<input type="text" id="startColumn" value="0" style="width: 20px;">
				RowNumber:<input type="text" id="rowNumber" value="3" style="width: 20px;">
				ColumnNumber:<input type="text" id="columnNumber" value="2" style="width: 20px;">
				<input id="mergeCells" type="button" value="Merge Cells" onClick="doClick(this);"></td>
			</tr>
			<tr>
				<th>Add/Remove Comment: </th>
				<td>StartRow:<input type="text" id="startRow_c" value="1" style="width: 20px;">
				StartColumn:<input type="text" id="startColumn_c" value="1" style="width: 20px;">
				Comment:<input type="text" id="comment" value="This is my comment.">
				<input id="addComment" type="button" value="Add Comment" onClick="doClick(this);">
				<input id="removeComment" type="button" value="Remove Comment" onClick="doClick(this);"></td>
			</tr>
		</table>
	</div>

	<div id="gridweb"></div>
</body>
</html>