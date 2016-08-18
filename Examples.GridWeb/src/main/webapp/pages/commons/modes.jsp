<%@ page language="java" contentType="text/html;charset=UTF-8" pageEncoding="UTF-8" isELIgnored="false"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<title>Worksheets - Aspose.Cells Grid Suite Demos</title>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<script type="text/javascript">
	var editMode = false;
	function doClick(method) {
		if($("#editMode:checked").val()){
			editMode = true;
		}else{
			editMode = false;
		}
		$.post("FunctionServlet", {
			editMode : editMode,
			flag : "editMode",//method
			gridwebuniqueid : $("#mycomponent").attr("webuniqueid")
		}, function(data) {
			$("#gridweb").html(data);
		}, "html");
	}
	
	$(document).ready(function(){
		
	 
		
		var method = {id:"reload"};
		doClick(method);
	});
</script>
</head>
<body>
	<div>
		<p>
			Click <b>Enable Editing</b> to see how demo toggles editable / read-only mode and
            displays data in the GridWeb Control.
		</p>
	</div>

	<div>
		<table>
			<tr>
				<td>
					<input type="checkbox" id="editMode" onClick="doClick(this);">Enable editing
				</td>
			</tr>
		</table>
	</div>
	
	<div id="gridweb"></div>
</body>
</html>