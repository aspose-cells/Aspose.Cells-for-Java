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
		$.post("FunctionServlet", {
			flag : method.id,
			gridwebuniqueid : $("#mycomponent").attr("webuniqueid")
		}, function(data) {
			$("#gridweb").html(data);
		}, "html");
	}

	$(document).ready(function() {
		
		//loadHead();//
		
		var method = {
			id : "autoFilter"
		};
		doClick(method);
	});
</script>
</head>
<body>
	<div>
		<p>
			Click <b>Submit Changes</b> ("V") button to see how demo displays auto-filter in
            <b>Row 5</b> enabling user to filter grid contents.
		</p>
	</div>

	<div>
		<table>
			<tr>
				<td></td>
			</tr>
		</table>
	</div>

	<div id="gridweb"></div>
</body>
</html>