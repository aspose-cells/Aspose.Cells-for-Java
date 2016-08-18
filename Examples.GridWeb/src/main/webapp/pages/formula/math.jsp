<%@ page language="java" contentType="text/html;charset=UTF-8" pageEncoding="UTF-8" isELIgnored="false"%>
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
	
	$(document).ready(function(){
		
		//loadHead();//
		
		var method = {id:"loadMathFile"};
		doClick(method);
	});
</script>
</head>
<body>
	<div>
		<p>
            This demo loads an existing file into an empty WebWorksheet to demonstrate how GridWeb
            applies typical <b>Math</b> formulas to grid cells and calculates formula values.
            Click <b>Reload</b> to reload initial data for the grid. You can also Save/Open
            the output in <b>XLS</b>Format by clicking the Save Button on GridWeb Control Bar.
        </p>
	</div>

	<div>
		<table>
			<tr>
				<th>Reload Data:</th>
				<td><input id="loadMathFile" type="button" value="Reload" onClick="doClick(this);"></td>
			</tr>
		</table>
	</div>
	
	<div id="gridweb"></div>
</body>
</html>