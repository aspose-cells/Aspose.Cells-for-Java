<%@ page language="java" contentType="text/html;charset=UTF-8" pageEncoding="UTF-8" isELIgnored="false"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<title>Worksheets - Aspose.Cells Grid Suite Demos</title>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<script type="text/javascript">
	var validation = true;
	function doClick(method) {
		if($("#validation:checked").val()){
			validation = true;
		}else{
			validation = false;
		}
		$.post("FunctionServlet", {
			validation : validation,
			flag : method.id,
			gridwebuniqueid : $("#mycomponent").attr("webuniqueid")
		}, function(data) {
			$("#gridweb").html(data);
		}, "html");
	}
	
	$(document).ready(function(){
		
		//loadHead();//
		
		var method = {id:"validation"};
		doClick(method);
	});
	
	function myvalidation1(source, value)
	{
		if (Number(value) > 10000)
			return true;
		else
			return false;
	}
</script>
</head>
<body>
	<div>
		<p>
			Click <b>Reload</b> to see how demo reloads data and applies validation rules so
            that invalid (not matching certain RegExp) values could not be entered in the GridWeb
            Control.
		</p>
	</div>

	<div>
		<table>
			<tr>
				<td>
					Input Entry Protection/Validation: 
					<input type="button" id="validation" onClick="doClick(this);" value="Reload Data">
				</td>
			</tr>
			<tr>
				<td>
					<input type="checkbox" id="validation" onClick="doClick(this);" checked="checked">Enable Force Validation
				</td>
			</tr>
		</table>
	</div>
	
	<div id="gridweb"></div>
</body>
</html>