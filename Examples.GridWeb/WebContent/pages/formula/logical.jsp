<%@ page language="java" contentType="text/html;charset=UTF-8" pageEncoding="UTF-8" isELIgnored="false"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<title>Worksheets - Aspose.Cells Grid Suite Demos</title>
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
		$.post("FeatureServlet", {
			flag : method.id,
			gridwebuniqueid : $("#mycomponent").attr("webuniqueid")
		}, function(data) {
			renderData(data);
		}, "html");
	}
	
	//页面加载
	$(document).ready(function(){
		
		//loadHead();//
		
		var method = {id:"loadLogicalFile"};
		doClick(method);
	});
</script>
</head>
<body>
	<div>
		<p>
            This demo loads an existing file into an empty WebWorksheet to demonstrate how GridWeb
            applies typical <b>Logical</b> formulas to grid cells and calculates formula values.
            Click <b>Reload</b> to reload initial data for the grid. You can also Save/Open
            the output in <b>XLS</b>Format by clicking the Save Button on GridWeb Control Bar.
        </p>
	</div>

	<div>
		<table>
			<tr>
				<th>Reload Data:</th>
				<td><input id="loadLogicalFile" type="button" value="Reload" onClick="doClick(this);"></td>
			</tr>
		</table>
	</div>
	
	<div id="gridweb"></div>
</body>
</html>