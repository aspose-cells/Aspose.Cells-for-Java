<%@ page language="java" contentType="text/html;charset=UTF-8" pageEncoding="UTF-8" isELIgnored="false"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<title>Show Add Button - Aspose.Cells Grid Suite Demos</title>
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
		
		var method = {id:"showAddButton"};
		doClick(method);
	});
</script>
</head>
<body>
	<div>
		<p>
			This demo show adding and deleting worksheets by adding the add and delete buttons to the toolbar.
		</p>
	</div>

	<div>
		<table>
			<tr>
				<th>Reload Data:</th>
				<td><input id=loadSampleFile type="button" value="Reload" onClick="doClick(this);"></td>
			</tr>
		</table>
	</div>
	
	<div id="gridweb"></div>
</body>
</html>