<%@ page language="java" contentType="text/html;charset=UTF-8" pageEncoding="UTF-8" isELIgnored="false"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<title>Worksheets - Aspose.Cells Grid Suite Demos-chart</title>
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
		
		var method = {id:"loadGroupRowCol"};
		doClick(method);
	});
</script>
</head>
<body>
	<div>
		<p>
            This demo loads an existing file into an empty WebWorksheet to demonstrate how GridWeb
            display various of charts.Try edit the cell value,then the related charts image will refresh automatically.
            User can also call method to refresh all the chart image during some event.<b>Check</b> <a href="./pages/commons/chartsubmit.jsp" target="_blank"><b>here </b></a> for another demo. 
            Click <b>Reload</b> to reload initial data for the grid. You can also Save/Open
            the output in <b>XLS</b>Format by clicking the Save Button on GridWeb Control Bar.
        </p>
	</div>

	<div>
		<table>
			<tr>
				<th>Reload Data:</th>
				<td><input id="loadGroupRowCol" type="button" value="Reload" onClick="doClick(this);"></td>
			</tr>
		</table>
	</div>
	
	<div id="gridweb"></div>
</body>
</html>