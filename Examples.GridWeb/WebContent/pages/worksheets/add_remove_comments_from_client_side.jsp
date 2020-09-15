<%@ page language="java" contentType="text/html;charset=UTF-8" pageEncoding="UTF-8" isELIgnored="false"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<title>Add/Remove Hyperlinks From Client Side - Aspose.Cells Grid Suite Demos</title>
 <script type="text/javascript">
 function addLink() {
     var linkinfo={};
     linkinfo.url="http://www.aspose.com",
     linkinfo.text="Link to Aspose Webbsite",
     gridwebinstance.getByIndex(0).rangeupdate(addCelllink,linkinfo);
 }
 function removeLink() {
     gridwebinstance.getByIndex(0).rangeupdate(delCelllink);
 }
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
		
		var method = {id:"addRemoveHyperlinkFromClientSide"};
		doClick(method);
	});
</script>
</head>
<body>
	<div>
		<p>
			This demo shows adding and deleting hyperlinks from client side.
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
	<div>
        <input type="button" value="Add Hyperlink" onclick="addLink()"/><br/><br/>
        <input type="button" value="Remove Hyperlink" onclick="removeLink()"/><br/><br/>
    </div>
	
	<div id="gridweb"></div>
</body>
</html>