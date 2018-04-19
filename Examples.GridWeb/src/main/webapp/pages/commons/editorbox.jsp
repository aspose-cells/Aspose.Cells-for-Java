<%@ page language="java" contentType="text/html;charset=UTF-8" pageEncoding="UTF-8" isELIgnored="false"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<title>Worksheets - Aspose.Cells Grid Suite Demos-show editor box</title>
 
<script type="text/javascript">
	var flag = false;
	function renderData(data)
	{    //the default stype part for  gridweb component  is Stylemycomponent
		$("#Stylemycomponent").remove(); 
		//need to render gridweb ,this will trigger reinit of gridweb component
		 //the default   name for  gridweb component is mycomponent
		gridwebinstance.remove("mycomponent") ;
		$("#gridweb").html(data);
	}
	 
	
	 
	
	function ShowEditorClick(method) {
		if($("#showEditorBox:checked").val()){
		  flag = true;
		 
		}else{
		  flag = false;
			 
		}
		$.post("FunctionServlet", {
			col:$("#column").val(),
			isshow : flag,
			flag : "showEditor",//method
			gridwebuniqueid : $("#mycomponent").attr("webuniqueid")
		}, function(data) {
			renderData(data);
		}, "html");
	}
	
	
	
	//页面加载
	$(document).ready(function(){
		
	 
		
		var method = {id:"reload"};
		doClick(method);
	});
</script>
</head>
<body>
	<div>
		<p>
			Click <b>Enable/Disable show editor box
                </b> to see how demo show editor box .
			<br>
			 
		</p>
	</div>

	<div>
		<table>
			 
			<tr>
				<td>
					<span id="rowspan">
					  <input type="checkbox" id="showEditorBox" onClick="ShowEditorClick(this);">Enable/Disable show editor box
					</span>
				</td>
			</tr>
			 
		</table>
	</div>
	
	<div id="gridweb"></div>
</body>
</html>