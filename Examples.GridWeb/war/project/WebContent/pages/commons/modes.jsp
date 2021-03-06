<%@ page language="java" contentType="text/html;charset=UTF-8" pageEncoding="UTF-8" isELIgnored="false"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<title>Worksheets - Aspose.Cells Grid Suite Demos</title>
 
<script type="text/javascript">
	var editMode = false;
	function renderData(data)
	{    //the default stype part for  gridweb component  is Stylemycomponent
		$("#Stylemycomponent").remove(); 
		//need to render gridweb ,this will trigger reinit of gridweb component
		 //the default   name for  gridweb component is mycomponent
		gridwebinstance.remove("mycomponent") ;
		$("#gridweb").html(data);
	}
	function doClick(method) {
		if($("#editMode:checked").val()){
			editMode = true;
			$("#rowspan").show();
			$("#colspan").show();
		}else{
			editMode = false;
			$("#rowspan").hide();
			$("#colspan").hide();
		}
		$.post("FunctionServlet", {
			editMode : editMode,
			flag : "editMode",//method
			gridwebuniqueid : $("#mycomponent").attr("webuniqueid")
		}, function(data) {
           renderData(data);
		}, "html");
	}
	
	
	function doRowEditClick(method) {
		if($("#editModeRow:checked").val()){
			editMode = true;
		 
		}else{
			editMode = false;
			 
		}
		$.post("FunctionServlet", {
			row:$("#row").val(),
			editMode : editMode,
			flag : "setRowReadonly",//method
			gridwebuniqueid : $("#mycomponent").attr("webuniqueid")
		}, function(data) {
           renderData(data);
		}, "html");
	}
	
	function doColEditClick(method) {
		if($("#editModeCol:checked").val()){
			editMode = true;
		 
		}else{
			editMode = false;
			 
		}
		$.post("FunctionServlet", {
			col:$("#column").val(),
			editMode : editMode,
			flag : "setColReadonly",//method
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
			Click <b>Enable Editing</b> to see how demo toggles editable / read-only mode.
			<br>
			also we can specify to row level and column level.
		</p>
	</div>

	<div>
		<table>
			<tr>
				<td>
					<input type="checkbox" id="editMode" onClick="doClick(this);">Enable editing
					
				 
				</td>
			</tr>
			<tr>
				<td>
					<span id="rowspan">
					set row <input type="text" id="row" > <input type="checkbox" id="editModeRow" onClick="doRowEditClick(this);">Enable/Disable row read-only
					</span>
				</td>
			</tr>
			<tr>
				<td>
	                <span id="colspan">
					set column <input type="text" id="column" > <input type="checkbox" id="editModeCol" onClick="doColEditClick(this);">Enable/Disable column read-only
					</span>
				</td>
			</tr>
		</table>
	</div>
	
	<div id="gridweb"></div>
</body>
</html>