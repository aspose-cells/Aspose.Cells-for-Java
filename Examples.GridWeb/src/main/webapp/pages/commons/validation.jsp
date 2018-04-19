<%@ page language="java" contentType="text/html;charset=UTF-8" pageEncoding="UTF-8" isELIgnored="false"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<title>Worksheets - Aspose.Cells Grid Suite Demos-validation</title>
 <script type="text/javascript">
	var validation = true;
function renderData(data)
	{    //the default stype part for  gridweb component  is Stylemycomponent
		$("#Stylemycomponent").remove(); 
		//need to render gridweb ,this will trigger reinit of gridweb component
		 //the default   name for  gridweb component is mycomponent
		gridwebinstance.remove("mycomponent") ;
		$("#gridweb").html(data);
	}
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
           renderData(data);
		}, "html");
	}
	
	//页面加载
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
	  function ValidationErrorClientFunctionCallback(cell, msg) {
        //Showing an alert message where "this" refers to GridWeb
        //  alert(this.id + msg);
         $("#errmsg").text("row:" + this.getCellRow(cell) + ",col:" + this.getCellColumn(cell)+" msg:"+msg);
       // console.log();
        // 
        var who = this;
        //async raise alert 
        
        //restore to valid value 
        who.setValid(cell);
        var key = this.acttab + "_" + this.getCellRow(cell) + "_" + this.getCellColumn(cell);
        lastselectvalue = localvalue[key];
        setInnerText(cell.children[0], lastselectvalue);
        // this.setCellValueByCell(cell, lastselectvalue); 
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
	<span id="errmsg"></span>
	<div id="gridweb"></div>
</body>
</html>