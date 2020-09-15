<%@ page language="java" contentType="text/html;charset=UTF-8"
	pageEncoding="UTF-8" isELIgnored="false"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Worksheets - Aspose.Cells Grid Suite Demos</title>
<%@include file="/head.jsp" %>
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
			style : method.value,
			gridwebuniqueid : $("#mycomponent").attr("webuniqueid")
		}, function(data) {
           renderData(data);
		}, "html");
	}

	//页面加载
	$(document).ready(function() {
		
		////loadHead();//
		
		var method = {
			id : "loadSkinsFile"
		};
		doClick(method);
	});
</script>
</head>
<body>
	<div>
		<p>
			Select a <b>style</b> from drop down to see how demo applies different styles to
            the GridWeb Control.
		</p>
	</div>

	<div>
		<table>
			<tr>
				<th>Select another style:</th>
				<td>
					<select id="changeStyle" onchange="doClick(this);">
						<option>=========</option>
						<option value="STANDARD">Standard</option>
						<option value="COLORFUL_1">Colorful1</option>
						<option value="COLORFUL_2">Colorful2</option>
						<option value="PROFESSIONAL_1">Professional1</option>
						<option value="PROFESSIONAL_2">Professional2</option>
						<option value="TRADITIONAL_1">Traditional1</option>
						<option value="TRADITIONAL_2">Traditional2</option>
						<option value="CustomStyle1">CustomStyle1</option>
						<option value="CustomStyle2">CustomStyle2</option>
					</select>
				</td>
			</tr>
		</table>
	</div>

	<div id="gridweb"></div>
</body>
</html>