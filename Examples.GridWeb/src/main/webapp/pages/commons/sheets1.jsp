<%@ page language="java" contentType="text/html;charset=UTF-8" pageEncoding="UTF-8"%>
<%
String path = request.getContextPath();
String basePath = request.getScheme()+"://"+request.getServerName()+":"+request.getServerPort()+path+"/";
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<base href="<%=basePath%>">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<script type="text/javascript" src="grid\acw_client\jquery-2.1.4.min.js"></script>
 <script type="text/javascript" src="grid\acw_client\acwmain.js"></script>
 <link href="grid/acw_client/menu.css" rel="stylesheet" type="text/css">
 <style>span.acwxc {overflow:hidden; border:none; display:block; white-space: pre;}</style>
 <style>span.rotation90 {width:100%; height:100%;border:none; -webkit-transform: rotate(-90deg);-moz-transform: rotate(-90deg);filter:progid:DXImageTransform.Microsoft.BasicImage(rotation=3);display:block}</style>
 <style>span.rotation-90 {filter:progid:DXImageTransform.Microsoft.BasicImage(rotation=1);width:100%; height:100%;border:none; -webkit-transform: rotate(90deg);-moz-transform: rotate(90deg);display:block}</style>
 <style>span.wrap {white-space: pre-wrap; white-space: -moz-pre-wrap; white-space: -pre-wrap; white-space: -o-pre-wrap; word-wrap: break-word; -ms-word-break: break-all; }</style>
<title>Worksheets - Aspose.Cells Grid Suite Demos</title>
<script type="text/javascript">
	function doClick(method) {
		alert($("#mycomponent").attr("webuniqueid"));
		$.post("TestGridWebServlet", {
			flag : method.id,gridwebuniqueid:$("#mycomponent").attr("webuniqueid")
		}, function(data) {
			$("#gridweb").html(data);
		  //	if($("#mycomponent").get(0)!=null)
		   // $("#mycomponent").get(0).gridajaxcalltest(data);
			 
				 
	 
			
		}, "html");
	}
	
	$(document).ready(function(){
		var method = {id:"reload"};
		doClick(method);
	});
</script>

</head>
<body>

	<div>
		<p>
			Click <b>Add</b> to see how demo adds a worksheet, <b>Add Copy</b> to
			add a copy of a worksheet and <b>Remove Active Sheet</b> to remove
			active worksheet. Click <b>Reload</b> to re-bind data from data
			source and display data in the GridWeb Control.
		</p>
	</div>

	<div>
		<table>
			<tr>
				<th>Reload Data:</th>
				<td><input id="reload" type="button" value="Reload" onClick="doClick(this);"></td>
			</tr>
			<tr>
				<th>Add Sheet:</th>
				<td><input id="add" type="button" value="Add" onClick="doClick(this);"></td>
			</tr>
			<tr>
				<th>Copy Sheet:</th>
				<td><input id="copy" type="button" value="Add Copy" onClick="doClick(this);"></td>
			</tr>
			<tr>
				<th>Remove Sheet:</th>
				<td><input id="remove" type="button"
					value="Remove Active Sheet" onClick="doClick(this);"></td>
			</tr>
			<tr>
				<th>change style,style1:</th>
				<td><input id="style1" type="button"
					value="style1" onClick="doClick(this);"></td>
			</tr>
			<tr>
				<th>change style,style2:</th>
				<td><input id="style2" type="button"
					value="style2" onClick="doClick(this);"></td>
			</tr>
			<tr>
				<th>change style,custstyle1:</th>
				<td><input id="custstyle1" type="button"
					value="custstyle1" onClick="doClick(this);"></td>
			</tr>
			<tr>
				<th>change style,custstyle2:</th>
				<td><input id="custstyle2" type="button"
					value="custstyle2" onClick="doClick(this);"></td>
			</tr>
			<tr>
				<th>change sheet:</th>
				<td><input id="changesheet" type="button"
					value="changesheet" onClick="doClick(this);"></td>
			</tr>
		</table>
	</div>
	
	<div id="gridweb"></div>
</body>
</html>