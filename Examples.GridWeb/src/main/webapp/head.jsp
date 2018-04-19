<%
String path = request.getContextPath();
String basePath = request.getScheme()+"://"+request.getServerName()+":"+request.getServerPort()+path+"/";
%>

<base href="<%=basePath%>">
<script type="text/javascript" language="javascript"
	src="grid/acw_client/acwmain.js"></script>
<script type="text/javascript" language="javascript"
	src="grid/acw_client/lang_en.js"></script>
<link href="grid/acw_client/menu.css" rel="stylesheet" type="text/css" />
<link rel="stylesheet" href="Scripts/jquery-ui.css">
<script src="Scripts/jquery-2.1.1.js"></script>
<script src="Scripts/jquery-ui.js"></script>
<style>
span.acwxc {
	overflow: hidden;
	border: none;
	display: block;
	white-space: pre;
}
 
span.rotation90 {
	width: 100%;
	height: 100%;
	border: none;
	-webkit-transform: rotate(-90deg);
	-moz-transform: rotate(-90deg);
	filter: progid:DXImageTransform.Microsoft.BasicImage(rotation=3 );
	display: block
}
 
span.rotation-90 {
	filter: progid:DXImageTransform.Microsoft.BasicImage(rotation=1 );
	width: 100%;
	height: 100%;
	border: none;
	-webkit-transform: rotate(90deg);
	-moz-transform: rotate(90deg);
	display: block
}
 
span.wrap {
	white-space: pre-wrap;
	white-space: -moz-pre-wrap;
	white-space: -pre-wrap;
	white-space: -o-pre-wrap;
	word-wrap: break-word;
	-ms-word-break: break-all;
}
span.acwcmmnt {position: relative;}
span.acwcmmnt:after { content: '';position: absolute;top: 0;right: 0;width: 0;height: 0; display: block;border-left: 7px solid transparent; border-bottom: 7px solid transparent;border-top: 7px solid #f00;}
</style>