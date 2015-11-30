<%
String path = request.getContextPath();
String basePath = request.getScheme()+"://"+request.getServerName()+":"+request.getServerPort()+path+"/";
%>
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE9"/>
<base href="<%=basePath%>">
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<script type="text/javascript" language="javascript"
	src="grid/acw_client/acwmain.js"></script>
<script type="text/javascript" language="javascript"
	src="grid/acw_client/lang_en.js"></script>
<link href="grid/acw_client/menu.css" rel="stylesheet" type="text/css" />
<style>
span.acwxc {
	overflow: hidden;
	border: none;
	display: block;
	white-space: pre;
}
</style>
<style>
span.rotation90 {
	width: 100%;
	height: 100%;
	border: none;
	-webkit-transform: rotate(-90deg);
	-moz-transform: rotate(-90deg);
	filter: progid:DXImageTransform.Microsoft.BasicImage(rotation=3 );
	display: block
}
</style>
<style>
span.rotation-90 {
	filter: progid:DXImageTransform.Microsoft.BasicImage(rotation=1 );
	width: 100%;
	height: 100%;
	border: none;
	-webkit-transform: rotate(90deg);
	-moz-transform: rotate(90deg);
	display: block
}
</style>
<style>
span.wrap {
	white-space: pre-wrap;
	white-space: -moz-pre-wrap;
	white-space: -pre-wrap;
	white-space: -o-pre-wrap;
	word-wrap: break-word;
	-ms-word-break: break-all;
}
</style>