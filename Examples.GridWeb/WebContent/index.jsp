<%@ page language="java" contentType="text/html;charset=UTF-8"
	pageEncoding="UTF-8" isELIgnored="false"%>
<%
	String path = request.getContextPath();
	String basePath = request.getScheme() + "://"
			+ request.getServerName() + ":" + request.getServerPort()
			+ path + "/";
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<base href="<%=basePath%>">
<title>Index</title>
<style type="text/css">
li {
	padding-top: 10px;
}
</style>
</head>

<body>
	<h2>Welcome to the Aspose.Cells.GridWeb Featured Demos!</h2>
	<h3>Note: These aspose-cells-grid support the following browsers:
		IE(or any IE kernel browsers), Mozilla, Mozilla Firefox and Opera. The
		IE 6.0 or new version is recommended.</h3>
	<ul>
		<!-- <li>
			<a target="_blank" href="pages/format/"></a>
			<div>
				<h5>Description</h5>
				
			</div>
		</li> -->
	
		 <li>
			<a target="_blank" href="pages/worksheets/add_remove_comments_from_client_side.jsp">Add/Remove Comments from Client Side</a>
			<div>
				<h5>Description</h5>
				Add/Remove Comments from Client Side 
			</div>
		</li>
		 <li>
			<a target="_blank" href="pages/worksheets/add_remove_hyperlink_from_client_side.jsp">Add/Remove Hyperlinks from Client Side</a>
			<div>
				<h5>Description</h5>
				Add/Remove Hyperlinks from Client Side 
			</div>
		</li>
		 <li>
			<a target="_blank" href="pages/worksheets/update_font_from_client_side.jsp">Update Font from Client Side</a>
			<div>
				<h5>Description</h5>
				Update Font from Client Side
			</div>
		</li>
		 <li>
			<a target="_blank" href="pages/worksheets/show_add_button.jsp">Show Add/Remove Worksheet Buttons</a>
			<div>
				<h5>Description</h5>
				Show Add/Remove Worksheet Buttons
			</div>
		</li>
		<li>
			<a target="_blank" href="pages/commons/create_content.jsp">Creating Contents</a>
			<div>
				<h5>Description</h5>
				This Demo creates a Worksheet from the scratch.
			</div>
		</li>
		<!-- not work yet
		<li>
			<a target="_blank" href="pages/commons/modes.jsp">Edit/ReadOnly Mode</a>
			<div>
				<h5>Description</h5>
				This Demo demonstrates the functioning of &quot;Edit&quot; and &quot;Read Only&quot; Modes of GridWeb Control.
			</div>
		</li>
		-->
		<li>
			<a target="_blank" href="pages/commons/editorbox.jsp">Show Editor box</a>
			<div>
				<h5>Description</h5>
				This Demo show the editor box mode of GridWeb Control.
			</div>
		</li>
		<li>
			<a target="_blank" href="pages/commons/sheets.jsp">Worksheets</a>
			<div>
				<h5>Description</h5>
				The Demo exhibits the manipulation (Add, Remove, Copy) of Sheets. 
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/commons/webcells.jsp">Cells</a>
			<div>
				<h5>Description</h5>
				The Demo exhibits the manipulation (Insertion, Deletion) of rows/columns, merging cells and adding/removing comments.  
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/commons/headerbar_commandbutton.jsp">HeaderBar &amp; CommandButton</a>
			<div>
				<h5>Description</h5>
				This Demo covers some useful properties of GridWeb Control. 
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/commons/freezepane.jsp">FreezePane Report</a>
			<div>
				<h5>Description</h5>
				This Demo Imports an Excel File from a source and demonstrates Freezing Panes. 
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/commons/freezepane_custom.jsp">Freeze/Unfreeze Panes</a>
			<div>
				<h5>Description</h5>
				This Demo expresses customized Freezing Panes. 
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/commons/hyperlinkdemo.jsp">Hyperlink &amp; CellImage</a>
			<div>
				<h5>Description</h5>
				The Demo presents the functionality of Hyperlink and CellImage Object.
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/commons/comments_ruid.jsp">Comment</a>
			<div>
				<h5>Description</h5>
				The Demo presents the functionality of cell comment.
			</div>
		</li>
		
	   <li>
			<a target="_blank" href="pages/commons/setborders.jsp">border setting</a>
			<div>
				<h5>Description</h5>
				The Demo presents the functionality of setborders api.
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/commons/customheaders.jsp">Custom Headers</a>
			<div>
				<h5>Description</h5>
				This Demo customizes the Labels of Column Headers.
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/formula/math.jsp">Math</a>
			<div>
				<h5>Description</h5>
				This Demo presents the exercise of Mathematical Functions.
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/formula/text_data.jsp">Text &amp; Data</a>
			<div>
				<h5>Description</h5>
				The Demo covers the practice session of String Functions. 
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/formula/statistical.jsp">Statistical</a>
			<div>
				<h5>Description</h5>
				The Demo presents the exercise of Statistical Functions. 
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/formula/datetime.jsp">Date &amp; Time</a>
			<div>
				<h5>Description</h5>
				The Demo presents the exercise of Date and Time Functions. 
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/formula/logical.jsp">Logical</a>
			<div>
				<h5>Description</h5>
				This Demo presents the Demonstration of Logical Functions. 
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/format/customformat.jsp">Custom Format</a>
			<div>
				<h5>Description</h5>
				This Demo presents an exercise of Custom Formats. 
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/format/dateandtime.jsp">Date & Time Format</a>
			<div>
				<h5>Description</h5>
				This Demo covers the exercise of Date and Time Formats. 
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/commons/change_style.jsp">Skins</a>
			<div>
				<h5>Description</h5>
				This Demo covers the Demonstration of GridWeb Control’s preset styles and custom styles.
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/commons/validation.jsp">Protection/Validation</a>
			<div>
				<h5>Description</h5>
				This Demo introduces the Data Validation capabilities of GridWeb Control.  
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/commons/pagination.jsp">Paginating Sheet</a>
			<div>
				<h5>Description</h5>
				This Demo Imports an Excel File from a source and divides the contents of the sheet into different pages. 
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/commons/largerows.jsp">load files in async way for large files</a>
			<div>
				<h5>Description</h5>
				This Demo Imports a large Excel File with huge rows. also you can use Paginating way.
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/commons/sort.jsp">Sort</a>
			<div>
				<h5>Description</h5>
				The Demo represents the sorting capabilities of GriWeb Control.  
			</div>
		</li>
		
		<li>
			<a target="_blank" href="pages/filter/autofilter.jsp">AutoFilter</a>
			<div>
				<h5>Description</h5>
				This Demo Imports an Excel File from a source and Set the AutoFilter feature.   
			</div>
		</li>
	
 	
		<li>
			<a target="_blank" href="pages/commons/events.jsp">Handling Events</a>
			<div>
				<h5>Description</h5>
				This Demo Demonstrates Event Handling related to GridWeb Control. 
			</div>
		</li>
		<li>
			<a target="_blank" href="pages/commons/clientfunction.jsp">Handling ajax call event</a>
			<div>
				<h5>Description</h5>
				This Demo Demonstrates how to use cell select callback event. 
			</div>
		</li>
		<li>
			<a target="_blank" href="pages/commons/ajaxmodify.jsp">another Handling ajax call event</a>
			<div>
				<h5>Description</h5>
				This Demo Demonstrates how to use cell select callback event. 
			</div>
		</li>
		<li>
			<a target="_blank" href="pages/commons/chartrefresh.jsp">display chart </a>
			<div>
				<h5>Description</h5>
				This Demo Imports an Excel File with charts. 
			</div>
		</li>
		 <li>
			<a target="_blank" href="pages/commons/pivottable.jsp">display pivottable </a>
			<div>
				<h5>Description</h5>
				This Demo Imports an Excel File with Pivottable. 
			</div>
		</li>
		 <li>
			<a target="_blank" href="pages/commons/grouprowcol.jsp">display row/column group </a>
			<div>
				<h5>Description</h5>
				This Demo Imports an Excel File with Group row/column. 
			</div>
		</li>
		 <li>
			<a target="_blank" href="pages/commons/displaycontrols.jsp">display controls </a>
			<div>
				<h5>Description</h5>
				This Demo Imports an Excel File with controls. 
			</div>
		</li>
		<!-- <li>
			<a target="_blank" href="pages/format/"></a>
			<div>
				<h5>Description</h5>
				
			</div>
		</li> -->
	</ul>
</body>
</html>
