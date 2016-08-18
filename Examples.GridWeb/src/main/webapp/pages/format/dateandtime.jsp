<%@ page language="java" contentType="text/html;charset=UTF-8"
	pageEncoding="UTF-8" isELIgnored="false"%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<%@include file="/head.jsp" %>
<title>Worksheets - Aspose.Cells Grid Suite Demos</title>
<script type="text/javascript" src="grid/acw_client/jquery-2.1.4.min.js"></script>
<script type="text/javascript">
	function doClick(method) {
		var select = $("#DropDownList1").val();
		var text = $("option[value="+ select +"]").text();
		debugger;
		$.post("FormatServlet", {
			flag : method.id,
			gridwebuniqueid : $("#mycomponent").attr("webuniqueid"),
			value : $("#value").val(),
			DropDownList1 : select,
			text:text
		}, function(data) {
			$("#gridweb").html(data);
		}, "html");
	}

	$(document).ready(function() {
		
		//loadHead();//
		
		var method = {
			id : "loadDateTimeFormatFile"
		};
		doClick(method);
	});
</script>
</head>
<body>
	<div>
		<p>
			Pick a date format from the list, enter a value (text) and click <b>Submit</b>
			to see how demo applies custom date format to a grid cell and
			displays your value in it.
		</p>
	</div>

	<div>
		<table>
			<tr>
				<th>Reload Data:</th>
				<td><input id="loadDateTimeFormatFile" type="button"
					value="Reload" onClick="doClick(this);"></td>
			</tr>
			<tr>
				<th>NumberType:</th>
				<td><select id="DropDownList1">
						<option value="14">Date1</option>
						<option value="15">Date2</option>
						<option value="16">Date3</option>
						<option value="17">Date4</option>
						<option value="18">Time1</option>
						<option value="19">Time2</option>
						<option value="20">Time3</option>
						<option value="21">Time4</option>
						<option value="22">Time5</option>
						<option value="45">Time6</option>
						<option value="46">Time7</option>
						<option value="47">Time8</option>
						<option value="27">EasternDate1</option>
						<option value="28">EasternDate2</option>
						<option value="29">EasternDate3</option>
						<option value="30">EasternDate4</option>
						<option value="31">EasternDate5</option>
						<option value="36">EasternDate6</option>
						<option value="50">EasternDate7</option>
						<option value="51">EasternDate8</option>
						<option value="52">EasternDate9</option>
						<option value="53">EasternDate10</option>
						<option value="54">EasternDate11</option>
						<option value="57">EasternDate12</option>
						<option value="58">EasternDate13</option>
						<option value="32">EasternTime1</option>
						<option value="33">EasternTime2</option>
						<option value="34">EasternTime3</option>
						<option value="35">EasternTime4</option>
						<option value="55">EasternTime5</option>
						<option value="56">EasternTime6</option>
				</select></td>
				<th>Input Value:</th>
				<td><input id="value" type="text"> <input
					id="dateAndTime" type="button" value="Submit"
					onClick="doClick(this);"></td>
			</tr>
		</table>
	</div>

	<div id="gridweb"></div>

	<div>
		<table class="dtTABLE" cellspacing="0" border="1" style="width: 500px;">
			<tr>
				<td width="33%"><font color="gray"><b>Value</b></font></td>
				<td width="33%"><font color="gray"><b>Type</b></font></td>
				<td width="33%"><font color="gray"><b>Format String</b></font>
				</td>
			</tr>
			<tr>
				<td width="33%">0</td>
				<td width="33%">General</td>
				<td width="33%">General</td>
			</tr>
			<tr>
				<td width="33%">1</td>
				<td width="33%">Decimal</td>
				<td width="33%">0</td>
			</tr>
			<tr>
				<td width="33%">2</td>
				<td width="33%">Decimal</td>
				<td width="33%">0.00</td>
			</tr>
			<tr>
				<td width="33%">3</td>
				<td width="33%">Decimal</td>
				<td width="33%">#,##0</td>
			</tr>
			<tr>
				<td width="33%">4</td>
				<td width="33%">Decimal</td>
				<td width="33%">#,##0.00</td>
			</tr>
			<tr>
				<td width="33%">5</td>
				<td width="33%">Currency</td>
				<td width="33%">$#,##0;($#,##0)</td>
			</tr>
			<tr>
				<td width="33%">6</td>
				<td width="33%">Currency</td>
				<td width="33%">$#,##0;[Red]($#,##0)</td>
			</tr>
			<tr>
				<td width="33%">7</td>
				<td width="33%">Currency</td>
				<td width="33%">$#,##0.00;($#,##0.00)</td>
			</tr>
			<tr>
				<td width="33%">8</td>
				<td width="33%">Currency</td>
				<td width="33%">$#,##0.00;[Red]($#,##0.00)</td>
			</tr>
			<tr>
				<td width="33%">9</td>
				<td width="33%">Percentage</td>
				<td width="33%">0%</td>
			</tr>
			<tr>
				<td width="33%">10</td>
				<td width="33%">Percentage</td>
				<td width="33%">0.00%</td>
			</tr>
			<tr>
				<td width="33%">11</td>
				<td width="33%">Scientific</td>
				<td width="33%">0.00E+00</td>
			</tr>
			<tr>
				<td width="33%">12</td>
				<td width="33%">Fraction</td>
				<td width="33%"># ?</td>
			</tr>
			<tr>
				<td width="33%">13</td>
				<td width="33%">Fraction</td>
				<td width="33%"># ???</td>
			</tr>
			<tr>
				<td width="33%">14</td>
				<td width="33%">Date</td>
				<td width="33%">m/d/yyyy</td>
			</tr>
			<tr>
				<td width="33%">15</td>
				<td width="33%">Date</td>
				<td width="33%">d-mmm-yy</td>
			</tr>
			<tr>
				<td width="33%">16</td>
				<td width="33%">Date</td>
				<td width="33%">d-mmm</td>
			</tr>
			<tr>
				<td width="33%">17</td>
				<td width="33%">Date</td>
				<td width="33%">mmm-yy</td>
			</tr>
			<tr>
				<td width="33%">18</td>
				<td width="33%">Time</td>
				<td width="33%">h:mm AM/PM</td>
			</tr>
			<tr>
				<td width="33%">19</td>
				<td width="33%">Time</td>
				<td width="33%">h:mm:ss AM/PM</td>
			</tr>
			<tr>
				<td width="33%">20</td>
				<td width="33%">Time</td>
				<td width="33%">h:mm</td>
			</tr>

			<tr>
				<td width="33%">21</td>
				<td width="33%">Time</td>
				<td width="33%">h:mm:ss</td>
			</tr>
			<tr>
				<td width="33%">22</td>
				<td width="33%">Time</td>
				<td width="33%">m/d/yyyy h:mm</td>
			</tr>
			<tr>
				<td width="33%">37</td>
				<td width="33%">Accounting</td>
				<td width="33%">#,##0;(#,##0)</td>
			</tr>
			<tr>
				<td width="33%">38</td>
				<td width="33%">Accounting</td>
				<td width="33%">#,##0;[Red](#,##0)</td>
			</tr>
			<tr>
				<td width="33%">39</td>
				<td width="33%">Accounting</td>
				<td width="33%">#,##0.00;(#,##0.00)</td>
			</tr>
			<tr>
				<td width="33%">40</td>
				<td width="33%">Accounting</td>
				<td width="33%">#,##0.00;[Red](#,##0.00)</td>
			</tr>
			<tr>
				<td width="33%">41</td>
				<td width="33%">Accounting</td>
				<td width="33%">_ * #,##0_ ;_ * (#,##0)_ ;_ * "-"_ ;_ @_</td>
			</tr>
			<tr>
				<td width="33%">42</td>
				<td width="33%">Currency</td>
				<td width="33%">_ $* #,##0_ ;_ $* (#,##0)_ ;_ $* "-"_ ;_ @_</td>
			</tr>
			<tr>
				<td width="33%">43</td>
				<td width="33%">Accounting</td>
				<td width="33%">_ * #,##0.00_ ;_ * (#,##0.00)_ ;_ * "-"??_ ;_
					@_</td>
			</tr>
			<tr>
				<td width="33%">44</td>
				<td width="33%">Currency</td>
				<td width="33%">_ $* #,##0.00_ ;_ $* (#,##0.00)_ ;_ $* "-"??_
					;_ @_</td>
			</tr>
			<tr>
				<td width="33%">45</td>
				<td width="33%">Time</td>
				<td width="33%">mm:ss</td>
			</tr>
			<tr>
				<td width="33%">46</td>
				<td width="33%">Time</td>
				<td width="33%">[h]:mm:ss</td>
			</tr>
			<tr>
				<td width="33%">47</td>
				<td width="33%">Time</td>
				<td width="33%">mm:ss.0</td>
			</tr>
			<tr>
				<td width="33%">48</td>
				<td width="33%">Scientific</td>
				<td width="33%">##0.0E+00</td>
			</tr>
			<tr>
				<td width="33%">49</td>
				<td width="33%">Text</td>
				<td width="33%">@</td>
			</tr>
		</table>
	</div>
</body>
</html>