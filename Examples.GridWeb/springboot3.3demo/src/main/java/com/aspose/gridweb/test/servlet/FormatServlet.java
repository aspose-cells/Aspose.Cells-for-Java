package com.aspose.gridweb.test.servlet;

import java.util.Date;

//
//import com.aspose.cells.GridCell;
//import com.aspose.cells.GridCells;
//import com.aspose.cells.GridTableItemStyle;
//import com.aspose.cells.GridWebBean;
//import com.aspose.cells.GridWorksheetCollection;
import com.aspose.gridweb.GridCell;
import com.aspose.gridweb.GridCells;
import com.aspose.gridweb.GridTableItemStyle;
import com.aspose.gridweb.GridWebBean;
import com.aspose.gridweb.test.TestGridWebBaseServlet;

import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;

public class FormatServlet extends TestGridWebBaseServlet {

	private static final long serialVersionUID = 1L;

	@Override
	public void reload(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {

		try {
			super.reloadfile(gridweb,request, "format.xls");

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void loadCustomFormatFile(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {

		this.reload(gridweb,request, response);

	 
		GridCells gridCells = gridweb.getActiveSheet().getCells();

		gridCells.get("A1").setValue("Custom Format");
		gridCells.get("A2").setValue("0.0");
		gridCells.get("A3").setValue("0.000");
		gridCells.get("A4").setValue("#,##0.0");
		gridCells.get("A5").setValue("US$#,##0;US$-#,##0");
		gridCells.get("A6").setValue("0.0%");
		gridCells.get("A7").setValue("0.000E+00");
		gridCells.get("A8").setValue("yyyy-m-d h:mm");

		gridCells.get("B1").setValue("Format Results");

		GridCell B2 = gridCells.get("B2");
		B2.setValue(12345.6789);
		B2.setCustom("0.0");

		GridCell B3 = gridCells.get("B3");
		B3.setValue(12345.6789);

		B3.setCustom("0.000");

		GridCell B4 = gridCells.get("B4");
		B4.setValue(543123456.789);

		B4.setCustom("#,##0.0");

		GridCell B5 = gridCells.get("B5");
		B5.setValue(-543123456.789);

		B5.setCustom("US$#,##0;US$-#,##0");

		GridCell B6 = gridCells.get("B6");
		B6.setValue(0.925687);

		B6.setCustom("0.0%");

		GridCell B7 = gridCells.get("B7");
		B7.setValue(-1234567890.5687);

		B7.setCustom("0.000E+00");

		GridCell B8 = gridCells.get("B8");
		B8.setValue(new Date());

		B8.setCustom("yyyy-m-d h:mm");

	}

	public void customFormat(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {

		this.reload(gridweb,request, response);

		 
		GridCells gridCells = gridweb.getActiveSheet().getCells();

		gridCells.get("A1").setValue("Custom Format");
		gridCells.get("A2").setValue(request.getParameter("format"));

		gridCells.get("B1").setValue("Format Results");
		GridCell B2 = gridCells.get("B2");
		///notice we use this api to automatically  convert string value
		B2.putValue(request.getParameter("value"),true);
		GridTableItemStyle B2Style = B2.getStyle();
		B2Style.setCustom(request.getParameter("format"));
		B2.setStyle(B2Style);
	}

	public void loadDateTimeFormatFile(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {

		this.reload(gridweb,request, response);

		
		GridCells gridCells = gridweb.getActiveSheet().getCells();

		gridCells.get("A1").setValue("Number Type");
		gridCells.get("A2").setValue("Date 1:");
		gridCells.get("A3").setValue("Date 2:");
		gridCells.get("A4").setValue("Date 3:");
		gridCells.get("A5").setValue("Date 4:");

		gridCells.get("A6").setValue("Time 1:");
		gridCells.get("A7").setValue("Time 2:");
		gridCells.get("A8").setValue("Time 3:");
		gridCells.get("A9").setValue("Time 4:");
		gridCells.get("A10").setValue("Time 5:");
		gridCells.get("A11").setValue("Time 6:");
		gridCells.get("A12").setValue("Time 7:");
		gridCells.get("A13").setValue("Time 8:");

		gridCells.get("A14").setValue("EasternDate 1:");
		gridCells.get("A15").setValue("EasternDate 2:");
		gridCells.get("A16").setValue("EasternDate 3:");
		gridCells.get("A17").setValue("EasternDate 4:");
		gridCells.get("A18").setValue("EasternDate 5:");
		gridCells.get("A19").setValue("EasternDate 6:");
		gridCells.get("A20").setValue("EasternDate 7:");
		gridCells.get("A21").setValue("EasternDate 8:");
		gridCells.get("A22").setValue("EasternDate 9:");
		gridCells.get("A23").setValue("EasternDate 10:");
		gridCells.get("A24").setValue("EasternDate 11:");
		gridCells.get("A25").setValue("EasternDate 12:");
		gridCells.get("A26").setValue("EasternDate 13:");

		gridCells.get("A27").setValue("EasternTime 1:");
		gridCells.get("A28").setValue("EasternTime 2:");
		gridCells.get("A29").setValue("EasternTime 3:");
		gridCells.get("A30").setValue("EasternTime 4:");
		gridCells.get("A31").setValue("EasternTime 5:");
		gridCells.get("A32").setValue("EasternTime 6:");

		gridCells.get("B1").setValue("Format Results");

		GridCell B2 = gridCells.get("B2");
		B2.setValue(new Date());
		 
		B2.setNumberType(14);
		 

		GridCell B3 = gridCells.get("B3");
		B3.setValue(new Date());
		 
		B3.setNumberType(15);
		 

		GridCell B4 = gridCells.get("B4");
		B4.setValue(new Date());
	 
		B4.setNumberType(16);
	 

		GridCell B5 = gridCells.get("B5");
		B5.setValue(new Date());
		 
		B5.setNumberType(17);
		 

		GridCell B6 = gridCells.get("B6");
		B6.setValue(new Date());
		 
		B6.setNumberType(18);
	 

		GridCell B7 = gridCells.get("B7");
		B7.setValue(new Date());
		 
		B7.setNumberType(19);
	 

		GridCell B8 = gridCells.get("B8");
		B8.setValue(new Date());
		 
		B8.setNumberType(20);
		 

		GridCell B9 = gridCells.get("B9");
		B9.setValue(new Date());
		 
		B9.setNumberType(21);
	 

		GridCell B10 = gridCells.get("B10");
		B10.setValue(new Date());
		 
		B10.setNumberType(22);
	 

		GridCell B11 = gridCells.get("B11");
		B11.setValue(new Date());
	 
		B11.setNumberType(45);
		 

		GridCell B12 = gridCells.get("B12");
		B12.setValue(new Date());
	 
		B12.setNumberType(46);
		 

		GridCell B13 = gridCells.get("B13");
		B13.setValue(new Date());
		 
		B13.setNumberType(47);
		 

		GridCell B14 = gridCells.get("B14");
		B14.setValue(new Date());
	 
		B14.setNumberType(27);
	 

		GridCell B15 = gridCells.get("B15");
		B15.setValue(new Date());
	 
		B15.setNumberType(28);
		 

		GridCell B16 = gridCells.get("B16");
		B16.setValue(new Date());
	 
		B16.setNumberType(29);
		 

		GridCell B17 = gridCells.get("B17");
		B17.setValue(new Date());
		 
		B17.setNumberType(30);
	 

		GridCell B18 = gridCells.get("B18");
		B18.setValue(new Date());
	 
		B18.setNumberType(31);
	 

		GridCell B19 = gridCells.get("B19");
		B19.setValue(new Date());
	 
		B19.setNumberType(36);
	 

		GridCell B20 = gridCells.get("B20");
		B20.setValue(new Date());
	 
		B20.setNumberType(50);
		 

		GridCell B21 = gridCells.get("B21");
		B21.setValue(new Date());
	 
		B21.setNumberType(51);
		 

		GridCell B22 = gridCells.get("B22");
		B22.setValue(new Date());
	 
		B22.setNumberType(52);
	 

		GridCell B23 = gridCells.get("B23");
		B23.setValue(new Date());
	 
		B23.setNumberType(53);
	 

		GridCell B24 = gridCells.get("B24");
		B24.setValue(new Date());
		 
		B24.setNumberType(54);
	 

		GridCell B25 = gridCells.get("B25");
		B25.setValue(new Date());
		 
		B25.setNumberType(57);
	 

		GridCell B26 = gridCells.get("B26");
		B26.setValue(new Date());
		 
		B26.setNumberType(58);
	 

		GridCell B27 = gridCells.get("B27");
		B27.setValue(new Date());
		 
		B27.setNumberType(32);
	 

		GridCell B28 = gridCells.get("B28");
		B28.setValue(new Date());
	 
		B28.setNumberType(33);
 

		GridCell B29 = gridCells.get("B29");
		B29.setValue(new Date());
		 
		B29.setNumberType(34);
	 

		GridCell B30 = gridCells.get("B30");
		B30.setValue(new Date());
	 
		B30.setNumberType(35);
	 

		GridCell B31 = gridCells.get("B31");
		B31.setValue(new Date());
	 
		B31.setNumberType(55);
	 

		GridCell B32 = gridCells.get("B32");
		B32.setValue(new Date());
		 
		B32.setNumberType(56);
	 
	}

	public void dateAndTime(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		
		this.reload(gridweb,request, response);
		
		String  value = (request.getParameter("value"));
		int numberType = Integer.parseInt(request.getParameter("DropDownList1"));
		String text = request.getParameter("text");

		 
		GridCells gridCells = gridweb.getActiveSheet().getCells();

		gridCells.get("A1").setValue("Number Type");
		gridCells.get("B1").setValue("Format Results");

		gridCells.get("A2").setValue(text);

		GridCell B2 = gridCells.get("B2");
		///notice we use this api to automatically  convert string value
		B2.putValue(value,true);

		B2.setNumberType(numberType);

	}

}
