package com.aspose.gridweb.test.servlet;

import java.lang.reflect.Field;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.aspose.gridweb.BorderStyle;
import com.aspose.gridweb.CellEventArgs;
import com.aspose.gridweb.Color;
import com.aspose.gridweb.GridCells;
import com.aspose.gridweb.GridHyperlink;
import com.aspose.gridweb.GridTableItemStyle;
import com.aspose.gridweb.GridWebBean;
import com.aspose.gridweb.GridWorksheet;
import com.aspose.gridweb.HorizontalAlign;
import com.aspose.gridweb.PresetStyle;
import com.aspose.gridweb.Unit;
import com.aspose.gridweb.VerticalAlign;
import com.aspose.gridweb.WorkbookEventHandler;
import com.aspose.gridweb.test.TestGridWebBaseServlet;
import com.aspose.gridweb.test.data.AjaxModifyCellEventHandler;
import com.aspose.gridweb.test.data.CUSCommandEventHandler;
import com.aspose.gridweb.test.data.ClickOnAjaxEventHandler;
import com.aspose.gridweb.test.data.MsgCellEventHandler;
import com.aspose.gridweb.test.data.MsgPutWorkbookHandler;
import com.aspose.gridweb.test.data.PageChangedEventHandler;
import com.aspose.gridweb.test.data.RowColEventHandler;
import com.aspose.gridweb.test.data.SortEventHandler;

public class FeatureServlet extends TestGridWebBaseServlet {
	private static final long serialVersionUID = 1L;

	 

	@Override
	public void reload(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {

		try {
			super.reloadfile(gridweb,request,"data.xls");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void loadFreezePaneFile(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		 
		try {
			super.reloadfile(gridweb,request,"freezepane.xls");
		} catch (Exception e) {
			e.printStackTrace();
		}

		 
		GridWorksheet gridWorksheet =gridweb.getActiveSheet();
		gridWorksheet.freezePanes(3, 3, 3, 3);
	}

	public void loadTestLargeFile(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		
		try {
			super.reloadfile(gridweb,request,"testlargerows.xlsx");
		} catch (Exception e) {
			e.printStackTrace();
		}
		gridweb.setPageSize(20);
		 
	}
public void loadLargeFileAsync(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {

		try {
			super.reloadfile(gridweb,request,"testlargerows.xlsx");
		} catch (Exception e) {
			e.printStackTrace();
		}
		gridweb.setEnableAsync(true);
		 
	}
	public void freezePane(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		int row = Integer.parseInt(request.getParameter("row"));
		int column = Integer.parseInt(request.getParameter("column"));
		int rowNumber = Integer.parseInt(request.getParameter("rowNumber"));
		int columnNumber = Integer.parseInt(request.getParameter("columnNumber"));

		
		GridWorksheet gridWorksheet = gridweb.getActiveSheet();
		gridWorksheet.freezePanes(row, column, rowNumber, columnNumber);
	}

	public void unfreezePane(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		
		
		GridWorksheet gridWorksheet =gridweb.getActiveSheet();
		gridWorksheet.unFreezePanes();
	}

	public void customHeaders(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		
		// gridWorkSheet
		GridWorksheet gridWorkSheet = gridweb.getActiveSheet();
		gridWorkSheet.setColumnCaption(0, "Product");
		gridWorkSheet.setColumnCaption(1, "Category");
		gridWorkSheet.setColumnCaption(2, "Price");
		gridWorkSheet.setRowCaption(2, "row2");
	 

		GridCells gridCells = gridWorkSheet.getCells();
		gridCells.get("A1").setValue("Aniseed Syrup");
		gridCells.get("A2").setValue("Boston Crab Meat");
		gridCells.get("A3").setValue("Chang");

		gridCells.get("B1").setValue("Condiments");
		gridCells.get("B2").setValue("Seafood");
		gridCells.get("B3").setValue("Beverages");
		gridCells.setColumnWidthPixel(0, 180);


	}

	public void loadDateTimeFile(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {

		try {
			super.reloadfile(gridweb,request,"datetime.xls");
		} catch (Exception e) {
			e.printStackTrace();
	}
	}
	public void loadPivotFile(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		 
		try {
			super.reloadfile(gridweb,request,"PivotTable.xls");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void loadTextAndDataFile(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		 
		try {
			super.reloadfile(gridweb,request,"TextAndData.xls");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void loadMathFile(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		 
		try {
			super.reloadfile(gridweb,request,"Math.xls");
			 
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public void loadChartFile(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		 
		try {
			super.reloadfile(gridweb,request,"charttest.xls");
			gridweb.setWidth(Unit.Pixel(1200));
			gridweb.setHeight(Unit.Pixel(700));
			 
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public void loadConditionFormatFile(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		
		try {
			super.reloadfile(gridweb,request,"conditionformat.xlsx");
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public void loadPivot(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		 
		try {
			super.reloadfile(gridweb,request,"pivottable.xls");
			gridweb.setWidth(Unit.Pixel(1200));
			gridweb.setHeight(Unit.Pixel(700));
			 
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void loadGroupRowCol(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		 
		try {
			super.reloadfile(gridweb,request,"grouprowcol.xlsx");
			gridweb.setRenderHiddenRow(true);
			gridweb.setWidth(Unit.Pixel(1200));
			gridweb.setHeight(Unit.Pixel(600));
			 
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void loadLargeRows(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		 
		try {
			super.reloadfile(gridweb,request,"employeesales.xls");
			gridweb.setEnableAsync(true);
			gridweb.setWidth(Unit.Pixel(1200));
			gridweb.setHeight(Unit.Pixel(700));
			 
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public void loadControls(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		 
		try {
			super.reloadfile(gridweb,request,"controls.xlsx");
			 
			gridweb.setWidth(Unit.Pixel(1200));
			gridweb.setHeight(Unit.Pixel(500));
			 
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void cellmodifyajax(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		 
		try {
			gridweb.CellModifiedOnAjax=new AjaxModifyCellEventHandler();
			 
			gridweb.setWidth(Unit.Pixel(1200));
			gridweb.setHeight(Unit.Pixel(500));
			 
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	
	
	public void loadChartFileSubmit(final GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		 
		try {
			super.reloadfile(gridweb,request,"charttest.xls");
			gridweb.setWidth(Unit.Pixel(1200));
			gridweb.setHeight(Unit.Pixel(700));
			//the default is true,so here we set false to avoid auto refreshing 
			gridweb.setAutoRefreshChart(false);
			WorkbookEventHandler SubmitCommand = new WorkbookEventHandler() {
				@Override
				public void handleCellEvent(Object arg0, CellEventArgs arg1) {

					gridweb.refreshChartShape();
				}
			};
			gridweb.SubmitCommand = SubmitCommand;
			 
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void loadLogicalFile(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
	 
		try {
			super.reloadfile(gridweb,request,"Logical.xls");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void loadStatisticalFile(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		 
		try {
			super.reloadfile(gridweb,request,"Statistical.xls");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void loadSkinsFile(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		 
		try {
			super.reloadfile(gridweb,request,"Skins.xls");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public void savecustomfile(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		String file = request.getParameter("filename");
		gridweb.saveCustomStyleFile(file);
	 
	}
	public void loadcustomfile(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		String file = request.getParameter("filename");
		gridweb.setCustomStyleFileName(file);
	 
	}

	public void changeStyle(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		String style = request.getParameter("style");

		if (style.startsWith("Custom")) {
			String basePath = request.getScheme() + "://" + request.getServerName() + ":" + request.getServerPort() + webPath
					+ "/";
			String url = basePath + "xml/" + style + ".xml";
			gridweb.setCustomStyleFileName(url);
			return;
		}

		Field[] fields = PresetStyle.class.getDeclaredFields();
		int presetStyle = PresetStyle.STANDARD;
		for (Field field : fields) {
			if (field.getName().equalsIgnoreCase(style)) {
				try {
					presetStyle = field.getInt(field.getName());
				} catch (IllegalArgumentException e) {
					e.printStackTrace();
				} catch (IllegalAccessException e) {
					e.printStackTrace();
				}
			}
		}
		gridweb.setPresetStyle(presetStyle);
	}

	public void pagination(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		 
		try {
			super.reloadfile(gridweb,request,"employeesales.xls");
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		gridweb.setPageSize(20);
	}

	public void sort(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		 
		try {
			super.reloadfile(gridweb,request,"sort.xls");
		} catch (Exception e) {
			e.printStackTrace();
		}

		
		// Creates sortting header style.
		GridTableItemStyle gridTableItemStyle = new GridTableItemStyle();
		gridTableItemStyle.setBorderStyle(BorderStyle.Outset);
		gridTableItemStyle.setBorderWidth(new Unit(2));
		gridTableItemStyle.setBorderColor(Color.getWhite());
		gridTableItemStyle.setBackColor(Color.getSilver());
		gridTableItemStyle.setHorizontalAlign(HorizontalAlign.Center);
		gridTableItemStyle.setVerticalAlign(VerticalAlign.Middle);

		// Creates Hyperlinks.
		  GridWorksheet sheet = gridweb.getWorkSheets().get(0);
		sheet.getCells().get("A1").copyStyle(gridTableItemStyle);
		int i = sheet.getHyperlinks().add("A1", 1, 1, "");
		GridHyperlink hlink = sheet.getHyperlinks().get(i);
		hlink.setAddress("CELLCMD:A1");
		hlink.setTextToDisplay("orderid");

		sheet.getCells().get("B1").copyStyle(gridTableItemStyle);
		i = sheet.getHyperlinks().add("B1", 1, 1, "");
		hlink = sheet.getHyperlinks().get(i);
		hlink.setAddress("CELLCMD:B1");
		hlink.setTextToDisplay("Sales Amout");

		sheet.getCells().get("C1").copyStyle(gridTableItemStyle);
		i = sheet.getHyperlinks().add("C1", 1, 1, "");
		hlink = sheet.getHyperlinks().get(i);
		hlink.setAddress("CELLCMD:C1");
		hlink.setTextToDisplay("Percent of Saler's Total");

		sheet.getCells().get("D1").copyStyle(gridTableItemStyle);
		i = sheet.getHyperlinks().add("D1", 1, 1, "");
		hlink = sheet.getHyperlinks().get(i);
		hlink.setAddress("CELLCMD:D1");
		hlink.setTextToDisplay("Percent of Country Total");

		final GridWorksheet sheet1 = gridweb.getWorkSheets().get(1);

		sheet1.getCells().get("A1").copyStyle(gridTableItemStyle);
		i = sheet1.getHyperlinks().add("A1", 1, 1, "");
		hlink = sheet1.getHyperlinks().get(i);
		hlink.setAddress("CELLCMD:1A1");
		hlink.setTextToDisplay("Product");

		sheet1.getCells().get("A2").copyStyle(gridTableItemStyle);
		i = sheet1.getHyperlinks().add("A2", 1, 1, "");
		hlink = sheet1.getHyperlinks().get(i);
		hlink.setAddress("CELLCMD:1A2");
		hlink.setTextToDisplay("Category");

		sheet1.getCells().get("A3").copyStyle(gridTableItemStyle);
		i = sheet1.getHyperlinks().add("A3", 1, 1, "");
		hlink = sheet1.getHyperlinks().get(i);
		hlink.setAddress("CELLCMD:1A3");
		hlink.setTextToDisplay("Package");

		sheet1.getCells().get("A4").copyStyle(gridTableItemStyle);
		i = sheet1.getHyperlinks().add("A4", 1, 1, "");
		hlink = sheet1.getHyperlinks().get(i);
		hlink.setAddress("CELLCMD:1A4");
		hlink.setTextToDisplay("Quantity");

		SortEventHandler se=new SortEventHandler();
		gridweb.CellCommand = se;
     	gridweb.setOnCellErrorClientFunction("statebagtest");
     	gridweb.setWidth(Unit.Pixel(567));

	}

	public void events(final GridWebBean gridweb,final HttpServletRequest request, final HttpServletResponse response) {
		this.reload(gridweb,request, response);
		
		gridweb.setPageSize(8);
		 
		// gridWorkSheet
		GridWorksheet gridWorkSheet =gridweb.getActiveSheet();
		gridWorkSheet.getCells().setColumnWidthPixel(0, 180);


		gridweb.SubmitCommand = new MsgPutWorkbookHandler("submit");


		gridweb.SaveCommand = new MsgPutWorkbookHandler("save");;


		gridweb.UndoCommand = new MsgPutWorkbookHandler("undo");;

				
		gridweb.SheetTabClick = new MsgPutWorkbookHandler("tabclick");;

		gridweb.CustomCommand = new CUSCommandEventHandler();


		gridweb.RowDoubleClick = new RowColEventHandler("RowDoubleClick");


		gridweb.ColumnDoubleClick = new RowColEventHandler("ColumnDoubleClick");


		gridweb.CellDoubleClick = new MsgCellEventHandler("CellDoubleClick");


		gridweb.CellClickOnAjax = new ClickOnAjaxEventHandler();


		gridweb.RowInserted = new RowColEventHandler("RowInserted");


		gridweb.RowDeleted = new RowColEventHandler("RowDeleted");

		 
		gridweb.RowDeleting = new RowColEventHandler("RowDeleting");

		 
		gridweb.ColumnInserted = new RowColEventHandler("ColumnInserted");

		 
		gridweb.ColumnDeleted =new RowColEventHandler("ColumnDeleted");

		 
		gridweb.ColumnDeleting = new RowColEventHandler("ColumnDeleting");

		 
		gridweb.CellCommand = new MsgCellEventHandler("CellCommand");;

		 
		gridweb.PageIndexChanged = new PageChangedEventHandler();
			}
	 
	public void clientfunc(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		
		 gridweb.CellClickOnAjax=new ClickOnAjaxEventHandler();
		 gridweb.setOnCellSelectedAjaxCallBackClientFunction("dealwithcellselectcallback");
	}

}
