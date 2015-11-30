package com.aspose.gridweb.test.servlet;

import java.lang.reflect.Field;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.aspose.gridweb.BorderStyle;
import com.aspose.gridweb.CellErrorHandler;
import com.aspose.gridweb.CellEventArgs;
import com.aspose.gridweb.CellEventHandler;
import com.aspose.gridweb.CellEventStringHandler;
import com.aspose.gridweb.Color;
import com.aspose.gridweb.CustomCommandEventHandler;
import com.aspose.gridweb.GridCellException;
import com.aspose.gridweb.GridCells;
import com.aspose.gridweb.GridHyperlink;
import com.aspose.gridweb.GridTableItemStyle;
import com.aspose.gridweb.GridWebBean;
import com.aspose.gridweb.GridWorksheet;
import com.aspose.gridweb.GridWorksheetCollection;
import com.aspose.gridweb.HorizontalAlign;
import com.aspose.gridweb.OnErrorActionQuery;
import com.aspose.gridweb.PresetStyle;
import com.aspose.gridweb.RowColumnEventArgs;
import com.aspose.gridweb.RowColumnEventHandler;
import com.aspose.gridweb.Unit;
import com.aspose.gridweb.VerticalAlign;
import com.aspose.gridweb.WorkbookEventHandler;
import com.aspose.gridweb.test.TestGridWebBaseServlet;

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

		GridWorksheetCollection gridWorksheetCollection = gridweb.getWorkSheets();
		GridWorksheet gridWorksheet = gridWorksheetCollection.get(gridWorksheetCollection.getActiveSheetIndex());
		gridWorksheet.freezePanes(3, 3, 3, 3);
	}

	public void freezePane(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		int row = Integer.parseInt(request.getParameter("row"));
		int column = Integer.parseInt(request.getParameter("column"));
		int rowNumber = Integer.parseInt(request.getParameter("rowNumber"));
		int columnNumber = Integer.parseInt(request.getParameter("columnNumber"));

		GridWorksheetCollection gridWorksheetCollection = gridweb.getWorkSheets();
		GridWorksheet gridWorksheet = gridWorksheetCollection.get(gridWorksheetCollection.getActiveSheetIndex());
		gridWorksheet.freezePanes(row, column, rowNumber, columnNumber);
	}

	public void unfreezePane(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		
		GridWorksheetCollection gridWorksheetCollection = gridweb.getWorkSheets();
		GridWorksheet gridWorksheet = gridWorksheetCollection.get(gridWorksheetCollection.getActiveSheetIndex());
		gridWorksheet.unFreezePanes();
	}

	public void customHeaders(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		
		GridWorksheetCollection gridWorksheetCollection = gridweb.getWorkSheets();
		gridWorksheetCollection.clear();
		int index = gridWorksheetCollection.add();
		// gridWorkSheet
		GridWorksheet gridWorkSheet = gridWorksheetCollection.get(index);
		gridWorkSheet.setColumnCaption(0, "Product");
		gridWorkSheet.setColumnCaption(1, "Category");
		gridWorkSheet.setColumnCaption(2, "Price");

		GridCells gridCells = gridWorkSheet.getCells();
		gridCells.get("A1").setValue("Aniseed Syrup");
		gridCells.get("A2").setValue("Boston Crab Meat");
		gridCells.get("A3").setValue("Chang");

		gridCells.get("B1").setValue("Condiments");
		gridCells.get("B2").setValue("Seafood");
		gridCells.get("B3").setValue("Beverages");

	}

	public void loadDateTimeFile(GridWebBean gridweb,HttpServletRequest request, HttpServletResponse response) {
		 
		try {
			super.reloadfile(gridweb,request,"datetime.xls");
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
		final GridWorksheet sheet = gridweb.getWorkSheets().get(0);
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

		CellEventHandler ce = new CellEventHandler() {
			public void handleCellEvent(Object sender, CellEventArgs e) {
				if (e.getArgument().toString().equals("A1")) {
					sheet.getCells().sort(1, 0, 20, 4, 0, true,true,false);
				} else if (e.getArgument().toString().equals("B1")) {
					sheet.getCells().sort(1, 0, 20, 4, 1, true,true,false);
				} else if (e.getArgument().toString().equals("C1")) {
					sheet.getCells().sort(1, 0, 20, 4, 2, true,true,false);
				} else if (e.getArgument().toString().equals("D1")) {
					sheet.getCells().sort(1, 0, 20, 4, 3, true,true,false);
				} else if (e.getArgument().toString().equals("1A1")) {
					sheet1.getCells().sort(0, 1, 4, 7, 0, true,true,true);
				} else if (e.getArgument().toString().equals("1A2")) {
					sheet1.getCells().sort(0, 1, 4, 7, 1, true,true,true);
				} else if (e.getArgument().toString().equals("1A3")) {
					sheet1.getCells().sort(0, 1, 4, 7, 2, true,true,true);
				} else if (e.getArgument().toString().equals("1A4")) {
					sheet1.getCells().sort(0, 1, 4, 7, 3, true,true,true);
				}
			}

		};
		gridweb.CellCommand = ce;

	}

	public void events(final GridWebBean gridweb,final HttpServletRequest request, final HttpServletResponse response) {
		this.reload(gridweb,request, response);
		
		gridweb.setPageSize(3);
		final GridWorksheetCollection gridWorksheetCollection = gridweb.getWorkSheets();
		// gridWorkSheet
		final GridWorksheet gridWorkSheet = gridWorksheetCollection.get(gridWorksheetCollection.getActiveSheetIndex());
		gridWorkSheet.getCells().setColumnWidthPixel(0, 180);

		WorkbookEventHandler SubmitCommand = new WorkbookEventHandler() {
			@Override
			public void handleCellEvent(Object arg0, CellEventArgs arg1) {

				// try {
				// request.getRequestDispatcher("/sample/pages/commons/event_info.jsp").forward(request,
				// response);
				// } catch (ServletException e) {
				// e.printStackTrace();
				// } catch (IOException e) {
				// e.printStackTrace();
				// }
				gridWorkSheet.getCells().get("A1").setValue("SubmitCommand");

				// out.println("<script type=\"text/javascript\">");
				// out.println("showMsg(123)");
				// out.println("</script>");
			}
		};
		gridweb.SubmitCommand = SubmitCommand;

		WorkbookEventHandler SaveCommand = new WorkbookEventHandler() {
			@Override
			public void handleCellEvent(Object arg0, CellEventArgs arg1) {
				gridWorkSheet.getCells().get("A1").setValue("SaveCommand");
			}
		};
		gridweb.SaveCommand = SaveCommand;

		WorkbookEventHandler UndoCommand = new WorkbookEventHandler() {
			@Override
			public void handleCellEvent(Object arg0, CellEventArgs arg1) {
				gridWorkSheet.getCells().get("A1").setValue("UndoCommand");
			}
		};
		gridweb.UndoCommand = UndoCommand;

		WorkbookEventHandler SheetTabClick = new WorkbookEventHandler() {
			@Override
			public void handleCellEvent(Object arg0, CellEventArgs arg1) {
				gridWorkSheet.getCells().get("A1").setValue("SheetTabClick");
			}
		};
		gridweb.SheetTabClick = SheetTabClick;

		WorkbookEventHandler SheetTabChange = new WorkbookEventHandler() {
			@Override
			public void handleCellEvent(Object arg0, CellEventArgs arg1) {
				GridWorksheet gridWorkSheet = gridWorksheetCollection.get(gridWorksheetCollection.getActiveSheetIndex());
				gridWorkSheet.getCells().get("A1").setValue("SheetTabChange");
			}
		};
		// gridweb.SheetTabChange = SheetTabChange;

		CellErrorHandler CellError = new CellErrorHandler() {
			@Override
			public void handleCellEvent(Object arg0, GridCellException arg1, OnErrorActionQuery arg2) {
				gridWorkSheet.getCells().get("A1").setValue("CellError");
			}
		};
		// gridweb.CellError = CellError;

		CustomCommandEventHandler CustomCommand = new CustomCommandEventHandler() {
			@Override
			public void handleCellEvent(Object arg0, String arg1) {
				gridWorkSheet.getCells().get("A1").setValue("CustomCommand");
			}
		};
		gridweb.CustomCommand = CustomCommand;

		RowColumnEventHandler RowDoubleClick = new RowColumnEventHandler() {
			@Override
			public void handleCellEvent(Object arg0, RowColumnEventArgs arg1) {
				gridWorkSheet.getCells().get("A1").setValue("RowDoubleClick");
			}
		};
		gridweb.RowDoubleClick = RowDoubleClick;

		RowColumnEventHandler ColumnDoubleClick = new RowColumnEventHandler() {
			@Override
			public void handleCellEvent(Object arg0, RowColumnEventArgs arg1) {
				gridWorkSheet.getCells().get("A1").setValue("ColumnDoubleClick");
			}
		};
		gridweb.ColumnDoubleClick = ColumnDoubleClick;

		CellEventHandler CellDoubleClick = new CellEventHandler() {
			@Override
			public void handleCellEvent(Object arg0, CellEventArgs arg1) {
				gridWorkSheet.getCells().get("A1").setValue("CellDoubleClick");
			}
		};
		gridweb.CellDoubleClick = CellDoubleClick;

		CellEventStringHandler CellClickOnAjax = new CellEventStringHandler() {
			@Override
			public String handleCellEvent(Object arg0, CellEventArgs arg1) {
				gridWorkSheet.getCells().get("A1").setValue("CellClickOnAjax");
				return null;
			}
		};
		gridweb.CellClickOnAjax = CellClickOnAjax;

		RowColumnEventHandler RowInserted = new RowColumnEventHandler() {
			@Override
			public void handleCellEvent(Object arg0, RowColumnEventArgs arg1) {
				gridWorkSheet.getCells().get("A1").setValue("RowInserted");
			}
		};
		gridweb.RowInserted = RowInserted;

		RowColumnEventHandler RowDeleted = new RowColumnEventHandler() {
			@Override
			public void handleCellEvent(Object arg0, RowColumnEventArgs arg1) {
				gridWorkSheet.getCells().get("A1").setValue("RowDeleted");
			}
		};
		gridweb.RowDeleted = RowDeleted;

		RowColumnEventHandler RowDeleting = new RowColumnEventHandler() {
			@Override
			public void handleCellEvent(Object arg0, RowColumnEventArgs arg1) {
				gridWorkSheet.getCells().get("A1").setValue("RowDeleting");
			}
		};
		gridweb.RowDeleting = RowDeleting;

		RowColumnEventHandler ColumnInserted = new RowColumnEventHandler() {
			@Override
			public void handleCellEvent(Object arg0, RowColumnEventArgs arg1) {
				gridWorkSheet.getCells().get("A1").setValue("ColumnInserted");
			}
		};
		gridweb.ColumnInserted = ColumnInserted;

		RowColumnEventHandler ColumnDeleted = new RowColumnEventHandler() {
			@Override
			public void handleCellEvent(Object arg0, RowColumnEventArgs arg1) {
				gridWorkSheet.getCells().get("A1").setValue("ColumnDeleted");
			}
		};
		gridweb.ColumnDeleted = ColumnDeleted;

		RowColumnEventHandler ColumnDeleting = new RowColumnEventHandler() {
			@Override
			public void handleCellEvent(Object arg0, RowColumnEventArgs arg1) {
				gridWorkSheet.getCells().get("A1").setValue("ColumnDeleting");
			}
		};
		gridweb.ColumnDeleting = ColumnDeleting;

		CellEventHandler CellCommand = new CellEventHandler() {
			@Override
			public void handleCellEvent(Object arg0, CellEventArgs arg1) {
				gridWorkSheet.getCells().get("A1").setValue("CellCommand");
			}
		};
		gridweb.CellCommand = CellCommand;

		WorkbookEventHandler PageIndexChanged = new WorkbookEventHandler() {
			@Override
			public void handleCellEvent(Object arg0, CellEventArgs arg1) {
				int row=(gridweb.getCurrentPageIndex())*gridweb.getPageSize();
				gridWorkSheet.getCells().get(row,0).setValue("PageIndexChanged"+(gridweb.getCurrentPageIndex()+1));
			}
		};
		gridweb.PageIndexChanged = PageIndexChanged;
	}

}
