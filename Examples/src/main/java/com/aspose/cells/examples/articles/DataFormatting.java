package com.aspose.cells.examples.articles;

import com.aspose.cells.BackgroundType;
import com.aspose.cells.BorderType;
import com.aspose.cells.Cell;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Column;
import com.aspose.cells.Range;
import com.aspose.cells.Row;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Style;
import com.aspose.cells.StyleFlag;
import com.aspose.cells.TextAlignmentType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class DataFormatting {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DataFormatting.class) + "articles/";
		// Create a new Workbook.
		Workbook workbook = new Workbook();
		// Obtain the cells of the first worksheet.
		Cells cells = workbook.getWorksheets().get(0).getCells();

		// Input the title on B1 cell.
		cells.get("B1").putValue("Western Product Sales 2006");

		// Insert some column headings in the second row.
		Cell cell = cells.get("B2");
		cell.putValue("January");
		cell = cells.get("C2");
		cell.putValue("February");
		cell = cells.get("D2");
		cell.putValue("March");
		cell = cells.get("E2");
		cell.putValue("April");
		cell = cells.get("F2");
		cell.putValue("May");
		cell = cells.get("G2");
		cell.putValue("June");
		cell = cells.get("H2");
		cell.putValue("July");
		cell = cells.get("I2");
		cell.putValue("August");
		cell = cells.get("J2");
		cell.putValue("September");
		cell = cells.get("K2");
		cell.putValue("October");
		cell = cells.get("L2");
		cell.putValue("November");
		cell = cells.get("M2");
		cell.putValue("December");
		cell = cells.get("N2");
		cell.putValue("Total");

		// Insert product names.
		cells.get("A3").putValue("Biscuits");
		cells.get("A4").putValue("Coffee");
		cells.get("A5").putValue("Tofu");
		cells.get("A6").putValue("Ikura");
		cells.get("A7").putValue("Choclade");
		cells.get("A8").putValue("Maxilaku");
		cells.get("A9").putValue("Scones");
		cells.get("A10").putValue("Sauce");
		cells.get("A11").putValue("Syrup");
		cells.get("A12").putValue("Spegesild");
		cells.get("A13").putValue("Filo Mix");
		cells.get("A14").putValue("Pears");
		cells.get("A15").putValue("Konbu");
		cells.get("A16").putValue("Kaviar");
		cells.get("A17").putValue("Zaanse");
		cells.get("A18").putValue("Cabrales");
		cells.get("A19").putValue("Gnocchi");
		cells.get("A20").putValue("Wimmers");
		cells.get("A21").putValue("Breads");
		cells.get("A22").putValue("Lager");
		cells.get("A23").putValue("Gravad");
		cells.get("A24").putValue("Telino");
		cells.get("A25").putValue("Pavlova");
		cells.get("A26").putValue("Total");

		// Input porduct sales data (B3:M25).
		cells.get("B3").putValue(5000);
		cells.get("C3").putValue(4500);
		cells.get("D3").putValue(6010);
		cells.get("E3").putValue(7230);
		cells.get("F3").putValue(5400);
		cells.get("G3").putValue(5030);
		cells.get("H3").putValue(3000);
		cells.get("I3").putValue(6000);
		cells.get("J3").putValue(9000);
		cells.get("K3").putValue(3300);
		cells.get("L3").putValue(2500);
		cells.get("M3").putValue(5510);

		cells.get("B4").putValue(4000);
		cells.get("C4").putValue(2500);
		cells.get("D4").putValue(6000);
		cells.get("E4").putValue(5300);
		cells.get("F4").putValue(7400);
		cells.get("G4").putValue(7030);
		cells.get("H4").putValue(4000);
		cells.get("I4").putValue(4000);
		cells.get("J4").putValue(5500);
		cells.get("K4").putValue(4500);
		cells.get("L4").putValue(2500);
		cells.get("M4").putValue(2510);

		cells.get("B5").putValue(2000);
		cells.get("C5").putValue(1500);
		cells.get("D5").putValue(3000);
		cells.get("E5").putValue(2500);
		cells.get("F5").putValue(3400);
		cells.get("G5").putValue(4030);
		cells.get("H5").putValue(2000);
		cells.get("I5").putValue(2000);
		cells.get("J5").putValue(1500);
		cells.get("K5").putValue(2200);
		cells.get("L5").putValue(2100);
		cells.get("M5").putValue(2310);

		cells.get("B6").putValue(1000);
		cells.get("C6").putValue(1300);
		cells.get("D6").putValue(2000);
		cells.get("E6").putValue(2600);
		cells.get("F6").putValue(5400);
		cells.get("G6").putValue(2030);
		cells.get("H6").putValue(2100);
		cells.get("I6").putValue(4000);
		cells.get("J6").putValue(6500);
		cells.get("K6").putValue(5600);
		cells.get("L6").putValue(3300);
		cells.get("M6").putValue(5110);

		cells.get("B7").putValue(3000);
		cells.get("C7").putValue(3500);
		cells.get("D7").putValue(1000);
		cells.get("E7").putValue(4500);
		cells.get("F7").putValue(5400);
		cells.get("G7").putValue(2030);
		cells.get("H7").putValue(3000);
		cells.get("I7").putValue(3000);
		cells.get("J7").putValue(4500);
		cells.get("K7").putValue(6000);
		cells.get("L7").putValue(3000);
		cells.get("M7").putValue(3000);

		cells.get("B8").putValue(5000);
		cells.get("C8").putValue(5500);
		cells.get("D8").putValue(5000);
		cells.get("E8").putValue(5500);
		cells.get("F8").putValue(5400);
		cells.get("G8").putValue(5030);
		cells.get("H8").putValue(5000);
		cells.get("I8").putValue(2500);
		cells.get("J8").putValue(5500);
		cells.get("K8").putValue(5200);
		cells.get("L8").putValue(5500);
		cells.get("M8").putValue(2510);

		cells.get("B9").putValue(4100);
		cells.get("C9").putValue(1500);
		cells.get("D9").putValue(1000);
		cells.get("E9").putValue(2300);
		cells.get("F9").putValue(3300);
		cells.get("G9").putValue(4030);
		cells.get("H9").putValue(5000);
		cells.get("I9").putValue(6000);
		cells.get("J9").putValue(3500);
		cells.get("K9").putValue(4300);
		cells.get("L9").putValue(2300);
		cells.get("M9").putValue(2110);

		cells.get("B10").putValue(2000);
		cells.get("C10").putValue(2300);
		cells.get("D10").putValue(3000);
		cells.get("E10").putValue(3300);
		cells.get("F10").putValue(3400);
		cells.get("G10").putValue(3030);
		cells.get("H10").putValue(3000);
		cells.get("I10").putValue(3000);
		cells.get("J10").putValue(3500);
		cells.get("K10").putValue(3500);
		cells.get("L10").putValue(3500);
		cells.get("M10").putValue(3510);

		cells.get("B11").putValue(4400);
		cells.get("C11").putValue(4500);
		cells.get("D11").putValue(4000);
		cells.get("E11").putValue(4300);
		cells.get("F11").putValue(4400);
		cells.get("G11").putValue(4030);
		cells.get("H11").putValue(5000);
		cells.get("I11").putValue(5000);
		cells.get("J11").putValue(4500);
		cells.get("K11").putValue(4400);
		cells.get("L11").putValue(4400);
		cells.get("M11").putValue(4510);

		cells.get("B12").putValue(2000);
		cells.get("C12").putValue(1500);
		cells.get("D12").putValue(3000);
		cells.get("E12").putValue(2300);
		cells.get("F12").putValue(3400);
		cells.get("G12").putValue(3030);
		cells.get("H12").putValue(3000);
		cells.get("I12").putValue(3000);
		cells.get("J12").putValue(2500);
		cells.get("K12").putValue(2500);
		cells.get("L12").putValue(1500);
		cells.get("M12").putValue(5110);

		cells.get("B13").putValue(4000);
		cells.get("C13").putValue(1400);
		cells.get("D13").putValue(1400);
		cells.get("E13").putValue(3300);
		cells.get("F13").putValue(3300);
		cells.get("G13").putValue(3730);
		cells.get("H13").putValue(3800);
		cells.get("I13").putValue(3600);
		cells.get("J13").putValue(2600);
		cells.get("K13").putValue(4600);
		cells.get("L13").putValue(1400);
		cells.get("M13").putValue(2660);

		cells.get("B14").putValue(3000);
		cells.get("C14").putValue(3500);
		cells.get("D14").putValue(3333);
		cells.get("E14").putValue(2330);
		cells.get("F14").putValue(3430);
		cells.get("G14").putValue(3040);
		cells.get("H14").putValue(3040);
		cells.get("I14").putValue(3030);
		cells.get("J14").putValue(1509);
		cells.get("K14").putValue(4503);
		cells.get("L14").putValue(1503);
		cells.get("M14").putValue(3113);

		cells.get("B15").putValue(2010);
		cells.get("C15").putValue(1520);
		cells.get("D15").putValue(3030);
		cells.get("E15").putValue(2320);
		cells.get("F15").putValue(3410);
		cells.get("G15").putValue(3000);
		cells.get("H15").putValue(3000);
		cells.get("I15").putValue(3020);
		cells.get("J15").putValue(2520);
		cells.get("K15").putValue(2520);
		cells.get("L15").putValue(1520);
		cells.get("M15").putValue(5120);

		cells.get("B16").putValue(2220);
		cells.get("C16").putValue(1200);
		cells.get("D16").putValue(3220);
		cells.get("E16").putValue(1320);
		cells.get("F16").putValue(1400);
		cells.get("G16").putValue(1030);
		cells.get("H16").putValue(3200);
		cells.get("I16").putValue(3020);
		cells.get("J16").putValue(2100);
		cells.get("K16").putValue(2100);
		cells.get("L16").putValue(1100);
		cells.get("M16").putValue(5210);

		cells.get("B17").putValue(1444);
		cells.get("C17").putValue(1540);
		cells.get("D17").putValue(3040);
		cells.get("E17").putValue(2340);
		cells.get("F17").putValue(1440);
		cells.get("G17").putValue(1030);
		cells.get("H17").putValue(3000);
		cells.get("I17").putValue(4000);
		cells.get("J17").putValue(4500);
		cells.get("K17").putValue(2500);
		cells.get("L17").putValue(4500);
		cells.get("M17").putValue(5550);

		cells.get("B18").putValue(4000);
		cells.get("C18").putValue(5500);
		cells.get("D18").putValue(3000);
		cells.get("E18").putValue(3300);
		cells.get("F18").putValue(3330);
		cells.get("G18").putValue(5330);
		cells.get("H18").putValue(3400);
		cells.get("I18").putValue(3040);
		cells.get("J18").putValue(2540);
		cells.get("K18").putValue(4500);
		cells.get("L18").putValue(4500);
		cells.get("M18").putValue(2110);

		cells.get("B19").putValue(2000);
		cells.get("C19").putValue(2500);
		cells.get("D19").putValue(3200);
		cells.get("E19").putValue(3200);
		cells.get("F19").putValue(2330);
		cells.get("G19").putValue(5230);
		cells.get("H19").putValue(2400);
		cells.get("I19").putValue(3240);
		cells.get("J19").putValue(2240);
		cells.get("K19").putValue(4300);
		cells.get("L19").putValue(4100);
		cells.get("M19").putValue(2310);

		cells.get("B20").putValue(7000);
		cells.get("C20").putValue(8500);
		cells.get("D20").putValue(8000);
		cells.get("E20").putValue(5300);
		cells.get("F20").putValue(6330);
		cells.get("G20").putValue(7330);
		cells.get("H20").putValue(3600);
		cells.get("I20").putValue(3940);
		cells.get("J20").putValue(2940);
		cells.get("K20").putValue(4600);
		cells.get("L20").putValue(6500);
		cells.get("M20").putValue(8710);

		cells.get("B21").putValue(4000);
		cells.get("C21").putValue(4500);
		cells.get("D21").putValue(2000);
		cells.get("E21").putValue(2200);
		cells.get("F21").putValue(2000);
		cells.get("G21").putValue(3000);
		cells.get("H21").putValue(3000);
		cells.get("I21").putValue(3000);
		cells.get("J21").putValue(4330);
		cells.get("K21").putValue(4420);
		cells.get("L21").putValue(4500);
		cells.get("M21").putValue(1330);

		cells.get("B22").putValue(2050);
		cells.get("C22").putValue(3520);
		cells.get("D22").putValue(1030);
		cells.get("E22").putValue(2000);
		cells.get("F22").putValue(3000);
		cells.get("G22").putValue(2000);
		cells.get("H22").putValue(2010);
		cells.get("I22").putValue(2210);
		cells.get("J22").putValue(2230);
		cells.get("K22").putValue(4240);
		cells.get("L22").putValue(3330);
		cells.get("M22").putValue(2000);

		cells.get("B23").putValue(1222);
		cells.get("C23").putValue(3000);
		cells.get("D23").putValue(3020);
		cells.get("E23").putValue(2770);
		cells.get("F23").putValue(3011);
		cells.get("G23").putValue(2000);
		cells.get("H23").putValue(6000);
		cells.get("I23").putValue(9000);
		cells.get("J23").putValue(4000);
		cells.get("K23").putValue(2000);
		cells.get("L23").putValue(5000);
		cells.get("M23").putValue(6333);

		cells.get("B24").putValue(1000);
		cells.get("C24").putValue(2000);
		cells.get("D24").putValue(1000);
		cells.get("E24").putValue(1300);
		cells.get("F24").putValue(1330);
		cells.get("G24").putValue(1390);
		cells.get("H24").putValue(1600);
		cells.get("I24").putValue(1900);
		cells.get("J24").putValue(1400);
		cells.get("K24").putValue(1650);
		cells.get("L24").putValue(1520);
		cells.get("M24").putValue(1910);

		cells.get("B25").putValue(2000);
		cells.get("C25").putValue(6600);
		cells.get("D25").putValue(3300);
		cells.get("E25").putValue(8300);
		cells.get("F25").putValue(2000);
		cells.get("G25").putValue(3000);
		cells.get("H25").putValue(6000);
		cells.get("I25").putValue(4000);
		cells.get("J25").putValue(7000);
		cells.get("K25").putValue(2000);
		cells.get("L25").putValue(5000);
		cells.get("M25").putValue(5500);

		// Add Month wise Summary formulas.
		cells.get("B26").setFormula("=SUM(B3:B25)");
		cells.get("C26").setFormula("=SUM(C3:C25)");
		cells.get("D26").setFormula("=SUM(D3:D25)");
		cells.get("E26").setFormula("=SUM(E3:E25)");
		cells.get("F26").setFormula("=SUM(F3:F25)");
		cells.get("G26").setFormula("=SUM(G3:G25)");
		cells.get("H26").setFormula("=SUM(H3:H25)");
		cells.get("I26").setFormula("=SUM(I3:I25)");
		cells.get("J26").setFormula("=SUM(J3:J25)");
		cells.get("K26").setFormula("=SUM(K3:K25)");
		cells.get("L26").setFormula("=SUM(L3:L25)");
		cells.get("M26").setFormula("=SUM(M3:M25)");

		// Add Product wise Summary formulas.
		cells.get("N3").setFormula("=SUM(B3:M3)");
		cells.get("N4").setFormula("=SUM(B4:M4)");
		cells.get("N5").setFormula("=SUM(B5:M5)");
		cells.get("N6").setFormula("=SUM(B6:M6)");
		cells.get("N7").setFormula("=SUM(B7:M7)");
		cells.get("N8").setFormula("=SUM(B8:M8)");
		cells.get("N9").setFormula("=SUM(B9:M9)");
		cells.get("N10").setFormula("=SUM(B10:M10)");
		cells.get("N11").setFormula("=SUM(B11:M11)");
		cells.get("N12").setFormula("=SUM(B12:M12)");
		cells.get("N13").setFormula("=SUM(B13:M13)");
		cells.get("N14").setFormula("=SUM(B14:M14)");
		cells.get("N15").setFormula("=SUM(B15:M15)");
		cells.get("N16").setFormula("=SUM(B16:M16)");
		cells.get("N17").setFormula("=SUM(B17:M17)");
		cells.get("N18").setFormula("=SUM(B18:M18)");
		cells.get("N19").setFormula("=SUM(B19:M19)");
		cells.get("N20").setFormula("=SUM(B20:M20)");
		cells.get("N21").setFormula("=SUM(B21:M21)");
		cells.get("N22").setFormula("=SUM(B22:M22)");
		cells.get("N23").setFormula("=SUM(B23:M23)");
		cells.get("N24").setFormula("=SUM(B24:M24)");
		cells.get("N25").setFormula("=SUM(B25:M25)");
		// Add Grand Total.
		cells.get("N26").setFormula("=SUM(N3:N25)");

		// Define a style object
		Style stl0 = workbook.createStyle();
		// Set a custom shading color of the cells.
		stl0.setForegroundColor(Color.fromArgb(155, 204, 255));
		// Set the solid background fill.
		stl0.setPattern(BackgroundType.SOLID);
		// Set a font.
		stl0.getFont().setName("Trebuchet MS");
		// Set the size.
		stl0.getFont().setSize(18);
		// Set the font text color.
		stl0.getFont().setColor(Color.getMaroon());
		// Set it bold
		stl0.getFont().setBold(true);
		// Set it italic.
		stl0.getFont().setItalic(true);
		// Define a style flag struct.
		StyleFlag flag = new StyleFlag();
		// Apply cell shading.
		flag.setCellShading(true);
		// Apply font.
		flag.setFontName(true);
		// Apply font size.
		flag.setFontSize(true);
		// Apply font color.
		flag.setFontColor(true);
		// Apply bold font.
		flag.setFontBold(true);
		// Apply italic attribute.
		flag.setFontItalic(true);
		// Get the first row in the first worksheet.
		Row row = workbook.getWorksheets().get(0).getCells().getRows().get(0);
		// Apply the style to it.
		row.applyStyle(stl0, flag);

		// Obtain the cells of the first worksheet.
		cells = workbook.getWorksheets().get(0).getCells();
		// Set the height of the first row.
		cells.setRowHeight(0, 30);

		// Define a style object adding a new style
		// to the collection list.
		Style stl1 = workbook.createStyle();
		// Set the rotation angle of the text.
		stl1.setRotationAngle(45);
		// Set the custom fill color of the cells.
		stl1.setForegroundColor(Color.fromArgb(0, 51, 105));
		// Set the solid background pattern for fill.
		stl1.setPattern(BackgroundType.SOLID);
		// Set the left border line style.
		stl1.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
		// Set the left border line color.
		stl1.getBorders().getByBorderType(BorderType.LEFT_BORDER).setColor(Color.getWhite());
		// Set the horizontal alignment to center aligned.
		stl1.setHorizontalAlignment(TextAlignmentType.CENTER);
		// Set the vertical alignment to center aligned.
		stl1.setVerticalAlignment(TextAlignmentType.CENTER);
		// Set the font.
		stl1.getFont().setName("Times New Roman");
		// Set the font size.
		stl1.getFont().setSize(10);
		// Set the font color.
		stl1.getFont().setColor(Color.getWhite());
		// Set the bold attribute.
		stl1.getFont().setBold(true);
		// Set the style flag struct.
		flag = new StyleFlag();
		// Apply the left border.
		flag.setLeftBorder(true);
		// Apply text rotation orientation.
		flag.setRotation(true);
		// Apply fill color of cells.
		flag.setCellShading(true);
		// Apply horizontal alignment.
		flag.setHorizontalAlignment(true);
		// Apply vertical alignment.
		flag.setVerticalAlignment(true);
		// Apply the font.
		flag.setFontName(true);
		// Apply the font size.
		flag.setFontSize(true);
		// Apply the font color.
		flag.setFontColor(true);
		// Apply the bold attribute.
		flag.setFontBold(true);
		// Get the second row of the first worksheet.
		row = workbook.getWorksheets().get(0).getCells().getRows().get(1);
		// Apply the style to it.
		row.applyStyle(stl1, flag);

		// Set the height of the second row.
		cells.setRowHeight(1, 48);

		// Define a style object adding a new style
		// to the collection list.
		Style stl2 = workbook.createStyle();
		// Set the custom cell shading color.
		stl2.setForegroundColor(Color.fromArgb(155, 204, 255));
		// Set the solid background pattern for fill color.
		stl2.setPattern(BackgroundType.SOLID);
		// Set the font.
		stl2.getFont().setName("Trebuchet MS");
		// Set the font color.
		stl2.getFont().setColor(Color.getMaroon());
		// Set the font size.
		stl2.getFont().setSize(10);
		// Set the style flag struct.
		flag = new StyleFlag();
		// Apply cell shading.
		flag.setCellShading(true);
		// Apply the font.
		flag.setFontName(true);
		// Apply the font color.
		flag.setFontColor(true);
		// Apply the font size.
		flag.setFontSize(true);

		// Get the first column in the first worksheet.
		Column col = workbook.getWorksheets().get(0).getCells().getColumns().get(0);
		// Apply the style to it.
		col.applyStyle(stl2, flag);

		// Define a style object adding a new style
		// to the collection list.
		Style stl3 = workbook.createStyle();
		// Set the custom cell filling color.
		stl3.setForegroundColor(Color.fromArgb(124, 199, 72));
		// Set the solid background pattern for fill color.
		stl3.setPattern(BackgroundType.SOLID);
		// Apply the style to A2 cell.
		cells.get("A2").setStyle(stl3);

		// Define a style object adding a new style
		// to the collection list.
		Style stl4 = workbook.createStyle();
		// Set the custom font text color.
		stl4.getFont().setColor(Color.fromArgb(0, 51, 105));
		// Set the bottom border line style.
		stl4.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
		// Set the bottom border line color to custom color.
		stl4.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setColor(Color.fromArgb(124, 199, 72));
		// Set the background fill color of the cells.
		stl4.setForegroundColor(Color.getWhite());
		// Set the solid fill color pattern.
		stl4.setPattern(BackgroundType.SOLID);
		// Set custom number format.
		stl4.setCustom("$#,##0.0");
		// Set a style flag struct.
		flag = new StyleFlag();
		// Apply font color.
		flag.setFontColor(true);
		// Apply cell shading color.
		flag.setCellShading(true);
		// Apply custom number format.
		flag.setNumberFormat(true);
		// Apply bottom border.
		flag.setBottomBorder(true);

		// Define a style object adding a new style
		// to the collection list.
		Style stl5 = workbook.createStyle();
		// Set the bottom borde line style.
		stl5.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
		// Set the bottom border line color.
		stl5.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setColor(Color.fromArgb(124, 199, 72));
		// Set the custom shading color of the cells.
		stl5.setForegroundColor(Color.fromArgb(250, 250, 200));
		// Set the solid background pattern for fillment color.
		stl5.setPattern(BackgroundType.SOLID);
		// Set custom number format.
		stl5.setCustom("$#,##0.0");
		// Set font text color.
		stl5.getFont().setColor(Color.getMaroon());

		// Create a named range of cells (B3:M25)in the first worksheet.
		Range range = workbook.getWorksheets().get(0).getCells().createRange("B3", "M25");
		// Name the range.
		range.setName("MyRange");
		// Apply the style to cells in the named range.
		range.applyStyle(stl4, flag);

		// Apply different style to alternative rows
		// in the range.
		for (int i = 0; i <= 22; i++) {
			for (int j = 0; j < 12; j++) {
				if (i % 2 == 0) {
					range.get(i, j).setStyle(stl5);

				}

			}
		}

		// Define a style object adding a new style
		// to the collection list.
		Style stl6 = workbook.createStyle();
		// Set the custom fill color of the cells.
		stl6.setForegroundColor(Color.fromArgb(0, 51, 105));
		// Set the background pattern for fill color.
		stl6.setPattern(BackgroundType.SOLID);
		// Set the font.
		stl6.getFont().setName("Arial");
		// Set the font size.
		stl6.getFont().setSize(10);
		// Set the font color
		stl6.getFont().setColor(Color.getWhite());
		// Set the text bold.
		stl6.getFont().setBold(true);
		// Set the custom number format.
		stl6.setCustom("$#,##0.0");
		// Set the style flag struct.
		flag = new StyleFlag();
		// Apply cell shading.
		flag.setCellShading(true);
		// Apply the arial font.
		flag.setFontName(true);
		// Apply the font size.
		flag.setFontSize(true);
		// Apply the font color.
		flag.setFontColor(true);
		// Apply the bold attribute.
		flag.setFontBold(true);
		// Apply the number format.
		flag.setNumberFormat(true);
		// Get the 26th row in the first worksheet which produces totals.
		row = workbook.getWorksheets().get(0).getCells().getRows().get(25);
		// Apply the style to it.
		row.applyStyle(stl6, flag);
		// Now apply this style to those cells (N3:N25) which
		// has product wise sales totals.
		for (int i = 2; i < 25; i++) {
			cells.get(1, 13).setStyle(stl6);

		}
		// Set N column's width to fit the contents.
		workbook.getWorksheets().get(0).getCells().setColumnWidth(13, 9.33);

		workbook.save(dataDir + "DataFormatting_out.xlsx", SaveFormat.XLSX);

	}
}
