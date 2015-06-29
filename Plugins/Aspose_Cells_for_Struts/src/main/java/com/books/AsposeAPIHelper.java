package com.books;

import java.util.List;
import java.util.Map;

import javax.servlet.ServletContext;
import javax.servlet.ServletOutputStream;

import com.aspose.cells.Style;
import com.aspose.cells.StyleFlag;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

/**
 * 
 * @author Adeel
 *
 */
public class AsposeAPIHelper {

	/**
	 * Creates excel sheet from list of book provided from grid.
	 * 
	 * @param out
	 *            the current scope OutputStream.
	 * @param books
	 *            books list as map containing attributes.
	 * @param context
	 *            the App ServletContext
	 * @see com.aspose.cells.Workbook
	 */
	public static void createAsposeExcelSheet(ServletOutputStream out,
			List<Map> books, ServletContext context) throws Exception {

		try {
			Workbook workbook = new Workbook();
			// Obtaining the reference of the first worksheet
			Worksheet sheet = workbook.getWorksheets().get(0);

			// Name the sheet
			sheet.setName("Books List");

			com.aspose.cells.Cells cells = sheet.getCells();

			// Setting the values to the cells
			com.aspose.cells.Cell cell = cells.get("F11");
			cell.setValue("Book Id");
			cell = cells.get("G11");
			cell.setValue("Book Name");
			cell = cells.get("H11");
			cell.setValue("AuthorName");
			cell = cells.get("I11");
			cell.setValue("Book Cost");
			Style style1 = sheet.getCells().get("A1").getStyle();
		
			// Set the number format.
			style1.setNumber(14);

			// Set the font color to red color.
			style1.getFont().setColor(com.aspose.cells.Color.getRed());
			style1.getFont().setBold(true);
			// Name the style.
			style1.setName("Heading");
			com.aspose.cells.Range range = cells.createRange("F11", "I11");
			// Initialize styleflag object.
			StyleFlag flag = new StyleFlag();

			// Set all formatting attributes on.
			flag.setAll(true);

			// Apply the style (described above)to the range.
			range.applyStyle(style1, flag);
			style1.update();
			int i = 12;
			for (Map book : books) {
				String bookId = book.get("BookId").toString();
				String bookName = book.get("BookName").toString();
				String bookAuthorName = book.get("AuthorName").toString();
				String bookCost = book.get("BookCost").toString();
				cell = cells.get("F" + i);
				cell.setValue(bookId);
				cell = cells.get("G" + i);
				cell.setValue(bookName);
				cell = cells.get("H" + i);
				cell.setValue(bookAuthorName);
				cell = cells.get("I" + i);
				cell.setValue(bookCost);
				i++;
			}

			// Save the sheet

			workbook.save(out, com.aspose.cells.SaveFormat.XLSX);

		} catch (Exception e) {
			throw new Exception(
					"Aspose: Unable to export to ms excel format.. some error occured",
					e);

		}
	}
}
