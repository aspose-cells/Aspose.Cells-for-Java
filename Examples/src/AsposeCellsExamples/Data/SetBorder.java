package AsposeCellsExamples.Data;

import com.aspose.cells.BorderType;
import com.aspose.cells.CellArea;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Color;
import com.aspose.cells.FormatCondition;
import com.aspose.cells.FormatConditionCollection;
import com.aspose.cells.FormatConditionType;
import com.aspose.cells.OperatorType;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class SetBorder {
	public static void main(String[] args) throws Exception {
		// Instantiating a Workbook object
	    Workbook workbook = new Workbook();
	    Worksheet sheet = workbook.getWorksheets().get(0);

	    // Adds an empty conditional formatting
	    int index = sheet.getConditionalFormattings().add();
	    FormatConditionCollection fcs = sheet.getConditionalFormattings().get(index);

	    // Sets the conditional format range.
	    CellArea ca = new CellArea();
	    ca.StartRow = 0;
	    ca.EndRow = 5;
	    ca.StartColumn = 0;
	    ca.EndColumn = 3;
	    fcs.addArea(ca);
	    // Adds condition.
	    int conditionIndex = fcs.addCondition(FormatConditionType.CELL_VALUE, OperatorType.BETWEEN, "50", "100");
	    // Sets the background color.
	    FormatCondition fc = fcs.get(conditionIndex);
	    
		Style style = fc.getStyle();
		style.setBorder(BorderType.LEFT_BORDER, CellBorderType.DASHED, Color.fromArgb(0, 255, 255));
		style.setBorder(BorderType.TOP_BORDER, CellBorderType.DASHED, Color.fromArgb(0, 255, 255));
		style.setBorder(BorderType.RIGHT_BORDER, CellBorderType.DASHED, Color.fromArgb(0, 255, 255));
		style.setBorder(BorderType.RIGHT_BORDER, CellBorderType.DASHED, Color.fromArgb(255, 255, 0));
		fc.setStyle(style);
	}
}
