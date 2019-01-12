package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.TextBox;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class AccessTextBoxName {
	public static void main(String[] args) throws Exception {

		Workbook workbook = new Workbook();

		Worksheet sheet = workbook.getWorksheets().get(0);

		int idx = sheet.getTextBoxes().add(10, 10, 10, 10);

		// Create a texbox with some text and assign it some name
		TextBox tb1 = sheet.getTextBoxes().get(idx);
		tb1.setName("MyTextBox");
		tb1.setText("This is MyTextBox");

		// Access the same textbox via its name
		TextBox tb2 = sheet.getTextBoxes().get("MyTextBox");

		// Displaying the text of the textbox accessed by its name
		System.out.println(tb2.getText());

	}

}
