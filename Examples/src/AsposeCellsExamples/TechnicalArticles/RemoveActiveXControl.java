package AsposeCellsExamples.TechnicalArticles;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class RemoveActiveXControl {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		//Source directory
		String sourceDir = Utils.Get_SourceDirectory();

		//Output directory
		String outputDir = Utils.Get_OutputDirectory();

		// Create a workbook
		Workbook workbook = new Workbook(sourceDir + "sampleUpdateActiveXComboBoxControl.xlsx");

		// Access first shape from first worksheet
		Shape shape = workbook.getWorksheets().get(0).getShapes().get(0);

		// Access ActiveX ComboBox Control and update its value
		if (shape.getActiveXControl() != null)
		{
			// Remove Shape ActiveX Control
			shape.removeActiveXControl();
		}

		// Save the workbook
		workbook.save(outputDir + "RemoveActiveXControl_out.xlsx");
		// ExEnd:1

		System.out.println("RemoveActiveXControl executed successfully.");
	}
}
