package AsposeCellsExamples.Workbook;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class AccessWebExtensionInformation {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the source directory.
		String sourceDir = Utils.Get_SourceDirectory();

		Workbook workbook = new Workbook(sourceDir + "WebExtensionsSample.xlsx");

		WebExtensionTaskPaneCollection taskPanes = workbook.getWorksheets().getWebExtensionTaskPanes();
		
		for (Object obj : taskPanes)
        {
			WebExtensionTaskPane taskPane = (WebExtensionTaskPane) obj;
			
			System.out.println("Width: " + taskPane.getWidth());
			System.out.println("IsVisible: " + taskPane.isVisible());
			System.out.println("IsLocked: " + taskPane.isLocked());
			System.out.println("DockState: " + taskPane.getDockState());
			System.out.println("StoreName: " + taskPane.getWebExtension().getReference().getStoreName());
			System.out.println("StoreType: " + taskPane.getWebExtension().getReference().getStoreType());
			System.out.println("WebExtension.Id: " + taskPane.getWebExtension().getId());
        }
        // ExEnd:1

		System.out.println("AccessWebExtensionInformation executed successfully.");
	}
}
