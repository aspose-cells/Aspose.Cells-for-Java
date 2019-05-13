package AsposeCellsExamples.Files.Utility;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class AccessWebExtensionInformation {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AccessWebExtensionInformation.class) + "Files/Utility/";

		Workbook workbook = new Workbook(dataDir + "WebExtensionsSample.xlsx");

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
