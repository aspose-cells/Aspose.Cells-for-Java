package com.aspose.cells.examples.articles;

import com.aspose.cells.examples.Utils;
import com.aspose.gridweb.CustomCommandButton;

public class CreatingCustomCommandButtons {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CreatingCustomCommandButtons.class) + "articles/";

		gridweb.setReqRes(request, response);
		// gridweb.importExcelFile(filePath);

		// Instantiating a CustomCommandButton object
		CustomCommandButton button = new CustomCommandButton();

		// Setting the command for button
		button.setCommand("MyButton");

		// Setting text of the button
		button.setText("MyButton");

		// Setting tooltip of the button
		button.setToolTip("My Custom Command Button");

		// Setting image URL of the button
		button.setImageUrl("icon.png");

		// Adding button to CustomCommandButtons collection of GridWeb
		gridweb.getCustomCommandButtons().add(button);

	}

}
