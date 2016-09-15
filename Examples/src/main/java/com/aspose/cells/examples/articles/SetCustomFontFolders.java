package com.aspose.cells.examples.articles;

import java.io.File;
import java.nio.file.Files;

import com.aspose.cells.FileFontSource;
import com.aspose.cells.FolderFontSource;
import com.aspose.cells.FontConfigs;
import com.aspose.cells.FontSourceBase;
import com.aspose.cells.MemoryFontSource;
import com.aspose.cells.examples.Utils;

public class SetCustomFontFolders {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SetCustomFontFolders.class) + "articles/";
		// Defining string variables to store paths to font folders & font file
		String fontFolder1 = dataDir + "/Arial";
		String fontFolder2 = dataDir + "/Calibri";
		String fontFile = dataDir + "/Arial/arial.ttf";

		// Setting first font folder with setFontFolder method
		// Second parameter directs the API to search the sub folders for font files
		FontConfigs.setFontFolder(fontFolder1, true);

		// Setting both font folders with setFontFolders method
		// Second parameter prohibits the API to search the sub folders for font files
		FontConfigs.setFontFolders(new String[] { fontFolder1, fontFolder2 }, false);

		// Defining FolderFontSource
		FolderFontSource sourceFolder = new FolderFontSource(fontFolder1, false);

		// Defining FileFontSource
		FileFontSource sourceFile = new FileFontSource(fontFile);

		// Defining MemoryFontSource
		byte[] bytes = Files.readAllBytes(new File(fontFile).toPath());
		MemoryFontSource sourceMemory = new MemoryFontSource(bytes);

		// Setting font sources
		FontConfigs.setFontSources(new FontSourceBase[] { sourceFolder, sourceFile, sourceMemory });

	}
}
