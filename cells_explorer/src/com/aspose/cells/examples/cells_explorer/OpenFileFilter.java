package com.aspose.cells.examples.cells_explorer;

import java.io.File;

import javax.swing.filechooser.FileFilter;

public class OpenFileFilter extends FileFilter
{
	
	public OpenFileFilter(String[] extensions, String description)
	{
		assert extensions != null && extensions.length > 0 : "Null or Empty OpenFileFilter extensions array.";
		for (String ext : extensions)
		{
			assert ext != null && !"".equals(ext) : "Null or Empty OpenFileFilter extension.";
		}
		
		mExtensions = extensions;
		mDescription = description;
	}
	
	public boolean accept(File f)
	{
		if (f.isDirectory())
		{
			return true;
		}
		
		for (String ext : mExtensions)
		{
			if (f.getName().endsWith(ext))
			{
				return true;
			}
		}
		
		return false;
	}
	
	public String getDescription()
	{
		return mDescription;
	}
	
	String[]	mExtensions;
	String		mDescription;
}
