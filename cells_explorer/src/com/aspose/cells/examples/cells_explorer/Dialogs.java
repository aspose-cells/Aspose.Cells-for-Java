package com.aspose.cells.examples.cells_explorer;

import java.io.File;
import java.text.MessageFormat;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

public class Dialogs
{

	private static String		mDocumentPath	= "";
	private static JFileChooser	mOpenDialog;
	

	private Dialogs()
	{
	}
	
	static
	{
		mOpenDialog = new JFileChooser();
		mOpenDialog.setAcceptAllFileFilterUsed(false);
		mOpenDialog.setFileFilter(Globals.OPEN_FILE_FILTER);
		mOpenDialog.setMultiSelectionEnabled(false);
		mOpenDialog.setFileSelectionMode(JFileChooser.FILES_ONLY);		
		mOpenDialog.setDialogTitle(Globals.OPEN_DOCUMENT_DIALOG_TITLE);
	}	

	public static String openDocument()
	{
		mOpenDialog.setCurrentDirectory(new File(mDocumentPath));
		
		if (mOpenDialog.showOpenDialog(Globals.mMainForm) == JFileChooser.APPROVE_OPTION)
		{
			File file = mOpenDialog.getSelectedFile();
			String fileName = file.getAbsolutePath();
			if (file.exists())
			{
				mDocumentPath = file.getParent();
				return fileName;
			}
			else
			{
				JOptionPane.showMessageDialog(Globals.mMainForm, MessageFormat
						.format("File \"{0}\" doesn't exist.", fileName),
						Globals.APPLICATION_TITLE, JOptionPane.ERROR_MESSAGE);
				return "";
			}
		}
		else
		{
			return "";
		}
	}
	
}
