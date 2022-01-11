package com.aspose.cells.examples.cells_explorer;

import java.io.File;
import java.text.MessageFormat;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

public class OpenFileDialog
{

	private static String		mDocumentPath	= "";
	private static JFileChooser	mOpenDialog;
	

	private OpenFileDialog()
	{
	}
	
	static
	{
		mOpenDialog = new JFileChooser();
		mOpenDialog.setAcceptAllFileFilterUsed(false);
		mOpenDialog.setFileFilter(GlobalConstant.OPEN_FILE_FILTER);
		mOpenDialog.setMultiSelectionEnabled(false);
		mOpenDialog.setFileSelectionMode(JFileChooser.FILES_ONLY);		
		mOpenDialog.setDialogTitle(GlobalConstant.OPEN_DOCUMENT_DIALOG_TITLE);
	}	

	public static String openDocument()
	{
		mOpenDialog.setCurrentDirectory(new File(mDocumentPath));
		
		if (mOpenDialog.showOpenDialog(GlobalConstant.mMainFrame) == JFileChooser.APPROVE_OPTION)
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
				JOptionPane.showMessageDialog(GlobalConstant.mMainFrame, MessageFormat
						.format("File \"{0}\" doesn't exist.", fileName),
						GlobalConstant.APPLICATION_TITLE, JOptionPane.ERROR_MESSAGE);
				return "";
			}
		}
		else
		{
			return "";
		}
	}
	
}
