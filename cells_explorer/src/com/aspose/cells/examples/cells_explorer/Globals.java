package com.aspose.cells.examples.cells_explorer;

import javax.swing.JTree;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.DefaultTreeModel;

public class Globals
{
    /**
     * This class is purely static, that's why we prevent instance creation by declaring the constructor as private.
     */
    private Globals()
    {
    }

    // Titles used within the application.
    static final String APPLICATION_TITLE = "Cells Explorer";
    static final String UNEXPECTED_EXCEPTION_DIALOG_TITLE = APPLICATION_TITLE + " - unexpected error occured";    
    static final String OPEN_DOCUMENT_DIALOG_TITLE = "Open File";
    
    static final OpenFileFilter OPEN_FILE_FILTER = new OpenFileFilter(
            new String[]{".xlsx", ".xlsm", ".xlsb", "xls"}, "Excel 2003-2016 files");
    /**
     * Reference for application's main form.
     */
    static MainForm mMainForm;
    
    /**
     * Reference for current Tree Model
     */
    static DefaultTreeModel mTreeModel;

    /**
     * Reference for the current Tree
     */
    static JTree mTree;

    /**
     * Reference for the current root node.
     */
    static DefaultMutableTreeNode mRootNode;
}
