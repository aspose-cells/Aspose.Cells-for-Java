package com.aspose.cells.examples.cells_explorer;

import java.awt.Cursor;
import java.awt.Dimension;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.util.Enumeration;

import javax.swing.JTree;
import javax.swing.event.TreeExpansionEvent;
import javax.swing.event.TreeSelectionEvent;
import javax.swing.event.TreeSelectionListener;
import javax.swing.event.TreeWillExpandListener;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.DefaultTreeModel;
import javax.swing.tree.ExpandVetoException;
import javax.swing.tree.TreeNode;
import javax.swing.tree.TreePath;
import javax.swing.tree.TreeSelectionModel;

import com.aspose.cells.examples.cells_explorer.model.CellsNode;
import com.aspose.cells.examples.cells_explorer.model.WorkbookUtil;

public class CellsMain implements TreeWillExpandListener, TreeSelectionListener, KeyListener
{
	public CellsMain() throws Exception
	{		
		GlobalConstant.mMainFrame = new MainFrame();
		
		// Get the screen size
		Toolkit toolkit = Toolkit.getDefaultToolkit();
		Dimension screenSize = toolkit.getScreenSize();
		
		// Calculate the frame location
		int x = (screenSize.width - GlobalConstant.mMainFrame.getWidth()) / 2;
		int y = (screenSize.height - GlobalConstant.mMainFrame.getHeight()) / 2;
		
		// Set the new frame location
		GlobalConstant.mMainFrame.setLocation(x, y);
		
		GlobalConstant.mMainFrame.setTitle(GlobalConstant.APPLICATION_TITLE);
		
		GlobalConstant.mMainFrame.addWindowListener(new WindowAdapter()
		{
			public void windowClosing(WindowEvent e)
			{
				onClose();
			}
		});		
		
		
		GlobalConstant.mMainFrame.menuOpen.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent evt)
			{
				onOpen();
			}
		});
		
		
		GlobalConstant.mMainFrame.menuExit.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent evt)
			{
				onClose();
			}
		});		
		
		
		GlobalConstant.mMainFrame.menuExpandAll.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent evt)
			{
				onExpandAll();
			}
		});
		
		GlobalConstant.mMainFrame.menuCollapseAll
				.addActionListener(new ActionListener()
				{
					public void actionPerformed(ActionEvent evt)
					{
						onCollapseAll();
					}
				});		
		
		GlobalConstant.mMainFrame.setVisible(true);
		
	}
	
	private void onClose()
	{
		GlobalConstant.mMainFrame.dispose();
	}
		
	private void onOpen()
	{
		
		try
		{
			String fileName = OpenFileDialog.openDocument();
			if (!"".equals(fileName))
			{
				GlobalConstant.mMainFrame.setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));				
				
				GlobalConstant.mMainFrame.setTitle(GlobalConstant.APPLICATION_TITLE + " - " + fileName);
				
				CellsNode rootNode = WorkbookUtil.loadExcelFile(fileName);
				
				GlobalConstant.mRootNode = CellsItem.createItem(rootNode).getTreeNode();
                GlobalConstant.mTreeModel = new DefaultTreeModel(GlobalConstant.mRootNode);
                
				GlobalConstant.mTree = new JTree(GlobalConstant.mTreeModel);
				GlobalConstant.mTree.setExpandsSelectedPaths(false);
				GlobalConstant.mTree.getSelectionModel().setSelectionMode(
						TreeSelectionModel.SINGLE_TREE_SELECTION);
				
				GlobalConstant.mTree.setShowsRootHandles(true);
				GlobalConstant.mTree.addTreeWillExpandListener(this);
				GlobalConstant.mTree.addTreeSelectionListener(this);
				GlobalConstant.mTree.addKeyListener(this);
				GlobalConstant.mMainFrame.treeScrollPane.setViewportView(GlobalConstant.mTree);
				
				TreePath path = new TreePath(GlobalConstant.mRootNode);
				
				((CellsItem) GlobalConstant.mRootNode.getUserObject()).onExpand();
				
				GlobalConstant.mTree.expandPath(path);
				GlobalConstant.mTree.setSelectionPath(path);
				

				GlobalConstant.mMainFrame.menuExpandAll.setEnabled(true);
				GlobalConstant.mMainFrame.menuCollapseAll.setEnabled(true);
			}
		}
		catch (Exception e)
		{
			GlobalConstant.mMainFrame.setCursor(null);
		}
		finally
		{
			// Set the cursor back to normal even if an exception occurs.
			GlobalConstant.mMainFrame.setCursor(null);
		}
	}	
	
	
	/**
	 * Expand all child nodes under the selected node.
	 */
	private void onExpandAll()
	{
		TreePath path = GlobalConstant.mTree.getSelectionPath();
		if (path != null)
		{
			GlobalConstant.mMainFrame.setCursor(Cursor
					.getPredefinedCursor(Cursor.WAIT_CURSOR));
			expandAll(GlobalConstant.mTree, path, true);
			GlobalConstant.mMainFrame.setCursor(null);
		}
	}
	
	/**
	 * Collapse all child nodes under the selected node
	 */
	private void onCollapseAll()
	{
		TreePath path = GlobalConstant.mTree.getSelectionPath();
		if (path != null)
		{
			GlobalConstant.mMainFrame.setCursor(Cursor
					.getPredefinedCursor(Cursor.WAIT_CURSOR));
			expandAll(GlobalConstant.mTree, path, false);
			GlobalConstant.mMainFrame.setCursor(null);
		}
	}
	
	@SuppressWarnings("rawtypes")
	private void expandAll(JTree tree, TreePath parent, boolean expand)
	{
		// Traverse children.
		TreeNode node = (TreeNode) parent.getLastPathComponent();
		
		// Expansion or collapse must be done from the bottom-up
		if (expand)
		{
			tree.expandPath(parent);
		}
		
		if (node.getChildCount() >= 0)
		{
			for (Enumeration e = node.children(); e.hasMoreElements();)
			{
				TreeNode n = (TreeNode) e.nextElement();
				TreePath path = parent.pathByAddingChild(n);
				expandAll(tree, path, expand);
			}
		}
		
		if (!expand)
		{
			tree.collapsePath(parent);
		}
	}
	
	/**
	 * Informs CellsItem class, which provides GUI representation of a cells node,
	 * that the corresponding TreeNode is about being expanded.
	 */
	public void treeWillExpand(TreeExpansionEvent event)
			throws ExpandVetoException
	{
		DefaultMutableTreeNode node = (DefaultMutableTreeNode) event.getPath().getLastPathComponent();
		
		if (node != null)
		{
			try
			{
				((CellsItem) node.getUserObject()).onExpand();
			}
			catch (Exception e)
			{
				throw new RuntimeException(e);
			}
		}
	}
	
	public void treeWillCollapse(TreeExpansionEvent event)
			throws ExpandVetoException
	{
	}
	
	/**
	 * Informs CellsItem class, which provides GUI representation of a cells node,
	 * that the corresponding TreeNode was selected.
	 */
	public void valueChanged(TreeSelectionEvent e)
	{
		DefaultMutableTreeNode node = (DefaultMutableTreeNode) GlobalConstant.mTree
				.getLastSelectedPathComponent();
		
		if (node == null)
		{
			return;
		}
		try
		{
			// This operation can take some time so we set the Cursor to WaitCursor.
			GlobalConstant.mMainFrame.setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
			// Show the text contained by selected cells node.
			CellsItem selectedItem = (CellsItem) node.getUserObject();
			GlobalConstant.mMainFrame.textArea.setText(selectedItem.getText());
			GlobalConstant.mMainFrame.textArea.moveCaretPosition(0);			
			
			
			// Restore cursor.
			GlobalConstant.mMainFrame.setCursor(null);
		}
		catch (Exception ex)
		{
			GlobalConstant.mMainFrame.textArea.setText("");
		}
	}


	@Override
	public void keyPressed(KeyEvent arg0)
	{
	}

	@Override
	public void keyReleased(KeyEvent arg0)
	{
	}

	@Override
	public void keyTyped(KeyEvent arg0)
	{
	}
}
