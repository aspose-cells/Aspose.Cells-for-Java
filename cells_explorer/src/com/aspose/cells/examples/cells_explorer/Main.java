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

public class Main implements TreeWillExpandListener, TreeSelectionListener, KeyListener
{
	public Main() throws Exception
	{		
		Globals.mMainForm = new MainForm();
		
		// Get the screen size
		Toolkit toolkit = Toolkit.getDefaultToolkit();
		Dimension screenSize = toolkit.getScreenSize();
		
		// Calculate the frame location
		int x = (screenSize.width - Globals.mMainForm.getWidth()) / 2;
		int y = (screenSize.height - Globals.mMainForm.getHeight()) / 2;
		
		// Set the new frame location
		Globals.mMainForm.setLocation(x, y);
		
		Globals.mMainForm.setTitle(Globals.APPLICATION_TITLE);
		
		Globals.mMainForm.addWindowListener(new WindowAdapter()
		{
			public void windowClosing(WindowEvent e)
			{
				onClose();
			}
		});		
		
		
		Globals.mMainForm.menuOpen.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent evt)
			{
				onOpen();
			}
		});
		
		
		Globals.mMainForm.menuExit.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent evt)
			{
				onClose();
			}
		});		
		
		
		Globals.mMainForm.menuExpandAll.addActionListener(new ActionListener()
		{
			public void actionPerformed(ActionEvent evt)
			{
				onExpandAll();
			}
		});
		
		Globals.mMainForm.menuCollapseAll
				.addActionListener(new ActionListener()
				{
					public void actionPerformed(ActionEvent evt)
					{
						onCollapseAll();
					}
				});		
		
		Globals.mMainForm.setVisible(true);
		
	}
	
	private void onClose()
	{
		Globals.mMainForm.dispose();
	}
		
	private void onOpen()
	{
		
		try
		{
			String fileName = Dialogs.openDocument();
			if (!"".equals(fileName))
			{
				Globals.mMainForm.setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));				
				
				Globals.mMainForm.setTitle(Globals.APPLICATION_TITLE + " - " + fileName);
				
				CellsNode rootNode = WorkbookUtil.loadExcelFile(fileName);
				
				Globals.mRootNode = Item.createItem(rootNode).getTreeNode();
                Globals.mTreeModel = new DefaultTreeModel(Globals.mRootNode);
                
				Globals.mTree = new JTree(Globals.mTreeModel);
				Globals.mTree.setExpandsSelectedPaths(false);
				Globals.mTree.getSelectionModel().setSelectionMode(
						TreeSelectionModel.SINGLE_TREE_SELECTION);
				
				Globals.mTree.setShowsRootHandles(true);
				Globals.mTree.addTreeWillExpandListener(this);
				Globals.mTree.addTreeSelectionListener(this);
				Globals.mTree.addKeyListener(this);
				Globals.mMainForm.treeScrollPane.setViewportView(Globals.mTree);
				
				TreePath path = new TreePath(Globals.mRootNode);
				
				((Item) Globals.mRootNode.getUserObject()).onExpand();
				
				Globals.mTree.expandPath(path);
				Globals.mTree.setSelectionPath(path);
				

				Globals.mMainForm.menuExpandAll.setEnabled(true);
				Globals.mMainForm.menuCollapseAll.setEnabled(true);
			}
		}
		catch (Exception e)
		{
			Globals.mMainForm.setCursor(null);
		}
		finally
		{
			// Set the cursor back to normal even if an exception occurs.
			Globals.mMainForm.setCursor(null);
		}
	}	
	
	
	/**
	 * Expand all child nodes under the selected node.
	 */
	private void onExpandAll()
	{
		TreePath path = Globals.mTree.getSelectionPath();
		if (path != null)
		{
			Globals.mMainForm.setCursor(Cursor
					.getPredefinedCursor(Cursor.WAIT_CURSOR));
			expandAll(Globals.mTree, path, true);
			Globals.mMainForm.setCursor(null);
		}
	}
	
	/**
	 * Collapse all child nodes under the selected node
	 */
	private void onCollapseAll()
	{
		TreePath path = Globals.mTree.getSelectionPath();
		if (path != null)
		{
			Globals.mMainForm.setCursor(Cursor
					.getPredefinedCursor(Cursor.WAIT_CURSOR));
			expandAll(Globals.mTree, path, false);
			Globals.mMainForm.setCursor(null);
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
	 * Informs Item class, which provides GUI representation of a cells node,
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
				((Item) node.getUserObject()).onExpand();
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
	 * Informs Item class, which provides GUI representation of a cells node,
	 * that the corresponding TreeNode was selected.
	 */
	public void valueChanged(TreeSelectionEvent e)
	{
		DefaultMutableTreeNode node = (DefaultMutableTreeNode) Globals.mTree
				.getLastSelectedPathComponent();
		
		if (node == null)
		{
			return;
		}
		try
		{
			// This operation can take some time so we set the Cursor to WaitCursor.
			Globals.mMainForm.setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
			// Show the text contained by selected cells node.
			Item selectedItem = (Item) node.getUserObject();
			Globals.mMainForm.textArea.setText(selectedItem.getText());
			Globals.mMainForm.textArea.moveCaretPosition(0);			
			
			
			// Restore cursor.
			Globals.mMainForm.setCursor(null);
		}
		catch (Exception ex)
		{
			Globals.mMainForm.textArea.setText("");
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
