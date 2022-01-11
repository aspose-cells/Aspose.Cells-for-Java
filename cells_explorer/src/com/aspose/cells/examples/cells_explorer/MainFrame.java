package com.aspose.cells.examples.cells_explorer;

import java.awt.Font;
import java.io.File;
import java.util.Enumeration;

import javax.swing.GroupLayout;
import javax.swing.UIManager;
import javax.swing.plaf.FontUIResource;

public class MainFrame extends javax.swing.JFrame
{

	private static final long	serialVersionUID	= 1L;
	
	public MainFrame()
	{
		try
		{
			for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels())
			{
				if ("Metal".equals(info.getName()))
				{
					javax.swing.UIManager.setLookAndFeel(info.getClassName());
					break;
				}
			}
		}
		catch (ClassNotFoundException ex)
		{
			java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(
					java.util.logging.Level.SEVERE, null, ex);
		}
		catch (InstantiationException ex)
		{
			java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(
					java.util.logging.Level.SEVERE, null, ex);
		}
		catch (IllegalAccessException ex)
		{
			java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(
					java.util.logging.Level.SEVERE, null, ex);
		}
		catch (javax.swing.UnsupportedLookAndFeelException ex)
		{
			java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(
					java.util.logging.Level.SEVERE, null, ex);
		}
		
		InitGlobalFont(new Font("alias", Font.PLAIN, 16)); 
		
		initComponents();
	}
	
	private void initComponents()
	{        
		//Globals.mTree
		treeScrollPane = new javax.swing.JScrollPane();	
		
		jScrollPane1 = new javax.swing.JScrollPane();
		textArea = new javax.swing.JTextPane();
		
		jMenuBar = new javax.swing.JMenuBar();
		jMenuFile = new javax.swing.JMenu();
		menuOpen = new javax.swing.JMenuItem();
		jSeparatorFile = new javax.swing.JPopupMenu.Separator();
		menuExit = new javax.swing.JMenuItem();
		jMenuView = new javax.swing.JMenu();
		menuExpandAll = new javax.swing.JMenuItem();
		menuCollapseAll = new javax.swing.JMenuItem();
		
		
		setDefaultCloseOperation(javax.swing.WindowConstants.DO_NOTHING_ON_CLOSE);
		setTitle("Cells Explorer");
		setIconImages(null);
		setModalExclusionType(java.awt.Dialog.ModalExclusionType.APPLICATION_EXCLUDE);
		setName("CellsExplorer");
		setPreferredSize(new java.awt.Dimension(1024, 768));
		
		treeScrollPane.setBackground(new java.awt.Color(255, 255, 255));
		treeScrollPane.setAutoscrolls(true);
		
		textArea.setEditable(false);
		jScrollPane1.setViewportView(textArea);	
		
		
		jMenuFile.setMnemonic('F');
		jMenuFile.setText("File");
		menuOpen.setMnemonic('O');
		menuOpen.setText("Open");
		jMenuFile.add(menuOpen);
		jMenuFile.add(jSeparatorFile);
		menuExit.setMnemonic('X');
		menuExit.setText("Exit");
		jMenuFile.add(menuExit);
		jMenuBar.add(jMenuFile);
		
		jMenuView.setMnemonic('V');
		jMenuView.setText("View");
		menuExpandAll.setMnemonic('E');
		menuExpandAll.setText("Expand All");
		menuExpandAll.setEnabled(false);
		jMenuView.add(menuExpandAll);
		menuCollapseAll.setMnemonic('C');
		menuCollapseAll.setText("Collapse All");
		menuCollapseAll.setEnabled(false);
		jMenuView.add(menuCollapseAll);
		jMenuBar.add(jMenuView);		
		setJMenuBar(jMenuBar);
		
		javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
		getContentPane().setLayout(layout);
		
		layout.setHorizontalGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)			    
				.addGroup(layout.createSequentialGroup()
				.addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
				.addGroup(layout.createSequentialGroup().addGap(15, 15,15)
				.addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
				.addGroup(layout.createSequentialGroup()
				.addComponent(treeScrollPane, GroupLayout.PREFERRED_SIZE, 269, GroupLayout.PREFERRED_SIZE)
				.addGap(12, 12, 12)
				.addComponent(jScrollPane1))))
				.addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()																
				.addGap(0, 0, 0)))
				.addContainerGap()));
		
		layout.setVerticalGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
				.addGroup(layout.createSequentialGroup()
				.addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)								
				.addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
				.addComponent(treeScrollPane)
				.addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 135, Short.MAX_VALUE))
				.addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
				.addGap(0, 0, 0)));
		
		pack();
	}
	
	private static void InitGlobalFont(java.awt.Font font)
    {
		  FontUIResource fontRes = new FontUIResource(font);
		  for (Enumeration<Object> keys = UIManager.getDefaults().keys(); keys.hasMoreElements(); )
		  {
			  Object key = keys.nextElement();
			  Object value = UIManager.get(key);
			  if (value instanceof FontUIResource)
			  {
				  UIManager.put(key, fontRes);
			  }
		  }
	}
	
	public String getResourcesDir()
	{
        File dir = new File(System.getProperty("user.dir"));
        dir = new File(dir, "src");
        dir = new File(dir, "resources");

        return dir.toString() + File.separator;
    }
	
	private javax.swing.JMenuBar				jMenuBar;
	private javax.swing.JMenu					jMenuFile;
	protected javax.swing.JMenuItem				menuExit;
	private javax.swing.JPopupMenu.Separator	jSeparatorFile;
	protected javax.swing.JMenuItem				menuOpen;
	
	private javax.swing.JMenu					jMenuView;
	protected javax.swing.JMenuItem				menuExpandAll;
	protected javax.swing.JMenuItem				menuCollapseAll;
		
	private javax.swing.JScrollPane				jScrollPane1;	
	protected javax.swing.JTextPane				textArea;	
	protected javax.swing.JScrollPane			treeScrollPane;
}