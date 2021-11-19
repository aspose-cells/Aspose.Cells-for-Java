package com.aspose.cells.examples.cells_explorer;

import javax.swing.tree.DefaultMutableTreeNode;

import com.aspose.cells.examples.cells_explorer.model.CellsNode;

/**
 * Base class used to provide GUI representation for cells nodes.
 */
public class Item
{
    private CellsNode mNode;
    private DefaultMutableTreeNode mTreeNode;
   
    public Item(CellsNode node)
    {
        mNode = node;
    }

    
    public CellsNode getNode()
    {
        return mNode;
    }

    /**
     * The display name for this Item. Can be customized by overriding this method in inheriting classes.
     */
    public String getName() throws Exception
    {
        return mNode.getNodeName();
    }

    /**
     * The text of the corresponding cells node.
     */
    public String getText() throws Exception
    {
        return mNode.getNodeContent();
    }

    /**
     * Creates a TreeNode for this item to be displayed in the Cells Explorer TreeView control.
     */
    public DefaultMutableTreeNode getTreeNode() throws Exception
    {
        if (mTreeNode == null)
        {
            mTreeNode = new DefaultMutableTreeNode(this);            

            if (mNode.getChildList().size() > 0)
            {
                mTreeNode.add(new DefaultMutableTreeNode("#dummy"));
            }
        }
        return mTreeNode;
    }



    /**
     * Provides lazy on-expand loading of underlying tree nodes.
     */
    public void onExpand() throws Exception
    {
        if ("#dummy".equals(getTreeNode().getFirstChild().toString()))
        {
            getTreeNode().removeAllChildren();
            Globals.mTreeModel.reload(getTreeNode());
            for (Object o : mNode.getChildList())
            {
                CellsNode n = (CellsNode) o;
                getTreeNode().add(Item.createItem(n).getTreeNode());
            }
        }
    }

    /**
     * Item class factory implementation.
     */
    public static Item createItem(CellsNode node)
    {        
        return new Item(node);
    }

    /**
     * Object.toString method used by Tree.
     */
    public String toString()
    {
        try 
        {
            return getName();
        } catch (Exception e)
        {
            throw new RuntimeException(e);
        }
        
    }

}