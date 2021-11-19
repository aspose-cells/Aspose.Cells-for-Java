package com.aspose.cells.examples.cells_explorer.model;

import java.util.ArrayList;
import java.util.UUID;


public class CellsNode
{
    public CellsNode()
    {
        this.setNodeId(UUID.randomUUID().toString());
        this.setDeleted(false);
        this.setChildList(new ArrayList<CellsNode>());
    }
   

    
    private String nodeId;
    public String getNodeId()
    {
    	return nodeId;
    }
    public void setNodeId(String value)
    {
    	this.nodeId = value;
    }

    private String nodeName;
    public String getNodeName()
    {
    	return nodeName;
    }
    public void setNodeName(String value)
    {
    	this.nodeName = value;
    }

  
    private String nodeContent;
    public String getNodeContent()
    {
    	return nodeContent;
    }
    public void setNodeContent(String value)
    {
    	this.nodeContent = value;
    }
    

    private boolean isDeleted;
    public boolean isDeleted()
    {
    	return isDeleted;
    }
    public void setDeleted(boolean value)
    {
    	this.isDeleted = value;
    }

 
    private int nodeType;
    public int getNodeType()
    {
    	return nodeType;
    }
    public void setNodeType(int value)
    {
    	this.nodeType = value;
    }
    

    private ArrayList<CellsNode> childList;
    public ArrayList<CellsNode> getChildList()
    {	
    	return childList;
    }
    public void setChildList(ArrayList<CellsNode> value)
    {
    	this.childList = value;
    }
    
    public void addChild(CellsNode node)
    {
        if (getChildList() != null)
        {
            getChildList().add(node);
        }
        
    }

}

