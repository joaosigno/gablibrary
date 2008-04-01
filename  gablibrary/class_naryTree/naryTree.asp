<!--#include file="treeNode.asp"-->
<%
'**************************************************************************************************************
'* GAB_LIBRARY Copyright (C) 2003	
'* License refer to license.txt		
'**************************************************************************************************************

'**************************************************************************************************************

'' @CLASSTITLE:		Tree
'' @CREATOR:		David Rankin, Michael Rebec, Michal Gabrukiewicz
'' @CREATEDON:		2006-06-14 11:19
'' @CDESCRIPTION:	Creates a n-ary tree. So you can store 0-n childs for every parent node
'' @REQUIRES:		-
'' @VERSION:		0.1

'**************************************************************************************************************
class NaryTree

	'private members
	private p_root 
	private foundNode
	
	'public members
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		set p_root = nothing
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	adds a node to the tree.
	'' @DESCRIPTION:	parent must be in the tree already!
	'' @PARAM:			[string]: - value. Value of the node
	'' @PARAM:			[string]: - parent. Value of the parent-node
	'' @RETURN:			[bool] true if added. False if not (e.g. parent not found)
	'**********************************************************************************************************
	public function addNode(value, parentValue)
		addNode = false
		
		if typename(p_root) = "Nothing" and parentValue = empty then
			set p_root = getNewNode(value)
			addNode = true
		elseif typename(p_root) <> "Nothing" then
			if p_root.value = parentValue then 
				'if not p_root.childs.exists(value) then
				p_root.childs.add lib.getUniqueID() & "", getNewNode(value)
				addNode = true
				'end if
			else
				set node = findNode(parentValue, p_root)
				if typename(node) <> "Nothing" then
					node.childs.add lib.getUniqueID() & "", getNewNode(value)
					addNode = true
				end if
			end if
		end if
	end function
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	adds 0-n nodes to the tree from an (unsorted/sorted) array
	'' @PARAM:			[dictionary]: - nodesDict. Dictionary with your node. key = nodeValue, value = parentValue
	''					if a node has no parent then enter an "empty"
	'**********************************************************************************************************
	public sub addNodes(nodesDict)
		addNodesRecursiv nodesDict, 0
	end sub
	
	'**********************************************************************************************************
	'* addNodesRecursiv 
	'**********************************************************************************************************
	private sub addNodesRecursiv(byVal nodesDict, index)
		if nodesDict.count > 0 then
			if index > nodesDict.count - 1 then
				str.write("ICE-2537 (endless loop detected. seems to be a not valid tree. check if every child has parent)")
				str.end()
			end if
			
			keys = nodesDict.keys
			items = nodesDict.items
			added = addNode(keys(index), items(index))
			if added then
				str.write(keys(index) & "<br>")
				nodesDict.remove(keys(index))
				addNodesRecursiv nodesDict, 0
			else
				addNodesRecursiv nodesDict, index + 1
			end if
		end if
	end sub
	
	'**********************************************************************************************************
	'' @SDESCRIPTION:	finds a given node
	'' @PARAM:			[string]: - nodeValue. Value of the node you are looking for
	'' @RETURN:			[treeNode] found tree node. nothing if not found
	'**********************************************************************************************************
	public function find(nodeValue)
		set find = findNode(nodeValue, p_root)
	end function
	
	'**********************************************************************************************************
	'* getChildren 
	'**********************************************************************************************************
	public function getChildren(value)
		set parent = findNode(value, p_root)
		getChildrenList(parent)
	end function
	
	'**********************************************************************************************************
	'* getChildrenList 
	'**********************************************************************************************************
	private function getChildrenList(parent)
		if parent.childs.count > 0 then
			for each child in parent.childs.items
				if child.childs.count > 0 then
					getChildrenList(child)
					str.writeln(child.value & ", ")
				else
					str.writeln(child.value & ", ")
				end if
			next
		end if
	end function
	
	'**********************************************************************************************************
	'* findParent 
	'**********************************************************************************************************
	private function findNode(nodeValue, startNode)
		'response.write("<br>nodeValue - " & nodeValue & " - Start:" & startNode.value)
		if startNode.value = p_root.value then set foundNode = nothing
		
		if nodeValue = startNode.value then
			set foundNode = startNode
		elseif startNode.childs.count > 0 then
			for each child in startNode.childs.items
				if child.value = nodeValue then
					set foundNode = child
				else
					set foundNode = findNode(nodeValue, child)
				end if
			next
		end if
		set findNode = foundNode
	end function
	
	'**********************************************************************************************************
	'* getNewNode 
	'**********************************************************************************************************
	private function getNewNode(value) 
		set node = new treeNode
		node.value = value
		set getNewNode = node
	end function
	
end class
lib.registerClass("NaryTree")
%>