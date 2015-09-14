Imports System.Xml
Imports Netron.Lithium
Module modLithium
  ' ==================================== LOCAL VARIABLES ===================================================
  Private ndxSelTree As XmlNode = Nothing     ' The selected tree (if any is selected)
  Private loc_intSelId As Integer = -1        ' ID of the selected node
  Private colCollapsedEtree As New StringColl ' Collection of EtreeIDs that are "collapsed"
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  MakeLitTree
  ' Goal :  Take the parsing tree [objParse] and display it on lithium control [objLitTree]
  ' History:
  ' 28-10-2011  ERK Created for NLPstudio
  ' 03-11-2011  ERK Adapted for Cesax psdx trees
  ' ---------------------------------------------------------------------------------------------------------
  Public Function MakeLitTree(ByRef objLitTree As LithiumControl, ByRef ndxTree As XmlNode, _
                              ByRef ndxSel As XmlNode) As Boolean
    Dim objRoot As ShapeBase      ' Root lithium object
    Dim objItem As ShapeBase      ' One item
    Dim intScrollH As Integer     ' Horizontal scroll
    Dim intScrollV As Integer     ' Vertical scroll
    Dim intI As Integer           ' Counter

    Try
      ' Validate -- nothing
      ' Note where the scroll positions are
      intScrollH = objLitTree.HorizontalScroll.Value
      intScrollV = objLitTree.VerticalScroll.Value
      '' =========== debug ==============
      'Debug.Print("MakeLitTree ENTRY....")
      '' ================================
      '' Check which are expanded
      'For Each objItem In objLitTree.Shapes
      '  If (objItem.NodeName = "eTree") AndAlso (Trim(objItem.NodeId) <> "") Then
      '    ' =========== debug ==============
      '    Debug.Print("Node [" & objItem.NodeId & "] " & IIf(objItem.Expanded, "e", "c") & " [" & objItem.Text & "]")
      '    ' ================================
      '    If (objItem.Expanded) Then
      '      ' Clear it from the collection
      '      colCollapsedEtree.DelItem(objItem.NodeId)
      '    Else
      '      ' Note this one is collapsed
      '      colCollapsedEtree.AddUnique(objItem.NodeId)
      '    End If
      '  End If
      'Next objItem

      ' Start making a lithium control tree
      objLitTree.NewDiagram()
      ' Make up the root
      objRoot = objLitTree.Root
      With objRoot
        If (ndxTree.Name = "forest") Then
          ' Show the location
          .Text = ndxTree.Attributes("TextId").Value & " " & ndxTree.Attributes("Location").Value
          .NodeName = "forest" : .NodeId = ndxTree.Attributes("forestId").Value
        Else
          ' Show the label
          .Text = ndxTree.Attributes("Label").Value
          .NodeName = "eTree" : .NodeId = ndxTree.Attributes("Id").Value
        End If
        .Visible = True
      End With
      ' Set the selected tree
      ndxSelTree = ndxSel
      If (ndxSel Is Nothing) Then
        loc_intSelId = -1
      Else
        loc_intSelId = ndxSel.Attributes("Id").Value
      End If
      ' Recursively add the nodes
      If (Not WalkTree(objRoot, ndxTree.ChildNodes)) Then Return False
      ' Expand the results
      objRoot.Expand()
      ' =========== debug ==============
      ' Debug.Print("MakeLitTree EXIT....")
      ' ================================
      ' Collapse what needs be
      For intI = objLitTree.Shapes.Count - 1 To 0 Step -1
        objItem = objLitTree.Shapes(intI)
        ' =========== debug ==============
        ' Debug.Print("Node [" & objItem.NodeId & "] " & IIf(objItem.Expanded, "e", "c") & " [" & objItem.Text & "]")
        'If (objItem.NodeId = 794) Then Stop
        ' ================================
        If (objItem.NodeName = "eTree") AndAlso (colCollapsedEtree.Exists(objItem.NodeId)) Then
          ' This needs collapsing --> expanded = false
          objItem.Collapse(True)
        End If
      Next intI
      ' Draw the tree
      objLitTree.DrawTree()
      'For Each objItem In objLitTree.Shapes
      '  ' Verification...
      '  If (Trim(objItem.NodeId) = "") Then Stop
      '  ' ================
      '  If (objItem.NodeName = "eTree") AndAlso (colCollapsedEtree.Exists(objItem.NodeId)) Then
      '    ' This needs collapsing --> expanded = false
      '    objItem.Collapse(True)
      '  End If
      'Next objItem
      ' Return to the scroll positions
      objLitTree.HorizontalScroll.Value = intScrollH
      objLitTree.VerticalScroll.Value = intScrollV
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLithium/MakeLitTree error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  litFind
  ' Goal :  Find node [ndxThis] within tree [objLitTree] and return the corresponding shapebase in [objThis]
  '         The direction [strDir] may be:
  '         "left"
  '         "right"
  '         "up"
  ' History:
  ' 15-05-2013  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function litMove(ByRef objLitTree As LithiumControl, ByRef ndxThis As XmlNode, ByVal strDir As String) As Boolean
    Dim objThis As ShapeBase = Nothing          ' One item
    Dim objSibl As Lithium.ShapeBase = Nothing  ' My sibling to the left
    Dim objPar As ShapeBase = Nothing           ' The parent

    Try
      ' Validate
      If (objLitTree Is Nothing) OrElse (ndxThis Is Nothing) Then Return False
      ' If we move left/right, then just re-do the children of the parent of [ndxThis]
      If (strDir = "left") OrElse (strDir = "right") Then
        ' There must be a parent node
        If (ndxThis.ParentNode Is Nothing) Then Return False
        ' Get the parent of [ndxThis] within the lithium control
        If (Not litFind(objLitTree, ndxThis.ParentNode, objThis)) Then Return False

      ElseIf (strDir = "up") Then
      Else
        ' Unknown command
      End If

      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLithium/litFind error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  litFind
  ' Goal :  Find node [ndxThis] within tree [objLitTree] and return the corresponding shapebase in [objThis]
  ' History:
  ' 15-05-2013  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function litFind(ByRef objLitTree As LithiumControl, ByRef ndxThis As XmlNode, _
                          ByRef objThis As ShapeBase) As Boolean
    Dim objItem As ShapeBase  ' One item
    Dim intId As Integer      ' The id

    Try
      ' Validate
      If (objLitTree Is Nothing) OrElse (ndxThis Is Nothing) Then Return False
      ' Get the id
      Select Case ndxThis.Name
        Case "forest"
          intId = ndxThis.Attributes("forestId").Value
        Case "eTree"
          intId = ndxThis.Attributes("Id").Value
        Case "eLeaf"  ' Get the ID of the parent
          intId = ndxThis.ParentNode.Attributes("Id").Value
      End Select
      ' Visit all items
      For Each objItem In objLitTree.Shapes
        ' Check if this is the right one
        Select Case ndxThis.Name
          Case "forest"
            ' Not yet implemented
          Case "eTree"
            If (objItem.NodeName = "eTree") Then
              ' Check the id
              If (intId = objItem.NodeId) Then objThis = objItem : Return True
            End If
          Case "eLeaf"
            ' The parent <eTree> must be the correct one
            If (objItem.NodeName = "eLeaf") Then
              ' Check parent's id
              If (intId = objItem.ParentNode.NodeId) Then objThis = objItem : Return True
            End If
        End Select
      Next objItem
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modLithium/litFind error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  CollapseChange
  ' Goal :  Process any possible changes in the collapse status
  ' History:
  ' 14-05-2013  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function CollapseChange(ByRef objLitTree As LithiumControl) As Boolean
    Dim objItem As ShapeBase      ' One item

    Try
      ' Check which are expanded
      For Each objItem In objLitTree.Shapes
        If (objItem.NodeName = "eTree") AndAlso (Trim(objItem.NodeId) <> "") Then
          '' =========== debug ==============
          'Debug.Print("Node [" & objItem.NodeId & "] " & IIf(objItem.Expanded, "e", "c") & " [" & objItem.Text & "]")
          '' ================================
          If (objItem.Expanded) Then
            '' ===== DEBUG
            'If (objItem.NodeId = 2857) Then Stop
            '' ================
            ' Clear it from the collection
            colCollapsedEtree.DelItem(colCollapsedEtree.Find(objItem.NodeId))
          Else
            ' Note this one is collapsed
            colCollapsedEtree.AddUnique(objItem.NodeId)
          End If
        End If
      Next objItem
      ' Debug.Print(colCollapsedEtree.Count)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLithium/CollapseChange error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  CollapseClear
  ' Goal :  Clear the collapse collection
  ' History:
  ' 14-05-2013  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function CollapseClear() As Boolean
    Try
      colCollapsedEtree.Clear()
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLithium/CollapseClear error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  CollapseSet
  ' Goal :  Set or clear the collapse state of node [ndxThis] in the collapse collection
  ' History:
  ' 14-05-2013  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function CollapseSet(ByRef ndxThis As XmlNode, ByVal bSet As Boolean) As Boolean
    Dim intId As Integer  ' Etree Id
    Dim intI As Integer   ' Index

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      If (ndxThis.Name = "eTree") Then
        ' Get the id
        intId = ndxThis.Attributes("Id").Value
        ' Check what we need to do
        If (bSet) Then
          ' Set this one
          colCollapsedEtree.AddUnique(intId)
        Else
          ' Find the element in the collection
          intI = colCollapsedEtree.Find(intId)
          If (intI >= 0) Then
            ' Delete it
            colCollapsedEtree.DelItem(intI)
          End If
        End If
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLithium/CollapseSet error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  NodeColor
  ' Goal :  Walk the parsing tree
  ' History:
  ' 18-06-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function NodeColor(ByRef ndxThis As XmlNode, ByVal bSelected As Boolean) As System.Drawing.Color
    Dim strNodeName As String
    Dim strNodeLabel As String

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return Color.White
      ' Get the node name and label
      strNodeName = ndxThis.Name
      Select Case strNodeName
        Case "eLeaf"
          ' Get the label
          strNodeLabel = ndxThis.Attributes("Type").Value
          Select Case strNodeLabel
            Case "Vern", "Punct"
              If (bSelected) Then
                Return Color.LightBlue
              Else
                Return Color.Ivory
              End If
            Case "Zero", "Star"
              If (bSelected) Then
                Return Color.DarkGreen
              Else
                Return Color.LightGreen
              End If
            Case Else
              Return Color.White
          End Select
        Case "eTree"
          ' Get the label
          strNodeLabel = ndxThis.Attributes("Label").Value
          Select Case strNodeLabel
            Case "META", "CODE"
              Return Color.LightGray
            Case Else
              If (bSelected) Then
                Return Color.DarkGoldenrod
              Else
                Return Color.SteelBlue
              End If
          End Select
        Case "forest"
          If (bSelected) Then
            Return Color.DarkGoldenrod
          Else
            Return Color.SteelBlue
          End If
        Case Else
          Return Color.White
      End Select
    Catch ex As Exception
      ' Show error
      HandleErr("modLithium/NodeColor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure: default color
      Return Color.White
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  WalkTree
  ' Goal :  Walk the parsing tree
  ' History:
  ' 18-10-2011  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Function WalkTree(ByRef objCurrentShape As ShapeBase, ByRef ndxList As XmlNodeList) As Boolean
    Dim ndxThis As XmlNode          ' One of the elements from [ndxList]
    Dim objChildShape As ShapeBase  ' Shape of one child
    Dim strToken As String          ' Value of this node
    Dim strLabel As String          ' The label of this node
    Dim strChildType As String      ' Node type of child

    Try
      ' Validate
      If (ndxList Is Nothing) OrElse (objCurrentShape Is Nothing) Then Return False
      ' Get
      ' Walk all children
      For Each ndxThis In ndxList
        strChildType = ndxThis.Name
        ' Check what node we have
        Select Case strChildType
          Case "eLeaf"
            ' This is a terminal node -- but what kind?
            Select Case ndxThis.Attributes("Type").Value
              Case "Vern", "Punct"
                ' Get the text of this token
                strToken = ndxThis.Attributes("Text").Value
                ' Translate this text to readible text
                strToken = VernToEnglish(strToken)
                objChildShape = objCurrentShape.AddChild(strToken)
                With objChildShape
                  .ShapeColor = NodeColor(ndxThis, False)
                  .NodeName = "eLeaf" : .NodeId = ndxThis.ParentNode.Attributes("Id").Value
                End With
              Case "Zero", "Star"
                ' Get the text of this token
                strToken = ndxThis.Attributes("Text").Value
                objChildShape = objCurrentShape.AddChild(strToken)
                With objChildShape
                  .ShapeColor = NodeColor(ndxThis, False)
                  .NodeName = "eLeaf" : .NodeId = ndxThis.ParentNode.Attributes("Id").Value
                End With
              Case Else
                ' No need to make this one
            End Select
          Case "eTree"
            ' We do NOT want to look at certain nodes...
            strLabel = ndxThis.Attributes("Label").Value
            If (Not DoLike(strLabel, "CODE|META|METADATA")) Then
              ' This is a part of the clause
              objChildShape = objCurrentShape.AddChild(strLabel)
              ' Is this within the selection?
              If (ndxThis.SelectSingleNode("ancestor-or-self::eTree[@Id=" & loc_intSelId & "]") Is Nothing) Then
                ' This is not selected, but default text
                objChildShape.ShapeColor = NodeColor(ndxThis, False)
              Else
                ' This is selected
                objChildShape.ShapeColor = NodeColor(ndxThis, True)
                objChildShape.IsSelected = True
              End If
              With objChildShape
                .NodeName = "eTree" : .NodeId = ndxThis.Attributes("Id").Value
              End With
              ' Do my children
              If (Not WalkTree(objChildShape, ndxThis.ChildNodes)) Then Return False
              ' Expand them
              ' objChildShape.Expand()
            End If
          Case Else
            ' We can skip this node!
        End Select
      Next ndxThis
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/WalkTree error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

End Module
