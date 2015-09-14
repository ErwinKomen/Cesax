Imports System.Xml
Module modQstack
  ' ================================= GLOBAL VARIABLES =========================================================
  Public tdlQstack As DataSet = Nothing     ' Query preparation stack
  Public strQstackFile As String = ""       ' Location of the Qstack file
  Public objQsEd As DgvHandle = Nothing     ' Handle
  ' ================================== LOCAL VARIABLES =========================================================
  Private colQstack As New NodeColl         ' Collection of nodes to build a query from
  Private bQstackInit As Boolean = False    ' Initialisation flag
  Private strQstackScheme As String = "Qstack.xsd"
  ' ============================================================================================================
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   eTreeQueryAdd
  ' Goal:   Add a node to the stack of nodes identified by the user as required to be detected by a query
  ' History:
  ' 31-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function eTreeQueryAdd(ByRef ndxThis As XmlNode, ByVal pntThis As Point) As Boolean
    Dim intIdx As Integer     ' Index of node
    Dim strVarName As String  ' Name of variable
    Dim strRelation As String ' Relation
    Dim ndxWork As XmlNode    ' One node
    Dim intI As Integer       ' Counter
    Dim intJ As Integer       ' Counter
    Dim colDesc As New NodeColl ' Collection of descendants

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      Select Case ndxThis.Name
        Case "eTree", "eLeaf"
          ' Okay, continue
        Case Else
          Status("Sorry, cannot build a query based on a node of type [" & ndxThis.Name & "]")
          Return False
      End Select
      ' Check if this node is in the stack already
      intIdx = colQstack.Find(ndxThis)
      If (intIdx >= 0) Then
        ' It is already there; is it the first (main) node?
        If (intIdx = 0) Then
          ' Ask if this is what the user wants
          Select Case MsgBox("Do you really want to reset the query-builder?", MsgBoxStyle.YesNoCancel)
            Case MsgBoxResult.No, MsgBoxResult.Cancel
              Return True
            Case Else
              ' Reset it 
              colQstack.Clear()
              Return True
          End Select
        End If
        ' Just delete it from the query stack
        colQstack.DelItem(colQstack.Find(ndxThis))
      Else
        ' Is this the first variable? Then don't ask
        If (colQstack.Count > 0) Then
          ' Think of a name for the variable
          strVarName = "cns" & colQstack.Count + 1
          ' Get a name for this Qstack variable
          pntThis.Y += 20
          Status("Press [Ok] when you have revised the short name (1-8 letters) for variable [" & strVarName & "]")
          If (Not frmMain.GetNewLabel(strVarName, pntThis)) Then Return False
          ' Show we have accepted the name
          Status("The name of this item is: [" & strVarName & "]")
        Else
          strVarName = "" ' Or "search"
        End If
        '' Validate ...
        'If (colQstack.Count > 0) Then
        '  ' Evaluate the latest addition: what is its relation to #0?
        '  strRelation = XmlRelation(colQstack.Item(0), ndxThis)
        '  Select Case strRelation
        '    Case "descendant"
        '      ' Add the whole chain of elements from lowest to highest
        '      ndxWork = ndxThis
        '      While (ndxWork.ParentNode IsNot colQstack.Item(0))
        '        ' Add parent node
        '        ndxWork = ndxWork.ParentNode
        '        colDesc.Add(ndxWork)
        '      End While
        '      ' Now add them from last to first
        '      For intI = colDesc.Count - 1 To 0 Step -1
        '        colQstack.Add(colDesc.Item(intI), "anc" & intI + 1 & "of" & strVarName)
        '      Next intI
        '    Case Else
        '      ' No smart ideas yet
        '  End Select
        'End If
        ' Add it to the stack
        colQstack.Add(ndxThis, strVarName)
      End If
      ' Show again what is in the stack
      If (Not ShowQstack()) Then Return False
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modQstack/eTreeQueryAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  Public Sub ClearQstack()
    Try
      colQstack.Clear()
    Catch ex As Exception
      ' Warn the user
      HandleErr("modQstack/ClearQstack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   InitQstack
  ' Goal:   Initialise the Qstack
  ' History:
  ' 01-11-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function InitQstack() As Boolean
    Try
      ' Already inited?
      If (Not bQstackInit) OrElse (tdlQstack Is Nothing) Then
        ' Create a Qstack datastructure
        strQstackFile = GetLocalDir("Cesax") & "\Qstack.xml"
        If (Not CreateDataSet(strQstackScheme, tdlQstack)) Then Return False
      End If
      ' Set initialisation flag
      bQstackInit = True
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modQstack/InitQstack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   QstackExport
  ' Goal:   Save the Qstack
  ' History:
  ' 01-11-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Sub QstackExport()
    Try
      ' Validate
      If (Not bQstackInit) OrElse (tdlQstack Is Nothing) Then Exit Sub
      ' Save the results
      tdlQstack.WriteXml(strQstackFile)
      Status("Qstack written to: " & strQstackFile)
    Catch ex As Exception
      ' Warn the user
      HandleErr("modQstack/QstackExport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   ShowQstack
  ' Goal:   Reveal the content of the query stack
  ' History:
  ' 31-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function ShowQstack() As Boolean
    Dim dtrFound() As DataRow   ' Result of SELECT
    Dim dtrThis As DataRow      ' Working entry
    Dim dtrMain As DataRow      ' Main node
    Dim strRelation As String   ' Relation the item has to the root
    Dim strRelMain As String    ' Main relation
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter
    Dim intK As Integer         ' Counter
    Dim strBest As String       ' Index of the best
    Dim arBest() As String = {"*sibling*", "*child*", "*parent*", "*ancestor*", "*descendant*", "*following*", "*preceding*"}

    Try
      ' Initialise
      If (Not InitQstack()) Then Return False
      ' make the form show, while we work
      If (Not frmQstack.Visible) Then
        frmQstack.Show(frmMain)
      End If
      ' Suspend layout right now
      frmQstack.dgvQstack.SuspendLayout()
      ' Clear the stack
      ClearTable(tdlQstack.Tables("entry"))
      ' Make sure this goes well
      tdlQstack.AcceptChanges()
      'Me.tbTreeHelp.Text = ""
      'Me.tbTreeHelp.Visible = True
      ' Do we have a main entry?
      If (colQstack.Count > 0) Then
        ' Add the main entry
        dtrThis = AddOneDataRow(tdlQstack, "entry", "entryId", "")
        With dtrThis
          ' Add characteristics
          .Item("Type") = "Search" : .Item("Value") = colQstack.Item(0).Attributes("Label").Value
          .Item("Number") = "0" : .Item("Base") = "-" : .Item("Relation") = "-"
          .Item("Name") = ""
          Select Case colQstack.Item(0).Name
            Case "eTree"
              dtrThis.Item("Node") = "phrase"
            Case "eLeaf"
              dtrThis.Item("Node") = "text"
            Case Else
              Status("Sorry, but we cannot handle nodes of type [" & colQstack.Item(0).Name & "]")
              Return False
          End Select
        End With
        ' Me.tbTreeHelp.AppendText("Look for: " & colQstack.Item(0).Attributes("Label").Value & vbCrLf & vbCrLf)
        If (colQstack.Count > 1) Then
          ' Add conditions
          ' Me.tbTreeHelp.AppendText("Conditions:" & vbCrLf)
          For intI = 1 To colQstack.Count - 1
            ' Figure out what the relation is for this one
            strRelation = XmlRelation(colQstack.Item(0), colQstack.Item(intI))
            strRelMain = strRelation
            ' Add this item as condition
            dtrThis = AddOneDataRow(tdlQstack, "entry", "entryId", "")
            dtrMain = dtrThis
            With dtrThis
              ' Add characteristics
              .Item("Type") = "Match" : .Item("Number") = intI : .Item("Base") = "search" : .Item("Relation") = strRelation
              .Item("Name") = colQstack.ItName(intI)
            End With
            Select Case colQstack.Item(intI).Name
              Case "eTree"
                'Me.tbTreeHelp.AppendText(" " & intI & "." & vbTab & "search/" & strRelation & vbTab & _
                '                         colQstack.Item(intI).Attributes("Label").Value & vbCrLf)
                dtrThis.Item("Value") = colQstack.Item(intI).Attributes("Label").Value
                dtrThis.Item("Node") = "phrase"
              Case "eLeaf"
                'Me.tbTreeHelp.AppendText(" " & intI & "." & vbTab & "search/" & strRelation & vbTab & _
                '                         colQstack.Item(intI).Attributes("Text").Value & vbCrLf)
                dtrThis.Item("Value") = colQstack.Item(intI).Attributes("Text").Value
                dtrThis.Item("Node") = "text"
            End Select
            ' Add relations with previous constituents, provided they are simple
            For intJ = 1 To intI - 1
              ' Figure out what the relation is for this one
              strRelation = XmlRelation(colQstack.Item(intJ), colQstack.Item(intI))
              If DoLike(strRelation, "*sibling*|*child*|*parent*") Then
                ' This is a simple relation, add it
                dtrThis = AddOneDataRow(tdlQstack, "entry", "entryId", "")
                With dtrThis
                  ' Add characteristics
                  .Item("Type") = "Order" : .Item("Number") = intI : .Item("Base") = intJ : .Item("Relation") = strRelation
                  .Item("Value") = ""
                End With
                Select Case colQstack.Item(intI).Name
                  Case "eTree"
                    dtrThis.Item("Node") = "phrase"
                  Case "eLeaf"
                    dtrThis.Item("Node") = "text"
                End Select
              End If
            Next intJ
            ' What kind of relation do we have for the "Match"?
            If (Not DoLike(strRelMain, "*sibling*|*child*|*parent*")) Then
              ' Check if we get a better match from one of the "Order"s
              dtrFound = tdlQstack.Tables("entry").Select("Number = " & intI & " AND Type = 'Order'")
              ' Walk them
              For intJ = 0 To dtrFound.Length - 1
                ' Is this a good one?
                strRelation = dtrFound(intJ).Item("Relation")
                If (DoLike(strRelation, "*sibling*|*child*|*parent*")) Then
                  ' Swap them
                  With dtrFound(intJ)
                    .Item("Type") = "Match" : .Item("Name") = dtrMain.Item("Name") : .Item("Value") = dtrMain.Item("Value")
                  End With
                  With dtrMain
                    .Item("Type") = "Order" : .Item("Name") = "" : .Item("Value") = ""
                  End With
                  ' And escape the loop
                  Exit For
                End If
              Next intJ
            End If

            'Next intI
            '' Add relations with previous constituents -- no matter which ones!
            'For intJ = 1 To intI - 1
            '  ' Figure out what the relation is for this one
            '  strRelation = XmlRelation(colQstack.Item(intJ), colQstack.Item(intI))
            '  ' Add this relation
            '  dtrThis = AddOneDataRow(tdlQstack, "entry", "entryId", "")
            '  With dtrThis
            '    ' Add characteristics
            '    .Item("Type") = "Order" : .Item("Number") = intI : .Item("Base") = intJ : .Item("Relation") = strRelation
            '    .Item("Value") = ""
            '  End With
            '  Select Case colQstack.Item(intI).Name
            '    Case "eTree"
            '      dtrThis.Item("Node") = "phrase"
            '    Case "eLeaf"
            '      dtrThis.Item("Node") = "text"
            '  End Select
            'Next intJ
            '' See if we can find a 'best' relation
            'For intK = 0 To arBest.Length - 1
            '  ' Check for those having this kind of relationship
            '  dtrFound = tdlQstack.Tables("entry").Select("Number = " & intI & " AND Relation LIKE '" & arBest(intK) & "'")
            '  ' Do we have something?
            '  If (dtrFound.Length > 0) Then
            '    ' note this item, and delete the others
            '    strBest = dtrFound(0).Item("Base").ToString
            '    dtrFound = tdlQstack.Tables("entry").Select("Number = " & intI & " AND Base <> '" & strBest & "'")
            '    For intJ = dtrFound.Length - 1 To 0 Step -1
            '      ' Remove this item
            '      dtrFound(intJ).Delete()
            '    Next intJ
            '    ' Make sure changes are accepted
            '    tdlQstack.AcceptChanges()
            '    ' Get out of here
            '    Exit For
            '  End If
            'Next intK

          Next intI
        End If
      End If
      ' Resume layout right now
      frmQstack.dgvQstack.ResumeLayout()
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modQstack/ShowQstack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
End Module
