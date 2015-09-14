Option Explicit On
Option Compare Binary
Imports System.Xml
Public Class frmFindX
  ' =============================== LOCAL Events ========================================
  Private WithEvents rtbThis As RichTextBox
  ' =============================== LOCAL Variables =====================================
  Private objFindOptions As Integer = RichTextBoxFinds.None
  Private intSelStart As Integer = 1  ' Last activated selection
  Private intSelLen As Integer = 0    ' Length of last selection
  Private bSelReset As Boolean = True ' Whether selection was reset successfully
  Private bInit As Boolean = False    ' Initialisation
  Private bDirty As Boolean = False   ' Dirty flag
  ' =============================== LOCAL CONSTANTS =====================================
  Private Const DEFAULT_XPATH As String = "./following::eTree[tb:Like(@Label, 'NP*')]"
  ' ------------------------------------------------------------------------------------
  ' Name:   chbWhole_CheckedChanged
  ' Goal:   Only find whole words
  ' History:
  ' 06-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub chbWhole_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbWhole.CheckedChanged
    ' Adapt the FindOptions value
    If (Me.chbWhole.Checked) Then
      ' Set this option
      objFindOptions += RichTextBoxFinds.WholeWord
    Else
      ' Delete this option
      objFindOptions -= RichTextBoxFinds.WholeWord
    End If
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   cmdNext_Click
  ' Goal:   Search for the next word
  ' History:
  ' 06-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub cmdNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNext.Click
    DoFind(True)
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   cmdPrev_Click
  ' Goal:   Search for the previous occurrance of the word
  ' History:
  ' 06-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdPrev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrev.Click
    DoFind(False)
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   cmdSearch_Click
  ' Goal:   Search for all occurrances of the current query in the current document
  ' History:
  ' 17-04-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSearch.Click
    Dim strSearch As String     ' The string to search
    Dim ndxList As XmlNodeList  ' Result of the query
    Dim intI As Integer         ' Counter

    Try
      ' Get the search string
      strSearch = Trim(Me.tbFind.Text)
      ' Validate
      If (strSearch = "") Then Status("First define a query") : Exit Sub
      ' Additional error shield for xpath search
      Try
        ' Try do the Xpath searching
        ndxList = pdxCurrentFile.SelectNodes(strSearch)
      Catch ex As Exception
        ' There probably is a syntax mismatch
        MsgBox("There was a syntax error in the query:" & vbCrLf & ex.Message)
        Exit Sub
      End Try
      ' Initialise the searching

      ' Walk through all the results
      For intI = 0 To ndxList.Count - 1

      Next intI

    Catch ex As Exception
      ' Show error
      HandleErr("frmFind/cmdSearch error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   DoFind
  ' Goal:   Generalized find function
  ' History:
  ' 07-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub DoFind(ByVal bForward As Boolean)
    Dim strSearch As String             ' The string to search
    Dim ndxForest As XmlNode = Nothing  ' Node of the currently selected forest
    Dim ndxSel As XmlNode = Nothing     ' Currently selected <eTree> element
    Dim ndxThis As XmlNode              ' Working node - the result of our search

    Try
      ' Get the current forest
      If (Not GetCurrentForest(ndxForest, ndxSel)) Then Exit Sub
      ' Get the search string
      strSearch = Me.tbFind.Text
      ' Try do the Xpath searching
      Try
        If (ndxSel Is Nothing) Then
          ndxThis = ndxForest.SelectSingleNode(strSearch, conTb)
        Else
          ' We have a selection: go one step further or one step back
          If (bForward) Then
            ndxSel = ndxSel.SelectSingleNode("./following::eTree")
          Else
            ndxSel = ndxSel.SelectSingleNode("./preceding::eTree")
          End If
          ndxThis = ndxSel.SelectSingleNode(strSearch, conTb)
        End If
      Catch ex As Exception
        ' There probably is a syntax mismatch
        MsgBox("There was a syntax error:" & vbCrLf & ex.Message)
        Exit Sub
      End Try
      ' Found something?
      If (Not ndxThis Is Nothing) Then
        ' Switch on any constituent
        bAnyConstituent = True
        ' Go to the resulting <eTree>, and select it
        GotoNode(ndxThis)
      End If
      ' Set focus back to the find form
      Me.Focus()
    Catch ex As Exception
      ' Show error
      HandleErr("frmFind/DoFind error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   Textbox
  ' Goal:   Set the textbox to the one now active
  ' History:
  ' 29-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public WriteOnly Property Textbox() As RichTextBox
    Set(ByVal value As RichTextBox)
      rtbThis = value
    End Set
  End Property

  ' ------------------------------------------------------------------------------------
  ' Name:   frmFind_Activated
  ' Goal:   What to do when the FIND form is activated
  ' History:
  ' 07-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmFind_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
    Try
      ' Access the Find textbox
      With Me.tbFind
        ' Does textbox contain text?
        If (.TextLength > 0) Then
          ' Select the text
          .Select(0, .TextLength)
        End If
        ' Check the content
        If (.Text = "") Then
          ' Put default text here
          .Text = DEFAULT_XPATH
        End If
        ' Set focus on the textbox
        .Focus()
        .SelectionStart = 0
        .SelectionLength = .TextLength
      End With
    Catch ex As Exception
      ' Show error
      HandleErr("frmFind/Activated error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   frmFindX_VisibleChanged
  ' Goal:   If necessary save changes
  ' History:
  ' 04-04-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmFindX_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
    Try
      ' Are we losing visibility?
      If (Not Me.Visible) AndAlso (bDirty) Then
        ' Possibly accept changes
        tdlSettings.AcceptChanges()
        ' Save changes
        XmlSaveSettings(strSetFile)
        ' Switch off dirty flag
        bDirty = False
      ElseIf (Me.Visible) Then
        ' We become visible again, so re-load the combobox
        If (Not InitXrelCombo()) Then
          ' Warn??
        End If
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmFind/VisibleChanged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   frmFind_KeyUp
  ' Goal:   Look for ESC key pressed
  ' History:
  ' 06-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmFind_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
    ' Check whether escape is pressed
    If (e.KeyCode = Keys.Escape) Then
      ' Leave this form
      Me.Hide()
    End If
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   frmFind_Load
  ' Goal:   Initialise form
  ' History:
  ' 06-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmFind_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Initialise the search string
    Me.tbFind.Text = ""
    ' Set my owner
    Me.Owner = frmMain
    ' Trigger initialisation
    Me.Timer1.Enabled = True
    ' Call activation
    frmFind_Activated(sender, e)
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   ShowSelect
  ' Goal:   Show the selection
  ' History:
  ' 07-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub ShowSelect()
    ' Access the RTB of the active form
    With rtbThis
      ' Set colours to selection
      .SelectionBackColor = Color.DarkBlue
      .SelectionColor = Color.White
      ' Store values of this selection
      intSelStart = .SelectionStart
      intSelLen = .SelectionLength
      ' Indicate that selection was not reset
      bSelReset = False
    End With
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   rtbThis_Click
  ' Goal:   Close FIND form when user clicks on [frmThis]
  ' History:
  ' 07-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub rtbThis_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtbThis.Click
    ' Clicking on the window where find takes place should close the Find window
    Me.Hide()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   rtbThis_SelectionChanged
  ' Goal:   Reset selection when user starts typing
  ' History:
  ' 07-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub rtbThis_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtbThis.SelectionChanged
    Dim intThisStart As Integer
    Dim intThisLen As Integer

    ' Do we need to act?
    If (bSelReset) Then Exit Sub
    With rtbThis
      ' Get the current selection
      intThisStart = .SelectionStart
      intThisLen = .SelectionLength
      ' Delete the previous selection
      If (intSelLen > 0) Then
        ' Select previous stuff
        .Select(intSelStart, intSelLen)
        ' Reverse colours again
        .SelectionBackColor = Color.White
        .SelectionColor = Color.Black
        ' Reset selection
        .Select(intThisStart, intThisLen)
      End If
    End With
    ' Indicate that selection was reset
    bSelReset = True
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Start initialisation
  ' History:
  ' 04-04-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Try
      ' Switch off timer
      Me.Timer1.Enabled = False
      ' Initialise the editor
      If (Not InitEditors()) Then
        Status("Cannot initialise the database results editor")
        Exit Sub
      End If
      ' Switch the init flag on
      bInit = True
    Catch ex As Exception
      ' Show error
      HandleErr("frmFindX/Timer error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   InitEditors
  ' Goal:   Initialise any editors with datagridviews and other controls
  '           that are bound to the dataset
  ' History:
  ' 04-04-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitEditors() As Boolean
    Try
      ' Do away with possible previous handlers
      DgvClear(objQryEd)      ' Query database Editor
      DgvClear(objQreRelEd)   ' Query database Editor
      ' Initialise the query editor (QryEd)
      InitQueryEditor()
      ' Initialise the query relations editor (QryRelEd)
      InitQryRelEditor()
      ' Return success
      Return True
    Catch ex As Exception
      ' Show user that there is an error
      HandleErr("frmFindX/InitEditors error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitQueryEditor
  ' Goal:   Initialise editors used for handling Xpath queries
  ' History:
  ' 04-04-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitQueryEditor() As Boolean
    Try
      ' Initialise the Results Database DGV handle
      objQryEd = New DgvHandle
      With objQryEd
        .Init(Me, tdlSettings, "Query", "QueryId", "Name ASC", "Name;Search", "", _
              "", "", Me.dgvQry, Nothing)
        .BindControl(Me.tbQryName, "Name", "textbox")
        .BindControl(Me.tbQrySearch, "Search", "textbox")
        .BindControl(Me.tbQryDescr, "Descr", "richtext")
        ' Set the parent table for the [Query] editor
        .ParentTable = "QueryList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Switch off multiselect
      Me.dgvQry.MultiSelect = False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show user that there is an error
      HandleErr("frmFindX/InitQueryEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitXrelCombo
  ' Goal:   Initialise the combobox for the Xrelations
  ' History:
  ' 20-04-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitXrelCombo() As Boolean
    Dim intI As Integer ' Counter
    Dim dtrFound() As DataRow ' Relation names

    Try
      ' Initialise the comboboxes
      dtrFound = tdlSettings.Tables("Xrel").Select("", "Name ASC")
      With Me.cboRelName
        ' Clear the entries
        .Items.Clear()
        For intI = 0 To dtrFound.Length - 1
          ' Copy the name of this color
          .Items.Add(dtrFound(intI).Item("Name"))
        Next intI
        ' Set the value to "Black"
        .SelectedIndex = 0
      End With
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show user that there is an error
      HandleErr("frmFindX/InitXrelCombo error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitQryRelEditor
  ' Goal:   Initialise editor used for handling Xpath query relations
  ' History:
  ' 04-04-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitQryRelEditor() As Boolean
    Try
      ' Initialise the comboboxes
      If (Not InitXrelCombo()) Then Return False
      ' Initialise the Results Database DGV handle
      objQreRelEd = New DgvHandle
      With objQreRelEd
        .Init(Me, tdlSettings, "QryRel", "QryRelId", "QryRelId ASC", "QryRelId;RelName;RelCond", "", _
              "", "", Me.dgvQryRel, AddressOf FillRelArg)
        .BindControl(Me.cboRelName, "RelName", "combo")
        .BindControl(Me.cboRelArg, "RelArg", "combo")
        .BindControl(Me.tbRelCond, "RelCond", "textbox")
        'Do NOT set the parent table for the [Query] editor -- it is determined by Qry SelectionChanged
        '.ParentTable = "Query"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Switch off multiselect
      Me.dgvQry.MultiSelect = False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show user that there is an error
      HandleErr("frmFindX/InitQryRelEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   dgvQry_SelectionChanged
  ' Goal:   Set the filter for the query relations
  ' History:
  ' 04-04-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub dgvQry_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvQry.SelectionChanged
    Dim intQueryId As Integer    ' THe ID of the currently selected query

    Try
      ' Is there a [PF object] already?
      If (Not objQryEd Is Nothing) AndAlso (Not objQreRelEd Is Nothing) Then
        ' Get the ID of the template
        intQueryId = objQryEd.SelectedId
        ' Set the filter for the [Template] section
        objQreRelEd.Filter = "QueryId=" & intQueryId
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFindX/dgvTemplate error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------
  ' Name:       MakeDirty()
  ' Goal:       Signal that something needs to be saved
  ' History:
  ' 04-04-2012  ERK Created
  '---------------------------------------------------------------
  Private Sub MakeDirty()
    ' Are we already set?
    If (Not bDirty) AndAlso (bInit) Then
      ' Set dirty bit
      bDirty = True
    End If
  End Sub

  Private Sub tbQryName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbQryName.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objQryEd, Me.tbQryName, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbQrySearch_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbQrySearch.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objQryEd, Me.tbQrySearch, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbQryDescr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbQryDescr.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objQryEd, Me.tbQryDescr, bInit)) Then MakeDirty()
  End Sub

  Private Sub cboRelName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboRelName.SelectedIndexChanged
    ' Make sure changes are processed
    If (CtlChanged(objQreRelEd, Me.cboRelName, bInit)) Then MakeDirty()
  End Sub
  Private Sub cboRelArg_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboRelArg.SelectedIndexChanged
    ' Make sure changes are processed
    If (CtlChanged(objQreRelEd, Me.cboRelArg, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbRelCond_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbRelCond.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objQreRelEd, Me.tbRelCond, bInit)) Then MakeDirty()
  End Sub

  '---------------------------------------------------------------
  ' Name:       mnuQryNew_Click()
  ' Goal:       Create a new Query
  ' History:
  ' 04-04-2012  ERK Created
  '---------------------------------------------------------------
  Private Sub mnuQryNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuQryNew.Click
    Try
      ' Try adding a new query
      TryAddNew(objQryEd, "", "Name", "<Provide query name>")
      ' Set the focus to the right point
      With Me.tbQryName
        .SelectAll()
        .Focus()
      End With
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFindX/QryNew error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub mnuRelNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRelNew.Click
    Dim intQueryId As Integer ' Currently selected query

    Try
      ' See if a query is selected
      intQueryId = objQryEd.SelectedId
      ' Check if we can go ahead
      If (intQueryId >= 0) Then
        ' Try adding a new query-relation
        TryAddNew(objQreRelEd, "", "QueryId", intQueryId)
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFindX/RelNew error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  TryAddNew
  ' Goal :  Try add a new element
  ' History:
  ' 02-10-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub TryAddNew(ByRef objThis As DgvHandle, ByVal strElement As String, _
                        ByVal ParamArray arFldVal() As String)
    Dim strElName As String = ""      ' Name of the new element
    Dim strText As String = ""        ' Initial text of the query
    Dim bRemNodes As Boolean = False  ' Whether to remove nodes or not
    Dim bPrintIdc As Boolean = False  ' Whether to print indices or not
    Dim arAllFldVal() As String       ' Array that combines all field/value combinations
    Dim intI As Integer               ' Counter

    Try
      ' Validate
      If (objThis Is Nothing) Then Exit Sub
      ' Check if we have an element supplied
      If (strElement = "") Then
        ' Make a unified array
        If (arFldVal.Length > 0) Then
          arAllFldVal = arFldVal
        Else
          ReDim arAllFldVal(0)
        End If
      Else
        ' Try get a name for this element
        ' strElName = GetQueryName("What is the name of the new " & strElement & "?")
        strElName = InputBox("What is the name of the new " & strElement & "?")
        ' Validate
        If (strElName = "") Then
          ' Warn the user
          MsgBox("You should provide the " & strElement & " with a name. Try again!")
          Exit Sub
        End If
        ' Does this name entry already exist?
        If (objThis.Exists("Name='" & strElName & "'")) Then
          ' Warn user and leave!
          MsgBox("A " & strElement & " with this name is already present in the list." & vbCrLf & _
                 "Try again with another name!")
          Exit Sub
        End If
        ' Make a unified array
        If (arFldVal.Length > 0) Then
          ReDim arAllFldVal(0 To arFldVal.Length + 1)
          arAllFldVal(0) = "Name"
          arAllFldVal(1) = strElName
          For intI = 0 To arFldVal.Length - 1
            arAllFldVal(2 + intI) = arFldVal(intI)
          Next intI
        Else
          ReDim arAllFldVal(0 To 1)
          arAllFldVal(0) = "Name"
          arAllFldVal(1) = strElName
        End If
      End If
      ' Call function to perform the action
      If (objThis.AddNew(arAllFldVal)) Then
        ' Set dirty flag
        MakeDirty()
      Else
        ' Failure: leave
        Exit Sub
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFindX/TryAddNew error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  GetQueryName
  ' Goal :  Get a name for the new element you want to add
  ' History:
  ' 02-06-2010  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Function GetQueryName(ByVal strQuestion As String) As String
    Dim strName As String = ""  ' The name to be returned
    Try
      ' Get a name for the query
      With frmGetName
        ' Set the owner
        .SetOwner(Me)
        ' Set the caption
        .Text = strQuestion
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.Cancel
            ' Don't continue
            Return ""
          Case Windows.Forms.DialogResult.OK
            ' Retrieve the query name etc.
            strName = .ElementName
        End Select
      End With
      ' Do trim the name...
      strName = Trim(strName)
      ' Return this name
      Return strName
    Catch ex As Exception
      ' Warn the user
      HandleErr("modMain/GetQueryName error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  '---------------------------------------------------------------
  ' Name:       mnuQryDel_Click()
  ' Goal:       Remove currently selected query
  ' History:
  ' 04-04-2012  ERK Created
  '---------------------------------------------------------------
  Private Sub mnuQryDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuQryDel.Click
    Try
      ' Is something selected or not?
      If (dgvQry.SelectedCells.Count > 0) Then
        ' Try to remove it
        TryRemove(objQryEd)
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFindX/QryDel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub mnuRelDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRelDel.Click
    Try
      ' Is something selected or not?
      If (dgvQryRel.SelectedCells.Count > 0) Then
        ' Try to remove it
        TryRemove(objQreRelEd)
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFindX/RelDel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  TryRemove
  ' Goal :  Delete the selected element from the dgv in [objThis]
  ' History:
  ' 02-06-2010  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub TryRemove(ByRef objThis As DgvHandle)
    Dim strQ As String = "Would you really like to delete the selected row?"

    Try
      ' First enquire
      If (MsgBox(strQ, MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes) Then
        ' TRy to remove the selected object
        If (Not objThis.Remove) Then
          ' Unsuccesful...
          Status("Unable to delete this element")
        Else
          ' Make sure dirty bit is set
          MakeDirty()
          ' Show status
          Status("Deleted")
        End If
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFindX/TryRemove error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  TryRTabControl1_SelectedIndexChangedemove
  ' Goal :  Initialise the view, depending on the tab page
  ' History:
  ' 17-04-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
    Try
      Select Case Me.TabControl1.SelectedTab.Name
        Case Me.tpQuery.Name
          ' Make sure the controls are refreshed
          objQryEd.Refresh()
          objQreRelEd.Refresh()
        Case Me.tpResults.Name
          ' Make sure the controls are refreshed
        Case Me.tpXpath.Name
      End Select
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFindX/Tabcontrol error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '' ---------------------------------------------------------------------------------------------------------
  '' Name :  dgvQryRel_SelectionChanged
  '' Goal :  Clear and fill the combobox with the available lines
  '' History:
  '' 17-04-2012  ERK Created
  '' ---------------------------------------------------------------------------------------------------------
  'Private Sub dgvQryRel_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvQryRel.SelectionChanged
  '  Dim intQryRelid As Integer    ' ID of selected relation
  '  Dim strArgVal As String       ' The argument value as it is right now
  '  Dim intI As Integer           ' Counter

  '  Try
  '    ' Check we are initialised
  '    If (Not bInit) Then Exit Sub
  '    ' Find out which line is selected
  '    intQryRelid = objQreRelEd.SelectedId
  '    ' Check
  '    If (intQryRelid < 0) Then Exit Sub
  '    ' Something is selected: access the combobox
  '    With Me.cboRelArg
  '      ' Clear the box
  '      .Items.Clear()
  '      ' Add the source line
  '      .Items.Add("s")
  '      ' Fill with the correct arguments
  '      For intI = 1 To intQryRelid - 1
  '        ' Add this value
  '        .Items.Add(intI)
  '      Next intI
  '    End With
  '  Catch ex As Exception
  '    ' Warn the user
  '    HandleErr("frmFindX/Tabcontrol error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
  '  End Try
  'End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  FillRelArg
  ' Goal :  Fill the input combobox
  ' History:
  ' 20-04-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub FillRelArg(ByVal intQryRelid As Integer)
    Dim dtrQc() As DataRow    ' Result of other query
    Dim intQueryId As Integer ' Currently selected query
    Dim intI As Integer       ' Counter

    Try
      ' Get query
      intQueryId = objQryEd.SelectedId
      ' validate
      If (intQryRelid < 0) OrElse (intQueryId < 0) Then Exit Sub
      ' Fill the [Input] combobox with valid items only
      With Me.cboRelArg
        ' Reset combobox
        .Items.Clear()
        ' At least add "Source" possibility
        .Items.Add("Source")
        ' Visit all possible [QryRelId] members
        For intI = 1 To intQryRelid - 1
          ' Does a line with this [QryRelId] exist within this query?
          dtrQc = tdlSettings.Tables("QryRel").Select("QueryId=" & intQueryId & "AND QryRelId=" & intI)
          If (dtrQc.Length > 0) Then
            ' Extract output from this line
            .Items.Add(intI & "/out")
          End If
        Next intI
      End With
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFindX/FillRelArg error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  mnuQryConvert_Click
  ' Goal :  Convert the currently selected query to Xpath code
  ' History:
  ' 20-04-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub mnuQryConvert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuQryConvert.Click
    ' Call the actual routine
    DoQryConvert()
  End Sub
  Private Sub DoQryConvert()
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim dtrXrel() As DataRow      ' Result of SELECT on Xrel
    Dim strSrc As String          ' The source value
    Dim strRelName As String      ' Relation name
    Dim strRelArg As String       ' Relation argument
    Dim strRelCond As String      ' Relation Condition
    Dim strXrelType As String     ' Type of Xpath relation
    Dim strXrelXname As String    ' String to be used for the Xpath relation
    Dim intQueryId As Integer     ' Currently selected query
    Dim intI As Integer           ' Counter
    Dim colQry As New StringColl  ' Where we build the query

    Try
      ' Get currently selected query
      intQueryId = objQryEd.SelectedId
      ' Validate
      If (intQueryId < 0) Then Exit Sub
      ' Find the value of "source"
      dtrFound = tdlSettings.Tables("Query").Select("QueryId=" & intQueryId)
      If (dtrFound.Length = 0) Then Status("First define a source") : Exit Sub
      strSrc = dtrFound(0).Item("Search")
      If (strSrc = "") Then Status("Source cannot be empty") : Exit Sub
      ' Start building the query
      colQry.Add("//eTree[tb:Like(@Label, '" & strSrc & "') ")
      ' Select all the Query elements with 'Source' as input
      dtrFound = tdlSettings.Tables("QryRel").Select("QueryId=" & intQueryId & " AND RelArg='Source'", "QryRelId ASC")
      ' Walk all the results
      For intI = 0 To dtrFound.Length - 1
        ' Access this entry
        With dtrFound(intI)
          ' Get the relation's name, argument and condition
          strRelName = .Item("RelName")
          strRelArg = .Item("RelArg")
          strRelCond = .Item("RelCond")
        End With
        ' Identify the [RelName] in [Xrel]
        dtrXrel = tdlSettings.Tables("Xrel").Select("Name='" & strRelName & "'")
        ' Validate
        If (dtrXrel.Length > 0) Then
          ' Get the Xrel type
          strXrelType = dtrXrel(0).Item("Type")
          strXrelXname = dtrXrel(0).Item("Xname")
          ' Action depends on the type
          Select Case strXrelType
            Case "Axis"
              colQry.Add(" and " & strXrelXname & "[tb:Like(@Label, '" & strRelCond & "')]")
            Case "Function"
              colQry.Add(" and (" & strXrelXname & " " & strRelCond & ")")
            Case "Attribute"
              colQry.Add(" and tb:Like(string(" & strXrelXname & "), '" & strRelCond & "')")
          End Select
        End If
      Next intI
      ' Finish the query
      colQry.Add("]")
      ' Put the query on the display
      Me.tbFind.Text = colQry.Text
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFindX/DoQryConvert error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Class