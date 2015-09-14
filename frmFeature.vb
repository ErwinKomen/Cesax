Imports System.Xml
Public Class frmFeature
  ' ===================================== LOCAL VARIABLES ===================================================
  Private bDirty As Boolean = False         ' Whether something has changed and needs to be saved
  Private bInit As Boolean = False          ' Whether we are initialized
  Private tblNPfeat As DataTable = Nothing  ' NP feature table
  Private tblFeat As DataTable = Nothing    ' Table to keep the features of this constituent
  Private objFeatDgv As DgvHandle = Nothing ' Handler for the datagrid view
  Private loc_ndxThis As XmlNode            ' The node we are now looking at
  Private loc_strNPfeatName As String = ""  ' Name of currently selected feature
  Private loc_strNPfeatVal As String = ""
  Private loc_bChanged As Boolean = False   ' Whether we have changed
  '----------------------------------------------------------------------------------------------------------
  ' Name:       Changed()
  ' Goal:       Provide the caller with a clue to know whether changes have been made
  ' History:
  ' 02-06-2010  ERK Created
  '----------------------------------------------------------------------------------------------------------
  Public ReadOnly Property Changed() As Boolean
    Get
      Changed = loc_bChanged
    End Get
  End Property
  '----------------------------------------------------------------------------------------------------------
  ' Name:       frmFeature_Load()
  ' Goal:       Set the initial matters for this form
  ' History:
  ' 01-06-2010  ERK Created
  '----------------------------------------------------------------------------------------------------------
  Private Sub frmFeature_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Me.Owner = frmMain
    ' Set my position
    Me.Top = frmMain.Top + frmMain.TabControl1.Top + frmMain.MenuStrip1.Height
    Me.Left = frmMain.Right - Me.Width
    ' Trigger initialisation
    Me.Timer1.Enabled = True
  End Sub

  '----------------------------------------------------------------------------------------------------------
  ' Name:       frmFeature_KeyDown()
  ' Goal:       Get the DEL button
  ' History:
  ' 02-06-2010  ERK Created
  '----------------------------------------------------------------------------------------------------------
  Private Sub frmFeature_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
    Try
      ' First check whether the user wants to remove a line from a datagridview...
      If (e.KeyCode = Keys.Delete) AndAlso (Me.dgvFeature.Focused) Then
        ' Try to delete it
        TryRemove(objFeatDgv)
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFeature/KeyDown error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub   ' Name of currently selected feature value
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  TryRemove
  ' Goal :  Delete the selected element from the dgv in [objThis]
  ' History:
  ' 02-06-2010  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub TryRemove(ByRef objThis As DgvHandle)
    Dim strQ As String = "Would you really like to delete the selected feature?"

    Try
      ' First enquire
      If (MsgBox(strQ, MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes) Then
        ' TRy to remove the selected object
        If (Not objThis.Remove) Then
          ' Unsuccesful...
          Status("Unable to delete this feature")
        Else
          ' Make sure dirty bit is set
          bDirty = True
          ' Show status
          Status("Deleted")
        End If
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFeature/TryRemove error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '----------------------------------------------------------------------------------------------------------
  ' Name:       Timer1_Tick()
  ' Goal:       Initialise the form
  ' History:
  ' 01-06-2010  ERK Created
  '----------------------------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Try
      ' Reset timer
      Me.Timer1.Enabled = False
      ' Wait until visibility
      While (Not Me.Visible)
        Application.DoEvents()
      End While
      ' Are we initialized already?
      If (bInit) Then Exit Sub
      ' Set up the table
      If (tblFeat Is Nothing) Then
        ' Direct the table to the correct place
        tblFeat = tdlSettings.Tables("Feature")
      Else
        ' Clear the table's rows
        ClearTable(tblFeat)
      End If
      ' Set up the datagrid view...
      DgvClear(objFeatDgv)
      objFeatDgv = New DgvHandle
      With objFeatDgv
        '.Init(Me, tdlSettings, "Feature", "FeatureId", "type", "name;value", "", _
        '    "", "", Me.dgvFeature, Nothing)
        .Init(Me, tdlSettings, "Feature", "FeatureId", "type ASC, name ASC", "type;name;value", "", _
            "", "", Me.dgvFeature, Nothing)
        ' Connect textboxes to the feature ingredients
        .BindControl(Me.tbFeatType, "type", "textbox")
        .BindControl(Me.tbFeatName, "name", "textbox")
        .BindControl(Me.tbFeatValue, "value", "richtext")
        ' Set the parent table for the [Feature] editor
        .ParentTable = "FeatureList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      '' Set up the table with NP features
      'If (tblNPfeat Is Nothing) Then
      '  ' Direct the table to the correct place
      '  tblNPfeat = tdlSettings.Tables("NPfeat")
      'Else
      '  ' Clear this table's rows
      '  ClearTable(tblNPfeat)
      'End If
      ' Reset the node we are looking at
      loc_ndxThis = Nothing
      ' Show we are initialized
      Status("Ready")
      bInit = True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFeature/Timer1 error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '----------------------------------------------------------------------------------------------------------
  ' Name:       Node()
  ' Goal:       Initialise the form with the information from this particular node
  ' History:
  ' 01-06-2010  ERK Created
  '----------------------------------------------------------------------------------------------------------
  Public Property Node() As XmlNode
    Get
      ' Return the node we are looking at
      Return loc_ndxThis
    End Get
    Set(ByVal ndxThis As XmlNode)
      Dim ndxType As XmlNodeList  ' Result of looking for <fs>
      Dim ndxName As XmlNodeList  ' List of <f> features
      Dim dtrNPfeatName() As DataRow  ' Result of a SELECT statement
      Dim strType As String       ' Type of this feature
      Dim strName As String       ' Name of this feature
      Dim strValue As String      ' Value of this feature
      Dim intI As Integer         ' Counter
      Dim intJ As Integer         ' Counter
      Dim intId As Integer = 1    ' Unique ID

      Try
        ' We should first be initialized
        While (Not bInit)
          ' Make sure we are interruptable
          Application.DoEvents()
        End While
        ' Validate (of course, this should not be possible...)
        If (ndxThis Is Nothing) Then Exit Property
        ' Set the node we are looking at
        loc_ndxThis = ndxThis
        ' Show the content of this node: the label and its text
        Me.tbFeatContent.Text = loc_ndxThis.Attributes("Label").Value
        ' Status(NodeText(loc_ndxThis))
        Me.StatusStrip1.Items(0).Text = NodeText(loc_ndxThis)
        ' Are there any rows?
        If (tblFeat.Rows.Count > 0) Then
          ' Clear the previous values from the table
          ClearTable(tblFeat)
        End If
        ' Get the NP features of this node, and process them
        ' ndxType = loc_ndxThis.SelectNodes("./fs[@type='NP']")
        ndxType = loc_ndxThis.SelectNodes("./fs")
        ' Walk the list
        For intI = 0 To ndxType.Count - 1
          ' Get the name of this classe
          strType = ndxType(intI).Attributes("type").Value
          ' Process this class of features
          ndxName = ndxType(intI).SelectNodes("./f")
          For intJ = 0 To ndxName.Count - 1
            ' Get the name and the value
            strName = ndxName(intJ).Attributes("name").Value
            strValue = ndxName(intJ).Attributes("value").Value
            ' Add this feature to the DataTable
            tblFeat.Rows.Add(intId, strType, strName, strValue)
            ' Increment the Id
            intId += 1
          Next intJ
        Next intI
        '' Fill the listbox with feature names
        'dtrNPfeatName = tblNPfeat.Select("", "Name ASC")
        '' Start out without a name and reset the listbox
        'strName = "" : Me.lboFeatName.Items.Clear()
        '' Walk through the datarow selection
        'For intI = 0 To dtrNPfeatName.Length - 1
        '  ' Do we need this element?
        '  If (strName <> dtrNPfeatName(intI).Item("Name") & "") Then
        '    ' Add this element to the listbox
        '    strName = dtrNPfeatName(intI).Item("Name") & ""
        '    Me.lboFeatName.Items.Add(strName)
        '  End If
        'Next intI
        ' Reset changes
        loc_bChanged = False
        ' Make sure the dgv handler is refreshed
        Me.objFeatDgv.Refresh()
      Catch ex As Exception
        ' Warn the user
        HandleErr("frmFeature/Node error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      End Try
    End Set
  End Property

  '----------------------------------------------------------------------------------------------------------
  ' Name:       cmdApply_Click()
  ' Goal:       Exit this form WITH saving
  ' History:
  ' 01-06-2010  ERK Created
  '----------------------------------------------------------------------------------------------------------
  Private Sub cmdApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdApply.Click
    Try
      ' Do we need to save changes?
      If (bDirty) Then
        ' Save the changes
        SaveChanges()
        Status("Changes have been saved")
      Else
        Status("No change")
      End If
      ' Just close the form
      ' Me.Visible = False
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFeature/Apply error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Make this form invisible
      Me.Visible = False
    End Try
  End Sub
  '----------------------------------------------------------------------------------------------------------
  ' Name:       cmdCancel_Click()
  ' Goal:       Exit this form without saving
  ' History:
  ' 01-06-2010  ERK Created
  '----------------------------------------------------------------------------------------------------------
  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Try
      ' Do we need to save changes?
      If (bDirty) Then
        ' Ask user
        Select Case MsgBox("Would you like to save changes in features?", MsgBoxStyle.YesNoCancel)
          Case MsgBoxResult.Yes
            ' First save changes
            SaveChanges()
          Case MsgBoxResult.No
          Case MsgBoxResult.Cancel
            ' Exit this subroutine
            Exit Sub
        End Select
      End If
      ' Just close the form
      Me.Visible = False
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFeature/Cancel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Make this form invisible
      Me.Visible = False
    End Try
  End Sub
  '----------------------------------------------------------------------------------------------------------
  ' Name:       cmdApply_Click()
  ' Goal:       Exit this form WITH saving
  ' History:
  ' 01-06-2010  ERK Created
  '----------------------------------------------------------------------------------------------------------
  Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
    Dim strId As String = ""  ' Name of node

    Try
      ' Ask for name
      strId = GetName("The name of this feature?")
      ' Check result
      If (strId = "") Then Status("Could not add feature without name") : Exit Sub
      ' Make room for a feature with this name
      If (Not FeatureAddOne(strId, "")) Then Status("Could not add the feature") : Exit Sub
      ' Indicate that the file has changed (although it actually only changes under cmdSave...)
      loc_bChanged = True
      ' Show what we want
      Me.StatusStrip1.Items(0).Text = "Supply [type] and [value] of the feature"
      Me.tbFeatType.Focus()
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFeature/Add error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Make this form invisible
      Me.Visible = False
    End Try
  End Sub
  Private Function FeatureAddOne(ByVal strFeatName As String, ByVal strFeatValue As String) As Boolean
    Dim dtrFound() As DataRow ' Result of looking for it
    Dim dtrThis As DataRow    ' one datarow
    Dim ndxThis As XmlNode = Nothing

    Try
      ' Validate
      If (tdlSettings Is Nothing) Then Return False
      ' Try find meta id
      dtrFound = tdlSettings.Tables("Feature").Select("name='" & strFeatName & "'")
      If (dtrFound.Length = 0) Then
        ' Add one element to [tdlSettings]
        dtrThis = AddOneDataRow(tdlSettings, "Feature", "FeatureId", "FeatureList")
        ' Add name and value to the feature
        With dtrThis
          .Item("type") = "default"
          .Item("name") = strFeatName
          .Item("value") = strFeatValue
        End With
      Else
        ' Retain the original feature, but change the value
        dtrThis = dtrFound(0)
      End If
      ' Make sure to select this new feature
      objFeatDgv.SelectDgvId(CInt(dtrThis.Item("FeatureId").ToString))
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("frmFeature/FeatureAddOne error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------------------------
  ' Name:       cmdApply_Click()
  ' Goal:       Exit this form WITH saving
  ' History:
  ' 01-06-2010  ERK Created
  '----------------------------------------------------------------------------------------------------------
  Private Sub cmdDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDel.Click
    Dim dtrFound() As DataRow       ' Result of SELECT
    Dim intSelId As Integer = 0     ' Currently selected feature
    Dim strFeatName As String = ""  ' name of feature
    Dim strFeatType As String = ""  ' Type of the feature

    Try
      Me.StatusStrip1.Items(0).Text = "Sorry, not yet implemented"
      ' Find the currently selected feature
      intSelId = objFeatDgv.SelectedId
      ' Validate
      If (intSelId <= 0) Then Status("First select a feature") : Exit Sub
      ' Find type and name of the feature
      dtrFound = tdlSettings.Tables("Feature").Select("FeatureId=" & intSelId)
      If (dtrFound.Length > 0) Then
        ' Get the name and the feature
        With dtrFound(0)
          strFeatName = .Item("name").ToString : strFeatType = .Item("type").ToString
        End With
        ' Try to delete the feature
        DelFeature(loc_ndxThis, strFeatType, strFeatName)
      End If
      ' Try to remove the currently selected feature from this list
      If (objFeatDgv.Remove()) Then
        ' Indicate that the file has been modified
        loc_bChanged = True
        ' Show status
        Status("Removed")
      Else
        Status("Could not remove feature")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFeature/Del error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Make this form invisible
      Me.Visible = False
    End Try
  End Sub
  '----------------------------------------------------------------------------------------------------------
  ' Name:       SaveChanges()
  ' Goal:       Save changes made by the user
  ' History:
  ' 02-06-2010  ERK Created
  '----------------------------------------------------------------------------------------------------------
  Private Sub SaveChanges()
    Dim intI As Integer   ' Counter

    Try
      ' Validate
      If (tblFeat Is Nothing) Or (Not bInit) Then Exit Sub
      '' Delete the <fs type="NP"> child with all its children
      'DelXmlNodeChild(loc_ndxThis, "fs", "type;NP")
      ' Walk all the features within [tblFeat]
      For intI = 0 To tblFeat.Rows.Count - 1
        ' Access this row
        With tblFeat.Rows(intI)
          ' Process this row
          AddFeature(pdxCurrentFile, loc_ndxThis, .Item("type"), .Item("name") & "", .Item("value") & "")
          '' Make sure the caller knows changes have been made
          'loc_bChanged = True
        End With
      Next intI
      ' Reset dirty flag
      bDirty = False
      ' Also add a history feature
      ' OLD: AddHistory(loc_ndxThis, "NP", "User")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFeature/SaveChanges error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '----------------------------------------------------------------------------------------------------------
  ' Name:       lboFeatName_SelectedIndexChanged()
  ' Goal:       When the [Name] listbox's index changes, then the [Value] listbox can be filled
  ' History:
  ' 02-06-2010  ERK Created
  '----------------------------------------------------------------------------------------------------------
  Private Sub lboFeatName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim strVariants As String ' Variants found
    Dim arVariant() As String ' Array of variants
    Dim intI As Integer       ' Counter

    Try
      ' Are we initialised?
      If (Not bInit) Then Exit Sub
      ' Does the table exist?
      If (tblNPfeat Is Nothing) Then Exit Sub
      '' Clear the current feature value table
      'Me.lboFeatValue.Items.Clear()
      '' Determine which feature NAME is selected
      'loc_strNPfeatName = Me.lboFeatName.SelectedItem
      '' Find out which NP feature values belong to this
      'dtrFound = tblNPfeat.Select("Name='" & loc_strNPfeatName & "'")
      '' Found anything?
      'If (dtrFound.Length > 0) Then
      '  ' Get this stuff
      '  strVariants = dtrFound(0).Item("Variants")
      '  ' Divide into an array
      '  arVariant = Split(strVariants, GetDelim(strVariants, vbCrLf, vbCr, vbLf))
      '  ' Walk all variants
      '  For intI = 0 To arVariant.Length - 1
      '    ' Add this item
      '    Me.lboFeatValue.Items.Add(arVariant(intI).Replace(";", vbTab))
      '  Next intI
      'End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFeature/lboFeatName error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ''----------------------------------------------------------------------------------------------------------
  '' Name:       lboFeatValue_DoubleClick()
  '' Goal:       Add the selected NP feature name/value to the current node's inventory
  '' History:
  '' 02-06-2010  ERK Created
  ''----------------------------------------------------------------------------------------------------------
  'Private Sub lboFeatValue_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs)
  '  Dim arVal() As String     ' Feature value
  '  Dim intId As Integer      ' New ID value
  '  Dim dtrFound() As DataRow ' Result of SELECT

  '  Try
  '    ' Are we initialised?
  '    If (Not bInit) Then Exit Sub
  '    ' Does the table exist?
  '    If (tblNPfeat Is Nothing) Then Exit Sub
  '    ' Find out what the feature value is
  '    arVal = Split(Me.lboFeatValue.SelectedItem, vbTab)
  '    ' The first element is the correct one
  '    loc_strNPfeatVal = arVal(0)
  '    ' Determine a new ID value
  '    intId = tblFeat.Rows.Count + 1
  '    ' Find out if there already is a row with [loc_strNPfeatName]
  '    dtrFound = tblFeat.Select("type='NP' AND name='" & loc_strNPfeatName & "'")
  '    If (dtrFound.Length = 0) Then
  '      ' Add this type/name/value to the correct table
  '      tblFeat.Rows.Add(intId, "NP", loc_strNPfeatName, loc_strNPfeatVal)
  '    Else
  '      ' Change the value that is already there
  '      dtrFound(0).Item("value") = loc_strNPfeatVal
  '    End If
  '    ' Note that we have change
  '    bDirty = True
  '  Catch ex As Exception
  '    ' Warn the user
  '    HandleErr("frmFeature/lboFeatValue error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
  '    ' Show that this value has not been added
  '    Status("The selected feature value has NOT been added")
  '  End Try
  'End Sub

  '----------------------------------------------------------------------------------------------------------
  ' Name:       tbFeatType_TextChanged()
  ' Goal:       Process changes!
  ' History:
  ' 07-11-2014  ERK Created
  '----------------------------------------------------------------------------------------------------------
  Private Sub tbFeatType_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbFeatType.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objFeatDgv, Me.tbFeatType, bInit)) Then bDirty = True
  End Sub
  Private Sub tbFeatName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbFeatName.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objFeatDgv, Me.tbFeatName, bInit)) Then bDirty = True
  End Sub
  Private Sub tbFeatValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbFeatValue.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objFeatDgv, Me.tbFeatValue, bInit)) Then bDirty = True
  End Sub
End Class