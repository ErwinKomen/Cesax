Imports System.Xml
Public Class frmDbReplace
  ' ================================= LOCAL VARIABLES =========================================================
  Private loc_strFilter As String = _
      "//Result[" & vbCrLf & _
      "  child::Feature[@Name='(FeatureName)' and tb:matches(@Value, '(FeatureValue)')]" & vbCrLf & _
      "]" & vbCrLf
  Private colFeature As New StringColl  ' Collection of feature names and values
  Private colResultA As New StringColl  ' Collection of result attribute/node names and values
  Private arResultAttr() As String = {"Search", "TextId", "Cat", "Period", "Notes", "Status"}
  Private arResultNode() As String = {"Text", "Psd"}
  Private intSelectedId As Integer      ' ID of result record that is selected
  Private bInit As Boolean = False      ' Initialisation flag
  Private bDirty As Boolean = False     ' Need saving or not?
  ' ============================================================================================================
  '---------------------------------------------------------------------------------------------------------
  ' Name:       XpathFilter()
  ' Goal:       Return the value of [loc_strFilter]
  ' History:
  ' 22-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public ReadOnly Property XpathFilter() As String
    Get
      Return loc_strFilter
    End Get
  End Property
  '---------------------------------------------------------------------------------------------------------
  ' Name:       frmDbReplace_Load()
  ' Goal:       Start-up material
  ' History:
  ' 22-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub frmDbReplace_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Set me somewhere
    Me.Owner = frmMain
    ' Start the timer for initialisation
    Me.Timer1.Enabled = True
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       Timer1_Tick()
  ' Goal:       initialise
  ' History:
  ' 22-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Dim ndxList As XmlNodeList  ' List of nodex
    Dim ndxThis As XmlNode      ' Working node
    Dim crsThis As Windows.Forms.Cursor = Cursor.Current
    Dim intI As Integer         ' Counter

    Try
      ' only continue if visible
      If (Not Me.Visible) Then Exit Sub
      ' Validate
      If (pdxCrpDbase Is Nothing) Then Status("No corpus results available") : Exit Sub
      ' Switch off the timer
      Me.Timer1.Enabled = False
      ' Show we are busy
      Status("Initialising...") : Cursor.Current = Cursors.WaitCursor
      ' Perform initialisations...
      ' (1) Determine the names of the features by looking at the first record that contains them
      ndxThis = pdxCrpDbase.SelectSingleNode("//Result")
      If (ndxThis Is Nothing) Then
        Status("The database does not contain results")
        Exit Sub
      End If
      ndxList = ndxThis.SelectNodes("./child::Feature")
      If (ndxList.Count > 0) Then
        ' Reset the combobox
        With Me.cboFindFeatureName
          ' Clear it
          .Items.Clear()
          ' Get a list of features and load the combobox
          For intI = 0 To ndxList.Count - 1
            .Items.Add(ndxList(intI).Attributes("Name").Value)
          Next intI
          ' Set the first one
          .SelectedIndex = 0
        End With
      End If
      ' (2) Add the names for Result characteristics
      With Me.cboFindResultName
        ' Clear
        .Items.Clear()
        ' Load the attributes
        For intI = 0 To arResultAttr.Length - 1
          .Items.Add(arResultAttr(intI))
        Next intI
        ' Load the nodes
        For intI = 0 To arResultNode.Length - 1
          .Items.Add(arResultNode(intI))
        Next intI
        ' Set the first one
        .SelectedIndex = 0
      End With
      ' (3) Load a default filter
      AdaptXpath()
      ' Adapt the combination editor
      InitEditors()
      ' Show we are ready
      bInit = True
      Status("Ready")
      Cursor.Current = crsThis
    Catch ex As Exception
      ' Show error
      HandleErr("frmDbReplace/Timer1 error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   InitEditors
  ' Goal:   Initialise the different editors with datagridviews and other controls
  '           that are bound to the dataset
  ' History:
  ' 16-10-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitEditors() As Boolean
    Try
      ' Do away with possible previous handlers
      DgvClear(objFRed)    ' Find and replacement Editor
      ' Initialise the Find and replacement Editor
      InitFindReplEditor()
      ' Return success
      InitEditors = True
    Catch ex As Exception
      ' Show user that there is an error
      HandleErr("frmSetting/InitEditors error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      InitEditors = False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitFindReplEditor
  ' Goal:   Initialise the editor for Find and replacement
  ' History:
  ' 23-09-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitFindReplEditor() As Boolean
    Try
      ' Initialise the Find and replacement handle
      objFRed = New DgvHandle
      With objFRed
        .Init(Me, tdlSettings, "FindRepl", "FindReplId", "FindReplId", "Type;Name;Find;Repl", "", _
              "", "", Me.dgvFindRepl, Nothing)
        .BindControl(Me.tbFRname, "Name", "textbox")
        .BindControl(Me.tbFRfind, "Find", "richtext")
        .BindControl(Me.tbFRrepl, "Repl", "richtext")
        ' Set the parent table for the [Find and replacement] editor
        .ParentTable = "FindReplList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitFindReplEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       cmdCancel_Click()
  ' Goal:       Close and provide the value "Cancel"
  ' History:
  ' 22-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.DialogResult = Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdAll_Click
  ' Goal:   Copy the annotation from one record to similar records
  ' History:
  ' 11-12-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAll.Click
    Try
      ' See if changes are needed
      If (bDirty) Then
        XmlSaveSettings(strSetFile)
        bDirty = False
      End If
      ' Prepare the dialog result
      Me.DialogResult = Windows.Forms.DialogResult.Ignore
      ' Close the window
      Me.Close()
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmDbReplace/All error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdStep_Click
  ' Goal:   Copy the annotation from one record to similar records
  ' History:
  ' 10-12-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdStep_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStep.Click
    Try
      ' See if changes are needed
      If (bDirty) Then
        XmlSaveSettings(strSetFile)
        bDirty = False
      End If
      ' Prepare the dialog result
      Me.DialogResult = Windows.Forms.DialogResult.OK
      ' Close the window
      Me.Close()
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmDbReplace/Step error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       cmdAdd_Click()
  ' Goal:       Add one feature name/value combination
  ' History:
  ' 22-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddF.Click
    Try
      ' Validate
      If (Me.cboFindFeatureName.SelectedIndex < 0) Then Exit Sub
      If (tdlSettings.Tables("FindRepl") Is Nothing) Then Exit Sub
      ' Add this item
      If (Not AddOneFindRepl(True, Me.cboFindFeatureName.SelectedItem, Me.tbFindFeatureValue.Text, Me.tbFindFeatureRepl.Text)) Then
        Status("Could not add this item")
      Else
        ' Adapt the Xpath filter expression
        AdaptXpath()
        ' Process changes
        tdlSettings.AcceptChanges()
        Status("Okay")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmDbReplace/cmdAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub cmdAddR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddR.Click
    Try
      ' Validate
      If (Me.cboFindResultName.SelectedIndex < 0) Then Exit Sub
      If (tdlSettings.Tables("FindRepl") Is Nothing) Then Exit Sub
      ' Add this item
      If (Not AddOneFindRepl(False, Me.cboFindResultName.SelectedItem, Me.tbFindResultValue.Text, Me.tbFindResultRepl.Text)) Then
        Status("Could not add this item")
      Else
        ' Adapt the Xpath filter expression
        AdaptXpath()
        ' Process changes
        tdlSettings.AcceptChanges()
        Status("Okay")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmDbReplace/cmdAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Function AddOneFindRepl(ByVal bIsFeature As Boolean, ByVal strName As String, ByVal strFind As String, ByVal strRepl As String) As Boolean
    Dim dtrFound() As DataRow ' Looking for info
    Dim dtrNew As DataRow     ' New item to the table

    Dim strType As String = ""  ' Part of the select statement that describes the type

    Try
      ' Make up the [strType]
      ''strType = IIf(bIsFeature, "Type='Feature'", "(Type='Attr' OR Type='Node')")
      If (bIsFeature) Then
        strType = "Feature"
      Else
        strType = IIf(Array.Exists(arResultAttr, Function(strIn As String) strIn = strName), "Attr", "Node")
      End If
      ' Find the correct table row
      dtrFound = tdlSettings.Tables("FindRepl").Select("Type='" & strType & "' AND Name='" & strName & "'")
      If (dtrFound.Length = 0) Then
        ' Add it
        dtrNew = AddOneDataRow(tdlSettings, "FindRepl", "FindReplId", "FindReplList")
      Else
        ' Take this row
        dtrNew = dtrFound(0)
      End If
      ' Adapt the row values
      With dtrNew
        .Item("Type") = strType ' IIf(Array.Exists(arResultAttr, Function(strIn As String) strIn = strName), "Attr", "Node")
        .Item("Name") = strName
        .Item("Find") = strFind
        .Item("Repl") = strRepl
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("frmDbReplace/AddOneFindRepl error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Function

  '---------------------------------------------------------------------------------------------------------
  ' Name:       cmdDel_Click()
  ' Goal:       Delete one feature name/value combination
  ' History:
  ' 22-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub cmdDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelF.Click
    Dim dtrFound() As DataRow ' Looking for info
    Dim intI As Integer       ' Counter

    Try
      ' Validate
      If (Me.cboFindFeatureName.SelectedIndex < 0) Then Exit Sub
      If (tdlSettings.Tables("FindRepl") Is Nothing) Then Exit Sub
      ' Find the correct table row
      dtrFound = tdlSettings.Tables("FindRepl").Select("Type='Feature' AND Name='" & cboFindFeatureName.SelectedItem & "'")
      ' Remove all results
      For intI = dtrFound.Length - 1 To 0 Step -1
        dtrFound(intI).Delete()
      Next intI
      ' make sure changes are processed
      tdlSettings.AcceptChanges()
      ' Adapt the Xpath filter expression
      AdaptXpath()
    Catch ex As Exception
      ' Show error
      HandleErr("frmDbReplace/AdaptXpath error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub cmdDelR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelR.Click
    Dim dtrFound() As DataRow ' Looking for info
    Dim intI As Integer       ' Counter

    Try
      ' Validate
      If (Me.cboFindResultName.SelectedIndex < 0) Then Exit Sub
      If (tdlSettings.Tables("FindRepl") Is Nothing) Then Exit Sub
      ' Find the correct table row
      dtrFound = tdlSettings.Tables("FindRepl").Select("(Type='Attr' OR Type='Node') AND Name='" & cboFindResultName.SelectedItem & "'")
      ' Remove all results
      For intI = dtrFound.Length - 1 To 0 Step -1
        dtrFound(intI).Delete()
      Next intI
      ' make sure changes are processed
      tdlSettings.AcceptChanges()
      ' Adapt the Xpath filter expression
      AdaptXpath()
    Catch ex As Exception
      ' Show error
      HandleErr("frmDbReplace/AdaptXpath error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       cmdReset_Click()
  ' Goal:       Remove all feature/value combinations
  ' History:
  ' 22-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdResetF.Click
    Try
      ' Clear the Find/Replace table from the settings dataset
      ClearTable(tdlSettings.Tables("FindRepl"))
      ' Make sure changes are accepted
      tdlSettings.AcceptChanges()
      ' Adapt the Xpath filter
      AdaptXpath()
    Catch ex As Exception
      ' Show error
      HandleErr("frmDbReplace/cmdReset error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
    'AdaptXpath()
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       AdaptXpath()
  ' Goal:       Build an Xpath expression and store it in loc_strFilter
  ' History:
  ' 22-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub AdaptXpath()
    Dim dtrFound() As DataRow ' Looking for info
    Dim bFirst As Boolean = True  ' Flag for the first feature/node
    Dim strXpath As String  ' Query
    Dim intI As Integer     ' Counter

    Try
      ' Are there any results?
      dtrFound = tdlSettings.Tables("FindRepl").Select("")
      If (dtrFound.Length = 0) Then
        ' Make simple query
        loc_strFilter = "//Result" & vbCrLf
      Else
        ' Start
        strXpath = "//Result[" & vbCrLf
        ' Walk all the results
        For intI = 0 To dtrFound.Length - 1
          ' Access this result
          With dtrFound(intI)
            ' Check if we need to find anything here
            If (.Item("Find").ToString <> "") Then
              ' Add preamble
              If (bFirst) Then
                strXpath &= "  "
                bFirst = False
              Else
                strXpath &= "  and "
              End If
              ' Action depends on the type of this result
              Select Case .Item("Type")
                Case "Attr"
                  ' This is an attribute
                  strXpath &= "tb:matches(string(@" & .Item("Name").ToString & "), '" & .Item("Find").ToString & "')" & vbCrLf
                Case "Node"
                  ' This is a child node
                  strXpath &= "tb:matches(string(child::" & .Item("Name").ToString & "), '" & .Item("Find").ToString & "')" & vbCrLf
                Case "Feature"
                  ' Add this one
                  strXpath &= "child::Feature[@Name = '" & .Item("Name").ToString & "' and tb:matches(@Value, '" & .Item("Find").ToString & "')]" & vbCrLf
              End Select
            End If
          End With
        Next intI
        ' Finish
        strXpath &= "]" & vbCrLf
        loc_strFilter = strXpath
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmDbReplace/AdaptXpath error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '---------------------------------------------------------------
  ' Name:       tbFRname_TextChanged()
  ' Goal:       Process changes
  ' History:
  ' 02-09-2013  ERK Created
  '---------------------------------------------------------------
  Private Sub tbFRname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbFRname.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objFRed, Me.tbFRname, bInit)) Then MakeDirty()
    AdaptXpath()
  End Sub

  Private Sub tbFRfind_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbFRfind.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objFRed, Me.tbFRfind, bInit)) Then MakeDirty()
    AdaptXpath()
  End Sub

  Private Sub tbFRrepl_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbFRrepl.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objFRed, Me.tbFRrepl, bInit)) Then MakeDirty()
    AdaptXpath()
  End Sub
  '---------------------------------------------------------------
  ' Name:       MakeDirty()
  ' Goal:       Signal that something needs to be saved
  ' History:
  ' 04-12-2008  ERK Created
  '---------------------------------------------------------------
  Private Sub MakeDirty()
    ' Are we already set?
    If (Not bDirty) AndAlso (bInit) Then
      ' Set dirty bit
      bDirty = True
    End If
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   CopyResults
  ' Goal:   Copy the prepared results throughout the database
  ' History:
  ' 10-12-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function CopyResults(ByRef bDirty As Boolean, Optional ByVal bStep As Boolean = False) As Boolean
    Dim dtrFound() As DataRow     ' Result of SELECT criterion
    Dim ndxList As XmlNodeList    ' List of nodes
    Dim ndxResult As XmlNode      ' One result
    Dim ndxWork As XmlNode        ' Working node (feature or inner node)
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intPtc As Integer         ' Where we are
    ' Dim intCount As Integer       ' Number of items to process
    Dim intSelectedId As Integer  ' The currently selected ID
    Dim bProceed As Boolean       ' Proceed or not?

    Try
      ' Initialized?
      If (Not bInit) Then Return False
      dtrFound = tdlSettings.Tables("FindRepl").Select("Repl <> ''")
      If (dtrFound.Length = 0) Then
        Status("There are no changes that I can made with these specifications")
        Return True
      End If
      '' Test: just count the numbers
      'Status("Counting...")
      'intCount = pdxResults.SelectNodes(loc_strFilter, conTb).Count
      ' Get a list of all results that need handling
      Status("Finding a list of items to replace...")
      ndxList = pdxCrpDbase.SelectNodes(loc_strFilter, conTb)
      Status("Continuing...")
      ' Warn user if there are no results
      If (ndxList.Count = 0) Then MsgBox("No results found for:" & vbCrLf & loc_strFilter) : Return True
      ' Walk through the nodes that need handling
      For intI = 0 To ndxList.Count - 1
        ' Show where I am
        intPtc = (intI + 1) * 100 \ ndxList.Count
        Status("Copying ... " & intPtc & "%", intPtc)
        ' Get a handle on this result
        ndxResult = ndxList(intI)
        ' Set proceding
        bProceed = True
        If (bStep) Then
          ' Select this item
          intSelectedId = ndxResult.Attributes("ResId").Value
          objResEd.SelectDgvId(intSelectedId)
          ' Ask user confirmation
          Select Case MsgBox("Shall I change this item?", MsgBoxStyle.YesNoCancel)
            Case MsgBoxResult.Yes
              bProceed = True
            Case MsgBoxResult.No
              bProceed = False
            Case MsgBoxResult.Cancel
              ' Opt out nicely
              Return True
          End Select
        End If
        ' Can we proceed?
        If (bProceed) Then
          ' Visit the list of changes that need to be made
          For intJ = 0 To dtrFound.Length - 1
            ' Check the kind of change that needs to take place
            Select Case dtrFound(intJ).Item("Type").ToString
              Case "Attr"
                ' Change the attribute's value
                AddXmlAttribute(pdxCrpDbase, ndxResult, dtrFound(intJ).Item("Name").ToString, dtrFound(intJ).Item("Repl").ToString)
                ' ndxResult.Attributes(dtrFound(intJ).Item("Name").ToString).Value = dtrFound(intJ).Item("Repl").ToString
                ' Tell that we have changed
                bDirty = True
              Case "Node"
                ' Get the correct node
                ndxWork = ndxResult.SelectSingleNode("./child::" & dtrFound(intJ).Item("Name").ToString)
                ' Make the child-node if it does not exist yet
                If (ndxWork Is Nothing) Then
                  ndxWork = AddXmlChild(ndxResult, dtrFound(intJ).Item("Name").ToString)
                End If
                ' Change then node's value
                If (ndxWork IsNot Nothing) Then
                  ' Perform the actual replacement
                  ndxWork.InnerText = dtrFound(intJ).Item("Repl").ToString
                  ' Tell that we have changed
                  bDirty = True
                End If
              Case "Feature"
                ' Find the feature we need to have
                ndxWork = ndxResult.SelectSingleNode("./child::Feature[@Name='" & dtrFound(intJ).Item("Name").ToString & "']")
                ' Make the child-node if it does not exist yet
                If (ndxWork Is Nothing) Then
                  ndxWork = AddXmlChild(ndxResult, "Feature", "Name", dtrFound(intJ).Item("Name").ToString)
                End If
                ' Change the feature's value attribute
                If (ndxWork IsNot Nothing) Then
                  ' Perform the actual change
                  'ndxWork.Attributes("Value").Value = dtrFound(intJ).Item("Repl").ToString
                  AddXmlAttribute(pdxCrpDbase, ndxWork, "Value", dtrFound(intJ).Item("Repl").ToString)
                  ' Tell that we have changed
                  bDirty = True
                End If
            End Select
          Next intJ
        End If
      Next intI
      ' Make sure changes are processed
      tdlResults.AcceptChanges()
      ' Return success
      Status("Find and Replace has finished")
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmDbReplace/CopyResults error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   Fill
  ' Goal:   Fill the values in [gpFeat], but keep [gpFeatTo] empty?
  ' History:
  ' 10-12-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Sub Fill(ByVal intId As Integer)
    Try
      ' Store the ID
      intSelectedId = intId
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmDbReplace/Fill error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoFill
  ' Goal:   Take over feature and other values from the currently selected record
  ' History:
  ' 10-12-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdFill_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFill.Click
    DoFill()
  End Sub
  Public Sub DoFill()
    Dim ndxResult As XmlNode      ' Currently selected result
    Dim ndxF As XmlNodeList       ' Features
    Dim strValue As String        ' The value to be taken
    Dim strName As String         ' Name
    Dim intI As Integer           '  Counter

    Try
      ' Get the currently selected one
      ndxResult = pdxCrpDbase.SelectSingleNode("//Result[@ResId=" & intSelectedId & "]")
      ' Find anything?
      If (ndxResult Is Nothing) Then
        ' Warn user (and exit)
        Status("First select a record from the database")
        Exit Sub
      End If
      ' Clear the table as we have it
      ClearTable(tdlSettings.Tables("FindRepl"))
      ' Fill the table with attribute values
      For intI = 0 To arResultAttr.Length - 1
        ' Check if there is a non-zero value
        strName = arResultAttr(intI)
        If (ndxResult.Attributes(strName) IsNot Nothing) Then
          strValue = ndxResult.Attributes(strName).Value
          ' Do not add empty values
          If (strValue <> "") Then
            ' Add this value to the database
            If (Not AddOneFindRepl(False, strName, strValue, "")) Then Status("Could not process [" & strName & "] with value [" & strValue & "]") : Exit Sub
          End If
        End If
      Next intI
      ' Continue with node values
      For intI = 0 To arResultNode.Length - 1
        ' Check if there is a non-zero value
        strName = arResultNode(intI)
        If (ndxResult.SelectSingleNode("./child::" & strName) IsNot Nothing) Then
          strValue = MyTrim(ndxResult.SelectSingleNode("./child::" & strName).InnerText)
          ' Do not add empty values
          If (strValue <> "") Then
            ' Add this value to the database
            If (Not AddOneFindRepl(False, strName, strValue, "")) Then Status("Could not process [" & strName & "] with value [" & strValue & "]") : Exit Sub
          End If
        End If
      Next intI
      ' Continue with feature values
      ndxF = ndxResult.SelectNodes("./child::Feature")
      For intI = 0 To ndxF.Count - 1
        strName = ndxF(intI).Attributes("Name").Value
        strValue = ndxF(intI).Attributes("Value").Value
        ' Do not add empty values
        If (strValue <> "") Then
          ' Add this value to the database
          If (Not AddOneFindRepl(True, strName, strValue, "")) Then Status("Could not process [" & strName & "] with value [" & strValue & "]") : Exit Sub
        End If
      Next intI
      ' Adapt the path
      AdaptXpath()
      ' Show where we are
      Status("You can make the required changes...")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmDbReplace/DoFill error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdFRdel_Click
  ' Goal:   Delete the selected criterion
  ' History:
  ' 02-09-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdFRdel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFRdel.Click
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim intId As Integer    ' ID of selected criterion

    Try
      ' Validate
      If (Not bInit) Then Exit Sub
      intId = objFRed.SelectedId
      If (intId < 0) Then Status("First select a criterion to delete") : Exit Sub
      ' Find the name and the type
      dtrFound = tdlSettings.Tables("FindRepl").Select("FindReplId=" & intId)
      If (dtrFound.Length > 0) Then
        ' Remove it
        dtrFound(0).Delete()
        tdlSettings.AcceptChanges()
      End If
      ' Show we are ready
      Status("Criterion deleted")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmDbReplace/cmdFRdel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdFRkeep_Click
  ' Goal:   Delete the selected criterion
  ' History:
  ' 02-09-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdFRkeep_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFRkeep.Click
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim intId As Integer    ' ID of selected criterion
    Dim intI As Integer     ' Counter

    Try
      ' Validate
      If (Not bInit) Then Exit Sub
      intId = objFRed.SelectedId
      If (intId < 0) Then Status("First select a criterion to keep") : Exit Sub
      ' Find the name and the type
      dtrFound = tdlSettings.Tables("FindRepl").Select("FindReplId<>" & intId)
      ' Walk all criteria
      For intI = dtrFound.Length - 1 To 0 Step -1
        ' Delete this one
        dtrFound(intI).Delete()
      Next intI
      tdlSettings.AcceptChanges()
      ' Show we are ready
      Status("Criteria deleted")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmDbReplace/cmdFRdel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

End Class