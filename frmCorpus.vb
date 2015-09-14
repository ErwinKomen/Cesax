Imports System.Xml
Public Class frmCorpus
  Private bInit As Boolean = False        ' Whether we are initialized or not
  Private intSelectedId As Integer = -1   ' ID of selected result
  Private arFeatIn() As String            ' Input features to look for
  Private arFeatOut() As String           ' Output features to make
  Private colFeatIn As New StringColl     ' Input features to look for
  Private colFeatOut As New StringColl    ' Output features to make
  Private strNotesIn As String = ""       ' Notes to look for
  Private strNotesOut As String = ""      ' Notes to change into
  Private strStatusIn As String = ""      ' Status to look for
  Private strStatusOut As String = ""     ' Status to change into
  Private strCatIn As String = ""         ' Category to look for
  Private strCatOut As String = ""        ' Category to change into
  Private strLastFocus As String = ""     ' Name of editable control that had focus last
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   frmCorpus_Load
  ' Goal:   Trigger initialisation
  ' History:
  ' 10-12-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub frmCorpus_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Try
      ' Trigger initialisation
      Me.Timer1.Enabled = True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmCorpus/Load error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdCancel_Click
  ' Goal:   Copy the annotation from one record to similar records
  ' History:
  ' 10-12-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Try
      ' Prepare the dialog result
      Me.DialogResult = Windows.Forms.DialogResult.Cancel
      ' Close the window
      Me.Close()
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmCorpus/Cancel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
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
      ' Prepare the results
      If (Not PrepareResults()) Then Status("Could not prepare the results") : Exit Sub
      ' Prepare the dialog result
      Me.DialogResult = Windows.Forms.DialogResult.OK
      ' Close the window
      Me.Close()
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmCorpus/Step error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdAll_Click
  ' Goal:   Copy the annotation from one record to similar records
  ' History:
  ' 11-12-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAll.Click
    Try
      ' Prepare the results
      If (Not PrepareResults()) Then Status("Could not prepare the results") : Exit Sub
      ' Prepare the dialog result
      Me.DialogResult = Windows.Forms.DialogResult.Ignore
      ' Close the window
      Me.Close()
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmCorpus/All error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdClear_Click
  ' Goal:   Clear everything except the textbox I am in right now
  ' History:
  ' 14-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click
    Dim tbThis As TextBox ' Feature textbox
    Dim intI As Integer   ' Counter

    Try
      ' Get all the features to look for
      For intI = 1 To intNumFeaturesAllowed
        ' Check input label
        tbThis = Me.gbFeat.Controls("tbDbFeat" & intI)
        If (tbThis IsNot Nothing) AndAlso (tbThis.Visible) Then
          ' Check if this is selected
          If (tbThis.Name <> strLastFocus) Then
            ' Clear the text
            tbThis.Text = ""
          End If
        End If
        ' Check output label
        tbThis = Me.gbFeatTo.Controls("tbDbToFeat" & intI)
        If (tbThis IsNot Nothing) AndAlso (tbThis.Visible) Then
          ' Check if this is selected
          If (tbThis.Name <> strLastFocus) Then
            ' Clear the text
            tbThis.Text = ""
          End If
        End If
      Next intI
      ' Check the other textboxes
      If (Me.tbCat.Name <> strLastFocus) Then Me.tbCat.Text = ""
      If (Me.tbCatTo.Name <> strLastFocus) Then Me.tbCatTo.Text = ""
      If (Me.tbDbNotes.Name <> strLastFocus) Then Me.tbDbNotes.Text = ""
      If (Me.tbDbToNotes.Name <> strLastFocus) Then Me.tbDbToNotes.Text = ""
      If (Me.cboStatus.Name <> strLastFocus) Then Me.cboStatus.Text = ""
      If (Me.cboStatusTo.Name <> strLastFocus) Then Me.cboStatusTo.Text = ""
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmCorpus/Clear error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   PrepareResults
  ' Goal:   Prepare (1) the input selection criteria and (2) the output selection ones
  ' History:
  ' 10-12-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function PrepareResults() As Boolean
    Dim lbThis As Label   ' Feature label
    Dim tbThis As TextBox ' Feature textbox
    Dim intI As Integer   ' Counter

    Try
      ' Initialized?
      If (Not bInit) Then Return False
      ' Get the specification for [Notes], [Cat] and [Status
      strNotesIn = Me.tbDbNotes.Text : strNotesOut = Me.tbDbToNotes.Text
      strCatIn = Me.tbCat.Text : strCatOut = Me.tbCatTo.Text
      strStatusIn = Me.cboStatus.SelectedItem : strStatusOut = Me.cboStatusTo.SelectedItem
      ' Clear feature lists
      colFeatIn.Clear() : colFeatOut.Clear()
      ' Get all the features to look for
      For intI = 1 To intNumFeaturesAllowed
        ' Check input label
        lbThis = Me.gbFeat.Controls("lbDbFeat" & intI)
        tbThis = Me.gbFeat.Controls("tbDbFeat" & intI)
        If (lbThis IsNot Nothing) AndAlso (tbThis IsNot Nothing) _
            AndAlso (lbThis.Visible) AndAlso (tbThis.Visible) Then
          ' Need to process?
          If (tbThis.Text <> "") Then
            ' Add to selection feature and value -- do not TRIM
            colFeatIn.Add(lbThis.Text, tbThis.Text)
          End If
        End If
        ' Check output label
        lbThis = Me.gbFeatTo.Controls("lbDbToFeat" & intI)
        tbThis = Me.gbFeatTo.Controls("tbDbToFeat" & intI)
        If (lbThis IsNot Nothing) AndAlso (tbThis IsNot Nothing) _
           AndAlso (lbThis.Visible) AndAlso (tbThis.Visible) Then
          ' Need to process?
          If (tbThis.Text <> "") Then
            ' Add to output feature and value -- do not TRIM
            colFeatOut.Add(lbThis.Text, tbThis.Text)
          End If
        End If
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmCorpus/PrepareResults error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   CopyResults
  ' Goal:   Copy the prepared results throughout the database
  ' History:
  ' 10-12-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function CopyResults(ByRef bDirty As Boolean, Optional ByVal bStep As Boolean = False) As Boolean
    'Dim dtrFound() As DataRow     ' Result of SELECT criterion
    'Dim dtrChild() As DataRow     ' One particular child
    Dim strFeat As String         ' Name of feature
    Dim strFvalue As String       ' Feature value
    Dim strSearch As String       ' Search string
    Dim ndxThis As XmlNode        ' Working node
    Dim ndxList As XmlNodeList    ' List of feature nodes
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intPtc As Integer         ' Where we are
    Dim intResCount As Integer    ' Number of results
    Dim intSelectedId As Integer  ' The currently selected ID
    Dim bOkay As Boolean          ' Flag
    Dim bProceed As Boolean       ' Proceed or not?

    Try
      ' Initialized?
      If (Not bInit) OrElse (pdxCrpDbase Is Nothing) Then Return False
      If ((colFeatIn.Count = 0) AndAlso strNotesIn = "" AndAlso strCatIn = "" AndAlso strStatusIn = "") OrElse _
        ((colFeatOut.Count = 0) AndAlso strNotesOut = "" AndAlso strCatOut = "" AndAlso strStatusOut = "") Then
        Status("There are no changes that I can made with these specifications")
        Return True
      End If
      ' Find the first result
      ndxThis = pdxCrpDbase.SelectSingleNode("./descendant::Result")
      If (ndxThis Is Nothing) Then Status("No results to be changed") : Return True
      ' Walk through the results
      intResCount = 1 + ndxThis.SelectNodes("./following-sibling::Result").Count
      intI = 0
      While (ndxThis IsNot Nothing)
        ' Show where I am
        intPtc = (intI + 1) * 100 \ intResCount
        Status("Copying ... " & intPtc & "%", intPtc)
        ' Check if this result fulfills the criteria
        ndxList = ndxThis.SelectNodes("./child::Feature")
        bOkay = True
        For intJ = 0 To ndxList.Count - 1
          ' Get the name of the feature
          strFeat = GetAttrValue(ndxList(intJ), "Name")
          ' Does this feature count?
          If (colFeatIn.Exists(strFeat)) Then
            strFvalue = GetAttrValue(ndxList(intJ), "Value")
            ' Make the search string
            strSearch = colFeatIn.Exmp(colFeatIn.Find(strFeat))
            ' Adapt the string for the brackets that may be in here
            strSearch = strSearch.Replace("]", "@@@")
            strSearch = strSearch.Replace("[", "[[]")
            strSearch = strSearch.Replace("@@@", "[]]")
            ' Does this child fulfill the selection criteria?
            If (Not strFvalue Like strSearch) Then
              ' It's not okay
              bOkay = False
              Exit For
            End If
          End If
        Next intJ
        ' Can we check other things?
        If (bOkay) Then
          ' Check the [Notes]
          If (strNotesIn <> "") Then
            ' Yes, check the notes
            If (GetAttrValue(ndxThis, "Notes") <> strNotesIn) Then bOkay = False
          End If
          ' Check the [Cat]
          If (bOkay) AndAlso (strCatIn <> "") Then
            ' Yes, check the cat
            If (GetAttrValue(ndxThis, "Cat") <> strCatIn) Then bOkay = False
          End If
          ' Check the [Status]
          If (bOkay) AndAlso (strStatusIn <> "") Then
            ' Yes, check the status
            If (GetAttrValue(ndxThis, "Status") <> strStatusIn) Then bOkay = False
          End If
        End If
        ' Do we have a match?
        If (bOkay) Then
          ' Set proceding
          bProceed = True
          If (bStep) Then
            ' Select this item
            intSelectedId = GetAttrValue(ndxThis, "ResId")
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
          ' Do we proceed?
          If (bProceed) Then
            ' We have a match, so replace the correct values
            For intJ = 0 To ndxList.Count - 1
              ' What feature is this?
              strFeat = GetAttrValue(ndxList(intJ), "Name")
              ' Does this one need replacement?
              If (colFeatOut.Exists(strFeat)) Then
                ' Replace the value of this feature
                AddXmlAttribute(pdxCrpDbase, ndxList(intJ), "Value", colFeatOut.Exmp(colFeatOut.Find(strFeat)))
              End If
            Next intJ

            ' Do we need to change the [notes] value?
            If (strNotesOut <> "") Then
              ' Yes, replace the notes value
              AddXmlAttribute(pdxCrpDbase, ndxThis, "Notes", strNotesOut)
            End If
            ' Do we need to change the [Cat] value?
            If (strCatOut <> "") Then
              ' Yes, replace the cat value
              AddXmlAttribute(pdxCrpDbase, ndxThis, "Cat", strCatOut)
            End If
            ' Do we need to change the [Status] value?
            If (strStatusOut <> "") Then
              ' Yes, replace the status value
              AddXmlAttribute(pdxCrpDbase, ndxThis, "Status", strStatusOut)
            End If
            ' Tell that we have changed
            bDirty = True
          End If
        End If
        ' Next result
        ndxThis = ndxThis.NextSibling : intI += 1
      End While

      '' Get all the results
      'dtrFound = tdlResults.Tables("Result").Select("", "ResId ASC")
      '' Anything?
      'If (dtrFound.Length = 0) Then Status("No results to be changed") : Return True
      '' Go through all of them
      'For intI = 0 To dtrFound.Length - 1
      '  ' Show where I am
      '  intPtc = (intI + 1) * 100 \ dtrFound.Length
      '  Status("Copying ... " & intPtc & "%", intPtc)
      '  ' Check if this fulfills the criteria
      '  With dtrFound(intI)
      '    '' Get all the children
      '    'dtrChild = .GetChildRows("Result_Feature")
      '    'bOkay = True
      '    'For intJ = 0 To dtrChild.Length - 1
      '    '  ' What feature is this?
      '    '  strFeat = dtrChild(intJ).Item("Name").ToString
      '    '  ' Does this feature count?
      '    '  If (colFeatIn.Exists(strFeat)) Then
      '    '    ' Make the search string
      '    '    strSearch = colFeatIn.Exmp(colFeatIn.Find(strFeat))
      '    '    ' Adapt the string for the brackets that may be in here
      '    '    strSearch = strSearch.Replace("]", "@@@")
      '    '    strSearch = strSearch.Replace("[", "[[]")
      '    '    strSearch = strSearch.Replace("@@@", "[]]")
      '    '    ' Does this child fulfill the selection criteria?
      '    '    If (Not dtrChild(intJ).Item("Value").ToString Like strSearch) Then
      '    '      ' It's not okay
      '    '      bOkay = False
      '    '      Exit For
      '    '    End If
      '    '  End If
      '    'Next intJ
      '    '' Can we check other things?
      '    'If (bOkay) Then
      '    '  ' Check the [Notes]
      '    '  If (strNotesIn <> "") Then
      '    '    ' Yes, check the notes
      '    '    If (dtrFound(intI).Item("Notes").ToString <> strNotesIn) Then bOkay = False
      '    '  End If
      '    '  ' Check the [Cat]
      '    '  If (bOkay) AndAlso (strCatIn <> "") Then
      '    '    ' Yes, check the cat
      '    '    If (dtrFound(intI).Item("Cat").ToString <> strCatIn) Then bOkay = False
      '    '  End If
      '    '  ' Check the [Status]
      '    '  If (bOkay) AndAlso (strStatusIn <> "") Then
      '    '    ' Yes, check the status
      '    '    If (dtrFound(intI).Item("Status").ToString <> strStatusIn) Then bOkay = False
      '    '  End If
      '    'End If
      '    ' Do we have a match?
      '    If (bOkay) Then
      '      ' Set proceding
      '      bProceed = True
      '      If (bStep) Then
      '        ' Select this item
      '        intSelectedId = dtrFound(intI).Item("ResId")
      '        objResEd.SelectDgvId(intSelectedId)
      '        ' Ask user confirmation
      '        Select Case MsgBox("Shall I change this item?", MsgBoxStyle.YesNoCancel)
      '          Case MsgBoxResult.Yes
      '            bProceed = True
      '          Case MsgBoxResult.No
      '            bProceed = False
      '          Case MsgBoxResult.Cancel
      '            ' Opt out nicely
      '            Return True
      '        End Select
      '      End If
      '      ' Do we proceed?
      '      If (bProceed) Then
      '        ' Signal that editing is coming
      '        dtrFound(intI).BeginEdit()
      '        ' We have a match, so replace the correct values
      '        For intJ = 0 To dtrChild.Length - 1
      '          ' What feature is this?
      '          strFeat = dtrChild(intJ).Item("Name").ToString
      '          ' Does this one need replacement?
      '          If (colFeatOut.Exists(strFeat)) Then
      '            ' Replace the value of this feature
      '            dtrChild(intJ).Item("Value") = colFeatOut.Exmp(colFeatOut.Find(strFeat))
      '          End If
      '        Next intJ
      '        ' Do we need to change the [notes] value?
      '        If (strNotesOut <> "") Then
      '          ' Yes, replace the notes value
      '          dtrFound(intI).Item("Notes") = strNotesOut
      '        End If
      '        ' Do we need to change the [Cat] value?
      '        If (strCatOut <> "") Then
      '          ' Yes, replace the cat value
      '          dtrFound(intI).Item("Cat") = strCatOut
      '        End If
      '        ' Do we need to change the [Status] value?
      '        If (strStatusOut <> "") Then
      '          ' Yes, replace the status value
      '          dtrFound(intI).Item("Status") = strStatusOut
      '        End If
      '        ' Tell that we have changed
      '        bDirty = True
      '      End If
      '    End If
      '  End With
      'Next intI
      ' Make sure changes are processed
      tdlResults.AcceptChanges()
      ' Return success
      Status("Annotation has been copied")
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmCorpus/CopyResults error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Copy the annotation from one record to similar records
  ' History:
  ' 10-12-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Dim dtrThis As DataRow        ' Currently selected datarow
    'Dim dtrChildren() As DataRow  ' Children of this datarow
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim strFname As String        ' Feature name
    'Dim strFvalue As String       ' Feature value
    Dim intI As Integer           '  Counter
    Dim lbThis As Label           ' Feature label
    Dim tbThis As TextBox         ' Feature textbox
    Dim ctlThis As Control        ' One control
    Dim ndxThis As XmlNode        ' Working node
    Dim ndxList As XmlNodeList    ' List of feature nodes

    Try
      ' Stop the timer
      Me.Timer1.Enabled = False
      ' Do what is necessary for initialisation
      ' Validate
      If (objResEd Is Nothing) Then Exit Sub
      ' Make sure we are not right now still selecting
      If (objResEd.IsSelecting) Then Exit Sub
      ' Get the datarow of the currently shown set
      dtrThis = objResEd.SelRow : ndxThis = ndxCurrentRes : If (ndxThis Is Nothing) Then Exit Sub
      ' Validate
      If (dtrThis Is Nothing) Then Exit Sub
      ' First swith off everything
      For intI = 1 To intNumFeaturesAllowed
        ctlThis = FindControl(Me.gbFeat, "lbDbFeat" & intI) : If (ctlThis IsNot Nothing) Then ctlThis.Visible = False
        ctlThis = FindControl(Me.gbFeatTo, "lbDbToFeat" & intI) : If (ctlThis IsNot Nothing) Then ctlThis.Visible = False
        ctlThis = FindControl(Me.gbFeat, "tbDbFeat" & intI) : If (ctlThis IsNot Nothing) Then ctlThis.Visible = False
        ctlThis = FindControl(Me.gbFeatTo, "tbDbToFeat" & intI) : If (ctlThis IsNot Nothing) Then ctlThis.Visible = False
        'Me.gbFeat.Controls("lbDbFeat" & intI).Visible = False
        'Me.gbFeatTo.Controls("lbDbToFeat" & intI).Visible = False
        'Me.gbFeat.Controls("tbDbFeat" & intI).Visible = False
        'Me.gbFeatTo.Controls("tbDbToFeat" & intI).Visible = False
      Next intI
      ' Check if this datarow has any children that contain features
      ndxList = ndxThis.SelectNodes("./child::Feature")
      For intI = 0 To ndxList.Count - 1
        ' Get feature name
        strFname = GetAttrValue(ndxList(intI), "Name")
        ' Get from label
        lbThis = Me.gbFeat.Controls("lbDbFeat" & intI + 1)
        ' If possible make label visible
        If (Not lbThis Is Nothing) Then lbThis.Visible = True : lbThis.Text = strFname
        ' Get to label
        lbThis = Me.gbFeatTo.Controls("lbDbToFeat" & intI + 1)
        ' Check
        If (Not lbThis Is Nothing) Then lbThis.Visible = True : lbThis.Text = strFname
        ' Get from textbox
        tbThis = Me.gbFeat.Controls("tbDbFeat" & intI + 1)
        ' Check
        If (Not tbThis Is Nothing) Then tbThis.Visible = True : tbThis.Text = ""
        ' Get to textbox
        tbThis = Me.gbFeatTo.Controls("tbDbToFeat" & intI + 1)
        ' Check
        If (Not tbThis Is Nothing) Then tbThis.Visible = True : tbThis.Text = ""
      Next intI
      'dtrChildren = dtrThis.GetChildRows("Result_Feature")
      'For intI = 0 To dtrChildren.Length - 1
      '  ' Process this child
      '  With dtrChildren(intI)
      '    ' Get label
      '    lbThis = Me.gbFeat.Controls("lbDbFeat" & intI + 1)
      '    ' Check
      '    If (Not lbThis Is Nothing) Then
      '      ' Make label visible
      '      lbThis.Visible = True
      '      lbThis.Text = .Item("Name").ToString
      '    End If
      '    ' Get label
      '    lbThis = Me.gbFeatTo.Controls("lbDbToFeat" & intI + 1)
      '    ' Check
      '    If (Not lbThis Is Nothing) Then
      '      ' Make label visible
      '      lbThis.Visible = True
      '      lbThis.Text = .Item("Name").ToString
      '    End If
      '    ' Get textbox
      '    tbThis = Me.gbFeat.Controls("tbDbFeat" & intI + 1)
      '    ' Check
      '    If (Not tbThis Is Nothing) Then
      '      tbThis.Visible = True
      '      tbThis.Text = ""
      '    End If
      '    ' Get textbox
      '    tbThis = Me.gbFeatTo.Controls("tbDbToFeat" & intI + 1)
      '    ' Check
      '    If (Not tbThis Is Nothing) Then
      '      tbThis.Visible = True
      '      tbThis.Text = ""
      '    End If
      '  End With
      'Next intI
      ' Set the sizes of the input and output feature arrays
      ReDim arFeatIn(0 To ndxList.Count - 1)
      ReDim arFeatOut(0 To ndxList.Count - 1)
      ' ReDim arFeatIn(0 To dtrChildren.Length - 1)
      ' ReDim arFeatOut(0 To dtrChildren.Length - 1)
      ' Fill the status comboboxes from the [tdlSettings]
      dtrFound = tdlSettings.Tables("Status").Select("", "Name ASC")
      Me.cboStatus.Items.Clear() : Me.cboStatusTo.Items.Clear()
      For intI = 0 To dtrFound.Length - 1
        Me.cboStatus.Items.Add(dtrFound(intI).Item("Name").ToString)
        Me.cboStatusTo.Items.Add(dtrFound(intI).Item("Name").ToString)
      Next intI
      ' Add event handler to different controls
      For Each ctlThis In Me.gbFeat.Controls
        ' Add event handler to this control
        AddHandler ctlThis.GotFocus, AddressOf Control_GotFocus
      Next ctlThis
      For Each ctlThis In Me.gbFeatTo.Controls
        ' Add event handler to this control
        AddHandler ctlThis.GotFocus, AddressOf Control_GotFocus
      Next ctlThis
      AddHandler tbDbNotes.GotFocus, AddressOf Control_GotFocus
      AddHandler tbDbNotes.GotFocus, AddressOf Control_GotFocus
      ' Show we are ready
      bInit = True
      Status("Ready")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmCorpus/Timer error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   Control_GotFocus
  ' Goal:   What to do when control lost focus
  ' History:
  ' 14-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub Control_GotFocus(ByVal sender As Object, ByVal e As EventArgs)
    Try
      ' Note the name of this control
      ' strLastFocus = Me.ActiveControl.Name --> somehow this does not work...
      strLastFocus = sender.Name
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmCorpus/Control_GotFocus error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
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
      ' start timer
      Me.tmeFill.Enabled = True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmCorpus/Fill error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoFill
  ' Goal:   Fill the values in [gpFeat], but keep [gpFeatTo] empty?
  ' History:
  ' 10-12-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Sub DoFill()
    'Dim dtrFound() As DataRow     ' Currently selected datarow
    'Dim dtrChildren() As DataRow  ' Children of this datarow
    Dim strValue As String        ' The value to be taken
    Dim intI As Integer           '  Counter
    Dim tbThis As TextBox         ' Feature textbox
    Dim ndxThis As XmlNode        ' Working node
    Dim ndxList As XmlNodeList    ' List of feature nodes

    Try
      ' Get the currently selected one
      ndxThis = ndxCurrentRes : If (ndxThis Is Nothing) Then Status("First select a record from the database") : Exit Sub
      ' Check if there are children containing a feature
      ndxList = ndxThis.SelectNodes("./child::Feature")
      If (ndxList.Count = 0) Then Status("First select a record from the database") : Exit Sub
      For intI = 0 To ndxList.Count - 1
        ' Process this feature
        tbThis = Me.gbFeat.Controls("tbDbFeat" & intI + 1)
        ' Check
        If (Not tbThis Is Nothing) Then
          tbThis.Visible = True
          ' Check what value to take
          strValue = GetAttrValue(ndxList(intI), "Value")
          tbThis.Text = IIf(DoLike(strValue, "-|,"), "", strValue)
        End If
      Next intI
      ' Check the value of [Notes]
      strValue = GetAttrValue(ndxThis, "Notes") : Me.tbDbNotes.Text = IIf(DoLike(strValue, "-|,"), "", strValue)
      ' Check the value of [Cat]
      strValue = GetAttrValue(ndxThis, "Cat") : Me.tbCat.Text = IIf(DoLike(strValue, "-|,"), "", strValue)
      ' Check the value of [Status]
      strValue = GetAttrValue(ndxThis, "Status") : Me.cboStatus.SelectedItem = IIf(DoLike(strValue, "-|,"), "", strValue)

      '' Get the currently selected one
      'dtrFound = tdlResults.Tables("Result").Select("ResId = " & intSelectedId)
      '' Found anything?
      'If (dtrFound.Length = 0) Then
      '  ' Warn user (and exit)
      '  Status("First select a record from the database")
      '  Exit Sub
      'Else
      '  ' Check if this datarow has any children that contain features
      '  dtrChildren = dtrFound(0).GetChildRows("Result_Feature")
      '  For intI = 0 To dtrChildren.Length - 1
      '    ' Process this child
      '    With dtrChildren(intI)
      '      ' Get textbox
      '      tbThis = Me.gbFeat.Controls("tbDbFeat" & intI + 1)
      '      ' Check
      '      If (Not tbThis Is Nothing) Then
      '        tbThis.Visible = True
      '        ' Check what value to take
      '        strValue = dtrChildren(intI).Item("Value")
      '        tbThis.Text = IIf(DoLike(strValue, "-|,"), "", strValue)
      '      End If
      '    End With
      '  Next intI
      '  ' Check the value of [Notes]
      '  strValue = dtrFound(0).Item("Notes").ToString : Me.tbDbNotes.Text = IIf(DoLike(strValue, "-|,"), "", strValue)
      '  ' Check the value of [Cat]
      '  strValue = dtrFound(0).Item("Cat").ToString : Me.tbCat.Text = IIf(DoLike(strValue, "-|,"), "", strValue)
      '  ' Check the value of [Status]
      '  strValue = dtrFound(0).Item("Status").ToString : Me.cboStatus.SelectedItem = IIf(DoLike(strValue, "-|,"), "", strValue)
      'End If
      ' Show where we are
      Status("You can make the required changes...")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmCorpus/DoFill error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tmeFill_Tick
  ' Goal:   Start filling when initialized
  ' History:
  ' 10-12-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tmeFill_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmeFill.Tick
    Try
      ' Check init
      If (Not bInit) Then Exit Sub
      ' Switch off the timer
      Me.tmeFill.Enabled = False
      ' Do the filling
      DoFill()
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmCorpus/DoFill error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub



End Class