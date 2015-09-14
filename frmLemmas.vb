Public Class frmLemmas
  ' =================================LOCAL VARIABLES===================================
  Private bInit As Boolean = False        ' Initialisation flag
  Private bDirty As Boolean = False       ' Dirty flag

  ' ------------------------------------------------------------------------------------
  ' Name:   frmLemmas_Load
  ' Goal:   Trigger initialisation
  ' History:
  ' 11-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmLemmas_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Trigger initialisation
    Me.Timer1.Enabled = True
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Call initialiser
  ' History:
  ' 11-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    ' Reset timer
    Me.Timer1.Enabled = False
    ' Call initialisation
    DoInit()
    ' Turn off Accept button
    Me.cmdOk.Enabled = False
    ' Show we are ready
    Status("Ready")
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   DoInit
  ' Goal:   Try to initialise
  ' History:
  ' 11-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub DoInit()
    Dim bAdded As Boolean = False ' Flag

    Try
      ' Is initialisation needed?
      If (Not bInit) Then
        ' Is there a settings object?
        If (tdlMorphDict Is Nothing) Then
          ' Exit, no initialisation
          Status("No initialisation yet (Morphdict object is not yet found) ...")
          Exit Sub
        End If
        ' Indicate we are initialised
        Status("Morphdict loaded from: " & strMorphDictFile)
      End If
      ' Initialise comboboxes and the different editors
      If (Not InitEditors()) Then
        ' Make sure initialisation is not set
        bInit = False
        ' Failure
        Exit Sub
      End If
      ' Set the init flag
      bInit = True
    Catch ex As Exception
      ' Show user that there is an error
      HandleErr("frmSetting/DoInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Make sure initialisation is not set
      bInit = False
    End Try
  End Sub
  '---------------------------------------------------------------
  ' Name:       cmdOk_Click()
  ' Goal:       Apply the indicated changes
  ' History:
  ' 11-03-2013  ERK Created
  '---------------------------------------------------------------
  Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    ' Try to save the settings file, and don't ask the user about it
    If (Not SaveMorphDict(strMorphDictFile, False)) Then
      Status("Unable to change settings to file " & strMorphDictFile)
    End If
  End Sub

  '---------------------------------------------------------------
  ' Name:       cmdCancel_Click()
  ' Goal:       Do not apply the indicated changes
  ' History:
  ' 11-03-2013  ERK Created
  '---------------------------------------------------------------
  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    ' Return negatively
    Me.DialogResult = Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   InitEditors
  ' Goal:   Initialise the different editors with datagridviews and other controls
  '           that are bound to the dataset
  ' History:
  ' 11-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitEditors() As Boolean
    Try
      ' Do away with possible previous handlers
      DgvClear(objM_Lem)    ' Lemma Editor
      DgvClear(objM_VP)     ' VernPos Editor
      DgvClear(objM_Morph)  ' Morph table Editor
      DgvClear(objM_Rem)    ' Remain table Editor
      ' Initialise the lemma editor
      If (Not InitLemmaEditor()) Then Return False
      ' Initialise the VernPos editor
      If (Not InitVernPosEditor()) Then Return False
      ' Initialise the Morph table editor
      If (Not InitMorphEditor()) Then Return False
      ' Initialise the Remain table editor
      If (Not InitRemainEditor()) Then Return False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show user that there is an error
      HandleErr("frmLemma/InitEditors error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitLemmaEditor
  ' Goal:   Initialise the editor for Lemma Types
  ' History:
  ' 23-09-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitLemmaEditor() As Boolean
    Try
      ' Initialise the Phrase Type DGV handle
      objM_Lem = New DgvHandle
      With objM_Lem
        .Init(Me, tdlMorphDict, "Lemma", "LemmaId", "Vern ASC, Pos ASC", "Vern;Pos;MWcat", "", _
               "", "", Me.dgvLemma, Nothing)
        .BindControl(Me.tbLem_Vern, "Vern", "textbox")
        .BindControl(Me.tbLem_Pos, "Pos", "textbox")
        .BindControl(Me.tbLem_MWcat, "MWcat", "textbox")
        ' Set the parent table for the [PhraseType] editor
        .ParentTable = "LemmaList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmLemma/InitLemmaEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitVernPosEditor
  ' Goal:   Initialise the editor for VernPos Types
  ' History:
  ' 23-09-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitVernPosEditor() As Boolean
    Try
      ' Initialise the Phrase Type DGV handle
      objM_VP = New DgvHandle
      With objM_VP
        ' Load the combobox
        With Me.cboVP_Type
          .Items.Clear()
          .Items.Add("LemmaOnly")
          .Items.Add("LemmaFeat")
          .Items.Add("Skip")
        End With
        ' Do other initialisations
        .Init(Me, tdlMorphDict, "VernPos", "VernPosId", "Vern ASC, Pos ASC", "Vern;Pos;Type", "", _
               "", "", Me.dgvVernPos, Nothing)
        .BindControl(Me.tbVP_Vern, "Vern", "textbox")
        .BindControl(Me.tbVP_Pos, "Pos", "textbox")
        .BindControl(Me.tbVP_PrntLabel, "PrntLabel", "textbox")
        .BindControl(Me.cboVP_Type, "Type", "combo")
        .BindControl(Me.tbVP_Lemma, "l", "textbox")
        .BindControl(Me.tbVP_Features, "f", "richtext")
        ' Set the parent table for the [PhraseType] editor
        .ParentTable = "VernPosList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmVernPos/InitVernPosEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitMorphEditor
  ' Goal:   Initialise the editor for Morph Types
  ' History:
  ' 23-09-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitMorphEditor() As Boolean
    Try
      ' Initialise the Phrase Type DGV handle
      objM_Morph = New DgvHandle
      With objM_Morph
        .Init(Me, tdlMorphDict, "Morph", "MorphId", "Vern ASC, Pos ASC, Label ASC, Freq DESC", "Vern;Pos;Label;Freq", "", _
               "", "", Me.dgvMorph, Nothing)
        .BindControl(Me.tbM_Vern, "Vern", "textbox")
        .BindControl(Me.tbM_Pos, "Pos", "textbox")
        .BindControl(Me.tbM_Label, "Label", "textbox")
        .BindControl(Me.tbM_Freq, "Freq", "textbox")
        .BindControl(Me.tbM_File, "File", "textbox")
        .BindControl(Me.tbM_forestId, "forestId", "textbox")
        .BindControl(Me.tbM_EtreeId, "EtreeId", "textbox")
        .BindControl(Me.tbM_Lemma, "l", "textbox")
        .BindControl(Me.tbM_Hist, "h", "textbox")
        .BindControl(Me.tbM_Feat, "f", "richtext")
        ' Set the parent table for the [PhraseType] editor
        .ParentTable = "MorphList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMorph/InitMorphEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitRemainEditor
  ' Goal:   Initialise the editor for Remain Types
  ' History:
  ' 23-09-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitRemainEditor() As Boolean
    Try
      ' Initialise the Phrase Type DGV handle
      objM_Rem = New DgvHandle
      With objM_Rem
        .Init(Me, tdlMorphDict, "Remain", "RemainId", "Freq DESC, Pos ASC, PrntLabel ASC, Vern ASC", "Vern;Pos;PrntLabel;Freq", "", _
               "", "", Me.dgvRemain, Nothing)
        .BindControl(Me.tbRem_Vern, "Vern", "textbox")
        .BindControl(Me.tbRem_Pos, "Pos", "textbox")
        .BindControl(Me.tbRem_PrntLabel, "PrntLabel", "textbox")
        .BindControl(Me.tbRem_Freq, "Freq", "textbox")
        .BindControl(Me.tbRem_File, "File", "textbox")
        .BindControl(Me.tbRem_forestId, "forestId", "textbox")
        .BindControl(Me.tbRem_EtreeId, "EtreeId", "textbox")
        ' Set the parent table for the [PhraseType] editor
        .ParentTable = "RemainList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmRemain/InitRemainEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  tbVP_Lemma_TextChanged
  ' Goal :  (1) Process the change immediately in [tdlMorphDict]
  '         (2) Set dirty flag...
  ' History:
  ' 11-03-2013  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub tbVP_Lemma_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbVP_Lemma.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objM_VP, Me.tbVP_Lemma, bInit)) Then MakeDirty()
  End Sub
  Private Sub cboVP_Type_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboVP_Type.SelectedIndexChanged
    ' Make sure changes are processed
    If (CtlChanged(objM_VP, Me.cboVP_Type, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbVP_Features_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbVP_Features.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objM_VP, Me.tbVP_Features, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbVP_Pos_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbVP_Pos.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objM_VP, Me.tbVP_Pos, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbVP_PrntLabel_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbVP_PrntLabel.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objM_VP, Me.tbVP_PrntLabel, bInit)) Then MakeDirty()
  End Sub
  '---------------------------------------------------------------
  ' Name:       MakeDirty()
  ' Goal:       Signal that something needs to be saved
  ' History:
  ' 11-03-2013  ERK Created
  '---------------------------------------------------------------
  Private Sub MakeDirty()
    ' Are we already set?
    If (Not bDirty) AndAlso (bInit) Then
      ' Set dirty bit
      bDirty = True
      ' Enable Cancel button
      Me.cmdOk.Enabled = True
    End If
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  cmdNew_Click
  ' Goal :  Make a new X, where X depends on which tabpage is visible
  ' History:
  ' 02-06-2010  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub cmdNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNew.Click
    Dim intLen As Integer = 0   ' Length

    Try
      ' Action depends on the tabpage we are on
      Select Case Me.TabControl1.SelectedTab.Name
        Case Me.tpCheck.Name
          ' Double check and execute
          ' TryAddNew(objCed, "Coreference type")
        Case Me.tpLemma.Name
          ' Double check and execute
          TryAddNew(objM_Lem, "Lemma", "Vern")
        Case Me.tpMorph.Name
          ' Double check and execute
          TryAddNew(objM_Morph, "OT-word", "Vern")
        Case Me.tpRemain.Name
          ' Double check and execute
          TryAddNew(objM_Rem, "Remaining word", "Vern")
        Case Me.tpVernPos.Name
          ' Double check and execute
          TryAddNew(objM_VP, "VernPos word", "Vern")
          ' Set focus on good position
          Me.tbVP_Pos.Focus()
      End Select
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmLemmas/New error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  TryAddNew
  ' Goal :  Try add a new element
  ' History:
  ' 02-10-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub TryAddNew(ByRef objThis As DgvHandle, ByVal strElement As String, _
                        ByVal strNameField As String, ByVal ParamArray arFldVal() As String)
    Dim strElName As String = ""      ' Name of the new element
    Dim strText As String = ""        ' Initial text of the query
    Dim bRemNodes As Boolean = False  ' Whether to remove nodes or not
    Dim bPrintIdc As Boolean = False  ' Whether to print indices or not
    Dim arAllFldVal() As String       ' Array that combines all field/value combinations
    Dim arEls() As String             ' more input elements, separated by space
    Dim strCond As String = ""        ' Condition
    Dim intI As Integer               ' Counter

    Try
      ' Validate
      If (objThis Is Nothing) Then Exit Sub
      If (strElement = "") Then strElement = "(element)"
      ' Try get a name for this element
      strElName = GetName("What is the name of the new " & strElement & "?")
      ' Validate
      If (strElName = "") Then
        ' Warn the user
        MsgBox("You should provide the " & strElement & " with a name. Try again!")
        Exit Sub
      End If
      arEls = Split(strElName, " ")
      If (arEls.Length > 1) Then
        strCond = strNameField & "='" & arEls(0) & "' AND Pos='" & arEls(1) & "'"
      Else
        strCond = strNameField & "='" & strElName & "'"
      End If
      ' Does this name entry already exist?
      If (objThis.Exists(strCond)) Then
        ' Warn user and leave!
        Select Case MsgBox("A " & strElement & " with this name is already present in the list." & vbCrLf & _
               "Would you still like to add it?", MsgBoxStyle.YesNoCancel)
          Case MsgBoxResult.No
            Exit Sub
          Case MsgBoxResult.Cancel
            Exit Sub
        End Select
      End If
      ' Make a unified array
      If (arFldVal.Length > 0) Then
        ReDim arAllFldVal(0 To arFldVal.Length + 1)
        arAllFldVal(0) = strNameField
        arAllFldVal(1) = arEls(0)
        For intI = 0 To arFldVal.Length - 1
          arAllFldVal(2 + intI) = arFldVal(intI)
        Next intI
      ElseIf (arEls.Length > 1) Then
        ReDim arAllFldVal(0 To 5)
        arAllFldVal(0) = strNameField
        arAllFldVal(1) = arEls(0)
        arAllFldVal(2) = "Pos"
        arAllFldVal(3) = arEls(1)
        arAllFldVal(4) = "Type"
        arAllFldVal(5) = "LemmaOnly"
      Else
        ReDim arAllFldVal(0 To 1)
        arAllFldVal(0) = strNameField
        arAllFldVal(1) = arEls(0)
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
      HandleErr("frmMorph/TryAddNew error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
    Dim strQ As String = "Would you really like to delete the selected row?"

    Try
      ' Action depends on the tabpage we are on
      Select Case Me.TabControl1.SelectedTab.Name
        Case Me.tpCheck.Name
          ' Double check and execute
          TryRemove(Nothing)
        Case Me.tpLemma.Name
          ' Double check and execute
          TryRemove(objM_Lem)
        Case Me.tpMorph.Name
          ' Double check and execute
          TryRemove(objM_Morph)
        Case Me.tpRemain.Name
          ' Double check and execute
          TryRemove(objM_Rem)
        Case Me.tpVernPos.Name
          ' Double check and execute
          TryRemove(objM_VP)
      End Select
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMorph/Del error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
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
      HandleErr("frmMorph/TryRemove error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------
  ' Name:       SaveMorphDict()
  ' Goal:       Save the results to the indicated location
  ' History:
  ' 04-12-2008  ERK Created
  '---------------------------------------------------------------
  Private Function SaveMorphDict(ByVal strFile As String, Optional ByVal bAsk As Boolean = True) _
    As Boolean
    Try
      ' Is saving needed?
      If (bDirty) AndAlso (bInit) Then
        ' Do we need to ask?
        If (bAsk) Then
          ' Ask user if he wants saving
          Select Case MsgBox("Save morphdict before continuing?", MsgBoxStyle.YesNoCancel)
            Case MsgBoxResult.Ok, MsgBoxResult.Yes
              Call DoSaving(strFile)
              ' Return success
              Status("MorphDict saved")
              Return True
            Case MsgBoxResult.No
              ' Still okay
              Status("okay")
              Return True
            Case MsgBoxResult.Cancel
              ' Return failure
              Status("Aborted")
              Return False
          End Select
        Else
          ' Save without asking
          Call DoSaving(strFile)
          Status("MorphDict saved")
          ' Return success
          Return True
        End If
      Else
        ' Return success ( but nothing is actually saved)
        Return True
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmRemain/SaveMorphDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Function
  '---------------------------------------------------------------
  ' Name:       DoSaving()
  ' Goal:       Save the results to the indicated location
  ' History:
  ' 04-12-2008  ERK Created
  '---------------------------------------------------------------
  Private Sub DoSaving(ByVal strFile As String)
    Dim strBkUp As String   ' Backup file

    Try
      ' Show what we are doing
      Status("Accepting changes")
      ' Accept changes
      tdlMorphDict.AcceptChanges()
      Application.DoEvents()
      ' Save the settings
      Status("Making backup... ")
      ' Make copy to backup file
      strBkUp = IO.Path.GetFileNameWithoutExtension(strFile) & "_" & Format(Today, "ddd") & ".bak"
      strBkUp = IO.Path.GetDirectoryName(strFile) & "\" & strBkUp
      IO.File.Copy(strFile, strBkUp, True)
      Application.DoEvents()
      ' Do the actual saving of the morphological dictionary
      Status("Saving to " & strFile)
      tdlMorphDict.WriteXml(strFile)
      Application.DoEvents()
      ' Turn off the dirty bit
      bDirty = False
      ' Turn off Accept button
      Me.cmdOk.Enabled = False
      ' Change the Cancel button
      Me.cmdCancel.Text = "E&xit"
      ' Show we are ready
      Status("Ready")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmRemain/DoSaving error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   dgvRemain_SelectionChanged
  ' Goal:   Show morphology information of the one that is currently selected
  ' History:
  ' 11-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub dgvRemain_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvRemain.SelectionChanged
    Dim dtrFound() As DataRow ' Result of SELEC
    Dim strVern As String     ' Currently selected vernacular
    Dim strPos As String      ' Currently selected POS
    Dim colShow As New StringColl ' What we want to show
    Dim intI As Integer       ' Counter

    Try
      ' Are we initialized?
      If (Not bInit) Then Exit Sub
      ' Get information of current selection
      strVern = Me.tbRem_Vern.Text
      strPos = Me.tbRem_Pos.Text
      ' Look in the MORPH table
      dtrFound = tdlMorphDict.Tables("Morph").Select("Vern='" & strVern & "' AND Pos='" & strPos & "'", "Label ASC, Freq DESC")
      For intI = 0 To dtrFound.Length - 1
        With dtrFound(intI)
          colShow.Add("Parent=[" & .Item("Label") & "]" & vbTab & _
                      "Freq=[" & .Item("Freq") & "]" & vbTab & _
                      "L=[" & .Item("l") & "]" & vbTab & _
                      "F=[" & .Item("f") & "]")
        End With
      Next intI
      ' Show the result
      Me.tbRem_Morph.Text = colShow.Text
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmRemain/dgvRemain error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Class