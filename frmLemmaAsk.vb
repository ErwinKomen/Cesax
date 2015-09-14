Public Class frmLemmaAsk
  Private bInit As Boolean = False
  Private bBusy As Boolean = False
  Private bAuto As Boolean = False    ' Flag that this is automatic adaptation
  ' ------------------------------------------------------------------------------------
  ' Name:   Word, Pos, Context, MorphLemma
  ' Goal:   Allow caller to set values for [Word] and [Pos] and [Context
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public WriteOnly Property Word() As String
    Set(ByVal value As String)
      Me.tbWord.Text = value
    End Set
  End Property
  Public WriteOnly Property Pos() As String
    Set(ByVal value As String)
      Me.tbPos.Text = value
    End Set
  End Property
  Public WriteOnly Property Context() As String
    Set(ByVal value As String)
      Me.wbContext.DocumentText = value
    End Set
  End Property
  Public WriteOnly Property MorphLemma() As String()
    Set(ByVal value As String())
      ' Sort the values
      Array.Sort(value)
      ' Clear what we have
      With Me.lboMorphLemma.Items
        .Clear()
        ' Add the whole array
        .AddRange(value)
      End With
    End Set
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   Lemma
  ' Goal:   Return the lemma that was chosen by the user
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public ReadOnly Property Lemma() As String
    Get
      Return Trim(Me.tbLemma.Text)
    End Get
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   frmLemmaAsk_Load
  ' Goal:   Load the form
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmLemmaAsk_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Trigger init
    Me.Timer1.Enabled = True
    Me.Owner = frmMain
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Initialise
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Try
      ' Await visibility
      If (Not Me.Visible) OrElse (binit) Then Exit Sub
      ' Switch off timer
      Me.Timer1.Enabled = False
      ' Perform initialisations...
      If (Not InitEditors()) Then
        Status("Cannot initialise the database results editor")
        Exit Sub
      End If
      ' Show we are initialised
      bInit = True
      Status("Ready")
    Catch ex As Exception
      ' Show error
      HandleErr("frmLemmaAsk/MorphLemmaGet error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   cmdOk_Click
  ' Goal:   Close the form
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    Me.DialogResult = Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub
  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.DialogResult = Windows.Forms.DialogResult.Cancel
    bInterrupt = True
    Me.Close()
  End Sub
  Private Sub cmdSaveStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveStop.Click
    Me.DialogResult = Windows.Forms.DialogResult.Yes
    Me.Close()
  End Sub
  Private Sub cmdSkip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSkip.Click
    Me.DialogResult = Windows.Forms.DialogResult.Ignore
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   lboMorphLemma_SelectedIndexChanged
  ' Goal:   Adapt the |chosen| lemma
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub lboMorphLemma_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lboMorphLemma.SelectedIndexChanged
    Dim strChosen As String = ""

    Try
      ' Get the chosen value
      strChosen = Me.lboMorphLemma.SelectedItem
      Me.tbLemma.Text = strChosen
      ' Set the value in the other box
      bAuto = True
      Me.tbOEdictKey.Text = strChosen
    Catch ex As Exception
      ' Show error
      HandleErr("frmLemmaAsk/lboMorphLemma error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   tbOEdictKey_TextChanged
  ' Goal:   Adapt the filter for the values shown in the [dgv]
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub tbOEdictKey_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbOEdictKey.TextChanged
    Try
      If (Not bInit) Then Exit Sub
      ' Are we on Auto?
      If (bAuto) Then
        bAuto = False : Exit Sub
      End If
      ' Start the timer
      Me.Timer2.Enabled = True
    Catch ex As Exception
      ' Show error
      HandleErr("frmLemmaAsk/lboOEdictKey: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Timer2_Tick
  ' Goal:   Adapt the filter for the values shown in the [dgv]
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
    Dim strFilter As String ' Filter

    Try
      If (Not bInit) OrElse (bBusy) Then Exit Sub
      ' Stop the timer
      Me.Timer2.Enabled = False
      ' Show we are busy
      bBusy = True
      ' Adapt the filter
      strFilter = Trim(Me.tbOEdictKey.Text)
      If (strFilter = "") Then
        objOEdictEd.Filter = ""
      Else
        ' What quotation mark do we need?
        If (InStr(strFilter, "'") > 0) Then
          objOEdictEd.Filter = "l LIKE " & """" & strFilter & "*" & """"
        Else
          objOEdictEd.Filter = "l LIKE '" & strFilter & "*'"
        End If
      End If
      ' Show we are ready
      bBusy = False
    Catch ex As Exception
      ' Show error
      HandleErr("frmLemmaAsk/lboOEdictKey: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' -----------------------------------------------------------------------------------
  ' Name:   InitEditors
  ' Goal:   Initialise any editors with datagridviews and other controls
  '           that are bound to the dataset
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitEditors() As Boolean
    Try
      ' Do away with possible previous handlers
      DgvClear(objOEdictEd)   ' Results database Editor
      ' Load the OE dictionary
      If (Not MorphReadOEdict()) Then Return False
      ' Initialise the results editor (ResEd)
      If (Not InitOEdictEditor()) Then Return False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show user that there is an error
      HandleErr("frmMain/InitEditors error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitOEdictEditor
  ' Goal:   Initialise the editor for the database
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitOEdictEditor() As Boolean
    Try
      ' Initialise the Results Database DGV handle
      objOEdictEd = New DgvHandle
      With objOEdictEd
        .Init(Me, tdlOEdict, "Entry", "EntryId", "l", "l;f", "", _
              "", "", Me.dgvOEdict, Nothing)
        .BindControl(Me.tbOElemma, "l", "textbox")
        .BindControl(Me.tbOEfeat, "f", "textbox")
        .BindControl(Me.tbOEpos, "Pos", "textbox")
        ' Set the parent table for the [Results] editor
        .ParentTable = "EntryList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Switch off multiselect
      Me.dgvOEdict.MultiSelect = False
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitOEdictEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   dgvOEdict_SelectionChanged
  ' Goal:   Show the definitions belonging to this entry
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub dgvOEdict_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvOEdict.SelectionChanged
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim intEntryId As Integer ' Selected entry
    Dim intI As Integer       ' Counter
    Dim colDef As New StringColl  ' The definitions

    Try
      ' Validate
      If (Not bInit) Then Exit Sub
      ' Make sure this lemma is shown
      Me.tbLemma.Text = Me.tbOElemma.Text
      ' Get the entry id
      intEntryId = objOEdictEd.SelectedId
      If (intEntryId < 0) Then Exit Sub
      ' Get all definitions
      dtrFound = tdlOEdict.Tables("Sense").Select("EntryId=" & intEntryId, "N ASC")
      ' Are there any senses?
      If (dtrFound.Length = 0) Then
        ' There are no senses, so perhaps there is a "conform" (=see elsewhere)?
        dtrFound = tdlOEdict.Tables("Entry").Select("EntryId=" & intEntryId)
        If (dtrFound.Length > 0) Then
          If (dtrFound(0).Item("Cf").ToString <> "") Then
            colDef.Add("See: " & dtrFound(0).Item("Cf").ToString)
          End If
        End If
      Else
        For intI = 0 To dtrFound.Length - 1
          ' Add this sense + definition
          With dtrFound(intI)
            colDef.Add(intI & ". " & .Item("Def"))
          End With
        Next intI
      End If
      ' Show the result
      Me.tbOEdictEntry.Text = colDef.Text
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmLemmaAsk/dgvOEdict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Class