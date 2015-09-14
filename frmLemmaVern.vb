Public Class frmLemmaVern
  Private bInit As Boolean = False
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   frmLemmaVern_Load
  ' Goal:   Load form
  ' History:
  ' 24-02-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub frmLemmaVern_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Set parenthood
    Me.Owner = frmMain
    ' Trigger initialisation
    Me.Timer1.Enabled = True
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Trigger and perform initialisation
  ' History:
  ' 24-02-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    ' Wait until we become visible
    While (Not Me.Visible)
      Application.DoEvents()
    End While
    ' Switch off timer
    Me.Timer1.Enabled = False
    ' Perform initialisation
    DoInit()
  End Sub
  Private Sub DoInit()
    Try
      ' Do away with possible previous handlers
      DgvClear(objLVed)   ' Results database Editor
      ' Initialise the dependency editor
      InitLemmaVernEditor()
      ' Show we are initialised
      bInit = True
      Status("ready")
    Catch ex As Exception
      ' Show error
      HandleErr("frmLemmaVern/DoInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   InitLemmaVernEditor
  ' Goal:   Initialise the dependency editor
  ' History:
  ' 04-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitLemmaVernEditor() As Boolean
    Dim intI As Integer '  Counter
    Dim arAction() As String = {"add", "change", "delete"}

    Try
      ' Initialise the dependency editor's DGV handle
      objLVed = New DgvHandle
      With objLVed
        .Init(Me, tdlLemmaVern, "Morph", "MorphId", "l ASC, Vern ASC", "MorphId;l;t;Vern;Pos;h", "", _
              "", "", Me.dgvLemmaVern, Nothing)
        .BindControl(Me.tbLemma, "l", "textbox")
        .BindControl(Me.tbPosHead, "t", "textbox")
        .BindControl(Me.tbVern, "Vern", "textbox")
        .BindControl(Me.tbPos, "Pos", "textbox")
        .BindControl(Me.cboAction, "h", "combo")
        ' Set the parent table for the [LemmaVern] editor
        .ParentTable = "MorphList"
        ' Fill the listbox for action
        With Me.cboAction
          .Items.Clear()
          For intI = 0 To arAction.Length - 1
            .Items.Add(arAction(intI))
          Next intI
        End With
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Switch off multiselect
      Me.dgvLemmaVern.MultiSelect = False
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmLemmaVern/InitLemmaVernEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   chbAllDelete_CheckedChanged
  ' Goal:   Mark all entries as "delete" or not
  ' History:
  ' 23-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub chbAllDelete_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbAllDelete.CheckedChanged
    Dim dtrFound() As DataRow ' Selection
    Dim intI As Integer       ' Counter
    Dim intPtc As Integer     ' Percentage

    Try
      dtrFound = tdlLemmaVern.Tables("Morph").Select("")
      For intI = 0 To dtrFound.Length - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ dtrFound.Length
        Status("Processing " & intPtc & "%", intPtc)
        ' Treat this entry
        With dtrFound(intI)
          If (Me.chbAllDelete.Checked) Then
            .Item("h") = "delete"
          Else
            .Item("h") = "add"
          End If
        End With
      Next intI
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmLemmaVern/chbAllDelete error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub


  Private Sub tbLemma_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbLemma.TextChanged
    CtlChanged(objLVed, Me.tbLemma, bInit)
    ' Make sure [Action] is set to "change"
    If (bInit) AndAlso (Me.cboAction.SelectedItem <> "change") Then Me.cboAction.SelectedItem = "change"
  End Sub
  Private Sub tbPosHead_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbPosHead.TextChanged
    CtlChanged(objLVed, Me.tbPosHead, bInit)
  End Sub
  Private Sub tbVern_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbVern.TextChanged
    CtlChanged(objLVed, Me.tbVern, bInit)
  End Sub
  Private Sub tbPos_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbPos.TextChanged
    CtlChanged(objLVed, Me.tbPos, bInit)
  End Sub
  Private Sub cboAction_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAction.SelectedIndexChanged
    CtlChanged(objLVed, Me.cboAction, bInit)
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.DialogResult = Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    ' Double check if this is really what the user wants
    Select Case MsgBox("Do you really want to close this window?", MsgBoxStyle.YesNoCancel)
      Case MsgBoxResult.No, MsgBoxResult.Cancel
        Exit Sub
    End Select
    Me.DialogResult = Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  Private Sub dgvLemmaVern_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvLemmaVern.KeyDown
    Try
      ' Validate
      If (bInit) Then
        Select Case e.KeyCode
          Case Keys.D
            ' Make sure this entry is marked as "delete"
            If (Me.cboAction.SelectedItem = "delete") Then
              Me.cboAction.SelectedItem = "add"
            Else
              Me.cboAction.SelectedItem = "delete"
            End If
            ' Me.cboAction.SelectedItem = "delete"
        End Select

      End If
    Catch ex As Exception

    End Try
  End Sub


End Class