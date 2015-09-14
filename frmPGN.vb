Imports System.Windows.Forms

Public Class frmPGN
  ' ================================= LOCAL VARIABLES =========================================================
  Private loc_strPgn As String = ""     ' Local copy of PGN
  Private loc_strNPtype As String = ""  ' Local copy of NPtype (Dem or Pro)
  ' ================================= LOCAL CONSTANTS =========================================================
  ' ===========================================================================================================
  ' ------------------------------------------------------------------------------------
  ' Name:   PGN
  ' Goal:   Get or set the PGN property
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Property PGN() As String
    Get
      ' Return the local copy
      PGN = loc_strPgn
    End Get
    Set(ByVal value As String)
      Dim arPgn() As String   ' Array of ambiguous values
      Dim intI As Integer     ' COunter

      ' Split the value into an array
      arPgn = Split(value, ";")
      ' Clear the listbox
      With Me.lboPGN
        .Items.Clear()
        ' Load the values
        For intI = 0 To UBound(arPgn)
          .Items.Add(arPgn(intI))
        Next intI
        ' Set the selected value
        .SelectedIndex = 0
      End With
    End Set
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   Clause
  ' Goal:   Set the HTML text of the [wbClause]
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public WriteOnly Property Clause() As String
    Set(ByVal value As String)
      Me.wbClause.DocumentText = value
    End Set
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   NPtype
  ' Goal:   Set the NP type (Dem or Pro)
  ' History:
  ' 10-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public ReadOnly Property NPtype() As String
    Get
      ' Return the local topy
      NPtype = loc_strNPtype
    End Get
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   OK_Button_Click
  ' Goal:   Return a positive dialogue result
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Cancel_Button_Click
  ' Goal:   Return negative dialogue result
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   frmPGN_Load
  ' Goal:   Trigger initialisation
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmPGN_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Set owner
    Me.Owner = frmMain
    Me.StartPosition = FormStartPosition.CenterScreen
    ' Fill the combobox
    With Me.cboNPtype
      .Items.Clear()
      .Items.Add("Pro")
      .Items.Add("Dem")
    End With
    ' Make sure the initialisation gets triggered
    Me.Timer1.Enabled = True
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Perform initialisation
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Try
      ' Switch off the timer
      Me.Timer1.Enabled = False
      ' Show the first element of the listbox
      Me.lboPGN.SelectedIndex = 0
      ' Show the first element of the combobox
      Me.cboNPtype.SelectedIndex = 0
      ' Show we are ready
      Status("ready")
      lboPGN_SelectedIndexChanged(sender, e)
      ' Set the selection right
      With Me.tbPGN
        .SelectionStart = 0
        .SelectionLength = .TextLength
      End With
      ' Set the focus on the correct element
      Me.tbPGN.Focus()
    Catch ex As Exception
      ' Show error
      HandleErr("frmPGN/Timer1 (initialisation) error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   lboPGN_DoubleClick
  ' Goal:   Make sure we end with a positive note
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub lboPGN_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lboPGN.DoubleClick
    ' Show the correct value
    Me.tbPGN.Text = Me.lboPGN.SelectedItem
    ' Set the dialog result
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    ' Leave this form
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   lboPGN_SelectedIndexChanged
  ' Goal:   Adapt what is displayed
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub lboPGN_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lboPGN.SelectedIndexChanged
    Try
      ' Validate
      If (Not Me.Visible) Then Exit Sub
      ' Show the correct value
      Me.tbPGN.Text = Me.lboPGN.SelectedItem
    Catch ex As Exception
      ' Show error
      HandleErr("frmPGN/PGN_SelectedIndexChanged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   tbPGN_TextChanged
  ' Goal:   Adapt the local copy
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub tbPGN_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbPGN.TextChanged
    ' Set the local copy
    loc_strPgn = Me.tbPGN.Text
  End Sub

  Private Sub cboNPtype_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboNPtype.SelectedIndexChanged
    ' Set the local copy
    loc_strNPtype = Me.cboNPtype.SelectedItem
  End Sub
End Class
