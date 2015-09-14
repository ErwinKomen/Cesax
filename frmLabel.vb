Imports System.Windows.Forms

Public Class frmLabel
  Private loc_strLabelText As String = ""
  Private bProcessing As Boolean = False
  Private frmThis As frmMain = Nothing      ' What we relate to
  Public Property LabelText()
    Get
      Return loc_strLabelText
    End Get
    Set(ByVal value)
      ' Show we are busy
      bProcessing = True
      loc_strLabelText = value
      ' Put this in the textbox
      Me.tbLabel.Text = loc_strLabelText
      ' Select the current label completely
      Me.tbLabel.Select(0, Me.tbLabel.TextLength)
      ' Set the focus to it
      Me.tbLabel.Focus()
      ' We are no longer busy
      bProcessing = False
    End Set
  End Property
  Public Sub SetLocation(ByRef pntThis As Point)
    Try
      Me.Top = frmThis.Top + frmThis.MenuStrip1.Height + pntThis.Y
      Me.Left = frmThis.Left + frmThis.litTree.Left + pntThis.X
    Catch ex As Exception

    End Try
  End Sub
  Public Sub SetForm(ByRef frmOwner As frmMain)
    frmThis = frmOwner
  End Sub
  Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  Private Sub tbLabel_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbLabel.TextChanged
    ' Are we busy?
    If (bProcessing) Then Exit Sub
    ' Change the local copy
    loc_strLabelText = Trim(Me.tbLabel.Text)
  End Sub

  Private Sub frmLabel_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    ' show I am owned
    Me.Owner = frmMain
    '' Adapt the location according to the width and height
    'Me.Top = Me.Top + Me.Height - Me.tbLabel.Height
    'Me.Left = Me.Left + Me.Width - Me.tbLabel.Width
    ' Adapt the size
    Me.Width = Me.tbLabel.Width
    Me.Height = Me.tbLabel.Height
  End Sub
End Class
