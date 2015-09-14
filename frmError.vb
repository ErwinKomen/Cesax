Public Class frmError
  Public WriteOnly Property ErrText() As String
    Set(ByVal value As String)
      Me.tbErrText.Text = value
    End Set
  End Property
  Private Sub cmdContinue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdContinue.Click
    ' Set the dialog result
    Me.DialogResult = Windows.Forms.DialogResult.Ignore
    ' Close myself
    Me.Close()
  End Sub

  Private Sub cmdInterrupt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInterrupt.Click
    ' Set the dialog result
    Me.DialogResult = Windows.Forms.DialogResult.Cancel
    ' Close myself
    Me.Close()
  End Sub

  Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
    ' Set the dialog result
    Me.DialogResult = Windows.Forms.DialogResult.OK
    ' Close myself
    Me.Close()
  End Sub

  Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
    ' Set the dialog result
    Me.DialogResult = Windows.Forms.DialogResult.Abort
    ' Close myself
    Me.Close()
  End Sub
End Class