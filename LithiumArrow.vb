Public Class LithiumArrow
  Inherits LithiumControl
  Protected Overrides Function IsInputKey( _
      ByVal keyData As System.Windows.Forms.Keys) As Boolean

    Select Case keyData
      Case Keys.Down, Keys.Up, Keys.Left, Keys.Right, Keys.Escape
        ' Allow my own routine to handle the input
        Return True
      Case Else
        ' Make sure standard input is handled
        Return MyBase.IsInputKey(keyData)
    End Select
  End Function

  Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)
    MyBase.OnKeyDown(e)
  End Sub
  Protected Overrides Sub OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs)
    MyBase.OnKeyPress(e)
  End Sub
End Class
