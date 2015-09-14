Public Class TabArrow
  Inherits System.Windows.Forms.TabControl
  Protected Overrides Function IsInputKey( _
    ByVal keyData As System.Windows.Forms.Keys) As Boolean

    Try
      Select Case keyData
        Case Keys.Down, Keys.Up, Keys.Left, Keys.Right, Keys.Escape
          ' Check who is selected
          If (Me.SelectedTab.Name = "tpTree") Then
            ' Allow my own routine to handle the input
            Return True
          End If
      End Select
      ' Make sure standard input is handled
      Return MyBase.IsInputKey(keyData)
    Catch ex As Exception
      ' Show error
      HandleErr("tabArrow/IsInputKey error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)
    Try
      ' Check who this is 
      If (Me.SelectedTab.Name = "tpTree") Then
        ' Raise a KeyDown event for the [litTree] control
        frmMain.litTreeKey(e)
      Else
        ' Make sure standard input is handled
        MyBase.OnKeyDown(e)
      End If
    Catch ex As Exception

    End Try
  End Sub
End Class
