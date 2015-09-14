Public Class frmErrWait
  Private bErrorViewed As Boolean = False  ' Whether the error message has been read
  Public Property ErrorViewed() As Boolean
    Get
      ErrorViewed = bErrorViewed
    End Get
    Set(ByVal value As Boolean)
      bErrorViewed = value
    End Set
  End Property
  Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    ' Indicate the error message has been read
    bErrorViewed = True
  End Sub

  Private Sub frmErrWait_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Set the owner
    Me.Owner = frmMain
  End Sub
End Class