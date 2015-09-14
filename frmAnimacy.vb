Public Class frmAnimacy
  ' ==================================== LOCAL VARIABLES ====================================================
  Private strAnimacy As String = "-" ' Value of the animacy
  ' =========================================================================================================
  ' ------------------------------------------------------------------------------------
  ' Name:   Animacy
  ' Goal:   Return the Animacy value selected
  ' History:
  ' 16-04-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Public ReadOnly Property Animacy() As String
    Get
      Return strAnimacy
    End Get
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   Chain
  ' Goal:   HTML code of the chain
  ' History:
  ' 16-04-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Public WriteOnly Property Chain() As String
    Set(ByVal value As String)
      ' Fill the Chain HTML
      Me.wbChain.DocumentText = value
    End Set
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   Percentage
  ' Goal:   Indicate where we are in terms of percentage
  ' History:
  ' 16-04-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Public WriteOnly Property Percentage() As Integer
    Set(ByVal value As Integer)
      Status("Progress " & value, value)
    End Set
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   frmAnimacy_Load
  ' Goal:   Figure out the animacy of this one chain
  ' History:
  ' 16-04-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmAnimacy_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Set the owner
    Me.Owner = frmMain
    ' Start timer
    Me.Timer1.Enabled = True
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Initialisations
  ' History:
  ' 16-04-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Try
      ' Switch off timer
      Me.Timer1.Enabled = False
      ' Set radiobuttons to initial value
      Me.rbUnknown.Checked = True
      ' Show we are ready
      Status("Ready")
    Catch ex As Exception
      ' Show error
      HandleErr("frmAnimacy/Timer1 error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  Private Sub rbAnim_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbAnim.CheckedChanged
    strAnimacy = "a"
  End Sub
  Private Sub rbInanim_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbInanim.CheckedChanged
    strAnimacy = "i"
  End Sub
  Private Sub rbUnknown_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbUnknown.CheckedChanged
    strAnimacy = "-"
  End Sub

  Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    Me.DialogResult = Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    ' Make sure interrupt is set
    bInterrupt = True
    Me.DialogResult = Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub
End Class