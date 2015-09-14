Public Class frmChain
  Private strDir As String = strWorkDir ' Directory chosen
  ' ------------------------------------------------------------------------------------
  ' Name:   Action
  ' Goal:   The action to perform
  ' History:
  ' 16-08-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Public ReadOnly Property Action() As String
    Get
      Return IIf(Me.rbRecalc.Checked, "Recalc", "Update")
    End Get
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   Action
  ' Goal:   The scope of the action
  ' History:
  ' 16-08-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Public ReadOnly Property Scope() As String
    Get
      Return IIf(Me.rbAll.Checked, "All", "Current")
    End Get
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   Dir
  ' Goal:   Directory where files are taken from
  ' History:
  ' 16-08-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Public ReadOnly Property Dir() As String
    Get
      ' Return the chosen directory
      Return strDir
    End Get
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   frmChain_Load
  ' Goal:   Load the form
  ' History:
  ' 16-08-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmChain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Me.Owner = frmMain
    ' Trigger timer
    Me.Timer1.Enabled = True
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Initialisations
  ' History:
  ' 16-08-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Try
      ' Switch off the timer
      Me.Timer1.Enabled = False
      ' Do initialisations
      Me.rbCurrent.Select()
      Me.rbRecalc.Select()
      Me.tbDir.Text = strDir
      Status("Ready")
    Catch ex As Exception
      ' Show error
      HandleErr("frmChain/Timer1 error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   cmdOk_Click
  ' Goal:   Exit positively
  ' History:
  ' 16-08-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    ' Set the dialog result
    Me.DialogResult = Windows.Forms.DialogResult.OK
    ' Close the form
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   cmdCancel_Click
  ' Goal:   Exit negatively
  ' History:
  ' 16-08-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    ' Set the dialog result
    Me.DialogResult = Windows.Forms.DialogResult.Cancel
    ' Close the form
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   cmdDir_Click
  ' Goal:   Choose a directory
  ' History:
  ' 16-08-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDir.Click
    If (Not GetDirName(Me.FolderBrowserDialog1, strDir, "Directory to search for PSDX files")) Then
      ' Show that there is something wrong
      Status("frmChain: there is a problem finding a directory")
    End If
    ' Otherwise set the name
    Me.tbDir.Text = strDir
  End Sub
End Class