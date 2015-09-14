Imports System.Windows.Forms

Public Class frmCategory
  ' ================================= LOCAL VARIABLES =========================================================
  Private loc_strCategory As String ' The category that has been selected
  Private loc_bCheck As Boolean     ' Whether to check for differences or just fill in the unknowns 
  Private loc_bOptions As Boolean = False  ' Whether to show the options or not
  Private loc_strFeatFile As String ' Feature definition file
  '  and those that were not done yet
  ' ============================================================================================================
  Public ReadOnly Property Category() As String
    Get
      ' Return the current category
      Return loc_strCategory
    End Get
  End Property
  Public ReadOnly Property Check() As Boolean
    Get
      ' Return whether differences should be checked or not
      Return loc_bCheck
    End Get
  End Property
  Public ReadOnly Property FeatFile() As String
    Get
      Return loc_strFeatFile
    End Get
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   Options
  ' Goal:   Either show or hide the options group
  ' History:
  ' 07-03-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Public WriteOnly Property ShowOptions() As Boolean
    Set(ByVal bShow As Boolean)
      Me.grpOptions.Visible = bShow
      loc_bOptions = bShow
    End Set
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   OK_Button_Click
  ' Goal:   Return positively
  ' History:
  ' 16-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Cancel_Button_Click
  ' Goal:   Return positively
  ' History:
  ' 16-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   frmCategory_Load
  ' Goal:   Set the timer
  ' History:
  ' 16-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmCategory_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Set timer
    Me.Timer1.Enabled = True
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Initialize
  ' History:
  ' 16-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Try
      ' Reset the timer
      Me.Timer1.Enabled = False
      ' Initially HIDE the group with options
      Me.grpOptions.Visible = loc_bOptions
      ' Set the correct INITIAL category
      Me.rbNP.Select()
    Catch ex As Exception
      ' Show error
      HandleErr("frmCategory/Timer1_Tick error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   rbNP_CheckedChanged
  ' Goal:   Set the correct category
  ' History:
  ' 16-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub rbNP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNP.CheckedChanged
    loc_strCategory = CAT_NP
  End Sub
  Private Sub rbAdv_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbAdv.CheckedChanged
    loc_strCategory = CAT_ADV
  End Sub
  Private Sub rbVb_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbVb.CheckedChanged
    loc_strCategory = CAT_VB
  End Sub
  Private Sub rbVbUnacc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbVbUnacc.CheckedChanged
    loc_strCategory = CAT_VBUNACC
  End Sub
  Private Sub tbVbType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbVbType.CheckedChanged
    loc_strCategory = CAT_VBTYPE
  End Sub
  Private Sub rbGr_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbGr.CheckedChanged
    loc_strCategory = CAT_GR
  End Sub
  Private Sub rbWh_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbWh.CheckedChanged
    loc_strCategory = CAT_WH
  End Sub
  Private Sub rbAnim_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbAnim.CheckedChanged
    loc_strCategory = CAT_ANIM
  End Sub
  Private Sub chbCheck_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbCheck.CheckedChanged
    ' Adapt the local value
    loc_bCheck = (Me.chbCheck.Checked)
  End Sub
  Private Sub rbCogn_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbCogn.CheckedChanged
    loc_strCategory = CAT_COGN
  End Sub
  Private Sub rbPrepNorm_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbPrepNorm.CheckedChanged
    loc_strCategory = CAT_PNORM
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   cmdFeatDefFile_Click
  ' Goal:   Select a feature definition file
  ' History:
  ' 15-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdFeatDefFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFeatDefFile.Click
    Dim dlgThis As New OpenFileDialog

    Try
      ' Ask for a file
      With dlgThis
        ' Start in the working directory
        .InitialDirectory = strWorkDir
        ' probably text files
        .Filter = "text file (*.txt)|*.txt|other type (*.*)|*.*"
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.Yes, Windows.Forms.DialogResult.OK
            ' Get the file
            loc_strFeatFile = .FileName
            ' Put it in the textbox
            Me.tbFeatDefFile.Text = loc_strFeatFile
          Case Else
            ' Exit
            Exit Sub
        End Select
      End With
    Catch ex As Exception
      ' Show error
      HandleErr("frmCategory/FeatDefFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

End Class
