Public Class frmConvert
  ' ============================= PRIVATE VARIABLES ===========================================
  Private strLang As String = ""  ' Selected language
  ' ===========================================================================================
  Public ReadOnly Property Language() As String
    Get
      ' Return the selected language
      Return strLang
    End Get
  End Property
  '---------------------------------------------------------------------------------------------------------
  ' Name:       frmConvert_Load()
  ' Goal:       Set main adjustments and trigger initialisation
  ' History:
  ' 15-11-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub frmConvert_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Set my owner
    Me.Owner = frmMain
    ' Start timer
    Me.Timer1.Enabled = True
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       Timer1_Tick()
  ' Goal:       Perform initialisation
  ' History:
  ' 15-11-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    ' Switch off timer
    Me.Timer1.Enabled = False
    ' Set default language
    Me.rbEnglishTranscribed.Checked = True
    ' Set status
    Status("Make a selection")
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       rbEnglish_CheckedChanged()
  ' Goal:       Set the indicated language
  ' History:
  ' 15-11-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub rbEnglishWritten_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbEnglishWritten.CheckedChanged
    strLang = "English" : Status(strLang)
  End Sub
  Private Sub rbEnglishTranscribed_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbEnglishTranscribed.CheckedChanged
    strLang = "EnglishTranscribed" : Status(strLang)
  End Sub
  Private Sub rbGerman_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbGerman.CheckedChanged
    strLang = "German" : Status(strLang)
  End Sub
  Private Sub rbDutch_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbDutch.CheckedChanged
    strLang = "Dutch" : Status(strLang)
  End Sub
  Private Sub rbSpanish_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSpanish.CheckedChanged
    strLang = "Spanish" : Status(strLang)
  End Sub
  Private Sub rbFrench_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbFrench.CheckedChanged
    strLang = "French" : Status(strLang)
  End Sub
  Private Sub rbWelsh_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbWelsh.CheckedChanged
    strLang = "Welsh" : Status(strLang)
  End Sub
  Private Sub rbLezgi_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbLezgi.CheckedChanged
    strLang = "Lezgi" : Status(strLang)
  End Sub
  Private Sub rbLak_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbLak.CheckedChanged
    strLang = "Lak" : Status(strLang)
  End Sub
  Private Sub rbChechen_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbChechen.CheckedChanged
    strLang = "Chechen" : Status(strLang)
  End Sub
  Private Sub rbOther_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbOther.CheckedChanged
    strLang = Me.tbEthno.Text : Status(strLang)
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       tbEthno_TextChanged()
  ' Goal:       Set the indicated language
  ' History:
  ' 15-11-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub tbEthno_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbEthno.TextChanged
    Try
      ' Take the language code from the [tbEthno]
      strLang = Me.tbEthno.Text : Status(strLang)
    Catch ex As Exception
      ' Show error
      HandleErr("frmConvert/tbEthno error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------------------------------------------------
  ' Name:       cmdOk_Click()
  ' Goal:       Return with the correct status
  ' History:
  ' 15-11-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    Me.DialogResult = Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub
  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.DialogResult = Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub


End Class