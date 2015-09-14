Public Class frmEnglish
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
  ' Name:       frmEnglish_Load()
  ' Goal:       Set main adjustments and trigger initialisation
  ' History:
  ' 15-11-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub frmEnglish_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
    Me.rbOldEnglish.Checked = True
    Me.Owner = frmMain
    ' Set status
    Status("Make a selection")
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       rbOldEnglish_CheckedChanged()
  ' Goal:       Set the indicated language
  ' History:
  ' 15-11-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub rbOldEnglish_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbOldEnglish.CheckedChanged
    strLang = "OE"
    Status("Old English (OE)")
  End Sub
  Private Sub rbOldEnglishBT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbOldEnglishBT.CheckedChanged
    strLang = "OEB"
    Status("Old English Bossworth & Toller (OEB)")
  End Sub
  Private Sub rbMiddleEnglish_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbMiddleEnglish.CheckedChanged
    strLang = "ME"
    Status("Middle English (ME)")
  End Sub
  Private Sub rbEarlyModEng_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbEarlyModEng.CheckedChanged
    strLang = "eModE"
    Status("Early Modern English (eModE)")
  End Sub
  Private Sub rbLateModEng_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbLateModEng.CheckedChanged
    strLang = "LmodE"
    Status("Late Modern English (LmodE)")
  End Sub
  Private Sub rbMiddleEnglishMED_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbMiddleEnglishMED.CheckedChanged
    strLang = "MED"
    Status("Middle English - MED")
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