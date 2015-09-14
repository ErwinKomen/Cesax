Public Class frmLanguage
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
  ' Name:       frmLanguage_Load()
  ' Goal:       Set main adjustments and trigger initialisation
  ' History:
  ' 15-11-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub frmLanguage_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
    Me.rbEnglish.Checked = True
    Me.Owner = frmMain
    ' Set status
    Status("Make a selection")
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       rbEnglish_CheckedChanged()
  ' Goal:       Set the indicated language
  ' History:
  ' 15-11-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub rbEnglish_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbEnglish.CheckedChanged
    strLang = "English"
    Status("English")
  End Sub
  Private Sub rbDutch_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbDutch.CheckedChanged
    strLang = "Dutch"
    Status("Dutch (CGN)")
  End Sub
  Private Sub rbChechen_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbChechen.CheckedChanged
    strLang = "Chechen"
    Status("Chechen")
  End Sub
  Private Sub rbWelsh_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbWelsh.CheckedChanged
    strLang = "Welsh"
    Status("Welsh")
  End Sub
  Private Sub rbGerman_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbGerman.CheckedChanged
    strLang = "German"
    Status("German")
  End Sub
  Private Sub rbSelf_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSelf.CheckedChanged
    strLang = "Self"
    Status("Self")
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