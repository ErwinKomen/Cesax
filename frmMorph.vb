Public Class frmMorph
  ' =========================================================================================================================
  Private loc_strSentence As String = ""  ' Local copy of sentence (in html)
  Private loc_strOption As String = ""    ' The option that is being chosen
  Private loc_arOption() As String        ' Array of options
  Private loc_OptionId As Integer = -1    ' ID of the chosen option
  Private loc_strInfo As String           ' The information
  Private loc_strLemma As String
  Private loc_strLabel As String          ' My own label value
  Private loc_strParent As String         ' Parent constituent's label
  ' =========================================================================================================================
  ' ------------------------------------------------------------------------------------
  ' Name:   Sentence
  ' Goal:   Set the html sentence that is being displayed
  ' History:
  ' 19-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public WriteOnly Property Sentence() As String
    Set(ByVal value As String)
      loc_strSentence = value
      Me.wbSentence.DocumentText = loc_strSentence
    End Set
  End Property
  Public WriteOnly Property Choices() As String()
    Set(ByVal value As String())
    End Set
  End Property
  Public WriteOnly Property Information() As String
    Set(ByVal value As String)
      Me.tbInfo.Text = value
      loc_strInfo = value
    End Set
  End Property
  Public WriteOnly Property OwnLabel() As String
    Set(ByVal value As String)
      loc_strLabel = value
    End Set
  End Property
  Public WriteOnly Property ParentLabel() As String
    Set(ByVal value As String)
      loc_strParent = value
    End Set
  End Property
  Public ReadOnly Property Features() As String
    Get
      Return MyTrim(Me.tbFeat.Text).Replace(vbLf, ";")
    End Get
  End Property
  'Public WriteOnly Property POS() As String
  '  Set(ByVal value As String)
  '    Me.tbPos.Text = value
  '  End Set
  'End Property
  'Public WriteOnly Property Features() As String
  '  Set(ByVal value As String)
  '    Me.tbFeat.Text = value
  '  End Set
  'End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   SetChoices
  ' Goal:   Set the choices
  ' History:
  ' 19-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub SetChoices(ByRef dtrList() As DataRow)
    Dim strShow As String ' What we show in the listbox
    Dim intI As Integer   ' Counter

    Try
      ' Fill the listbox
      With Me.lboOptions
        .Items.Clear()
        For intI = 0 To dtrList.Length - 1
          ' Combine into a nice string
          With dtrList(intI)
            strShow = .Item("Vern") & " (" & .Item("l") & ")" & vbTab & .Item("Pos") & "/" & .Item("Label") & vbTab & _
                "[" & .Item("Freq") & "]" & vbTab & .Item("f")
          End With
          ' Show this string
          .Items.Add(strShow)
        Next intI
        ' Select number #0
        .SelectedIndex = 0
      End With
    Catch ex As Exception
      ' Show error
      HandleErr("frmMorph/SetChoices error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   Choice
  ' Goal:   Give the ID of the choice made by the user
  ' History:
  ' 19-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public ReadOnly Property Choice() As Integer
    Get
      Return loc_OptionId
    End Get
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   frmMorph_Load
  ' Goal:   Trigger initialisations
  ' History:
  ' 19-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmMorph_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Trigger timer
    Me.Timer1.Enabled = True
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Perform initialisations
  ' History:
  ' 19-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Try
      ' Switch off timer
      Me.Timer1.Enabled = False
      ' Do initialisations here...

    Catch ex As Exception
      ' Show error
      HandleErr("frmMorph/Timer1 error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   cmdOk_Click
  ' Goal:   Leave this form OK or Canceled
  ' History:
  ' 19-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    Me.DialogResult = Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub
  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.DialogResult = Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub
  Private Sub cmdNone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNone.Click
    Me.DialogResult = Windows.Forms.DialogResult.None
    Me.Close()
  End Sub
  Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
    Me.DialogResult = Windows.Forms.DialogResult.No
    Me.Close()
  End Sub
  Private Sub cmdFull_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFull.Click
    Me.DialogResult = Windows.Forms.DialogResult.Ignore ' Full generalization
    Me.Close()
  End Sub
  Private Sub cmdRestricted_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRestricted.Click
    Me.DialogResult = Windows.Forms.DialogResult.Retry  ' Restricted generalization
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   lboOptions_DoubleClick
  ' Goal:   Select the option chosen by the user and return positively
  ' History:
  ' 19-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub lboOptions_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lboOptions.DoubleClick
    Try
      Me.DialogResult = Windows.Forms.DialogResult.OK
      Me.Close()
    Catch ex As Exception
      ' Show error
      HandleErr("frmMorph/OptionsDblClick error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   lboOptions_SelectedIndexChanged
  ' Goal:   Select the option chosen by the user
  ' History:
  ' 19-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub lboOptions_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lboOptions.SelectedIndexChanged
    Dim intPos As Integer       ' Where we have features
    Dim strFeat As String = ""  ' Features
    Dim arSel() As String       ' Selected

    Try
      ' Validate
      If (Me.lboOptions.Items.Count = 0) Then Exit Sub
      ' Change the chosen option
      loc_strOption = Me.lboOptions.SelectedItem
      loc_OptionId = Me.lboOptions.SelectedIndex
      ' Get array of selection
      arSel = Split(loc_strOption, vbTab)
      ' Get features: after the 4th tab
      strFeat = arSel(arSel.Length - 1).Replace(";", vbCrLf)
      ' Set the features
      Me.tbFeat.Text = strFeat
      ' Get the lemma
      intPos = InStr(loc_strOption, "(")
      loc_strLemma = Mid(loc_strOption, intPos + 1)
      intPos = InStr(loc_strLemma, ")")
      loc_strLemma = Microsoft.VisualBasic.Strings.Left(loc_strLemma, intPos - 1)
      ' Set generalization texts
      Me.tbFull.Text = loc_strLabel & "/" & loc_strParent & " +lemma +features"
      Me.tbRestricted.Text = loc_strLabel & "/" & loc_strParent & " +lemma (without features)"
    Catch ex As Exception
      ' Show error
      HandleErr("frmMorph/SelectionChanged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Class