Public Class frmMorphCheck
  ' =========================================================================================================================
  Private loc_strSentence As String = ""  ' Local copy of sentence (in html)
  Private loc_strOption As String = ""    ' The option that is being chosen
  Private loc_arOption() As String        ' Array of options
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
  Public WriteOnly Property POS() As String
    Set(ByVal value As String)
      Me.tbPos.Text = value
    End Set
  End Property
  Public WriteOnly Property Features() As String
    Set(ByVal value As String)
      Me.tbFeat.Text = value
    End Set
  End Property
  Public WriteOnly Property Lemma() As String
    Set(ByVal value As String)
      Me.tbLemma.Text = value
    End Set
  End Property
  Public WriteOnly Property LemmaDictCat() As String
    Set(ByVal value As String)
      Me.tbMWcat.Text = value
    End Set
  End Property
  Public WriteOnly Property LemmaDictPOS() As String
    Set(ByVal value As String)
      Me.tbCheck.Text = value
    End Set
  End Property
  Public WriteOnly Property MorphDict() As String()
    Set(ByVal value As String())
      Dim intI As Integer   ' Counter

      ' Process this one!!
      With Me.lboMorphDict
        .Items.Clear()
        For intI = 0 To value.Length - 1
          .Items.Add(value(intI))
        Next intI
      End With
    End Set
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   Choice
  ' Goal:   Give the ID of the choice made by the user
  ' History:
  ' 19-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public ReadOnly Property Choice() As String
    Get
      Return loc_strOption
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

  Private Sub rbAllowPOS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbAllowPOS.CheckedChanged
    loc_strOption = "AllowPOS"
  End Sub
  Private Sub rbChangeToLemmaDictPOS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbChangeToLemmaDictPOS.CheckedChanged
    loc_strOption = "ChangeToLemmaDictPOS"
  End Sub
  Private Sub rbDeleteLemmaVernPOS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbDeleteLemmaVernPOS.CheckedChanged
    loc_strOption = "DeleteLemmaVernPOS"
  End Sub
  Private Sub rbDeleteLemmaVern_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbDeleteLemmaVern.CheckedChanged
    loc_strOption = "DeleteLemmaVern"
  End Sub

  Private Sub lboMorphDict_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lboMorphDict.SelectedIndexChanged
    Dim strLine As String   ' One line

    Try
      ' Validate
      If (Me.lboMorphDict.Items.Count = 0) Then Exit Sub
      ' Get the selected ID
      strLine = Me.lboMorphDict.SelectedItem
      strLine = Microsoft.VisualBasic.Strings.Left(strLine, InStr(strLine, vbTab) - 1)
      loc_strOption = strLine
    Catch ex As Exception
      ' Show error
      HandleErr("frmMorph/MorphDictIndex error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Class