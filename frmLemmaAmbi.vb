Public Class frmLemmaAmbi
  Private arLemma() As String   ' Array of lemma's
  Private loc_intSel As Integer ' Selected lemma
  Private loc_strFeat As String ' Features
  Private bInit As Boolean = False  ' Initialisation flag
  ' ==========================================================================================================================
  Public WriteOnly Property Vern() As String
    Set(ByVal value As String)
      Me.tbVern.Text = value
    End Set
  End Property
  Public WriteOnly Property Pos() As String
    Set(ByVal value As String)
      Me.tbPos.Text = value
    End Set
  End Property
  Public Property Feat() As String
    Get
      Return loc_strFeat
    End Get
    Set(ByVal value As String)
      loc_strFeat = value
      Me.tbFeat.Text = loc_strFeat
    End Set
  End Property
  Public WriteOnly Property Lemma() As String
    Set(ByVal value As String)
      Dim intI As Integer ' Counter

      ' Convert lemma into array
      arLemma = Split(value, ";")
      With Me.lboLemma
        .Items.Clear()
        For intI = 1 To arLemma.Length - 1
          ' Add this item
          .Items.Add(arLemma(intI).Replace("|", vbTab))
        Next intI
        ' Set to default position
        .SelectedIndex = 0 : loc_intSel = 0
      End With
    End Set
  End Property
  Public ReadOnly Property LemmaChosen() As String
    Get
      Dim arSel() As String   ' Selected entry

      ' Return the lemma we have chosen
      If (loc_intSel < 0) Then
        Return ""
      Else
        arSel = Split(lboLemma.Items(loc_intSel), vbTab)
        Return arSel(0)
      End If
    End Get
  End Property

  Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    Me.DialogResult = Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub
  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.DialogResult = Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub
  Private Sub cmdSkip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSkip.Click
    Me.DialogResult = Windows.Forms.DialogResult.Ignore
    Me.Close()
  End Sub

  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Try
      ' Reset timer
      Me.Timer1.Enabled = False
      ' Perform initialisation
      Me.Owner = frmMain
      Me.StartPosition = FormStartPosition.CenterParent
      loc_intSel = -1
      Application.DoEvents()
      bInit = True
      Status("ready")
    Catch ex As Exception

    End Try
  End Sub

  Private Sub frmLemmaAmbi_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Trigger initialisation timer
    Me.Timer1.Enabled = True
  End Sub

  Private Sub lboLemma_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lboLemma.DoubleClick
    Me.DialogResult = Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  Private Sub lboLemma_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lboLemma.SelectedIndexChanged
    Try
      ' Validate
      If (Not bInit) Then Exit Sub
      ' Make sure we have the right index
      loc_intSel = Me.lboLemma.SelectedIndex
      ' See if we can show a dictionary entry here
      Me.tbDict.Text = Split(lboLemma.Items(loc_intSel), vbTab)(0)
      ' Show dictionary entry
      DictShow()
    Catch ex As Exception

    End Try
  End Sub

  Private Sub DictShow(Optional ByVal bPartly As Boolean = False)
    Dim dtrFound() As DataRow     ' dictionary entry
    Dim dtrSense() As DataRow     ' Senses
    Dim colHtml As New StringColl ' HTML text
    Dim intEntryId As Integer     ' ID of this entry
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter

    Try
      ' Validate
      If (Not bInit) Then Exit Sub
      ' See if we can show a dictionary entry here
      colHtml.Add("<html><body>")
      If (bPartly) Then
        dtrFound = tdlOEdict.Tables("Entry").Select("l LIKE '" & Me.tbDict.Text & "*'")
      Else
        ' Need an accect match
        dtrFound = tdlOEdict.Tables("Entry").Select("l='" & Me.tbDict.Text & "'")
      End If
      For intI = 0 To dtrFound.Length - 1
        With dtrFound(intI)
          ' Start off the entry
          colHtml.Add("<b>" & intI + 1 & ": <font color='red'>" & .Item("l") & "</font></b>")
          If (.Item("Cf").ToString <> "") Then
            ' This just points to something else
            colHtml.Add("<i>see: </i>" & .Item("Cf"))
          Else
            ' Check POS
            If (.Item("Pos").ToString <> "") Then
              colHtml.Add("pos=" & .Item("Pos"))
            End If
            ' Add a break
            colHtml.Add("<br>")
            ' Get all the senses
            ' Get the entry id
            intEntryId = .Item("EntryId")
            If (intEntryId < 0) Then Exit Sub
            ' Get all definitions
            dtrSense = tdlOEdict.Tables("Sense").Select("EntryId=" & intEntryId, "N ASC")
            For intJ = 0 To dtrSense.Length - 1
              ' Add this sense + definition
              With dtrSense(intJ)
                colHtml.Add(intI + 1 & ":" & intJ + 1 & ". " & .Item("Def") & "<br>")
              End With
            Next intJ
          End If
          ' Add paragraph mark
          colHtml.Add("<p>")
        End With
      Next intI
      ' Finish HTML
      colHtml.Add("</body></html>")
      ' Show it
      Me.wbDict.DocumentText = colHtml.Text
      Application.DoEvents()
    Catch ex As Exception

    End Try
  End Sub
  Private Sub tbFeat_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbFeat.TextChanged
    loc_strFeat = Me.tbFeat.Text
  End Sub

  Private Sub tbDict_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDict.TextChanged
    ' Show dictionary entry
    DictShow(True)
  End Sub
End Class