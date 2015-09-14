Imports System.Windows.Forms
Imports System.Xml
Public Class frmTagSync
  ' ======================================== PRIVATE VARIABLES ============================================================
  Dim loc_OEtagOffset As Integer = 0  ' How many words do we have to go left (minus) or right (plus) in the OEtag text?
  Dim loc_YcoeOffset As Integer = 0   ' How many words do we have to go left (minus) or right (plus) in the YCOE  text?
  Dim loc_intCurrent As Integer       ' Current offset
  Dim loc_intStart As Integer         ' Current start in [arTagged] array
  Dim loc_intLen As Integer           ' Current length
  Dim loc_arTagged() As String        ' Local copy of the array
  Dim loc_ndxList As XmlNodeList      ' Local copy of list
  Dim loc_strLoc As String = ""       ' Location
  Dim loc_strOEfile As String = ""    ' Name of OE file
  Dim loc_intForestId As Integer = -1 ' @forestId of the line found with cmdYcoeSearch
  ' =======================================================================================================================
  Public ReadOnly Property OEtagOffset() As Integer
    Get
      Return loc_OEtagOffset
    End Get
  End Property
  Public ReadOnly Property YcoeOffset() As Integer
    Get
      Return loc_YcoeOffset
    End Get
  End Property
  Public ReadOnly Property ForestId() As Integer
    Get
      Return loc_intForestId
    End Get
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   cmdYes_Click, cmdNo, cmdCancel
  ' Goal:   Signal the result of the operation
  ' History:
  ' 04-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdYes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdYes.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.Yes
    Me.Close()
  End Sub
  Private Sub cmdNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNo.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.No
    Me.Close()
  End Sub
  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   frmTagSync_Load, cmdNo, cmdCancel
  ' Goal:   Trigger initialisation
  ' History:
  ' 04-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmTagSync_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Set my parent
    Me.Owner = frmMain
    ' Start initialisation timer
    Me.Timer1.Enabled = True
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   frmTagSync_VisibleChanged
  ' Goal:   Show where we are when we become visible
  ' History:
  ' 07-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmTagSync_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
    Try
      ' Are we becoming visible?
      If (Me.Visible) Then
        ' Show the location
        Status(loc_strLoc)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmTagSync/VisibleChanged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Perform initialisation (rudimentary)
  ' History:
  ' 04-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    ' Switch off timer
    Me.Timer1.Enabled = False
    ' Perform initialisations
    loc_OEtagOffset = 0
    loc_YcoeOffset = 0
    '' Show we are ready
    'Status("Ready")
    'Me.wbOEtagged.DocumentText = "(empty)"
    'Me.wbYCOE.DocumentText = "(empty)"
  End Sub
  Public WriteOnly Property LocShow() As String
    Set(ByVal value As String)
      loc_strLoc = value
    End Set
  End Property
  Public WriteOnly Property OEfile() As String
    Set(ByVal value As String)
      loc_strOEfile = value
      Me.tbOEfile.Text = loc_strOEfile
    End Set
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   OEtagged
  ' Goal:   Set the contents of [wbOEtagged]
  ' History:
  ' 04-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub OEtagged(ByRef arTagged() As String, ByVal intStart As Integer, ByVal intLen As Integer, ByVal intThis As Integer)
    Try
      ' Store current offset
      loc_intStart = intStart : loc_intLen = intLen
      loc_arTagged = arTagged
      ' Perform actual action
      ShowOEtagged(arTagged, intStart, intLen, intThis)
    Catch ex As Exception
      ' Show error
      HandleErr("frmTagSync/OEtagged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Public Sub ShowOEtagged(ByRef arTagged() As String, ByVal intStart As Integer, ByVal intLen As Integer, ByVal intThis As Integer)
    Dim colHtml As New StringColl
    Dim intI As Integer     ' Counter
    Dim intLine As Integer  ' Position where a line number is
    Dim intPos As Integer   ' Position in [arTagged] array
    Dim strWord As String   ' One word
    Dim bHaveThis As Boolean = False  ' Did we encounter [intThis] or not?

    Try
      ' Validate
      If (arTagged Is Nothing) Then Exit Sub
      If (arTagged.Length = 0) Then Exit Sub
      If (intStart < 0) Then Exit Sub
      ' Check for end of file
      If (intStart >= arTagged.Length - 1) Then
        ' This is end of file reached
        Me.DialogResult = System.Windows.Forms.DialogResult.Abort
        Me.Close()
        Exit Sub
      End If
      ' Start html output
      colHtml.Add("<html><body>")
      ' Catch the starting position
      intPos = intStart
      ' Trace back to any possible preceding line number
      intLine = intPos
      While (intLine > 0) AndAlso (Not DoLike(arTagged(intLine), "_t(l*|_r(l*|_r(p*"))
        ' Go to the preceding one
        intLine -= 1
      End While
      ' Show whatever precedes until the preceding line number
      For intI = intLine To intPos - 1
        If (DoLike(arTagged(intI), "_t(l*|_r(l*|_r(p*")) Then
          ' This is a line indicator: show where we are!
          strWord = Mid(arTagged(intI), 5)
          colHtml.Add("(" & strWord & " ")
        Else
          ' Get the word
          strWord = TagToWord(arTagged(intI)) & " "
          ' Is this the target word?
          If (intI = intThis) Then
            ' Add red and bolded word
            colHtml.Add("<b><font color='red'>" & strWord & "</font></b>")
            bHaveThis = True
          Else
            ' Normal black word
            colHtml.Add(strWord)
          End If
        End If
        ' Are we skipping over [intThis]?
        If (Not bHaveThis) AndAlso (intI = intThis) Then
          Stop
        End If
      Next intI
      ' Loop through the words
      For intI = 0 To intLen - 1
        ' Skip empty places in the [arTagged] array
        While (intPos < arTagged.Length - 1) AndAlso (TagToWord(arTagged(intPos)) = "")
          ' Check what instruction this position is
          If (InStr(arTagged(intPos), "_r(bos") = 1) Then
            ' This is the start of a line
            colHtml.Add("<br>")
          ElseIf (DoLike(arTagged(intPos), "_t(l*|_r(l*|_r(p*")) Then
            ' This is a line indicator: show where we are!
            strWord = Mid(arTagged(intPos), 5)
            ' But we could be positioned right here...
            If (intPos = intThis) Then
              colHtml.Add("<br>(<b><font color='red'>" & strWord & " </font></b>")
            Else
              colHtml.Add("<br>(" & strWord & " ")
            End If
          End If
          '' Are we skipping over [intThis]?
          'If (Not bHaveThis) AndAlso (intPos = intThis) Then
          '  Stop
          'End If
          ' Go to the next position
          intPos += 1
        End While
        ' Skip and grey-out foreign words
        While (intPos < arTagged.Length - 1) AndAlso (InStr(arTagged(intPos), "_t(f") > 0 OrElse InStr(arTagged(intPos), "_f") > 0)
          ' Get the word
          strWord = TagToWord(arTagged(intPos)) & " "
          ' Is this the target word?
          If (intPos = intThis) Then
            ' Add red and bolded word
            colHtml.Add("<b><font color='blue'>" & strWord & "</font></b>")
            bHaveThis = True
          Else
            ' This is a foreign language, so grey it out, but do show it
            colHtml.Add("<font color='grey'>" & strWord & "</font>")
          End If
          ' Go to the next position
          intPos += 1
        End While
        ' Validate
        If (intPos < arTagged.Length) Then
          ' Get the word
          strWord = TagToWord(arTagged(intPos)) & " "
          ' Is this the target word?
          If (intPos = intThis) Then
            ' Add red and bolded word
            colHtml.Add("<b><font color='red'>" & strWord & "</font></b>")
            bHaveThis = True
          Else
            ' Normal black word
            colHtml.Add(strWord)
          End If
        End If
        ' Go to the next position
        intPos += 1
      Next intI
      ' At least try to show to the next _eos sign
      While (intPos < arTagged.Length - 2) AndAlso (InStr(arTagged(intPos), "_eos") = 0)
        ' Show this word
        ' Get the word
        strWord = TagToWord(arTagged(intPos)) & " "
        ' This is a foreign language, so grey it out, but do show it
        colHtml.Add(strWord)
        ' Go to next position
        intPos += 1
      End While
      ' Finish html output
      colHtml.Add("</body></html>")
      Me.wbOEtagged.DocumentText = colHtml.Text
    Catch ex As Exception
      ' Show error
      HandleErr("frmTagSync/ShowOEtagged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   YCOE
  ' Goal:   Set the contents of [wbYCOE]
  ' History:
  ' 04-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub YCOE(ByRef ndxList As XmlNodeList, ByVal intThis As Integer)

    Try
      ' Store current offset
      loc_intCurrent = intThis
      loc_ndxList = ndxList
      ' Perform the action
      ShowYcoe(ndxList, intThis)
    Catch ex As Exception
      ' Show error
      HandleErr("frmTagSync/YCOE error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   GetLoc
  ' Goal:   Get the current location
  ' History:
  ' 04-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetLoc() As String
    Try
      Return IO.Path.GetFileNameWithoutExtension(strCurrentFile) & ":" & _
                   loc_ndxList(loc_intCurrent).SelectSingleNode("./ancestor::forest[1]").Attributes("forestId").Value
    Catch ex As Exception
      ' Show error
      HandleErr("frmTagSync/YCOE error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ("error")
    End Try
  End Function
  Private Sub ShowYcoe(ByRef ndxList As XmlNodeList, ByVal intThis As Integer)
    Dim ndxLine As XmlNodeList    ' One line
    Dim colHtml As New StringColl
    Dim intI As Integer           ' Counter
    Dim intCount As Integer = 10  ' Number of lines to get

    Try
      ' Validate
      If (ndxList Is Nothing) Then Exit Sub
      If (ndxList.Count = 0) Then Exit Sub
      ' Get location
      loc_strLoc = GetLoc()
      ' Show status
      Application.DoEvents()
      Status(loc_strLoc)
      ' Start html output
      colHtml.Add("<html><body>")
      ' Get the preceding <forest>, if existing
      ndxLine = ndxList(intThis).SelectNodes("./ancestor::forest/preceding-sibling::forest[1]/descendant::eLeaf[" & strLeafCondition & "]")
      ' Add this line
      AddOneLine(ndxLine, -1, colHtml) : colHtml.Add("<br>")
      ' Add the line we are at
      AddOneLine(ndxList, intThis, colHtml) : colHtml.Add("<br>")
      ' Show some more following lines
      For intI = 1 To intCount
        ' Get the next <forest>, if existing
        ndxLine = ndxList(intThis).SelectNodes("./ancestor::forest/following-sibling::forest[" & intI & "]/descendant::eLeaf[" & strLeafCondition & "]")
        ' Add this line
        AddOneLine(ndxLine, -1, colHtml) : colHtml.Add("<br>")
      Next intI
      ' Finish html output
      colHtml.Add("</body></html>")
      Me.wbYCOE.DocumentText = colHtml.Text
    Catch ex As Exception
      ' Show error
      HandleErr("frmTagSync/ShowYcoe error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   AddOneLine
  ' Goal:   Add the list of <eLeaf> elements in [ndxList] to [colHtml]
  ' History:
  ' 05-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub AddOneLine(ByRef ndxList As XmlNodeList, ByVal intThis As Integer, ByRef colHtml As StringColl)
    Dim strWord As String   ' One word
    Dim intI As Integer     ' Counter

    Try
      ' Validate
      If (ndxList Is Nothing) Then Exit Sub
      ' Do we actually have a list?
      If (ndxList.Count > 0) Then
        ' Find the <forest> and its location
        strWord = ndxList(0).SelectSingleNode("./ancestor::forest[1]").Attributes("Location").Value
        ' Check if we can just take the second part
        If (InStr(strWord, ":") > 0) Then
          strWord = Mid(strWord, InStr(strWord, ":") + 1)
        End If
        colHtml.Add("[" & strWord & "] ")
      End If
      ' Loop through the words
      For intI = 0 To ndxList.Count - 1
        ' Get the word
        strWord = StringOEtoTagged(LCase(ndxList(intI).Attributes("Text").Value)).Replace("$", "") & " "
        ' Is this the target word?
        If (intI = intThis) Then
          ' Add red and bolded word
          colHtml.Add("<b><font color='red'>" & strWord & "</font></b>")
        ElseIf (ndxList(intI).ParentNode.Attributes("Label").Value = "FW") Then
          ' Add grey word
          colHtml.Add("<font color='grey'>" & strWord & "</font>")
        Else
          ' Normal black word
          colHtml.Add(strWord)
        End If
      Next intI
    Catch ex As Exception
      ' Show error
      HandleErr("frmTagSync/AddOneLine error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  Private Sub cmdOEtagLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOEtagLeft.Click
    ' Let the offset wander to the left
    loc_OEtagOffset -= 1
    ' Show what we have right now
    ShowOEtagged(loc_arTagged, loc_intStart, loc_intLen, loc_intCurrent + loc_intStart + loc_OEtagOffset)
    ' Also make sure that we reset the other offset
    loc_YcoeOffset = 0
    ' Show what we have right now
    ShowYcoe(loc_ndxList, loc_intCurrent)
  End Sub

  Private Sub cmdOEtagRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOEtagRight.Click
    ' Let the offset wander to the left
    loc_OEtagOffset += 1
    ' Check if we need to adapt the length to the right
    If (loc_intCurrent + loc_OEtagOffset > loc_intLen) Then loc_intLen += 1
    ' Show what we have right now
    ShowOEtagged(loc_arTagged, loc_intStart, loc_intLen, loc_intCurrent + loc_intStart + loc_OEtagOffset)
    ' Also make sure that we reset the other offset
    loc_YcoeOffset = 0
    ' Show what we have right now
    ShowYcoe(loc_ndxList, loc_intCurrent)
  End Sub

  Private Sub cmdYcoeLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdYcoeLeft.Click
    ' Let the offset wander to the left
    loc_YcoeOffset -= 1
    ' Show what we have right now
    ShowYcoe(loc_ndxList, loc_intCurrent + loc_YcoeOffset)
    ' Also make sure that we reset the other offset
    loc_OEtagOffset = 0
    ' Show what we have right now
    ShowOEtagged(loc_arTagged, loc_intStart, loc_intLen, loc_intCurrent + loc_intStart)
  End Sub

  Private Sub cmdYcoeRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdYcoeRight.Click
    ' Let the offset wander to the left
    loc_YcoeOffset += 1
    ' Show what we have right now
    ShowYcoe(loc_ndxList, loc_intCurrent + loc_YcoeOffset)
    ' Also make sure that we reset the other offset
    loc_OEtagOffset = 0
    ' Show what we have right now
    ShowOEtagged(loc_arTagged, loc_intStart, loc_intLen, loc_intCurrent + loc_intStart)
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   cmdYcoeNextLine_Click
  ' Goal:   Signal that we have to go the next line 
  ' History:
  ' 05-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdYcoeNextLine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdYcoeNextLine.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.Retry
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   cmdYcoeSearch_Click
  ' Goal:   Try to find the next line in YCOE that matches the current loc_arTagged[loc_intStart]
  ' History:
  ' 05-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdYcoeSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdYcoeSearch.Click
    Dim ndxFor As XmlNode   ' Current forest
    Dim strWord As String   ' One word from [ndxList]
    Dim strTagd As String   ' One word from [arTagged]
    Dim intId As Integer    ' Id we are considering
    Dim intJ As Integer     ' Counter
    Dim bFound As Boolean   ' Whether we have found a matching line

    Try
      ' Find the current forest
      ndxFor = loc_ndxList(loc_intCurrent).SelectSingleNode("./ancestor::forest/following-sibling::forest[1]")
      While (ndxFor IsNot Nothing)
        ' Check if this is indeed a forest
        If (ndxFor.Name = "forest") Then
          ' Show where we are
          Status("TagSync - considering forestId = " & ndxFor.Attributes("forestId").Value)
          ' Determine the [ndxList] of this line
          loc_ndxList = ndxFor.SelectNodes("./descendant::eLeaf[" & strLeafCondition & "]")
          ' =========== DEBUG =============
          'If (ndxFor.Attributes("forestId").Value = 1147) Then Stop
          'Debug.Print(loc_arTagged(loc_intStart - 1), loc_arTagged(loc_intStart + 1))
          ' ===============================
          ' Are there any results?
          If (loc_ndxList.Count > 0) Then
            ' Compare at least 2 words
            bFound = True
            For intJ = 0 To 1
              ' Get one word from loc_ndxList
              strWord = StringOEtoTagged(LCase(loc_ndxList(intJ).Attributes("Text").Value)).Replace("$", "")
              ' Get corresponding word from arTagged
              strTagd = TagToWord(loc_arTagged(loc_intStart + intJ))
              ' Compare the two words
              If (Not OEtaggedWordsLike(strWord, strTagd)) Then
                ' If one word is not correct, then that's it!
                bFound = False
                Exit For
              End If
            Next intJ
            ' Do we have a result?
            If (bFound) Then
              ' We have a match! -- get the forest id
              intId = ndxFor.Attributes("forestId").Value
              ' This should be in the local copy
              loc_intForestId = intId
              ' Make sure this result ripples through...
              Me.DialogResult = System.Windows.Forms.DialogResult.Ignore
              Me.Close()
              Exit Sub
            End If

          End If
        End If
        ' Go to the next forest
        ndxFor = ndxFor.NextSibling
        ' Are we finishing?
        If (ndxFor Is Nothing) Then
          ' Ask user whether we should start at the beginning
          Select Case MsgBox("TagSync has not found results so far." & vbCrLf & _
                             "Re-start searching at the beginning?", MsgBoxStyle.YesNoCancel)
            Case MsgBoxResult.Cancel
              Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
              Me.Close() : Exit Sub
            Case MsgBoxResult.No
              Exit Sub
            Case MsgBoxResult.Yes
              ndxFor = pdxCurrentFile.SelectSingleNode("//descendant::forest[1]")
          End Select
        End If
      End While
      Status("TagSync: Could not find a match")
    Catch ex As Exception
      ' Show error
      HandleErr("frmTagSync/YcoeSearch error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   cmdOEtagSearch_Click
  ' Goal:   Advance the loc_OEtagOffset such, that the word in loc_arTagged matches the
  '           currently selected word in Ycoe
  ' History:
  ' 06-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdOEtagSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOEtagSearch.Click
    Dim strWord As String   ' One word from [ndxList]
    Dim strTagd As String   ' One word from [arTagged]
    Dim intI As Integer     ' Counter

    Try
      ' Get the currently selected word in Ycoe
      strWord = StringOEtoTagged(LCase(loc_ndxList(loc_intCurrent).Attributes("Text").Value)).Replace("$", "")
      ' Walk through the loc_artagged array
      For intI = loc_intStart + loc_intCurrent + loc_OEtagOffset To loc_arTagged.Length - 1
        ' Get the current word in arTagged
        strTagd = TagToWord(loc_arTagged(intI))
        ' Compare the two words
        If (OEtaggedWordsLike(strWord, strTagd)) Then
          ' We found the first word that matches -- take it!
          loc_OEtagOffset = intI - loc_intStart - loc_intCurrent
          ' Show what we have right now
          ShowOEtagged(loc_arTagged, loc_intStart, loc_intLen, loc_intCurrent + loc_intStart + loc_OEtagOffset)
          ' Exit nicely
          Exit Sub
        End If
      Next intI
      ' If we got here, we are not able to find anything useful
      Status("TagSync: could not find a match")
    Catch ex As Exception
      ' Show error
      HandleErr("frmTagSync/OEtagSearch error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   cmdYcoeGoto_Click
  ' Goal:   Goto the forest line indicated in the textbox
  ' History:
  ' 06-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdYcoeGoto_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdYcoeGoto.Click
    Dim ndxFor As XmlNode   ' Current forest
    Dim strGoto As String   ' Line to go to
    Dim intId As Integer    ' Forest id

    Try
      ' Get the line
      strGoto = Trim(Me.tbYcoeLine.Text)
      ' Validate
      If (Not IsNumeric(strGoto)) Then Exit Sub
      ' Find the first forest
      ndxFor = pdxCurrentFile.SelectSingleNode("//descendant::forest[1]")
      ' Try to find the place
      ndxFor = ndxFor.SelectSingleNode("./following-sibling::forest[@forestId = " & strGoto & "]", conTb)
      ' ndxFor = ndxFor.SelectSingleNode("//descendant::forest[@forestId = " & strGoto & "]")
      ' Found anything?
      If (ndxFor Is Nothing) Then Status("Could not find this node") : Exit Sub
      ' We have a match! -- get the forest id
      intId = ndxFor.Attributes("forestId").Value
      ' This should be in the local copy
      loc_intForestId = intId
      ' Make sure this result ripples through...
      Me.DialogResult = System.Windows.Forms.DialogResult.Ignore
      Me.Close()
    Catch ex As Exception
      ' Show error
      HandleErr("frmTagSync/OEtagSearch error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Class
