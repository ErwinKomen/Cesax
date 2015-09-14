Option Strict On
Option Explicit On
Option Compare Binary
Public Class frmFind
  Private WithEvents rtbThis As RichTextBox
  Private objFindOptions As Integer = RichTextBoxFinds.None
  Private intSelStart As Integer = 1  ' Last activated selection
  Private intSelLen As Integer = 0    ' Length of last selection
  Private bSelReset As Boolean = True ' Whether selection was reset successfully
  ' ------------------------------------------------------------------------------------
  ' Name:   chbWhole_CheckedChanged
  ' Goal:   Only find whole words
  ' History:
  ' 06-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub chbWhole_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbWhole.CheckedChanged
    ' Adapt the FindOptions value
    If (Me.chbWhole.Checked) Then
      ' Set this option
      objFindOptions += RichTextBoxFinds.WholeWord
    Else
      ' Delete this option
      objFindOptions -= RichTextBoxFinds.WholeWord
    End If
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   chbCase_CheckedChanged
  ' Goal:   Search case sensitive
  ' History:
  ' 06-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub chbCase_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbCase.CheckedChanged
    ' Adapt the FindOptions value
    If (Me.chbCase.Checked) Then
      ' Set this option
      objFindOptions += RichTextBoxFinds.MatchCase
    Else
      ' Delete this option
      objFindOptions -= RichTextBoxFinds.MatchCase
    End If
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   cmdNext_Click
  ' Goal:   Search for the next word
  ' History:
  ' 06-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub cmdNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNext.Click
    DoFind(True)
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   cmdPrev_Click
  ' Goal:   Search for the previous occurrance of the word
  ' History:
  ' 06-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdPrev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrev.Click
    DoFind(False)
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   DoFind
  ' Goal:   Generalized find function
  ' History:
  ' 07-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub DoFind(ByVal bForward As Boolean)
    Dim strSearch As String   ' The string to search
    Dim intPos As Integer     ' Current position

    Try
      ' Get the search string
      strSearch = Me.tbFind.Text
      ' Get current position
      intPos = rtbThis.SelectionStart
      ' Do we need to find backwards?
      If (bForward) Then
        ' Add to the selection
        intPos += 1
        ' Check the validity of this value
        If (intPos >= rtbThis.TextLength) Then intPos = 1
      Else
        ' Set reverse find option
        objFindOptions += RichTextBoxFinds.Reverse
        ' Subtract from selection
        intPos -= 1
        ' Check the validity of this value
        If (intPos < 1) Then intPos = rtbThis.TextLength - 1
      End If
      ' Access the active form
      With rtbThis
        ' Use the RTB find function, depending on the direction
        If (bForward) Then
          ' Forward finding takes place in the range [intPos,--]
          intPos = .Find(strSearch, intPos, CType(objFindOptions, System.Windows.Forms.RichTextBoxFinds))
        Else
          ' Backward finding must take place in the range [0,intPos]
          intPos = .Find(strSearch, 0, intPos, CType(objFindOptions, System.Windows.Forms.RichTextBoxFinds))
        End If
        ' Did we find something?
        If (intPos > 0) Then
          ' Now select what we have found
          .Focus()
          .Select(intPos, Len(strSearch))
          ' Indicate status
          Status("Found expression at: " & intPos)
          ' Show selection
          ShowSelect()
        Else
          ' Indicate that there is nothing more
          .Focus()
          Status("No (more) matches")
        End If
      End With
      ' Do we need to reset reverse?
      If (Not bForward) Then
        ' Set reverse find option
        objFindOptions -= RichTextBoxFinds.Reverse
      End If
      ' Set focus back to the find form
      Me.Focus()
    Catch ex As Exception
      ' Show error
      HandleErr("frmFind/DoFind error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   Textbox
  ' Goal:   Set the textbox to the one now active
  ' History:
  ' 29-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public WriteOnly Property Textbox() As RichTextBox
    Set(ByVal value As RichTextBox)
      rtbThis = value
    End Set
  End Property

  ' ------------------------------------------------------------------------------------
  ' Name:   frmFind_Activated
  ' Goal:   What to do when the FIND form is activated
  ' History:
  ' 07-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmFind_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
    Try
      ' Access the Find textbox
      With Me.tbFind
        ' Does textbox contain text?
        If (.TextLength > 0) Then
          ' Select the text
          .Select(0, .TextLength)
        End If
        ' Set focus on the textbox
        .Focus()
      End With
    Catch ex As Exception
      ' Show error
      HandleErr("frmFind/Activated error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   frmFind_KeyUp
  ' Goal:   Look for ESC key pressed
  ' History:
  ' 06-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmFind_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
    ' Check whether escape is pressed
    If (e.KeyCode = Keys.Escape) Then
      ' Leave this form
      Me.Hide()
    End If
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   frmFind_Load
  ' Goal:   Initialise form
  ' History:
  ' 06-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmFind_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Initialise the search string
    Me.tbFind.Text = ""
    ' Set my owner
    Me.Owner = frmMain
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   ShowSelect
  ' Goal:   Show the selection
  ' History:
  ' 07-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub ShowSelect()
    ' Access the RTB of the active form
    With rtbThis
      ' Set colours to selection
      .SelectionBackColor = Color.DarkBlue
      .SelectionColor = Color.White
      ' Store values of this selection
      intSelStart = .SelectionStart
      intSelLen = .SelectionLength
      ' Indicate that selection was not reset
      bSelReset = False
    End With
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   rtbThis_Click
  ' Goal:   Close FIND form when user clicks on [frmThis]
  ' History:
  ' 07-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub rtbThis_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtbThis.Click
    ' Clicking on the window where find takes place should close the Find window
    Me.Hide()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   rtbThis_SelectionChanged
  ' Goal:   Reset selection when user starts typing
  ' History:
  ' 07-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub rtbThis_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rtbThis.SelectionChanged
    Dim intThisStart As Integer
    Dim intThisLen As Integer

    ' Do we need to act?
    If (bSelReset) Then Exit Sub
    With rtbThis
      ' Get the current selection
      intThisStart = .SelectionStart
      intThisLen = .SelectionLength
      ' Delete the previous selection
      If (intSelLen > 0) Then
        ' Select previous stuff
        .Select(intSelStart, intSelLen)
        ' Reverse colours again
        .SelectionBackColor = Color.White
        .SelectionColor = Color.Black
        ' Reset selection
        .Select(intThisStart, intThisLen)
      End If
    End With
    ' Indicate that selection was reset
    bSelReset = True
  End Sub
End Class