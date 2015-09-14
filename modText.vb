Imports System.Xml
Module modText
  ' ===================================== LOCAL CONSTANTS ===========================================================
  Private loc_strSpace As String = " " & vbTab & vbCrLf ' What should be regarded as white space
  Private loc_strPunct As String = ".,<>;:'[]{}-_=+\|)(*&^%$#@!~`" & """" & "«»"
  Private loc_strAlpha As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'-0123456789" & _
                          "абвгдеёжзийклмнопрстуфхцчшщъыьэюяІІАБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ" & _
                          "áéóûÁÉÓÛáéíóúýÁÉÍÓÚÝàèìòùÀÈÌÒÙ"
  Private arNoSplit() As String = {"n", "an", "iin", "in", "chu", "ra", "gha", "ghachu"}
  Private arClitic() As String = {"q", "tie", "m"}
  ' ===================================== Global variables =======================================================
  Private intEtreeId As Integer = 0    ' Globally available ID 
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  DoWords
  ' Goal :  Parse a line of text into nodes under [ndxParent]
  ' Note :  Several features are added:
  '           feat/Cat  - the word category
  '           feat/Cyr  - cyrillic version of the word
  '           feat/mbt  - MBT's choice for the word's POS
  ' History:
  ' 22-03-2011  ERK Created
  ' 17-05-2012  ERC Added [strCyrLine]
  ' ----------------------------------------------------------------------------------------------------------
  Public Function DoWords(ByRef ndxParent As XmlNode, ByVal strIn As String, _
                          ByRef intSt As Integer, _
                          ByRef intEn As Integer, Optional ByVal bAuto As Boolean = False) As Boolean
    Dim ndxWork As XmlNode = Nothing  ' Working node
    Dim ndxEtree As XmlNode = Nothing ' The <eTree> element
    Dim strWord As String = ""    ' Word we found
    Dim strLatWord As String = "" ' Latin (lookup) word
    Dim strCyrWord As String = "" ' The corresponding cyrillic word
    Dim strAlt As String = ""     ' Alternative
    Dim strPrev As String = ""    ' Previous word
    Dim strPunc As String = ""    ' Punctuation we found
    Dim strType As String         ' Kind of punctuation
    Dim strPos As String = ""     ' Part of speech
    Dim strPosMbt As String = ""  ' Category guess from MBT
    Dim strCat As String = ""     ' The final category of this word
    Dim strFeat As String = ""    ' Features
    Dim strItem As String = ""    ' One item in the POS chooser
    Dim intId As Integer = 0      ' Word id number within this clause (starting with 0)
    Dim intPos As Integer         ' Position within string
    Dim intStart As Integer       ' Start position
    Dim intSpaces As Integer = 0  ' Number of spaces before a word

    Try
      ' Do we have anything at all?
      If (strIn = "") Then Return False
      ' Start with 1
      intPos = 1 : intSt = 1 : intEtreeId = 1
      ' Go into a loop
      Do
        Application.DoEvents()
        ' Check for interrupt
        If (bInterrupt) Then Return False
        ' Copy previous word
        strPrev = strWord
        ' Try to find a word
        intStart = intPos
        If (bnf_Word(strIn, intPos, strWord, bAuto)) AndAlso (strWord <> "") Then
          ' Add an <eTree> entry
          strPos = "X"
          ndxEtree = AddEtreeChild(ndxParent, intEtreeId, strPos, intStart, intPos - 1)
          ' Add the <eLeaf> under this one
          ndxWork = AddEleafChild(ndxEtree, "Vern", strWord, intStart, intPos - 1)
          ' Continue with the ID
          intId += 1
        End If
        ' Try to find punctuation
        intStart = intPos
        While (bnf_Punct(strIn, intPos, strWord))
          ' Check the kind of punctuation
          Select Case strWord
            Case ",", ":", ";", "<", ">"
              strType = ","
            Case ".", "!", "?"
              strType = "."
            Case """", "«", "»"
              strType = "&quot;"
            Case Else
              strType = "-"
          End Select
          ' Add an <eTree> entry
          ndxWork = AddEtreeChild(ndxParent, intEtreeId, strType, intStart, intPos - 1)
          ' Add the <eLeaf> under this one
          ' ndxWork = AddEleafChild(ndxWork, "Punct", XmlNamed(strWord), intStart, intPos - 1)
          ndxWork = AddEleafChild(ndxWork, "Punct", strWord, intStart, intPos - 1)
          ' Continue with the ID
          intId += 1
        End While
        ' Skip space
        bnf_Space(strIn, intPos)
      Loop While (intPos <= strIn.Length)
      ' Check the end position
      intEn = intPos - 1
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      MsgBox("modText/DoWords error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  bnf_Word
  ' Goal :  Anything can go in a word, except for a space sign and a left bracket (
  ' History:
  ' 14-12-2009  ERK Created
  ' 06-03-2013  ERK A word may not start with a hyphen
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function bnf_Word(ByRef strIn As String, ByRef intPos As Integer, _
           ByRef strBack As String, ByVal bAuto As Boolean) As Boolean
    Dim intBack As Integer = intPos   ' Roll back position
    Dim strWord As String             ' The word we come up with
    Dim intHyph As Integer            ' Position of hyphen
    Dim intI As Integer               ' Counter
    Dim bAsk As Boolean = True        ' Need to ask

    Try
      ' Try to get anything except SPACE signs and )
      While (InStr(loc_strAlpha, Mid(strIn, intPos, 1)) > 0) AndAlso (intPos <= Len(strIn))
        '' ================= DEBUGGING
        'If (Mid(strIn, intPos, 2) = "3-") Then Stop
        '' =================================
        ' Check for word-initial hyphens
        If ((intPos = 1) OrElse Mid(strIn, intPos - 1, 1) = " ") AndAlso (Mid(strIn, intPos, 1) = "-") Then
          strBack = "" : Return True
        End If
        ' Check clitic
        If (intPos > intBack) AndAlso (bnf_Clitic(strIn, intPos)) Then Exit While
        ' Go to the next position
        intPos += 1
      End While
      ' Possibly add period
      If (intPos < Len(strIn)) AndAlso (Mid(strIn, intPos, 1) = ".") Then
        ' Check if what follows after the period is empty or not
        If (Trim(Mid(strIn, intPos + 1)) <> "") Then
          ' We should add the period
          intPos += 1
        End If
      End If
      ' Are there word-medial hyphens?
      strWord = Mid(strIn, intBack, intPos - intBack)
      ' ================= DEBUGGING
      ' If (strWord = "3-gha") Then Stop
      ' =================================
      intHyph = InStr(strWord, "-")
      If (intHyph > 0) Then
        ' Check for allowable suffixes
        For intI = 0 To arNoSplit.Length - 1
          If (Mid(strWord, intHyph + 1) = arNoSplit(intI)) Then
            ' We have a word with a hyphen that needs to stay in one piece
            bAsk = False : strBack = strWord : Return True
          End If
        Next intI
        If (Not bAuto) AndAlso (bAsk) Then
          ' Ask user whether this word needs to be split or taken together
          Select Case MsgBox("Would you like this word to stay as is (Y) or split (N)?" & vbCrLf & _
                             "  [" & strWord & "]", MsgBoxStyle.YesNoCancel)
            Case MsgBoxResult.Cancel
              bInterrupt = True : Return False
            Case MsgBoxResult.No
              ' Set the position back!
              intPos = intBack + intHyph - 1
          End Select
        Else
          ' Split by default
          intPos = intBack + intHyph - 1
        End If
      End If
      ' Return the result
      strBack = Mid(strIn, intBack, intPos - intBack)
      ' Return success
      bnf_Word = (intPos > intBack)
    Catch ex As Exception
      ' Note error
      MsgBox("modText/bnf_Word error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  bnf_Punct
  ' Goal :  Punctuation is defined as not a word and not a space
  ' History:
  ' 14-12-2009  ERK Created
  ' 29-06-2013  ERK Allow only repeated punctuation of the same type
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function bnf_Punct(ByRef strIn As String, ByRef intPos As Integer, _
           ByRef strBack As String) As Boolean
    Dim intBack As Integer = intPos   ' Roll back position
    Dim strPunct As String            ' The specific punctuation we find

    Try
      ' Validate
      If (intPos > Len(strIn)) Then Return False
      ' Get the first character
      strPunct = Mid(strIn, intPos, 1)
      ' Check for the first punctuation sign
      If (strPunct = "-" OrElse InStr(loc_strAlpha & loc_strSpace, strPunct) = 0) Then
        ' Go to the next position
        intPos += 1
        ' Check if more of this are following
        While (intPos <= Len(strIn)) AndAlso (Mid(strIn, intPos, 1) = strPunct)
          intPos += 1
        End While
      End If
      ' Return the result
      strBack = Mid(strIn, intBack, intPos - intBack)
      ' Return success
      bnf_Punct = (intPos > intBack)
    Catch ex As Exception
      ' Note error
      MsgBox("modText/bnf_Punct error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  bnf_Space
  ' Goal :  Skip any number of spaces
  ' History:
  ' 14-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function bnf_Space(ByRef strIn As String, ByRef intPos As Integer) As Integer
    Dim intNum As Integer = 0 ' Number of spaces skipped

    Try
      ' Try skip spaces
      While (InStr(loc_strSpace, Mid(strIn, intPos, 1)) > 0) AndAlso (intPos <= Len(strIn))
        ' Skip one space
        intPos += 1
        intNum += 1
      End While
      ' Return the number of spaces actually skipped
      bnf_Space = intNum
    Catch ex As Exception
      ' Note error
      MsgBox("modText/bnf_Space error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  bnf_Clitic
  ' Goal :  Check if a clitic is following
  ' History:
  ' 17-09-2011  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function bnf_Clitic(ByRef strIn As String, ByRef intPos As Integer) As Boolean
    Dim intI As Integer               ' Counter

    Try
      ' Do we have a hyphen?
      If (Mid(strIn, intPos, 1) = "-") Then
        ' Walk all clitic possibilities
        For intI = 0 To arClitic.Length - 1
          ' Is this clitic present?
          If (Mid(strIn, intPos + 1, arClitic(intI).Length) = arClitic(intI)) Then
            ' Is it followed by a non-alpha character?
            If (intPos + 1 + arClitic(intI).Length >= strIn.Length) OrElse _
               (InStr(loc_strAlpha, Mid(strIn, intPos + 1 + arClitic(intI).Length, 1)) = 0) Then
              ' This is a clitic
              Return True
            End If
          End If
        Next intI
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Warn the user
      MsgBox("modText/bnf_Clitic error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  NextWord
  ' Goal :  Get the next word, starting from [intPos]
  ' Note :  This is actually a tokenizer, but then a VERY SIMPLE one, and one that makes errors...
  ' History:
  ' 08-08-2011  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function NextWord(ByRef strIn As String, ByRef intPos As Integer, ByRef strWord As String) As Boolean
    Dim intBack As Integer = intPos   ' Roll back position

    Try
      ' Validate
      If (strIn = "") Then Return False
      ' Skip any non-word characters
      While (InStr(loc_strAlpha, Mid(strIn, intPos, 1)) = 0) AndAlso (intPos <= Len(strIn))
        ' Go to the next position
        intPos += 1
      End While
      ' Check for end of line
      If (intPos >= strIn.Length) Then Return False
      ' Adapt the rollback position?
      If (intPos > intBack) Then intBack = intPos
      ' Find the end of the word
      While (InStr(loc_strAlpha, Mid(strIn, intPos, 1)) > 0) AndAlso (intPos <= Len(strIn))
        ' Go to the next position
        intPos += 1
      End While
      ' Return the result
      strWord = Mid(strIn, intBack, intPos - intBack)
      ' Return success
      Return (intPos > intBack)
    Catch ex As Exception
      ' Warn the user
      MsgBox("modText/NextWord error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
End Module
