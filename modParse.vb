Module modParse
  ' ===================================== LOCAL VARIABLES ===========================================================
  ' ===================================== LOCAL CONSTANTS ===========================================================
  Private strSpace As String = " " & vbTab & vbCrLf
  Private strNoWord As String = ")" & strSpace
  Private SPEC_CHAR_IN As String = "tadegTADEG"
  Private SPEC_CHAR_OUT As String = "þæđëġÞÆĐËĠ"

  ' -------------------------------------------------------------------------------------------------------
  ' Name: VernToEnglish
  ' Goal: Convert a vernacular OE text into intelligable English
  ' Notes:
  '       - Special characters are defined as '+' + character:
  '         +t, +a,
  ' History:
  ' 27-11-2008    ERK Created
  ' 19-12-2008    ERK Adapted for VB2005 TreeBank module
  ' -------------------------------------------------------------------------------------------------------
  Public Function VernToEnglish(ByVal strText As String) As String
    Dim strOut As String = "" ' Output to be build up
    Dim intI As Integer       ' Position in the string

    ' Check all characters of the input
    For intI = 1 To Len(strText)
      ' Check this character for a key
      Select Case Mid(strText, intI, 1)
        Case "+"  ' This could be a special character
          ' Is the next character a special one?
          If (InStr(SPEC_CHAR_IN, Mid(strText, intI + 1, 1)) = 0) Then
            ' Just copy the input
            strOut &= Mid(strText, intI, 1)
          Else
            ' Goto next character
            intI = intI + 1
            ' Copy the translation of the special character
            strOut &= GetSpecChar(Mid(strText, intI, 1))
          End If
        Case Else ' Just copy the input
          strOut &= Mid(strText, intI, 1)
      End Select
    Next intI
    ' Return the string we have now made
    VernToEnglish = strOut
  End Function
  ' -------------------------------------------------------------------------------------------------------
  ' Name: GetSpecChar
  ' Goal: Given a special character with a + sign, convert into "normal" English
  ' History:
  ' 27-11-2008    ERK Created
  ' -------------------------------------------------------------------------------------------------------
  Private Function GetSpecChar(ByVal strIn As String) As String
    Dim intPos As Integer

    ' Get position in input string
    intPos = InStr(SPEC_CHAR_IN, Left(strIn, 1))
    If (intPos > 0) Then
      GetSpecChar = Mid(SPEC_CHAR_OUT, intPos, 1)
    Else
      ' Output the input with + sign
      GetSpecChar = "+" & strIn
    End If
  End Function
End Module
