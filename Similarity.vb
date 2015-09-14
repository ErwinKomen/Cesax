Imports System.Text.RegularExpressions
Public Class Similarity
  ' ================================= General =================================================================
  ' Name:   Similarity
  ' Goal:   Functions that calculate a similarity measure between two words
  ' History:
  ' 30-05-2013  ERK Copied from internet
  ' 03-06-2013  ERK Added my own, Levenstein-based, similarity measure
  ' ============================================================================================================
  ' ================================= LOCAL VARIABLES =========================================================
  Private loc_arLev(40, 40) As Integer   ' Contains the Levenstein-like similarity values
  ' i)	Vertaling: th=2, dh=3, ae=4, hw=5, sc=6, ea=7, eo=8, ie=9
  Private arSeqIn() As String = {"th", "dh", "ae", "hw", "sc"}
  Private arSeqOut() As String = {"2", "3", "4", "5", "6"}
  Private lev_strSubstOrg As String = "abcdefghijklmnopqrstuvwxyz23456789"
  Private lev_strSubstCs1 As String = "e ktaw  y c nmu    do uci 327 x4ee"
  Private lev_strSubstCs2 As String = "op  ou ce q   ebk   w  se td      "
  Private lev_strInsDGain As String = "0000002110000000000000001000000000"
  Private lev_strVowel As String = "a|e|i|o|u|y|4"
  Private lev_strAnyVowel As String = "aeiouy4"
  Private lev_intEq As Integer = 7
  ' ============================================================================================================
  ' -----------------------------------------------------------------------------------------
  ' Name:   LevenSimi
  ' Goal:   Calculate a similarity measure between S and T
  '         Basically use the Levenstein distance matrix (n+1)*(m+1)
  '         Use different 'awards' for:
  '           e = equal           - fixed 7
  '           s = substitute      - range 4-6
  '           d/i = delete/insert - range 0-1
  '         Exceptions:
  '           Insertion/deletion of a simple "g" should have a gain of 6 (almost perfect)
  '         Requirements: 
  '           e   >  s  
  '           s/2 >= d
  ' Return: The percentage match from 0 to 100, where 100 is perfect match
  ' History:
  ' 03-06-2013  ERK Created
  ' 04-06-2013  ERK Added back-tracking of the result and evaluation of the i, d, s actions
  ' -----------------------------------------------------------------------------------------
  Public Function LevenSimi(ByVal strS_in As String, ByVal strT_in As String) As Integer
    Dim strS As String      ' Converted version of strS_in
    Dim strT As String      ' Converted version of strT_in
    ' Dim strOut As String    ' Output
    Dim arS() As Char       ' Character array version of S
    Dim arT() As Char       ' Character array versino of T
    Dim chS As Char         ' Current character from S
    Dim chT As Char         ' Current character from T
    Dim chSprec As Char     ' Preceding character in S
    Dim intM As Integer     ' dimension of S
    Dim intN As Integer     ' Dimension of T
    Dim intI As Integer     ' Counter
    Dim intJ As Integer     ' Counter
    Dim intGain As Integer  ' Total gain
    Dim intSubst As Integer ' Totalnumber of substitutions

    Try
      ' Do simple comparison
      If (strS_in = strT_in) Then Return 100
      If (strS_in = "") OrElse (strT_in = "") Then Return 0
      ' ============== DEBUG ==============
      'If (strS_in = "ahge" AndAlso strT_in = "graeg") Then Stop
      ' ===================================
      ' Convert the strings
      strS = LevSeqToNum(strS_in, True) : strT = LevSeqToNum(Left(strT_in, 40), True)
      arS = strS.ToCharArray : arT = strT.ToCharArray
      ' Determine M and N
      intN = strS.Length : intM = strT.Length
      ' Initialize row 0: all insertions
      For intI = 1 To intN
        ' Add the gain of inserting S[intI] at the beginning
        loc_arLev(intI, 0) = LevInsDelGain(arS(intI - 1)) + loc_arLev(intI - 1, 0)
      Next intI
      ' Initialize column 0: all deletions
      For intJ = 1 To intM
        ' Add the gain of deleting T[intJ]
        loc_arLev(0, intJ) = LevInsDelGain(arT(intJ - 1)) + loc_arLev(0, intJ - 1)
      Next intJ
      ' Walk the rows and columns and determine the step with the most gain
      For intI = 1 To intN
        For intJ = 1 To intM
          ' Get current characters from S and T
          chS = arS(intI - 1) : chT = arT(intJ - 1)
          ' Decide what the most gainful step is for element [j,i] (row=j, column=i)
          loc_arLev(intI, intJ) = Max3( _
             loc_arLev(intI - 1, intJ) + LevInsDelGain(chT), _
             loc_arLev(intI, intJ - 1) + LevInsDelGain(chS), _
             loc_arLev(intI - 1, intJ - 1) + LevSubstGain(chS, chT))
        Next intJ
      Next intI
      ' Maximum resulting similarity is in the last row+column
      intGain = loc_arLev(intN, intM)
      '' ======== DEBUG =============
      'Debug.Print(strS_in & " > " & strT_in & " = " & intGain & "%")
      'For intI = 0 To intN
      '  strOut = "[" & intI & "] "
      '  For intJ = 0 To intM
      '    strOut &= loc_arLev(intI, intJ) & vbTab
      '  Next intJ
      '  Debug.Print(strOut)
      'Next intI
      'Stop
      '' ============================

      ' Back-track the result by starting in (0,0) and following the path
      intI = 0 : intJ = 0 : chSprec = "-" : intSubst = 0
      While (intI < intN AndAlso intJ < intM)
        ' Get current character
        chS = arS(intI)
        ' Check what the best path is from this one on
        If (loc_arLev(intI, intJ + 1) > loc_arLev(intI + 1, intJ + 1)) AndAlso (loc_arLev(intI, intJ + 1) > loc_arLev(intI + 1, intJ)) Then
          ' Right is largest; Advance indices
          intJ += 1
          ' Check if insertion of this character is allowed
          ' If (Not LevAcceptInsDel(arT(intJ - 1), chSprec)) Then intGain = 0
        ElseIf (loc_arLev(intI + 1, intJ + 1) > loc_arLev(intI + 1, intJ)) AndAlso (loc_arLev(intI + 1, intJ + 1) > loc_arLev(intI + 1, intJ)) Then
          ' Diagonal is largest; Advance indices
          intI += 1 : intJ += 1
          ' Check if substitution of this character is allowed for the other one
          ' If (Not LevAcceptSubst(arS(intI - 1), arT(intJ - 1))) Then intGain = 0
          ' Keep tack of previous and current S character
          chSprec = chS
          intSubst += 1
        Else
          ' Below is largest; Advance indices
          intI += 1
          ' Check if deletion of this character is allowed
          ' If (Not LevAcceptInsDel(arS(intI - 1), chSprec)) Then intGain = 0
          ' Keep tack of previous and current S character
          chSprec = chS
        End If
      End While
      ' Check the number of substitutions w.r.t. the maximum string length
      If ((100 * intSubst \ Math.Max(intM, intN)) < 75) Then
        intGain = 0
      End If

      ' Normalize the total figure: divide by equality gain and by the total number of M+N
      intGain = intGain * 100 \ (Math.Max(intM, intN) * lev_intEq)
      '' ======== DEBUG =============
      'Debug.Print(strS_in & " > " & strT_in & " = " & intGain & "%")
      'For intI = 0 To intN
      '  strOut = "[" & intI & "] "
      '  For intJ = 0 To intM
      '    strOut &= loc_arLev(intI, intJ) & vbTab
      '  Next intJ
      '  Debug.Print(strOut)
      'Next intI
      '' ============================
      ' Return this gain
      Return intGain
    Catch ex As Exception
      ' Show error
      HandleErr("Similarity/LevenSimi error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return 0
    End Try
  End Function
  ' -----------------------------------------------------------------------------------------
  ' Name:   LevenDist
  ' Goal:   Calculate a similarity measure between S and T
  '         Basically use the Levenstein distance matrix (n+1)*(m+1)
  '         Use different 'awards' for:
  '           e = equal           - fixed 0
  '           s = substitute      - fixed -2
  '           d/i = delete/insert - fixed -2
  ' Return: Number of subst/del/ins operations and the number of equals
  ' History:
  ' 06-06-2013  ERK Created
  ' -----------------------------------------------------------------------------------------
  Public Function LevenDist(ByVal strS_in As String, ByVal strT_in As String, ByRef intEquals As Integer) As Integer
    Dim strS As String      ' Converted version of strS_in
    Dim strT As String      ' Converted version of strT_in
    Dim arS() As Char       ' Character array version of S
    Dim arT() As Char       ' Character array versino of T
    Dim chS As Char         ' Current character from S
    Dim chT As Char         ' Current character from T
    Dim chSprec As Char     ' Preceding character in S
    Dim intM As Integer     ' dimension of S
    Dim intN As Integer     ' Dimension of T
    Dim intI As Integer     ' Counter
    Dim intJ As Integer     ' Counter
    Dim intGain As Integer  ' Total gain
    Dim intSubst As Integer ' Totalnumber of substitutions
    ' Dim intEquals As Integer ' Number of equals
    Dim intInsDelGain As Integer = -2
    Dim intSubstGain As Integer = -2

    Try
      ' Do simple comparison
      If (strS_in = strT_in) Then Return 100
      If (strS_in = "") OrElse (strT_in = "") Then Return 0
      ' ============== DEBUG ==============
      'If (strS_in = "ahge" AndAlso strT_in = "graeg") Then Stop
      ' ===================================
      ' Convert the strings
      strS = LevSeqToNum(strS_in, True) : strT = LevSeqToNum(Left(strT_in, 40), True)
      arS = strS.ToCharArray : arT = strT.ToCharArray
      ' Determine M and N
      intN = strS.Length : intM = strT.Length : intEquals = 0 : intSubst = 0
      ' Initialize row 0: all insertions
      For intI = 1 To intN
        ' Add the gain of inserting S[intI] at the beginning
        loc_arLev(intI, 0) = intInsDelGain + loc_arLev(intI - 1, 0)
      Next intI
      ' Initialize column 0: all deletions
      For intJ = 1 To intM
        ' Add the gain of deleting T[intJ]
        loc_arLev(0, intJ) = intInsDelGain + loc_arLev(0, intJ - 1)
      Next intJ
      ' Walk the rows and columns and determine the step with the most gain
      For intI = 1 To intN
        For intJ = 1 To intM
          ' Get current characters from S and T
          chS = arS(intI - 1) : chT = arT(intJ - 1)
          ' Check for similarity
          If (chS = chT) Then
            ' They are equal!
            loc_arLev(intI, intJ) = loc_arLev(intI - 1, intJ - 1)
          Else
            ' Decide what the most gainful step is for element [j,i] (row=j, column=i)
            loc_arLev(intI, intJ) = Max3( _
               loc_arLev(intI - 1, intJ) + intInsDelGain, _
               loc_arLev(intI, intJ - 1) + intInsDelGain, _
               loc_arLev(intI - 1, intJ - 1) + intSubstGain)
          End If
        Next intJ
      Next intI
      ' Maximum resulting similarity is in the last row+column
      intGain = loc_arLev(intN, intM)
      ' Back-track the result by starting in (0,0) and following the path
      intI = 0 : intJ = 0 : chSprec = "-" : intSubst = 0 : chS = "-"
      While (intI < intN OrElse intJ < intM)
        ' Get current character
        If (intI > 0 AndAlso intI < intN) Then chS = arS(intI - 1)
        If (intJ > 0 AndAlso intJ < intM) Then chT = arT(intJ - 1)
        ' If (intJ > intM) Then Stop
        ' Check what the best path is from this one on
        If (intI >= intN AndAlso intJ < intM) OrElse (intJ < intM AndAlso (loc_arLev(intI, intJ + 1) > loc_arLev(intI + 1, intJ + 1)) AndAlso (loc_arLev(intI, intJ + 1) > loc_arLev(intI + 1, intJ))) Then
          ' Right is largest; Advance indices
          intJ += 1
          ' Check if insertion of this character is allowed
          ' If (Not LevAcceptInsDel(arT(intJ - 1), chSprec)) Then intGain = 0
          intSubst += 2
        ElseIf (intJ < intM) AndAlso (intI < intN) AndAlso (loc_arLev(intI + 1, intJ + 1) > loc_arLev(intI + 1, intJ)) AndAlso (loc_arLev(intI + 1, intJ + 1) > loc_arLev(intI, intJ + 1)) Then
          ' Diagonal is largest; Advance indices
          intI += 1 : intJ += 1
          ' Check if substitution of this character is allowed for the other one
          ' If (Not LevAcceptSubst(arS(intI - 1), arT(intJ - 1))) Then intGain = 0
          '' Keep tack of previous and current S character
          'chSprec = chS
          ' Check for equality
          If (loc_arLev(intI - 1, intJ - 1) = loc_arLev(intI, intJ)) Then
            ' They are equal
            intEquals += 2
          ElseIf (InStr(lev_strAnyVowel, chS) > 0 AndAlso chT = "g") OrElse _
                 (InStr(lev_strAnyVowel, chT) > 0 AndAlso chS = "g") Then
            intEquals += 1 : intSubst += 1
          Else
            intSubst += 2
          End If
        Else
          ' Below is largest; Advance indices
          intI += 1
          ' Check if deletion of this character is allowed
          ' If (Not LevAcceptInsDel(arS(intI - 1), chSprec)) Then intGain = 0
          '' Keep tack of previous and current S character
          'chSprec = chS
          intSubst += 2
        End If
      End While
      ' Compare the substitutions and the equals for a simple test
      If (intEquals < intSubst) Then Return 0
      ' Calculate a measure
      intGain = (100 * intEquals) / (intSubst + 2 * Math.Max(intM, intN))
      '' Check the number of substitutions w.r.t. the maximum string length
      'If ((100 * intSubst \ Math.Max(intM, intN)) < 75) Then
      '  intGain = 0
      'End If
      ' Check for high numbers
      'If (intGain >= 75) Then
      '  Stop
      'End If

      '' Normalize the total figure: divide by equality gain and by the total number of M+N
      'intGain = intGain * 100 \ (Math.Max(intM, intN) * lev_intEq)
      ' Return this gain
      Return intGain
    Catch ex As Exception
      ' Show error
      HandleErr("Similarity/LevenDist error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return 0
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   LevAcceptSubst
  ' Goal:   Check if substitution of S by T is allowed
  ' History:
  ' 03-06-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function LevAcceptSubst(ByVal chS As Char, ByVal chT As Char) As Boolean
    Try
      ' Allow substitution of one vowel by another vowel
      If (DoLike(chS, lev_strVowel)) Then
        If (DoLike(chT, lev_strVowel)) Then
          Return True
        ElseIf (chS = "i" AndAlso chT = "j") Then
          Return True
        Else
          Return False
        End If
      Else
        ' Character in S is a consonant -- only allow certain substitutions
        Select Case chS
          Case "c"
            Return (chT = "k")
          Case "k"
            Return (chT = "c")
          Case "2"
            Return (chT = "3")
          Case "3"
            Return (chT = "2")
          Case "x"
            Return (chT = "6")
          Case "6"
            Return (chT = "x")
          Case "j"
            Return (chT = "i")
          Case Else
            Return False
        End Select
        ' Return (LevSubstGain(chS, chT) > 4)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("Similarity/LevAcceptSubst error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   LevAcceptInsDel
  ' Goal:   Check if insertion/deletion of X is allowed
  ' History:
  ' 03-06-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function LevAcceptInsDel(ByVal chX As Char, ByVal chPrev As Char) As Boolean
    Try
      ' Check the character...
      Select Case chX
        Case "g", "h" ' Allow insertion/deletion of g/h anywhere
          Return True
        Case "a", "e", "i", "o", "u", "y", "4"
          ' Allow insertion/deletion of vowels only if preceding one is a vowel too
          Return (DoLike(chPrev, lev_strVowel))
        Case Else
          ' Other insertions/deletions are not allowed
          Return False
      End Select
    Catch ex As Exception
      ' Show error
      HandleErr("Similarity/LevAcceptSubst error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   Max3
  ' Goal:   Give the maximum of numbers i,j,k
  ' History:
  ' 03-06-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function Max3(ByVal intI As Integer, ByVal intJ As Integer, ByVal intK As Integer) As Integer
    Try
      Return Math.Max(intI, Math.Max(intJ, intK))
    Catch ex As Exception
      ' Show error
      HandleErr("Similarity/Max3 error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return 0
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   LevInsDelGain
  ' Goal:   What are the 'benefits' of inserting S?
  '         The result must be an integer 0, 1 or 2
  ' History:
  ' 18-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function LevInsDelGain(ByVal chS As Char) As Integer
    Dim intIdx As Integer ' Index
    Dim intGain As Integer  ' Costs

    Try
      ' Get the index within lev_strSubstOrg
      intIdx = Asc(LCase(chS)) - Asc("a") + 1
      ' Check range --> return extra large distance if needed
      If (intIdx <= 0) OrElse (intIdx > lev_strSubstOrg.Length) Then
        ' There only is the minimal default gain, which is: nothing
        intGain = 0
      Else
        ' Calculate the gain -- there could actually be some gain
        intGain = CInt(Mid(lev_strInsDGain, intIdx, 1))
      End If
      ' Make sure zero-gain is lowered to -1
      If (intGain = 0) Then intGain = -1
      ' Return the insertion costs (which are 1 or 2)
      Return intGain
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/LevInsDelGain error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   LevSubstGain
  ' Goal:   What are the gainst of substituting S for T?
  '         The result is 3-5 for substitution, and 6 for equality
  ' History:
  ' 03-06-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function LevSubstGain(ByVal chS As Char, ByVal chT As Char) As Integer
    Dim intIdx As Integer ' Index

    Try
      ' Are they equal? --> no cost
      If (chS = chT) Then Return lev_intEq
      ' Get the index within lev_strSubstOrg
      intIdx = Asc(LCase(chS)) - Asc("a") + 1
      ' Check range --> return minimum if not found
      If (intIdx <= 0) OrElse (intIdx > lev_strSubstOrg.Length) Then Return 3
      ' Check if the costs are 1, 2 or 3
      If (Mid(lev_strSubstCs1, intIdx, 1) = chT) Then
        ' Level 1 substitution has the most gain
        Return 6
      ElseIf (Mid(lev_strSubstCs2, intIdx, 1) = chT) Then
        ' Level 2 substitution still has some gain
        Return 4
      End If
      ' No special case: gain of substitution is lowest
      ' Was: Return 4
      Return -2
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/LevSubstGain error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return 0
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   LevSeqToNum
  ' Goal:   Convert certain char sequences into numbers
  ' History:
  ' 31-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function LevSeqToNum(ByVal strIn As String, ByVal bForward As Boolean) As String
    Dim intI As Integer   ' Counter

    Try
      ' Validate
      If (strIn = "") Then Return ""
      ' Check all possibilities
      For intI = 0 To arSeqIn.Length - 1
        ' Forward or backward?
        If (bForward) Then
          strIn = strIn.Replace(arSeqIn(intI), arSeqOut(intI))
        Else
          strIn = strIn.Replace(arSeqOut(intI), arSeqIn(intI))
        End If
      Next intI
      ' Return the result 
      Return strIn
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/LevSeqToNum error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   CompareStrings
  ' Goal:   This function implements string comparison algorithm based on character pair similarity
  ' Source: http://www.catalysoft.com/articles/StrikeAMatch.html
  ' Return: The percentage match from 0.0 to 1.0 where 1.0 is 100%
  ' History:
  ' 30-05-2013  ERK Copied from internet
  ' ------------------------------------------------------------------------------------
  Public Function CompareStrings(ByVal str1 As String, ByVal str2 As String) As Double
    Try
      ' Dim pairs1 As List(Of String) = WordLetterPairs(str1.ToUpper())
      ' Dim pairs2 As List(Of String) = WordLetterPairs(str2.ToUpper())
      Dim pairs1 As List(Of String) = WordLetterPairs(str1)
      Dim pairs2 As List(Of String) = WordLetterPairs(str2)
      ' Initialise number of intersections
      Dim intersection As Integer = 0
      ' Calculate the baseline
      Dim union As Integer = pairs1.Count + pairs2.Count
      ' Check against zero-dividing
      If (union = 0) Then Return 0

      For i As Integer = 0 To pairs1.Count - 1
        For j As Integer = 0 To pairs2.Count - 1
          If pairs1(i) = pairs2(j) Then
            intersection += 1
            pairs2.RemoveAt(j)
            'Must remove the match to prevent "GGGG" from appearing to match "GG" with 100% success
            Exit For
          End If
        Next
      Next
      ' Return the result
      Return (2.0 * intersection) / union
    Catch ex As Exception
      ' Show error
      HandleErr("Similarity/CompareStrings error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return 0
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   WordLetterPairs
  ' Goal:   Get all letter pairs for each individual word in the string
  ' History:
  ' 30-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function WordLetterPairs(ByVal str As String) As List(Of String)
    Dim AllPairs As New List(Of String)()
    Dim Words As String()

    Try
      ' Simplification
      If Not String.IsNullOrEmpty(str) Then
        ' Find the pairs of characters
        Dim PairsInWord As [String]() = LetterPairs(str)
        For p As Integer = 0 To PairsInWord.Length - 1
          AllPairs.Add(PairsInWord(p))
        Next
      End If
      ' Skip this original code
      If (False) Then
        ' Tokenize the string and put the tokens/words into an array
        Words = Regex.Split(str, "\s")

        ' For each word
        For w As Integer = 0 To Words.Length - 1
          If Not String.IsNullOrEmpty(Words(w)) Then
            ' Find the pairs of characters
            Dim PairsInWord As [String]() = LetterPairs(Words(w))
            For p As Integer = 0 To PairsInWord.Length - 1
              AllPairs.Add(PairsInWord(p))
            Next
          End If
        Next
      End If
      ' Return all the pairs
      Return AllPairs
    Catch ex As Exception
      ' Show error
      HandleErr("Similarity/WordLetterPairs error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   LetterPairs
  ' Goal:   Generates an array containing every two consecutive letters in the input string
  ' History:
  ' 30-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function LetterPairs(ByVal str As String) As String()
    Try
      ' Validate
      Dim numPairs As Integer = str.Length - 1

      Dim pairs As String() = New String(numPairs - 1) {}

      For i As Integer = 0 To numPairs - 1
        pairs(i) = str.Substring(i, 2)
      Next
      ' Return the result
      Return pairs
    Catch ex As Exception
      ' Show error
      HandleErr("Similarity/LetterPairs error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   New
  ' Goal:   Initialisation of the Levenstein array
  ' History:
  ' 30-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub New()
    Try
      ' Initialize element (0)(0) (no gain there)
      loc_arLev(0, 0) = 0
    Catch ex As Exception
      ' Show error
      HandleErr("Similarity/New error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Class
