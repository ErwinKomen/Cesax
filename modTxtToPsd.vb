Imports System.IO
Imports System.Text.RegularExpressions
Module modTxtToPsd
  ' ===================================== GLOBAL VARIABLES ==========================================================
  ' ===================================== LOCAL VARIABLES ===========================================================
  Private strStanfordParser As String = ""  ' Location of the stanford parser
  Private strStanfordGerman As String = ""  ' German definition file
  Private strStanfordEnglish As String = "" ' Englishdefinition file
  Private strHttpDir As String = "http://erwinkomen.ruhosting.nl/software/stp/"
  Private strStanfordShellIntroGerman As String = ""
  Private strStanfordShellIntroEnglish As String = ""
  Private strStanfordShellIntroEngTrans As String = ""
  Private strStanfordShellTokenize As String = ""
  Private bCsError As Boolean = False       ' Error in CS processing
  Private colCSoutput As New StringColl     ' Output of the CS process
  Private colCSerror As New StringColl      ' Output of the CS process
  Private intCsOutLine As Integer = 0       ' Progress for the CS process
  Private intCsErrLine As Integer = 0       ' Progress for the CS process
  Private strShort As String = ""           ' Short file name
  ' ===================================== LOCAL CONSTANTS ===========================================================
  Private Const STANFORD_PARSER As String = "stanford-parser.jar"
  Private Const STANFORD_GERMAN As String = "germanPCFG.ser.gz"
  Private Const STANFORD_ENGLISH As String = "englishPCFG.ser.gz"
  ' Private Const STANFORD_INTRO As String = "-mx800m -classpath " & """"
  Private Const STANFORD_INTRO As String = "-mx1200m -classpath " & """"
  Private Const STANFORD_LOC As String = " edu.stanford.nlp.parser.lexparser.LexicalizedParser -retainTMPSubcategories -outputFormat " & _
    """" & "penn" & """" & " " & """"
  Private Const STANFORD_TAGGED As String = " edu.stanford.nlp.parser.lexparser.LexicalizedParser -retainTMPSubcategories -outputFormat " & _
    """" & "penn" & """" & _
    " -sentences newline -tagSeparator _  " & _
    " -tokenizerFactory edu.stanford.nlp.process.WhitespaceTokenizer -tokenizerMethod newCoreLabelTokenizerFactory " & " " & """"
  Private Const STANFORD_TOKEN As String = " edu.stanford.nlp.process.PTBTokenizer "
  '----------------------------------------------------------------------------------------
  ' Name:       ConvertOneTxtToPsd()
  ' Goal:       Convert one file from source to destination
  ' History:
  ' 14-03-2013  ERK Created
  ' 07-01-2014  ERK Added "EnglishTranscribed"
  '----------------------------------------------------------------------------------------
  Public Function ConvertOneTxtToPsd(ByVal strSrcFile As String, ByVal strDstFile As String, _
                       ByVal strLanguage As String, Optional ByVal bDoAsk As Boolean = False) As Boolean
    Dim strCommand As String = "" ' The command shell
    Dim strDir As String = ""     ' Directory where we are working
    Dim strOut As String = ""     ' The output we receive
    Dim strErr As String = ""     ' The error we receive
    Dim strSrcMod As String = ""  ' Name of the modified source file
    Dim strErrFile As String = "" ' Name of error file

    Try
      ' Does the destination already exist
      If (IO.File.Exists(strDstFile)) Then
        ' Check the date of the destination compared to the source
        If (IO.File.GetLastWriteTime(strDstFile) > IO.File.GetLastWriteTime(strSrcFile)) Then
          ' Ask if we need to overwrite
          If (bDoAsk) Then
            Select Case MsgBox("Would you like to overwrite the existing psd file?", MsgBoxStyle.YesNoCancel)
              Case MsgBoxResult.No
                Return True
              Case MsgBoxResult.Cancel
                Return False
              Case Else
                ' Just continue!!
            End Select
          Else
            ' Do not ask, but overwrite!
            ' WAS: Return true
          End If
        End If
      End If
      ' Check and possibly download parser
      If (Not ParserCheck()) Then Logging("ConvertOneTxtToPsd: could not load parser") : Return False
      ' Action depends on language
      Select Case strLanguage
        Case "EnglishTranscribed"
          ' Try to follow a sentence-by-sentence approach
          If (Not TxtToPsdBySentence(strSrcFile, strDstFile, strLanguage)) Then Return False
        Case Else
          ' Think of a modified filename
          strSrcMod = IO.Path.GetDirectoryName(strSrcFile) & "\" & IO.Path.GetFileNameWithoutExtension(strSrcFile) & ".mod"
          ' Copy and improve source > modified
          If (Not ConvertOneTxtToMod(strSrcFile, strSrcMod, strLanguage)) Then Return False
          ' Get the directory
          strDir = IO.Path.GetDirectoryName(strSrcMod)
          strShort = IO.Path.GetFileNameWithoutExtension(strSrcMod)
          ' Try to convert one file
          Select Case strLanguage
            Case "German"
              strCommand = strStanfordShellIntroGerman & """" & strSrcMod & """"
            Case "English"
              strCommand = strStanfordShellIntroEnglish & """" & strSrcMod & """"
            Case "EnglishTranscribed"
              strCommand = strStanfordShellIntroEngTrans & """" & strSrcMod & """"
            Case Else
              Return False
          End Select
          If (Not ExecuteOneJavaCommand(strCommand, strDir, strOut, strErr)) Then
            ' Something went wrong
            Logging("modTxtToPsd/Convert/Execute error")
            Return False
          End If
          ' add the output into the destination file
          IO.File.WriteAllText(strDstFile, strOut)
          ' Make name of error file
          strErrFile = IO.Path.GetDirectoryName(strDstFile) & "\" & IO.Path.GetFileNameWithoutExtension(strDstFile) & ".err"
          IO.File.WriteAllText(strErrFile, strErr)
          ' Remove the MOD file
          IO.File.Delete(strSrcMod)
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modTxtToPsd/ConvertOneTxtToPsd error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       TxtToPsdBySentence()
  ' Goal:       Check if the parser is there, and if not: download it
  ' History:
  ' 14-03-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function TxtToPsdBySentence(ByVal strIn As String, ByVal strOut As String, ByVal strLanguage As String) As Boolean
    Dim strParse As String    ' Parse command
    Dim strToken As String    ' Tokenize command
    Dim strFileTxt As String  ' Temporary text file
    Dim strFileTok As String  ' Temporary file for a tokenized line
    Dim strFilePsd As String  ' Temporary file for a psd line
    Dim strText As String     ' Text of input file
    Dim arSent() As String    ' Array of sentences
    Dim strSent As String     ' One sentence
    Dim strTmpOut As String   ' Temporary output
    Dim strTmpErr As String   ' Temporary error
    Dim strDir As String      ' Directory
    Dim strFind As String     ' What we need to find
    Dim intPos As Integer     ' Position in string
    Dim chrThis As Char       ' one character
    Dim arLine1() As String   ' First tokenized line
    Dim arLineF() As String   ' Filtered line
    Dim arLineC() As String   ' Clean line
    Dim arInsertF() As String ' Indicates that this element has been filtered away or not
    Dim intSize As Integer = 0    ' Size of F
    Dim intI As Integer           ' Counter
    Dim intSnt As Integer = 0     ' Sentence
    Dim intBreak As Integer       ' Position of break
    Dim intK As Integer           ' Counter
    Dim srIn As StreamReader = Nothing  ' File I read from
    Dim srOut As StreamWriter = Nothing ' File I write to
    Dim mclThis As MatchCollection      ' Collection of hits

    Try
      ' Validate
      If (strIn = "") OrElse (Not IO.File.Exists(strIn)) Then Return False
      ' Determine temporary file names
      strFileTxt = GetSetDir() & "\temp.txt"
      strFileTok = GetSetDir() & "\temp.tok"
      strFilePsd = GetSetDir() & "\temp.psd"
      ' Determine the parse and tokenize commands
      strToken = strStanfordShellTokenize & "temp.txt"
      strParse = strStanfordShellIntroEnglish & """" & strFileTok & """"
      ' Initialisations
      strDir = IO.Path.GetDirectoryName(strIn)
      ' Start reading input
      srIn = New StreamReader(strIn)
      ' Start writing to output (Note: doesn't really work with UTF8)
      ' srOut = New StreamWriter(strOut, False, System.Text.Encoding.UTF8)
      srOut = New StreamWriter(strOut)
      ' Loop through file
      Do
        ' Check for interrupt
        If (bInterrupt) Then Return False
        intSnt += 1
        ' Show which input line we are
        ' Status("TxtToPsd sentence = " & intSnt)
        ' Read one line
        strText = srIn.ReadLine
        ' Have something?
        If (strText <> "") AndAlso (Not DoLike(strText, "<DID*|<S>|<F>|<h>")) Then
          ' Loop through text
          For intPos = 1 To strText.Length
            ' Get current character
            chrThis = Mid(strText, intPos, 1)
            ' Consider this element
            Select Case AscW(chrThis)
              Case 160, 65533
                strText = Left(strText, intPos - 1) & " " & Mid(strText, intPos + 1)
              Case &H2019
                strText = Left(strText, intPos - 1) & "'" & Mid(strText, intPos + 1)
              Case 8230
                strText = Left(strText, intPos - 1) & "..." & Mid(strText, intPos + 1)
              Case 231, 8216
                strText = Left(strText, intPos - 1) & "c" & Mid(strText, intPos + 1)
              Case 232, 233, 234, 235
                strText = Left(strText, intPos - 1) & "e" & Mid(strText, intPos + 1)
              Case 239
                strText = Left(strText, intPos - 1) & "i" & Mid(strText, intPos + 1)
              Case 228, 227, 226
                strText = Left(strText, intPos - 1) & "a" & Mid(strText, intPos + 1)
              Case Is >= 128
                ' Convert unknown characters into spaces

                Stop
                ' strText = Left(strText, intPos - 1) & "i" & Mid(strText, intPos + 1)
              Case Else
                ' Just continue
            End Select
          Next intPos
          ' (1a) Preprocessing of (em) etc
          strText = strText.Replace("(em)", "<em>")
          strText = strText.Replace("(mm)", "<mm>")
          strText = strText.Replace("(er)", "<er>")
          strText = strText.Replace("(eh)", "<eh>")
          strText = strText.Replace("(erm)", "<erm>")
          strText = strText.Replace("(hu)", "<hu>")
          strText = strText.Replace("(e)", "<eh>")
          strText = strText.Replace(":", "<P:>")
          strText = strText.Replace(".", "<P.>")
          strText = strText.Replace("=", "<P=>")
          'If (InStr(strText, ".") > 0) Then
          '  Stop
          'End If
          ' (1b) break sentences into smaller ones, based on sentence dividers like [and so] and [so]
          arSent = MySplitSent(strText)
          For intK = 0 To arSent.Length - 1
            ' Get this sentence
            strSent = Trim(arSent(intK))
            ' Check for empty sentences
            If (strSent <> "") Then
              ' (2) Tokenize this sentence
              IO.File.WriteAllText(strFileTxt, strSent)
              strTmpErr = "" : strTmpOut = ""
              ExecuteOneJavaCommand(strToken, GetSetDir, strTmpOut, strTmpErr, True)

              ' (3) Transform into array
              arLine1 = Split(strTmpOut, vbCrLf)
              ' (4) Junk away filtering
              ReDim arLineF(0 To 0) : intSize = 0
              For intI = 0 To arLine1.Length - 1
                If (Not DoLike(arLine1(intI), "</A>|</B>|<h>|<S>|<F>")) Then
                  ' Add to output
                  ReDim Preserve arLineF(0 To intSize)
                  arLineF(intSize) = arLine1(intI)
                  intSize += 1
                End If
              Next intI
              ReDim arInsertF(0 To intSize - 1)
              ' (5) Make a clean array
              ReDim arLineC(0 To 0) : intSize = 0
              For intI = 0 To arLineF.Length - 1
                'If (DoLike(arLineF(intI), "<.>|<=>|<:>")) Then
                '  Stop
                'End If
                If (Left(arLineF(intI), 1) = "<") Then
                  ' Get the text
                  arLineF(intI) = arLineF(intI).Replace("<", "")
                  arLineF(intI) = arLineF(intI).Replace(">", "")
                  arLineF(intI) = arLineF(intI).Replace("/", "")
                  arInsertF(intI) = Trim(arLineF(intI))
                Else
                  ' Indicate this is normal text
                  arInsertF(intI) = ""
                  ' Add to output
                  ReDim Preserve arLineC(0 To intSize)
                  arLineC(intSize) = arLineF(intI)
                  intSize += 1
                End If
              Next intI
              ' (6) Write array as file
              IO.File.WriteAllLines(strFileTok, arLineC)
              ' (7) Perform stanford parser on this line
              strTmpErr = "" : strTmpOut = ""
              ExecuteOneJavaCommand(strParse, GetSetDir, strTmpOut, strTmpErr, True)
              ' Check
              If (arLineC.Length <> arLineF.Length) Then
                ' Walk through all the entries in [arLineF]
                intPos = 0
                For intI = 0 To arLineF.Length - 1
                  ' Skip empty entries
                  If (arLineF(intI) <> "") Then
                    ' Do we need to insert this one?
                    If (arInsertF(intI) = "") Then
                      ' Find this entry
                      strFind = " " & arLineF(intI) & ")"
                      intPos = InStr(intPos + 1, strTmpOut, strFind)
                      If (intPos = 0) Then
                        ' is this sentence being skipped?
                        If (InStr(strTmpOut, "Sentence skipped") > 0) Then
                          ' Okay, make sure we continue with something else
                          Exit For
                        Else
                          Stop
                        End If
                      Else
                        intPos += strFind.Length
                      End If
                    Else
                      ' We need to insert it
                      If (intPos = 0) Then
                        ' Is there any output at all?
                        If (strTmpOut = "") Then strTmpOut = "(ROOT (S ))"
                        ' Insert at beginning: find first space after (ROOT (X
                        With Regex.Match(strTmpOut, "\(\w+\s+\(\w+\s+")
                          intPos = .Index + .Length
                        End With
                        ' Find the next space
                        'intPos = InStr(strTmpOut, " ")
                        'intPos = InStr(intPos + 1, strTmpOut, "(")
                        'intPos = InStr(intPos + 1, strTmpOut, " ")
                        strFind = "(META " & arInsertF(intI) & ") "
                        strTmpOut = Left(strTmpOut, intPos) & strFind & Mid(strTmpOut, intPos + 1)
                        intPos += strFind.Length
                      Else
                        ' Insert after previous intPos
                        ' Stop
                        ' Find this entry
                        ' strFind = " " & arLineF(intI) & ")"
                        'intPos = InStr(intPos + 1, strTmpOut, strFind)
                        strTmpOut = Left(strTmpOut, intPos - 1) & " (META " & arLineF(intI) & ") " & Mid(strTmpOut, intPos)
                        ' Stop
                      End If
                    End If
                  End If

                Next intI
              End If
              ' Add the output
              srOut.WriteLine(strTmpOut)
            End If

          Next intK



          'Else
          '  strText = ""
        End If
        '' Write the adapted line
        'If (strText <> "") Then srOut.WriteLine(strText)
      Loop Until strText Is Nothing
      ' Finish files
      srIn.Close()
      srOut.Close()
      ' Return success
      Return True
    Catch ex As Exception
      ' Finish files
      If (srIn IsNot Nothing) Then srIn.Close()
      If (srOut IsNot Nothing) Then srOut.Close()
      ' Give error
      HandleErr("modTxtToPsd/TxtToPsdBySentence error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       MySplitSent()
  ' Goal:       Try split sentences
  ' History:
  ' 09-01-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function MySplitSent(ByVal strText As String) As String()
    Dim mclThis As MatchCollection
    Dim arBack(0 To 0) As String

    Try
      ' ========== DEBUG -=============
      'If (InStr(strText, "but <er>") > 0) Then
      '  Stop
      'End If
      ' ================================
      mclThis = Regex.Matches(strText, " and so | so | but \<er\>| and \<er\>| and and | and \<P\.\>| and yet| and yes | yes again ")
      ReDim arBack(0 To mclThis.Count)
      If (mclThis.Count = 0) Then
        arBack(0) = strText
      Else
        ' We have some stuff
        For intK = 0 To mclThis.Count - 1
          ' Get this chunk
          If (mclThis.Count = 1) Then
            ' There IS one hit, which divides the matter into two!
            ' Debug.Print(mclThis.Item(0).Index)
            ' There is only one chunk
            arBack(0) = Left(strText, mclThis.Item(0).Index + 1)
            arBack(1) = Mid(strText, mclThis.Item(0).Index + 2)
          ElseIf (intK = 0) Then
            ' First part
            ' Debug.Print("First: [" & Left(strText, mclThis.Item(1).Index + 1) & "]")
            arBack(intK) = Left(strText, mclThis.Item(1).Index + 1)
          ElseIf (intK = mclThis.Count - 1) Then
            ' Last part
            ' Stop
            ' Debug.Print("Last: [" & Mid(strText, mclThis.Item(intK).Index + 1) & "]")
            arBack(intK) = Mid(strText, mclThis.Item(intK).Index + 1)
          Else
            ' Somewhere in between
            ' Stop
            ' Debug.Print("Mid: [" & Mid(strText, mclThis.Item(intK).Index + 1, mclThis.Item(intK + 1).Index - mclThis.Item(intK).Index) & "]")
            arBack(intK) = Mid(strText, mclThis.Item(intK).Index + 1, mclThis.Item(intK + 1).Index - mclThis.Item(intK).Index)
          End If
        Next intK
      End If
      ' Return the array
      Return arBack
    Catch ex As Exception
      ' Give error
      HandleErr("MySplitSent/TxtToPsdBySentence error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return Nothing
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       ConvertOneTxtToMod()
  ' Goal:       Check if the parser is there, and if not: download it
  ' History:
  ' 14-03-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function ConvertOneTxtToMod(ByVal strIn As String, ByVal strOut As String, ByVal strLanguage As String) As Boolean
    Dim strText As String     ' Text of input file
    Dim intPos As Integer     ' Position in string
    Dim chrThis As Char       ' one character
    Dim srIn As StreamReader = Nothing  ' File I read from
    Dim srOut As StreamWriter = Nothing ' File I write to

    Try
      ' Validate
      If (strIn = "") OrElse (Not IO.File.Exists(strIn)) Then Return False
      ' Start reading input
      srIn = New StreamReader(strIn)
      ' Start writing to output (Note: doesn't really work with UTF8)
      ' srOut = New StreamWriter(strOut, False, System.Text.Encoding.UTF8)
      srOut = New StreamWriter(strOut)
      ' Loop through file
      Do
        ' Read one line
        strText = srIn.ReadLine
        ' Have something?
        If (strText <> "") AndAlso (InStr(strText, "<DID") = 0) Then
          ' Check how to deal with < and > signs
          Select Case strLanguage
            Case "EnglishTranscribed"
              ' Attempt to translate xml and other signs of transcription
              strText = ConvertTranscription(strText)
            Case Else
              ' Convert symbols
              strText = strText.Replace("<", "&lt;")
              strText = strText.Replace(">", "&gt;")
          End Select
          ' Loop through text
          For intPos = 1 To strText.Length
            ' Get current character
            chrThis = Mid(strText, intPos, 1)
            ' Consider this element
            Select Case AscW(chrThis)
              Case 160, 65533
                strText = Left(strText, intPos - 1) & " " & Mid(strText, intPos + 1)
              Case &H2019
                strText = Left(strText, intPos - 1) & "'" & Mid(strText, intPos + 1)
              Case 8230
                strText = Left(strText, intPos - 1) & "..." & Mid(strText, intPos + 1)
              Case 231, 8216
                strText = Left(strText, intPos - 1) & "c" & Mid(strText, intPos + 1)
              Case 232, 233, 234
                strText = Left(strText, intPos - 1) & "e" & Mid(strText, intPos + 1)
              Case Is >= 128
                ' Convert unknown characters into spaces
                Stop
                ' strText = Left(strText, intPos - 1) & " " & Mid(strText, intPos + 1)
              Case Else
                ' Just continue
            End Select
          Next intPos
        End If
        ' Write the adapted line
        srOut.WriteLine(strText)
      Loop Until strText Is Nothing
      ' Finish files
      srIn.Close()
      srOut.Close()
      ' Return success
      Return True
    Catch ex As Exception
      ' Finish files
      If (srIn IsNot Nothing) Then srIn.Close()
      If (srOut IsNot Nothing) Then srOut.Close()
      ' Give error
      HandleErr("modTxtToPsd/ConvertOneTxtToMod error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       ConvertTranscription()
  ' Goal:       Convert the text...
  ' History:
  ' 07-01-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function ConvertTranscription(ByVal strIn As String) As String
    Dim intPos As Integer   ' Position in string
    Dim intEnd As Integer   ' Closing tag position
    Dim intI As Integer     ' Counter
    Dim intJ As Integer     ' Counter
    Dim intK As Integer     ' Counter
    Dim strFind As String   ' The string to look for
    Dim strEnd As String    ' Ending tag
    Dim strText As String   ' temporary text
    Dim arMeta() As String = {"A", "B", "S", "F", "er", "em", "mm", "laughs", "overlap", "overlap ", "coughs"}
    Dim arDel() As String = {"A", "B", "h", "F", "S"}
    Dim arRound() As String = {"er", "em", "mm", "erm"}
    Dim arStart() As String = {"<", "("}
    Dim arStartDel() As String = {"<", "(", "</"}
    Dim arEnd() As String = {">", "/>", " />", ")"}
    Dim strFw As String = "foreign"
    Dim arFw() As String

    Try
      ' Convert round brackets to lt/gt brackets
      For intI = 0 To arRound.Length - 1
        ' Define what we are looking for
        strFind = "(" & arRound(intI) & ")"
        ' Start looking...
        intPos = InStr(strIn, strFind)
        While (intPos > 0)
          ' Replace
          strIn = Left(strIn, intPos - 1) & "<" & arRound(intI) & ">" & _
                  Mid(strIn, intPos + strFind.Length + 1)
          ' Look afresh
          intPos = InStr(strIn, strFind)
        End While
      Next intI
      ' Find speaker beginnings and other stuff
      For intI = 0 To arMeta.Length - 1
        ' Look at beginnings and ends
        For intJ = 0 To arStart.Length - 1
          For intK = 0 To arEnd.Length - 1
            ' What are we looking for?
            strFind = arStart(intJ) & arMeta(intI) & arEnd(intK)
            ' Start looking...
            intPos = InStr(strIn, strFind)
            While (intPos > 0)
              ' Replace
              strIn = Left(strIn, intPos - 1) & arMeta(intI) & "_UH " & Mid(strIn, intPos + strFind.Length + 1)
              ' strIn = Left(strIn, intPos - 1) & arMeta(intI) & "_META " & Mid(strIn, intPos + strFind.Length + 1)
              ' strIn = Left(strIn, intPos - 1) & arMeta(intI) & "_FRAG " & Mid(strIn, intPos + strFind.Length + 1)
              ' strIn = Left(strIn, intPos - 1) & arMeta(intI) & "_X " & Mid(strIn, intPos + strFind.Length + 1)
              ' Look afresh
              intPos = InStr(strIn, strFind)
            End While
          Next intK
        Next intJ
      Next intI
      ' Delete what is necessary
      For intI = 0 To arDel.Length - 1
        ' Look at beginnings and ends
        For intJ = 0 To arStartDel.Length - 1
          For intK = 0 To arEnd.Length - 1
            ' What are we looking for?
            strFind = arStartDel(intJ) & arDel(intI) & arEnd(intK)
            ' Start looking...
            intPos = InStr(strIn, strFind)
            While (intPos > 0)
              ' Replace
              strIn = Left(strIn, intPos - 1) & Mid(strIn, intPos + strFind.Length + 1)
              ' Look afresh
              intPos = InStr(strIn, strFind)
            End While
          Next intK
        Next intJ
      Next intI
      ' Look for foreign word stretches
      strFind = "<" & strFw & ">" : strEnd = "</" & strFw & ">"
      intPos = InStr(strIn, strFind)
      While (intPos > 0)
        ' Look for closing tag
        intEnd = InStr(intPos, strIn, strEnd)
        If (intEnd > 0) Then
          ' Convert and delete and so forth
          arFw = Split(Trim(Mid(strIn, intPos + strFind.Length, intEnd - intPos - strFind.Length)), " ")
          strText = Join(arFw, "_FW ")
          strIn = Left(strIn, intPos - 1) & strText & "_FW " & Mid(strIn, intEnd + strEnd.Length + 1)
        Else
          intEnd = InStr(intPos, strIn, "<" & strFw & "/>")
          If (intEnd > 0) Then
            ' Convert and delete and so forth
            arFw = Split(Trim(Mid(strIn, intPos + strFind.Length, intEnd - intPos - strFind.Length)), " ")
            strText = Join(arFw, "_FW ")
            strIn = Left(strIn, intPos - 1) & strText & "_FW " & Mid(strIn, intEnd + strEnd.Length + 1)
          Else
            Stop
          End If
        End If
        ' Look for another FW
        intPos = InStr(strIn, strFind)
      End While
      ' Return the result
      Return strIn
    Catch ex As Exception
      ' Give error
      HandleErr("modTxtToPsd/ConvertTranscription error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return ""
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       ParserCheck()
  ' Goal:       Check if the parser is there, and if not: download it
  ' History:
  ' 14-03-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function ParserCheck() As Boolean
    Dim strSave As String ' Local destination

    Try
      ' Get destination directory
      strSave = GetSetDir()
      strStanfordGerman = strSave & "\" & STANFORD_GERMAN
      strStanfordEnglish = strSave & "\" & STANFORD_ENGLISH
      strStanfordParser = strSave & "\" & STANFORD_PARSER
      ' Check for the parser and the definitions
      If (Not IO.File.Exists(strStanfordParser)) Then
        ' Try to download it
        Status("Trying to download " & STANFORD_PARSER & "...")
        My.Computer.Network.DownloadFile(strHttpDir & STANFORD_PARSER, strStanfordParser, "", "", True, 5000, True)
        If (Not IO.File.Exists(strStanfordParser)) Then Return False
      End If
      ' Check for German definitions
      If (Not IO.File.Exists(strStanfordGerman)) Then
        ' Try to download it
        Status("Trying to download " & STANFORD_GERMAN & "...")
        My.Computer.Network.DownloadFile(strHttpDir & STANFORD_GERMAN, strStanfordGerman, "", "", True, 5000, True)
        If (Not IO.File.Exists(strStanfordGerman)) Then Return False
      End If
      ' Construct the parser command introduction
      strStanfordShellIntroGerman = STANFORD_INTRO & strStanfordParser & """" & STANFORD_LOC & strStanfordGerman & """" & " "
      ' Check for English definitions
      If (Not IO.File.Exists(strStanfordEnglish)) Then
        ' Try to download it
        Status("Trying to download " & STANFORD_ENGLISH & "...")
        My.Computer.Network.DownloadFile(strHttpDir & STANFORD_ENGLISH, strStanfordEnglish, "", "", True, 5000, True)
        If (Not IO.File.Exists(strStanfordEnglish)) Then Return False
      End If
      ' Construct the parser command introduction
      strStanfordShellIntroEnglish = STANFORD_INTRO & strStanfordParser & """" & STANFORD_LOC & strStanfordEnglish & """" & " "
      strStanfordShellIntroEngTrans = STANFORD_INTRO & strStanfordParser & """" & STANFORD_TAGGED & strStanfordEnglish & """" & " "
      strStanfordShellTokenize = STANFORD_INTRO & strStanfordParser & """" & STANFORD_TOKEN & " "
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modTxtToPsd/ParserCheck error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  ExecuteOneTxtToPsd
  ' Goal :  Execute one CS.jar command, and return when ready
  ' History:
  ' 13-09-2010  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function ExecuteOneJavaCommand(ByVal strArgs As String, ByVal strWorkDir As String, _
                  ByRef strOut As String, ByRef strErr As String, Optional ByVal bSilent As Boolean = False) As Boolean
    Dim objThis As New Process        ' Current process
    Dim bIsRunning As Boolean = True  ' Assume the process is running
    Dim intLast As Integer            ' Last shown from [colCs]

    Try
      ' Initialise error output
      strErr = "" : colCSoutput.Clear() : intCsOutLine = 0 : intCsErrLine = 0 : colCSerror.Clear()
      bCsError = False
      ' Set the properties
      With objThis.StartInfo
        ' Set the window style of this process
        .WindowStyle = ProcessWindowStyle.Minimized
        ' Supply the filename
        .FileName = "java.exe"
        ' Supply the arguments for this process
        .Arguments = strArgs
        ' Set the working directory correctly
        .WorkingDirectory = strWorkDir
        .CreateNoWindow = True
        ' Redirect the standard error output
        .UseShellExecute = False
        .RedirectStandardError = True
        .RedirectStandardOutput = True
        ' .RedirectStandardOutput = True
        ' Start filling up the "error" collection
        colCSerror.Add(.FileName & " " & .Arguments)
      End With
      ' Set our event handler to asynchronously read the output.
      AddHandler objThis.OutputDataReceived, AddressOf ReceiveCSoutput
      AddHandler objThis.ErrorDataReceived, AddressOf ReceiveCSerror
      Try
        ' Actually start the process
        objThis.Start()
        ' Start asynchronous reading on the output
        objThis.BeginOutputReadLine()
        ' Start asynchronous reading on the error
        objThis.BeginErrorReadLine()
      Catch ex As Exception
        ' If we cannot start up this process, then perhaps "java.exe" is not reachable
        MsgBox("modMain/ExecuteOneTxtToPsd: cannot run <stanford-parser>" & vbCrLf & vbCrLf & _
               "This could happen if <java.exe> is not reachable via your %PATH% system variable." & vbCrLf & _
               " Check this by opening a command prompt and typing:" & vbCrLf & _
               "     java.exe" & vbCrLf & _
               " If this is the problem, then adapt the Path system variable (see your computer's documentation).")
        bInterrupt = True
        Return False
      End Try
      ' Show we are loading
      If (Not bSilent) Then Status("Loading parser...")
      ' What has been shown already?
      intLast = colCSerror.Count - 1
      ' Find out when it finishes
      While (bIsRunning) And (Not bCsError)
        ' Allow events
        Application.DoEvents()
        '' Show progress in the line
        'frmMain.tsLine.Text = CStr(intCsOutLine) & "/" & CStr(intCsErrLine)
        ' Check for interrupt
        If (bInterrupt) Then
          ' Ask user for confirmation
          ' Double check anyway
          If (MsgBox("Would you really like to quit the current txt to psd conversion?", MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes) Then
            ' Make sure the subprocess is being quit
            objThis.Kill()
            ' Quit
            Return False
          End If
        End If
        ' Reset the interrupt flag
        bInterrupt = False
        ' Chack status
        bIsRunning = (Not objThis.WaitForExit(100))
        ' Show where we are
        If (Not bSilent) Then Status("TxtToPsd [" & strShort & "]: " & intCsOutLine)
        ' Also show the error messages in the [LOGGING] window
        ' Logging(colCSerror.Text)
        If (colCSerror.Count > 0) AndAlso (colCSerror.Count - 1 > intLast) Then
          If (InStr(colCSerror.Item(colCSerror.Count - 1), "virtual") > 0) Then Stop
          ' Stop
          Logging(colCSerror.Item(colCSerror.Count - 1))
          intLast = colCSerror.Count - 1
        End If
        '' Add standard error
        'strErr &= objThis.StandardError.ReadLine
      End While
      ' Make sure both Error text and Output get compiled
      strErr = colCSerror.Text
      strOut = colCSoutput.Text
      ' Check if there was an error
      If (bCsError) Then
        ' Add the output to the error stuff
        strErr &= "OUTPUT:" & vbCrLf & " =================================== " & vbCrLf & strOut
        'Else
        '  strOut = colCSoutput.Text
      End If
      '' Clear the TsLine output
      'frmMain.tsLine.Text = ""
      ' Return success 
      Return True
    Catch ex As Exception
      ' Warn user
      HandleErr("modTxtToPsd/ExecuteOneTxtToPsd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Make sure the subprocess is being quit
      If (objThis IsNot Nothing) Then objThis.Kill()
      ' Make sure we are interrupted!
      bInterrupt = True
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  ReceiveCSoutput
  ' Goal :  Read the output from the CS process
  ' History:
  ' 05-10-2010  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Private Sub ReceiveCSoutput(ByVal sendingProcess As Object, _
        ByVal outLine As DataReceivedEventArgs)
    Try
      ' Keep track of the number of lines
      intCsOutLine += 1
      ' Collect the command output.
      If (Not String.IsNullOrEmpty(outLine.Data)) Then
        ' Add it to the output collection
        colCSoutput.Add(outLine.Data)
      End If
    Catch ex As Exception
      ' Warn user
      HandleErr("modTxtToPsd/ReceiveCSoutput error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  ReceiveCSerror
  ' Goal :  Read the error from the CS process
  ' History:
  ' 05-10-2010  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Private Sub ReceiveCSerror(ByVal sendingProcess As Object, _
        ByVal outLine As DataReceivedEventArgs)
    Try
      ' Keep track of the number of lines
      intCsErrLine += 1
      ' Collect the command output.
      If (Not String.IsNullOrEmpty(outLine.Data)) Then
        ' Check for errors
        If (LCase(outLine.Data).Contains("error message")) Then
          ' Set the error flag
          bCsError = True
        End If
        ' Add it to the output collection
        colCSerror.Add(outLine.Data)
      End If
    Catch ex As Exception
      ' Warn user
      HandleErr("modTxtToPsd/ReceiveCSerror error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '----------------------------------------------------------------------------------------
  ' Name:       PrepareOneTxtToPsdx()
  ' Goal:       Prepare conversion of one file from source (con/ll) to destination (psdx)
  '             Alternative: tokenization of text to destination (psdx)
  ' History:
  ' 16-08-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function PrepareOneTxtToPsdx(ByVal strSrcFile As String, ByVal strDstFile As String, _
                       ByVal strLanguage As String, Optional ByVal bDoAsk As Boolean = False, _
                       Optional ByVal bFromConll As Boolean = True) As Boolean
    Dim strCommand As String = "" ' The command shell
    Dim strDir As String = ""     ' Directory where we are working
    Dim strOut As String = ""     ' The output we receive
    Dim strErr As String = ""     ' The error we receive
    Dim strSrcMod As String = ""  ' Name of the modified source file
    Dim strErrFile As String = "" ' Name of error file
    Dim strText As String = ""    ' String of file content
    Dim strBrk As String = ""     ' Break symbol used

    Try
      ' Does the destination already exist
      If (IO.File.Exists(strDstFile & ".psdx")) Then
        ' Check the date of the destination compared to the source
        If (IO.File.GetLastWriteTime(strDstFile & ".psdx") > IO.File.GetLastWriteTime(strSrcFile)) Then
          ' Ask if we need to overwrite
          If (bDoAsk) Then
            Select Case MsgBox("Would you like to overwrite the existing PSDX file?", MsgBoxStyle.YesNoCancel)
              Case MsgBoxResult.No
                Return True
              Case MsgBoxResult.Cancel
                Return False
              Case Else
                ' Just continue!!
            End Select
          Else
            ' Do not ask, but overwrite!
            ' WAS: Return true
          End If
        End If
      End If
      ' Think of a modified filename
      strSrcMod = IO.Path.GetDirectoryName(strSrcFile) & "\" & IO.Path.GetFileNameWithoutExtension(strSrcFile) & ".mod"
      ' Copy and improve source > modified
      If (Not ConvertOneTxtToMod(strSrcFile, strSrcMod, strLanguage)) Then Return False
      ' Get the directory
      strDir = IO.Path.GetDirectoryName(strSrcMod)
      strShort = GetShortFileName(strSrcMod)
      ' Is it conversion from CONLL?
      If (bFromConll) Then
        ' Try to convert one MOD file into a CONLL dependency file
        If (Not ConvertOneConllxToPsdx(strSrcFile, strDstFile & ".psdx", strLanguage)) Then Return False
      Else
        ' This is tokenization of an existing text...
        If (Not ConvertOneTxtToPsdx(strSrcFile, strDstFile & ".psdx", strLanguage)) Then Return False
      End If
      ' Remove the MOD file
      IO.File.Delete(strSrcMod)
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modTxtToPsd/ConvertOneTxtToPsdx error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
End Module
