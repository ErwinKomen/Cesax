Imports System.Text.RegularExpressions
Module modTree
  ' =========================================== LOCAL VARIABLE =================================================
  Private rgChunk As New Regex("\s\w$\/\w$(\s|\))")     ' Chunk parsing
  Private rgLbr As New Regex("\(")                      ' Left brackets
  Private rgRbr As New Regex("\)")                      ' Right brackets
  Private loc_strSpace As String = " " & vbTab & vbCrLf ' Spaces
  Private loc_strEnd As String = " )" & vbTab & vbCrLf  ' Spaces
  ' ============================================================================================================
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  DoExport
  ' Goal :  Export coreference information of type [strExpType] to the indicated file [strOutFile]
  ' History:
  ' 01-07-2010  ERK Created
  ' 11-10-2012  ERK Added the IS-output component from Cesac2008 (topic guesser)
  ' 11-04-2013  ERK Added "Irat" -- interrater agreement statistics (from Cesac)
  ' 24-04-2023  ERK Added CoNLL-x dependency output
  ' ----------------------------------------------------------------------------------------------------------
  Public Function DoExport(ByVal strExpType As String, ByVal strOutFile As String) As Boolean
    Dim strBack As String = ""  ' Output gathered

    Try
      ' Recursively put the output into [strBack]
      If (Not TravTree(strExpType, strBack)) Then
        ' Was interrupt pressed?
        If (Not bInterrupt) Then
          ' An error has occurred
          Status("DoExport: error in determining the coreference rating")
        End If
        ' Return failure
        Return False
      End If
      ' What output form is needed?
      Select Case strExpType
        Case "PsdOutput", "PsdSimple", "PsdPTB", "PdeOutput", "ISoutput", "PosSimple", _
             "PosOrg", "Token", "Rating", "DepCoNLL-X", "ConLLpos", "ConLLXpos", _
             "PosPTB", "Folia"
          ' Save result in output file in UTF8 form - which is still recognized by Excel
          IO.File.WriteAllText(strOutFile, strBack, System.Text.Encoding.UTF8)
        Case "Unicode"
          ' Save result in output file in Unicode form
          IO.File.WriteAllText(strOutFile, strBack, System.Text.Encoding.Unicode)
      End Select
      ' Show we are ready
      Status("Export " & strExpType & " was saved to: " & strOutFile)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modTree/DoExport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  ChunkToPsd
  ' Goal :  Convert chunk-psd to real psd
  ' Note :  This assumes an additional (ROOT is needed around chunks...
  ' History:
  ' 05-11-2013  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Private Function ChunkToPsd(ByVal strInFile As String, ByVal strOutFile As String) As Boolean
    Dim strLine As String         ' One line
    Dim strOut As String          ' Output
    Dim strId As String = ""      ' ID of sentence
    Dim strNext As String         ' Verification next letter
    Dim strWord As String         ' The actual word
    Dim strWork As String         ' Working string
    Dim colWord As New StringColl ' Word start and end
    Dim colPos As New StringColl  ' POS start and end
    Dim regWord As Match
    Dim intSlash As Integer       ' Position of slash
    Dim intBr As Integer = 0      ' Number of brackets
    Dim intPos As Integer         ' Position
    Dim intI As Integer           ' Counter
    Dim intLine As Integer = 0    ' Current line
    Dim rdThis As IO.StreamReader
    Dim wrThis As IO.StreamWriter

    Try
      ' Validate
      If (strInFile = "") OrElse (Not IO.File.Exists(strInFile)) OrElse (strOutFile = "") Then Return False
      ' Read input file
      rdThis = New IO.StreamReader(strInFile)
      ' Start output file
      wrThis = New IO.StreamWriter(strOutFile, False, System.Text.Encoding.UTF8)
      ' Determine start of ID 
      strId = Left(IO.Path.GetFileNameWithoutExtension(strInFile).Replace(" ", "_"), 6) & "_"
      ' Read lines
      While (Not rdThis.EndOfStream)
        ' Read one line
        strLine = rdThis.ReadLine
        ' Initialise
        colWord.Clear() : colPos.Clear()
        ' Parse this line
        intSlash = InStr(strLine, "/")
        While (intSlash > 0)
          ' Additional check: whatever follows a / must be a letter
          strNext = Mid(strLine, intSlash + 1, 1)
          If (Regex.IsMatch(strNext, "\w")) Then
            ' Find start of word
            intPos = intSlash - 1
            ' Skip until a space appears
            While (InStr(loc_strSpace, Mid(strLine, intPos, 1)) = 0) AndAlso (intPos > 1)
              intPos -= 1
            End While
            ' Mark the word start and end
            colWord.Add(intPos + 1, intSlash)
            ' Get the POS
            intPos = intSlash + 1
            While (InStr(loc_strEnd, Mid(strLine, intPos, 1)) = 0) AndAlso (intPos <= strLine.Length)
              intPos += 1
            End While
            ' Mark the POS start and en
            colPos.Add(intSlash + 1, intPos)
          Else
            ' Replace this slash by a space
            strLine = Left(strLine, intSlash - 1) & " " & Mid(strLine, intSlash + 1)
          End If
          ' Look for next one
          intSlash = InStr(intSlash + 1, strLine, "/")
        End While
        ' Make a new line-string
        If (colWord.Count = 0) Then
          strOut = strLine
        Else
          intPos = 1 : strOut = ""
          For intI = 0 To colWord.Count - 1
            ' Validate
            If (colWord.Item(intI) > intPos) Then
              ' add what comes before the word start
              strOut &= Mid(strLine, intPos, colWord.Item(intI) - intPos)
              ' Get the word
              strWord = Mid(strLine, colWord.Item(intI), colWord.Exmp(intI) - colWord.Item(intI))
              ' Check for *initial* punctuation
              If (strWord.Length > 1) Then
                strWork = Left(strWord, 1)
                If (Not Regex.IsMatch(strWork, "(\w|')")) Then
                  ' =========== DEBUG ================
                  'Debug.Print("ChunkToPsd non letter = x" & Hex(AscW(strWork)))
                  ' ==================================
                  ' Check
                  If (Hex(AscW(strWork)) <> "FEFF") Then
                    ' Get the initial punctuation out
                    Select Case strWork
                      Case "("
                        strWork = "-LRB-"
                      Case ")"
                        strWork = "-RRB-"
                    End Select
                    strOut &= " (, " & strWork & ") "
                  End If
                  ' Adapt the word...
                  strWord = Mid(strWord, 2)
                End If
              End If
              ' Check for *final* punctuation
              If (strWord.Length > 1) Then
                strWork = Right(strWord, 1)
                If (Regex.IsMatch(strWork, "(\w|')")) Then
                  ' This is a normal word
                  strWork = ""
                Else
                  ' Get the initial punctuation out
                  Select Case strWork
                    Case "("
                      strWork = "-LRB-"
                    Case ")"
                      strWork = "-RRB-"
                  End Select
                  strWork = " (, " & strWork & ") "
                  strWord = Left(strWord, strWord.Length - 1)
                End If
              Else
                strWork = ""
              End If
              ' Get the word itself
              strWord = strWord.Replace("(", "[") : strWord = strWord.Replace(")", "]")
              strOut &= " (" & Mid(strLine, colPos.Item(intI), colPos.Exmp(intI) - colPos.Item(intI)) & " " & _
                        strWord & ") "
              ' Add any final punctuation
              strOut &= strWork
            Else
              Stop
            End If
            ' Move on with POS
            intPos = colPos.Exmp(intI)
          Next intI
          ' Add the part that comes afterwards
          strOut &= Mid(strLine, intPos)
        End If
        ' Change tabs for spaces
        strOut = strOut.Replace(vbTab, " ")
        strOut = strOut.Replace(vbCr, " ")
        ' Keep track of the amount of brackets
        intBr += rgLbr.Matches(strOut).Count - rgRbr.Matches(strOut).Count
        ' Check if this starts a sentence
        intPos = InStr(strOut, "(S")
        If (intPos > 0) Then
          ' Does this start a sentence?
          If ((strOut.Length = intPos + 2) OrElse (InStr(loc_strSpace, Mid(strOut, intPos + 2, 1)) > 0)) Then
            ' (1) Add a start node
            strOut = "(ROOT " & strOut
            ' (2) Keep count of line number
            intLine += 1
          ElseIf (InStr(strOut, "(SEF") > 0) Then
            ' Okay, continue
          Else
            ' Stop
            ' Double check!
          End If
        End If
        ' Are we back to zero?
        If (intBr = 0) Then
          ' Check if this is an empty node
          If (intPos > 0) AndAlso (rgRbr.Matches(strOut).Count = 1) Then
            ' This is an empty node!
            strOut = "" : intLine -= 1
          Else
            ' Add information to finish it off
            strOut &= " (ID " & strId & Format(intLine, "0000") & "))"
          End If
        End If
        ' Write the line
        wrThis.WriteLine(strOut)
      End While
      ' Close the files
      wrThis.Close()
      rdThis.Close()
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modTree/ChunkToPsd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  DoImport
  ' Goal :  Import information of type [strExpType] to the indicated file [strInFile]
  ' History:
  ' 01-07-2010  ERK Created
  ' 05-11-2013  ERK Added ChunkParsed
  ' ----------------------------------------------------------------------------------------------------------
  Public Function DoImport(ByVal strImpType As String, ByVal strInFile As String) As Boolean
    Dim strBack As String = ""  ' Output gathered
    Dim strDst As String = ""   ' Destination file name
    Dim strTmp As String = ""   ' Temporary file
    Dim filSave As New System.Windows.Forms.SaveFileDialog

    Try
      ' Action depends on the import type
      Select Case strImpType
        Case "Treebank", "ChunkParsed"
          ' Prepare a destination file
          strDst = IO.Path.GetDirectoryName(strInFile) & "\" & IO.Path.GetFileNameWithoutExtension(strInFile) & ".psdx"
          ' Get destination file name
          ' Get the filename from the user
          With filSave
            ' The initial directory is the one we already know
            .InitialDirectory = strWorkDir
            ' Set the default extention
            .DefaultExt = "psdx"
            ' Set the default filter
            .Filter = "Psd XML files|*.psdx"
            ' Assign default file name to the FileSave dialog
            'strCRPfile = strFile
            .FileName = strDst
            ' Show the actual dialog to the user
            Select Case .ShowDialog()
              Case Windows.Forms.DialogResult.OK, Windows.Forms.DialogResult.Yes
                ' Get the filename that the user has finally selected
                strDst = .FileName
              Case Else
                ' Aborted, so exit
                Return False
            End Select
          End With
          ' Do we need preprocessing?
          If (strImpType = "ChunkParsed") Then
            ' This is PSD, but then chunk-parsed; it needs a preprocessing step
            strTmp = GetDocDir() & "\" & IO.Path.GetFileNameWithoutExtension(IO.Path.GetTempFileName) & ".psd"
            If (Not ChunkToPsd(strInFile, strTmp)) Then Return False
            ' Convert the pre-processed file into PSDX
            If (Not ConvertOnePsdToPsdx(strTmp, strDst, "", False)) Then Return False
            ' Remove the temporary file
            IO.File.Delete(strTmp)
          Else
            ' Convert from PSD to PSDX
            If (Not ConvertOnePsdToPsdx(strInFile, strDst, "", False)) Then
              ' Failure
              Return False
            End If
          End If
          ' Do FileOpen from this temporary place
          If (Not LoadOneFile(strDst)) Then
            ' Failure
            Return False
          End If
        Case "PdeInput"
          ' Read in the translation
          If (Not ReadTranslation(strInFile)) Then
            ' Failure
            Return False
          End If
        Case "ChainDict"
          ' Read the chain dictionary information into our own system
          strChainDict = strInFile
          ' Show we can't do it yet
          MsgBox("modTree/DoImport: cannot handle reading chain dictionary yet")
          Return False
        Case "ConLLX"
          ' Convert CONLLX format to PSDX without further ado
          strDst = IO.Path.GetDirectoryName(strInFile) & "\" & _
            IO.Path.GetFileNameWithoutExtension(strInFile) & ".psdx"
          If (Not ConvertOneConllxToPsdx(strInFile, strDst, "ask")) Then Return False
          ' Load the temporary file
          If (Not LoadOneFile(strDst)) Then Return False
        Case "Folia", "folia"
          ' Convert folia to psdx without further ado
          strDst = IO.Path.GetDirectoryName(strInFile) & "\" & _
            IO.Path.GetFileNameWithoutExtension(strInFile) & ".psdx"
          If (Not ConvertOneFoliaToPsdx(strInFile, strDst)) Then Return False
          ' Load the temporary file
          If (Not LoadOneFile(strDst)) Then Return False
      End Select
      ' Show we are ready
      Status("Importing " & strImpType & " was completed from: " & strInFile)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modTree/DoImport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoCurrentDep
  ' Goal:   Add dependency relations to the current file
  ' History:
  ' 24-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function DoCurrentDep() As Boolean
    Dim strBack As String = ""
    Try
      Return TravTree("DepChe", strBack)
    Catch ex As Exception
      ' Show error
      HandleErr("modTree/DoCurrentDep error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  TravTree
  ' Goal :  Call a recursively working function to travel the tree...
  ' History:
  ' 01-07-2010  ERK Created
  ' 11-10-2012  ERK Added the IS-output component from Cesac2008 (topic guesser)
  ' ----------------------------------------------------------------------------------------------------------
  Private Function TravTree(ByVal strAction As String, ByRef strBack As String) As Boolean
    Dim objTotal As New StringColl  ' All the different nodes
    Dim ndxForest As Xml.XmlNode    ' Currently looked at forest item
    Dim intI As Integer             ' Counter
    Dim intPtc As Integer           ' Percentage

    Try
      ' Validate
      If (strAction = "") Then Return False
      If (pdxList Is Nothing) Then Return False
      ' Initialisations
      objTotal.Clear()
      Select Case strAction
        Case "PdeOutput"
          ' Add header
          objTotal.Add("<TEI><forestGrp>")
        Case "ISoutput"
          ' Add header
          objTotal.Add("<html><body><h2>Topic guessing (Information Structure)</h2>")
          ' Beginning of HTML file for information structure purposes
          objTotal.Add("<table cellSpacing=0 cellPadding=0 border=1>")
          objTotal.Add("<tr><td><b>Loc</b></td>" & _
                             "<td><b>Type</b></td>" & _
                             "<td><b>Why</b></td>" & _
                             "<td><b>TopPos</b></td>" & _
                             "<td><b>TopCat</b></td>" & _
                             "<td><b>Topic</b></td>" & _
                             "<td><b>Syntax</b></td>" & _
                             "</tr>")
        Case "Rating"
          ' Set headings of the columns for the CSV file
          objTotal.Add("Location;Reference;Type")
      End Select
      ' Go through the list
      For intI = 0 To pdxList.Count - 1
        ' Show where we are
        intPtc = (100 * (intI + 1)) \ pdxList.Count
        Status("Processing " & strAction & " " & intPtc & "%", intPtc)
        ' Also check for interrupt
        If (bInterrupt) Then Return False
        ' Clear the string
        strBack = ""
        ' Get the current forest
        ndxForest = pdxList(intI)
        Select Case strAction
          Case "DepCoNLL-X"
            objTotal.Add(GetCoNLL(ndxForest))
          Case "ConLLXpos", "ConLLpos"
            objTotal.Add(GetConLLpos(ndxForest))
          Case Else
            ' Perform the actual action
            If (Not TravNode(pdxList(intI), strAction, strBack)) Then
              ' There is a problem!!!
              Status("modTree/Travtree: there is a problem at line " & intI)
              Return False
            End If
            strBack = MyTrim(strBack)
            ' Is there any return?
            If (Trim(strBack) <> "") Then
              ' Store the output
              objTotal.Add(strBack)
              ' strBack = MyTrim(strBack)
            End If
        End Select
      Next intI
      Select Case strAction
        Case "PdeOutput"
          ' Add footer
          objTotal.Add("</forestGrp></TEI>")
        Case "ISoutput"
          ' Add header
          objTotal.Add("</table></body></html>")
      End Select
      ' Get the output on the correct place
      strBack = objTotal.Text
      ' Show we finish positively
      Status("Done")
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modTree/TravTree error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function
End Module
