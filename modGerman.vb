Module modGerman
  ' ===================================== GLOBAL VARIABLES ==========================================================
  ' ===================================== LOCAL VARIABLES ===========================================================
  Private strStanfordParser As String = ""  ' Location of the stanford parser
  Private strStanfordGerman As String = ""  ' German definition file
  Private strStanfordEnglish As String = "" ' Englishdefinition file
  Private strHttpDir As String = "http://erwinkomen.ruhosting.nl/software/stp/"
  Private strStanfordShellIntroGerman As String = ""
  Private strStanfordShellIntroEnglish As String = ""
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
  Private Const STANFORD_INTRO As String = "-mx800m -classpath " & """"
  Private Const STANFORD_LOC As String = " edu.stanford.nlp.parser.lexparser.LexicalizedParser -retainTMPSubcategories -writeOutputFiles  -outputFormat " & _
    """" & "penn" & """" & " " & """"
  '----------------------------------------------------------------------------------------
  ' Name:       ConvertOneTxtToPsd()
  ' Goal:       Convert one file from source to destination
  ' History:
  ' 14-03-2013  ERK Created
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
      ' Think of a modified filename
      strSrcMod = IO.Path.GetDirectoryName(strSrcFile) & "\" & IO.Path.GetFileNameWithoutExtension(strSrcFile) & ".mod"
      ' Copy and improve source > modified
      If (Not ConvertOneTxtToMod(strSrcFile, strSrcMod)) Then Return False
      ' Check and possibly download parser
      If (Not ParserCheck()) Then Logging("ConvertOneTxtToPsd: could not load parser") : Return False
      ' Get the directory
      strDir = IO.Path.GetDirectoryName(strSrcMod)
      strShort = IO.Path.GetFileNameWithoutExtension(strSrcMod)
      ' Try to convert one file
      Select Case strLanguage
        Case "German"
          strCommand = strStanfordShellIntroGerman & """" & strSrcMod & """"
        Case "English"
          strCommand = strStanfordShellIntroEnglish & """" & strSrcMod & """"
        Case Else
          Return False
      End Select
      If (Not ExecuteOneTxtToPsd(strCommand, strDir, strOut, strErr)) Then
        ' Something went wrong
        Logging("modGerman/Convert/Execute error")
        Return False
      End If
      ' add the output into the destination file
      IO.File.WriteAllText(strDstFile, strOut)
      ' Make name of error file
      strErrFile = IO.Path.GetDirectoryName(strDstFile) & "\" & IO.Path.GetFileNameWithoutExtension(strDstFile) & ".err"
      IO.File.WriteAllText(strErrFile, strErr)
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modGerman/ConvertOneTxtToPsd error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       ConvertOneTxtToMod()
  ' Goal:       Check if the parser is there, and if not: download it
  ' History:
  ' 14-03-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function ConvertOneTxtToMod(ByVal strIn As String, ByVal strOut As String) As Boolean
    Dim strText As String ' Text of input file
    Dim intPos As Integer ' Position in string
    Dim chrThis As Char   ' one character

    Try
      ' Validate
      If (strIn = "") OrElse (Not IO.File.Exists(strIn)) Then Return False
      ' Read input file
      strText = IO.File.ReadAllText(strIn)
      ' Convert symbols
      strText = strText.Replace("<", "&lt;")
      strText = strText.Replace(">", "&gt;")
      ' Loop through text
      For intPos = 1 To strText.Length
        ' Get current character
        chrThis = Mid(strText, intPos, 1)
        ' Consider this element
        Select Case AscW(chrThis)
          Case 160, 65533
            strText = Left(strText, intPos - 1) & " " & Mid(strText, intPos + 1)
          Case Is >= 128
            ' Convert unknown characters into spaces
            strText = Left(strText, intPos - 1) & " " & Mid(strText, intPos + 1)
            ' Old code:
            'HandleErr("ConvertOneTxtToMod: there is an unexpected upper-ascii element:" & _
            '          "char=[" & chrThis & "] hex=[" & Hex(AscW(chrThis)) & "] dec=[" & AscW(chrThis) & "]")
            'bInterrupt = True
            'Return False
          Case Else
            ' Just continue
        End Select
      Next intPos
      'strText = strText.Replace(ChrW(160), " ")
      ' Write file
      IO.File.WriteAllText(strOut, strText)
      ' IO.File.WriteAllText(strOut, strText)
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modGerman/ConvertOneTxtToMod error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False

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
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modGerman/ParserCheck error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
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
  Private Function ExecuteOneTxtToPsd(ByVal strArgs As String, ByVal strWorkDir As String, _
                  ByRef strOut As String, ByRef strErr As String) As Boolean
    Dim objThis As New Process        ' Current process
    Dim bIsRunning As Boolean = True  ' Assume the process is running

    Try
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
      End With
      ' Initialise error output
      strErr = "" : colCSoutput.Clear() : intCsOutLine = 0 : intCsErrLine = 0 : colCSerror.Clear()
      bCsError = False
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
      Status("Loading parser...")
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
        Status("TxtToPsd [" & strShort & "]: " & intCsOutLine)
        ' Also show the error messages in the [tbGenSource] window
        With frmMain.tbGenSource
          ' Append the text to the LOG textbox in the General page -- add a CRLF
          .Text = colCSerror.Text
          ' Zorgen dat we naar het einde scrollen
          '.SelectionStart = .TextLength
          '.ScrollToCaret()
        End With
        '' Add standard error
        'strErr &= objThis.StandardError.ReadLine
      End While
      'strErr = colCSerror.Text
      'strOut = colCSoutput.Text
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
      HandleErr("modGerman/ExecuteOneTxtToPsd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
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
      HandleErr("modGerman/ReceiveCSoutput error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
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
      HandleErr("modGerman/ReceiveCSerror error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Module
