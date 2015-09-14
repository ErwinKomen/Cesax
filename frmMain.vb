Imports System.Xml
Imports Netron
Imports Netron.Lithium
Public Class frmMain
  ' ================================= GLOBAL VARIABLES =========================================================
  Private tabHidden As New TabControl
  ' ================================= LOCAL VARIABLES =========================================================
  Private bDirty As Boolean = False     ' Whether something needs to be saved or not...
  Private bResDirty As Boolean = False  ' Whether results need to be changed
  Private bResRecord As Boolean = False ' One record of Results has changed
  Private bResStatus As Boolean = False ' Status of record needs to be changed
  Private bRefDirty As Boolean = False  ' Whether RefChain needs to be changed
  Private bResSel As Boolean = False    ' Results selection changes
  Private bRevision As Boolean = False  ' Whether a revision note should be added or not
  Private bInit As Boolean = False      ' Initialisation flag
  Private bPsdxInit As Boolean = False  ' Whether a PSD file has been loaded
  Private bEdtLoad As Boolean = True    ' Whether the editor is loading
  Private bTransBusy As Boolean = False ' Translation is busy
  Private bAskReady As Boolean = False  ' Whether user has given an anwer
  Private bTargetInit As Boolean = False  ' Whether the listview for the target is initialized or not
  Private bTreeMove As Boolean = False    ' Inside or outside the treemove 
  Private bLabelSelect As Boolean = False ' Busy getting a label
  Private bEditMode As Boolean = False    ' Mode to edit the tree in the "Tree" tab page
  Private intLabelState As DialogResult   ' State of the label
  Private loc_intDstId As Integer = -1    ' Answer given by user
  Private strReportName As String = "MissingPronouns"   ' Default report name
  Private WithEvents tbPde As New RichTextBox           ' A RICH textbox that will hold the translation
  Private intZoom As Integer = 100        ' Number of zoom
  Private bDbFilter As Boolean = False    ' DbFilter is busy
  Private intCurrentResId As Integer = -1 ' ID of currently selected DbResult record
  Private arFtLabel() As Label            ' Feature label array
  Private arFtTxtBox() As TextBox         ' Textbox

  ' =============================== LOCAL CONSTANTS ======================================
  Private Const HELP_CS As String = "http://corpussearch.sourceforge.net/"
  Private Const HELP_XQ As String = "http://www.w3.org/XML/Query/"
  Private Const HELP_IEXPLORE As String = "C:\Program Files\Internet Explorer\iexplore.exe"
  Private Const QUICK_REFERENCE As String = "CesaxQuickStart.htm"
  Private Const HELP_MANUAL As String = "http://erwinkomen.ruhosting.nl/software/Cesax/Cesax_Manual.pdf"
  Private Const HELP_HTML As String = "http://erwinkomen.ruhosting.nl/software/Cesax/Cesax_Manual.htm"
  ' ============================================================================================================
  '---------------------------------------------------------------------------------------------------------
  ' Name:       EnglishPeriod()
  ' Goal:       Retrieve the English period user wants to work with
  ' History:
  ' 29-11-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Function EnglishPeriod(ByRef strAbbr As String) As Boolean
    Try
      With frmEnglish
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Establish the language
            strAbbr = .Language
          Case Else
            Return False
        End Select
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/EnglishPeriod error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       CheckPeriodFile()
  ' Goal:       Check whether there is a valid period information file
  ' History:
  ' 12-07-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Function CheckPeriodFile() As Boolean
    Dim bNeedPer As Boolean = False   ' Whethe the period definition file should be determined

    Try
      ' Check if the PeriodDefinition file is known -- otherwise don't allow File/Open
      If (strPeriodFile = "") Then
        bNeedPer = True
      ElseIf (Not IO.File.Exists(strPeriodFile)) Then
        bNeedPer = True
      End If
      If (bNeedPer) Then
        ' Do we have internet connection?
        If (My.Computer.Network.IsAvailable) Then
          ' Ask the user whether he/she wants to download one herself
          With frmLanguage
            Select Case .ShowDialog
              Case Windows.Forms.DialogResult.OK
                ' Start the period file directory
                strPeriodFile = GetSetDir() & "\"
                ' Try to download the period file
                Select Case .Language
                  Case "English"
                    strPeriodFile &= "EnglishPeriods.xml"
                  Case "Chechen"
                    strPeriodFile &= "ChechenPeriods.xml"
                  Case "Dutch"
                    strPeriodFile &= "CgnPeriod.xml"
                  Case "Welsh"
                    strPeriodFile &= "WelshPeriod.xml"
                  Case "German"
                    strPeriodFile &= "GermanPeriod.xml"
                  Case "Self"
                    ' User wants to do it himself
                    Status("Supply the period file in Tools/Settings/General")
                    Logging("Supply the period file in Tools/Settings/General")
                    strPeriodFile = ""
                    Return True
                End Select
                ' Check if it, perchance, already exists here
                If (Not IO.File.Exists(strPeriodFile)) Then
                  ' Try to download it
                  If (Not TryLoadFile(strPeriodFile)) Then
                    ' Warn user
                    Status("Unable to download period file for [" & .Language & "]")
                    Logging("Unable to download period file for [" & .Language & "]")
                    ' Use default period file instead
                    If (Not CreateDefaultPeriodFile()) Then Return False
                    ' We're fine
                    Return True
                  End If
                End If
                ' Set the period file setting
                SetTableSetting(tdlSettings, "PeriodDefinition", strPeriodFile)
                ' Save the settings
                XmlSaveSettings(strSetFile)
                ' Now the periods need to be read!!
                TryReadPeriods()
              Case Else
                ' Warn the user
                Status("Supply the period file in Tools/Settings/General")
                Logging("Added a new period definition file." & vbCrLf & _
                        "If you make use of existing periods, please supply the period file in Tools/Settings/General")
                If (Not CreateDefaultPeriodFile()) Then Return False
                ' We're fine
                Return True
            End Select
          End With
        Else
          ' Make use of default period file
          Status("Supply the period file in Tools/Settings/General")
          Logging("Added a new period definition file." & vbCrLf & _
                  "If you make use of existing periods, please supply the period file in Tools/Settings/General")
          If (Not CreateDefaultPeriodFile()) Then Return False
          ' We're fine
          Return True
        End If

        '' Warn theuser
        'MsgBox("You have not supplied a valid PeriodFile (location)." & vbCrLf & _
        '       "Go to Tools/Settings, supply a valid period file, and try File/Open again.")
        'Return False
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/CheckPeriodFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CreateDefaultPeriodFile
  ' Goal:   Create a default period file
  ' History:
  ' 01-11-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function CreateDefaultPeriodFile() As Boolean
    Dim pdxPer As New XmlDocument     ' Document
    Dim ndxThis As XmlNode            ' PeriodInfo

    Try
      ' Create a default period definition file
      strPeriodFile = GetSetDir() & "\_DefaultPeriods.xml"
      ' Create this file
      pdxPer.LoadXml(strDefPeriodText)
      ' Add a PeriodInfo header
      SetXmlDocument(pdxPer)
      ndxThis = AddXmlChild(pdxPer.SelectSingleNode("./descendant::PeriodList"), "PeriodInfo", _
                            "File", strPeriodFile, "attribute", _
                            "Comment", "Automatically created", "attribute", _
                            "Goal", "Default", "attribute", _
                            "Created", "", "attribute", _
                            "Changed", "", "attribute")
      ' Save the xml document
      pdxPer.Save(strPeriodFile)
      ' Now the periods need to be read!!
      TryReadPeriods()
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/CreateDefaultPeriodFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   TryLoadFile
  ' Goal:   See if the indicated file exists, and if not, try download it
  ' History:
  ' 19-09-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function TryLoadFile(ByRef strFile As String) As Boolean
    Dim strDstDir As String         ' Destination directory
    Dim strName As String           ' Name of this file
    Dim intI As Integer             ' Counter
    Dim arDefFile() As String = {"http://erwinkomen.ruhosting.nl/software/CorpusStudio/EnglishPeriods.xml", _
                                 "http://erwinkomen.ruhosting.nl/software/CorpusStudio/ChechenPeriods.xml", _
                                 "http://erwinkomen.ruhosting.nl/software/CorpusStudio/CgnPeriod.xml"}

    Try
      ' Validate
      If (strFile = "") Then Return False
      ' A period file is defined -- see if it exists
      If (Not IO.File.Exists(strFile)) Then
        ' Get the directory itself
        strDstDir = IO.Path.GetDirectoryName(strFile)
        ' Do we actually have a destination directory?
        If (strDstDir = "") Then
          ' Name a destination directory, which depends on this application
          strDstDir = IO.Path.GetDirectoryName(strSetFile)
        End If
        ' Get the name of this file
        strName = IO.Path.GetFileName(strFile)
        ' See if we can load it from our own directory to the one indicated here
        For intI = 0 To arDefFile.Length - 1
          ' Is this the correct file?
          If (IO.Path.GetFileName(arDefFile(intI)) = strName) Then
            ' Now we can bother about the destination directory
            If (Not IO.Directory.Exists(strDstDir)) Then
              ' Try create this directory
              Try
                IO.Directory.CreateDirectory(strDstDir)
              Catch ex As Exception
                ' No bother about exceptions
              End Try
              ' See if it has been established
              If (Not IO.Directory.Exists(strDstDir)) Then
                ' Ask user if he/she wants this default file to be downloaded anyway
                With Me.FolderBrowserDialog1
                  ' Set initial directory
                  .SelectedPath = strWorkDir
                  ' Set message
                  .Description = "About to download default file [" & strName & "]" & vbCrLf & _
                  "Please specify the directory where it should be copied to"
                  Select Case .ShowDialog
                    Case Windows.Forms.DialogResult.OK, Windows.Forms.DialogResult.Yes
                    Case Windows.Forms.DialogResult.No, Windows.Forms.DialogResult.Cancel
                      ' Don't do it
                      Return False
                  End Select
                  ' Get the new location
                  strDstDir = .SelectedPath
                End With
              End If
            End If
            ' Combine the directory and the file
            strFile = strDstDir & "\" & strName
            ' Try to download it from the internet
            Status("Trying to download " & arDefFile(intI) & "...")
            My.Computer.Network.DownloadFile(arDefFile(intI), strFile, "", "", True, 5000, True)
            ' Exit with success
            Status("ok...")
            Return True
          End If
        Next intI
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show what's wrong
      MsgBox("frmMain/TryLoadFile: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failre
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       CheckChainDict()
  ' Goal:       Check whether there is a valid path for the chain dictionary
  ' History:
  ' 12-07-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Function CheckChainDict() As Boolean
    Dim bNeedChain As Boolean = False ' Whether the period definition file should be determined
    Dim strPath As String = ""        ' Path of chain dictionary

    Try
      ' Check if the Chain dictionary file is known -- otherwise don't allow File/Open
      If (strChainDict = "") Then
        bNeedChain = True
      ElseIf (Not IO.Directory.Exists(IO.Path.GetDirectoryName(strChainDict))) Then
        bNeedChain = True
      End If
      If (bNeedChain) Then
        ' Make a default path for the chain dictionary and warn the user in a log
        strChainDict = GetDocDir() & "\ChainDictionary.xml"
        ' Save this setting
        SetTableSetting(tdlSettings, "ChainDictionary", strChainDict)
        ' Save this
        XmlSaveSettings(strSetFile)
        ' Warn the user
        Logging("You have not supplied a valid path/name for the Chain Dictionary." & vbCrLf & _
                "Cesax has supplied a default: " & strChainDict & vbCrLf & _
               "If you would like to change it, go to Tools/Settings, supply a valid chain dictionary location, and try File/Open again.")
        Return False
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/CheckChainDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      Return False
    End Try
  End Function

  '---------------------------------------------------------------------------------------------------------
  ' Name:       mnuFileNew_Click()
  ' Goal:       Create a new psdx file
  ' History:
  ' 22-05-2015  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub mnuFileNew_Click(sender As System.Object, e As System.EventArgs) Handles mnuFileNew.Click
    Dim strFile As String = ""        ' One file

    Try
      ' Check if the PeriodDefinition file is known -- otherwise don't allow File/Open
      If (Not CheckPeriodFile()) Then Exit Sub
      ' Check the chain dictionary
      If (Not CheckChainDict()) Then Exit Sub
      ' Should the old file be saved?
      If (bDirty) Then
        ' Try save the current file...
        If (Not SaveOne(True, False, True)) Then
          ' Exit without quitting
          Exit Sub
        End If
      End If
      ' Ask for filename, text name, meta information etc
      With frmNewText
        ' Set preliminary directory: the directory where the last text comes from
        If (strCurrentFile = "") Then
          .Directory = strWorkDir
        Else
          .Directory = IO.Path.GetDirectoryName(strCurrentFile)
        End If
        ' Ask for information
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Get the filename for further processing
            strFile = .FileName
            ' Check the file name
            If (Not IO.File.Exists(strFile)) Then
              Status("No proper file has been created")
              Exit Sub
            End If
          Case Windows.Forms.DialogResult.Cancel
            ' Leave gracefully...
            Status("No file has been created")
            Exit Sub
        End Select
      End With
      ' Show what has been created
      Status("Created: " & strFile)
      ' Load the file properly (this also sets "strCurrentFile" and so on)
      LoadPsdx(strFile)
      ' Set this as the most recent file
      If (Not SetRecent(strFile)) Then Status("Could not set 'recent' information") : Exit Sub
      ' Check which menu items should be displayed
      CheckRecent()
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/FileNew error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------------------------------------------------
  ' Name:       mnuFileOpen_Click()
  ' Goal:       Open a psdx file
  ' History:
  ' 23-04-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub mnuFileOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileOpen.Click
    Dim bChanged As Boolean = False   ' Indicates that we are not changed
    Dim strFile As String = ""        ' One file

    Try
      ' Check if the PeriodDefinition file is known -- otherwise don't allow File/Open
      If (Not CheckPeriodFile()) Then Exit Sub
      ' Check the chain dictionary
      If (Not CheckChainDict()) Then Exit Sub
      ' Should the old file be saved?
      If (bDirty) Then
        ' Try save the current file...
        If (Not SaveOne(True, False, True)) Then
          ' Exit without quitting
          Exit Sub
        End If
      End If
      ' Start with current file
      strFile = strCurrentFile
      ' Locate the file that should be opened
      If (GetFileName(Me.OpenFileDialog1, strWorkDir, strFile, "PSD xml files|*.psdx")) Then
        ' Try open this file
        If (Not LoadOneFile(strFile)) Then
          ' Just exit
          Exit Sub
        End If
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/FileOpen error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------------------------------------------------
  ' Name:       CheckRecent()
  ' Goal:       Check which Recent menu items should be displayed
  ' History:
  ' 10-03-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Sub CheckRecent()
    Dim intI As Integer   ' Counter
    Dim strFile As String ' One file
    Dim ctrThis As ToolStripMenuItem = Nothing

    Try
      ' Loop
      For intI = 1 To 5
        ' Check this one
        strFile = GetTableSetting(tdlSettings, "Recent" & intI)
        ' control depends on which one
        Select Case intI
          Case 1
            ctrThis = Me.mnuFileRecent1
          Case 2
            ctrThis = Me.mnuFileRecent2
          Case 3
            ctrThis = Me.mnuFileRecent3
          Case 4
            ctrThis = Me.mnuFileRecent4
          Case 5
            ctrThis = Me.mnuFileRecent5
        End Select
        ' We have something?
        If (strFile = "") Then
          ' Reset this one
          ctrThis.Visible = False
        Else
          ' Adapt this one
          With ctrThis
            ' Set the name
            .Text = "&" & intI & " " & strFile
            ' Make it visible
            .Visible = True
          End With
        End If
      Next intI
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/CheckRecent error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------------------------------------------------
  ' Name:       OpenRecent()
  ' Goal:       Open the indicated PSDX file.
  ' History:
  ' 10-03-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub OpenRecent(ByVal intNum As Integer)
    Dim strFile As String  ' File name

    Try
      ' Check if the PeriodDefinition file is known -- otherwise don't allow File/Open
      If (Not CheckPeriodFile()) Then Exit Sub
      ' Check the chain dictionary
      If (Not CheckChainDict()) Then Exit Sub
      ' Should the old file be saved?
      If (bDirty) Then
        ' Try save the current file...
        If (Not SaveOne(True, False, True)) Then
          ' Exit without quitting
          Exit Sub
        End If
      End If
      ' Get the name of the file
      strFile = GetTableSetting(tdlSettings, "Recent" & CStr(intNum))
      ' Try open this file
      LoadOneFile(strFile)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/OpenRecent error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------------------------------------------------
  ' Name:       mnuFileRecentX_Click()
  ' Goal:       Open a psdx file that has recently been opened.
  ' History:
  ' 10-03-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub mnuFileRecent1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileRecent1.Click
    OpenRecent(1)
  End Sub
  Private Sub mnuFileRecent2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileRecent2.Click
    OpenRecent(2)
  End Sub
  Private Sub mnuFileRecent3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileRecent3.Click
    OpenRecent(3)
  End Sub
  Private Sub mnuFileRecent4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileRecent4.Click
    OpenRecent(4)
  End Sub
  Private Sub mnuFileRecent5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileRecent5.Click
    OpenRecent(5)
  End Sub
  '---------------------------------------------------------------------------------------------------------
  ' Name:       GetFullPsdxName()
  ' Goal:       Try to get the full psdx name of the file
  ' History:
  ' 05-10-2012  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function GetFullPsdxName(ByRef strFile As String) As Boolean
    ' Dim arFile() As String      ' Result of looking in the working directory
    Dim strDir As String        ' The directory
    Dim strFull As String = ""  ' Full name of file we find
    ' Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (strFile = "") Then Return False
      ' Check if we can access the directory
      strDir = IO.Path.GetDirectoryName(strFile)
      ' Check if this file exists in the current directory
      If (Not IO.File.Exists(strFile)) Then
        ' Take only the file + extension
        strFile = IO.Path.GetFileName(strFile)
        ' First look for it in the [tbDbSrcDir]
        strDir = Trim(Me.tbDbSrcDir.Text)
        If (GetFileInDirectory(strDir, strFile, strFull)) Then
          ' Return positively
          strFile = strFull : Return True
        End If
        ' Try look in the working directory
        strDir = strWorkDir
        If (GetFileInDirectory(strDir, strFile, strFull)) Then
          ' Return positively
          strFile = strFull : Return True
        End If
      Else
        ' This already is a full file, so we are safe as it is
        Return True
      End If
      ' Return negatively: we did not find it
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/GetFullPsdxName error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       LoadPsdx()
  ' Goal:       Open a psdx file
  ' History:
  ' 23-04-2010  ERK Created
  ' 05-10-2012  ERK If the file does not exist, then start looking for it in the [strWorkDir]
  ' 05-02-2013  ERK Added "interactive" option
  '---------------------------------------------------------------------------------------------------------
  Public Sub LoadPsdx(ByVal strFile As String, Optional ByVal bInteractive As Boolean = True, Optional ByVal intLine As Integer = 0)
    Dim bChanged As Boolean = False   ' Indicates that we are not changed
    ' Dim ndxLast As XmlNode            ' One node

    Try
      ' Set the file that should be opened
      strCurrentFile = strFile
      ' Indicate we are not initialized yet, and initialize the EtreeId
      bPsdxInit = False
      ' Show we are loading
      Status("Loading file...")
      ' Signal that the editor is loading
      bEdtLoad = True
      ' Load up the text
      Me.tbEdtMain.Text = "(Please wait while loading file...)"
      ' Make sure the [TransFile] is set too
      strCurrentTrans = IO.Path.GetDirectoryName(strCurrentFile) & "\" & _
        IO.Path.GetFileNameWithoutExtension(strCurrentFile) & ".psdy"
      ' Switch over to the correct tabpage
      Me.TabControl1.SelectedTab = Me.tpEdit
      ' Set the XML document to nothing
      pdxCurrentFile = Nothing
      ' Make sure the focus is NOT on the RTB where we are loading...
      Me.tbEdtFeatures.Focus()
      ' Read the Schema file and the newly selected PSDX file
      Status("Loading file...")
      If (ReadXmlDoc(strCurrentFile, pdxCurrentFile)) Then
        ' We are in the process of reading it... Enable stuff
        MenuItems(True, False)
        ' Initialise the tei metadata stuff
        If (Not TeiMetaDataInit()) Then Logging("Error: could not initialise metadata") : Exit Sub
        ' Unpack the file and show it...
        ' (This also adds sections if needed)
        ' (This also gets the maximum EtreeId)
        ShowPsdx(Me.tbEdtMain, bChanged)
        ' Possibly correct translation language(s) in the file
        If (Not TransLangCheck(bChanged)) Then Logging("Warning: could not check and repair translation languages")
        ' Get the language(s) in the file
        If (Not TransLangInit(Me.cboGenTransLang)) Then Logging("Warning: there was a problem initialising the Translation Languages")
        If (Not TransLangInit(Me.cboDepTransLang)) Then Logging("Warning: there was a problem initialising the Translation Languages")
        ' Check interrupt
        If (bInterrupt) Then Exit Sub
        ' Initially reset the dirty bit, unless changed
        SetDirty(bChanged)
        ' Show we are initialized
        bPsdxInit = True
        ' If we are changed, this is because a section has been added!!!
        If (bChanged) Then
          ' It is changed if we have automatically made a start section
          Status("One or more sections have automatically been added (that's why you will be asked to save the file).")
        Else
          ' There already were sections...
          Status("Ready")
        End If
      Else
        ' Show failure
        Status("Failed to load " & strCurrentFile)
      End If
      ' Fill in the name of the document
      Me.tbName.Text = IO.Path.GetFileNameWithoutExtension(strCurrentFile)
      strCurrentPeriod = GetPeriod(pdxCurrentFile, False)
      Me.tbGenPeriod.Text = strCurrentPeriod
      ' Fill in the other information of the document on the general tab
      Me.tbGenTitle.Text = GetFileDesc(pdxCurrentFile, "title")
      Me.tbGenDistributor.Text = GetFileDesc(pdxCurrentFile, "distributor")
      Me.tbGenSource.Text = GetFileDesc(pdxCurrentFile, "bibl")
      Me.tbGenAuthor.Text = GetFileDesc(pdxCurrentFile, "author")
      Me.tbGenCreaManu.Text = GetFileDesc(pdxCurrentFile, "manuscript")
      Me.tbGenCreaOrig.Text = GetFileDesc(pdxCurrentFile, "original")
      Me.tbGenSubType.Text = GetFileDesc(pdxCurrentFile, "subtype")
      Me.tbGenEditor.Text = GetFileDesc(pdxCurrentFile, "editor")
      Me.tbGenLngId.Text = GetFileDesc(pdxCurrentFile, "ident", "language")
      Me.tbGenLngName.Text = GetFileDesc(pdxCurrentFile, "name", "language")
      ' Set the current etho language
      strCurrentEthno = GetFileDesc(pdxCurrentFile, "ident", "language")
      If (strCurrentEthno = "") Then
        ' Do we have a period?
        If (strCurrentPeriod <> "") Then
          If (DoLike(strCurrentPeriod, "O[1-4]|O[1-4][1-4]|M[1-4]|M[1-4][1-4]|E[1-3]|B[1-3]")) Then
            strCurrentEthno = "eng_hist_" & strCurrentPeriod
          End If
        End If
      End If
      ' Reset the tree editor undo-stack
      InitTreeStack()
      ' Show the revision history
      ShowRevision()
      ' Set the selection point in the beginning of the text
      Me.tbEdtMain.SelectionStart = 1
      ' Put the focus back to the editor
      Me.tbEdtMain.Focus()
      ' Validate
      If (Not pdxSection Is Nothing) Then
        ' Do we have sections?
        If (pdxSection.Count <= 1) Then
          ' We have to go to section 0, because there is only one section
          ShowSection(Me.tbEdtMain, 0)
        ElseIf (bInteractive) Then
          ' Give user the possibility to go to a section right now
          SectionGoto()
        Else
          ' Check what section needs to be shown, depending on the line-number
          If (intLine = 0) Then
            ' Just go to the first section
            ShowSection(Me.tbEdtMain, 0)
          Else
            ShowSection(Me.tbEdtMain, LineToSection(intLine))
          End If
        End If
      End If
      ' Clear the collapse tree
      CollapseClear()
      ' Reset the dirty flag
      SetDirty(False)
      If (Not TeiLakInit()) Then Logging("Could not initialize Lak", True)
      If (Not TeiCheInit()) Then Logging("Could not initialize Che", True)
      ' Show that the editor is not loading
      bEdtLoad = False
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/LoadPsdx error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       mnuFileSave_Click()
  ' Goal:       Show the Help-About form
  ' History:
  ' 23-04-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub mnuFileSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileSave.Click
    Try
      ' Try to save it
      If (Not SaveOne(False, False, True)) Then
        ' Something went wrong
        Status("Unable to save the information")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/FileSave error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       mnuFileSaveAs_Click()
  ' Goal:       Show the Help-About form
  ' History:
  ' 23-04-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub mnuFileSaveAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileSaveAs.Click
    Try
      ' Try to save it
      If (Not SaveOne(False, True, True)) Then
        ' Something went wrong
        Status("Unable to save the information")
      Else
        ' Make sure file is set as most recent one
        SetRecent(strCurrentFile)
        ' Display must be adapted
        CheckRecent()
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/FileSave error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       mnuFileExport_Click()
  ' Goal:       In principle the user should be able to export information into the following way(s):
  '             (1) Penn-Treebank 2 bracketed labelling (psd)
  '             (2) ???
  ' History:
  ' 29-04-2010  ERK Created
  ' 11-10-2012  ERK Added the IS-output component from Cesac2008 (topic guesser)
  '---------------------------------------------------------------------------------------------------------
  Private Sub mnuFileExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileExport.Click
    Dim strExpType As String    ' Kind of export type chosen
    Dim strFile As String       ' Output file for the export information
    Dim strFilter As String     ' Filter for output storage
    Dim intFiltIdx As Integer   ' Index of chosen filter
    Dim arFilter() As String = { _
      "Penn-Treebank-2 full (psd)|*.psd", _
      "Penn-Treebank-2 simple (psd)|*.psd", _
      "CoNLL-x dependency (dep)|*.dep", _
      "Folia (xml)|*.xml", _
      "English translation (psdy)|*.psdy", _
      "Interrater agreement (csv)|*.csv", _
      "POS - original (pos)|*.pos", _
      "POS - simple (pos)|*.pos", _
      "POS - Penn Treebank (pos)|*.pos", _
      "Token (tok)|*.tok"}
    Dim arExt() As String = {"psd", "psd", "dep", "xml", "psdy", "csv", "pos", "pos", "pos", "tok"}
    Dim arExpType() As String = {"PsdOutput", "PsdSimple", "DepCoNLL-X", "Folia", _
                                 "PdeOutput", "Rating", _
                                 "PosOrg", "PosSimple", "PosPTB", "Token"}

    Try
      ' Validate
      If (strCurrentFile = "") Then
        ' Show our worries
        Status("First open a file!")
        Exit Sub
      End If
      ' Stel het filter element in op alle soorten bestanden die we willen kunnen exporteren
      strFilter = Join(arFilter, "|")
      ' Get the output file name and location from the user
      With Me.SaveFileDialog1
        ' Default filename is derived from current PSDx file
        strFile = IO.Path.GetDirectoryName(strCurrentFile) & "\" & _
          IO.Path.GetFileNameWithoutExtension(strCurrentFile)
        ' Take the correct default type
        intFiltIdx = 1
        .Filter = strFilter
        ' Make sure the PSD is chosen by default
        .FilterIndex = intFiltIdx
        ' Set the correct default extension
        .DefaultExt = arExt(intFiltIdx - 1)
        ' Assign default file name to the FileSave dialog
        .FileName = strFile
        ' Show the actual dialog to the user
        Select Case .ShowDialog()
          Case Windows.Forms.DialogResult.OK, Windows.Forms.DialogResult.Yes
            ' Get the filename that the user has finally selected
            strFile = .FileName
            ' Get the filterindex back
            intFiltIdx = .FilterIndex
          Case Else
            ' Aborted, so exit
            Exit Sub
        End Select
      End With
      ' Determine the export type, depending on the filter index
      strExpType = arExpType(intFiltIdx - 1)
      ' Perform the chosen export using the general subroutine for this
      If (Not DoExport(strExpType, strFile)) Then
        ' Don't know what to do now - probably called function already gives error message
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/FileExport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       mnuFileImport_Click()
  ' Goal:       Allow the user to import information of the following kind(s):
  '             (1) Chain dictionary (xml)
  '             (2) Translation files
  ' History:
  ' 06-07-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub mnuFileImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileImport.Click
    Dim strImpType As String    ' Kind of import type chosen
    Dim strFile As String       ' Input file for the import information
    Dim strFilter As String     ' Filter for input storage
    Dim intFiltIdx As Integer   ' Index of chosen filter
    Dim arFilter() As String = { _
      "Treebank file (psd)|*.psd", _
      "Chunk-parsed (psd)|*.psd", _
      "Chain dictionary (xml)|*.xml", _
      "Dependency (conll-X)|*.conll", _
      "English translation (psdy)|*.psdy", _
      "Folia file (xml)|*.xml"}
    Dim arExt() As String = {"psd", "psd", "xml", "conll", "psdy", "xml"}
    Dim arImpType() As String = {"Treebank", "ChunkParsed", "ChainDict", "ConLLX", _
                                 "PdeInput", "Folia"}

    Try
      ' Stel het filter element in op alle soorten bestanden die we willen kunnen exporteren
      strFilter = Join(arFilter, "|")
      ' Get the input file
      With Me.OpenFileDialog1
        ' Take the correct default type
        intFiltIdx = GetTableSetting(tdlSettings, "ImportFilter", "1")
        .Filter = strFilter
        ' Make sure the PSD is chosen by default
        .FilterIndex = intFiltIdx
        ' Set the correct default extension
        .DefaultExt = arExt(intFiltIdx - 1)
        ' Go to the default working directory
        .InitialDirectory = GetTableSetting(tdlSettings, "ImportDir", strWorkDir)
        ' Set the correct file name
        .FileName = "*." & .DefaultExt
        ' Show the actual dialog to the user
        Select Case .ShowDialog()
          Case Windows.Forms.DialogResult.OK, Windows.Forms.DialogResult.Yes
            ' Get the filename that the user has finally selected
            strFile = .FileName
            ' Get the filterindex back
            intFiltIdx = .FilterIndex
            ' Store the filter index for future
            SetTableSetting(tdlSettings, "ImportFilter", intFiltIdx)
            SetTableSetting(tdlSettings, "ImportDir", IO.Path.GetDirectoryName(strFile))
            ' Save the settings
            XmlSaveSettings(strSetFile)
          Case Else
            ' Aborted, so exit
            Exit Sub
        End Select
      End With
      ' Determine the export type, depending on the filter index
      strImpType = arImpType(intFiltIdx - 1)
      ' Perform the chosen export using the general subroutine for this
      If (DoImport(strImpType, strFile)) Then
        ' Set the dirty flag
        SetDirty(True)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/FileImport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------------------------------------------------
  ' Name:       mnuFileExit_Click()
  ' Goal:       Try to exit nicely
  ' History:
  ' 23-04-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub mnuFileExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileExit.Click
    ' Should something be saved?
    If (bDirty) Then
      ' Try to save this
      If (Not SaveOne(True, False, True)) Then
        ' Ask if user wants to exit anyway
        Select Case MsgBox("Saving the results was not succesfull." & vbCrLf & _
                           "Exit CESAX anyway?", MsgBoxStyle.YesNoCancel)
          Case MsgBoxResult.No, MsgBoxResult.Cancel
            ' Don't exit the PROGRAM yet...
            Exit Sub
        End Select
      End If
      ' Also save the chain dictionary if necessary
      TrySaveChainDict()
    End If
    ' Check if results database should be changed
    If (bResDirty) Then
      ' Ask if user wants to exit anyway
      Select Case MsgBox("Would you like to save the results in the database before leaving?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Yes
          ' Try save the database
          TrySaveResDb(strResultDb)
          ' Reset dirty bit
          SetResDirty(False)
        Case MsgBoxResult.No
        Case MsgBoxResult.Cancel
          'Leave
          Exit Sub
      End Select
    End If
    ' Exit the program
    End
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       mnuTransShow_Click()
  ' Goal:       Show the translation, focusing on the currently selected clause
  ' History:
  ' 04-03-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub mnuTransShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTransShow.Click
    Try
      ' Call the appropriate function
      ShowTranslation(True)
      ' Possible extension: save the HTML text
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/TransShow error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------------------------------------------------
  ' Name:       mnuTransFirst_Click()
  ' Goal:       Select one particular line to go to
  ' History:
  ' 03-03-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub mnuTransFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTransFirst.Click
    Try
      ' Validate
      If (pdxList.Count = 0) Then Exit Sub
      ' Go to the first line
      GotoTransLine(0)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/TransFirst error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub mnuTransLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTransLast.Click
    Dim intLine As Integer  ' Line we need to go to

    Try
      ' Validate
      If (pdxList.Count = 0) Then Exit Sub
      ' Which mode are we in?
      If (Me.mnuTransSection.Checked) Then
        ' Go to the last line of the section
        intLine = intSectLast
      Else
        ' Take the last line overall
        intLine = pdxList.Count - 1
      End If
      ' Go to the last line
      GotoTransLine(intLine)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/TransLast error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub mnuTransNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTransNext.Click
    Try
      ' Validate
      If (pdxList.Count = 0) Then Exit Sub
      ' Do the moving elsewhere
      TransNext(True)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/TransNext error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub mnuTransPrevious_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTransPrevious.Click
    Try
      ' Validate
      If (pdxList.Count = 0) Then Exit Sub
      ' Do the moving elsewhere
      TransNext(False)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/TransPrev error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub mnuTransGoto_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTransGoto.Click
    Dim intLine As Integer  ' Line we need to go to
    Dim strBack As String   ' Response

    Try
      ' Validate
      If (pdxList.Count = 0) Then Exit Sub
      ' Check what line the user wants to go to
      strBack = InputBox("Line to jump to", "Goto line", intLastTransLine + 1) - 1
      If (strBack = "") Then Exit Sub
      If (Not IsNumeric(strBack)) Then Exit Sub
      intLine = CInt(strBack)
      ' Check the result
      If (intLine >= 0 AndAlso intLine < pdxList.Count) Then
        ' Jump to this line
        GotoTransLine(intLine)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/TransGoto error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------------------------------------------------
  ' Name:       TransNext()
  ' Goal:       Go to the next or the previous line in the translation
  ' History:
  ' 04-03-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub TransNext(ByVal bNext As Boolean)
    Dim intLine As Integer  ' Line we need to go to

    Try
      ' Which mode are we in?
      If (Me.mnuTransSection.Checked) Then
        If (bNext) Then
          ' Go to the next line within this section
          If (intLastTransLine < intSectLast) Then
            intLine = intLastTransLine + 1
            GotoTransLine(intLine)
          End If
        Else
          ' Go to the previous line within this section
          If (intLastTransLine > intSectFirst) Then
            intLine = intLastTransLine - 1
            GotoTransLine(intLine)
          End If
        End If
      Else
        If (bNext) Then
          ' Go to the next line overall
          If (intLastTransLine < pdxList.Count + 1) Then
            intLine = intLastTransLine + 1
            GotoTransLine(intLine)
          End If
        Else
          ' Go to the previous line overall
          If (intLastTransLine > 0) Then
            intLine = intLastTransLine - 1
            GotoTransLine(intLine)
          End If
        End If
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/TransNext error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------------------------------------------------
  ' Name:       mnuTransCreate_Click()
  ' Goal:       Show original and translation until the first line not yet done
  ' History:
  ' 25-10-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub mnuTransCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTransCreate.Click
    Dim intFirstTrans As Integer  ' First line that is untranslated
    Try

      ' Find the first line that is not yet translated
      intFirstTrans = GetFirstTrans()
      ' Go to this particular line
      GotoTransLine(intFirstTrans)
      ' Show we are ready
      Status("Ready")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/TransCreate error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       mnuTransAddLang_Click()
  ' Goal:       Add a language into which a translation can be made
  ' History:
  ' 15-03-2014  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub mnuTransAddLang_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTransAddLang.Click
    Dim strLang As String = ""  ' Ethno 3-letter code of language
    Dim bChanged As Boolean = False ' Changes

    Try
      ' Ask for ethno code
      strLang = InputBox("Give the 3-letter code of the translation-language", "Translation language", "nld")
      ' Check for valid answer
      strLang = Trim(strLang)
      If (Len(strLang) <> 3) Then Status("You should provide a THREE-letter code") : Exit Sub
      ' Add this language in all forests
      If (Not AddTransLang(strLang, bChanged)) Then Status("There was an error") : Exit Sub
      ' Check for changes
      If (bChanged) Then SetDirty(True)
      ' Show we are ready
      Status("Translation language [" & strLang & "] has been added")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/TransAddLang error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------------------------------------------------
  ' Name:       mnuTransSection_Click()
  ' Goal:       Toggle between the scope of the translation navigation command:
  '             Checked   = only show current section
  '             Unchecked = show whole file
  ' History:
  ' 07-03-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub mnuTransSection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTransSection.Click
    Try
      ' Toggle the setting
      Me.mnuTransSection.Checked = (Not Me.mnuTransSection.Checked)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/TransSection error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------------------------------------------------
  ' Name:       GotoTransLine()
  ' Goal:       Show original and translation until the first line not yet done
  ' History:
  ' 25-10-2010  ERK Created
  ' 03-03-2011  ERK Also show the remaining part
  '---------------------------------------------------------------------------------------------------------
  Public Sub GotoTransLine(ByVal intLine As Integer)
    Dim strText As String = ""    ' The text to be translated
    Dim strPrevOrg As String = "" ' Original so far
    Dim strPrevTrn As String = "" ' Translation so far
    Dim strNextOrg As String = "" ' Original afterwards
    Dim strNextTrn As String = "" ' Translation afterwards
    Dim strThisOrg As String = "" ' Original of this line
    Dim strThisTrn As String = "" ' Translation of this line

    Try
      ' Validation (Don't know what to validate yet)
      ' Switch to the correct tab page
      Me.TabControl1.SelectedTab = Me.tpTranslation
      ' Do we need to stay within the section?
      If (Me.mnuTransSection.Checked) Then
        ' Check whether we are inside the section
        If (intLine < intSectFirst) OrElse (intLine > intSectLast) Then
          ' Warn the user
          Status("The line you want to go to (" & intLine + 1 & ")is outside the current selection [" & _
                 intSectFirst + 1 & "-" & intSectLast + 1 & "]")
          Exit Sub
        End If
      End If
      ' Signal we are busy
      Status("Getting the translation so far...")
      Me.tbAddNewTrn.Text = "Please wait..."
      ' Show the whole text (original + translation), particularly focusing on the 
      '   first untranslated line
      strText = GetTextTrans(intLine, Me.mnuTransSection.Checked, strPrevOrg, strPrevTrn, _
                             strThisOrg, strThisTrn, strNextOrg, strNextTrn)
      ' Show the text before and after
      With Me.tbAddOrg
        ' Clear previous
        .Text = ""
        ' Load original so far
        .SelectionStart = 1
        .SelectionColor = Color.Black
        .SelectedText = strPrevOrg
        ' Add last line in BLUE
        .SelectionStart = .TextLength
        .SelectionColor = Color.Blue
        .SelectedText = strThisOrg
        ' Scroll to the end of the translation
        .SelectionStart = .TextLength
        .SelectionColor = Color.Black
        ' Make sure this is taken into effect
        .Focus()
      End With
      ' Provide the translation so far
      With Me.tbAddTrn
        .Text = strPrevTrn
        ' Possibly add existing translation
        If (strThisTrn <> "") Then .AppendText(strThisTrn)
        ' Scroll to the end of the translation
        .SelectionStart = .TextLength
        ' Make sure this is taken into effect
        .Focus()
      End With
      ' Set following text and translation
      Me.tbNextOrg.Text = strNextOrg
      Me.tbNextTrn.Text = strNextTrn
      ' Make sure the following text is shown from the start
      ' Scroll to the start of the translation
      Me.tbNextOrg.SelectionStart = 1
      Me.tbNextTrn.SelectionStart = 1
      ' Show the current translation or ask for a new translation
      With Me.tbAddNewTrn
        ' Set focus on asking for a new translation
        .Focus()
        If (strText = "") Then
          ' Ask for a new translation
          .Text = "<Add translation here>"
          .SelectionStart = 0
          .SelectionLength = .TextLength
        Else
          ' Show the current translation
          .Text = strText
          .SelectionStart = .TextLength
        End If
      End With
      '' Make sure the focus is on asking for a new translation
      'Me.tbAddNewTrn.Focus()
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/GotoTransLine error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------------------------------------------------
  ' Name:       mnuHelpAbout_Click()
  ' Goal:       Show the Help-About form
  ' History:
  ' 23-04-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub mnuHelpAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpAbout.Click
    ' Show the About form
    frmAbout.ShowDialog()
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       mnuHelpQuick_Click()
  ' Goal:       Show Quick help for the user (show a HTML document with help)
  ' History:
  ' 23-04-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub mnuHelpQuick_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpQuick.Click
    Try
      ' Show the Quick Reference Help information form
      With frmHelp
        .Show()
        .Help = IO.Path.GetDirectoryName(Application.ExecutablePath) & "\" & QUICK_REFERENCE
      End With
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/HelpQuick error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       mnuHelpHtml_Click()
  ' Goal:       Open the manual (html) from the internet
  ' History:
  ' 14-07-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub mnuHelpHtml_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpHtml.Click
    Try
      With frmHelp
        .Show()
        .Help = HELP_HTML
      End With
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/HelpHtml error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       mnuHelpManual_Click()
  ' Goal:       Open the manual (PDF) from the internet
  ' History:
  ' 15-03-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub mnuHelpManual_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpManual.Click
    Dim strFile As String   ' Where we are expecting the file to be
    Dim strLocal As String  ' Local position of the file
    Dim strTemp As String   ' Local file name

    Try
      ' Get the name ofthe file
      strFile = HELP_MANUAL
      ' Get the local position of the file
      strLocal = My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\Cesax\" & _
        IO.Path.GetFileName(strFile)
      ' Check if the file is already there
      If (IO.File.Exists(strLocal)) Then
        ' Can we get access to the Internet file?
        If (My.Computer.Network.IsAvailable) Then
          ' Download it to a local file
          strTemp = IO.Path.GetTempPath & IO.Path.GetFileNameWithoutExtension(IO.Path.GetTempFileName) & ".pdf"
          My.Computer.Network.DownloadFile(strFile, strTemp, "", "", True, 1000, True)
          ' See if the file on internet is newer
          If (IO.File.GetLastAccessTimeUtc(strTemp) > IO.File.GetLastAccessTimeUtc(strLocal)) Then
            ' Copy it
            IO.File.Copy(strTemp, strLocal, True)
          End If
        End If
      Else
        ' Does the directory exist?
        If (Not IO.Directory.Exists(IO.Path.GetDirectoryName(strLocal))) Then
          ' Make the directory
          IO.Directory.CreateDirectory(IO.Path.GetDirectoryName(strLocal))
        End If
        ' Internet access?
        If (My.Computer.Network.IsAvailable) Then
          ' Try downloading
          Status("Trying to download " & IO.Path.GetFileName(strFile) & "...")
          My.Computer.Network.DownloadFile(strFile, strLocal, "", "", True, 1000, True)
          ' Exit with success
          Status("ok...")
        Else
          ' Failure...
          Status("You need to be connected to the internet")
          Exit Sub
        End If
      End If
      ' Double check
      If (Not IO.File.Exists(strLocal)) Then
        ' Warn user
        Status("The PDF manual has not been correctly downloaded")
        Exit Sub
      End If
      ' Show the locally available PDF file
      System.Diagnostics.Process.Start(strLocal)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/HelpManual error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try

  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       mnuToolsSettings_Click()
  ' Goal:       Open form to adjust settings
  ' History:
  ' 23-04-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsSettings_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsSettings.Click
    ' Show the form where the settings can be adjusted
    frmSetting.Show()
  End Sub
  '---------------------------------------------------------------------------------------------------------
  ' Name:       mnuToolsUpdate_Click()
  ' Goal:       Try to move stuff from the shadow to my own settings
  ' History:
  ' 14-03-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsUpdateSettings.Click
    Try
      ' Validate
      If (Not bInit) Then Exit Sub
      ' Check the settings
      If (Not CheckSettings()) Then Exit Sub
      ' Show we have updated
      Status("Update ready")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/mnuToolsUpdate_Click error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsSoftwareUpdate_Click
  ' Goal:   Check for update in the software
  ' History:
  ' 25-03-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsSoftwareUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsSoftwareUpdate.Click
    Try
      ' Attempt to do an update in the application
      UpdateApplication(False)
    Catch ex As Exception
      ' Show message
      HandleErr("frmMain/ToolsUpdate error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  Private Sub frmMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
    ' Stop
  End Sub

  'Private Sub frmMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
  '  Try
  '    ' Check which form is selected
  '    If (litTree.Visible) Then
  '      litTree_KeyDown(sender, e)
  '    Else
  '      ' Let others handle it
  '      e.Handled = False
  '    End If
  '  Catch ex As Exception
  '    ' Warn the user
  '    HandleErr("frmMain/frmMain_KeyDown error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
  '  End Try
  'End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       frmMain_Load()
  ' Goal:       Trigger initialisation and show who we are
  ' History:
  ' 23-04-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Trigger initialisation
    Me.Timer1.Enabled = True
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       Timer1_Tick()
  ' Goal:       Try to start initialisation
  ' History:
  ' 23-04-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Dim intHangingIndent As Integer = 7   ' Pixels hanging indent
    Dim bMandatory As Boolean             ' Do we have to check for mandatory updates only?

    Try
      ' Switch off timer
      Me.Timer1.Enabled = False
      ' Disable menu items that cannot be used unless something is loaded
      MenuItems(False, False)
      ' Look for any possible software message
      LookForUpdMsg(strSoftwareMsg)
      Me.TabControl2.SelectedTab = Me.tpFeatures
      ' SoftwareMessage()
      ' Are we supposed to check for updates?
      bMandatory = (Not bUpdStartup)
      ' Try updating...
      Status("Looking for updates...")
      UpdateApplication(bMandatory)
      ' Show user what we are doing
      Status("Reading settings file. Please wait...")
      mnuViewDependency_Click(sender, e)
      ' Try to get a settings file
      strSetFile = GetSettingsFile()
      ' Perform any other initialisations
      If (Not DoMainInit()) Then Exit Sub
      ' Set up margins for some textboxes
      Me.tbEdtMain.SelectionHangingIndent = intHangingIndent
      Me.tbEdtFeatures.SelectionHangingIndent = intHangingIndent
      Me.tbEdtSyntax.SelectionHangingIndent = intHangingIndent
      Me.tbAutoLog.SelectionHangingIndent = intHangingIndent
      Me.tbAskProblem.SelectionHangingIndent = intHangingIndent
      Me.tbMainPde.SelectionHangingIndent = intHangingIndent
      Me.tbGenRevision.SelectionHangingIndent = intHangingIndent
      Me.tbResPsd.SelectionHangingIndent = intHangingIndent
      Me.tbRefItemContext.SelectionHangingIndent = intHangingIndent
      ' Set the initial text for the textboxes
      Me.tbEdtFeatures.Text = "(no file loaded)"
      Me.tbEdtMain.Text = "(no file loaded)"
      Me.tbEdtSyntax.Text = "(no file loaded)"
      Me.tbMainPde.Text = "(no file loaded)"
      Me.tbName.Text = "(no file loaded)"
      ' Are we okay?
      If (strSetFile = "") Then
        ' Show user failure
        Status("Settings have not been loaded yet")
        bInit = False
      Else
        ' Show success
        Status("Settings have been loaded from: " & strSetFile)
        Logging("Settings have been loaded from: " & strSetFile)
        bInit = True
      End If
      ' Set the graphics object
      InitGraphics(Me.tbEdtMain)
      ' Initialize dragdrop effects
      Me.tbAddNewTrn.AllowDrop = True
      ' Check for recent files
      CheckRecent()
      ' Check for recent corpus files
      CheckCrpRec()
      ' Initialise FindX results dataset
      InitFindResults()
      ' Create meta object
      If (Not CreateDataSet("MetaH.xsd.txt", tdlMeta)) Then
        Status("Cannot create meta set") : bInit = False
      End If
      ' Show that we are initialized
      Status("Initialized.")
      Logging("You need to open a file before the tabpages become active...")
    Catch ex As Exception
      ' Give message
      Select Case MsgBox("Initialisation error: " & ex.Message & vbCrLf & "Would you like to try initialisation again?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Yes
          ' Switch the timer on again
          Me.Timer1.Enabled = True
        Case MsgBoxResult.No, MsgBoxResult.Cancel
          ' This means quitting the program
          Status("Exiting...")
          End
      End Select
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   InitFindResults
  ' Goal:   Initialise the dataset that will help show the Find results
  ' History:
  ' 17-04-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub InitFindResults()
    Try
      ' Validate
      If (Not tdlFindX Is Nothing) Then tdlFindX.Dispose() : tdlFindX = Nothing
      ' Create the dataset
      If (Not CreateDataSet("FindResult.xsd", tdlFindX)) Then
        Status("frmMain/InitFindresults: Could not create [FindResult] dataset")
      End If
    Catch ex As Exception
      ' Show message to user
      HandleErr("frmMain/InitFindResults error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   MenuItems
  ' Goal:   Set or Clear menu items that should only work when the content is right
  ' History:
  ' 29-07-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub MenuItems(ByVal bProjectLoaded As Boolean, ByVal bIsXML As Boolean)
    ' Set or clear available menu items
    Me.mnuFileSave.Enabled = bProjectLoaded
    Me.mnuFileSaveAs.Enabled = bProjectLoaded
    'Me.TabControl1.Enabled = bProjectLoaded
    ' Show our status
    If (bProjectLoaded) Then
      Logging("Project has been loaded")
    End If
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   SaveOneTrans
  ' Goal:   Check whether the translation needs to be saved, and then do it in the correct way
  ' History:
  ' 29-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function SaveOneTrans(ByVal bDoAsk As Boolean, ByVal bSaveAs As Boolean, ByVal bOverWrite As Boolean) As Boolean
    Try
      ' Are there any changes that need to be saved?
      If ((bDirty) OrElse (bSaveAs)) Then
        ' Do we need to ask the user for permission to save?
        If (bDoAsk) Then
          ' Ask the user
          Select Case MsgBox("Would you like to save changes to the current Psdx translation file?", MsgBoxStyle.YesNoCancel)
            Case MsgBoxResult.Cancel
              ' Leave this subroutine!!
              Return False
            Case MsgBoxResult.No
              ' No need to do anything then
            Case MsgBoxResult.Yes
              ' Invoke subroutine to actually save the work
              If (Not SaveOpenWorkTrans(bSaveAs, bOverWrite)) Then
                ' Return failure and leave
                Return False
              End If
          End Select
        Else
          ' Save whatever needs to be saved
          If (Not SaveOpenWorkTrans(bSaveAs, bOverWrite)) Then
            ' Return failure and leave
            Return False
          End If
        End If
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show message to user
      HandleErr("frmMain/SaveOneTrans error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SaveOne
  ' Goal:   Check whether something needs to be saved, and then do it in the correct way
  ' History:
  ' 29-12-2008  ERK Created shell
  ' ------------------------------------------------------------------------------------
  Private Function SaveOne(ByVal bDoAsk As Boolean, ByVal bSaveAs As Boolean, ByVal bOverWrite As Boolean) As Boolean
    ' Are there any changes that need to be saved?
    If ((bDirty) OrElse (bSaveAs)) Then
      ' Allow user to add revision information
      If (bRevision) Then AddRevision()
      ' Do we need to ask the user for permission to save?
      If (bDoAsk) Then
        ' Ask the user
        Select Case MsgBox("Would you like to save changes to the current Psdx file?", MsgBoxStyle.YesNoCancel)
          Case MsgBoxResult.Cancel
            ' Leave this subroutine!!
            Return False
          Case MsgBoxResult.No
            ' No need to do anything then
          Case MsgBoxResult.Yes
            ' Invoke subroutine to actually save the work
            If (Not SaveOpenWork(bSaveAs, bOverWrite)) Then
              ' Return failure and leave
              Return False
            End If
        End Select
      Else
        ' Save whatever needs to be saved
        If (Not SaveOpenWork(bSaveAs, bOverWrite)) Then
          ' Return failure and leave
          Return False
        End If
      End If
    End If
    ' Return success
    SaveOne = True
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SaveOpenWork
  ' Goal:   Save any open work. As the user when this is appropriate
  ' History:
  ' 06-12-2008  ERK Created shell
  ' ------------------------------------------------------------------------------------
  Private Function SaveOpenWork(ByVal bSaveAs As Boolean, ByVal bOverWrite As Boolean) As Boolean
    Dim strFile As String   ' Name of file to save

    Try
      ' Save anything necessary from the General tab
      If (bDirty) Then
        ' Do save changes...
        If (Not AddFileDesc(pdxCurrentFile, "title", Me.tbGenTitle.Text)) Then Return False
        If (Not AddFileDesc(pdxCurrentFile, "distributor", Me.tbGenDistributor.Text)) Then Return False
        If (Not AddFileDesc(pdxCurrentFile, "bibl", Me.tbGenSource.Text)) Then Return False
      End If
      ' Check filename
      If (strCurrentFile = "") OrElse (bSaveAs) Then
        ' Get the filename from the user
        With Me.SaveFileDialog1
          ' The initial directory is the one we already know
          .InitialDirectory = strWorkDir
          ' The name is derived from the name of this project
          strFile = Me.tbName.Text & ".psdx"
          ' Set the default extention
          .DefaultExt = "psdx"
          ' Set the default filter
          .Filter = "Psd XML files|*.psdx"
          ' Check and possibly make correct directory
          If (Not CheckDir(strWorkDir)) Then
            ' Return failure
            SaveOpenWork = False
            ' Leave
            Exit Function
          End If
          ' Assign default file name to the FileSave dialog
          'strCRPfile = strFile
          .FileName = strFile
          ' Show the actual dialog to the user
          Select Case .ShowDialog()
            Case Windows.Forms.DialogResult.OK, Windows.Forms.DialogResult.Yes
              ' Get the filename that the user has finally selected
              strCurrentFile = .FileName
            Case Else
              ' Aborted, so exit
              Exit Function
          End Select
        End With
      End If
      ' Check if we are overwriting
      If (IO.File.Exists(strCurrentFile)) AndAlso (Not bOverWrite) Then
        Select Case MsgBox("Overwrite existing " & strCurrentFile & "?", MsgBoxStyle.YesNoCancel)
          Case MsgBoxResult.Cancel
            ' Indicate user
            Status("Saving was canceled")
            ' Do nothing
            Return True
          Case MsgBoxResult.No
            ' Try to go and save using File Save As
            bSaveAs = True
            ' Try to go and save using File Save As
            Return SaveOpenWork(True, False)
          Case MsgBoxResult.Yes
            ' Show status
            Status("Saving the file...")
            ' Save the CRP information to the selected filename as XML
            pdxCurrentFile.Save(strCurrentFile)
            ' Show status
            Status("File has been saved: " & strCurrentFile & " (" & Now & ")")
        End Select
      Else
        ' Show status
        Status("Saving the file...")
        ' Save the CRP information to the selected filename as XML
        pdxCurrentFile.Save(strCurrentFile)
        ' Show status
        Status("File has been saved: " & strCurrentFile & " (" & Now & ")")
      End If
      ' Reset the dirty boolean
      SetDirty(False)
      ' Return success
      SaveOpenWork = True
    Catch ex As Exception
      ' Show message to user
      HandleErr("frmMain/SaveOpenWork error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SaveOpenWorkTrans
  ' Goal:   Save any open work. As the user when this is appropriate
  ' History:
  ' 29-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function SaveOpenWorkTrans(ByVal bSaveAs As Boolean, ByVal bOverWrite As Boolean) As Boolean
    Dim strFile As String         ' Name of file to save
    Dim bResult As Boolean = True ' The result of saving

    Try
      ' Check filename
      If (strCurrentTrans = "") OrElse (bSaveAs) Then
        ' Get the filename from the user
        With Me.SaveFileDialog1
          ' The initial directory is the one we already know
          .InitialDirectory = strWorkDir
          ' The name is derived from the name of this project
          strFile = Me.tbName.Text & ".psdy"
          ' Set the default extention
          .DefaultExt = "psdy"
          ' Set the default filter
          .Filter = "Psd XML translation files|*.psdy"
          ' Check and possibly make correct directory
          If (Not CheckDir(strWorkDir)) Then
            ' Return failure
            Return False
          End If
          ' Assign default file name to the FileSave dialog
          .FileName = strFile
          ' Show the actual dialog to the user
          Select Case .ShowDialog()
            Case Windows.Forms.DialogResult.OK, Windows.Forms.DialogResult.Yes
              ' Get the filename that the user has finally selected
              strCurrentTrans = .FileName
            Case Else
              ' Aborted, so exit
              Return False
          End Select
        End With
      End If
      ' Check if we are overwriting
      If (IO.File.Exists(strCurrentTrans)) AndAlso (Not bOverWrite) Then
        Select Case MsgBox("Overwrite existing " & strCurrentTrans & "?", MsgBoxStyle.YesNoCancel)
          Case MsgBoxResult.Cancel
            ' Indicate user
            Status("Saving was canceled")
            ' Do nothing
            Return True
          Case MsgBoxResult.No
            ' Try to go and save using File Save As
            Return SaveOpenWorkTrans(True, False)
          Case MsgBoxResult.Yes
            ' Do actual translation
            bResult = DoTransSave(strCurrentTrans)
        End Select
      Else
        ' Do actual translation
        bResult = DoTransSave(strCurrentTrans)
      End If
      ' Reset the dirty boolean
      SetDirty(False)
      ' Return success
      Return bResult
    Catch ex As Exception
      ' Show message to user
      HandleErr("frmMain/SaveOpenWorkTrans error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CheckDir
  ' Goal:   Check whether directory exists, and give user option to create it
  ' History:
  ' 05-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CheckDir(ByVal strDir As String) As Boolean
    If (Not IO.Directory.Exists(strDir)) Then
      Select Case MsgBox("Would you first like me to create directory " & strDir & "?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Yes
          ' Create this directory
          IO.Directory.CreateDirectory(strDir)
          ' return success
          CheckDir = True
        Case MsgBoxResult.No
          ' return success
          CheckDir = True
        Case MsgBoxResult.Cancel
          ' Abort operation
          CheckDir = False
        Case Else
          CheckDir = False
      End Select
    Else
      CheckDir = True
    End If
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SetDirty
  ' Goal:   Set or reset the dirty bit
  ' History:
  ' 29-07-2009  ERK Created
  ' 21-12-2009  ERK This sub should be PUBLIC, so that DgvHandle can call it too.
  ' ------------------------------------------------------------------------------------
  Public Sub SetDirty(ByVal bValue As Boolean)
    ' Only set it after initialisation
    If (bInit) Then
      ' Set or reset the dirty bit
      bDirty = bValue
      ' Also indicate that a revision note should be added
      bRevision = True
    End If
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   tbEdtMain_GotFocus
  ' Goal:   Make sure the find form knows where we are
  ' History:
  ' 29-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub tbEdtMain_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbEdtMain.GotFocus
    ' Set the textbox property of the find form
    objFind.Textbox = Me.tbEdtMain
  End Sub
  Private Sub tbEdtFeatures_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbEdtFeatures.GotFocus
    ' Set the textbox property of the find form
    objFind.Textbox = Me.tbEdtFeatures
  End Sub
  Private Sub tbEdtSyntax_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbEdtSyntax.GotFocus
    ' Set the textbox property of the find form
    objFind.Textbox = Me.tbEdtSyntax
  End Sub
  Private Sub tbMainPde_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbMainPde.GotFocus
    ' Set the textbox property of the find form
    objFind.Textbox = Me.tbMainPde
  End Sub
  Private Sub tbAddOrg_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbAddOrg.GotFocus
    ' Set the textbox property of the find form
    objFind.Textbox = Me.tbAddOrg
  End Sub
  Private Sub tbAddTrn_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbAddTrn.GotFocus
    ' Set the textbox property of the find form
    objFind.Textbox = Me.tbAddTrn
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   tbEdtMain_KeyPress
  ' Goal:   Capture the SPACE when applied to EdtMain
  ' History:
  ' 28-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub tbEdtMain_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbEdtMain.KeyPress
    Dim bIncrement As Boolean   ' Whether we should increment or not

    Try
      ' Make sure the editor is visible
      If (Not Me.TabControl1.SelectedTab Is Me.tpEdit) Then Exit Sub
      ' Is the editor loading?
      If (Not bEdtLoad) Then
        ' Should we increment or decrement?
        bIncrement = Not ControlPressed()
        ' Was "space" pressed?
        Select Case e.KeyChar
          Case "+"
            ' Try to select MORE constituents
            ShowConstituent(Me.tbEdtMain, False)
          Case "-"
            ' Try to select LESS constituents
            ShowConstituent(Me.tbEdtMain, True)
          Case " "
            ' (De)select the source or the goal of a coreference link
            SelectCoref(Me.tbEdtMain)
          Case "p", "s"
            ' Try to find the first constituent that points to me
            FindSource(GetSelectedConst())
          Case "n"
            ' Try to find the first constituent that points to me
            FindSource(GetSelectedConst(), "*PRN*|*PRD*")
          Case "i"
            ' Try to find the first constituent that points to me with "Inferred"
            FindSource(GetSelectedConst(), "*PRN*|*PRD*", False)
          Case "a"
            ' Try to go to the antecedent
            FindAntecedent(GetSelectedConst())
          Case ChrW(Keys.Escape)
            ' Clear all coreference selections
            ClearCoref(Me.tbEdtMain)
          Case Else
            ' Show user we cannot edit this text
            Status("Editing of the text here is not possible")
        End Select
        ' Show we have handled it
        e.Handled = True
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/KeyPress error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '' ------------------------------------------------------------------------------------
  '' Name:   tbEdtMain_MouseClick
  '' Goal:   Find out where we are and show word information
  '' History:
  '' 30-01-2014  ERK Created
  '' ------------------------------------------------------------------------------------
  'Private Sub tbEdtMain_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
  '  Dim ndxThis As XmlNode      ' Selected constituent
  '  Dim ndxLeaf As XmlNode      ' Leaf
  '  Dim strVern As String       ' Vernacular
  '  Dim strLemma As String = "" ' Lemma
  '  Dim strLabel As String = "" ' Label
  '  Dim strDef As String = ""   ' Definition
  '  Dim strFeats As String = "" ' Features
  '  Dim objHtml As New StringColl

  '  Try
  '    ndxThis = ndxCurrentNode
  '    If (ndxThis IsNot Nothing) Then
  '      Select Case ndxThis.Name
  '        Case "forest"
  '          Stop
  '        Case "eTree"
  '          ' Try to get vernacular
  '          ndxLeaf = ndxThis.SelectSingleNode("./descendant::eLeaf[@Type='Vern'][1]")
  '          If (ndxLeaf IsNot Nothing) Then
  '            strVern = ndxLeaf.Attributes("Text").Value
  '            ' Get the lemma and the definition
  '            If (LemmaOneWord(ndxThis, strLemma, strDef)) Then
  '              strLabel = ndxThis.Attributes("Label").Value
  '              strFeats = GetFeatVect(ndxThis, "l", "h")
  '              ' Build an explanation
  '              objHtml.Add("<html><body><table>")
  '              ' Put it all together
  '              objHtml.Add("<tr><td valign='top'>" & VernToEnglish(strVern) & "</td>" & _
  '                          "<td valign='top'><font size='1' color='blue'>" & strLabel & "</font></td><td valign='top'><font color='red'>" & strLemma & "</font></td>" & _
  '                          "<td valign='top'><font size='1' color='green'>" & strFeats.Replace("|", ", ") & "</font></td><td valign='top'>" & strDef & "</td></tr>")
  '              objHtml.Add("</table></body></html>")
  '              Me.wbWord.DocumentText = objHtml.Text
  '              Me.wbWord.Visible = True
  '            End If
  '          End If
  '        Case "eLeaf"
  '          Stop
  '      End Select
  '    End If
  '  Catch ex As Exception
  '    ' Show error
  '    HandleErr("frmMain/tbEdtMain_MouseClick error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
  '  End Try
  'End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   tbEdtMain_Resize
  ' Goal:   The graphics object needs to be re-initialized when the form is resized
  ' History:
  ' 18-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub tbEdtMain_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbEdtMain.Resize
    ' Reinitialize graphics
    InitGraphics(Me.tbEdtMain)
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   tbEdtMain_SelectionChanged
  ' Goal:   Determine the line number and the character position
  '         These values are made available for the whole form
  ' History:
  ' 26-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub tbEdtMain_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbEdtMain.SelectionChanged
    Try
      ' Is the editor loading?
      If (Not bEdtLoad) And (Not bColoring) Then
        ' Perform what is needed for the selection change in the editor
        EditorSelectionChanged(Me.tbEdtMain, Me.tbEdtSyntax, Me.tbMainPde, Me.tbEdtFeatures)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/EdtMain error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuEditGotoLine_Click
  ' Goal:   Go to the line selected by the user
  ' History:
  ' 15-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuEditGotoLine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditGotoLine.Click
    Dim intLine As Integer = 0  ' Line we want to go to
    Dim strBack As String       ' Return from [InputBox()]

    Try
      ' Check if we have lines loaded
      If (pdxList Is Nothing) Then
        ' Nothing to be shown
        Status("First load a text")
        Exit Sub
      End If
      ' Note the line where we are right now
      intLine = GetCurrentLine()
      strBack = InputBox("Line to jump to", "Goto line", intLine)
      ' Check our return
      If (strBack = "") OrElse (Not IsNumeric(strBack)) Then Exit Sub
      ' Get the number
      intLine = CInt(strBack)
      ' Try going there
      EditGotoLine(intLine)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuEditGotoLine error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   EditGotoLine
  ' Goal:   Go to the line selected by the user
  ' History:
  ' 15-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub EditGotoLine(ByVal intLine As Integer)
    Dim bStatus As Boolean = bEdtLoad

    Try
      ' Validate
      If (intLine < intSectFirst + 1 OrElse intLine > intSectLast + 1) Then
        ' Do as if we are loading
        bEdtLoad = True
        ' Change sections
        If (Not ShowSection(Me.tbEdtMain, LineToSection(intLine))) Then bEdtLoad = bStatus : Exit Sub
        ' Reset status
        bEdtLoad = bStatus
      End If
      If (intLine >= intSectFirst) AndAlso (intLine <= intSectLast + 1) Then
        ' Determine which line it is for us
        intLine -= intSectFirst + 1
        ' Jump selection to this line
        Me.tbEdtMain.SelectionStart = arLeft(intLine)
        ' Scroll to this line
        Me.tbEdtMain.ScrollToCaret()
      Else
        ' Something is still wrong
        Status("Could not jump to line [" & intLine & "]")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/EditGotoLine error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub


  ' ------------------------------------------------------------------------------------
  ' Name:   mnuEditGotoAuto_Click
  ' Goal:   Go to the automatically added reference with user-given number
  ' History:
  ' 15-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuEditGotoAuto_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditGotoAuto.Click
    Dim strStatFile As String     ' Output file for statistics
    Dim ndxThis As XmlNode        ' Working node
    Dim intId As Integer          ' Id to start looking from (last read)
    Dim intI As Integer           ' Counter
    Dim intNum As Integer         ' Total number of automatically added nodes
    Dim intPos As Integer         ' Current number of automatically added node
    Dim intAutoNum As Integer = 0 ' Number of the auto stuff we want to go to
    Dim dtrFound() As DataRow     ' Result of selection

    Try
      ' Check if we have lines loaded
      If (pdxList Is Nothing) Then
        ' Nothing to be shown
        Status("First load a text")
        Exit Sub
      End If
      ' Make sure interrupt is not on
      bInterrupt = False
      ' Read the correct statistics file
      strStatFile = AutoStatInit(strCurrentFile, intCurrentSection, False)
      ' Get all the automatically added coreference links
      dtrFound = tdlAutoStat.Tables("Stat").Select("LinkType LIKE 'Auto*'")
      ' Determine the total number of automatically added nodes
      intNum = dtrFound.Length
      ' Initially assume position 1
      intPos = 1
      ' Get the <eTree> node we are at right now
      ndxThis = CurrentItem()
      ' Has something been selected?
      If (ndxThis Is Nothing) Then
        ' Nothing selected, so take the first one
        intId = -1
      Else
        ' Something selected, so look from here on
        intId = ndxThis.Attributes("Id").Value
        ' find out where this Id is in the [dtrFound] table
        For intI = 0 To dtrFound.Length - 1
          ' Is this the closest one?
          If (dtrFound(intI).Item("NodeId") >= intId) Then
            ' Okay suggest this one
            intPos = intI
            Exit For
          End If
        Next intI
      End If
      ' Note the line where we are right now
      intAutoNum = InputBox("Auto number to jump to", "Goto auto", intPos)
      ' Validate the answer
      If (intAutoNum >= 0) AndAlso (intAutoNum < intNum) Then
        ' Show where we are
        CorefProgress(intAutoNum & "/" & intNum)
        ' The number is within range, so start looking from there for the first possibility
        If (Not GotoFirstAuto(dtrFound, intAutoNum, ndxThis)) Then
          ' We've had them all
          Status("No more automatically added nodes to check!")
        End If
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/EditGotoAuto error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuEditFind_Click
  ' Goal:   Find something within the active window
  ' History:
  ' 06-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuEditFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditFind.Click
    ' Open the find dialog form
    objFind.Show()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuEditFindNext_Click
  ' Goal:   Find the next occurrance
  ' History:
  ' 21-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuEditFindNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditFindNext.Click
    objFind.DoFind(True)
  End Sub

  ' ------------------------------------------------------------------------------------------------------
  ' Name:   mnuEditFindXpath_Click
  ' Goal:   Find an element using an Xpath query, and taking the currently selected item as starting point
  ' History:
  ' 19-03-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------
  Private Sub mnuEditFindXpath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditFindXpath.Click
    Try
      ' Open the find dialog form
      objFindX.Show()
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuEditFindXpath error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuEditFeature_Click
  ' Goal:   Open de feature editing form
  ' History:
  ' 01-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuEditFeature_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditFeature.Click
    Dim ndxThis As Xml.XmlNode  ' The currently selected node

    Try
      ' Find out what the currently selected node is
      Select Case Me.TabControl1.SelectedTab.Name
        Case Me.tpEdit.Name
          ndxThis = ndxCurrentNode
        Case Me.tpTree.Name
          ndxThis = ndxCurrentTreeNode
        Case Else
          Status("Go to tab page [Edit] or [Tree]")
          Exit Sub
      End Select
      'ndxThis = CurrentItem()
      ' Validate
      If (ndxThis Is Nothing) Then
        ' Show that we first need to select a node
        Status("First select a node")
      Else
        With objFeat
          ' Show the form
          .Show()
          ' Set the node
          .Node = ndxThis
          ' Wait while it is visible
          While .Visible
            ' Do whatever needs doing
            Application.DoEvents()
          End While
          ' See whether changes were made
          If (.Changed) Then
            ' The features need to be shown again
            ShowFeat(Me.tbEdtFeatures)
            ' Make sure we are marked as dirty
            SetDirty(True)
          End If
        End With
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuEditFeature error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   FeatShowAdapt
  ' Goal:   Make sure the features of the correct node are being shown
  ' History:
  ' 07-11-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub FeatShowAdapt(ByRef ndxThis As XmlNode)
    Try
      ' Validate
      If (ndxThis Is Nothing) Then Exit Sub
      ' Access the feature show
      With objFeat
        ' Are we visible?
        If (.Visible) Then
          ' Change the node
          .Node = ndxThis
        End If
      End With
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/FeatShowAdapt error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMaxent_Click
  ' Goal:   Train maximum entropy coreference resolution based on the files in the
  '           indicated directory
  ' History:
  ' 07-11-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsMaxent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMaxent.Click
    Dim strDirIn As String = ""   ' input directory
    Dim strInFile As String       ' One input file
    Dim arInFile() As String      ' Array of input files
    Dim intI As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage

    Try
      ' Validate
      If (Not bInit) Then Exit Sub
      ' Check the settings
      If (Not CheckSettings()) Then Exit Sub
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      Else
        ' Start automatic machine now
        bAutoBusy = True
      End If
      ' Make sure interrupt is not on
      bInterrupt = False
      ' Let user select the directory where the input files are located
      With frmDirs
        ' Initialise the directoryes
        .SrcDir = strWorkDir
        ' Make sure user selects only one directory
        .AskOneDir()
        ' Elicit the directory
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Yes, we get a directory
            strDirIn = .SrcDir
          Case Else
            ' We better leave
            Status("Did not execute Tools/Maxent")
            Exit Sub
        End Select
      End With
      ' Initialise maximum entropy training
      If (Not InitTrain()) Then
        Status("Could not initialize training module")
        Exit Sub
      End If
      '' Initialise the stack of actions
      'ActionInit()
      ' Set the richtextbox local copy
      InitAutoCoref(Me.tbEdtMain)
      ' Read the chain dictionary
      TryReadChainDict()
      ' Clear the appropriate textbox
      Me.tbAutoLog.Text = ""
      ' Switch on max ent training
      bMaxEntTraining = True
      ' Double check existence
      If (IO.Directory.Exists(strDirIn)) Then
        ' Consider all *.psdx input files
        arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx")
        For intI = 0 To arInFile.Count - 1
          ' Show where we are
          intPtc = (intI + 1) * 100 \ arInFile.Count
          Status("Processing " & intPtc & "%", intPtc)
          ' Get the input file name
          strInFile = arInFile(intI)
          ' Determine coreference information automatically, starting from the point we are right now...
          If (Not DoTrainCoref(strInFile)) Then
            ' This is as far as I have gotten...
            MsgBox("There is an error in the maximum entropy training")
          End If
        Next intI
      End If
      ' Switch to this tabpage
      Me.TabControl1.SelectedTab = Me.tpReport
      ' Switch off max ent training
      bMaxEntTraining = False
      ' Finish the autocoreferencing mode
      FinishAutoCoref()
      ' Switch off automatic machine
      bAutoBusy = False
      ' Show we are ready
      Status("Done")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsMaxent error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Switch off max ent training
      bMaxEntTraining = False
      ' Finish the autocoreferencing mode
      FinishAutoCoref()
      ' Switch off automatic machine
      bAutoBusy = False
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsAuto_Click
  ' Goal:   Semi-automatically resolve coreferences as far as possible
  ' History:
  ' 26-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsAuto_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsAuto.Click
    Try
      bTraining = False
      DoCorefResolution()
      bTraining = False
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsAuto error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsTrain_Click
  ' Goal:   Build up training data for Semi-automatic coreference resolution
  ' History:
  ' 01-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsTrain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsTrain.Click
    Try
      bTraining = True
      DoCorefResolution()
      bTraining = False
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsTrain error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   DoCorefResolution
  ' Goal:   Semi-automatically resolve coreferences as far as possible
  ' History:
  ' 26-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub DoCorefResolution()
    Dim strStatFile As String ' Output file for statistics

    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) OrElse (Not bInit) Then Status("Not initialized yet or no file loaded") : Exit Sub
      ' Check the settings
      If (Not CheckSettings()) Then Exit Sub
      ' Initialise training if needed
      If (bTraining) Then
        If (Not InitTraining("state")) Then Logging("Could not initialise training") : Exit Sub
        If (Not InitTraining("coref")) Then Logging("Could not initialise training") : Exit Sub
      End If
      ' Go to the correct tabpage
      Me.TabControl1.SelectedTab = Me.tpEdit
      ' Start automatic machine and make sure we are not in FULL auto mode (semi-auto only)
      bAutoBusy = True : bIsFullyAuto = False
      ' Set the editor to coreference mode
      SetRefEditMode(True) : Logging("Reference-editing mode is now switched on (See Reference/ReferenceEditing).")
      ' Make sure interrupt is not on
      bInterrupt = False
      ' Initialise the statistics report for this file/section combination
      strStatFile = AutoStatInit(strCurrentFile, intCurrentSection)
      ' Initialise the stack of actions
      ActionInit()
      ' Set the richtextbox local copy
      InitAutoCoref(Me.tbEdtMain)
      ' Read the chain dictionary
      TryReadChainDict()
      ' Clear the appropriate textbox
      Me.tbAutoLog.Text = ""
      ' Determine coreference information automatically, starting from the point we are right now...
      If (Not DoAutoCoref()) Then
        ' Have we been interrupted?
        If (bInterrupt) Then
          ' Switch off interruption
          bInterrupt = False
          ' Just give the status
          Status("DoCorefResolution: Autocoreferencing stopped due to INTERRUPT")
        Else
          ' This is as far as I have gotten...
          Status("DoCorefResolution: There is an error in the automatic coreferencing")
        End If
      End If
      ' Finish the autocoreferencing mode
      FinishAutoCoref()
      ' Save the chain dictionary
      TrySaveChainDict()
      ' Switch off automatic machine
      bAutoBusy = False
      Status("Saving statistics")
      If (Not AutoStatFinish(strStatFile)) Then
        Logging("Could not save autostat report at " & strStatFile)
      End If
      ' Save training data
      If (bTraining) Then
        If (Not SaveTrainingCoref("state")) Then
          Logging("Could not save training data")
        End If
        If (Not SaveTrainingCoref("coref")) Then
          Logging("Could not save training data")
        End If
      End If
      ' Finish the action report for this file
      FinishActionReport()
      Status("Ready")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/DoCorefResolution error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Also finish the autocoreferencing mode
      FinishAutoCoref()
      ' Switch off automatic machine
      bAutoBusy = False : bTraining = False
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   FinishActionReport
  ' Goal:   Finish the action report for this section
  ' History:
  ' 26-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub FinishActionReport()
    Dim strReport As String   ' The actual report
    Dim strName As String     ' Name of file
    Dim strDirStat As String  ' Directory name
    Dim strFile As String     ' Ouptut file name

    Try
      ' Get the report
      strReport = ActionReport()
      ' Validate
      If (strReport = "") Then Exit Sub
      ' Finish the action report for this file
      strName = IO.Path.GetFileNameWithoutExtension(strCurrentFile)
      strDirStat = IO.Path.GetDirectoryName(strCurrentFile) & "\stat"
      strFile = strDirStat & "\" & strName & "-" & intCurrentSection + 1 & "-action.htm"
      ' Check if this report already exists
      If (IO.File.Exists(strFile)) Then
        ' append the report
        Status("Appending action report")
        IO.File.AppendAllText(strFile, strReport)
      Else
        ' Save it
        Status("Saving action report")
        IO.File.WriteAllText(strFile, strReport)
      End If
      Status("Done")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/FinishActionReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsFullAuto_Click
  ' Goal:   Automatically resolve coreferences as far as possible for all files in a directory
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsFullAuto_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsFullAuto.Click
    Dim strDirIn As String = ""   ' input directory
    Dim strDirOut As String = ""  ' output directory
    Dim strInFile As String       ' One input file
    Dim strOutFile As String      ' One output file
    Dim strDirPsd As String = ""  ' The PSD directory
    Dim strDirStat As String = "" ' The directory for statistics
    Dim strName As String         ' The name of the file
    Dim strPeriod As String = ""  ' The period we are dealing with
    Dim arInFile() As String      ' Array of input files
    Dim intI As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage

    Try
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      End If
      ' Check if the PeriodDefinition file is known -- otherwise don't allow File/Open
      If (Not CheckPeriodFile()) Then Exit Sub
      ' Let user select the directory where the input files are located
      With frmDirs
        ' Initialise the directoryes
        .SrcDir = strWorkDir
        .DstDir = .SrcDir & "\FullAuto"
        ' Elicit the directory
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Yes, we get a directory
            strDirIn = .SrcDir
            strDirOut = .DstDir
            ' Make sure the output directory exists
            If (Not IO.Directory.Exists(strDirOut)) Then IO.Directory.CreateDirectory(strDirOut)
            ' Check whether something is already present in the output directory
            If (IO.Directory.GetFiles(strDirOut).Length > 0) Then
              ' There are already files in there, so ask if user wants to overwrite
              Select Case MsgBox("There are already files in the output directory " & strDirOut & vbCrLf & _
                                 "Would you like to overwrite the existing output?", MsgBoxStyle.YesNoCancel)
                Case MsgBoxResult.Yes
                  ' No problem, just continue
                  bDoOverwrite = True
                Case MsgBoxResult.No
                  ' We may not overwrite
                  bDoOverwrite = False
                Case MsgBoxResult.Cancel
                  ' Exit
                  Exit Sub
              End Select
            End If
          Case Else
            ' We better leave
            Status("Did not execute Tools/FullAuto")
            Exit Sub
        End Select
      End With
      ' Double check existence
      If (IO.Directory.Exists(strDirIn)) Then
        ' Consider all *.psdx input files
        arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx")
        ' Check the period for each file
        For intI = 0 To arInFile.Count - 1
          ' Check if this file has its period available
          If (Not HasPeriod(arInFile(intI), strPeriod)) Then
            ' Request the period for this file
            If (Not AskUserPeriod(arInFile(intI), strPeriod)) Then Exit Sub
          End If
        Next intI
        ' Make directories for Statistics result and for PSD output
        strDirStat = strDirOut & "\stat"
        If (Not IO.Directory.Exists(strDirStat)) Then IO.Directory.CreateDirectory(strDirStat)
        strDirPsd = strDirOut & "\psd"
        If (Not IO.Directory.Exists(strDirPsd)) Then IO.Directory.CreateDirectory(strDirPsd)
        ' Perform fully auto for each file
        For intI = 0 To arInFile.Count - 1
          ' Show where we are
          intPtc = (intI + 1) * 100 \ arInFile.Count
          Status("FullAuto " & intPtc & "%", intPtc)
          ' Get the input file name
          strInFile = arInFile(intI)
          ' Get the name of the file
          strName = IO.Path.GetFileNameWithoutExtension(strInFile)
          ' Determine the output file name
          strOutFile = strDirOut & "\" & strName & ".psdx"
          ' Does this output already exist?
          If (Not bDoOverwrite) AndAlso (IO.File.Exists(strOutFile)) Then
            ' We must not overwrite, and the output already exists
            ' Show where we are in log function
            Logging("FullAuto " & intI + 1 & "/" & arInFile.Count & " " & _
                    strName & " - skipped")
          Else
            ' Show where we are in log function
            Logging("FullAuto " & intI + 1 & "/" & arInFile.Count & " " & _
                    strName)
            ' Initialise the stack of actions
            ActionInit()
            ' Initialise profiling
            ProfileInit()
            ' Read the chain dictionary
            TryReadChainDict()
            ' Call the correct function
            ' (AutoStat is done for each section of the input file!!!)
            If (Not OneFullAutoFile(strInFile, strOutFile)) Then
              ' Was interrupt pressed?
              If (bInterrupt) Then
                ' Reset interrupt
                bInterrupt = False
                Exit For
              End If
              ' There was an error, so stop
              MsgBox("There is an error in fully automatic coreference resolution")
              Exit Sub
            End If
          End If
          ' Check to see if any action has taken place
          If (IsAction()) Then
            ' Finish the action report for this file
            strOutFile = strDirStat & "\" & strName & "-action.htm"
            Status("Saving action report")
            IO.File.WriteAllText(strOutFile, ActionReport)
            ' Save the profile information for this file
            strOutFile = strDirStat & "\" & strName & "-profile.htm"
            Status("Saving profile report")
            IO.File.WriteAllText(strOutFile, ProfileOview)
            ' Save the file in PSD form
            strOutFile = strDirPsd & "\" & strName & "-cesax.psd"
            Status("Saving PSD output")
            If (Not DoExport("PsdOutput", strOutFile)) Then
              If (bInterrupt) Then
                ' Reset the interrupt
                bInterrupt = False
                ' Show user
                Status("Save PSD is aborted due to interrupt")
                Exit Sub
              Else
                Logging("Could not save PSD output to " & strOutFile)
              End If
            End If
            ' Try to save the chain dictionary
            TrySaveChainDict()
          End If
        Next intI
      End If
      ' Show what we are doing
      Status("Making report...")
      ' Calculate this report and put it in place
      Me.wbReport.DocumentText = "Full automatic resolving is ready (without report)."
      ' Switch to this tabpage
      Me.TabControl1.SelectedTab = Me.tpReport
      ' Show we are ready
      Status("Done")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsFullAuto_Click error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Also finish the autocoreferencing mode
      FinishAutoCoref()
      ' Switch off automatic machine
      bAutoBusy = False
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   DirHasStat
  ' Goal:   Check if the indicated directory has *-stat.xml files
  ' History:
  ' 22-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function DirHasStat(ByVal strIn As String) As Boolean
    Try
      ' Validate
      If (strIn = "") Then Return False
      If (Not IO.Directory.Exists(strIn)) Then Return False
      ' Check if this directory has statistics information XML files
      Return (IO.Directory.GetFiles(strIn, "*stat*.xml").Length > 0)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/DirHasStat error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsFullStat_Click
  ' Goal:   Give statistics of Fully Automatically resolved coreference resolution
  '           based on the stat/*-stat.xml files in the indicated directory
  ' History:
  ' 21-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsFullStat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsFullStat.Click
    Dim strDirIn As String = ""   ' input directory
    Dim strDirOut As String = ""  ' output directory
    Dim strInFile As String       ' One input file
    Dim strOutFile As String      ' One output file
    Dim strReport As String       ' The HTML content of the report
    Dim arInFile() As String      ' Array of input files
    Dim intI As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage

    Try
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      End If
      ' Initialise AutoStatistics report
      InitFullStat()
      ' Let user select an input and output directory
      With frmDirs
        ' Initialise the directories myself
        .SrcDir = strWorkDir
        ' Make sure user selects only one directory
        .AskOneDir()
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Get the source directory: check which directory has the correct files
            If (DirHasStat(.SrcDir)) Then
              strDirIn = .SrcDir
            ElseIf (DirHasStat(.SrcDir & "\FullAuto\stat")) Then
              strDirIn = .SrcDir & "\FullAuto\stat"
            ElseIf (DirHasStat(.SrcDir & "\FullAuto")) Then
              strDirIn = .SrcDir & "\FullAuto"
            ElseIf (DirHasStat(.SrcDir & "\stat")) Then
              strDirIn = .SrcDir & "\stat"
            Else
              ' Unable to get the directory...
              strDirIn = ""
            End If
            ' Check whether correct directory was selected
            If (strDirIn = "") Then
              ' Warn the user
              MsgBox("Could not find statistics information XML files under " & .SrcDir)
              ' Return failure
              Status("Full Auto statistics have not been determined")
              Exit Sub
            Else
              ' Consider all *-stat.xml input files
              arInFile = IO.Directory.GetFiles(strDirIn, "*-stat.xml")
              For intI = 0 To arInFile.Count - 1
                ' Show where we are
                intPtc = (intI + 1) * 100 \ arInFile.Count
                Status("Processing " & intPtc & "%", intPtc)
                ' Get the input file name
                strInFile = arInFile(intI)
                ' Show where we are
                Logging("Full Auto statistics " & intI + 1 & "/" & arInFile.Count & " - " & _
                        IO.Path.GetFileNameWithoutExtension(strInFile))
                ' Call the correct function
                If (Not OneFullStatFile(strInFile)) Then
                  ' There was an error, so stop
                  Status("There was an error in determining Full Auto statistics")
                  Exit Sub
                End If
              Next intI
            End If
        End Select
      End With
      ' Calculate this report and put it in place
      strReport = FullStatReport()
      Me.wbReport.DocumentText = strReport
      ' Switch to this tabpage
      Me.TabControl1.SelectedTab = Me.tpReport
      ' Save the report on an appropriate location
      strOutFile = strDirIn & "\FullAutoStatistics.htm"
      IO.File.WriteAllText(strOutFile, strReport)
      ' Show there are some errors
      Status("Ready determining Full Auto Statistics. Results are in: " & strOutFile)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsFullStat_Click error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsAutoReport_Click
  ' Goal:   Give a report of the coreferences that were made automatically thus far
  ' History:
  ' 28-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsAutoReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsAutoReport.Click
    Try
      ' Open the appropriate tabpage
      Me.TabControl1.SelectTab(Me.tpReport)
      ' get the autocoref report
      Me.wbReport.DocumentText = ActionReport()
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsAutoReport_Click error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsVerify_Click
  ' Goal:   Make a full overview of all the coreferences, so that the user can verify
  '           whether these are correct
  ' History:
  ' 21-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsVerify_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsVerify.Click
    Dim strDirIn As String = ""   ' input directory
    Dim strOutFile As String      ' One output file
    Dim strReport As String       ' The HTML content of the report

    Try
      ' Open the appropriate tabpage
      Me.TabControl1.SelectTab(Me.tpReport)
      ' Make the verification overview of the currently loaded file...
      strReport = VerifyReport()
      Me.wbReport.DocumentText = strReport
      ' Switch to this tabpage
      Me.TabControl1.SelectedTab = Me.tpReport
      ' Go to the correct directory
      strOutFile = strWorkDir & "\stat"
      ' Check if directory exists
      If (Not IO.Directory.Exists(strOutFile)) Then IO.Directory.CreateDirectory(strOutFile)
      ' Save the report on an appropriate location
      strOutFile &= "\" & Me.tbName.Text & "-Overview.htm"
      IO.File.WriteAllText(strOutFile, strReport)
      ' Show there are some errors
      Status("Ready producing a verification report. Results are in: " & strOutFile)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsVerify_Click error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsNextAuto_Click
  ' Goal:   Go to the next automatically made coreference, in order to check it
  ' History:
  ' 26-11-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsNextAuto_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsNextAuto.Click
    Dim strStatFile As String ' Output file for statistics
    Dim ndxThis As XmlNode    ' Working node
    Dim intId As Integer      ' Id to start looking from (last read)
    Dim intNum As Integer     ' Total number of automatically added nodes
    Dim intPos As Integer     ' Current number of automatically added node
    Dim dtrFound() As DataRow ' Result of selection

    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) OrElse (Not bInit) Then Exit Sub
      ' Go to the correct tabpage
      Me.TabControl1.SelectedTab = Me.tpEdit
      ' Make sure interrupt is not on
      bInterrupt = False
      ' Read the correct statistics file
      strStatFile = AutoStatInit(strCurrentFile, intCurrentSection, False)
      ' Determine the total number of automatically added nodes
      intNum = tdlAutoStat.Tables("Stat").Select("LinkType LIKE 'Auto*'").Length
      ' Get the <eTree> node we are at right now
      ndxThis = CurrentItem()
      ' Has something been selected?
      If (ndxThis Is Nothing) Then
        ' Nothing selected, so take the first one
        intId = -1
      Else
        ' Something selected, so look from here on
        intId = ndxThis.Attributes("Id").Value
      End If
      ' Find the next Id in the AutoStat table
      dtrFound = tdlAutoStat.Tables("Stat").Select("NodeId>" & intId & " AND LinkType LIKE 'Auto*'", "NodeId ASC")
      ' Determine the number where we are with checking
      intPos = intNum - dtrFound.Length
      ' Show where we are
      CorefProgress(intPos & "/" & intNum)
      If (Not GotoFirstAuto(dtrFound, 0, ndxThis)) Then
        ' We've had them all
        Status("No more automatically added nodes to check!")
      End If

    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsNextAuto_Click error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   GotoFirstAuto
  ' Goal:   Go to the first automatically possible point in [dtrFound]
  '         Start counting from [intStart]
  '         Don't end up at [ndxLast]
  ' History:
  ' 30-11-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GotoFirstAuto(ByRef dtrFound() As DataRow, ByVal intStart As Integer, _
                                 ByRef ndxLast As XmlNode) As Boolean
    Dim ndxThis As XmlNode  ' Working node
    Dim bCheck As Boolean   ' Flag to see if this one shold be treated
    Dim intId As Integer    ' Id of the node we are checking

    Try
      ' Go through these, and find the first one that meets criteria
      For intI = intStart To dtrFound.Length - 1
        With dtrFound(intI)
          ' Is this one we should verify?
          bCheck = False
          Select Case .Item("LinkType")
            Case "AutoLocal"
              ' Check for reference type
              Select Case .Item("RefType")
                Case strRefNew, strRefNewVar, strRefInert, strRefInferred ' No need to check
                Case Else           ' Check all other types!
                  bCheck = True
              End Select
            Case "Auto"
              ' Check for reference type
              Select Case .Item("RefType")
                Case strRefNew, strRefNewVar, strRefInferred  ' No need to check
                Case strRefIdentity, strRefCross
                  ' These definitely need verification
                  bCheck = True
                Case Else   ' Check all other types!
                  ' Warn the user of the category
                  MsgBox("ToolsNextAuto Auto/Reftype=" & .Item("RefType"))
                  bCheck = True
              End Select
          End Select
          ' Need for checking?
          If (bCheck) Then
            ' Go to this item
            intId = .Item("NodeId")
            ndxThis = IdToNode(intId)
            ' Additional check: if we are *PRN*, then no checking is needed
            If (Not DoLike(ndxThis.Attributes("Label").Value, "*PRN*|*-1|*-2|*-3|*-4|*-5")) Then
              ' Is this a coreferencing node?
              If (CorefDstId(ndxThis) < 0) Then
                ' Is this the one we should not end up with?
                If (Not EqualNodes(ndxThis, ndxLast)) Then
                  ' Go to this node
                  GotoNode(ndxThis)
                  ' Leave!
                  Return True
                End If
              Else
                ' If this is 1st person to 1st person coreference, then we don't need to check
                Dim strMyPGN As String = GetFeature(ndxThis, "NP", "PGN")
                Dim strAntPGN As String = GetFeature(CorefDst(ndxThis), "NP", "PGN")

                If (Not DoLike(Strings.Left(strMyPGN, 1), "1|2")) OrElse _
                   (Strings.Left(strAntPGN, 1) <> Strings.Left(strMyPGN, 1)) Then
                  ' Is this the one we should not end up with?
                  If (Not EqualNodes(ndxThis, ndxLast)) Then
                    ' Go to this node
                    GotoNode(ndxThis)
                    ' Leave!
                    Return True
                  End If
                End If
              End If
            End If
          End If
        End With
      Next intI
      ' We have not actually gone somewhere
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/GotoFirstAuto error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsLearn_Click
  ' Goal:   Learn from previous coreferencing
  ' History:
  ' 10-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsLearn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsLearn.Click
    Dim strDirIn As String = ""   ' input directory
    Dim strInFile As String       ' One input file
    Dim arInFile() As String      ' Array of input files
    Dim intI As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage

    Try
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      End If
      ' Let user select the directory where the input files are located
      With frmDirs
        ' Initialise the directoryes
        .SrcDir = strWorkDir & "\Adapted"
        ' Make sure user selects only one directory
        .AskOneDir()
        ' Elicit the directory
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Yes, we get a directory
            strDirIn = .SrcDir
          Case Else
            ' We better leave
            Status("Did not execute Tools/Learn")
            Exit Sub
        End Select
      End With
      ' Initialise Learning report
      InitLearn()
      ' Double check existence
      If (IO.Directory.Exists(strDirIn)) Then
        ' Consider all *.psdx input files
        arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx")
        For intI = 0 To arInFile.Count - 1
          ' Show where we are
          intPtc = (intI + 1) * 100 \ arInFile.Count
          Status("Processing " & intPtc & "%", intPtc)
          ' Get the input file name
          strInFile = arInFile(intI)
          ' Determine coreference information automatically, starting from the point we are right now...
          If (Not DoLearnCoref(strInFile)) Then
            ' This is as far as I have gotten...
            MsgBox("There is an error in the coreference learning")
          End If
        Next intI
      End If
      ' Show what we are doing
      Status("Making report...")
      ' Calculate this report and put it in place
      Me.wbReport.DocumentText = LearnReport()
      ' Switch to this tabpage
      Me.TabControl1.SelectedTab = Me.tpReport
      ' Show we are ready
      Status("Done")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsLearn_Click error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsEnglish_Click
  ' Goal:   Produces PSD files from all the TXT files in the selected directory
  ' History:
  ' 25-03-2013  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsEnglish_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsEnglish.Click
    Dim strLang As String = "English"  ' Language

    Try
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      End If
      ' Start the process
      If (BatchConvertTxtToPsd(strLang)) Then
        ' Show we are done
        Status("Ready producing PSD files.")
      Else
        ' Show there was a problem
        Status("There was a problem producing PSD files for [" & strLang & "]")
      End If
      ' Make sure we indicate nothing has changed
      SetDirty(False)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsEnglish error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsTxtToPsd_Click
  ' Goal:   Produces PSD files from all the TXT files in the selected directory
  ' History:
  ' 25-03-2013  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsTxtToPsd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsTxtToPsd.Click
    Dim strLang As String = ""  ' Language

    Try
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      End If
      ' Determine the language
      With frmConvert
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            strLang = .Language
          Case Else
            Exit Sub
        End Select
      End With
      ' Start the process
      If (BatchConvertTxtToPsd(strLang)) Then
        ' Show we are done
        Status("Ready producing PSD files.")
      Else
        ' Show there was a problem
        Status("There was a problem producing PSD files for [" & strLang & "]")
      End If
      ' Make sure we indicate nothing has changed
      SetDirty(False)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsTxtToPsd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '' ---------------------------------------------------------------------------------------------------------
  '' Name:   mnuToolsSpanish_Click
  '' Goal:   Produces PSDX files from all the TXT files in the selected directory
  '' History:
  '' 16-08-2013  ERK Created
  '' ---------------------------------------------------------------------------------------------------------
  'Private Sub mnuToolsSpanish_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  '  Try
  '    ' Check if automachine is busy
  '    If (bAutoBusy) Then
  '      MsgBox("First stop auto coreference resolution (F11)")
  '      Exit Sub
  '    End If
  '    If (BatchConvertConllToPsdx("Spanish")) Then
  '      ' Show we are done
  '      Status("Ready producing PSDX files.")
  '    Else
  '      ' Show there was a problem
  '      Status("There was a problem producing PSDX files for Spanish")
  '    End If
  '    ' Make sure we indicate nothing has changed
  '    SetDirty(False)
  '  Catch ex As Exception
  '    ' Show error
  '    HandleErr("frmMain/mnuToolsSpanish error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
  '  End Try
  'End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsConLLX_Click
  ' Goal:   Produces PSDX files from all the TXT files in the selected directory
  ' History:
  ' 16-08-2013  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsConLLX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsConLLX.Click
    Try
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      End If
      If (BatchConvertConllToPsdx()) Then
        ' Show we are done
        Status("Ready producing PSDX files.")
      Else
        ' Show there was a problem
        Status("There was a problem producing PSDX files from CONLL-X")
      End If
      ' Make sure we indicate nothing has changed
      SetDirty(False)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsConLLX error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsAlpinoRead_Click
  ' Goal:   Produces PSDX files from all the alpino .XML files in the selected directory structure
  '         Assumption: each 'text' has a director with the name of the text 
  '           that contains one .xml file for each sentence
  ' History:
  ' 29-06-2015  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsAlpinoRead_Click(sender As System.Object, e As System.EventArgs) Handles mnuToolsAlpinoRead.Click
    Dim strLang As String = ""  ' Language of this corpus

    Try
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      End If
      bInterrupt = False
      ' Determine the language
      With frmConvert
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            strLang = .Language
          Case Else
            Exit Sub
        End Select
      End With
      ' Start conversion
      If (BatchConvertToPsdx("alpino", LangToEthno(strLang))) Then
        ' Show we are done
        Status("Ready producing PSDX files.")
      Else
        ' Show there was a problem
        Status("There was a problem producing PSDX files from the Alpino (ds) format")
      End If
      ' Make sure we indicate nothing has changed
      SetDirty(False)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsAlpinoRead error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsFoliaRead_Click
  ' Goal:   Produces PSDX files from all the FOLIA.XML files in the selected directory
  '         This extracts different files from the folia file, if it contains different texts
  ' History:
  ' 17-02-2014  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsFoliaRead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsFoliaRead.Click
    Dim strLang As String = ""  ' Language of this corpus

    Try
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      End If
      bInterrupt = False
      ' Determine the language
      With frmConvert
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            strLang = .Language
          Case Else
            Exit Sub
        End Select
      End With
      ' Start conversion
      If (BatchConvertToPsdx("folia", LangToEthno(strLang))) Then
        ' Show we are done
        Status("Ready producing PSDX files.")
      Else
        ' Show there was a problem
        Status("There was a problem producing PSDX files from the Folia format")
      End If
      ' Make sure we indicate nothing has changed
      SetDirty(False)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsFoliaRead error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsFoliaWrite_Click
  ' Goal:   Produces Folia files from all the Psdx files in the selected directory
  ' History:
  ' 12-03-2014  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsFoliaWrite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsFoliaWrite.Click
    Dim strLang As String = ""  ' Language of this corpus

    Try
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      End If
      bInterrupt = False
      ' Determine the language
      With frmConvert
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            strLang = .Language
          Case Else
            Exit Sub
        End Select
      End With
      ' Start conversion
      If (BatchConvertToFolia("psdx", LangToEthno(strLang))) Then
        ' Show we are done
        Status("Ready producing Folia files.")
      Else
        ' Show there was a problem
        Status("There was a problem producing Folia files")
      End If
      ' Make sure we indicate nothing has changed
      SetDirty(False)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsFolia error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolAdelheitToXml_Click
  ' Goal:   Convert improved 5-column Adelheit to its XML format
  ' History:
  ' 01-04-2014  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub mnuToolAdelheitToXml_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolAdelheitToXml.Click
    Dim strFile As String = ""
    Dim strDstFile As String = ""

    Try
      ' Ask for the input file
      If (Not GetFileName(Me.OpenFileDialog1, strWorkDir, strFile, "Adelheid improved (*.txt)|*.txt")) Then Status("Could not open file") : Exit Sub
      ' Think of a destnation file
      strDstFile = IO.Path.GetDirectoryName(strFile) & "\" & IO.Path.GetFileNameWithoutExtension(strFile) & ".xml"
      ' Convert the file
      If (AdelheitToXml(strFile, strDstFile)) Then
        Status("Resulting XML is: " & strDstFile)
      Else
        Status("There was an error producing an XML from Adelheit")
      End If

    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsAdelheit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsFwToFolia_Click
  ' Goal:   Produces Folia files from all the Fieldwork XML files in the selected directory
  ' History:
  ' 27-03-2014  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsFwToFolia_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsFwToFolia.Click
    Dim strLang As String = ""  ' Language of this corpus

    Try
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      End If
      bInterrupt = False
      ' Determine the language
      With frmConvert
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            strLang = .Language
          Case Else
            Exit Sub
        End Select
      End With
      ' Start conversion
      If (BatchConvertToFolia("Flex", LangToEthno(strLang))) Then
        ' Show we are done
        Status("Ready producing Folia files.")
      Else
        ' Show there was a problem
        Status("There was a problem producing Folia files")
      End If
      ' Make sure we indicate nothing has changed
      SetDirty(False)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsFwToFolia error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
    ' ---------------------------------------------------------------------------------------------------------
    ' Name:   mnuToolsTiger_Click
    ' Goal:   Produces PSDX files from all the TIG files in the selected directory
    ' History:
    ' 27-01-2014  ERK Created
    ' ---------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsTiger_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsTiger.Click
    Dim strLang As String = ""  ' Language of this corpus

    Try
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      End If
      bInterrupt = False
      ' Determine the language
      With frmConvert
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            strLang = .Language
          Case Else
            Exit Sub
        End Select
      End With
      ' Load the definitions for the introduction
      If (strLang = "Dutch") Then
        ' Load the definitions
        If (Not LoadDutchCgnInfo()) Then Status("Could not load CGN info") : Exit Sub
      ElseIf (strLang = "German") Then
        ' Make and/or load the definitions
        If (Not LoadGermanTigerInfo()) Then Status("Could not load Tiger info") : Exit Sub
      End If
      ' Start conversion
      If (BatchConvertToPsdx("tiger", LangToEthno(strLang))) Then
        ' Show we are done
        Status("Ready producing PSDX files.")
      Else
        ' Show there was a problem
        Status("There was a problem producing PSDX files from the Tiger format")
      End If
      ' Make sure we indicate nothing has changed
      SetDirty(False)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsTiger error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsGerman_Click
  ' Goal:   Produces PSD files from all the CON or CONLL files in the selected directory
  ' History:
  ' 14-03-2013  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsGerman_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Try
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      End If
      If (BatchConvertTxtToPsd("German")) Then
        ' Show we are done
        Status("Ready producing PSD files.")
      Else
        ' Show there was a problem
        Status("There was a problem producing PSD files for German")
      End If
      ' Make sure we indicate nothing has changed
      SetDirty(False)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsGerman error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name:   BatchConvertTxtToPsd
  ' Goal:   Produces PSD files from all the TXT files in the selected directory
  '         Do this for language [strLanguage]
  ' History:
  ' 14-03-2013  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Function BatchConvertTxtToPsd(ByVal strLanguage As String) As Boolean
    Dim strDirIn As String = ""   ' input directory
    Dim strDirOut As String = ""  ' output directory
    Dim strInFile As String       ' One input file
    Dim strOutFile As String      ' One output file
    Dim arInFile() As String      ' Array of input files
    Dim intI As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage
    Dim bOverwrite As Boolean = False   ' Whether we may overwrite or not
    Dim bAsk As Boolean = True          ' Need to ask?

    Try
      ' Let user select an input and output directory
      With frmDirs
        ' Initialise the directories myself
        ' .SrcDir = strWorkDir
        .SrcDir = GetTableSetting(tdlSettings, "BatchConvertSrc", strWorkDir)
        .DstDir = GetTableSetting(tdlSettings, "BatchConvertDst", .SrcDir)
        '' If the directory does not exist, then make it
        'If (Not IO.Directory.Exists(.DstDir)) Then IO.Directory.CreateDirectory(.DstDir)
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Get the source and destination directories
            strDirIn = .SrcDir
            strDirOut = .DstDir
            ' Adapt table settings
            SetTableSetting(tdlSettings, "BatchConvertSrc", strDirIn)
            SetTableSetting(tdlSettings, "BatchConvertDst", strDirOut)
            ' Save the settings
            XmlSaveSettings(strSetFile)
            ' Check whether correct directories were selected
            If (IO.Directory.Exists(strDirIn)) AndAlso (IO.Directory.Exists(strDirOut)) Then
              ' Consider all *.psdx input files
              arInFile = IO.Directory.GetFiles(strDirIn, "*.txt")
              ' Do we have something?
              If (arInFile.Count = 0) Then
                ' Try STP
                arInFile = IO.Directory.GetFiles(strDirIn, "*.stp")
              End If
              For intI = 0 To arInFile.Count - 1
                ' Show where we are
                intPtc = (intI + 1) * 100 \ arInFile.Count
                Status("Processing " & intPtc & "%", intPtc)
                ' Get the input file name
                strInFile = arInFile(intI)
                ' Determine the output file name
                strOutFile = strDirOut & "\" & IO.Path.GetFileNameWithoutExtension(strInFile) & _
                  ".psd"
                ' Show where we are
                Logging("Producing PSD " & intI + 1 & "/" & arInFile.Count & " - " & IO.Path.GetFileNameWithoutExtension(strInFile))
                ' Check existence
                If (Not bOverwrite) AndAlso (IO.File.Exists(strOutFile)) AndAlso (bAsk) Then
                  ' Ask if overwriting is okay
                  If (MsgBox("Can I overwrite files like: [" & strOutFile & "]?", MsgBoxStyle.YesNoCancel) <> MsgBoxResult.Yes) Then
                    ' No overwriting, so exit 
                    ' Return False
                    bOverwrite = False : bAsk = False
                  Else
                    ' We can overwrite, so put to true
                    bOverwrite = True : bAsk = False
                  End If
                End If
                If (bOverwrite) OrElse (Not IO.File.Exists(strOutFile)) Then
                  ' Convert from text to PSD
                  If (Not ConvertOneTxtToPsd(strInFile, strOutFile, strLanguage, False)) Then
                    ' Check for interrupt
                    If (bInterrupt) Then Status("Interrupted with error...")
                    ' There was an error, so stop
                    Return False
                  Else
                    ' Check for interrupt
                    If (bInterrupt) Then Status("Interrupted by user...") : Return False
                  End If
                End If
              Next intI
            End If
        End Select
      End With
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/BatchConvertTxtToPsd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name:   BatchConvertConllToPsdx
  ' Goal:   Produces PSDX files from all the CON or CONLL files in the selected directory
  '         Do this for language [strLanguage]
  ' History:
  ' 16-08-2013  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Function BatchConvertConllToPsdx() As Boolean
    Dim strDirIn As String = ""     ' input directory
    Dim strDirOut As String = ""    ' output directory
    Dim strInFile As String         ' One input file
    Dim strOutFile As String        ' One output file
    Dim strLanguage As String = ""  ' Language
    Dim arInFile() As String        ' Array of input files
    Dim intI As Integer             ' Counter
    Dim intPtc As Integer           ' Percentage
    Dim bOverwrite As Boolean = False    ' Whether we may overwrite or not

    Try
      ' Let user select an input and output directory
      With frmDirs
        ' Initialise the directories myself
        .SrcDir = IIf(strLastConvert = "", strWorkDir, strLastConvert)
        .DstDir = .SrcDir
        '' If the directory does not exist, then make it
        'If (Not IO.Directory.Exists(.DstDir)) Then IO.Directory.CreateDirectory(.DstDir)
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Get the source and destination directories
            strDirIn = .SrcDir
            strDirOut = .DstDir
            ' Save this directory, if it is different from what we have
            If (strLastConvert <> strDirIn) Then
              strLastConvert = strDirIn
              SetTableSetting(tdlSettings, "LastConvertDir", strLastConvert)
              XmlSaveSettings(strSetFile)
            End If
            ' Derive the language somehow
            If (Not DeriveLanguage(strDirIn, strLanguage)) Then Return False
            ' Validate language
            Select Case strLanguage
              Case "Spanish", "es"  ' Okay, implemented
                strLanguage = "es"
              Case "Dutch", "nl"    ' Okay, implemented
                strLanguage = "nl"
              Case Else       ' Not implemented
                Status("Not implemented for language [" & strLanguage & "]")
                Return False
            End Select
            ' Check whether correct directories were selected
            If (IO.Directory.Exists(strDirIn)) AndAlso (IO.Directory.Exists(strDirOut)) Then
              ' Consider all *.con or *.conll input files
              arInFile = IO.Directory.GetFiles(strDirIn, "*.con")
              If (arInFile.Count = 0) Then
                arInFile = IO.Directory.GetFiles(strDirIn, "*.conll")
              End If
              ' Now continue, assuming all is implemented
              For intI = 0 To arInFile.Count - 1
                ' Show where we are
                intPtc = (intI + 1) * 100 \ arInFile.Count
                Status("Processing " & intPtc & "%", intPtc)
                ' Get the input file name
                strInFile = arInFile(intI)
                ' Determine the output file name, but WITHOUT EXTENSION yet
                strOutFile = strDirOut & "\" & IO.Path.GetFileNameWithoutExtension(strInFile)
                ' Show where we are
                Logging("Producing PSDX " & intI + 1 & "/" & arInFile.Count & " - " & IO.Path.GetFileNameWithoutExtension(strInFile))
                ' Check existence of final PSDX result
                If (Not bOverwrite) AndAlso (IO.File.Exists(strOutFile & ".psdx")) Then
                  ' Ask if overwriting is okay
                  If (MsgBox("Can I overwrite files like: [" & strOutFile & ".psdx" & "]?", MsgBoxStyle.YesNoCancel) <> MsgBoxResult.Yes) Then
                    ' No overwriting, so exit 
                    Return False
                  End If
                  ' We can overwrite, so put to true
                  bOverwrite = True
                End If
                ' Convert from Text to Conll7 to Psdx
                If (Not PrepareOneTxtToPsdx(strInFile, strOutFile, strLanguage, False)) Then
                  ' Check for interrupt
                  If (bInterrupt) Then Status("Interrupted with error...")
                  ' There was an error, so stop
                  Return False
                Else
                  ' Check for interrupt
                  If (bInterrupt) Then Status("Interrupted by user...") : Return False
                End If
              Next intI
            End If
        End Select
      End With
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/BatchConvertConllToPsdx error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      Return False
    End Try
  End Function

  ' ---------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsTokenize_Click
  ' Goal:   Do tokenization on existing psdx files
  ' History:
  ' 30-05-2015  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsTokenize_Click(sender As System.Object, e As System.EventArgs) Handles mnuToolsTokenize.Click
    Dim strDirIn As String = "" ' Input directory
    Dim arFile() As String      ' Array of input files
    Dim intI As Integer           ' Counter
    Dim strInFile As String       ' One input file
    Dim intPtc As Integer         ' Percentage

    Try
      ' Reset interrupt
      bInterrupt = False
      ' Ask for input files
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, "Directory with the .psdx files to tokenize", strLastConvert)) Then
        Status("Aborted")
        Exit Sub
      End If
      ' Validate
      If (Not IO.Directory.Exists(strDirIn)) Then Status("COuld not open directory") : Exit Sub
      ' Get all the files here
      arFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories)
      If (arFile.Length = 0) Then Status("There are no .psdx files in this directory") : Exit Sub
      ' Ask for confirmation
      Select Case MsgBox("DO you really want to (re-)tokenize .psdx files in [" & strDirIn & "]?", MsgBoxStyle.YesNo)
        Case MsgBoxResult.No
          Status("Aborted")
          Exit Sub
      End Select
      ' Get going
      For intI = 0 To arFile.Length - 1
        ' Where are we?
        intPtc = (100 * (intI + 1)) \ arFile.Length
        Status("Processing " & intPtc & "%", intPtc)
        strInFile = arFile(intI)
        ' Perform tokenization
        If (Not OneTokenizePsdx(strInFile)) Then Status("Could not tokenize [" + strInFile & "]") : Exit Sub
        ' Save the results
        pdxCurrentFile.Save(strInFile)
        ' Show we've done this one
        Logging(strInFile)
      Next intI
      Status("Ready tokenizing " & arFile.Length & " texts")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ToolsTokenize error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsBatch_Click
  ' Goal:   Produces PSDX files from all the PSD (or STP = stanford parser) files in the selected directory
  ' History:
  ' 22-05-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsBatch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsBatch.Click
    Dim strDirIn As String = ""   ' input directory
    Dim strDirOut As String = ""  ' output directory
    Dim strDirApp As String = ""  ' Part of input directory to be appended to output directory for this file
    Dim strInFile As String       ' One input file
    Dim strOutFile As String      ' One output file
    Dim strShort As String        ' Short file name
    Dim strLang As String         ' Language
    Dim strEthno As String        ' Ethno code
    Dim arInFile() As String      ' Array of input files
    Dim intI As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage
    Dim bOverwrite As Boolean = False    ' Whether we may overwrite or not
    Dim bAsked As Boolean = False         ' Asked user?

    Try
      ' Check if automachine is busy
      If (bAutoBusy) Then Status("First stop auto coreference resolution (F11)") : Exit Sub
      ' Initialize
      arInFile = Nothing
      ' Let user select an input and output directory
      With frmDirs
        ' Initialise the directories myself
        .SrcDir = GetTableSetting(tdlSettings, "LastConvertDir", strWorkDir)
        .DstDir = GetTableSetting(tdlSettings, "LastConvertDirDst", .SrcDir)
        '' If the directory does not exist, then make it
        'If (Not IO.Directory.Exists(.DstDir)) Then IO.Directory.CreateDirectory(.DstDir)
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Get the source and destination directories
            strDirIn = .SrcDir
            strDirOut = .DstDir
            ' Check whether correct directories were selected
            If (IO.Directory.Exists(strDirIn)) AndAlso (IO.Directory.Exists(strDirOut)) Then
              ' Consider all *.psdx input files
              arInFile = IO.Directory.GetFiles(strDirIn, "*.psd", IO.SearchOption.AllDirectories)
              ' Do we have something?
              If (arInFile.Count = 0) Then
                ' Try STP
                arInFile = IO.Directory.GetFiles(strDirIn, "*.stp", IO.SearchOption.AllDirectories)
                ' Any results?
                If (arInFile.Count = 0) Then
                  ' Try UPSD
                  arInFile = IO.Directory.GetFiles(strDirIn, "*.upsd", IO.SearchOption.AllDirectories)
                End If
              End If
            End If
          Case Else
            Status("Aborted")
            Exit Sub
        End Select
      End With
      ' Save this directory, if it is different from what we have
      If (strLastConvert <> strDirIn) Then
        strLastConvert = strDirIn
        SetTableSetting(tdlSettings, "LastConvertDir", strLastConvert)
        XmlSaveSettings(strSetFile)
      End If
      If (strLastConvertDst <> strDirOut) Then
        strLastConvertDst = strDirOut
        SetTableSetting(tdlSettings, "LastConvertDirDst", strLastConvertDst)
        XmlSaveSettings(strSetFile)
      End If
      ' Validate
      If (arInFile Is Nothing) Then Status("No files found") : Exit Sub
      ' Determine the language
      With frmConvert
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            strLang = .Language
          Case Else
            Exit Sub
        End Select
      End With
      For intI = 0 To arInFile.Count - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ arInFile.Count
        Status("Processing " & intPtc & "%", intPtc)
        ' Get the input file name
        strInFile = arInFile(intI)
        ' Find name without dots
        strShort = IO.Path.GetFileNameWithoutExtension(strInFile)
        While (InStr(strShort, ".") > 0)
          strShort = IO.Path.GetFileNameWithoutExtension(strShort)
        End While
        ' Determine the output file name, but WITHOUT EXTENSION yet
        strDirApp = Mid(IO.Path.GetDirectoryName(strInFile), strDirIn.Length + 1)
        strOutFile = strDirOut & strDirApp & "\" & strShort
        ' Check existence of directory
        If (Not IO.Directory.Exists(strDirOut & strDirApp)) Then
          IO.Directory.CreateDirectory(strDirOut & strDirApp)
        End If
        ' Add the extension
        strOutFile &= ".psdx"

        '' Determine the output file name
        'strOutFile = strDirOut & "\" & IO.Path.GetFileNameWithoutExtension(strInFile) & _
        '  ".psdx"
        ' Show where we are
        Logging("Producing PSDX " & intI + 1 & "/" & arInFile.Count & " - " & IO.Path.GetFileNameWithoutExtension(strInFile))
        ' Check existence
        If (Not bOverwrite) AndAlso (IO.File.Exists(strOutFile)) AndAlso (Not bAsked) Then
          ' Ask if overwriting is okay
          If (MsgBox("Can I overwrite files like: [" & strOutFile & "]?", MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes) Then
            '' No overwriting, so exit 
            'Exit Sub
            ' We can overwrite, so put to true
            bOverwrite = True
          End If
          ' Indicate we have ased
          bAsked = True
        End If
        ' Combine overwrite and exist
        If (Not IO.File.Exists(strOutFile)) OrElse (bOverwrite) Then
          ' Convert from PSD to PSDX
          If (Not ConvertOnePsdToPsdx(strInFile, strOutFile, strLang, False)) Then
            ' Check for interrupt
            If (bInterrupt) Then Status("Interrupted with error...")
            ' There was an error, so stop
            Exit Sub
          Else
            ' Check for interrupt
            If (bInterrupt) Then Status("Interrupted by user...") : Exit Sub
          End If
        End If
      Next intI
      ' Show we are done
      Status("Ready producing PSDX files.")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsBatch error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   MakePsds
  ' Goal:   Produces PSD files from all the PSDX files in the selected directory
  ' History:
  ' 05-10-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub MakePsds(ByVal strType As String)
    Dim strDirIn As String = ""   ' input directory
    Dim strDirOut As String = ""  ' output directory
    Dim strInFile As String       ' One input file
    Dim strOutFile As String      ' One output file
    Dim arInFile() As String      ' Array of input files
    Dim intI As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage
    Dim bOverwrite As Boolean = False    ' Whether we may overwrite or not

    Try
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      End If
      ' Let user select an input and output directory
      With frmDirs
        ' Initialise the directories myself
        .SrcDir = strWorkDir
        .DstDir = .SrcDir & "\psd"
        ' If the directory does not exist, then make it
        If (Not IO.Directory.Exists(.DstDir)) Then IO.Directory.CreateDirectory(.DstDir)
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Get the source and destination directories
            strDirIn = .SrcDir
            strDirOut = .DstDir
            ' Check whether correct directories were selected
            If (IO.Directory.Exists(strDirIn)) AndAlso (IO.Directory.Exists(strDirOut)) Then
              ' Consider all *.psdx input files
              arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx")
              For intI = 0 To arInFile.Count - 1
                ' Show where we are
                intPtc = (intI + 1) * 100 \ arInFile.Count
                Status("Processing " & intPtc & "%", intPtc)
                ' Get the input file name
                strInFile = arInFile(intI)
                ' Determine the output file name
                strOutFile = strDirOut & "\" & IO.Path.GetFileNameWithoutExtension(strInFile) & _
                  ".psd"
                ' Show where we are
                Logging("Producing PSD " & intI + 1 & "/" & arInFile.Count & " - " & IO.Path.GetFileNameWithoutExtension(strInFile))
                ' Check existence
                If (Not bOverwrite) AndAlso (IO.File.Exists(strOutFile)) Then
                  ' Ask if overwriting is okay
                  If (MsgBox("Can I overwrite files like: [" & strOutFile & "]?", MsgBoxStyle.YesNoCancel) <> MsgBoxResult.Yes) Then
                    ' No overwriting, so exit 
                    Exit Sub
                  End If
                  ' We can overwrite, so put to true
                  bOverwrite = True
                End If
                ' Call the correct function to produce PSD
                If (Not OnePsdFile(strInFile, strOutFile, strType)) Then
                  ' Check for interrupt
                  If (bInterrupt) Then Status("Interrupted...")
                  ' There was an error, so stop
                  Exit Sub
                End If
              Next intI
            End If
        End Select
      End With
      ' Show we are done
      Status("Ready producing PSD files.")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsPsd_Click error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsPsdSimple_Click
  ' Goal:   Produces PSD files (simple) from all the PSDX files in the selected directory
  ' History:
  ' 21-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsPsdSimple_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsPsdSimple.Click
    MakePsds("PsdSimple")
  End Sub
  Private Sub mnuToolsPsdPTB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsPsdPTB.Click
    MakePsds("PsdPTB")
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsPsd_Click
  ' Goal:   Produces PSD files from all the PSDX files in the selected directory
  ' History:
  ' 05-10-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsPsd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsPsd.Click
    MakePsds("PsdOutput")
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsFeatNPtype_Click
  ' Goal:   Correct the NPtype feature for all psdx files in the input directory
  '           (or those under it)
  ' Note:   If you want to use this item, make it VISIBLE in the design page of [frmMain.vb]
  ' History:
  ' 01-09-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsFeatNPtype_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim strDirIn As String = ""   ' input directory
    Dim strInFile As String       ' One input file
    Dim strCategory As String     ' The category we are checking
    Dim arInFile() As String      ' Array of input files
    ' Dim arFeatDef(0) As String    ' Feature definitions - dummy
    Dim pdxFile As XmlDocument    ' One input file as XML
    Dim intI As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage

    Try
      ' Ask for an input directory
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, _
                    "Root directory containing psdx files for which the NPtype will be recalculated", strWorkDir)) Then Exit Sub
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      End If
      ' Make sure interrupt is not on
      bInterrupt = False
      ' Initialise missing pronoun report
      InitMissing()
      ' Initialise pronoun tracking
      InitCatTrack()
      ' Initially clear the error report
      Me.wbError.DocumentText = "(empty)"
      ' Set the category we are renewing
      strCategory = CAT_NPTYPE
      ' Get all the input files
      arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories)
      ' Walk through them one by one
      For intI = 0 To arInFile.Length - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ arInFile.Count
        Status("Processing " & intPtc & "%", intPtc)
        ' Check for interrupt
        If (bInterrupt) Then
          ' We have been interrupted
          Status("Processing has been interrupted")
          bInterrupt = False
          Exit Sub
        End If
        ' What file are we working with?
        strInFile = arInFile(intI) : pdxFile = Nothing
        CorefProgress(intI + 1 & "/" & arInFile.Length & ": " & IO.Path.GetFileNameWithoutExtension(strInFile))
        ' Try read this file into an XML structure
        If (ReadXmlDoc(strInFile, pdxFile)) Then
          ' Calculate the features for this category
          If (Not OneDoFeaturesPdx(pdxFile, strCategory, True)) Then
            ' Some error has occurred
            Status("There was a problem reading file: " & strInFile & vbCrLf & "Processing has stopped")
            Exit Sub
          End If
          ' Write the result to the destination
          pdxFile.Save(strInFile)
        Else
          ' Could not read the file, so return failure
          Status("Could not read the XML file: " & strInFile & vbCrLf & "Processing has stopped")
          Exit Sub
        End If
      Next intI
      ' Reset the progress indicator
      CorefProgress("-")
      ' Show that we are ready
      Status("Ready")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsFeatNPtype_Click error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuMorphCategories_Click
  ' Goal:   Gather all the POS categories, asking whether they are "Main" or "Derived"
  '         And if they are derived, what their "Head" is
  ' History:
  ' 13-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuMorphCategories_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMorphCategories.Click
    Dim strCatRepFile As String = ""    ' YCOE tagging report file
    Dim strDirIn As String = ""         ' Input directory with YCOE files

    Try
      ' We need access to [tdlMorphDict]
      MorphDictIni()
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      ' Let user select an input directory
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, "Give directory where morphologically enriched OE files are located", _
        "d:\data files\corpora\English\xml\Adapted\OE")) Then Exit Sub
      If (MorphCategories(strDirIn, strCatRepFile)) Then
        ' Put the report in place, by navigating to the file
        Status("Opening the categories report...")
        Me.wbReport.Navigate(strCatRepFile)
        ' Show we are ready
        Status("Ready collecting categories information (see report)")
        Me.TabControl1.SelectedTab = Me.tpReport
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsOEmorphCategories error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMorphTaggedReport_Click
  ' Goal:   Give a report of what is already tagged in the texts
  ' History:
  ' 11-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsMorphTaggedReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMorphTaggedReport.Click
    Dim strTagRepFile As String = ""    ' YCOE tagging report file
    Dim strDirIn As String = ""         ' Input directory with YCOE files

    Try
      ' Let user select an input directory
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, "Give directory where morphologically enriched OE files are located", _
        "d:\data files\corpora\English\xml\Adapted\OE")) Then Exit Sub
      If (MakeTaggingReport(strDirIn, strTagRepFile)) Then
        ' Put the report in place, by navigating to the file
        Status("Opening the YCOE tagging report...")
        Me.wbReport.Navigate(strTagRepFile)
        ' Show we are ready
        Status("Ready collecting YCOE tagging information (see report)")
        Me.TabControl1.SelectedTab = Me.tpReport
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsOEmorphTaggedReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMorphRemainReport_Click
  ' Goal:   Give a report of what is already tagged in the texts
  ' History:
  ' 11-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsMorphRemainReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMorphRemainReport.Click
    Dim strTagRepFile As String = ""    ' YCOE tagging report file
    Dim strDirIn As String = ""         ' Input directory with YCOE files

    Try
      ' Let user select an input directory
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, "Give directory where morphologically enriched OE files are located", _
        "d:\data files\corpora\English\xml\Adapted\OE")) Then Exit Sub
      ' We need access to [tdlMorphDict], to create a list of what remains
      ' Initialise morphological dictionary
      MorphDictIni()
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      ' Clear the table for remainders
      ClearTable(tdlMorphDict.Tables("Remain"))
      '' Make sure changes are processed
      'Status("processing removal of old rows...")
      'tdlSettings.AcceptChanges()
      'Status("continuing...")
      ' Start making remain report (html)
      If (MakeRemainReport(strDirIn, "OE", strTagRepFile)) Then
        ' Put the report in place, by navigating to the file
        Status("Opening the YCOE remain report...")
        Me.wbReport.Navigate(strTagRepFile)
        ' Show we are ready
        Status("Ready collecting YCOE remain information (see report)")
        Me.TabControl1.SelectedTab = Me.tpReport
      End If

    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsOEmorphRemainReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMorphSuffix_Click
  ' Goal:   Derive suffix-rewrite rules in table [Morph]
  '         Look for candidates in table [Remain], against lemma's in [Morph] and [Lemma]
  '         Add candidate Vern/Pos/Lemma combinations in [VernPos]
  ' History:
  ' 16-03-2013  ERK Created first part
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsMorphSuffix_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMorphSuffix.Click
    Try
      ' Initialise morphological dictionary
      MorphDictIni()
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      ' Try sweeping
      If (Not MorphSuffixRules()) Then Status("Interrupt pressed or Suffix-Rewrite met an error")
      ' Ask whether we should save
      Select Case MsgBox("Would you like to adapt the [VernPos] table using the suffix rules?", MsgBoxStyle.YesNo)
        Case MsgBoxResult.Yes
          ' Adapt [VernPos] with the suffix rules
          ' TODO: implement...

          ' Save the result
          tdlMorphDict.WriteXml(strMorphDictFile)
      End Select
      ' Show we are ready
      Status("Blessings!")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsOEmorphSuffixRewrite error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMorphDerive_Click
  ' Goal:   Derive suffix-rewrite rules from table [VernPos]
  ' History:
  ' 13-05-2013  ERK Created first part
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsMorphDerive_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMorphDerive.Click
    Try
      ' Initialise morphological dictionary
      MorphDictIni()
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      ' Try sweeping
      If (Not MorphDeriveRules()) Then Status("Interrupt pressed or Derive-Rewrite met an error")
      ' Ask whether we should save
      Select Case MsgBox("Would you like to adapt the [VernPos] table using the derive rules?", MsgBoxStyle.YesNo)
        Case MsgBoxResult.Yes
          ' Adapt [VernPos] with the suffix rules
          ' TODO: implement...

          ' Save the result
          tdlMorphDict.WriteXml(strMorphDictFile)
      End Select
      ' Show we are ready
      Status("Blessings!")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsOEmorphDeriveRewrite error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMorphLemmaAmbi_Click
  ' Goal:   Check all items in table [VernPos] of type "LemmaAmbi"
  '         Ask user if one lemma should be taken, and if so, which one
  ' History:
  ' 26-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsMorphLemmaAmbi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMorphLemmaAmbi.Click
    Try
      ' Initialise morphological dictionary
      MorphDictIni()
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      ' Try lemmatization using rewrite rules
      If (Not MorphLemmaAmbi()) Then Status("Interrupt pressed or LemmaAmbi met an error")
      ' Ask whether we should save
      Select Case MsgBox("Would you like to save the adapted [VernPos] table?", MsgBoxStyle.YesNo)
        Case MsgBoxResult.Yes
          tdlMorphDict.WriteXml(strMorphDictFile)
      End Select
      ' Show we are ready
      Status("Blessings!")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsOEmorphLemmatizeRewrite error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMorphLemmatizeRewrite_Click
  ' Goal:   Using the suffix-rewrite rules in table [Suffix],
  '           attempt to determine the lemma for words that have not yet been lemmatized
  '         Action depends on the number of different lemma's that are possible
  '         zero - leave as is
  '         one  - accept this as lemma
  '         more - add a list of lemma's, possibly extended with confidence?
  ' History:
  ' 22-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsMorphLemmatizeRewrite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMorphLemmatizeRewrite.Click
    Try
      ' Initialise morphological dictionary
      MorphDictIni()
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      ' Try lemmatization using rewrite rules
      If (Not MorphLemmatizeRewrite()) Then Status("Interrupt pressed or LemmatizeRewrite met an error")
      ' Ask whether we should save
      Select Case MsgBox("Would you like to save the adapted [VernPos] table?", MsgBoxStyle.YesNo)
        Case MsgBoxResult.Yes
          tdlMorphDict.WriteXml(strMorphDictFile)
      End Select
      ' Show we are ready
      Status("Blessings!")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsOEmorphLemmatizeRewrite error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMorphLemmatizeByDict_Click
  ' Goal:   Using the XML dictionary (Bosworth & Toller derived)
  '           attempt to determine the lemma for main entries, such as ADJ, VB, N, N^N, NR, NR^N and so on
  ' History:
  ' 08-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsMorphLemmatizeByDict_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMorphLemmatizeByDict.Click
    Try
      ' Initialise morphological dictionary
      MorphDictIni()
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      ' Try read the OE dictionary
      If (Not MorphReadOEdict()) Then Status("Could not read the OE dictionary database") : Exit Sub
      ' Try to extract prefixes from OEdict into MorphDict table [Prefix]
      If (Not MorphPfxOEdict()) Then Status("Could not extract prefixes from OE dictionary database") : Exit Sub
      ' Try lemmatization using rewrite rules
      If (Not MorphLemmatizeByDict()) Then Status("Interrupt pressed or LemmatizeRewrite met an error")
      ' Ask whether we should save
      Select Case MsgBox("Would you like to save the adapted [VernPos] table?", MsgBoxStyle.YesNo)
        Case MsgBoxResult.Yes
          tdlMorphDict.WriteXml(strMorphDictFile)
      End Select
      ' Show we are ready
      Status("Blessings!")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsOEmorphLemmatizeByDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMorphLemmatizeByTable_Click
  ' Goal:   Process the resolved entries from table [Remain] into table [VernPos]
  '         The resolved entries are in a tab-separated text file
  ' History:
  ' 27-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsMorphLemmatizeByTable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMorphLemmatizeByTable.Click
    Dim strFile As String = ""  ' File with the resolve information we are looking for

    Try
      ' Initialise morphological dictionary
      MorphDictIni()
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      ' Ask for the resolution tab-separated file
      If (Not GetFileName(Me.OpenFileDialog1, GetDocDir(), strFile, "tab-separated text file (*.csv)|*.csv")) Then Exit Sub
      ' Process this file
      If (Not MorphResolveTable(strFile)) Then Status("Could not process resolution information") : Exit Sub
      ' Ask whether we should save
      Select Case MsgBox("Would you like to save the adapted [VernPos] table?", MsgBoxStyle.YesNo)
        Case MsgBoxResult.Yes
          tdlMorphDict.WriteXml(strMorphDictFile)
      End Select
      ' Show we are ready
      Status("Blessings!")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsOEmorphLemmatizeByTable error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMorphSweep_Click
  ' Goal:   Address the entries from table [Remain] using information from [Morph]
  '           and then adding entries into table [VernPos]
  ' History:
  ' 14-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsMorphSweep_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMorphSweep.Click
    Try
      ' Initialise morphological dictionary
      MorphDictIni()
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      ' Try sweeping
      If (Not MorphSweep()) Then Status("Interrupt pressed or Sweeping met an error")
      ' Ask whether we should save
      Select Case MsgBox("Would you like to save the adapted [VernPos] table?", MsgBoxStyle.YesNo)
        Case MsgBoxResult.Yes
          tdlMorphDict.WriteXml(strMorphDictFile)
      End Select
      ' Show we are ready
      Status("Blessings!")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsOEmorphSweep error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMorhVerbs_Click
  ' Goal:   Try to lemmatize the verbs
  ' History:
  ' 14-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsMorhVerbs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMorhVerbs.Click
    Dim strDirIn As String = ""         ' Input directory with YCOE files
    Dim strInFile As String = ""        ' One input file
    Dim arInFile() As String            ' Array of input files
    Dim intI As Integer                 ' Counter
    Dim intPtc As Integer               ' Percentage

    Try
      ' Initialise morphological dictionary
      MorphDictIni()
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      ' Let user select an input directory
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, "Give directory where morphologically enriched OE files are located", _
        "d:\data files\corpora\English\xml\Adapted\OE")) Then Exit Sub
      ' Consider all *.psdx input files
      arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories)
      For intI = 0 To arInFile.Count - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ arInFile.Count
        Status("Processing " & intPtc & "%", intPtc)
        ' Get the input file name
        strInFile = arInFile(intI)
        ' Also show it on the main form
        Logging("OE-morphdict propagation [" & IO.Path.GetFileNameWithoutExtension(strInFile) & "] (" & intI + 1 & "/" & arInFile.Count & ")")
        ' Process this one file
        If (Not OneMorphPropaVerb(strInFile)) Then
          ' Inform user
          Status("Error in processing: " & strInFile)
          Exit Sub
        End If
      Next intI

      ' Show we are ready
      Status("Blessings!")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsOEmorphVerbs error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMorphAnalyzer_Click
  ' Goal:   Try to read the "analyzer.sql" file from Ondrej Tichy and transform it into
  '           something I can use
  ' History:
  ' 08-02-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsMorphAnalyzer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMorphAnalyzer.Click
    If (ReadMorphOEanalyzer()) Then
      Status("Ready")
    Else
      Status("There was an error")
    End If
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMorphOEtexts_Click
  ' Goal:   Go through the data text-by-text, asking the user to supply the lemma
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsMorphOEtexts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMorphOEtexts.Click
    Dim strDirIn As String = ""         ' Input directory with YCOE files
    Dim strDirOut As String = ""        ' Output directory 
    Dim strInFile As String = ""        ' One input file
    Dim strMorphRepFile As String = ""  ' Morphology report file
    Dim strTagRepFile As String = ""    ' YCOE tagging report file
    Dim arInFile() As String            ' Array of input files
    Dim colRep As New StringColl        ' Checking report
    Dim intI As Integer                 ' Counter
    Dim intPtc As Integer               ' Percentage
    Dim intDone As Integer = 0          ' Number that has been done
    Dim intNeed As Integer = 0          ' Number that still need doing
    Dim intD As Integer = 0
    Dim intN As Integer = 0
    Dim bNeeded As Boolean = False
    Dim bChanged As Boolean = False     ' Something changed in dataset?

    Try
      ' Initialise morphological dictionary
      Status("Morphological dictionary...")
      bInterrupt = False
      MorphDictIni()
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      ' Read the OE dictionary
      Status("OE dictionary...")
      If (Not MorphReadOEdict()) Then Status("Could not open OE dictionary") : Exit Sub
      ' Let user select an input directory
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, "Give directory where morphologically enriched OE files are located", _
        "d:\data files\corpora\English\xml\Adapted\OE")) Then Exit Sub
      ' Consider all *.psdx input files
      arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories)
      For intI = 0 To arInFile.Count - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ arInFile.Count
        Status("Processing " & intPtc & "%", intPtc)
        ' Get the input file name
        strInFile = arInFile(intI)
        ' Also show it on the main form
        Logging("OE-lemmatization [" & IO.Path.GetFileNameWithoutExtension(strInFile) & "] (" & intI + 1 & "/" & arInFile.Count & ")")
        ' Process this one file
        If (Not OneMorphPropaText(strInFile, bChanged)) Then
          ' Inform user
          Status("Interrupt or Error in processing: " & strInFile)
          Exit Sub
        End If
        ' Anything changed?
        If (bChanged) Then
          Status("Saving dataset...")
          tdlMorphDict.WriteXml(strMorphDictFile)
          bChanged = False
        End If
      Next intI
      ' Show we are ready
      Status("Ready")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsOEmorphTexts error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMorphOE_Click
  ' Goal:   Visit all YCOE files and store morphological information:
  '         1) Look for M-features
  '         2) Note the POS and the parent's <eTree> label
  '         3) Note the kinds of M-features
  ' History:
  ' 18-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsMorphProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMorphProcess.Click
    Dim strDirIn As String = ""         ' Input directory with YCOE files
    Dim strDirOut As String = ""        ' Output directory 
    Dim strInFile As String = ""        ' One input file
    Dim strOutFile As String = ""       ' One input file
    Dim strMorphRepFile As String = ""  ' Morphology report file
    Dim strTagRepFile As String = ""    ' YCOE tagging report file
    Dim arInFile() As String            ' Array of input files
    Dim colRep As New StringColl        ' Checking report
    Dim intI As Integer                 ' Counter
    Dim intPtc As Integer               ' Percentage
    Dim intDone As Integer = 0          ' Number that has been done
    Dim intNeed As Integer = 0          ' Number that still need doing
    Dim intD As Integer = 0
    Dim intN As Integer = 0
    Dim bNeeded As Boolean = False

    Try
      ' Initialise morphological dictionary
      MorphDictIni()
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      ' Let user select an input directory
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, "Give directory where morphologically enriched OE files are located", _
        "d:\data files\corpora\English\xml\Adapted\OE")) Then Exit Sub
      ' Make sure input and output directory are the same
      strDirOut = strDirIn
      ' Check where we are
      If (tdlMorphDict.Tables("Morph").Rows.Count > 0) Then
        ' Ask user
        Select Case MsgBox("Morphology information is already available." & vbCrLf & _
                           "Would you like to start this process afresh?", MsgBoxStyle.YesNoCancel)
          Case MsgBoxResult.Cancel
            ' Leave!
            Exit Sub
          Case MsgBoxResult.Yes
            bNeeded = True
          Case MsgBoxResult.No
            bNeeded = False
        End Select
      Else
        ' Needs doing
        bNeeded = True
      End If
      ' Do we need doing?
      If (bNeeded) Then
        Status("Deleting old information...")
        MorphDictIni(True)
        '' Initialise the [Morph] table in the dictionary
        'ClearTable(tdlMorphDict.Tables("Morph"))
        ' Consider all *.psdx input files
        arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories)
        For intI = 0 To arInFile.Count - 1
          ' Show where we are
          intPtc = (intI + 1) * 100 \ arInFile.Count
          Status("Processing " & intPtc & "%", intPtc)
          ' Get the input file name
          strInFile = arInFile(intI)
          ' Make the name of the output file
          strOutFile = strDirOut & "\" & IO.Path.GetFileName(strInFile)
          ' Also show it on the main form
          Logging("OE-morphology [" & IO.Path.GetFileNameWithoutExtension(strInFile) & "] (" & intI + 1 & "/" & arInFile.Count & ")")
          ' Process this one file
          If (Not OneMorphGatherFile(strInFile, intD, intN)) Then
            ' Inform user
            Status("Error in processing: " & strInFile)
            Exit Sub
          End If
          ' Add totals
          intDone += intD : intNeed += intN
        Next intI
        ' Write the morphological dictionary
        Status("Writing morphological dictionary...")
        tdlMorphDict.WriteXml(strMorphDictFile)
      End If
      ' Does the user want to check and correct morphology, comparing OE-tagged lemma list with YCOE file?
      Select Case MsgBox("Would you like to check the added M-lemmas in YCOE against the OE-tagged lemma dictionary?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Cancel
          Exit Sub
        Case MsgBoxResult.Yes
          ' Consider all *.psdx input files
          arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories)
          For intI = 0 To arInFile.Count - 1
            ' Show where we are
            intPtc = (intI + 1) * 100 \ arInFile.Count
            Status("Processing " & intPtc & "%", intPtc)
            ' Get the input file name
            strInFile = arInFile(intI)
            ' Make the name of the output file
            strOutFile = strDirOut & "\" & IO.Path.GetFileName(strInFile)
            ' Also show it on the main form
            Logging("OE-morphology [" & IO.Path.GetFileNameWithoutExtension(strInFile) & "] (" & intI + 1 & "/" & arInFile.Count & ")")
            ' First attempt morphology correction
            If (Not OneMorphCorrectFile(strInFile, intD)) Then
              ' Inform user
              Status("Error in processing: " & strInFile)
              Exit Sub
            End If
            ' Process changes
            If (intD > 0) Then
              ' Write the morphological dictionary
              Status("Writing morphological dictionary...")
              tdlMorphDict.WriteXml(strMorphDictFile)
            End If
          Next intI
          Select Case MsgBox("Would you like to re-make the morphological dictionary now?", MsgBoxStyle.YesNoCancel)
            Case MsgBoxResult.Cancel
              Exit Sub
            Case MsgBoxResult.Yes
              Status("Deleting old information...")
              MorphDictIni(True)
              '' Initialise the [Morph] table in the dictionary
              'ClearTable(tdlMorphDict.Tables("Morph"))
              ' Consider all *.psdx input files
              arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories)
              For intI = 0 To arInFile.Count - 1
                ' Show where we are
                intPtc = (intI + 1) * 100 \ arInFile.Count
                Status("Processing " & intPtc & "%", intPtc)
                ' Get the input file name
                strInFile = arInFile(intI)
                ' Make the name of the output file
                strOutFile = strDirOut & "\" & IO.Path.GetFileName(strInFile)
                ' Also show it on the main form
                Logging("OE-morphology [" & IO.Path.GetFileNameWithoutExtension(strInFile) & "] (" & intI + 1 & "/" & arInFile.Count & ")")
                ' Process this one file
                If (Not OneMorphGatherFile(strInFile, intD, intN)) Then
                  ' Inform user
                  Status("Error in processing: " & strInFile)
                  Exit Sub
                End If
                ' Add totals
                intDone += intD : intNeed += intN
              Next intI
              ' Write the morphological dictionary
              Status("Writing morphological dictionary...")
              tdlMorphDict.WriteXml(strMorphDictFile)
          End Select
      End Select
      ' Possibly give YCOE tagging report
      Select Case MsgBox("Would you like to have a report on what is tagged in YCOE?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Cancel
          Exit Sub
        Case MsgBoxResult.Yes
          If (MakeTaggingReport(strDirIn, strTagRepFile)) Then
            ' Put the report in place, by navigating to the file
            Status("Opening the YCOE tagging report...")
            Me.wbReport.Navigate(strTagRepFile)
            ' Show we are ready
            Status("Ready collecting YCOE tagging information (see report)")
            Me.TabControl1.SelectedTab = Me.tpReport
          End If
      End Select
      ' Possibly give morphology report
      Select Case MsgBox("Would you like to have a morphology report?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Cancel
          Exit Sub
        Case MsgBoxResult.Yes
          ' Prepare tag-report anyway
          Status("Preparing morphology report...")
          If (MorphologyReport(strMorphRepFile)) Then
            ' Put the report in place, by navigating to the file
            Status("Opening the morphology report...")
            Me.wbReport.Navigate(strMorphRepFile)
            ' Show we are ready
            Status("Ready collecting morphology information (see report)")
            Me.TabControl1.SelectedTab = Me.tpReport
          End If
      End Select
      ' Do [tdlMorphDict] adaptation
      Select Case MsgBox("Would you like to adapt the [morphdict]?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Cancel
          Exit Sub
        Case MsgBoxResult.Yes
          ' Make sure that @h features are in the @h and not in the @f field
          If (Not MorphDictAdapt()) Then Status("Problem in MorphDictAdapt") : Exit Sub
          ' Write the morphological dictionary
          Status("Writing morphological dictionary...")
          tdlMorphDict.WriteXml(strMorphDictFile)
      End Select
      ' Now go through the files to update them with the newly acquired information (if user wishes so)
      Select Case MsgBox("Would you like to propagate the [morphdict] information to the whole YCOE?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Cancel
          Status("Blessings!")
          Exit Sub
        Case MsgBoxResult.Yes
          ' Consider all *.psdx input files
          arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories)
          For intI = 0 To arInFile.Count - 1
            ' Show where we are
            intPtc = (intI + 1) * 100 \ arInFile.Count
            Status("Processing " & intPtc & "%", intPtc)
            ' Get the input file name
            strInFile = arInFile(intI)
            ' =============== Debugging ===============
            'If (InStr(strInFile, "james") > 0) Then Stop
            ' =========================================
            ' Make the name of the output file (NOTE: NOT CURRENTLY USED)
            strOutFile = strDirOut & "\" & IO.Path.GetFileName(strInFile)
            ' Also show it on the main form
            Logging("OE-morphdict propagation [" & IO.Path.GetFileNameWithoutExtension(strInFile) & "] (" & intI + 1 & "/" & arInFile.Count & ")")
            ' Process this one file
            If (Not OneMorphPropaDict(strInFile, strOutFile, "OE")) Then
              ' Inform user
              Status("Error in processing: " & strInFile)
              Exit Sub
            End If
          Next intI
      End Select
      ' Now go through the files to update them with the newly acquired information (if user wishes so)
      Select Case MsgBox("Would you like to propagate the other morphology information to the whole YCOE?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.No, MsgBoxResult.Cancel
          Status("Blessings!")
          Exit Sub
      End Select
      ' Consider all *.psdx input files
      arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories)
      For intI = 0 To arInFile.Count - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ arInFile.Count
        Status("Processing " & intPtc & "%", intPtc)
        ' Get the input file name
        strInFile = arInFile(intI)
        ' =============== Debugging ===============
        'If (InStr(strInFile, "james") > 0) Then Stop
        ' =========================================
        ' Make the name of the output file (NOTE: NOT CURRENTLY USED)
        strOutFile = strDirOut & "\" & IO.Path.GetFileName(strInFile)
        ' Also show it on the main form
        Logging("OE-morphology adaptation [" & IO.Path.GetFileNameWithoutExtension(strInFile) & "] (" & intI + 1 & "/" & arInFile.Count & ")")
        ' Process this one file
        If (Not OneMorphPropaFile(strInFile, strOutFile, False)) Then
          ' Inform user
          Status("Error in processing: " & strInFile)
          Exit Sub
        End If
      Next intI
      ' Write the morphological dictionary
      Status("Writing morphological dictionary...")
      tdlMorphDict.WriteXml(strMorphDictFile)
      ' Possibly give YCOE tagging report
      Select Case MsgBox("Would you like to have a report on what is now tagged in YCOE?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Cancel
          Exit Sub
        Case MsgBoxResult.Yes
          ' Reset the tagging report
          TagInit()
          ' Make it afresh
          If (MakeTaggingReport(strDirIn, strTagRepFile)) Then
            ' Put the report in place, by navigating to the file
            Status("Opening the YCOE tagging report...")
            Me.wbReport.Navigate(strTagRepFile)
            ' Show we are ready
            Status("Ready collecting YCOE tagging information (see report)")
            Me.TabControl1.SelectedTab = Me.tpReport
          End If
      End Select
      ' Signal we are ready
      Status("Morphology has been adapted")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/ToolsMorphProcess error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMorphEdit_Click
  ' Goal:   Open the form that deals with the different tables in the tdlMorphDict
  '         (Not just lemma's...)
  ' History:
  ' 11-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsMorphEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMorphEdit.Click
    Try
      ' Is the correct datatable loaded?
      If (tdlMorphDict Is Nothing) Then
        ' Try load it
        MorphDictIni()
        ' Double check
        If (tdlMorphDict Is Nothing) Then Status("Could not initialize morphdict") : Exit Sub
      End If
      ' Go to the lemma window
      With frmLemmas
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            Status("Ready")
          Case Windows.Forms.DialogResult.Cancel
            Status("Aborted")
        End Select
      End With
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/ToolsMorphEdit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMorphHeadRules_Click
  ' Goal:   Get a list of the possible children each XP can have
  ' History:
  ' 27-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsMorphHeadRules_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMorphHeadRules.Click
    Dim strResult As String = ""  ' The result

    Try
      ' Call my function
      If (TravelXPchildren(strResult)) Then
        Me.wbReport.DocumentText = strResult
        Status("Done")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/ToolsMorphHeadRules error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsFeatTagged_Click
  ' Goal:   Check all YCOE files to see where they co-incide with OE-tagged (=Toronto) ones
  '         Add the features from the OE-tagged corpus where possible
  ' History:
  ' 29-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsFeatTagged_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsFeatTagged.Click
    Dim strDirIn As String = ""   ' Input directory with YCOE files
    Dim strDirOut As String = ""  ' Output directory 
    Dim strInFile As String = ""  ' One input file
    Dim strOutFile As String = "" ' One input file
    Dim arInFile() As String      ' Array of input files
    Dim colRep As New StringColl  ' Checking report
    Dim intI As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage
    Dim intDone As Integer = 0    ' Number that has been done
    Dim intNeed As Integer = 0    ' Number that still need doing
    Dim intD As Integer = 0
    Dim intN As Integer = 0

    Try
      ' Let user select an input and output directory
      With frmDirs
        ' Initialise the directories myself
        .SrcDir = "d:\data files\corpora\English\xml\Adapted\OE"
        .DstDir = "d:\data files\corpora\English\xml\AdaptedNew\OE"
        ' Initialise missing pronoun report
        InitMissing()
        ' Initialise pronoun tracking
        InitCatTrack()
        ' If the directory does not exist, then make it
        If (Not IO.Directory.Exists(.DstDir)) Then IO.Directory.CreateDirectory(.DstDir)
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Get the source and destination directories
            strDirIn = .SrcDir
            strDirOut = .DstDir
        End Select
      End With
      ' Start tag-comment report
      TagInit()
      colRep.Add("<h4>Totals</h4><p><table><tr><td>File</td><td>Done</td><td>Need</td></tr>")
      ' Reset interrupt
      bInterrupt = False
      ' Consider all *.psdx input files
      arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories)
      For intI = 0 To arInFile.Count - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ arInFile.Count
        Status("Processing " & intPtc & "%", intPtc)
        ' Get the input file name
        strInFile = arInFile(intI)
        ' Make the name of the output file
        strOutFile = strDirOut & "\" & IO.Path.GetFileName(strInFile)
        ' Log this file
        TagComment("Processing [" & IO.Path.GetFileNameWithoutExtension(strInFile) & "]")
        If (InStr(strInFile, "james") > 0) Then Stop
        ' Also show it on the main form
        Logging("OEtagging [" & IO.Path.GetFileNameWithoutExtension(strInFile) & "] (" & intI + 1 & "/" & arInFile.Count & ")")
        ' Process this one file
        If (Not OneTaggedFeaturesFile(strInFile, strOutFile, intD, intN)) Then
          ' Inform user
          Status("Error in processing: " & strInFile)
          ' Open tag-report
          Me.wbReport.DocumentText = "<html><body>" & TagReport() & "</body></html>"
          Me.TabControl1.SelectedTab = Me.tpReport
          Exit Sub
        End If
        ' Report
        colRep.Add("<tr><td>" & IO.Path.GetFileNameWithoutExtension(strInFile) & "</td><td>" & intD & "</td><td>" & intN & "</td></tr>")
        ' Add totals
        intDone += intD : intNeed += intN
      Next intI
      ' Report
      colRep.Add("<tr><td>TOTAL</td><td>" & intDone & "</td><td>" & intNeed & "</td></tr></table>")
      ' Add totals to tag report
      TagComment("Totals: done = " & intDone & " need = " & intNeed)
      ' Prepare tag-report anyway
      Me.wbReport.DocumentText = "<html><body>" & TagReport() & colRep.Text & "</body></html>"
      ' Show we are ready
      Status("Ready adding OE-tagged features (see report)")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/ToolsFeatTagged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsCalcFeatures_Click
  ' Goal:   Calculate particular features for all the XPs in the PSDX files in the selected directory
  ' History:
  ' 18-05-2010  ERK Created
  ' 28-01-2011  ERK Adapted to be able to get all kinds of features for all files in a directory
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsCalcFeatures_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim strDirIn As String = ""   ' input directory
    Dim strDirOut As String = ""  ' output directory
    Dim strInFile As String       ' One input file
    Dim strOutFile As String      ' One output file
    Dim strCategory As String     ' The category of features the user would like to renew
    Dim strReport As String       ' Text of report (HTML)
    Dim strFeatDef As String = "" ' Feature definition file
    Dim arInFile() As String      ' Array of input files
    Dim bCheck As Boolean         ' Check differences (recalculate) or not
    Dim intI As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage

    Try
      ' Reset interrupt
      bInterrupt = False
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      End If
      ' Get the correct category to be worked with
      With frmCategory
        ' Make sure options are shown
        .ShowOptions = True
        ' Ask the user's opinion
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Get the result
            strCategory = .Category
            bCheck = .Check
            ' Do category specific matters
            Select Case strCategory
              Case CAT_ADV
                ' Initialise the feature dictionary
                If (Not MakeFeatDict(strCategory)) Then
                  Status("Unable to initialize the feature dictionary")
                  Exit Sub
                End If
              Case CAT_NP
              Case CAT_PNORM, CAT_VB, CAT_VBTYPE
                ' Get the file with the relation between preposition form and normalized (standardized) preposition = Lemma??
                strFeatDef = .FeatFile
                ' make a feature dictionary
                If (Not MakeFeatDict(strCategory, strFeatDef)) Then
                  Status("Unable to initialize feature dictionary")
                  Exit Sub
                End If
              Case CAT_VBUNACC  ' Unaccusatives that have been coded in a parallel *psdx corpus
                ' The file is the first one of similar ones in a direactory of parallel psdx files
                str_UnaccDir = IO.Path.GetDirectoryName(.FeatFile)
              Case CAT_COGN
            End Select
          Case Else
            ' No use to continue
            Exit Sub
        End Select
      End With
      ' Let user select an input and output directory
      With frmDirs
        ' Initialise the directories myself
        .SrcDir = strWorkDir
        .DstDir = .SrcDir & "\Adapted"
        ' Initialise missing pronoun report
        InitMissing()
        ' Initialise pronoun tracking
        InitCatTrack()
        ' If the directory does not exist, then make it
        If (Not IO.Directory.Exists(.DstDir)) Then IO.Directory.CreateDirectory(.DstDir)
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Get the source and destination directories
            strDirIn = .SrcDir
            strDirOut = .DstDir
            ' Check whether correct directories were selected
            If (IO.Directory.Exists(strDirIn)) AndAlso (IO.Directory.Exists(strDirOut)) Then
              ' Consider all *.psdx input files
              arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories)
              For intI = 0 To arInFile.Count - 1
                ' Show where we are
                intPtc = (intI + 1) * 100 \ arInFile.Count
                Status("Processing " & intPtc & "%", intPtc)
                ' Get the input file name
                strInFile = arInFile(intI)
                ' Determine the output file name
                strOutFile = strDirOut & "\" & IO.Path.GetFileName(strInFile)
                ' Show where we are
                Logging("Calculate " & strCategory & " features " & intI + 1 & "/" & _
                        arInFile.Count & " - " & IO.Path.GetFileNameWithoutExtension(strInFile))
                ' Call the correct function
                If (Not OneCalcFeaturesFile(strInFile, strOutFile, strCategory, bCheck)) Then
                  ' There was an error, so stop
                  Exit Sub
                End If
              Next intI
            End If
        End Select
      End With
      ' Try get the pronoun tracking report
      If (HasCatTrack()) Then
        ' Calculate this report and put it in place
        strReport = CatTrackHtml(strCategory)
        strOutFile = strDirIn & "\" & strCategory & "_Features.htm"
        ' Add report location
        strReport = strOutFile & "<p>" & strReport
        Me.wbReport.DocumentText = strReport
        ' Switch to this tabpage
        Me.TabControl1.SelectedTab = Me.tpReport
        ' Save the report on an appropriate location
        IO.File.WriteAllText(strOutFile, strReport)
      End If
      ' Try get the missing pronouns report
      If (HasMissing()) Then
        ' Calculate the report and put it in place
        Me.wbError.DocumentText = MissingHtml(strCategory)
        ' Show there are some errors
        Status("Ready adapting " & strCategory & " features. There are missing pronouns reported in the <Errors> page.")
      Else
        ' Show we are done
        Status("Ready adapting " & strCategory & " features.")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsCalcFeatures_Click error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsShowFeatures_Click
  ' Goal:   Give a report of the features for NP, Adv or Verb
  '         1 - NP
  '         2 - Adv
  '         3 - V
  ' History:
  ' 14-02-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsShowFeatures_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim strDirIn As String = ""   ' input directory
    Dim strDirOut As String = ""  ' output directory
    Dim strCategory As String     ' The category of features the user would like to renew
    Dim strReport As String       ' Text of report (HTML)
    Dim strOutFile As String      ' One output file
    Dim strName As String = ""    ' Name of category
    Dim strType As String = ""    ' Type of this category (part of the name)
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intK As Integer           ' Counter
    Dim dtrFound() As DataRow     ' Result of selection statement
    Dim objOut As New StringColl  ' Output string collection
    Dim arWord() As String        ' List of words

    Try
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      End If
      ' Get the correct category to be worked with
      With frmCategory
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Get the result
            strCategory = .Category
          Case Else
            ' No use to continue
            Exit Sub
        End Select
      End With
      ' If this is "Adv", then we know what to do
      Select Case strCategory
        Case CAT_NP       ' NP ???
        Case CAT_VB       ' Verb types
        Case CAT_VBUNACC
        Case CAT_PNORM    ' Normalized prepositions
        Case CAT_ADV
          dtrFound = tdlSettings.Tables("Category").Select("Name LIKE 'Adv-*'", "Type ASC")
          ' Initialise
          strType = "" : objOut.Add("<html><body><h2>Report on features for: " & strCategory & "</h2><p>")
          ' Go through all the results
          For intI = 0 To dtrFound.Length - 1
            ' Access this element
            With dtrFound(intI)
              ' Check whether this is a new type
              If (strType <> .Item("Type")) Then
                ' Copy the new type
                strType = .Item("Type")
                ' Make a new section
                objOut.Add("<h3>Type: " & strType & "</h3>")
              End If
              ' Start period output table
              objOut.Add("<table><tr><td>Period</td><td>" & strCategory & "/" & strType & "</td></tr>")
              ' Output all periods
              For intJ = 0 To UBound(arPeriod)
                ' Is there anything for this period?
                If (.Item(arPeriod(intJ)).ToString <> "") Then
                  ' Output this period
                  objOut.Add("<tr><td valign=top>" & arPeriod(intJ) & "</td><td valign=top>")
                  arWord = Split(.Item(arPeriod(intJ)), "|")
                  ' Sort the results
                  Array.Sort(arWord)
                  ' Output the first one
                  objOut.Add(arWord(0))
                  ' Show the results
                  For intK = 1 To UBound(arWord)
                    objOut.Add(", " & arWord(intK))
                  Next intK
                  ' Finish this period
                  objOut.Add("</td></tr>")
                End If
              Next intJ
              ' Finish this table
              objOut.Add("</table><p>")
            End With
          Next intI
          ' Finish the report
          objOut.Add("</body></html>")
      End Select
      ' Try to store the result
      strReport = objOut.Text
      If (strReport = "") Then
        ' Warn the user
        Status("No results were found for " & strCategory)
      Else
        ' Calculate this report and put it in place
        strOutFile = strWorkDir & "\" & strCategory & "_Features.htm"
        ' Add report location
        strReport = strOutFile & "<p>" & strReport
        Me.wbReport.DocumentText = strReport
        ' Switch to this tabpage
        Me.TabControl1.SelectedTab = Me.tpReport
        ' Save the report on an appropriate location
        IO.File.WriteAllText(strOutFile, strReport)
        ' Show we are done
        Status("Features of " & strCategory & " are stored in: " & strOutFile)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsShowFeatures_Click error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsFeatDict_Click
  ' Goal:   Make a dictionary of the features for NP, Adv or Verb
  '         1 - NP
  '         2 - Adv
  '         3 - V
  ' History:
  ' 15-02-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsFeatDict_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim strDirIn As String = ""   ' input directory
    Dim strDirOut As String = ""  ' output directory
    Dim strCategory As String     ' The category of features the user would like to renew
    Dim strReport As String       ' Text of report (HTML)
    Dim strOutFile As String      ' One output file
    Dim strName As String = ""    ' Name of category
    Dim strType As String = ""    ' Type of this category (part of the name)
    Dim strVern As String = ""    ' The vernacular lexeme
    Dim strBare As String = ""    ' The bare vernacular item
    Dim intI As Integer           ' Counter
    Dim dtrFound() As DataRow     ' Result of selection statement
    Dim objOut As New StringColl  ' Output string collection

    Try
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      End If
      ' Get the correct category to be worked with
      With frmCategory
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Get the result
            strCategory = .Category
          Case Else
            ' No use to continue
            Exit Sub
        End Select
      End With
      ' Initialise the feature dictionary
      If (Not FeatDictInit(True)) Then Exit Sub
      ' If this is "Adv", then we know what to do
      Select Case strCategory
        Case CAT_NP
        Case CAT_VB
        Case CAT_VBUNACC
        Case CAT_PNORM
        Case CAT_ADV
          ' Initialise
          strType = "" : objOut.Add("<html><body><h2>Feature dictionary for: " & strCategory & "</h2><p>")
          ' Make the feature dictionary
          If (Not MakeFeatDict(strCategory)) Then Exit Sub
          ' Make a datarow collection
          dtrFound = tblFeatDict.Select("", "Bare ASC, Vern ASC, Cat ASC, Type ASC, Period ASC")
          For intI = 0 To dtrFound.Length - 1
            ' Access this element
            With dtrFound(intI)
              ' Add this element to the output
              objOut.Add("<br><b>" & .Item("Vern") & "</b>&nbsp;-&nbsp;<i>" & .Item("Cat") & "/" & _
                         .Item("Type") & "</i>&nbsp;" & .Item("Period"))
            End With
          Next intI
          ' Finish the report
          objOut.Add("</body></html>")
      End Select
      ' Try to store the result
      strReport = objOut.Text
      If (strReport = "") Then
        ' Warn the user
        Status("No results were found for " & strCategory)
      Else
        ' Calculate this report and put it in place
        strOutFile = strWorkDir & "\" & strCategory & "_FeatDict.htm"
        ' Add report location
        strReport = strOutFile & "<p>" & strReport
        Me.wbReport.DocumentText = strReport
        ' Switch to this tabpage
        Me.TabControl1.SelectedTab = Me.tpReport
        ' Save the report on an appropriate location
        IO.File.WriteAllText(strOutFile, strReport)
        ' Show we are done
        Status("Feature dictionary of " & strCategory & " is stored in: " & strOutFile)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsFeatDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsRenewFeatures_Click
  ' Goal:   Calculate features of a user-defined category newly or again 
  '           for the file that is currently opened
  ' History:
  ' 08-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsRenewFeatures_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim strCategory As String   ' The category of features the user would like to renew\
    Dim bCheck As Boolean       ' Whether features should be double-checked
    Dim strFile As String       ' Name of file
    Dim strFeatDef As String = "" ' Feature definition file
    Dim strDir As String        ' Name of directory
    ' Dim arFeatDef(0) As String  ' Feature definitions

    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Exit Sub
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      End If
      ' Make sure interrupt is not on
      bInterrupt = False
      ' Initialise missing pronoun report
      InitMissing()
      ' Initialise pronoun tracking
      InitCatTrack()
      ' Initially clear the error report
      Me.wbError.DocumentText = "(empty)"
      ' Initialisations
      strDir = "" : strFile = ""
      ' Get the correct category to be worked with
      With frmCategory
        ' Make sure options are shown
        .ShowOptions = True
        ' Ask for the user's wishes
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Get the result
            strCategory = .Category
            ' Should we double check those that were done already?
            bCheck = .Check
            ' Do category specific matters
            Select Case strCategory
              Case CAT_ADV
                ' Initialise the feature dictionary
                If (Not MakeFeatDict(strCategory)) Then
                  Status("Unable to initialize the feature dictionary")
                End If
              Case CAT_NP, CAT_GR, CAT_WH, CAT_ANIM, CAT_PNORM
                ' No initialisation needed
              Case CAT_PNORM, CAT_VB
                ' Get the file with the relation between preposition form and normalized (standardized) preposition = Lemma??
                strFeatDef = .FeatFile
                If (strFeatDef = "") OrElse (Not IO.File.Exists(strFeatDef)) Then Status("Cannot open feature definitions") : Exit Sub
                ' make a feature dictionary
                If (Not MakeFeatDict(strCategory, strFeatDef)) Then
                  Status("Unable to initialize feature dictionary")
                  Exit Sub
                End If
              Case CAT_VBUNACC  ' Unaccusatives that have been coded in a parallel *psdx corpus
                ' Get the first file
                strFeatDef = .FeatFile
                If (strFeatDef = "") OrElse (Not IO.File.Exists(strFeatDef)) Then Status("Cannot open feature definitions") : Exit Sub
                ' The file is the first one of similar ones in a direactory of parallel psdx files
                str_UnaccDir = IO.Path.GetDirectoryName(strFeatDef)
              Case CAT_COGN
                ' Get the directory name
                strDir = IO.Path.GetDirectoryName(strCurrentFile)
                If (strDir <> "") Then
                  ' We must have the "stat" subdirectory
                  strDir &= "\stat"
                  ' Create if it does not exist
                  If (Not IO.Directory.Exists(strDir)) Then IO.Directory.CreateDirectory(strDir)
                End If
                ' Make filename
                strFile = strDir & "\" & IO.Path.GetFileNameWithoutExtension(strCurrentFile) & "_cogn.htm"
            End Select
          Case Else
            ' No use to continue
            Exit Sub
        End Select
      End With
      ' Calculate the features for this category
      If (OneDoFeaturesPdx(pdxCurrentFile, strCategory, bCheck)) Then
        ' Try get the pronoun tracking report
        If (HasCatTrack()) Then
          ' Calculate this report and put it in place
          Me.wbReport.DocumentText = CatTrackHtml(strCategory)
          ' Switch to this tabpage
          Me.TabControl1.SelectedTab = Me.tpReport
        End If
        ' Try get the missing pronouns report
        If (HasMissing()) Then
          ' Calculate the report and put it in place
          Me.wbError.DocumentText = MissingHtml(strCategory)
          ' Show there are some errors
          Status("Ready adapting " & strCategory & " features. There are missing " & strCategory & _
                 "s reported in the <Errors> page.")
        Else
          ' Show we are done
          Status("Ready adapting " & strCategory & " features.")
        End If
        ' Make sure the dirty flag is set
        SetDirty(True)
        ' Are we doing cognitive states?
        If (strCategory = CAT_COGN) Then
          ' Switch to this tabpage
          Me.TabControl1.SelectedTab = Me.tpReport
          ' Save the file
          IO.File.WriteAllText(strFile, Me.wbReport.DocumentText, System.Text.Encoding.UTF8)
          Status("File saved: " & strFile)
        End If
      ElseIf (bInterrupt) Then
        ' Check if saving is needed
        If (bNeedSaving) Then SetDirty(True)
        ' We have been interrupted
        Status("Processing has been interrupted")
        bInterrupt = False
      Else
        Status("There was an error renewing " & strCategory & " features")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ToolsRenewFeatures error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuCorpusLoad_Click
  ' Goal:   Load a CrpResults database
  ' History:
  ' 20-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuCorpusLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusLoad.Click
    Dim strInFile As String = ""  ' Location of input file

    Try
      ' Ask user where the results are
      If (Not GetFileName(Me.OpenFileDialog1, strCrpDir, strInFile, "Corpus Results database (*.xml)|*.xml")) Then
        Status("Could not locate input file")
        Exit Sub
      End If
      ' Try load it 
      If (Not CorpusLoad(strInFile)) Then
        ' Do what??
        Status("Unable to load the corpus database")
      End If
      ' Reset the dirty flag
      SetResDirty(False)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuCorpusLoad error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuCorpusAddFile_Click
  ' Goal:   Add the contents of a file to an existing CrpResults database
  ' Note:   The corpus results must already be loaded
  '         Only those with status "Done" are added, and their status is changed into DoneImp
  ' History:
  ' 10-10-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuCorpusAddFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusAddFile.Click
    Dim strInFile As String = ""  ' Location of input file

    Try
      ' Check if anything is loaded
      If (tdlResults Is Nothing) OrElse (tdlResults.Tables("Result").Rows.Count = 0) Then
        ' Tell user he first needs to load results!
        Status("First load results through Corpus/Load or Corpus/Open")
        Exit Sub
      End If
      ' Ask user where the results are
      If (Not GetFileName(Me.OpenFileDialog1, strCrpDir, strInFile, "Corpus Results database (*.xml)|*.xml")) Then
        Status("Could not locate file to be added")
        Exit Sub
      End If
      ' Try load it 
      If (Not CorpusAddFile(strInFile)) Then
        ' Do what??
        Status("Unable to load the corpus database you want to add")
      End If
      ' Reset the dirty flag
      SetResDirty(False)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuCorpusAddFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   DepCurrentShow
  ' Goal:   Show the current psdx file as a dependency one and allow editing it
  ' History:
  ' 04-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function DepCurrentShow() As Boolean
    Dim strDepFile As String  ' File containing the dependency information for the [dgvDep]
    Dim strLang As String = ""  ' Language
    Dim intI As Integer         ' Counter
    Dim intForestId As Integer  ' value of forestId attribute
    Dim bPosOnly As Boolean     ' Only POS or more?

    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Status("First load a psdx file") : Return False
      ' Reset interrupt
      bInterrupt = False
      ' Set save mode
      bDepChg = True
      ' Initialisations
      strDepFile = GetSetDir() & "\CurrentDep-tmp.xml"
      If (Not DeriveLanguage(IO.Path.GetDirectoryName(strCurrentFile), strLang)) Then Status("Cannot derive language from: " & IO.Path.GetDirectoryName(strCurrentFile)) : Return False
      ' Show what we are doing
      Status("Working on dependency...")
      ' Should we create a new dataset?
      If (Not tdlDependency Is Nothing) Then
        ' Properly dispose
        tdlDependency.Dispose()
        tdlDependency = Nothing
      End If
      ' Check if there is any dependency information available
      bPosOnly = (pdxCurrentFile.SelectSingleNode("./descendant::fs[@type='con']/child::f[@name='drel']") Is Nothing)
      ' Check what we need to do
      If (bPosOnly) Then
        Status("Creating POS-only dependency...")
        ' We can look at POS only
        If (Not CreateDepPos(strDepFile, strLang)) Then Status("Could not create dependency xml file") : Return False
      Else
        ' We need to calculate full dependency
        Status("Calculating dependency...")
        ' Make Conll dependency for currently loaded file, if this is needed
        If (Not CreateDepFull(strDepFile, strLang)) Then Status("Could not create dependency xml file") : Return False
      End If
      ' Load the created depfile
      Status("Loading dependency...")
      If (Not ReadDataset("Dependency.xsd", strDepFile, tdlDependency)) Then Status("Could not read DepFile as dataset") : Return False
      ' Show we are working
      ' TODO: ???

      ' Do away with possible previous handlers
      DgvClear(objDepEd)   ' Results database Editor
      ' Initialise the dependency editor
      InitDepEditor()
      ' Color the rows, depending on the forest number
      With Me.dgvDep
        For intI = 0 To .RowCount - 1
          ' Access this row
          With .Rows(intI)
            Debug.Print("forestId = " & .Cells("forestId").Value)
            If (.Cells("forestId").Value / 2 = .Cells("forestId").Value \ 2) Then
              ' Color green
              .DefaultCellStyle.BackColor = Color.LightSkyBlue

            Else
              ' Color white
              .DefaultCellStyle.BackColor = Color.White
            End If
          End With
        Next intI
      End With
      bDepChg = False
      ' Return success
      Status("Dependency ready")
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/DepCurrentShow error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   CorpusLoad
  ' Goal:   Load a corpus database
  ' History:
  ' 21-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CorpusLoad(ByVal strFile As String) As Boolean
    Dim intI As Integer     ' Counter
    Dim intPrev As Integer  ' Number where recent file with same name is already kept
    Dim arSel() As String   ' Array of features
    Dim crsThis As Cursor = Me.Cursor

    Try
      ' Validate
      If (strFile = "") Then Return False
      If (Not IO.File.Exists(strFile)) Then Return False
      ' Reset interrupt
      bInterrupt = False
      ' Set the global filename
      strResultDb = strFile
      '' Save the filename for the next time
      'SetTableSetting(tdlSettings, "CrpLast", strResultDb)
      strCrpLast = strResultDb
      ' Make sure changes are saved
      XmlSaveSettings(strSetFile)
      ' Should we create a new dataset?
      If (Not tdlResults Is Nothing) Then
        ' Properly dispose
        tdlResults.Dispose()
        tdlResults = Nothing
      End If
      ' Show that we are loading
      Status("Loading corpus database: " & IO.Path.GetFileNameWithoutExtension(strFile) & " ...")
      Me.Cursor = Cursors.WaitCursor
      If (Not LoadOneDbFile(strFile)) Then Status("Could not load file") : Return False
      ' ============== DEBUG =============
      '' Load the database
      'If (Not XmlToDataset(strCrpResults, strFile, tdlResults, pdxResults)) Then Status("Could not open corpus database") : Exit Function
      ' ==================================
      ' Show it is loaded
      Status("Initialising editor...")
      ' Log projtype/CRPname
      TryAppendLog("corpus", strFile)
      ' Initialise the editor
      If (Not InitEditors()) Then
        Status("Cannot initialise the database results editor")
        Return False
      End If
      ' Display several more GENERAL values
      With tdlResults.Tables("General").Rows(0)
        Me.tbDbName.Text = .Item("ProjectName").ToString
        Me.tbDbCreated.Text = .Item("Created").ToString
        Me.tbDbSrcDir.Text = .Item("SrcDir").ToString
        Me.tbDbNotes.Text = .Item("Notes").ToString
        Me.tbDbAnalysis.Text = .Item("Analysis").ToString
      End With
      ' Initialise filter:
      strDbaseXpathFilter = ""
      ' Note my own file's name
      Me.tbDbFile.Text = strFile
      ' Set the feature selection combobox
      If (Not SetFeatAnalysis()) Then Status("Could not build feature selection box") : Return False
      ' Check if we need to make another menu item visible
      If (InStr(LCase(strFile), "cleft") > 0) Then
        ' Allow user to use the Corpus/Fix command
        'Me.mnuCorpusFix.Visible = True
        Me.mnuCorpusClefts.Visible = True
      End If
      ' Switch to the correct tabpage
      Me.TabControl1.SelectedTab = Me.tpDbase
      ' Check if this file already exists in the Recent ones
      intPrev = GetCrpRecNumber(strFile)
      ' Found it?
      If (intPrev > 0) Then
        ' Yes, we found it in the recent number --> Swap recent files
        For intI = intPrev To 2 Step -1
          ' Copy settings from [i-1] to [i]
          SetTableSetting(tdlSettings, "CorpusRecent" & intI, GetTableSetting(tdlSettings, "CorpusRecent" & intI - 1))
          ' Make sure changes are saved
          XmlSaveSettings(strSetFile)
        Next intI
      Else
        ' No, so shift the recent ones
        For intI = 5 To 2 Step -1
          ' Copy settings from [i-1] to [i]
          SetTableSetting(tdlSettings, "CorpusRecent" & intI, GetTableSetting(tdlSettings, "CorpusRecent" & intI - 1))
        Next intI
        ' Make sure changes are saved
        XmlSaveSettings(strSetFile)
      End If
      ' Make sure the Recent1 setting gets adapted
      If (strResultDb <> GetTableSetting(tdlSettings, "CorpusRecent1")) Then
        SetTableSetting(tdlSettings, "CorpusRecent1", strResultDb)
        ' Make sure changes are saved
        XmlSaveSettings(strSetFile)
      End If
      ' Make sure the correct stuff is displayed
      CheckCrpRec()
      ' Make sure first record is being shown
      objResEd.SelectDgvId(1)
      dgvDbResult_SelectionChanged(Nothing, Nothing)
      ' Show we are ready
      Me.Cursor = crsThis
      Status("Corpus has been loaded")
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/CorpusLoad error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       SetFeatAnalysis()
  ' Goal:       Build the [cboDbSelect] combobox
  ' History:
  ' 25-08-2014  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Function SetFeatAnalysis() As Boolean
    Dim arSel() As String   ' Array of features
    Dim intI As Integer

    Try
      ' Set the feature selection combobox
      With Me.cboDbSelect
        .Items.Clear()
        ' Fill an array
        arSel = Split(Me.tbDbAnalysis.Text, ";")
        For intI = 0 To UBound(arSel)
          .Items.Add(arSel(intI))
        Next intI
        ' Select the first one
        .SelectedItem = 0
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/SetFeatAnalysis error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  '---------------------------------------------------------------------------------------------------------
  ' Name:       CheckCrpRec()
  ' Goal:       Check which Recent corpus database menu items should be displayed
  ' History:
  ' 10-03-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Sub CheckCrpRec()
    Dim intI As Integer   ' Counter
    Dim strFile As String ' One file
    Dim ctrThis As ToolStripMenuItem = Nothing

    Try
      ' Loop
      For intI = 1 To 5
        ' Check this one
        strFile = GetTableSetting(tdlSettings, "CorpusRecent" & intI)
        ' control depends on which one
        Select Case intI
          Case 1
            ctrThis = Me.mnuCorpusRecent1
          Case 2
            ctrThis = Me.mnuCorpusRecent2
          Case 3
            ctrThis = Me.mnuCorpusRecent3
          Case 4
            ctrThis = Me.mnuCorpusRecent4
          Case 5
            ctrThis = Me.mnuCorpusRecent5
        End Select
        ' We have something?
        If (strFile = "") Then
          ' Reset this one
          ctrThis.Visible = False
        Else
          ' Adapt this one
          With ctrThis
            ' Set the name
            .Text = "&" & intI & " " & strFile
            ' Make it visible
            .Visible = True
          End With
        End If
      Next intI
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/CheckCrpRec error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------------------------------------------------
  ' Name:       CorpusRecent()
  ' Goal:       Open the indicated corpus file.
  ' History:
  ' 10-03-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub CorpusRecent(ByVal intNum As Integer)
    Dim strFile As String  ' File name

    Try
      ' Get the name of the file
      strFile = GetTableSetting(tdlSettings, "CorpusRecent" & CStr(intNum))
      ' Try load this one
      CorpusLoad(strFile)
      ' Reset the dirty flag
      SetResDirty(False)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/CorpusRecent error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
    'Try
    '  ' Get the name of the file
    '  strFile = GetTableSetting(tdlSettings, "Recent" & CStr(intNum))
    '  ' Try open this file
    '  LoadOneFile(strFile)
    'Catch ex As Exception
    '  ' Show error
    '  HandleErr("frmMain/CorpusRecent error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    'End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuCorpusRecent1_Click
  ' Goal:   Load the most recente corpus database
  ' History:
  ' 21-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuCorpusRecent1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusRecent1.Click
    CorpusRecent(1)
  End Sub
  Private Sub mnuCorpusRecent2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusRecent2.Click
    CorpusRecent(2)
  End Sub
  Private Sub mnuCorpusRecent3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusRecent3.Click
    CorpusRecent(3)
  End Sub
  Private Sub mnuCorpusRecent4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusRecent4.Click
    CorpusRecent(4)
  End Sub
  Private Sub mnuCorpusRecent5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusRecent5.Click
    CorpusRecent(5)
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuCorpusReorder_Click
  ' Goal:   Load the corpus database, re-order it
  ' History:
  ' 23-05-2011  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuCorpusReorder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusReorder.Click
    Dim strInFile As String = ""  ' Location of input file

    Try
      ' Ask user where the results are
      If (Not GetFileName(Me.OpenFileDialog1, strCrpDir, strInFile, "Corpus Results database (*.xml)|*.xml")) Then
        Status("Could not locate input file")
        Exit Sub
      End If
      ' Go to the correct tabpage
      Me.TabControl1.SelectedTab = Me.tpDbase
      '' Re order the database
      'CorpusReOrder(strInFile)
      ' Try load it 
      If (Not CorpusLoad(strInFile)) Then
        ' Do what??
        Status("Unable to load the corpus database")
      End If
      ' Perform re-ordering by Period etc.
      If (Not CorpusOrderBy()) Then
        Status("Could  not re-order the corpus")
        Exit Sub
      End If
      ' Set the dirty flag (due to re-ordering)
      SetResDirty(True)
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/CorpusLoadReOrder error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuCorpusSave_Click
  ' Goal:   Save the CrpResults database that is now open
  ' History:
  ' 20-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuCorpusSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusSave.Click
    Try
      ' Perform any actions before changes
      dgvDbResult_BeforeSelChanged(intCurrentResId)
      ' Do the saving
      TrySaveResDb(strResultDb)
      ' Reset the dirty flag
      SetResDirty(False)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/CorpusSave error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuCorpusReport_Click
  ' Goal:   Transform the corpus that is loaded into a report
  ' History:
  ' 03-01-2011  ERK Created
  ' 26-09-2013  ERK Adapted for XmlDocument processing
  ' ------------------------------------------------------------------------------------
  Private Sub mnuCorpusReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusReport.Click
    Dim colThis As New StringColl ' Collection of results in HTML
    Dim colCsv As New StringColl  ' Collection of results in CSV
    Dim bOkay As Boolean = True   ' Whether all is fine
    Dim strFile As String         ' Output file name
    Dim strCsvFile As String      ' File for CSV output
    Dim strText As String         ' HTML results
    Dim ndxRes As XmlNode         ' Result node
    Dim arColumn(0) As String     ' Array of header columns
    Dim intNum As Integer         ' Number of results
    Dim intPtc As Integer         ' Percentage
    Dim intI As Integer           ' Counter
    Dim strAttrList As String = "Period;TextId;Search;forestId;eTreeId;Status;Cat"

    Try
      ' Validate
      If (pdxCrpDbase Is Nothing) Then Status("No corpus results are available") : Exit Sub
      ' Get first result
      ndxRes = pdxCrpDbase.SelectSingleNode("./descendant::Result")
      If (ndxRes Is Nothing) Then Status("No corpus results are available") : Exit Sub
      ' Get number of results
      intNum = 1 + ndxRes.SelectNodes("./following-sibling::Result").Count
      ' Think of a good HTML file name
      strFile = IO.Path.GetDirectoryName(strResultDb) & "\" & IO.Path.GetFileNameWithoutExtension(strResultDb) & ".htm"
      IO.File.WriteAllText(strFile, "", System.Text.Encoding.UTF8)
      ' Think of a good CSV file name
      strCsvFile = IO.Path.GetDirectoryName(strResultDb) & "\" & IO.Path.GetFileNameWithoutExtension(strResultDb) & ".txt"
      IO.File.WriteAllText(strCsvFile, "", System.Text.Encoding.UTF8)
      ' Initialisations
      colThis.Clear() : colCsv.Clear()
      ' Make columns for Period, TextId, Search, Cat
      strText = "Period;TextId;Search;forestId;eTreeId;Status;Cat"
      ' Add columns for the features that are being used
      If (Not GetResultsDbaseFeatures(strAttrList, arColumn, "Name", ndxRes)) Then
        ' Something went wrong
        Status("frmMain/CorpusReport problem") : Exit Sub
      End If
      ' Start producing corpus results in HTML form
      colThis.Add("<html><head>Corpus Database Results</head><body><p>The database contains the following fields:")
      colThis.Add("<table><tr><td>" & Join(arColumn, "</td></tr><tr><td>") & "</td></tr></table>")
      ' Store column names in  csv
      colCsv.Add(Join(arColumn, vbTab))
      ' Walk through the lines
      intI = 0
      While (ndxRes IsNot Nothing)
        ' Show where we are
        intPtc = (intI + 1) * 100 \ intNum
        Status("Corpus Report " & intPtc & "%", intPtc)
        If (bInterrupt) Then Exit Sub
        ' Add Period, TextId, Search, forest, eTree, Cat for this row
        If (Not GetResultsDbaseFeatures(strAttrList, arColumn, "Value", ndxRes)) Then
          ' Something went wrong
          Status("frmMain/CorpusReport problem") : Exit Sub
        End If
        colCsv.Add(Join(arColumn, vbTab))
        ' Flush and clear
        If (colCsv.Count > 512) Then
          colCsv.Flush(strCsvFile) : colCsv.Clear()
        End If
        ' Next line
        ndxRes = ndxRes.NextSibling : intI += 1
      End While

      colThis.Add("<p><table><tr><td>Number of records</td><td>" & intNum & "</td></tr>")
      colThis.Add("<tr><td>Database location (text file)</td><td>" & strCsvFile & "</td></tr>")
      colThis.Add("</table>")
      ' Finish the HTML
      colThis.Add("</body></html>")
      ' Gather the results
      'strText = colThis.Text
      colThis.Flush(strFile) : colThis.Clear()
      colCsv.Flush(strCsvFile) : colCsv.Clear()
      ' Put the results in the correct tabpage
      Status("Opening result...")
      Me.wbReport.Navigate(strFile)
      Status("Ready")
      Me.TabControl1.SelectedTab = Me.tpReport
      ' Show what we have done
      Logging("Corpus database <html> result summary in: " & strFile)
      Logging("Corpus database <csv> result in: " & strCsvFile)
      Status("Saved results in " & strFile)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/CorpusReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuCorpusPrepareSPSS_Click
  ' Goal:   Make three files for processing in SPSS:
  '         1) Tab-separated file
  '         2) Syntax file
  ' History:
  ' 10-01-2013  ERK Created
  ' 26-09-2013  ERK Transformed for work with XmlDocument
  ' ------------------------------------------------------------------------------------
  Private Sub mnuCorpusPrepareSPSS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusPrepareSPSS.Click
    Dim colThis As New StringColl   ' Collection of results in HTML
    Dim colTab As New StringColl    ' Collection of results in tab-delimited format
    Dim colText As New StringColl   ' Results in text format
    Dim colVars As New StringColl   ' Tables of variables
    Dim colSyntax As New StringColl ' SPSS syntax file
    Dim colSynTab As New StringColl ' Tab-delimited importing for SPSS syntax file
    Dim ndxRes As XmlNode           ' Result node
    Dim arColumn(0) As String       ' Array of header columns
    Dim bOkay As Boolean = True     ' Whether all is fine
    Dim strFile As String           ' Output file name
    Dim strTabFile As String        ' Name of tab-delimited output file
    Dim strSpsFile As String        ' SPS file
    Dim strText As String           ' HTML results
    Dim strLast As String           ' Last item should get period added
    Dim colName As New StringColl   ' The names of the different features
    Dim colValue() As StringColl    ' Place that contains the different values
    Dim arFtTypes() As String       ' Feature types
    Dim intI As Integer             ' Counter
    Dim intJ As Integer             ' Counter
    Dim intPtc As Integer           ' Percentage
    Dim intNum As Integer           ' Number of features
    Dim intResCount As Integer      ' Number of results
    Dim strAttrList As String = "Period;TextId;forestId;eTreeId;Cat"
    Dim intAttrCount As Integer = Split(strAttrList, ";").Length
    Dim bUseSyntab As Boolean = False ' SYSTEM setting to use the SYNTAB feature or not

    Try
      ' Validate
      If (pdxCrpDbase Is Nothing) Then Status("No corpus results are available") : Exit Sub
      ' Get first result
      ndxRes = pdxCrpDbase.SelectSingleNode("./descendant::Result")
      If (ndxRes Is Nothing) Then Status("No corpus results are available") : Exit Sub
      ' Get number of results
      intResCount = 1 + ndxRes.SelectNodes("./following-sibling::Result").Count
      ' Initialisations
      colThis.Clear() : colTab.Clear()
      strTabFile = IO.Path.GetDirectoryName(strResultDb) & "\" & IO.Path.GetFileNameWithoutExtension(strResultDb) & "_spss.txt"
      ' Clear file
      IO.File.WriteAllText(strTabFile, "", System.Text.Encoding.UTF8)
      ' Start producing corpus results in HTML form
      colThis.Add("<h3>Datatable to be copied to SPSS (numerical)</h3><table>")
      colText.Add("<h3>Datatable to be copied to SPSS (text)</h3><table>")
      ' Start producing spss syntax file
      colSyntax.Add("*********************************")
      colSyntax.Add("* Define variable labels")
      colSyntax.Add("*********************************" & vbCrLf)
      ' Make columns for Period, TextId, Search, Cat
      If (Not GetResultsDbaseFeatures(strAttrList, arColumn, "Name", ndxRes)) Then Status("Problem getting feature list") : Exit Sub
      ' Add to tabulated output
      colTab.Add(Join(arColumn, vbTab))
      ' Get the feature names into [colName]
      For intI = intAttrCount To arColumn.Length - 1
        colName.Add(arColumn(intI))
      Next intI
      ' Get number of features
      intNum = colName.Count
      ' Get feature settings
      With frmSpss
        ' Ask for features
        .Features = Split(colName.Semi, ";")
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.Yes, Windows.Forms.DialogResult.OK
            ' Now we know what the types are, retrieve them...
            arFtTypes = .FtTypes
          Case Else
            ' Exit
            Status("Function aborted")
            Exit Sub
        End Select
      End With
      ' Make sure we are being seen
      Application.DoEvents()
      If (bUseSyntab) Then
        ' Produce the spss syntax-file data needed to import the tab-delimited data file
        colSynTab.Add("*************************************************************************")
        colSynTab.Add("* Open the tab-delimited data files ")
        colSynTab.Add("* N.B: If loading the text file does not work, then delete this section! ")
        colSynTab.Add("*************************************************************************" & vbCrLf)
        colSynTab.Add("GET DATA")
        colSynTab.Add("  /TYPE=TXT")
        colSynTab.Add("  /FILE=" & """" & strTabFile & """")
        colSynTab.Add("  /DELCASE=LINE")
        colSynTab.Add("  /DELIMITERS=" & """" & "\t" & """")
        colSynTab.Add("  /ARRANGEMENT=DELIMITED")
        colSynTab.Add("  /FIRSTCASE=2")
        colSynTab.Add("  /IMPORTCASE=ALL")
        colSynTab.Add("  /VARIABLES=")
        colSynTab.Add("  Period A3")
        colSynTab.Add("  TextId A20")
        colSynTab.Add("  forestId F4.0")
        colSynTab.Add("  eTreeId F6.0")
        colSynTab.Add("  Cat A20")
        ' Walk through all the variables ...
        For intI = 0 To arFtTypes.Length - 1
          ' is this the last or not?
          If (intI = arFtTypes.Length - 1) Then
            ' This is the last
            strLast = "."
          Else
            ' Not the last
            strLast = ""
          End If
          ' Check what type this is
          Select Case arFtTypes(intI)
            Case "Nos"
              ' Treat as text --> Don't allow more than 256 characters
              colSynTab.Add("  " & colName.Item(intI) & " A256" & strLast)
            Case "Ind", "Dep"
              ' Must be numerical! --> don't allow more than 6 numbers (1 million entries)
              colSynTab.Add("  " & colName.Item(intI) & " F6.0" & strLast)
          End Select
        Next intI
        colSynTab.Add("CACHE.")
        colSynTab.Add("EXECUTE.")
        colSynTab.Add("DATASET NAME DataSet1 WINDOW=FRONT." & vbCrLf)
      End If

      ' Prepare the feature value arrays (this is to convert feature string into value
      ReDim colValue(0 To intNum - 1)
      For intI = 0 To colValue.Length - 1
        ' Make room for this one
        colValue(intI) = New StringColl
      Next intI
      ' Walk through the lines in <Result>
      intI = 0
      While (ndxRes IsNot Nothing)
        ' Show where we are
        intPtc = (intI + 1) * 100 \ intResCount
        Status("Spss preparing " & intPtc & "%", intPtc)
        If (bInterrupt) Then Exit Sub
        ' Add Period, TextId, Search, forest, eTree, Cat for this row
        If (Not GetResultsDbaseFeatures(strAttrList, arColumn, "Value", ndxRes)) Then
          ' Something went wrong
          Status("frmMain/CorpusPrepareSpss problem") : Exit Sub
        End If
        ' Walk the features
        For intJ = 0 To colName.Count - 1
          ' What kind of feature is this?
          Select Case arFtTypes(intJ)
            Case "Nos"
              ' No changes are needed for this feature
            Case "Dep", "Ind"
              ' Add this value to the possible values for this item
              colValue(intJ).AddUnique(arColumn(intAttrCount + intJ))
              ' Get the numerical value for this feature value
              arColumn(intAttrCount + intJ) = colValue(intJ).Find(arColumn(intAttrCount + intJ)) + 1
            Case Else
              Status("Unrecognized feature type [" & arFtTypes(intJ) & "]")
              Exit Sub
          End Select
        Next intJ
        ' Process changes in the [colTab] and possibly straight into [strTabFile]
        colTab.Add(Join(arColumn, vbTab))
        ' Flush and clear
        If (colTab.Count > 512) Then colTab.Flush(strTabFile) : colTab.Clear()
        ' Next line
        ndxRes = ndxRes.NextSibling : intI += 1
      End While
      ' Flush the last part of the results
      colTab.Flush(strTabFile) : colTab.Clear()

      ' Make the tables for the variables
      For intI = 0 To colName.Count - 1
        ' Only add table for values that need it
        If (arFtTypes(intI) <> "Nos") Then
          ' Create a table for the values of this variable
          colVars.Add("<h3>" & colName.Item(intI) & "</h3><table><tr><td>Name</td><td>Value</td></tr>")
          ' Create an SPSS division
          colSyntax.Add("* Feature: " & colName.Item(intI) & vbCrLf & vbCrLf & _
                        "VALUE LABELS " & colName.Item(intI))
          For intJ = 0 To colValue(intI).Count - 1
            ' Add a row for this value
            colVars.Add("<tr><td>" & colValue(intI).Item(intJ) & "</td><td>" & intJ + 1 & "</td></tr>")
            ' Add an entry for this value -- add a period for the last line
            colSyntax.Add(intJ + 1 & "  '" & UnQuote(colValue(intI).Item(intJ)) & "'" & _
                          IIf(intJ = colValue(intI).Count - 1, ".", ""))
          Next intJ
          ' Finish this table
          colVars.Add("</table>")
          ' Finish syntax
          colSyntax.Add("EXECUTE.")
        End If
      Next intI
      ' Make preamble for all files
      strFile = IO.Path.GetDirectoryName(strResultDb) & "\" & IO.Path.GetFileNameWithoutExtension(strResultDb)
      ' Gather the results
      strText = "<html><head>Corpus Database Results</head><body><p>" & _
                "<table><tr><td>Corpus results:</td><td>" & strFile & "_corpus.htm</td></tr>" & _
                "<tr><td>SPSS syntax file</td><td>" & strFile & ".sps</td></tr>" & _
                "<tr><td>SPSS tab-delimited data</td><td>" & strTabFile & "</td></tr>" & _
                "</table><p>" & _
                colVars.Text & _
                "</table></body></html>"
      ' Put the results in the correct tabpage
      Me.wbReport.DocumentText = strText
      Me.TabControl1.SelectedTab = Me.tpReport
      ' (1) Think of a good HTML file name
      strFile = IO.Path.GetDirectoryName(strResultDb) & "\" & IO.Path.GetFileNameWithoutExtension(strResultDb) & "_corpus.htm"
      ' Put results into this HTML file in UTF8
      IO.File.WriteAllText(strFile, strText, System.Text.Encoding.UTF8)
      Logging("Corpus results overview: " & strFile)
      ' (2) Think of a good SPSS syntax file name
      strSpsFile = IO.Path.GetDirectoryName(strResultDb) & "\" & IO.Path.GetFileNameWithoutExtension(strResultDb) & ".sps"
      ' Prepare the syntax text
      strText = colSynTab.Text & colSyntax.Text
      ' Put results into this HTML file in UTF8
      IO.File.WriteAllText(strSpsFile, strText, System.Text.Encoding.UTF8)
      Logging("SPSS syntax file: " & strSpsFile)
      ' (3) Think of a good SPSS file name for the tab-delimited text file
      '' Prepare the tab-delimited text
      'strText = colTab.Text
      '' Put results into this HTML file in UTF8
      'IO.File.WriteAllText(strTabFile, strText, System.Text.Encoding.UTF8)
      Logging("SPSS tab-delimited data: " & strTabFile)
      ' Show what we have done
      Status("Result file names: see [General] tab")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/CorpusPrepareSPSS error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuCorpusPrepareSPSS_Click
  ' Goal:   Make an HTML report where we prepare the data to be exported into SPSS
  ' History:
  ' 10-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub CorpusPrepareSpss_Org()
    Dim colThis As New StringColl   ' Collection of results in HTML
    Dim colTab As New StringColl    ' Collection of results in tab-delimited format
    Dim colText As New StringColl   ' Results in text format
    Dim colVars As New StringColl   ' Tables of variables
    Dim colSyntax As New StringColl ' SPSS syntax file
    Dim colSynTab As New StringColl ' Tab-delimited importing for SPSS syntax file
    Dim dtrThis As DataRow          ' One datarow
    Dim dtrChildren() As DataRow    ' Feature children of <Result>
    Dim dtrFound() As DataRow       ' Sorted collection
    Dim bOkay As Boolean = True     ' Whether all is fine
    Dim strFile As String           ' Output file name
    Dim strTabFile As String        ' Name of tab-delimited output file
    Dim strSpsFile As String        ' SPS file
    Dim strText As String           ' HTML results
    Dim strTab As String            ' Tab-delimited results
    Dim strLast As String           ' Last item should get period added
    Dim colName As New StringColl   ' The names of the different features
    Dim colValue() As StringColl    ' Place that contains the different values
    Dim arFtTypes() As String       ' Feature types
    Dim intI As Integer             ' Counter
    Dim intJ As Integer             ' Counter
    Dim intPtc As Integer           ' Percentage
    Dim intNum As Integer           ' Number of features
    Dim intName As Integer          ' The numerical value of this feature name
    Dim intValue As Integer         ' The numerical value of this feature value
    ' Dim objSpss As Object           ' SPSS object

    Try
      ' Check if there are any results
      If (tdlResults Is Nothing) Then
        bOkay = False
      ElseIf (tdlResults.Tables("Result") Is Nothing) Then
        bOkay = False
      ElseIf (tdlResults.Tables("Result").Rows.Count = 0) Then
        bOkay = False
      End If
      If (Not bOkay) Then
        Status("No corpus results are available")
        Exit Sub
      End If
      ' Initialisations
      colThis.Clear() : colTab.Clear()
      strTabFile = IO.Path.GetDirectoryName(strResultDb) & "\" & IO.Path.GetFileNameWithoutExtension(strResultDb) & "_spss.txt"
      ' Start producing corpus results in HTML form
      colThis.Add("<h3>Datatable to be copied to SPSS (numerical)</h3><table>")
      colText.Add("<h3>Datatable to be copied to SPSS (text)</h3><table>")
      ' Start producing spss syntax file
      colSyntax.Add("*********************************")
      colSyntax.Add("* Define variable labels")
      colSyntax.Add("*********************************" & vbCrLf)
      ' Make columns for Period, TextId, Search, Cat
      strText = "<tr><td>Period</td><td>TextId</td><td>forestId</td><td>eTreeId</td><td>Cat</td>"
      strTab = "Period" & vbTab & "TextId" & vbTab & "forestId" & vbTab & "eTreeId" & vbTab & "Cat"
      ' Find out what other columns there are
      dtrThis = tdlResults.Tables("Result").Rows(0)
      ' Check if this datarow has any children that contain features
      dtrChildren = dtrThis.GetChildRows("Result_Feature") : intNum = 0
      For intI = 0 To dtrChildren.Length - 1
        ' Skip the Pde variable
        If (dtrChildren(intJ).Item("Name") <> "Pde") Then
          ' Add the heading for this child
          strText &= "<td>" & dtrChildren(intI).Item("Name") & "</td>"
          strTab &= vbTab & dtrChildren(intI).Item("Name")
          ' Increment the feature number
          intNum += 1
          ' Add the feature name
          colName.Add(dtrChildren(intI).Item("Name").ToString)
        End If
      Next intI
      ' Finish this line
      strText &= "</tr>"
      colThis.Add(strText)
      colText.Add(strText)
      colTab.Add(strTab)
      ' Get feature settings
      With frmSpss
        ' Ask for features
        .Features = Split(colName.Semi, ";")
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.Yes, Windows.Forms.DialogResult.OK
            ' Now we know what the types are, retrieve them...
            arFtTypes = .FtTypes
          Case Else
            ' Exit
            Status("Function aborted")
            Exit Sub
        End Select
      End With
      ' Make sure we are being seen
      Application.DoEvents()
      ' Produce the spss syntax-file data needed to import the tab-delimited data file
      colSynTab.Add("*************************************************************************")
      colSynTab.Add("* Open the tab-delimited data files ")
      colSynTab.Add("* N.B: If loading the text file does not work, then delete this section! ")
      colSynTab.Add("*************************************************************************" & vbCrLf)
      colSynTab.Add("GET DATA")
      colSynTab.Add("  /TYPE=TXT")
      colSynTab.Add("  /FILE=" & """" & strTabFile & """")
      colSynTab.Add("  /DELCASE=LINE")
      colSynTab.Add("  /DELIMITERS=" & """" & "\t" & """")
      colSynTab.Add("  /ARRANGEMENT=DELIMITED")
      colSynTab.Add("  /FIRSTCASE=2")
      colSynTab.Add("  /IMPORTCASE=ALL")
      colSynTab.Add("  /VARIABLES=")
      colSynTab.Add("  Period A3")
      colSynTab.Add("  TextId A20")
      colSynTab.Add("  forestId F4.0")
      colSynTab.Add("  eTreeId F6.0")
      colSynTab.Add("  Cat A20")
      ' Walk through all the variables ...
      For intI = 0 To arFtTypes.Length - 1
        ' is this the last or not?
        If (intI = arFtTypes.Length - 1) Then
          ' This is the last
          strLast = "."
        Else
          ' Not the last
          strLast = ""
        End If
        ' Check what type this is
        Select Case arFtTypes(intI)
          Case "Nos"
            ' Treat as text --> Don't allow more than 256 characters
            colSynTab.Add("  " & colName.Item(intI) & " A256" & strLast)
          Case "Ind", "Dep"
            ' Must be numerical! --> don't allow more than 6 numbers (1 million entries)
            colSynTab.Add("  " & colName.Item(intI) & " F6.0" & strLast)
        End Select
      Next intI
      colSynTab.Add("CACHE.")
      colSynTab.Add("EXECUTE.")
      colSynTab.Add("DATASET NAME DataSet1 WINDOW=FRONT." & vbCrLf)

      ' Prepare the feature value arrays
      ReDim colValue(0 To intNum - 1)
      For intI = 0 To colValue.Length - 1
        ' Make room for this one
        colValue(intI) = New StringColl
      Next intI
      ' Sort the lines
      dtrFound = tdlResults.Tables("Result").Select("", "Status ASC, ResId ASC")
      ' Go through all the lines
      For intI = 0 To dtrFound.Length - 1
        ' Show where we are
        intPtc = (intI * 1) * 100 \ dtrFound.Length
        Status("Spss preparing " & intPtc & "%", intPtc)
        ' Add Period, TextId, Search, Cat for this row
        ' Added: also forestId and eTreeId
        strText = "<tr><td>" & dtrFound(intI).Item("Period") & "</td>" & _
          "<td>" & dtrFound(intI).Item("TextId") & "</td>" & _
          "<td>" & dtrFound(intI).Item("forestId") & "</td>" & _
          "<td>" & dtrFound(intI).Item("eTreeId") & "</td>" & _
          "<td>" & dtrFound(intI).Item("Cat") & "</td>"
        With dtrFound(intI)
          strTab = .Item("Period") & vbTab & .Item("TextId") & vbTab & .Item("forestId") & vbTab & .Item("eTreeId") & vbTab & .Item("Cat")
        End With
        ' Add the features for this row
        dtrChildren = dtrFound(intI).GetChildRows("Result_Feature")
        For intJ = 0 To dtrChildren.Length - 1
          ' Skip the Pde variable
          If (dtrChildren(intJ).Item("Name") <> "Pde") Then
            ' What kind of feature is this?
            Select Case arFtTypes(intJ)
              Case "Nos"
                ' Add the text-value of this feature
                strText &= "<td>" & dtrChildren(intJ).Item("Value") & "</td>"
                strTab &= vbTab & dtrChildren(intJ).Item("Value")
              Case "Dep", "Ind"
                ' Get the numerical value for this feature name
                intName = colName.Find(dtrChildren(intJ).Item("Name").ToString)
                ' Add this value to the possible values for this item
                colValue(intName).AddUnique(dtrChildren(intJ).Item("Value"))
                ' Get the numerical value for this feature value
                intValue = colValue(intName).Find(dtrChildren(intJ).Item("Value")) + 1
                ' Add the value of this feature child
                strText &= "<td>" & intValue & "</td>"
                strTab &= vbTab & intValue
              Case Else
                Status("Unrecognized feature type [" & arFtTypes(intJ) & "]")
                Exit Sub
            End Select
          End If
        Next intJ
        ' Finish this line
        strText &= "</tr>"
        colThis.Add(strText)
        colTab.Add(strTab)
        ' Add Period, TextId, Search, Cat for this row
        strText = "<tr><td>" & dtrFound(intI).Item("Period") & "</td>" & _
            "<td>" & dtrFound(intI).Item("forestId") & "</td>" & _
            "<td>" & dtrFound(intI).Item("eTreeId") & "</td>" & _
            "<td>" & dtrFound(intI).Item("TextId") & "</td>" & _
            "<td>" & dtrFound(intI).Item("Cat") & "</td>"
        ' Prepare the text output
        For intJ = 0 To dtrChildren.Length - 1
          ' Skip the Pde variable
          If (dtrChildren(intJ).Item("Name") <> "Pde") Then
            ' Add the value of this feature child
            strText &= "<td>" & dtrChildren(intJ).Item("Value") & "</td>"
          End If
        Next intJ
        ' Finish this line
        strText &= "</tr>"
        colText.Add(strText)
      Next intI
      'With tdlResults.Tables("Result")
      '  For intI = 0 To .Rows.Count - 1
      '  Next intI
      'End With
      ' Finish the table
      colThis.Add("</table>")
      colText.Add("</table>")
      ' Make the tables for the variables
      For intI = 0 To colName.Count - 1
        ' Only add table for values that need it
        If (arFtTypes(intI) <> "Nos") Then
          ' Create a table for the values of this variable
          colVars.Add("<h3>" & colName.Item(intI) & "</h3><table><tr><td>Name</td><td>Value</td></tr>")
          ' Create an SPSS division
          colSyntax.Add("* Feature: " & colName.Item(intI) & vbCrLf & vbCrLf & "VALUE LABELS " & colName.Item(intI))
          For intJ = 0 To colValue(intI).Count - 1
            ' Add a row for this value
            colVars.Add("<tr><td>" & colValue(intI).Item(intJ) & "</td><td>" & intJ + 1 & "</td></tr>")
            ' Add an entry for this value -- add a period for the last line
            colSyntax.Add(intJ + 1 & "  '" & colValue(intI).Item(intJ) & "'" & _
                          IIf(intJ = colValue(intI).Count - 1, ".", ""))
          Next intJ
          ' Finish this table
          colVars.Add("</table>")
          ' Finish syntax
          colSyntax.Add("EXECUTE.")
        End If
      Next intI
      ' Gather the results
      strText = "<html><head>Corpus Database Results</head><body><p>" & colVars.Text & colThis.Text & colText.Text & "</body></html>"
      ' Put the results in the correct tabpage
      Me.wbReport.DocumentText = strText
      Me.TabControl1.SelectedTab = Me.tpReport
      ' (1) Think of a good HTML file name
      strFile = IO.Path.GetDirectoryName(strResultDb) & "\" & IO.Path.GetFileNameWithoutExtension(strResultDb) & "_corpus.htm"
      ' Put results into this HTML file in UTF8
      IO.File.WriteAllText(strFile, strText, System.Text.Encoding.UTF8)
      Logging("Corpus results: " & strFile)
      ' (2) Think of a good SPSS syntax file name
      strSpsFile = IO.Path.GetDirectoryName(strResultDb) & "\" & IO.Path.GetFileNameWithoutExtension(strResultDb) & ".sps"
      ' Prepare the syntax text
      strText = colSynTab.Text & colSyntax.Text
      ' Put results into this HTML file in UTF8
      IO.File.WriteAllText(strSpsFile, strText, System.Text.Encoding.UTF8)
      Logging("SPSS syntax file: " & strSpsFile)
      ' (3) Think of a good SPSS file name for the number-table
      strFile = IO.Path.GetDirectoryName(strResultDb) & "\" & IO.Path.GetFileNameWithoutExtension(strResultDb) & "_spss.htm"
      ' Prepare the syntax text
      strText = colThis.Text
      ' Put results into this HTML file in UTF8
      IO.File.WriteAllText(strFile, strText, System.Text.Encoding.UTF8)
      Logging("SPSS numerical data: " & strFile)
      ' (4) Think of a good SPSS file name for the tab-delimited text file
      ' Prepare the tab-delimited text
      strText = colTab.Text
      ' Put results into this HTML file in UTF8
      IO.File.WriteAllText(strTabFile, strText, System.Text.Encoding.UTF8)
      Logging("SPSS tab-delimited data: " & strTabFile)
      ' Show what we have done
      Status("Result file names: see [General] tab")
      '' Ask if we should start a thread with the SPSS file
      'Select Case MsgBox("Would you like to start SPSS with the results just made?")
      '  Case MsgBoxResult.Yes
      '    Shell(strSpsFile, AppWinStyle.NormalFocus)
      'End Select
      '' Ask if we should start a thread with the SPSS file
      'Select Case MsgBox("Alternative to start SPSS with the results just made?")
      '  Case MsgBoxResult.Yes
      '    objSpss = GetObject(strSpsFile)
      'End Select
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/CorpusPrepareSPSS_Org error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuCorpusBayesian_Click
  ' Goal:   Perform a naive Bayesian classifier and report on the results
  ' History:
  ' 12-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuCorpusBayesian_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusBayesian.Click
    Dim colThis As New StringColl   ' Collection of results in HTML
    Dim colTab As New StringColl    ' Collection of results in tab-delimited format
    Dim colText As New StringColl   ' Results in text format
    Dim dtrThis As DataRow          ' One datarow
    Dim dtrChildren() As DataRow    ' Feature children of <Result>
    Dim dtrFound() As DataRow       ' Sorted collection
    Dim dtrTest() As DataRow        ' Test set
    Dim bOkay As Boolean = True     ' Whether all is fine
    Dim strFile As String           ' Output file name
    Dim strText As String           ' HTML results
    Dim strClassName As String      ' Name of class we are working with
    Dim strFeatName As String       ' Name of feature
    Dim strFeatVal As String        ' Value of feature
    Dim colName As New StringColl   ' The names of the different features
    Dim colValue() As StringColl    ' Place that contains the different values
    Dim colClass As New StringColl  ' The names of the different class outputs
    Dim arClassPrior() As Double    ' class prior probabilities
    Dim arClassPred() As Double     ' Prediction for each class for one item in test set
    Dim arFtTypes() As String       ' Feature types
    Dim intI As Integer             ' Counter
    Dim intJ As Integer             ' Counter
    Dim intK As Integer             ' Counter
    Dim intPtc As Integer           ' Percentage
    Dim intPtcTrain As Integer = 70 ' 70% training set and 30% test set (fixed)
    Dim intClassFreq As Integer     ' Number of occurrances for this class in total training set
    Dim intFeatClassFreq As Integer ' Number of occurrances for this feature/class combination
    Dim intClassMax As Integer      ' Best class
    Dim intScore As Double          ' Score
    Dim intNum As Integer           ' Number of features
    Dim intName As Integer          ' The numerical value of this feature name
    Dim tdlBayes As DataSet = Nothing
    Dim tblActPred As New DataTable ' Datatable for the results
    Dim tblBayes As New DataTable   ' Training set: feature name, value and class
    'Dim arKey(0 To 3) As DataColumn

    Try
      ' Check if there are any results
      If (tdlResults Is Nothing) Then
        bOkay = False
      ElseIf (tdlResults.Tables("Result") Is Nothing) Then
        bOkay = False
      ElseIf (tdlResults.Tables("Result").Rows.Count = 0) Then
        bOkay = False
      End If
      If (Not bOkay) Then
        Status("No corpus results are available")
        Exit Sub
      End If
      ' Initialisations
      If (Not CreateDataSet("Bayes.xsd", tdlBayes)) Then Exit Sub
      tblActPred = tdlBayes.Tables("ActPred")
      tblBayes = tdlBayes.Tables("Naive")
      colThis.Clear()
      ' Start producing corpus results in HTML form
      ' Find out what other columns there are
      dtrThis = tdlResults.Tables("Result").Rows(0)
      ' Check if this datarow has any children that contain features
      dtrChildren = dtrThis.GetChildRows("Result_Feature") : intNum = 0
      For intI = 0 To dtrChildren.Length - 1
        ' Skip the Pde variable
        If (dtrChildren(intJ).Item("Name") <> "Pde") Then
          ' Increment the feature number
          intNum += 1
          ' Add the feature name
          colName.Add(dtrChildren(intI).Item("Name").ToString)
        End If
      Next intI
      ' Get feature settings
      With frmSpss
        ' Ask for features
        .Features = Split(colName.Semi, ";")
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.Yes, Windows.Forms.DialogResult.OK
            ' Now we know what the types are, retrieve them...
            arFtTypes = .FtTypes
            ' The flag "UseCat" should have been set. Check this...
            If (Not .UseCatAsDependant) Then
              MsgBox("Corpus/Bayesian only works with the flag [Use Cat as dependant] set")
              Status("There was an error")
              Exit Sub
            End If
          Case Else
            ' Exit
            Status("Function aborted")
            Exit Sub
        End Select
      End With
      ' Make sure we are being seen
      Application.DoEvents()
      ' Make sure we are not interrupted
      dgvDbResult.DataSource = Nothing
      ' Get all the lines
      dtrFound = tdlResults.Tables("Result").Select("", "Status ASC, ResId ASC")
      ' Divide the lines in training and test set by using the [Select] tag
      For intI = 0 To dtrFound.Length - 1
        ' Where are we?
        intPtc = (intI + 1) * 100 \ dtrFound.Length
        Status("Dividing training/test " & intPtc & "%", intPtc)
        ' Check training percentage
        If (Rnd() * 100 <= intPtcTrain) Then
          ' Add it to the trainingset
          dtrFound(intI).Item("Select") = "train"
        Else
          ' Add it to the testset
          dtrFound(intI).Item("Select") = "test"
        End If
      Next intI
      Status("Accepting changes")
      tdlResults.AcceptChanges()
      Status("Continuing training...")
      ' Get all the training set lines
      dtrFound = tdlResults.Tables("Result").Select("Select='train'", "ResId ASC")
      ' Allow interrupt again
      If (Not InitResultsEditor()) Then Status("Could not re-initialize") : Exit Sub
      ' Make sure the dirty flag is cleared
      bResDirty = False
      ' Calculate the class prior values -- the probabilities that a particular class occurs
      For intI = 0 To dtrFound.Length - 1
        colClass.AddUnique(dtrFound(intI).Item("Cat").ToString)
      Next intI
      ' Make room for class prior and class prediction
      ReDim arClassPrior(0 To colClass.Count - 1)
      ReDim arClassPred(0 To colClass.Count - 1)
      ' Retrieve frequencies and divide them by the total amount of samples --> p(C[i])
      For intI = 0 To arClassPrior.Length - 1
        arClassPrior(intI) = colClass.Freq(intI) / dtrFound.Length
      Next intI
      ' Prepare the feature value arrays
      ReDim colValue(0 To intNum - 1)
      For intI = 0 To colValue.Length - 1
        ' Make room for this one
        colValue(intI) = New StringColl
      Next intI
      ' Go through all the lines of the training set
      intK = 0
      For intI = 0 To dtrFound.Length - 1
        ' Show where we are
        intPtc = (intI * 1) * 100 \ dtrFound.Length
        Status("Bayes training " & intPtc & "%", intPtc)
        ' Check the feature values for this row
        dtrChildren = dtrFound(intI).GetChildRows("Result_Feature")
        For intJ = 0 To dtrChildren.Length - 1
          ' Skip the Pde variable
          If (dtrChildren(intJ).Item("Name") <> "Pde") Then
            ' What kind of feature is this?
            Select Case arFtTypes(intJ)
              Case "Nos"
                ' No action
              Case "Dep", "Ind"
                ' Get the numerical value for this feature name
                intName = colName.Find(dtrChildren(intJ).Item("Name").ToString)
                ' Add this value to the possible values for this item
                colValue(intName).AddUnique(dtrChildren(intJ).Item("Value"))
                ' Add the combination feature name/value and class into the table
                intK += 1
                tblBayes.Rows.Add(intK, dtrChildren(intJ).Item("Name"), dtrChildren(intJ).Item("Value"), dtrFound(intI).Item("Cat"))
              Case Else
                Status("Unrecognized feature type [" & arFtTypes(intJ) & "]")
                Exit Sub
            End Select
          End If
        Next intJ
      Next intI
      Status("Starting testing...")
      ' Get the test set
      dtrTest = tdlResults.Tables("Result").Select("Select='test'", "ResId ASC")
      ' Go through the test set
      For intI = 0 To dtrTest.Length - 1
        ' Show where we are
        intPtc = (intI * 1) * 100 \ dtrTest.Length
        Status("Bayes testing " & intPtc & "%", intPtc)
        ' Determine the prediction for each of the classes
        For intJ = 0 To arClassPred.Length - 1
          ' Get the name of this class
          strClassName = colClass.Item(intJ)
          ' Initialise the prediction with the p(C[i])
          arClassPred(intJ) = arClassPrior(intJ)
          ' Get the frequency for this class
          intClassFreq = colClass.Freq(intJ)
          ' Visit all the features
          dtrChildren = dtrTest(intI).GetChildRows("Result_Feature")
          For intK = 0 To dtrChildren.Length - 1
            ' Make sure this is an independant variable
            If (arFtTypes(intK) = "Ind") AndAlso (dtrChildren(intK).Item("Name") <> "Pde") Then
              ' Get the name and value of this feature
              strFeatName = dtrChildren(intK).Item("Name")
              strFeatVal = dtrChildren(intK).Item("Value")
              ' Extract the number of times this feature name/value for this class occurs
              intFeatClassFreq = tblBayes.Select("Fname='" & strFeatName & "' AND Fvalue='" & strFeatVal & "'" & _
                                                 " AND Class='" & strClassName & "'").Length
              ' Check how often the feature with this name and value occurs for this particular class
              'intFeatClassFreq = tdlResults.Tables("Feature").Select("Parent.Select='train' AND Name = '" & _
              '    strFeatName & "' AND Value = '" & strFeatVal.Replace("'", "''") & "'" & _
              '    " AND Parent.Cat='" & strClassName & "'").Length
              If (intFeatClassFreq < 10) Then intFeatClassFreq += 1
              ' ============ DEBUG ============
              ' tdlResults.WriteXml("D:\Data files\Corpora\English\xml\Cesaxed\Test.xml")
              ' If (intFeatClassFreq > 1) Then Stop
              If (intFeatClassFreq > intClassFreq) Then Stop
              ' ===============================
              ' Adapt the prediction value
              arClassPred(intJ) *= intFeatClassFreq / intClassFreq
            End If
          Next intK
        Next intJ
        ' Check out which of the classes gets the highest score
        intClassMax = -1 : intScore = 0
        For intJ = 0 To arClassPred.Length - 1
          If (arClassPred(intJ) > intScore) Then
            intClassMax = intJ : intScore = arClassPred(intJ)
          End If
        Next intJ
        ' Process actual and prediction
        tblActPred.Rows.Add(intI, dtrTest(intI).Item("Cat").ToString, colClass.Item(intClassMax))
      Next intI
      ' Save the table with actual and predicted values as XML file
      ' (1) Think of a good HTML file name
      strFile = IO.Path.GetDirectoryName(strResultDb) & "\" & IO.Path.GetFileNameWithoutExtension(strResultDb) & "_bayes.xml"
      ' (2) Save it
      tdlBayes.WriteXml(strFile)
      ' Make array of actual versus predicted
      colThis.Add("<table><tr><td>Actual</td><td>Predicted</td><td>Count</td></tr>")
      For intI = 0 To arClassPred.Length - 1
        For intJ = 0 To arClassPred.Length - 1
          colThis.Add("<tr><td>" & colClass.Item(intI) & "</td><td>" & colClass.Item(intJ) & "</td>" & _
                      "<td>" & tblActPred.Select("Actual='" & colClass.Item(intI) & _
                                        "' AND Pred='" & colClass.Item(intJ) & "'").Length & "</td>" & _
                      "</tr>")
        Next intJ
      Next intI
      ' Finish table
      colThis.Add("</table>")
      ' Gather the results
      strText = "<html><head>Corpus Database Results</head><body><p>" & colThis.Text & "</body></html>"
      ' Put the results in the correct tabpage
      Me.wbReport.DocumentText = strText
      Me.TabControl1.SelectedTab = Me.tpReport
      ' (1) Think of a good HTML file name
      strFile = IO.Path.GetDirectoryName(strResultDb) & "\" & IO.Path.GetFileNameWithoutExtension(strResultDb) & "_bayes.htm"
      ' Put results into this HTML file in UTF8
      IO.File.WriteAllText(strFile, strText, System.Text.Encoding.UTF8)
      Logging("Bayes results: " & strFile)
      ' Show what we have done
      Status("Result file names: see [General] tab")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/CorpusBayesian error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuCorpusPrepareTimbl_Click
  ' Goal:   Make a tab-separated trainingset and testset for use in TIMBL
  '         - The feature values may be symbolic, so not numbers?
  '         - Take 70% trainingset + 30% testset, or make this user-adjustable?
  '         - Allow user to select the periods that need to be included?
  ' History:
  ' 15-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuCorpusPrepareTimbl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusPrepareTimbl.Click
    Dim arPerSel() As String        ' Array of periods that need to be taken up in this set
    Dim strPer As String = ""       ' Selection string for periods
    Dim strLine As String = ""      ' One line for the trainingset or for the testset
    Dim strDelim As String = ","    ' The delimiter we are using
    Dim strFile As String           ' Output file name
    Dim strCat As String = ""       ' Category output
    Dim ndxRes As XmlNode           ' Result node
    Dim intPtcTrain As Integer      ' Percentage training
    Dim intPtcTest As Integer       ' Percentage testset
    Dim intI As Integer             ' Counter
    Dim intJ As Integer             ' Counter
    Dim intPtc As Integer           ' Percentage
    Dim bOkay As Boolean = True     ' Whether all is fine
    Dim colTrain As New StringColl  ' Trainingset
    Dim colTest As New StringColl   ' Testset
    Dim colThis As New StringColl   ' Collection of results (in tab-delimited text file)
    Dim arColumn(0) As String       ' Array of header columns
    Dim intNum As Integer           ' Number of results
    Dim strAttrList As String = ""

    Try
      ' Validate
      If (pdxCrpDbase Is Nothing) Then Status("No corpus results are available") : Exit Sub
      ' Get first result
      ndxRes = pdxCrpDbase.SelectSingleNode("./descendant::Result")
      If (ndxRes Is Nothing) Then Status("No corpus results are available") : Exit Sub
      ' Get number of results
      intNum = 1 + ndxRes.SelectNodes("./following-sibling::Result").Count
      ' Get the parameters
      With frmTimbl
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.Cancel
            ' Exit gracefully
            Exit Sub
        End Select
        ' retrieve the parameters
        intPtcTest = .PtcTest
        intPtcTrain = .PtcTrain
        arPerSel = .Periods
        ' Validate
        If (arPerSel.Length = 0) Then Status("You need to select one or more periods") : Exit Sub
      End With
      ' Think of a good file name
      strFile = IO.Path.GetDirectoryName(strResultDb) & "\" & IO.Path.GetFileNameWithoutExtension(strResultDb) & "_timbl"
      ' Initialise the files
      IO.File.WriteAllText(strFile & ".train", "", System.Text.Encoding.UTF8)
      IO.File.WriteAllText(strFile & ".test", "", System.Text.Encoding.UTF8)
      ' Walk through the lines
      intI = 0
      While (ndxRes IsNot Nothing)
        ' Show where we are
        intPtc = (intI + 1) * 100 \ intNum
        Status("preparing timbl " & intPtc & "%", intPtc)
        If (bInterrupt) Then Exit Sub
        ' Check if this is the correct period
        strPer = GetAttrValue(ndxRes, "Period")
        If (strPer <> "") AndAlso (Array.Exists(arPerSel, Function(strThis As String) strThis = strPer)) Then
          ' Add Period, TextId, Search, forest, eTree, Cat for this row
          If (Not GetResultsDbaseFeatures(strAttrList, arColumn, "Value", ndxRes)) Then
            ' Something went wrong
            Status("frmMain/CorpusPrepareTimbl problem") : Exit Sub
          End If
          ' Fill in blanks
          For intJ = 0 To arColumn.Length - 1
            If (arColumn(intJ) = "") Then arColumn(intJ) = "-"
          Next intJ
          ' Get category
          strCat = GetAttrValue(ndxRes, "Cat") : If (strCat = "") Then strCat = "-"
          ' Combine the line and add the "Cat" as output
          strLine = Join(arColumn, strDelim) & strDelim & strCat
          ' Select whether we add this line to the training or to the testset
          If (Rnd() * 100 <= intPtcTrain) Then
            ' Add it to the trainingset
            colTrain.Add(strLine)
            If (colTrain.Count > 512) Then colTrain.Flush(strFile & ".train") : colTrain.Clear()
          Else
            ' Add it to the testset
            colTest.Add(strLine)
            If (colTest.Count > 512) Then colTest.Flush(strFile & ".test") : colTest.Clear()
          End If
        End If
        ' Next line
        ndxRes = ndxRes.NextSibling : intI += 1
      End While
      ' Finish files
      colTrain.Flush(strFile & ".train") : colTrain.Clear()
      colTest.Flush(strFile & ".test") : colTest.Clear()

      ' Prepare message for user
      colThis.Add("<html><body><h3>Timbl preparation</h3><table>")
      colThis.Add("<tr><td>training set:</td><td>" & strFile & ".train" & "</td></tr>")
      colThis.Add("<tr><td>test set:</td><td>" & strFile & ".test" & "</td></tr>")
      colThis.Add("<tr><td>delimiter:</td><td>" & IIf(strDelim = vbTab, "TAB", strDelim) & "</td></tr>")
      colThis.Add("</table>")
      ' Finish what we want to show
      colThis.Add("</body></html>")
      ' Put the information for the user in the correct tabpage
      Me.wbReport.DocumentText = colThis.Text
      Me.TabControl1.SelectedTab = Me.tpReport
      ' Show what we have done
      Status("Saved results in " & strFile)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/CorpusPrepareTIMBL error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuCorpusClefts_Click
  ' Goal:   Analyze the CLEFT corpus that has been loaded
  ' History:
  ' 04-01-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuCorpusClefts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusClefts.Click
    Dim colThis As New StringColl ' HTML result of the analysis
    Dim strFile As String         ' File name of results
    Dim strText As String         ' Working string
    Dim arPeriods() As String     ' All the periods in which there are results
    Dim strPeriods As String      ' One HTML row for the periods
    Dim arPerVal() As Integer     ' Number of occurrances of a particular feature value 
    Dim arFeatVal() As String     ' The different feature values
    Dim arFeatName() As String    ' The different features
    Dim arAnalysis() As String    ' The names of the features selected for analysis
    Dim arCleftISstate() As String = {"TopCom", "ComTop", "ComCom", "TopTop", "Emph", "Wh", "Other"}
    Dim arCleftIStype() As String = {"TopCom", "ComTop", "ComCom", "TopTop", "Wh"}
    Dim arCleftCat() As String = {"Subject", "Object", "NonArgNP", "PParg", "Adjunct", "Other"}
    Dim arCleftCoref() As String = {"New", "Inferred", "Identity", "Temp", "Assumed"}
    Dim arCleftFiles() As String = {"Texts"}
    Dim intI As Integer           ' Counter
    Dim bOkay As Boolean = True   ' Whether all is fine
    Dim bFlag As Boolean          ' Remember the dirty condition

    Try
      ' Check if there are any results
      If (tdlResults Is Nothing) Then
        bOkay = False
      ElseIf (tdlResults.Tables("Result") Is Nothing) Then
        bOkay = False
      ElseIf (tdlResults.Tables("Result").Rows.Count = 0) Then
        bOkay = False
      End If
      If (Not bOkay) Then
        Status("No corpus results are available")
        Exit Sub
      End If
      ' Initialisations
      colThis.Clear() : bFlag = bResDirty
      ' Start producing corpus results in HTML form
      colThis.Add("<html><head>Corpus Database Analysis</head><body><p>")
      ' Get all the periods
      arPeriods = GetResFieldValues(tdlResults.Tables("Result").Select(""), "Period")
      ReDim arPerVal(0 To UBound(arPeriods))
      ' Make columns for the periods
      strPeriods = "<tr><td>Value</td>"
      For intI = 0 To UBound(arPeriods)
        strPeriods &= "<td>" & arPeriods(intI) & "</td>"
      Next intI
      strPeriods &= "</tr>"
      ' Get the names of all the different features
      arFeatName = GetResFieldValues(tdlResults.Tables("Feature").Select(""), "Name", True)
      ' Get the names of the features selected for analysis
      arAnalysis = Split(Me.tbDbAnalysis.Text, ";")
      ' Check the result
      If (arAnalysis.Length = 0) OrElse (arAnalysis.Length = 1 AndAlso arAnalysis(0) = "") Then
        ' Copy all the names
        arAnalysis = arFeatName.Clone
        ' Set the value
        Me.tbDbAnalysis.Text = Join(arAnalysis, ";")
      End If
      ' Go through all the different features
      For intI = 0 To UBound(arAnalysis)
        ' Double check if there is something at all
        If (arAnalysis(intI) = "") Then Exit For
        ' Get the array of possible feature values 
        arFeatVal = GetResFieldValues(tdlResults.Tables("Feature").Select("Name='" & arAnalysis(intI) & "'"), "Value", True)
        ' Process this element from [arAnalysis]
        If (Not OneCorpusCleftTable(colThis, arPerVal, arFeatVal, arPeriods, strPeriods, arAnalysis(intI))) Then
          ' Show there is something wrong
          MsgBox("CorpusClefts error. The following table could not be processed: " & arAnalysis(intI))
          ' Continue with the next table!
        End If
      Next intI         ' Features
      ' Make a separate table for the IS-Status of clefts
      If (Not OneCorpusCleftTable(colThis, arPerVal, arCleftISstate, arPeriods, strPeriods, "CleftISstate")) Then
        ' Show there is something wrong
        MsgBox("CorpusClefts error. Could not process the [CleftISstate] table.")
      End If
      ' Make a separate table for the IS-Type of clefts
      If (Not OneCorpusCleftTable(colThis, arPerVal, arCleftIStype, arPeriods, strPeriods, "CleftIStype")) Then
        ' Show there is something wrong
        MsgBox("CorpusClefts error. Could not process the [CleftIStype] table.")
      End If
      ' Make a separate table for the corrected "CleftedCat" type
      If (Not OneCorpusCleftTable(colThis, arPerVal, arCleftCat, arPeriods, strPeriods, "CleftedCatNoTemp")) Then
        ' Show there is something wrong
        MsgBox("CorpusClefts error. Could not process the [CleftedCatNoTemp] table.")
      End If
      ' Make a separate table for the corrected "CleftedCoref" type
      If (Not OneCorpusCleftTable(colThis, arPerVal, arCleftCoref, arPeriods, strPeriods, "CleftedCorefNoTemp")) Then
        ' Show there is something wrong
        MsgBox("CorpusClefts error. Could not process the [CleftedCorefNoTemp] table.")
      End If
      ' Make a separate table with the number of files in which a cleft is found per period
      If (Not OneCorpusCleftTable(colThis, arPerVal, arCleftFiles, arPeriods, strPeriods, "Files")) Then
        ' Show there is something wrong
        MsgBox("CorpusClefts error. Could not process the [Files] table.")
      End If
      ' Finish the HTML
      colThis.Add("</body></html>")
      ' Gather the results
      strText = colThis.Text
      ' Put the results in the correct tabpage
      Me.wbReport.DocumentText = strText
      Me.TabControl1.SelectedTab = Me.tpReport
      ' Think of a good HTML file name
      strFile = IO.Path.GetDirectoryName(strResultDb) & "\" & IO.Path.GetFileNameWithoutExtension(strResultDb) & _
        "-Analysis.htm"
      ' Put results into this HTML file in UTF8
      IO.File.WriteAllText(strFile, strText, System.Text.Encoding.UTF8)
      ' Reset the dirty condition
      bResDirty = bFlag
      ' Show what we have done
      Status("Saved analysis in " & strFile)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/CorpusClefts error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuCorpusStatus_Click
  ' Goal:   Set the focus on the status combobox of the database tab
  ' History:
  ' 12-05-2011  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuCorpusStatus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusStatus.Click
    Try
      ' Go to the correct tabpage
      Me.TabControl1.SelectedTab = Me.tpDbase
      ' Set the focus on the correct place
      Me.cboResStatus.Focus()
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/splitAddTop error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuCorpusSaveAs_Click
  ' Goal:   Save the corpus database as something else
  ' History:
  ' 23-05-2011  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuCorpusSaveAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusSaveAs.Click
    Dim strFile As String = strResultDb ' Where to save to

    Try
      ' Go to the correct tabpage
      Me.TabControl1.SelectedTab = Me.tpDbase
      ' Elicit the resulting DB name
      With Me.SaveFileDialog1
        .FileName = strFile
        .Filter = "Corpus Result database (*.xml)|*.xml"
        .InitialDirectory = IO.Path.GetDirectoryName(strFile)
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK, Windows.Forms.DialogResult.Yes
            ' Get the file name
            strFile = .FileName
          Case Else
            Exit Sub
        End Select
      End With
      ' Perform any actions before changes
      dgvDbResult_BeforeSelChanged(intCurrentResId)
      ' Try saving
      TrySaveResDb(strFile)
      ' Change the global value of the result dbase
      strResultDb = strFile
      ' Adapt last saved filename
      strCrpLast = strResultDb
      ' Note my own file's name
      Me.tbDbFile.Text = strFile
      ' Reset the dirty flag
      SetResDirty(False)
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/CorpusSaveAs error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuCorpusResultAdd_Click
  ' Goal:   Add a new line to the corpus database (if loaded)
  ' History:
  ' 08-09-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuCorpusResultAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusResultAdd.Click
    Try
      ' Validate: are we on the right tab page?
      If (Me.TabControl1.SelectedTab.Name <> Me.tpDbase.Name) Then Status("First go to the corpus tab page") : Exit Sub
      ' Validate: is a corpus database loaded?
      If (tdlResults Is Nothing) OrElse (tdlResults.Tables("Result") Is Nothing) Then
        ' Warn user
        Status("First load a corpus database (e.g. made with CorpusStudio)")
        Exit Sub
      End If
      ' Try to add a new row
      TryAddNewResult(objResEd)
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/CorpusResultAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuCorpusResultDelete_Click
  ' Goal:   Delete the currently selected line from the corpus database (if loaded)
  ' History:
  ' 08-09-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuCorpusResultDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusResultDelete.Click
    Try
      ' Validate: are we on the right tab page?
      If (Me.TabControl1.SelectedTab.Name <> Me.tpDbase.Name) Then Status("First go to the corpus tab page") : Exit Sub
      ' Validate: is a corpus database loaded?
      If (tdlResults Is Nothing) OrElse (tdlResults.Tables("Result") Is Nothing) Then
        ' Warn user
        Status("First load a corpus database (e.g. made with CorpusStudio)")
        Exit Sub
      End If
      ' Validate: is one row selected in the corpus?
      If (objResEd.SelectedId >= 0) Then
        ' Indicate that we are busy
        Status("Removing -- please wait...")
        ' Try to remove this row
        TryRemove(objResEd)
        ' Show that we are done
        Status("Done")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/CorpusResultDelete error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuCorpusFilter_Click
  ' Goal:   Filter the corpus result database using Xpath expression
  ' History:
  ' 22-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuCorpusFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusFilter.Click
    Dim ndxList As XmlNodeList        ' Result of [Select]
    Dim ndxCrp As XmlNodeList         ' List from CrpDbase
    Dim strXpathFilter As String = "" ' Xpath filter expression
    Dim intI As Integer               ' Counter
    Dim intJ As Integer               ' Counter
    Dim intPtc As Integer             ' Percentage

    Try
      ' Validate: are we on the right tab page?
      If (Me.TabControl1.SelectedTab.Name <> Me.tpDbase.Name) Then Status("First go to the corpus tab page") : Exit Sub
      ' Validate: is a corpus database loaded?
      If (tdlResults Is Nothing) OrElse (tdlResults.Tables("Result") Is Nothing) Then
        ' Warn user
        Status("First load a corpus database (e.g. made with CorpusStudio)")
        Exit Sub
      End If
      ' Access the filter form
      With frmDbFilter
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.Cancel
            Exit Sub
        End Select
        ' Retrieve the filter expression
        strXpathFilter = .XpathFilter
      End With
      ' Make sure we are not interrupted
      dgvDbResult.DataSource = Nothing : ndxCrp = Nothing
      ' (1) reset @Select for all //Result
      Status("Preparing clearing...")
      Try
        ndxCrp = pdxCrpDbase.SelectNodes(strXpathFilter, conTb)
      Catch ex As Exception
        ' The syntax of the filter is not correct...
        Status("The syntax of the filter is not correct:" & vbCrLf & ex.Message)
        ' Allow interrupt again
        If (Not InitResultsEditor()) Then Status("Could not re-initialize")
        ' We need to leave anyway
        Exit Sub
      End Try
      ' Set the value of the global filter
      strDbaseXpathFilter = strXpathFilter
      ' Select the whole list for the database
      ndxList = pdxResults.SelectNodes("//Result") : intJ = 0
      For intI = 0 To ndxList.Count - 1
        ' Show where we are 
        intPtc = (intI + 1) * 100 \ ndxList.Count
        Status("Adapt filter " & intPtc & "%", intPtc)
        ' Adapt feature
        If (intJ < ndxCrp.Count) AndAlso (ndxCrp(intJ).Attributes("ResId").Value = ndxList(intI).Attributes("ResId").Value) Then
          ' Set the selection in the dgv-owned [pdxResults]
          ' ndxList(intI).Attributes("Hidden").Value = "1"
          AddXmlAttribute(pdxResults, ndxList(intI), "Hidden", "1")
          ' Advance the J counter
          intJ += 1
        Else
          ' Clear this selection
          ' ndxList(intI).Attributes("Hidden").Value = "0"
          AddXmlAttribute(pdxResults, ndxList(intI), "Hidden", "0")
        End If
        'If (ndxCrp(intI).SelectSingleNode(strXpathFilter, conTb) Is Nothing) Then
        'Else
        'End If
      Next intI
      ' Allow interrupt again
      If (Not InitResultsEditor()) Then Status("Could not re-initialize") : Exit Sub
      ' (3) Filter on the selected ones
      objResEd.Filter = "Hidden = '1'"
      ' (4) Show how many are selected
      Status("Results selected: " & Me.dgvDbResult.RowCount)
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/CorpusFilter error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuCorpusFindRepl_Click
  ' Goal:   Find and replace annotation...
  ' History:
  ' 02-09-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuCorpusFindRepl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusFindRepl.Click
    Try
      ' Validate: are we on the right tab page?
      If (Me.TabControl1.SelectedTab.Name <> Me.tpDbase.Name) Then Status("First go to the corpus tab page") : Exit Sub
      ' Give a warning
      Select Case MsgBox("WARNING: the find-and-replace function is not 100% checked yet with the new engine in place." & vbCrLf & _
                         "Make sure you have a backup of the database before continuing.", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.No, MsgBoxResult.Cancel
          Status("Okay, smart move...")
          Exit Sub
      End Select
      ' Validate: is a corpus database loaded?
      If (tdlResults Is Nothing) OrElse (tdlResults.Tables("Result") Is Nothing) Then
        ' Warn user
        Status("First load a corpus database (e.g. made with CorpusStudio)")
        Exit Sub
      End If
      ' Check if a filter has been set
      If (objResEd.Filter <> "") Then
        Select Case MsgBox("You currently have a filter set to [" & objResEd.Filter & "]" & vbCrLf & _
                           "The filter function is incompatible with Find/Replace." & vbCrLf & _
                           "Would you like to switch off filtering and continue (Y) or not (N)?", MsgBoxStyle.YesNoCancel)
          Case MsgBoxResult.No, MsgBoxResult.Cancel
            Exit Sub
        End Select
        ' Switch off filtering
        objResEd.Filter = ""
        If (objResEd.SelectedId < 0) Then objResEd.SelectDgvId(1)
      End If
      ' Validate: is one row selected in the corpus?
      If (objResEd.SelectedId >= 0) Then
        ' Ask the user for copying details
        With frmDbReplace
          ' Fill with the values from the currently selected id
          .Fill(objResEd.SelectedId)
          ' See what showing givus us
          Select Case .ShowDialog
            Case Windows.Forms.DialogResult.OK
              ' Copy the results step by step, asking for confirmation in each step
              If (Not .CopyResults(bResDirty, True)) Then Status("Error while copying") : Exit Sub
            Case Windows.Forms.DialogResult.Ignore  ' This means: Do everything in one go!
              ' Ask for confirmation
              Select Case MsgBox("Changes cannot be undone! " & vbCrLf & _
                                 "You are advised to save the database before proceding" & vbCrLf & _
                                 "Are you sure you want to proceed?", MsgBoxStyle.YesNoCancel)
                Case MsgBoxResult.Yes
                  ' Indicate that we are busy
                  Status("Copying -- please wait...")
                  ' Copy the results
                  If (Not .CopyResults(bResDirty, False)) Then Status("Error while copying") : Exit Sub
                Case Else
                  ' Opt out
                  Exit Sub
              End Select
            Case Windows.Forms.DialogResult.Cancel
              ' Opt out!
              Exit Sub
          End Select
        End With
        ' Show that we are done
        Status("Done")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/CorpusFindRepl error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuCorpusFeaturesToTexts_Click
  ' Goal:   Copy the annotation from one record to similar records
  ' History:
  ' 02-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuCorpusFeaturesToTexts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusExpFeatTexts.Click
    Dim strCurrentFilter As String = "" ' Filter to be used
    Dim strFeatCat As String = ""       ' Feature category
    Dim strFeatName As String = ""      ' Feature name
    Dim strFeatValues As String = ""    ' List of allowable feature values
    Dim strFeatBlackV As String = ""    ' Blacklist
    Dim strFeatDbase As String = ""     ' Name of the feature from the database 
    Dim strTextDir As String = ""       ' Directory
    Dim bUseAttribute As Boolean = False  ' Change attribute or not?

    Try
      ' Validate
      If (tdlResults Is Nothing) OrElse (tdlResults.Tables("Result") Is Nothing) Then
        ' Warn user
        Status("First load a corpus database (e.g. made with CorpusStudio)")
        Exit Sub
      End If
      ' Ask user if he wants to apply a filter to the data that is going to be added
      Select Case MsgBox("NOTE: make sure you have a backup of your corpus text files!!" & vbCrLf & vbCrLf & _
                         "If you only need features to be added to texts from PART of your data, " & _
                         "you should set a filter first (Corpus/FilterFeatures)" & vbCrLf & _
                         "Would you like to stop and set/change the filter first?", vbYesNoCancel)
        Case MsgBoxResult.No
          ' We can continue!
        Case MsgBoxResult.Cancel, MsgBoxResult.Yes
          ' stop
          Status("Action aborted by user")
      End Select
      ' Get any currently selected filter
      strCurrentFilter = strDbaseXpathFilter : bUseAttribute = False
      ' Ask user for the necessary information
      With frmFeatToText
        ' Set the directory for the texts - take it from the currently loaded corpus
        .TextDir = Me.tbDbSrcDir.Text
        ' Pass on the filter for the user to be seen
        .Filter = strCurrentFilter
        ' Make sure database features are set
        If (Not .SetDbaseFeatures) Then Status("Could not set database features") : Exit Sub
        ' Ask user for information
        Select Case .ShowDialog
          Case DialogResult.Cancel
            Status("Action aborted by user")
            Exit Sub
        End Select
        ' Okay, retrieve the information we need to have...
        strFeatCat = .FeatCat
        strFeatName = .FeatName
        strFeatValues = .FeatWhiteList
        strFeatBlackV = .FeatBlackList
        strTextDir = .TextDir
        strFeatDbase = .FeatDbase
      End With
      ' Check if user wants to rename the label
      If (strFeatDbase = strFeatName) AndAlso (DoLike(strFeatName, "Label|from|to")) Then
        Select Case MsgBox("Would you like to change the attribute @" & strFeatName & " of the nodes " & _
                           "instead of adding a feature named [" & strFeatName & _
                           "]?", MsgBoxStyle.YesNoCancel)
          Case MsgBoxResult.Cancel
            Status("Aborted") : Exit Sub
          Case MsgBoxResult.No
          Case MsgBoxResult.Yes
            bUseAttribute = True
        End Select
      End If
      ' Process all files with the information above
      If (Not DoCorpusFeaturesToTexts(strTextDir, strCurrentFilter, strFeatDbase, strFeatCat, strFeatName, _
                                      strFeatValues, strFeatBlackV, bUseAttribute)) Then Status("There was an error") : Exit Sub
      ' Show that we are done
      Status("Done")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/CorpusFeaturesToTexts error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuCorpusExpFeatFile_Click
  ' Goal:   Export features to a file
  ' History:
  ' 10-07-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuCorpusExpFeatFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusExpFeatFile.Click
    Dim strCurrentFilter As String = "" ' Filter to be used
    Dim strFeatValues As String = ""    ' List of allowable feature values
    Dim strFeatBlackV As String = ""    ' Blacklist
    Dim strFeatDbase As String = ""     ' Name of the feature from the database 
    Dim strExpDir As String = ""        ' Directory
    Dim bUseStatus As Boolean = False   ' Copy status or not
    Dim bUseNotes As Boolean = False    ' Copy notes or not
    Dim bUsePde As Boolean = False      ' Copy pde or not

    Try
      ' Validate
      If (tdlResults Is Nothing) OrElse (tdlResults.Tables("Result") Is Nothing) Then
        ' Warn user
        Status("First load a corpus database (e.g. made with CorpusStudio)")
        Exit Sub
      End If
      ' Get any currently selected filter
      strCurrentFilter = strDbaseXpathFilter
      If (strCurrentFilter = "") Then
        ' Ask user if he wants to apply a filter to the data that is going to be added
        Select Case MsgBox("If you only need features to be exported from PART of your data, " & _
                           "you should set a filter first (Corpus/FilterFeatures)" & vbCrLf & _
                           "Would you like to stop and set/change the filter first?", vbYesNoCancel)
          Case MsgBoxResult.No
            ' We can continue!
          Case MsgBoxResult.Cancel, MsgBoxResult.Yes
            ' stop
            Status("Action aborted by user")
            Exit Sub
        End Select
      End If
      ' Ask user for the necessary information
      With frmFeatTransfer
        ' Set the directory where the feature file is going to be saved
        .ExportDir = IO.Path.GetDirectoryName(Me.tbDbFile.Text)
        ' Pass on the filter for the user to be seen
        .Filter = strCurrentFilter
        ' Set mode to export
        .Mode = "exp"
        ' Make sure database features are set
        If (Not .SetDbaseFeatures) Then Status("Could not set database features") : Exit Sub
        ' Ask user for information
        Select Case .ShowDialog
          Case DialogResult.Cancel
            Status("Action aborted by user")
            Exit Sub
        End Select
        ' Okay, retrieve the information we need to have...
        strFeatValues = .FeatWhiteList
        strFeatBlackV = .FeatBlackList
        strExpDir = .ExportDir
        strFeatDbase = .FeatDbase
        bUseStatus = .UseStatus
        bUseNotes = .UseNotes
        bUsePde = .UsePde
      End With
      ' Process all files with the information above
      If (Not DoCorpusFeaturesToFile(strExpDir, strCurrentFilter, strFeatDbase, strFeatValues, _
              strFeatBlackV, bUseStatus, bUseNotes, bUsePde, strExpDir)) Then Status("There was an error") : Exit Sub
      ' Show that we are done
      Status("Done")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/CorpusExpFeatFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuCorpusAddFeature_Click
  ' Goal:   Add one feature (with default value) to the currently loaded database
  ' History:
  ' 23-08-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuCorpusAddFeature_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusAddFeature.Click
    Dim strFeatName As String = ""  ' Name of new feature
    Dim strFeatValue As String = "" ' Default value for the new feature
    Dim bOkay As Boolean = True     ' Feature name is okay

    Try
      ' Validate
      If (tdlResults Is Nothing) OrElse (tdlResults.Tables("Result") Is Nothing) Then
        ' Warn user
        Status("First load a corpus database (e.g. made with CorpusStudio)")
        Exit Sub
      End If
      ' Elicit the feature name and value
      Do
        ' Get feature name and value
        With frmFeatAdd
          Select Case .ShowDialog
            Case System.Windows.Forms.DialogResult.Cancel, Windows.Forms.DialogResult.No
              Status("Canceled")
              Exit Sub
            Case Windows.Forms.DialogResult.OK, Windows.Forms.DialogResult.Yes
              strFeatName = .FeatureName
              strFeatValue = .FeatureValue
          End Select
        End With
        ' Check if the feature with this name does not exist already in the database
        If (FeatureExists(strFeatName)) Then
          Status("A feature with the name [" & strFeatName & "] already exists. Use a different name")
          bOkay = False
        End If
      Loop Until bOkay
      ' Add the feature to the database
      If (Not FeatureAdd(strFeatName, strFeatValue)) Then Status("Could not add the feature") : Exit Sub
      ' Add the feature name to the analysis
      Me.tbDbAnalysis.Text &= ";" & strFeatName
      ' Make sure first record is being shown
      objResEd.SelectDgvId(1)
      dgvDbResult_SelectionChanged(Nothing, Nothing)
      ' Set the status to dirty
      SetResDirty(True)
      ' Show that we are done
      Status("Done")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/CorpusAddFeature error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuCorpusImpFeatFile_Click
  ' Goal:   Import features from a file
  ' History:
  ' 10-07-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuCorpusImpFeatFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusImpFeatFile.Click
    Dim strImpFtFile As String = "" ' File from which we import
    Dim strCurrentFilter As String = "" ' Filter to be used
    Dim strFeatValues As String = ""    ' List of allowable feature values
    Dim strFeatBlackV As String = ""    ' Blacklist
    Dim strFeatDbase As String = ""     ' Name of the feature from the database 
    Dim strResFilter As String = ""     ' Results filter
    Dim strMethod As String = ""        ' Method used
    Dim strType As String = ""          ' File type
    Dim intSelId As Integer = 0         ' Selected id
    Dim bUseStatus As Boolean = False   ' Copy status or not
    Dim bUseNotes As Boolean = False    ' Copy notes or not
    Dim bUsePde As Boolean = False      ' Copy pde or not

    Try
      ' Validate
      If (tdlResults Is Nothing) OrElse (tdlResults.Tables("Result") Is Nothing) Then
        ' Warn user
        Status("First load a corpus database (e.g. made with CorpusStudio)")
        Exit Sub
      End If
      ' Get any currently selected filter
      strCurrentFilter = strDbaseXpathFilter
      If (strCurrentFilter = "") Then
        ' Ask user if he wants to apply a filter to the data that is going to be added
        Select Case MsgBox("If you only need features to be imported to PART of the database data, " & _
                           "you should set a filter first (Corpus/FilterFeatures)" & vbCrLf & _
                           "Would you like to *STOP* and set/change the filter first?", vbYesNoCancel)
          Case MsgBoxResult.No
            ' We can continue!
          Case MsgBoxResult.Cancel, MsgBoxResult.Yes
            ' stop
            Status("Action aborted by user")
            Exit Sub
        End Select
      End If
      ' Ask user for the necessary information
      With frmFeatTransfer
        ' Set the import directory from where we expect a feature file to be
        .ExportDir = IO.Path.GetDirectoryName(Me.tbDbFile.Text)
        ' Pass on the filter for the user to be seen
        .Filter = strCurrentFilter
        ' Set mode to import
        .Mode = "imp"
        ' Set the method
        .Method = "ResId"
        ' Make sure database features are set
        If (Not .SetDbaseFeatures) Then Status("Could not set database features") : Exit Sub
        ' Ask user for information
        Select Case .ShowDialog
          Case DialogResult.Cancel
            Status("Action aborted by user")
            Exit Sub
        End Select
        ' Okay, retrieve the information we need to have...
        strFeatValues = .FeatWhiteList
        strFeatBlackV = .FeatBlackList
        strImpFtFile = .ImportFile
        strFeatDbase = .FeatDbase
        bUseStatus = .UseStatus
        bUseNotes = .UseNotes
        bUsePde = .UsePde
        strMethod = .Method
      End With
      If (bUseStatus) Then
        intSelId = objResEd.SelectedId
        ' Make sure we are not interrupted
        strResFilter = objResEd.Filter
        dgvDbResult.DataSource = Nothing
      End If
      ' Process all files with the information above
      strType = IO.Path.GetExtension(strImpFtFile).ToLower
      Select Case strType
        Case ".csv"
          If (Not DoCorpusFeaturesFromCsv(strImpFtFile, strCurrentFilter, strFeatDbase, strMethod, strFeatValues, _
                  strFeatBlackV, bUseStatus, bUseNotes, bUsePde)) Then Status("There was an error") : Exit Sub
        Case ".xml"
          If (Not DoCorpusFeaturesFromFile(strImpFtFile, strCurrentFilter, strFeatDbase, strMethod, strFeatValues, _
                  strFeatBlackV, bUseStatus, bUseNotes, bUsePde)) Then Status("There was an error") : Exit Sub
        Case Else
          ' Cannot handle this
          Status("Cannot handle file of type: " & strType)
      End Select
      If (bUseStatus) Then
        ' Allow interrupt again
        If (Not InitResultsEditor()) Then Status("Could not re-initialize") : Exit Sub
        ' (3) Filter on the selected ones
        objResEd.Filter = strResFilter
        objResEd.SelectDgvId(intSelId)
      End If
      SetResDirty(True)
      ' Show that we are done
      Status("Done")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/CorpusImpFeatFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuCorpusCopyAnnotation_Click
  ' Goal:   Copy the annotation from one record to similar records
  ' History:
  ' 10-12-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuCorpusCopyAnnotation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusCopyAnnotation.Click
    Try
      ' Validate: are we on the right tab page?
      If (Me.TabControl1.SelectedTab.Name <> Me.tpDbase.Name) Then Status("First go to the corpus tab page") : Exit Sub
      ' Validate: is a corpus database loaded?
      If (tdlResults Is Nothing) OrElse (tdlResults.Tables("Result") Is Nothing) Then
        ' Warn user
        Status("First load a corpus database (e.g. made with CorpusStudio)")
        Exit Sub
      End If
      ' Validate: is one row selected in the corpus?
      If (objResEd.SelectedId >= 0) Then
        ' Ask the user for copying details
        With frmCorpus
          ' Fill with the values from the currently selected id
          .Fill(objResEd.SelectedId)
          ' See what showing givus us
          Select Case .ShowDialog
            Case Windows.Forms.DialogResult.OK
              ' Copy the results step by step, asking for confirmation in each step
              If (Not .CopyResults(bResDirty, True)) Then Status("Error while copying") : Exit Sub
            Case Windows.Forms.DialogResult.Ignore  ' This means: Do everything in one go!
              ' Ask for confirmation
              Select Case MsgBox("Changes cannot be undone! " & vbCrLf & _
                                 "You are advised to save the database before proceding" & vbCrLf & _
                                 "Are you sure you want to proceed?", MsgBoxStyle.YesNoCancel)
                Case MsgBoxResult.Yes
                  ' Indicate that we are busy
                  Status("Copying -- please wait...")
                  ' Copy the results
                  If (Not .CopyResults(bResDirty, False)) Then Status("Error while copying") : Exit Sub
                Case Else
                  ' Opt out
                  Exit Sub
              End Select
            Case Windows.Forms.DialogResult.Cancel
              ' Opt out!
              Exit Sub
          End Select
        End With
        ' Show that we are done
        Status("Done")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/CorpusCopyAnnotation error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  TryRemove
  ' Goal :  Delete the selected element from the dgv in [objThis]
  ' History:
  ' 02-06-2010  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub TryRemove(ByRef objThis As DgvHandle)
    Dim strQ As String = "Would you really like to delete the selected row?"
    Dim bValue As Boolean ' Value of the dirty bit

    Try
      ' First enquire
      If (MsgBox(strQ, MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes) Then
        ' Save the value of the dirty bit
        bValue = bDirty
        ' TRy to remove the selected object
        If (Not objThis.Remove) Then
          ' Unsuccesful...
          Status("Unable to delete this element")
        Else
          ' Make sure dirty bit is set
          bResDirty = True
          ' Show status
          Status("Deleted")
        End If
        ' Reset this value
        bDirty = bValue
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/TryRemove error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  TryAddNewResult
  ' Goal :  Try add a new result element
  ' History:
  ' 02-10-2009  ERK Created from TryAddNew
  ' 27-09-2011  ERK Specialized for Result
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub TryAddNewResult(ByRef objThis As DgvHandle)
    Dim strElName As String = ""      ' Name of the new element
    Dim strText As String = ""        ' Initial text of the query
    Dim strFeatName As String = ""    ' Name of feature
    Dim bRemNodes As Boolean = False  ' Whether to remove nodes or not
    Dim bPrintIdc As Boolean = False  ' Whether to print indices or not
    Dim dtrFirst As DataRow           ' The first database item
    Dim dtrChild As DataRow           ' Child row
    Dim dtrFound() As DataRow         ' Resutl of Select
    Dim dtrNew As DataRow             ' New child to be added

    Try
      ' Validate
      If (objThis Is Nothing) Then Exit Sub
      ' Are any results available?
      If (tdlResults.Tables("Result").Rows.Count = 0) Then
        ' In this case we cannot add anything
        Status("Cannot add results from scratch")
        Exit Sub
      End If
      ' Get pointer to the first database item
      dtrFirst = tdlResults.Tables("Result").Rows(0)
      ' Open the correct form
      With frmGetResult
        ' Ask for input
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.Yes, Windows.Forms.DialogResult.OK
            ' Call function to perform the action
            If (objThis.AddNew("Search", .ResLoc, "TextId", .ResTextId, _
                               "Text", .ResText, "File", .ResFile, _
                               "Period", .ResPeriod, "Psd", .ResPsd, "forestId", .ResForestId, _
                               "eTreeId", .ResEtreeId, "Notes", "-")) Then
              ' Get the datarow associated with me
              dtrFound = tdlResults.Tables("Result").Select("ResId = " & objThis.SelectedId)
              If (dtrFound.Length > 0) Then
                ' Go through all the features associated with the first database row
                For Each dtrChild In dtrFirst.GetChildRows("Result_Feature")
                  ' Get the name of the feature
                  strFeatName = dtrChild.Item("Name")
                  ' Add such a child under me
                  dtrNew = tdlResults.Tables("Feature").NewRow
                  dtrNew.Item("Name") = strFeatName : dtrNew.Item("Value") = ""
                  dtrNew.SetParentRow(dtrFound(0))
                  tdlResults.Tables("Feature").Rows.Add(dtrNew)
                Next dtrChild
              End If
              ' Set dirty flag
              bResDirty = True
            Else
              ' Failure: leave
              Exit Sub
            End If
            ' Show other values may now be added
            Status("Now add the other values in the database window")
          Case Else
            ' Show it was canceled
            Status("Adding a new result was canceled")
        End Select
      End With
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/TryAddNew error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   OneCorpusCleftTable
  ' Goal:   Process the analysis of one type
  ' History:
  ' 05-01-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function OneCorpusCleftTable(ByRef colThis As StringColl, ByRef arPerVal() As Integer, _
      ByRef arFeatVal() As String, ByRef arPeriods() As String, ByVal strPeriods As String, _
      ByVal strName As String) As Boolean
    Dim strText As String       ' Working string
    Dim strClfCoref As String   ' Value of feature "CleftedCoref"
    Dim strClfCat As String     ' Value of feature "CleftedCat"
    Dim strClsStat As String    ' Value of feature "ClauseStatus"
    Dim strFocType As String    ' Value of feature "FocusType"
    Dim intK As Integer         ' Counter
    Dim intJ As Integer         ' Counter
    Dim intL As Integer         ' Counter
    Dim intFiles As Integer = 0 ' Number of files used
    Dim strFile As String = ""  ' One file
    Dim dtrFound() As DataRow   ' Result of SELECT
    Dim dtrParent As DataRow    ' The parent of the currently found <Feature> node

    Try
      ' Validate
      If (colThis Is Nothing) Then Return False
      If (arPerVal.Length = 0) OrElse (arFeatVal.Length = 0) OrElse (strPeriods = "") Then Return False
      ' Initialisations
      strClfCoref = "" : strClsStat = "" : strFocType = "" : strClfCat = ""
      ' Initialize table for results for this feature
      colThis.Add("<h2>" & strName & "</h2><table>" & strPeriods)
      ' Count results for all different feature values
      For intJ = 0 To UBound(arFeatVal)
        ' Start a row for this feature value
        strText = "<tr><td>" & arFeatVal(intJ) & "</td>"
        ' Reset the period counts
        For intK = 0 To UBound(arPerVal)
          arPerVal(intK) = 0
        Next intK
        ' Action depends on the type of table we are to make
        Select Case strName
          Case "Files"    ' Calculate the number of files used per period
            ' Get all Result elements
            dtrFound = tdlResults.Tables("Result").Select("", "PeriodNo ASC, TextId ASC")
            ' Go through all the features
            For intK = 0 To UBound(dtrFound)
              ' Restrict ourselves to this result
              With dtrFound(intK)
                ' Get the index of this period
                intL = GetPerVal(arPeriods, .Item("Period"))
                If (intL < 0) Then
                  ' There is a problem...
                  MsgBox("OneCorpusCleftTable: There is a problem!")
                  Return False
                End If
                ' Check if the file name has change
                If (LCase(IO.Path.GetFileNameWithoutExtension(dtrFound(intK).Item("File"))) <> strFile) Then
                  ' Adapt counter
                  arPerVal(intL) += 1
                  ' Show difference
                  Debug.Print("   [" & strFile & "] --> [" & LCase(IO.Path.GetFileNameWithoutExtension(dtrFound(intK).Item("File"))) & "]")
                  ' Adapt the file name
                  strFile = LCase(IO.Path.GetFileNameWithoutExtension(dtrFound(intK).Item("File")))
                  ' Show result
                  Debug.Print(.Item("Period") & ": " & arPerVal(intL) & " - " & strFile)
                End If
              End With
            Next intK
          Case "CleftISstate"
            ' Get all Result elements
            dtrFound = tdlResults.Tables("Result").Select("")
            ' Go through all the features
            For intK = 0 To UBound(dtrFound)
              ' Restrict ourselves to this result
              With dtrFound(intK)
                ' Get the index of this period
                intL = GetPerVal(arPeriods, .Item("Period"))
                If (intL < 0) Then
                  ' There is a problem...
                  MsgBox("OneCorpusCleftTable: There is a problem!")
                  Return False
                End If
                ' Get the different feature values linked to this particular datarow
                If (Not GetFeatureValue(dtrFound(intK), "CleftedCoref", strClfCoref)) Then Return False
                If (Not GetFeatureValue(dtrFound(intK), "ClauseStatus", strClsStat)) Then Return False
                If (Not GetFeatureValue(dtrFound(intK), "FocusType", strFocType)) Then Return False
                ' Check if this result fulfills the conditions
                Select Case arFeatVal(intJ)
                  Case "TopCom" ' Topic-Comment:   CleftedCoref=Ref, ClauseStat=New, FocusType=none|Time|Reason|Purpose
                    ' Check the focustype
                    If (DoLike(strFocType, "none|Time|Reason|Purpose")) Then
                      If (strClfCoref <> "New") AndAlso (strClsStat = "New") Then arPerVal(intL) += 1
                    End If
                  Case "ComTop" ' Comment-Topic:   CleftedCoref=New, ClauseStat=Ref, FocusType=none
                    ' Check the focustype
                    If (strFocType = "none") Then
                      If (strClfCoref = "New") AndAlso (strClsStat <> "New") Then arPerVal(intL) += 1
                    End If
                  Case "ComCom" ' Comment-Comment: CleftedCoref=New, ClauseStat=New, FocusType=none
                    ' Check the focustype
                    If (strFocType = "none") Then
                      If (strClfCoref = "New") AndAlso (strClsStat = "New") Then arPerVal(intL) += 1
                    End If
                  Case "TopTop" ' Topic-Topic: CleftedCoref=Ref, ClauseStat=Ref, FocusType=none
                    ' Check the focustype
                    If (strFocType = "none") Then
                      If (strClfCoref <> "New") AndAlso (strClsStat <> "New") Then arPerVal(intL) += 1
                    End If
                  Case "Wh"
                    ' Check the focustype
                    If (strFocType = "Wh") Then arPerVal(intL) += 1
                  Case "Emph"   ' Emphatic:        FocusType=Contrast|Emph
                    ' Check the focustype
                    If (DoLike(strFocType, "Contrast|Contrast;*|Emph")) Then arPerVal(intL) += 1
                  Case "Other"  ' Other type: cleftedCoref<>New, ClauseStat<>New
                    If (DoLike(strFocType, "Time|Reason|Purpose")) Then
                      If (strClfCoref = "New") OrElse (strClsStat <> "New") Then arPerVal(intL) += 1
                    End If
                End Select
              End With
            Next intK
          Case "CleftIStype"
            ' Get all Result elements
            dtrFound = tdlResults.Tables("Result").Select("")
            ' Go through all the features
            For intK = 0 To UBound(dtrFound)
              ' Restrict ourselves to this result
              With dtrFound(intK)
                ' Get the index of this period
                intL = GetPerVal(arPeriods, .Item("Period"))
                If (intL < 0) Then
                  ' There is a problem...
                  MsgBox("OneCorpusCleftTable: There is a problem!")
                  Return False
                End If
                ' Get the different feature values linked to this particular datarow
                If (Not GetFeatureValue(dtrFound(intK), "CleftedCoref", strClfCoref)) Then Return False
                If (Not GetFeatureValue(dtrFound(intK), "ClauseStatus", strClsStat)) Then Return False
                If (Not GetFeatureValue(dtrFound(intK), "FocusType", strFocType)) Then Return False
                ' Check if this result fulfills the conditions
                Select Case arFeatVal(intJ)
                  Case "TopCom" ' Topic-Comment:   CleftedCoref=Ref, ClauseStat=New, FocusType=none|Time|Reason|Purpose
                    ' Check the focustype
                    If (strFocType <> "Wh") Then
                      If (strClfCoref <> "New") AndAlso (strClsStat = "New") Then arPerVal(intL) += 1
                    End If
                  Case "ComTop" ' Comment-Topic:   CleftedCoref=New, ClauseStat=Ref, FocusType=none
                    ' Check the focustype
                    If (strFocType <> "Wh") Then
                      If (strClfCoref = "New") AndAlso (strClsStat <> "New") Then arPerVal(intL) += 1
                    End If
                  Case "ComCom" ' Comment-Comment: CleftedCoref=New, ClauseStat=New, FocusType=none
                    ' Check the focustype
                    If (strFocType <> "Wh") Then
                      If (strClfCoref = "New") AndAlso (strClsStat = "New") Then arPerVal(intL) += 1
                    End If
                  Case "TopTop" ' Topic-Topic: CleftedCoref=Ref, ClauseStat=Ref, FocusType=none
                    ' Check the focustype
                    If (strFocType <> "Wh") Then
                      If (strClfCoref <> "New") AndAlso (strClsStat <> "New") Then arPerVal(intL) += 1
                    End If
                  Case "Wh"
                    ' Check the focustype
                    If (strFocType = "Wh") Then arPerVal(intL) += 1
                End Select
              End With
            Next intK
          Case "CleftedCatNoTemp"
            ' Get all Result elements
            dtrFound = tdlResults.Tables("Result").Select("")
            ' Go through all the features
            For intK = 0 To UBound(dtrFound)
              ' Restrict ourselves to this result
              With dtrFound(intK)
                ' Get the index of this period
                intL = GetPerVal(arPeriods, .Item("Period"))
                If (intL < 0) Then
                  ' There is a problem...
                  MsgBox("OneCorpusCleftTable: There is a problem!")
                  Return False
                End If
                ' Get the different feature values
                If (Not GetFeatureValue(dtrFound(intK), "CleftedCat", strClfCat)) Then Return False
                If (Not GetFeatureValue(dtrFound(intK), "FocusType", strFocType)) Then Return False
                ' Exclude those with "Time" focustype
                If (strFocType <> "Time") Then
                  ' Check if this is the correct one
                  If (arFeatVal(intJ) = strClfCat) Then
                    ' Increment the counter for this one
                    arPerVal(intL) += 1
                  End If
                ElseIf (strFocType = "Time") AndAlso (arFeatVal(intJ) = "Temp") Then
                  arPerVal(intL) += 1
                End If
              End With
            Next intK
          Case "CleftedCorefNoTemp"
            ' Get all Result elements
            dtrFound = tdlResults.Tables("Result").Select("")
            ' Go through all the features
            For intK = 0 To UBound(dtrFound)
              ' Restrict ourselves to this result
              With dtrFound(intK)
                ' Get the index of this period
                intL = GetPerVal(arPeriods, .Item("Period"))
                If (intL < 0) Then
                  ' There is a problem...
                  MsgBox("OneCorpusCleftTable: There is a problem!")
                  Return False
                End If
                ' Get the different feature values
                If (Not GetFeatureValue(dtrFound(intK), "CleftedCoref", strClfCoref)) Then Return False
                If (Not GetFeatureValue(dtrFound(intK), "FocusType", strFocType)) Then Return False
                ' Exclude those with "Time" focustype
                If (strFocType <> "Time") Then
                  ' Check if this is the correct one
                  If (arFeatVal(intJ) = strClfCoref) Then
                    ' Increment the counter for this one
                    arPerVal(intL) += 1
                  End If
                ElseIf (strFocType = "Time") AndAlso (arFeatVal(intJ) = "Temp") Then
                  arPerVal(intL) += 1
                End If
              End With
            Next intK
          Case Else
            ' Select the features with this particular value
            dtrFound = tdlResults.Tables("Feature").Select("Name='" & strName & _
                                "' AND Value='" & arFeatVal(intJ).Replace("'", "''") & "'")
            ' Go through all the features with this name and value
            For intK = 0 To UBound(dtrFound)
              ' Get my parent
              dtrParent = dtrFound(intK).GetParentRow("Result_Feature")
              ' Add the result to the appropriate period
              intL = GetPerVal(arPeriods, dtrParent.Item("Period").ToString)
              If (intL >= 0) Then
                ' Increment the appropriate counter
                arPerVal(intL) += 1
              Else
                ' There is a problem...
                MsgBox("OneCorpusCleftTable: There is a problem!")
                Return False
              End If
            Next intK
        End Select
        ' Go through the findings
        For intK = 0 To UBound(arPerVal)
          strText &= "<td>" & arPerVal(intK).ToString & "</td>"
        Next intK
        ' Finish the row for this feature value
        strText &= "</tr>"
        ' Store the row in the string collection
        colThis.Add(strText)
      Next intJ
      ' Finish this table
      colThis.Add("</table><p>")
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/OneCorpusCleftTable error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetFeatureValue
  ' Goal:   Get the feature [value] for the indicated [strName]
  ' History:
  ' 05-01-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetFeatureValue(ByRef dtrThis As DataRow, ByVal strName As String, ByRef strValue As String) As Boolean
    Dim dtrChildren() As DataRow  ' All my children
    Dim intI As Integer           ' Counter

    Try
      ' Validate
      If (dtrThis Is Nothing) Then Return False
      If (strName = "") Then Return False
      ' Get all my Feature children
      dtrChildren = dtrThis.GetChildRows("Result_Feature")
      For intI = 0 To UBound(dtrChildren)
        ' Is this the child?
        If (dtrChildren(intI).Item("Name") = strName) Then
          ' Return the value
          strValue = dtrChildren(intI).Item("Value")
          Return True
        End If
      Next intI
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/OneCorpusCleftTable error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuCorpusFix_Click
  ' Goal:   Apply a fix to the CLEFT corpus that has been loaded
  ' History:
  ' 04-01-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuCorpusFix_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusFix.Click
    Dim strText As String         ' Working string
    Dim arText() As String        ' Feature value split up
    Dim dtrThis As DataRow        ' One datarow
    Dim dtrFeat As DataRow        ' One feature datarow
    Dim dtrChildren() As DataRow  ' Feature children of <Result>
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim bOkay As Boolean = True   ' Whether all is fine
    Dim bHasFocType As Boolean    ' Whether the feature FocusType is there or not

    Try
      ' Check if there are any results
      If (tdlResults Is Nothing) Then
        bOkay = False
      ElseIf (tdlResults.Tables("Result") Is Nothing) Then
        bOkay = False
      ElseIf (tdlResults.Tables("Result").Rows.Count = 0) Then
        bOkay = False
      End If
      If (Not bOkay) Then
        Status("No corpus results are available")
        Exit Sub
      End If
      ' Go through all the lines
      With tdlResults.Tables("Result")
        For intI = 0 To .Rows.Count - 1
          ' Get this datarow
          dtrThis = .Rows(intI)
          ' Initialise for this row
          bHasFocType = False : ReDim arText(0)
          ' Walk through the features for this row
          dtrChildren = .Rows(intI).GetChildRows("Result_Feature")
          For intJ = 0 To dtrChildren.Length - 1
            ' Is this the correct child?
            If (dtrChildren(intJ).Item("Name") = "CleftedCoref") Then
              ' Get the value of feature "CleftedCoref"
              strText = dtrChildren(intJ).Item("Value")
              arText = Split(strText, ";")
              ' Replace "CleftedCoref"
              dtrChildren(intJ).Item("Value") = arText(0)
            ElseIf (dtrChildren(intJ).Item("Name") = "FocusType") Then
              ' Add the value of the focus type here
              dtrChildren(intJ).Item("Value") = arText(1)
              bHasFocType = True
            End If
          Next intJ
          ' Check focus type
          If (Not bHasFocType) Then
            ' Make a new <Feature> child
            dtrFeat = tdlResults.Tables("Feature").NewRow
            With dtrFeat
              .Item("Name") = "FocusType"
              If (arText.Length > 1) Then
                .Item("Value") = arText(1)
              Else
                .Item("Value") = "none"
              End If
              .SetParentRow(dtrThis)
            End With
            tdlResults.Tables("Feature").Rows.Add(dtrFeat)
          End If
        Next intI
      End With
      ' Make dirty
      SetResDirty(True)
      ' Show what we have done
      Status("Fix has been applied.")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/CorpusReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuRefEdit_Click
  ' Goal:   Set or clear ref-editing mode
  ' History:
  ' 01-03-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuRefEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRefEdit.Click
    Try
      SetRefEditMode(Me.mnuRefEdit.Checked)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/RefEdit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub SetRefEditMode(ByVal bSet As Boolean)
    Try
      ' Adapt the ref-editing mode
      bRefEditMode = bSet
      ' Make sure that what is shown reflects [bRefEditMode]
      If (Me.mnuRefEdit.Checked <> bRefEditMode) Then
        Me.mnuRefEdit.Checked = bRefEditMode
      End If
      '' Show what?
      'If (bRefEditMode) Then
      '  Me.tbRefEdit.Visible = True
      'Else
      '  Me.tbRefEdit.Visible = False
      'End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/SetRefEditMode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   tbRefEdit_Click
  ' Goal:   Switch off ref-editing mode
  ' History:
  ' 03-04-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub tbRefEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbRefEdit.Click
    ' Switch off editing
    SetRefEditMode(False)
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuRefAdd_Click
  ' Goal:   Open the reference addition menu
  ' History:
  ' 02-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuRefAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRefAdd.Click
    ' Validate
    If (Not bRefEditMode) Then Status("First enter reference-editing mode") : Exit Sub
    ' Show what we are doing
    Status("Open reference addition menu")
    ' Show this form
    With frmRef
      ' Show it with dialog
      Select Case .ShowDialog
        Case Windows.Forms.DialogResult.Cancel
          ' Do nothing
          Exit Sub
        Case Windows.Forms.DialogResult.OK
          ' Add the coreferences
          If (AddCoref(.RefType)) Then
            ' Indicate dirty
            SetDirty(True)
            ' Clear all coref information
            ClearCoref(Me.tbEdtMain)
          End If
      End Select
    End With
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuRefAddId_Click
  ' Goal:   Try to add an "Identity" reference to the selected set of stuff
  ' History:
  ' 02-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuRefAddId_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRefAddId.Click
    ' Validate
    If (Not bRefEditMode) Then Status("First enter reference-editing mode") : Exit Sub
    ' Add the corefernce stuff
    If (AddCoref(strRefIdentity)) Then
      ' Indicate dirty
      SetDirty(True)
      ' Clear all coref information
      ClearCoref(Me.tbEdtMain)
    End If
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuRefAddWorld_Click
  ' Goal:   Try to add an "Assumed" reference to the selected set of stuff
  ' History:
  ' 07-09-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuRefAddWorld_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRefAddWorld.Click
    ' Validate
    If (Not bRefEditMode) Then Status("First enter reference-editing mode") : Exit Sub
    ' Add the corefernce stuff
    If (AddCoref(strRefAssumed)) Then
      ' Indicate dirty
      SetDirty(True)
      ' Clear all coref information
      ClearCoref(Me.tbEdtMain)
    End If
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuRefDel_Click
  ' Goal:   Delete the currently selected coreference (if any)
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuRefDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRefDel.Click
    Try
      ' Validate
      If (Not bRefEditMode) Then Status("First enter reference-editing mode") : Exit Sub
      ' Show we are busy
      bEdtLoad = True
      ' Delete the current coreference relation
      If (DelCoref(Me.tbEdtMain, False)) Then
        ' Indicate dirty
        SetDirty(True)
      End If
      ' Show we can be interrupted again
      bEdtLoad = False
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/RefDel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuRefDelChain_Click
  ' Goal:   Delete the currently selected coreference (down on)
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuRefDelChain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRefDelChain.Click
    Try
      ' Validate
      If (Not bRefEditMode) Then Status("First enter reference-editing mode") : Exit Sub
      ' Show we are busy
      bEdtLoad = True
      ' Delete the current coreference relation and all that is related to it
      If (DelCoref(Me.tbEdtMain, True)) Then
        ' Indicate dirty
        SetDirty(True)
      End If
      ' Show we can be interrupted again
      bEdtLoad = False
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/RefDelChain error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuRefTopic_Click
  ' Goal:   Try to do topic-guessing
  ' History:
  ' 11-10-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuRefTopic_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRefTopic.Click
    Dim strDir As String = ""   ' Directory
    Dim strFile As String = ""  ' The name of the output file (default)

    Try
      ' Validate
      If (strCurrentFile = "") Then
        ' Show our worries
        Status("First open a file!")
        Exit Sub
      End If
      ' Get the directory
      strDir = IO.Path.GetDirectoryName(strCurrentFile) & "\is"
      ' Create directory if non-existent
      If (Not IO.Directory.Exists(strDir)) Then IO.Directory.CreateDirectory(strDir)
      ' Default filename is derived from current PSDx file
      strFile = strDir & "\" & IO.Path.GetFileNameWithoutExtension(strCurrentFile) & ".htm"
      ' Perform the chosen export using the general subroutine for this
      If (Not DoExport("ISoutput", strFile)) Then
        ' Don't know what to do now - probably called function already gives error message
        Exit Sub
      End If
      ' Show the produced html file in the report tab
      Me.TabControl1.SelectedTab = Me.tpReport
      Status("Opening result...")
      Me.wbReport.Navigate(strFile)
      ' Show we are there
      Status("ISoutput: " & strFile)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/RefTopic error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuRefList_Click
  ' Goal:   Give an overview of ALL coreferential chains in this document
  ' History:
  ' 12-07-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuRefList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRefList.Click
    Dim strAction As String = ""  ' Action to perform
    Dim strScope As String = ""   ' Scope of this action 
    Dim arFile() As String        ' List of appropriate files
    Dim strFile As String         ' One particular file
    Dim strDir As String = ""     ' Directory where we have to look
    Dim strPer As String = ""     ' General period
    Dim colGlobal As New StringColl ' Global reflist information
    Dim colProt As New StringColl ' Protagonist measures
    Dim colHtml As New StringColl ' HTML output collection
    Dim strOutFile As String      ' Name of the output file
    Dim strReport As String       ' Text of the report
    Dim intI As Integer           ' Counter
    Dim arPeriod() As String = {"OE", "ME", "eModE", "LmodE"}
    Dim tblFile As New DataTable  ' A table to host file & period
    Dim dtrFound() As DataRow     ' Ordered selection of files

    Try
      ' Reset interrupt status
      bInterrupt = False
      ' Find out what action should be taken
      With frmChain
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Set the parameters
            strAction = .Action
            strScope = .Scope
            strDir = .Dir
          Case Windows.Forms.DialogResult.Cancel
            ' Show we are leaving
            Status("Action aborted")
            ' Leave the program
            Exit Sub
        End Select
      End With
      ' If we come here, then we can continue
      ' Action depends on the scope
      Select Case strScope
        Case "All"
          ' Check all psdx files in the current directory
          arFile = IO.Directory.GetFiles(strDir, "*.psdx")
          ' Initialise the file/period table
          With tblFile
            .Columns.Add("File", System.Type.GetType("System.String"))
            .Columns.Add("Period", System.Type.GetType("System.String"))
          End With
          ' Check out the periods of all files
          For intI = 0 To UBound(arFile)
            ' Determine the period for this file
            If (HasPeriod(arFile(intI), strPer)) Then
              ' Add this file & period to the table
              tblFile.Rows.Add(arFile(intI), strPer)
            Else
              Status("Cannot determine the period of: " & arFile(intI))
            End If
          Next intI
          ' Order the result
          dtrFound = tblFile.Select("", "Period ASC, File ASC")
          ' Start overall HTML output
          colHtml.Add("<html><body><h4>General statistics</h4>")
          ' Walk through all files
          For intI = 0 To dtrFound.Length - 1
            ' Get this file
            strFile = dtrFound(intI).Item("File")
            ' Show what we are doing
            Status("Processing file " & intI + 1 & "/" & dtrFound.Length & " " & _
                   IO.Path.GetFileNameWithoutExtension(strFile))
            If (Not DoCorefList(strFile, strAction, False, colGlobal, colProt)) Then
              ' Is this interrupt?
              If (bInterrupt) Then
                Status("Coreferential chains: interrupted")
              Else
                ' There is a problem with this file
                Status("frmMain/Reflist: there is a problem processing " & arFile(intI))
              End If
              Exit Sub
            End If
            ' TODO: room to COMBINE information for global purposes
            ' Add global measures
            ' colHtml.Add(colGlobal.Text)
          Next intI
          For intI = 0 To colGlobal.Count - 1
            ' Finish this table
            colGlobal.Item(intI) &= "</table>" & vbCrLf
          Next intI
          ' Finish the GLOBAL report
          colHtml.Add(colGlobal.Text)
          colHtml.Add(colProt.Text)
          colHtml.Add("</body></html>")
          ' Calculate this report and put it in place
          strFile = arFile(0)
          strOutFile = IO.Path.GetDirectoryName(strFile) & "\" & _
            IO.Path.GetFileNameWithoutExtension(strFile) & "_ChainReport.htm"
          ' Add report location
          strReport = strOutFile & "<p>" & colHtml.Text
          ' Save the report on an appropriate location
          IO.File.WriteAllText(strOutFile, strReport)
          ' Publish the report with global measures on the correct tab page
          Me.wbReport.Navigate(strOutFile)
          ' Switch to this tabpage
          Me.TabControl1.SelectedTab = Me.tpReport
          ' Try indicate we are ready
          Status("Coreferential chains: ready")
          ' Did we have a current file?
          If (strCurrentFile <> "") Then
            ' Restore the current file
            If (Not ReadXmlDoc(strCurrentFile, pdxCurrentFile)) Then
              Status("frmMain/RefList: There was a problem recovering the current file")
            End If
            ' Make sure we initialise the [pdxList]
            If (Not InitCurrentFile()) Then
              Status("frmMain/RefList: Could not initialise the current file")
            End If
          End If
        Case "Current"
          ' Do we have a current file
          If (strCurrentFile = "") Then
            Status("First load a file")
            Exit Sub
          End If
          ' Check out current file
          If (Not DoCorefList(strCurrentFile, strAction, True, colGlobal, colProt)) Then
            ' Is this interrupt?
            If (bInterrupt) Then
              Status("Coreferential chains: interrupted")
            Else
              ' There is a problem with this file
              Status("Failed making coreferential chains for: " & strCurrentFile)
            End If
          End If
      End Select
      ' Initialise ref chain editor
      If (Not InitRefChain()) Then
        Status("Could not initialise the RefChain editor")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/RefList error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuRefLearn_Click
  ' Goal:   Learn referential "state" prediction
  '         This gathers training data for a memory-based learning process
  ' History:
  ' 09-04-2013  ERK Created
  ' 13-08-2013  ERK Added preparation for models that only look for 1 of the 5 states
  ' ------------------------------------------------------------------------------------
  Private Sub mnuRefLearn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRefLearn.Click
    Dim strDirIn As String = ""   ' Input directory
    Dim strFileIn As String       ' Input file
    Dim strFile As String         ' Name of the file into which we save the [tdlResults] database
    Dim arFile() As String        ' Input files
    Dim arSel() As String         ' Analysis features
    Dim strRow As String          ' One row
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim dtrChild() As DataRow     ' Children
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intK As Integer           ' Counter
    Dim intPtcTrain As Integer = 70
    Dim bSmall As Boolean = False ' Use a small feature vector
    Dim arState() As String = {"Assumed", "New", "Identity", "Inferred", "Inert"}
    'Dim dtrNew As DataRow         ' General-purpose datarow

    Try
      ' Initialise building training data
      If (Not InitTraining("reflearn")) Then Status("Could not initialize referential state training") : Exit Sub
      ' Initialize directory
      strDirIn = GetTableSetting(tdlSettings, "LastRefLearn")
      If (strDirIn = "") Then strDirIn = strWorkDir
      ' Get the correct input directory
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, "Directory where learn-files are located", strDirIn)) Then
        Status("Could not open input directory")
        Exit Sub
      End If
      ' Ask whether we want to use a small vector (ICHL21) or a larger one
      Select Case MsgBox("Would you like to use a small (ICHL21) vector?(Y)", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Yes
          bSmall = True
        Case MsgBoxResult.No
          bSmall = False
        Case MsgBoxResult.Cancel
          Exit Sub
      End Select
      ' Check what we returned
      If (Not IO.Directory.Exists(strDirIn)) Then Exit Sub
      ' Do away with any old results database
      If (Not tdlResults Is Nothing) Then
        ' Properly dispose
        tdlResults.Dispose()
        tdlResults = Nothing
      End If
      ' Initialise the collection for the results
      colCrpResults.Clear() : intCrpResultsId = 0
      ' Initialise database and add <general> section
      colCrpResults.Add("<?xml version='1.0' standalone='yes'?>")
      colCrpResults.Add("<CrpOview>")
      colCrpResults.Add(" <General>")
      colCrpResults.Add("  <ProjectName>RefStateTest (auto-generated)</ProjectName>")
      colCrpResults.Add("  <Created>" & Format(Now, "s") & "</Created>")
      colCrpResults.Add("  <DstDir>" & IO.Path.GetDirectoryName(strCurrentTrainingRefLearn) & "</DstDir>")
      colCrpResults.Add("  <SrcDir>" & strDirIn & "</SrcDir>")
      colCrpResults.Add("  <Notes>Database with results of referential state experiment</Notes>")
      colCrpResults.Add("  <Analysis></Analysis>")
      colCrpResults.Add(" </General>")
      '' Create a new dataset where we will store our results
      'If (Not CreateDataSet(strCrpResults, tdlResults)) Then
      '  Status("Could not create a dataset for the corpus results")
      '  Exit Sub
      'End If
      '' Check if there is a general section
      'If (tdlResults.Tables("General").Rows.Count = 0) Then
      '  ' Add a <General> section
      '  dtrNew = AddOneDataRowWithParent(tdlResults, "General", "", Nothing)
      '  ' Add the characteristics of this project
      '  With dtrNew
      '    .Item("ProjectName") = "RefStateTest (auto-generated)"
      '    .Item("Created") = Now
      '    .Item("DstDir") = IO.Path.GetDirectoryName(strCurrentTrainingRefLearn)
      '    .Item("SrcDir") = strDirIn
      '    .Item("Notes") = "Database with results of referential state experiment"
      '    .Item("Analysis") = ""
      '  End With
      'End If
      '' Speed up editing on the [tdlResults]
      'tdlResults.Tables("Result").BeginLoadData()
      'tdlResults.Tables("Feature").BeginLoadData()
      ' Save this directory, if it is different from what we have
      SetTableSetting(tdlSettings, "LastRefLearn", strDirIn)
      XmlSaveSettings(strSetFile)
      ' Get all the necessary files in this directory (plus subdirectories)
      arFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories)
      bInterrupt = False
      ' Walk the files
      For intI = 0 To arFile.Length - 1
        ' Get this file
        strFileIn = arFile(intI)
        ' Try to learn from this one
        If (Not OneRefLearnFile(strFileIn, bSmall)) Then
          ' Warn the user and exit
          Status("RefLearn: Could not process file: " & strFileIn)
          Exit Sub
        End If
      Next intI
      ' Determine my own file name
      ' WAS: strFile = IO.Path.GetDirectoryName(strCurrentTrainingRefLearn) & "\RefStLearn_" & Format(Now, "yyyyMMMd_HH-mm") & ".xml"
      strFile = strDirIn & "\RefStLearn_" & Format(Now, "yyyyMMMd_HH-mm") & ".xml"
      strResultDb = strFile
      ' Note my own file's name
      Me.tbDbFile.Text = strResultDb
      ' Finish the database
      colCrpResults.Add("</CrpOview>")
      IO.File.WriteAllText(strResultDb, colCrpResults.Text)
      ' Load the database
      If (Not ReadDataset(strCrpResults, strResultDb, tdlResults)) Then Status("Could not open dataset") : Exit Sub
      ' Divide the data into test and trainingset
      dtrFound = tdlResults.Tables("Result").Select("")
      For intI = 0 To dtrFound.Length - 1
        If (Rnd() * 100 < intPtcTrain) Then
          ' Training set
          dtrFound(intI).Item("Select") = "train"
        Else
          ' Test set
          dtrFound(intI).Item("Select") = "test"
        End If
      Next intI
      For intI = 0 To arState.Length - 1
        ' Show where we are
        Status("Making results for state: " & arState(intI))
        ' Set file name
        strFile = strCurrentTrainingRefLearn & "_" & arState(intI)
        ' Gather training results
        colCrpResults.Clear()
        dtrFound = tdlResults.Tables("Result").Select("Select = 'train'")
        For intJ = 0 To dtrFound.Length - 1
          ' Get my child features
          dtrChild = dtrFound(intJ).GetChildRows("Result_Feature")
          strRow = ""
          For intK = 0 To dtrChild.Length - 1
            If (strRow <> "") Then strRow &= vbTab
            strRow &= dtrChild(intK).Item("Value").ToString
          Next intK
          ' Add the resulting state
          strRow &= vbTab & IIf(dtrFound(intJ).Item("Cat").ToString = arState(intI), arState(intI), "Other")
          ' Add the whole row
          colCrpResults.Add(strRow)
        Next intJ
        IO.File.WriteAllText(strFile & ".train", colCrpResults.Text)
        ' Gather testset results
        colCrpResults.Clear()
        dtrFound = tdlResults.Tables("Result").Select("Select = 'test'")
        For intJ = 0 To dtrFound.Length - 1
          ' Get my child features
          dtrChild = dtrFound(intJ).GetChildRows("Result_Feature")
          strRow = ""
          For intK = 0 To dtrChild.Length - 1
            If (strRow <> "") Then strRow &= vbTab
            strRow &= dtrChild(intK).Item("Value").ToString
          Next intK
          ' Add the resulting state
          strRow &= vbTab & IIf(dtrFound(intJ).Item("Cat").ToString = arState(intI), arState(intI), "Other")
          ' Add the whole row
          colCrpResults.Add(strRow)
        Next intJ
        IO.File.WriteAllText(strFile & ".test", colCrpResults.Text)
      Next intI
      ' Save training and test files for each of the 5 sets
      ' Initialise the editor
      If (Not InitEditors()) Then
        Status("Cannot initialise the database results editor")
        Exit Sub
      End If
      ' Display three more GENERAL values
      With tdlResults.Tables("General").Rows(0)
        Me.tbDbName.Text = .Item("ProjectName").ToString
        Me.tbDbCreated.Text = .Item("Created").ToString
        Me.tbDbSrcDir.Text = .Item("SrcDir").ToString
        Me.tbDbNotes.Text = .Item("Notes").ToString
        Me.tbDbAnalysis.Text = .Item("Analysis").ToString
      End With
      ' Set the feature selection combobox
      With Me.cboDbSelect
        .Items.Clear()
        ' Fill an array
        arSel = Split(Me.tbDbAnalysis.Text, ";")
        For intI = 0 To UBound(arSel)
          .Items.Add(arSel(intI))
        Next intI
        ' Select the first one
        .SelectedItem = 0
      End With
      '' Speed up editing on the [tdlResults]
      'tdlResults.Tables("Feature").EndLoadData()
      'tdlResults.Tables("Result").EndLoadData()
      ' Switch to the correct tabpage
      Me.TabControl1.SelectedTab = Me.tpDbase
      ' Make sure we reset the currently loaded file
      pdxCurrentFile = Nothing
      ' Save the results
      If (SaveTrainingCoref("reflearn")) Then
        Logging("Result saved in: " & strCurrentTrainingRefLearn & ".train (and .test)")
        ' SHow we are ready
        Status("Ready")
      Else
        ' Some save error occurred
        Status("Could not save referential state training data")
      End If
      '' Save the [tdlResults]
      'tdlResults.WriteXml(strResultDb)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/RefLearn error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuRefIrat_Click
  ' Goal:   Compare the current text with a second version of this text and provide measures
  '           for interrater agreement, including Cohen's Kappa
  ' History:
  ' 12-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuRefIrat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRefIrat.Click
    Dim pdxFile As New XmlDocument  ' The second file read as XML
    Dim strFileIn As String = ""    ' Second input file
    Dim strReport As String = ""    ' Report of interrater agreement

    Try
      ' Check if we now have a file opened
      If (pdxCurrentFile Is Nothing) Then
        Status("Open one version of the text using File/Open first")
        Exit Sub
      End If
      ' We have one file -- ask for the other one
      If (Not GetFileName(Me.OpenFileDialog1, IO.Path.GetDirectoryName(strCurrentFile), strFileIn, "PSDX file|*.psdx")) Then Exit Sub
      ' Open this file as xml
      pdxFile.Load(strFileIn)
      ' Perform interrater agreement between them
      If (Not DoIrat(pdxCurrentFile, pdxFile, strReport)) Then Status("There was a problem getting IRAT") : Exit Sub
      ' Show the report
      Me.wbReport.DocumentText = strReport
      Me.TabControl1.SelectedTab = Me.tpReport
      ' Show we are ready
      Status("Ready")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/RefIrat error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   DoCorefList
  ' Goal:   Give an overview of ALL coreferential chains in this document
  ' History:
  ' 12-07-2011  ERK Created
  ' 31-01-2012  ERK Added protagonist counting
  ' 09-02-2012  ERK Toevoegen van [RootEtreeId] aan iedere <Chain>
  ' ------------------------------------------------------------------------------------
  Private Function DoCorefList(ByVal strFile As String, ByVal strAction As String, ByVal bShowHtml As Boolean, _
                               ByRef colGlobal As StringColl, ByRef colProt As StringColl) As Boolean
    Dim colBack As New StringColl         ' Result of our efforts
    Dim strOutFile As String              ' Name of output file
    Dim strOutXml As String = ""          ' Name of the XML output file with the chain information
    Dim strReport As String               ' Resulting HTML report
    ' Dim tdlRefChain As DataSet = Nothing  ' Dataset containing the referential chains
    Dim bXmlOk As Boolean = True          ' Whether we need to make the XML file or not

    Try
      ' Validate
      If (strFile = "") OrElse (Not IO.File.Exists(strFile)) Then Status("File does not exist") : Return False
      ' Check if this file needs loading
      If (strCurrentFile <> strFile) Then
        ' We need to load this file
        If (ReadXmlDoc(strFile, pdxCurrentFile)) Then
          ' All right, so no problem??
        End If
      End If
      ' Validate: check if a text has been loaded
      If (pdxCurrentFile Is Nothing) Then Status("First load a file") : Return False
      ' Check if this file has ANY coref at all
      If (pdxCurrentFile.SelectSingleNode("//fs[@type='coref']") Is Nothing) Then
        ' This file has no coref, so skip it
        Return True
      End If
      ' Calculate the XML file for the chain information
      strOutXml = IO.Path.GetDirectoryName(strFile) & "\" & _
        IO.Path.GetFileNameWithoutExtension(strFile) & "_Chains.xml"
      ' Check if this file exists
      If (IO.File.Exists(strOutXml)) Then
        ' Is it older than I am?
        If (strAction = "Recalc") OrElse (IO.File.GetLastWriteTime(strFile) > IO.File.GetLastWriteTime(strOutXml)) Then
          ' We need to make it
          bXmlOk = False
        End If
      Else
        ' It does not exist, so we need to make it
        bXmlOk = False
      End If
      ' Can we load the XML or do we need to make it?
      If (bXmlOk) Then
        ' Load the xml
        If (Not ReadDataset("RefChain.xsd", strOutXml, tdlRefChain)) Then Status("Cannot load dataset") : Return False
        ' Check if the LogBase matches the current one
        If (GetTableSetting(tdlRefChain, "LogBase") <> intLogBase) Then
          ' The LogBase does not match, so re-do the statistics
          If (Not RefListStat(tdlRefChain, strFile)) Then Status("Unable to do statistics on Refchain dataset") : Return False
          ' Set the new logbase
          SetTableSetting(tdlRefChain, "LogBase", intLogBase)
          ' Store the XML output file
          tdlRefChain.WriteXml(strOutXml)
        End If
      Else
        ' Create an XML data document
        If (Not CreateDataSet("RefChain.xsd", tdlRefChain)) Then Status("Unable to create a dataset") : Return False
        ' Make the dataset
        If (Not RefListXML(tdlRefChain, strFile)) Then Status("Unable to make the RefChain dataset") : Return False
        ' Do the statistics on the data
        If (Not RefListStat(tdlRefChain, strFile)) Then Status("Unable to do statistics on Refchain dataset") : Return False
        ' Store the XML output file
        tdlRefChain.WriteXml(strOutXml)
      End If
      ' Add table rows for global output
      If (Not RefListGlobal(tdlRefChain, strFile, colGlobal, colProt)) Then Return False
      ' Start producing a list
      If (RefListHTML(colBack, tdlRefChain)) Then
        ' Get the report
        strReport = colBack.Text
        ' Calculate this report and put it in place
        strOutFile = IO.Path.GetDirectoryName(strFile) & "\" & _
          IO.Path.GetFileNameWithoutExtension(strFile) & "_Chains.htm"
        ' Add report location
        strReport = strOutFile & "<p>" & strReport
        ' Save the report on an appropriate location
        IO.File.WriteAllText(strOutFile, strReport)
        ' Need to show it?
        If (bShowHtml) Then
          Me.wbReport.DocumentText = strReport
          ' Switch to this tabpage
          Me.TabControl1.SelectedTab = Me.tpReport
          ' Show we are done
          Status("Chain list has been stored in: " & strOutFile)
        End If
      Else
        ' There was a problem
        If (bInterrupt) Then
          ' Show user
          Status("Producing a reference list was interrupted")
          ' Reset interrupt
          bInterrupt = False
        Else
          ' Show user
          Status("There was a problem producing a reference list")
        End If
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/DoCorefList error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuRefCloCh_Click
  ' Goal:   Go to the closest chain (function not yet used)
  ' History:
  ' 12-07-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuRefCloCh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRefCloCh.Click
    Try
      ' Show we are not yet implemented
      MsgBox("This function is not (yet) implemented")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/RefClosestChain error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuMustFirst_Click
  ' Goal:   Select the first constituent in the [objCoref] window that must get a 
  '           Coreference attached to it
  ' History:
  ' 06-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuMustFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMustFirst.Click
    ' Call the appropriate routine
    If (Not FirstMust(Me.tbEdtMain)) Then
      ' Something went wrong...
      Status("Unable to find the first constituent that must receive coreference information")
    End If
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuMustNext_Click
  ' Goal:   Select the first constituent in the [objCoref] window that must get a 
  '           Coreference attached to it
  ' History:
  ' 06-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuMustNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMustNext.Click
    ' Call the appropriate routine
    If (Not NextMust(Me.tbEdtMain)) Then
      ' Something went wrong...
      Status("Unable to find a next constituent that must receive coreference information")
    End If
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuMustPrevious_Click
  ' Goal:   Select the first constituent in the [objCoref] window that must get a 
  '           Coreference attached to it
  ' History:
  ' 06-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuMustPrev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMustPrev.Click
    ' Call the appropriate routine
    If (Not PrevMust(Me.tbEdtMain)) Then
      ' Something went wrong...
      Status("Unable to find a previous constituent that must receive coreference information")
    End If
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuEditPrev_Click
  ' Goal:   Go to the previous line
  ' History:
  ' 10-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuEditPrev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditPrev.Click
    Dim intResId As Integer = -1

    Try
      ' Action depends on the tab page that is active
      Select Case Me.TabControl1.SelectedTab.Name
        Case Me.tpDbase.Name
          ' Try go to the next line in the Results database
          If (Me.dgvDbResult.RowCount > 0) Then
            intResId = objResEd.SelectedId
            intResId -= 1
            objResEd.SelectDgvId(intResId)
          End If
        Case Else
          ' Try the previous line
          If (Not PrevLine()) Then
            ' Impossible...
            Status("Unable to go to the previous line")
          End If
          ' Make sure the cursor in the main window scrolls to the line we are now at
          GoToCurrentLine(Me.tbEdtMain)
      End Select
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/EditPrev error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuEditNext_Click
  ' Goal:   Go to the next line
  ' History:
  ' 10-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuEditNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditNext.Click
    Dim intResId As Integer = -1

    Try
      ' Action depends on the tab page that is active
      Select Case Me.TabControl1.SelectedTab.Name
        Case Me.tpDbase.Name
          ' Try go to the next line in the Results database
          If (Me.dgvDbResult.RowCount > 0) Then
            intResId = objResEd.SelectedId
            intResId += 1
            objResEd.SelectDgvId(intResId)
          End If
        Case Else
          ' Try the next line
          If (Not NextLine()) Then
            ' Impossible...
            Status("Unable to go to the next line")
          End If
          ' Make sure the cursor in the main window scrolls to the line we are now at
          GoToCurrentLine(Me.tbEdtMain)
      End Select
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/EditNext error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   cmdSyntPrev_Click, cmdSyntFoll_Click, cmdTreePrev_Click, cmdTreeNext_Click
  ' Goal:   Show previous/next line in syntax/tree view
  ' History:
  ' 21-04-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdSyntPrev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSyntPrev.Click
    Try
      ' Try going there
      If (Not PrevLine()) Then Status("Unable to go to the previous line") : Exit Sub
      ' make sure we scroll there
      GoToCurrentLine(Me.tbEdtMain)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/SyntPrev error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub cmdSyntFoll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSyntFoll.Click
    Try
      ' Try going there
      If (Not NextLine()) Then Status("Unable to go to the next line") : Exit Sub
      ' make sure we scroll there
      GoToCurrentLine(Me.tbEdtMain)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/SyntPrev error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub cmdTreePrev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTreePrev.Click
    Try
      ' Try going there
      If (Not PrevLine()) Then Status("Unable to go to the previous line") : Exit Sub
      ' make sure we scroll there
      GoToCurrentLine(Me.tbEdtMain)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/SyntPrev error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub cmdTreeNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTreeNext.Click
    Try
      ' Try going there
      If (Not NextLine()) Then Status("Unable to go to the next line") : Exit Sub
      ' make sure we scroll there
      GoToCurrentLine(Me.tbEdtMain)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/SyntPrev error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub cmdDepPrev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDepPrev.Click
    Try
      ' Try going there
      If (Not PrevLine()) Then Status("Unable to go to the previous line") : Exit Sub
      ' make sure we scroll there
      GoToCurrentLine(Me.tbEdtMain)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/DepPrev error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub cmdDepNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDepNext.Click
    Try
      ' Try going there
      If (Not NextLine()) Then Status("Unable to go to the next line") : Exit Sub
      ' make sure we scroll there
      GoToCurrentLine(Me.tbEdtMain)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/DepPrev error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   cmdEditUndo_Click
  ' Goal:   Show previous/next line in syntax/tree view
  ' History:
  ' 21-04-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdEditUndo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEditUndo.Click
    Try
      ' Action depends on tab page
      Select Case Me.TabControl1.SelectedTab.Name
        Case Me.tpTree.Name
          If (PopTreeStack()) Then
            Status("Last action was undone")
          End If
        Case Else
          Status("Undo is not implemented for this page...")
      End Select
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/EditUndo error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   cboSyntStack_SelectedIndexChanged
  ' Goal:   Undo until the selected item
  ' History:
  ' 21-04-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cboSyntStack_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboSyntStack.SelectedIndexChanged
    Try
      ' Action depends on tab page
      Select Case Me.TabControl1.SelectedTab.Name
        Case Me.tpTree.Name
          If (PopTreeStack(Me.cboSyntStack.SelectedIndex)) Then
            Status("Last action was undone")
          End If
        Case Else
          Status("Undo is not implemented for this page...")
      End Select
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/cboSyntStack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   SectionText
  ' Goal:   Set the text of the [tsSectionLabel] element
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub SectionText(ByVal strText As String)
    Try
      ' Find the correct component and set the text
      Me.tsSectionLabel.Text = strText
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/SectionText error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuSectionFirst_Click
  ' Goal:   Load the first section
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuSectionFirst_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSectionFirst.Click
    Dim bStatus As Boolean = bEdtLoad ' Keep the original status of the flag

    Try
      ' Validate
      If (pdxSection Is Nothing) Then Exit Sub
      If (pdxSection.Count = 0) Then Exit Sub
      ' Do as if the editor is loading
      bEdtLoad = True
      ' Find and go to the indicated section
      ShowSection(Me.tbEdtMain, 0)
      '' Only show the translation if this page is actually visible
      'If (Me.TabControl1.SelectedTab.Name = Me.tpTranslation.Name) Then
      '  ShowTranslation(Me.tbTranslation, intCurrentSection)
      'End If
      ' Set focus to the main editor
      Me.tbEdtMain.Focus()
      ' Switch off editor loading
      bEdtLoad = bStatus
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuSectionFirst` error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuSectionLast_Click
  ' Goal:   Load the last section
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuSectionLast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSectionLast.Click
    Dim bStatus As Boolean = bEdtLoad ' Keep the original status of the flag

    Try
      ' Validate
      If (pdxSection Is Nothing) Then Exit Sub
      If (pdxSection.Count = 0) Then Exit Sub
      ' Do as if the editor is loading
      bEdtLoad = True
      ' Find and go to the indicated section
      ShowSection(Me.tbEdtMain, pdxSection.Count - 1)
      '' Only show the translation if this page is actually visible
      'If (Me.TabControl1.SelectedTab.Name = Me.tpTranslation.Name) Then
      '  ShowTranslation(Me.tbTranslation, intCurrentSection)
      'End If
      ' Set focus to the main editor
      Me.tbEdtMain.Focus()
      ' Switch off editor loading
      bEdtLoad = bStatus
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuSectionLast error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuSectionNext_Click
  ' Goal:   Load the next section
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuSectionNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSectionNext.Click
    Dim bStatus As Boolean = bEdtLoad ' Keep the original status of the flag

    Try
      ' Validate
      If (pdxSection Is Nothing) Then Exit Sub
      If (pdxSection.Count = 0) Then Exit Sub
      If (intCurrentSection >= pdxSection.Count - 1) Then Exit Sub
      ' Do as if the editor is loading
      bEdtLoad = True
      ' Find and go to the indicated section
      ShowSection(Me.tbEdtMain, intCurrentSection + 1)
      '' Only show the translation if this page is actually visible
      'If (Me.TabControl1.SelectedTab.Name = Me.tpTranslation.Name) Then
      '  ShowTranslation(Me.tbTranslation, intCurrentSection)
      'End If
      ' Set focus to the main editor
      Me.tbEdtMain.Focus()
      ' Switch off editor loading
      bEdtLoad = bStatus
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuSectionNext error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuSectionPrevious_Click
  ' Goal:   Load the previous section
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuSectionPrevious_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSectionPrevious.Click
    Dim bStatus As Boolean = bEdtLoad ' Keep the original status of the flag

    Try
      ' Validate
      If (pdxSection Is Nothing) Then Exit Sub
      If (pdxSection.Count = 0) Then Exit Sub
      If (intCurrentSection <= 0) Then Exit Sub
      ' Do as if the editor is loading
      bEdtLoad = True
      ' Find and go to the indicated section
      ShowSection(Me.tbEdtMain, intCurrentSection - 1)
      '' Only show the translation if this page is actually visible
      'If (Me.TabControl1.SelectedTab.Name = Me.tpTranslation.Name) Then
      '  ShowTranslation(Me.tbTranslation, intCurrentSection)
      'End If
      ' Set focus to the main editor
      Me.tbEdtMain.Focus()
      ' Switch off editor loading
      bEdtLoad = bStatus
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuSectionPrevious error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuSectionGoto_Click
  ' Goal:   Offer user the choice to go to a specific section
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuSectionGoto_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSectionGoto.Click
    Dim bStatus As Boolean = bEdtLoad ' Keep the original status of the flag

    Try
      ' Validate
      If (pdxSection Is Nothing) Then Exit Sub
      If (pdxSection.Count = 0) Then Exit Sub
      ' Call the appropriate function
      SectionGoto()
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuSectionGoto error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   SectionGoto
  ' Goal:   Offer user the choice to go to a specific section
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub SectionGoto()
    Dim bStatus As Boolean = bEdtLoad ' Keep the original status of the flag

    Try
      ' Validate
      If (pdxSection Is Nothing) Then Exit Sub
      If (pdxSection.Count = 0) Then Exit Sub
      ' Find and go to the indicated section
      With frmSection
        ' Show it
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Show this section
            GoShowSection(.Section)
            '' Do as if the editor is loading
            'bEdtLoad = True
            '' Get the focus away from me
            'Me.tbEdtFeatures.Focus()
            '' Go to the section number selected
            'ShowSection(Me.tbEdtMain, .Section)
            ' '' Only show the translation if this page is actually visible
            ''If (Me.TabControl1.SelectedTab.Name = Me.tpTranslation.Name) Then
            ''  ShowTranslation(Me.tbTranslation, intCurrentSection)
            ''End If
            '' Set focus to the main editor
            'Me.tbEdtMain.Focus()
            '' Switch off editor loading
            'bEdtLoad = bStatus
        End Select
      End With
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/SectionGoto error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GoShowSection
  ' Goal:   Go to the indicated section, and show it
  ' History:
  ' 01-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Sub GoShowSection(ByVal intSect As Integer)
    Dim bStatus As Boolean = bEdtLoad ' Keep the original status of the flag

    Try
      ' Do as if the editor is loading
      bEdtLoad = True
      ' Get the focus away from me
      Me.tbEdtFeatures.Focus()
      ' Go to the section number selected
      ShowSection(Me.tbEdtMain, intSect)
      '' Only show the translation if this page is actually visible
      'If (Me.TabControl1.SelectedTab.Name = Me.tpTranslation.Name) Then
      '  ShowTranslation(Me.tbTranslation, intCurrentSection)
      'End If
      ' Set focus to the main editor
      Me.tbEdtMain.Focus()
      ' Switch off editor loading
      bEdtLoad = bStatus
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/GoShowSection error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuSectionInsert_Click
  ' Goal:   Insert a section before this line
  ' History:
  ' 13-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuSectionInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSectionInsert.Click
    Try
      ' Are we on the right tabpage?
      If (Me.TabControl1.SelectedTab.Name <> Me.tpEdit.Name) Then
        ' Warn user
        Status("You are trying to insert a section, but you are not on the text-editor page")
        Exit Sub
      End If
      ' Do as if the editor is loading
      bEdtLoad = True
      ' Call the appropriate subroutine
      SectionInsert(Me.tbEdtMain)
      ' Switch off editor loading
      bEdtLoad = False
      ' Set the dirty flag
      SetDirty(True)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuSectionInsert error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuSectionDelete_Click
  ' Goal:   Delete the section that is present before this line (if any)
  ' History:
  ' 13-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuSectionDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSectionDelete.Click
    Try
      ' Are we on the right tabpage?
      If (Me.TabControl1.SelectedTab.Name <> Me.tpEdit.Name) Then
        ' Warn user
        Status("You are trying to delete a section, but you are not on the text page")
        Exit Sub
      End If
      ' Do as if the editor is loading
      bEdtLoad = True
      ' Call the appropriate subroutine
      SectionDelete(Me.tbEdtMain)
      ' Switch off editor loading
      bEdtLoad = False
      ' Set the dirty flag
      SetDirty(True)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuSectionDelete error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Switch off editor loading
      bEdtLoad = False
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuSectionDelAll_Click
  ' Goal:   Delete all section breaks
  ' History:
  ' 14-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuSectionDelAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSectionDelAll.Click
    Try
      ' Do as if the editor is loading
      bEdtLoad = True
      If (DelAllSections()) Then
        Status("All sections have been deleted")
      Else
        Status("Not all sections were deleted - there is an error")
      End If
      ' Switch off editor loading
      bEdtLoad = False
      ' Set the dirty flag
      SetDirty(True)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuSectionDelAll_Click error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Switch off editor loading
      bEdtLoad = False
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuSectionDerive_Click
  ' Goal:   Try to derive all section breaks based on (CODE <heading>) entries
  ' History:
  ' 14-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuSectionDerive_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSectionDerive.Click
    Try
      ' Do as if the editor is loading
      bEdtLoad = True
      If (AutoAddSections()) Then
        Status("Sections have been derived")
      Else
        Status("Section derivation was not successful")
      End If
      ' Switch off editor loading
      bEdtLoad = False
      ' Set the dirty flag
      SetDirty(True)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuSectionDerive_Click error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Switch off editor loading
      bEdtLoad = False
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   tbEdtMain_VScroll
  ' Goal:   Redraw the position when the user makes a vertical scroll
  ' History:
  ' 18-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub tbEdtMain_VScroll(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbEdtMain.VScroll
    Try
      ' Start the Lines drawing timer
      Me.tmeLines.Enabled = True
      ' Make sure the timer starts AGAIN
      Me.tmeLines.Start()
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/tbEdtMain_VScroll error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   tmeLines_Tick
  ' Goal:   Redraw the position when the user makes a vertical scroll
  ' History:
  ' 18-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub tmeLines_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmeLines.Tick
    Dim bStatus As Boolean = bEdtLoad   ' Retain original status

    Try
      ' Switch off the Lines drawing timer (resets it)
      Me.tmeLines.Enabled = False
      ' Are we initialized?
      If ((bInit) AndAlso (Me.tbEdtMain.Visible)) And (Not bEdtLoad) And (Not bAutoBusy) Then
        ' Do as if the editor is loading
        bEdtLoad = True
        ' Redraw what needs to be redrawn
        ShowPosition(Me.tbEdtMain)
        ' Switch off editor loading
        bEdtLoad = bStatus
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/tmeLines_Tick error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   TrySaveReport
  ' Goal:   Try save the report to the location the user specifies
  ' History:
  ' 28-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub TrySaveReport()
    Dim strFile As String   ' File to save to

    Try
      ' Get the filename from the user
      With Me.SaveFileDialog1
        ' The initial directory is the one we already know
        .InitialDirectory = strWorkDir
        ' The name is derived from the name of this project
        strFile = strReportName & ".psdx"
        ' Set the default extention
        .DefaultExt = "html"
        ' Set the default filter
        .Filter = "Report HTML file|*.html"
        ' Assign default file name to the FileSave dialog
        .FileName = strFile
        ' Show the actual dialog to the user
        Select Case .ShowDialog()
          Case Windows.Forms.DialogResult.OK, Windows.Forms.DialogResult.Yes
            ' Get the filename that the user has finally selected
            strFile = .FileName
          Case Else
            ' Aborted, so exit
            Status("File is not saved")
            Exit Sub
        End Select
      End With
      ' If we come here, then the user has selected a filename properly - do the saving...
      IO.File.WriteAllText(strFile, Me.wbReport.DocumentText)
      ' Show success
      Status("File saved to " & strFile)
    Catch ex As Exception
      ' Show error
      HandleErr("frmReport/TrySaveReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  Private Sub TabControl1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TabControl1.KeyDown
    ' Is the TREE tab selected?
    If (Me.TabControl1.SelectedTab Is Me.tpTree) Then
      ' Let the other handle it
      litTree_KeyDown(sender, e)
    End If
    ' Stop
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   TabControl1_SelectedIndexChanged
  ' Goal:   Try save the report to the location the user specifies
  ' History:
  ' 01-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
    Dim bHide As Boolean = True ' Whether we need to hide the label

    Try
      ' Are we initialised?
      If (bEdtLoad) OrElse (Not bInit) Then Exit Sub
      ' Make sure tree stack is not shown
      Me.cboSyntStack.Visible = False
      ' Check which index we now are
      Select Case Me.TabControl1.SelectedTab.Name
        Case Me.tpEdit.Name
          ' Check if redrawing is needed
          If (bEdtDirty) Then
            If (Not eTreeRedoText()) Then Logging("Could not redraw editor") : Exit Sub
            ' Reset dirty flag
            bEdtDirty = False
          End If
        Case Me.tpSyntax.Name
          ' Calculate and show the syntax of this line
          ShowSyntax(Me.tbEdtSyntax)
        Case Me.tpTranslation.Name
          ' Calculate and show the translation for this current section
          ShowTranslation(False)
        Case Me.tpTree.Name
          ' Calculate and show the syntax of this line in a tree
          ShowTree(Me.litTree)
          '' Reset the query node stack
          'colQstack.Clear()
          ' Reset visibility of help function
          Me.mnuSyntaxEditHelp.Checked = False
          Me.tbTreeHelp.Visible = False
          bHide = False
          ' Show tree edit stack?
          If (TreeStackSize() > 0) Then
            Me.cboSyntStack.Visible = True
          End If
        Case Me.tpDep.Name
          ' Show current dependency
          DepCurrentShow()
        Case Else
      End Select
      ' Auto switch-off editmode if needed
      If (Me.TabControl1.SelectedTab.Name <> Me.tpTree.Name) Then
        ' Switch mode off
        eTreeEditMode("reset")
      End If
      ' Need to hide the label?
      If (bHide) Then
        ' Hide it
        Me.tbLabel.Visible = False
        ' Reset the label busy flag
        bLabelSelect = False
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/TabControl error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   DoAskListView
  ' Goal:   Ask the user which antecedent for [strSrcNode] to choose from a set of antecedents
  ' History:
  ' 17-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function DoAskListView(ByVal strQuestion As String, ByRef ndxSrcNode As Xml.XmlNode, ByRef ndxAntNode As Xml.XmlNode, _
        ByRef tblSrc As DataTable, ByRef tblDst As DataTable, ByRef intDstId As Integer, _
        ByRef bIsInferred As Boolean) As DialogResult
    Dim strSrcNode As String  ' The text of the source node
    Dim strItem As String     ' This element
    Dim intAntId As Integer   ' ID of potential antecedent
    Dim intI As Integer       ' Counter

    Try
      ' Validate
      If (tblSrc Is Nothing) OrElse (ndxSrcNode Is Nothing) Then Return Windows.Forms.DialogResult.Cancel
      ' Get the source node's text
      strSrcNode = NodeInfo(ndxSrcNode)
      ' Set the question
      Me.tbAskProblem.Text = NodeInfo(ndxSrcNode) & " >> " & NodeInfo(ndxAntNode) & vbCrLf & strQuestion
      ' Set the targets
      SetTargets(ndxSrcNode.Attributes("Id").Value, tblSrc, tblDst)
      ' Make stuff visible
      Me.panAsk.Visible = True
      ' Make sure the translation window scrolls along
      ShowPde(Me.tbEdtMain, Me.tbMainPde, GetAutoPos())
      ' Deselect all previous ones
      For intI = 0 To Me.lvAskTarget.Items.Count - 1
        Me.lvAskTarget.Items(intI).Selected = False
      Next intI
      ' Give a message
      Status("Give the referent for " & strSrcNode)
      ' Has an antecedent been given?
      If (ndxAntNode Is Nothing) Then
        ' Set the focus to the listbox and select the first element
        Me.lvAskTarget.Items(0).Selected = True
      Else
        ' Start by selecting "0" just in case
        Me.lvAskTarget.Items(0).Selected = True
        ' Set the focus to the listbox element with the correct node ID
        intAntId = ndxAntNode.Attributes("Id").Value
        ' Try find this antecedent
        For intI = 0 To Me.lvAskTarget.Items.Count - 1
          ' Consider this element
          strItem = Me.lvAskTarget.Items(intI).SubItems(0).Text
          ' Is it numeric?
          If (IsNumeric(strItem)) Then
            ' See if this is the correct one
            If (CInt(strItem) = intAntId) Then
              ' This is the correct one, so select it
              Me.lvAskTarget.Items(intI).Selected = True
              ' Exit the loop
              Exit For
            End If
          End If
        Next intI
      End If
      Me.lvAskTarget.Focus()
      ' Wait for answer
      bAskReady = False
      While (Not bAskReady)
        ' Make sure we are interruptable
        Application.DoEvents()
      End While
      ' Pick up the correct resulting ID
      intDstId = loc_intDstId
      ' Make stuff invisible again
      Me.panAsk.Visible = False
      Me.TabControl2.SelectedTab = Me.tpCoref
      ' Check whether user pressed "Interrupt"
      If (bInterrupt) Then
        ' Return a more negative result
        Return Windows.Forms.DialogResult.Abort
      End If
      ' Check whether SHIFT was pressed
      bIsInferred = ShiftPressed()
      ' Return positive result
      Return Windows.Forms.DialogResult.OK
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/DoAskListView error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return DialogResult.Cancel
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MakeColumn
  ' Goal:   Make a column header for listview
  ' History:
  ' 01-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function MakeColumn(ByVal strName As String, ByVal intWidth As Integer, _
                              ByVal intAlign As HorizontalAlignment) As ColumnHeader
    Dim colThis As ColumnHeader   ' One columnheader

    Try
      ' Validate
      If (strName = "") Then Return Nothing
      ' Set the column headers
      colThis = New ColumnHeader(strName)
      With colThis
        ' Set the name
        .Name = strName
        ' Set the text
        .Text = strName
        ' Set horizontal alignment
        .TextAlign = intAlign
        ' Set width
        .Width = intWidth
      End With
      ' REturn the result
      Return colThis
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/GetColumn error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   TargetInit
  ' Goal:   Initialise the listview
  ' History:
  ' 01-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub TargetInit()
    Try
      ' Validate
      If (bTargetInit) Then Exit Sub
      ' Access the listview
      With Me.lvAskTarget
        ' Set the view
        .View = View.Details
        ' Don't allo user to edit items
        .LabelEdit = False
        ' Don't allow column reordering
        .AllowColumnReorder = False
        ' Don't display check boxes
        .CheckBoxes = False
        ' Make sure each row is completely selected
        .FullRowSelect = True
        ' Don't know about the gridlines...
        .GridLines = True
        ' Don't allow more than 1 selection
        .MultiSelect = False
        ' Clear all columns
        .Columns.Clear()
        ' Make the new columns
        .Columns.Add(MakeColumn("Id", 0, HorizontalAlignment.Left))
        .Columns.Add(MakeColumn("Loc", -2, HorizontalAlignment.Left))
        .Columns.Add(MakeColumn("Eval", -2, HorizontalAlignment.Right))
        .Columns.Add(MakeColumn("Ant", -2, HorizontalAlignment.Left))
        .Columns.Add(MakeColumn("Chain", -2, HorizontalAlignment.Left))
        ' Make sure the first column is NOT visible (see above)
        ' Initialize the width of the [Ant] column
        .Columns("Ant").Width = 200
      End With
      ' Show we are initialised
      bTargetInit = True
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/TargetInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetLvItem
  ' Goal:   Get one item for the listview
  ' History:
  ' 01-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function GetLvItem(ByVal strId As String, ByVal strLine As String, ByVal strEval As String, _
       ByVal strConstituent As String, ByVal strEqual As String) As ListViewItem
    Dim lvItem As ListViewItem  ' One listview item

    Try
      ' Create a listview item
      lvItem = New ListViewItem(strId)
      ' Populate its subitem parts
      With lvItem.SubItems
        .Add(strLine)
        .Add(strEval)
        .Add(strConstituent)
        .Add(strEqual)
      End With
      lvItem.Name = strId
      ' Return the result
      Return lvItem
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/TargetInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   SetTargets
  ' Goal:   Set the listbox from the source and the antecedent stacks
  ' History:
  ' 17-06-2010  ERK Created
  ' 01-07-2010  ERK Changed to listview instead of listbox
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub SetTargets(ByVal intSrcId As Integer, ByRef tblSrcStack As DataTable, ByRef tblAntStack As DataTable)
    Dim intI As Integer         ' Counter
    Dim dtrAnt() As DataRow     ' Sorted antecedent stack
    Dim dtrThis As DataRow      ' One element
    Dim strConst As String      ' The constituent
    Dim strEqual As String      ' What to put in the [Equal] column

    Try
      ' Validate
      If (tblSrcStack Is Nothing) Then Exit Sub
      ' Initialise again
      TargetInit()
      ' Access the listview
      With Me.lvAskTarget
        ' Clear the listview
        .Items.Clear()
        ' Add "New" and "Assumed"
        .Items.Add(GetLvItem(strRefNew, strRefNew, "-", "this is a [New] NP, not linked", ""))
        .Items.Add(GetLvItem(strRefAssumed, strRefAssumed, "-", "[Assumed] knowledge (hearer-old)", ""))
        .Items.Add(GetLvItem(strRefInert, strRefInert, "-", "referentially [Inert]", ""))
        .Items.Add(GetLvItem("Skip", "Skip", "-", "don't do this element right now", ""))
      End With
      ' Read the source stack one at a time backwards...
      For intI = tblSrcStack.Rows.Count - 1 To 0 Step -1
        ' Get this element
        dtrThis = tblSrcStack.Rows(intI)
        ' Only treat this element if the ID is smaller than our own ID
        If (dtrThis.Item("Id") < intSrcId) Then
          ' Get the constituent
          strConst = "[" & dtrThis.Item("Label") & " " & VernToEnglish(dtrThis.Item("Vern")) & "]"
          ' Put the source stack element in the listbox, including its ID
          Me.lvAskTarget.Items.Add(GetLvItem(dtrThis.Item("Id"), dtrThis.Item("Loc"), "-", strConst, ""))
        ElseIf (dtrThis.Item("Id") > intSrcId) Then
          ' These are cataphoric elements, so should be clearly labeled as such
          strConst = "cataphoric[" & dtrThis.Item("Label") & " " & VernToEnglish(dtrThis.Item("Vern")) & "]"
          ' Put the source stack element in the listbox, including its ID
          Me.lvAskTarget.Items.Add(GetLvItem(dtrThis.Item("Id"), dtrThis.Item("Loc"), "-", strConst, ""))
        End If
      Next intI
      ' Are there antecedents?
      If (tblAntStack Is Nothing) Then Exit Sub
      ' Make a sorted antecedent stack of NPs, up to the maximum number
      dtrAnt = tblAntStack.Select("Level=0", "Id DESC")
      ' --- Debug.Print(tblAntStack.Rows(406).Item("Level"))
      ' Put the antecedent stack's elements in there (up to a maximum)
      For intI = 0 To dtrAnt.Length - 1
        ' Get this element
        dtrThis = dtrAnt(intI)
        ' Get the constituent
        strConst = "[" & dtrThis.Item("Label") & " " & VernToEnglish(dtrThis.Item("Vern")) & "]"
        ' Decide what to put in the [Equal] column
        If (dtrThis.Item("Equal") = "") Then
          ' Do we have a head?
          If (dtrThis.Item("Head") = "") Then
            strEqual = ""
          Else
            strEqual = "=" & dtrThis.Item("Head")
          End If
        Else
          strEqual = dtrThis.Item("Equal")
        End If
        ' Put the source stack element in the listbox, including its ID
        Me.lvAskTarget.Items.Add(GetLvItem(dtrThis.Item("Id"), dtrThis.Item("Loc"), dtrThis.Item("Eval"), strConst, _
                                          strEqual))
        If (Not bShowAllAnt) Then
          ' Check if we are beyond our limit
          If (intI > GetAntMax()) Then Exit For
        End If
      Next intI
      ' Make a sorted antecedent stack of the IPs, up to the maximum number divided by two
      dtrAnt = tblAntStack.Select("Level=1", "Id DESC")
      ' Put the antecedent stack's elements in there (up to a maximum)
      For intI = 0 To dtrAnt.Length - 1
        ' Get this element
        dtrThis = dtrAnt(intI)
        ' Get the constituent
        strConst = "[" & dtrThis.Item("Label") & " " & VernToEnglish(dtrThis.Item("Vern")) & "]"
        ' Put the source stack element in the listbox, including its ID
        Me.lvAskTarget.Items.Add(GetLvItem(dtrThis.Item("Id"), dtrThis.Item("Loc"), dtrThis.Item("Eval"), strConst, "(IP)"))
        ' Check if we are beyond our limit
        If (intI > GetAntMax() \ 2) Then Exit For
      Next intI
      ' Success!
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/SetTargets error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsStop_Click
  ' Goal:   Stop autocoreferencing
  ' History:
  ' 17-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsStop.Click
    Try
      ' Make sure we stop asking around
      bInterrupt = True
      bAskReady = True
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ToolsStop error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuViewDependency_Click
  ' Goal:   Show or Hide the dependency tab page
  ' History:
  ' 03-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuViewDependency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewDependency.Click
    ' Toggle the visibility of the dependency tab page
    If (Me.mnuViewDependency.Checked) Then
      Me.tpDep.Parent = Me.TabControl1
      DepCurrentShow()
    Else
      Me.tpDep.Parent = Me.tabHidden
    End If
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuViewAuto_Click
  ' Goal:   Show or Hide the autocoreferencing stuff
  ' History:
  ' 18-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuViewAuto_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewAuto.Click
    ' Toggle the visibility of the autocoref log
    Me.panAutoLog.Visible = (Me.mnuViewAuto.Checked)
    Me.tbAutoLog.Visible = (Me.mnuViewAuto.Checked)
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuViewText_Click
  ' Goal:   View the complete text
  ' History:
  ' 13-07-2011  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuViewText_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewText.Click
    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Exit Sub
      ' Go to correct tab page
      Me.TabControl1.SelectedTab = Me.tpReport
      Me.wbReport.DocumentText = "Working..."
      ' Find content
      ViewFileContent("org")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ViewText error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuViewLemmas_Click
  ' Goal:   View the lemma's + glosses of all the words in all the lines
  ' History:
  ' 30-01-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuViewLemmas_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewLemmas.Click
    Dim strFile As String = ""

    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Status("First load a text") : Exit Sub
      ' Go to correct tab page
      Me.TabControl1.SelectedTab = Me.tpReport
      Me.wbReport.DocumentText = "Working..."
      ' Find content
      If (LemmaFileView(False, strFile)) Then
        ' Show it
        Me.wbReport.Navigate(strFile)
        Status("Ready")
      Else
        Status("There was an error")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ViewText error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuViewTrans_Click
  ' Goal:   View the complete translation of the text
  ' History:
  ' 13-07-2011  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuViewTrans_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewTrans.Click
    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Exit Sub
      If (strCurrentTransLang = "") Then Status("There is no translation language available") : Exit Sub
      ' Go to correct tab page
      Me.TabControl1.SelectedTab = Me.tpReport
      Me.wbReport.DocumentText = "Working..."
      ' Find content
      ViewFileContent(strCurrentTransLang)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ViewTranslation error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuViewGloss_Click
  ' Goal:   View the complete text, and a line by line gloss into PDE
  ' History:
  ' 27-11-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuViewGloss_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewGloss.Click
    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Exit Sub
      If (strCurrentTransLang = "") Then Status("There is no translation language available") : Exit Sub
      ' Go to correct tab page
      Me.TabControl1.SelectedTab = Me.tpReport
      Me.wbReport.DocumentText = "Working..."
      ' Find content
      ViewFileContent("org", strCurrentTransLang)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ViewTranslation error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuViewSelectTrans_Click
  ' Goal:   Select the (back) translation to be shown and to be available for editing
  ' History:
  ' 17-03-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuViewSelectTrans_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewSelectTrans.Click
    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Exit Sub
      ' Direct the user to the combobox
      Me.TabControl1.SelectedTab = Me.tpGeneral
      Me.cboGenTransLang.Focus()
      ' Show all is well
      Status("Select the language here")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ViewSelectTrans error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuViewChart_Click
  ' Goal:   Produce a "chart" for the current section of the text
  ' History:
  ' 12-03-2012  ERK Created
  ' 09-07-2013  ERK Add option for all sections of text + make sure V2 column stays 'clean' of other constituents
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuViewChart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuViewChart.Click
    Dim bSubCl As Boolean = False ' Subclause property
    Dim bAllSc As Boolean = False ' Include all sections in charting or not?
    Dim strOutType As String = "" ' Type of output
    Dim intV1ptc As Integer = 85  ' Percentage for V1
    Dim intV2ptc As Integer = 95  ' Percentage for V2
    Dim intSptc As Integer = 90   ' Percentage for subject

    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Exit Sub
      ' Get the options
      With frmChart
        ' Initialise the values
        .V1ptc = intV1ptc
        .V2ptc = intV2ptc
        .Sptc = intSptc
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK, Windows.Forms.DialogResult.Yes
            ' Get the properties
            bSubCl = .SubCl
            bAllSc = .AllSections
            strOutType = .OutputType
            intV1ptc = .V1ptc
            intV2ptc = .V2ptc
            intSptc = .Sptc
          Case Else
            Exit Select
        End Select
      End With
      ' Go to correct tab page
      Me.TabControl1.SelectedTab = Me.tpReport
      Me.wbReport.DocumentText = "Working..."
      ' Make sure we know the sections
      SectionCalculate()
      ' Produce chart
      If (MakeChart(bSubCl, bAllSc, strOutType, intV1ptc, intV2ptc, intSptc)) Then
        Status("Ready")
      Else
        Status("There was an error producing the chart")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ViewChart error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '' ------------------------------------------------------------------------------------------------------------
  '' Name:   mnuViewEndnote_Click
  '' Goal:   Check all the endnote references in a document
  '' History:
  '' 28-11-2012  ERK Created
  '' ------------------------------------------------------------------------------------------------------------
  'Private Sub mnuViewEndnote_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
  '  Dim strBack As String = ""    ' What we return

  '  Try
  '    ' Make alist of references in my dissertation
  '    DissertationList(strBack)
  '    ' Display the list
  '    Me.wbReport.DocumentText = strBack
  '  Catch ex As Exception
  '    ' Show error
  '    HandleErr("frmMain/ViewEndnote error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
  '  End Try
  'End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   lvAskTarget_DoubleClick
  ' Goal:   The user has selected one element
  ' History:
  ' 17-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub lvAskTarget_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvAskTarget.DoubleClick
    AskTargetChoiceLv()
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   lvAskTarget_KeyPress
  ' Goal:   Detect whether user presses "Enter" to indicate his choice
  ' History:
  ' 17-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub lvAskTarget_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles lvAskTarget.KeyPress
    ' Check which key is pressed (if this is ENTER)
    If (e.KeyChar = ChrW(Keys.Enter)) Then
      ' Okay, user pressed "ENTER", so invoke the Target subroutine
      AskTargetChoiceLv()
    ElseIf (e.KeyChar = ChrW(Keys.Escape)) Then
      ' User pressed ESC, which indicates that we skip this one
      loc_intDstId = -4
      bAskReady = True
    End If
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   lvAskTarget_SelectedIndexChanged
  ' Goal:   Show the constraints for this element in the list
  ' History:
  ' 25-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub lvAskTarget_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvAskTarget.SelectedIndexChanged
    Dim intIdx As Integer ' The selected index of the listbox
    Dim strItem As String ' The selected item

    Try
      ' Is something selected?
      If (Me.lvAskTarget.SelectedIndices.Count = 0) Then Exit Sub
      ' Try to get the selected index
      intIdx = Me.lvAskTarget.SelectedIndices(0)
      ' Check what it is
      If (intIdx >= 0) AndAlso (Me.lvAskTarget.Items.Count > 0) Then
        ' Some kind of stuff was pressed
        strItem = Trim(Me.lvAskTarget.Items(intIdx).SubItems(0).Text)
        ' See if this is numeric
        If (IsNumeric(strItem)) Then
          ' Get the "Id" value of the antecedent's node
          intIdx = CInt(strItem)
          ' Now show the constraints for this antecedent
          Me.ToolTip1.SetToolTip(Me.lvAskTarget, ShowConstraints(intIdx))
        End If
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/lvAskTarget_SelectedIndexChanged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   AskTargetChoiceLv
  ' Goal:   User has selected one element in [lvAskTarget]
  ' History:
  ' 17-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub AskTargetChoiceLv()
    Dim intIdx As Integer       ' Index of selected item
    Dim strItem As String = ""  ' Text of the selected item

    Try
      ' Validate
      If (Me.lvAskTarget.SelectedIndices.Count = 0) Then Exit Sub
      ' Find out which element has been selected
      intIdx = Me.lvAskTarget.SelectedIndices(0)
      strItem = Me.lvAskTarget.Items(intIdx).SubItems(0).Text
      ' Action depends on the item that has been selected
      Select Case strItem
        Case strRefNew
          ' Indicate a new element is assumed
          loc_intDstId = -2
        Case strRefAssumed
          ' Indicate an "Assumed" relation has to be made
          loc_intDstId = -3
        Case "Skip"
          ' Indicate that we don't want to do this right now
          loc_intDstId = -4
        Case strRefInert
          ' Indicate an "Inert" relation has to be made
          loc_intDstId = -5
        Case Else
          ' Make sure the result is numeric
          If (IsNumeric(strItem)) Then
            ' Set the ID number of the target
            loc_intDstId = CInt(strItem)
          Else
            ' There's been an error
            loc_intDstId = -1
          End If
      End Select
      ' Make sure the flag is set
      bAskReady = True
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/AskTargetChoiceLv error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '' ------------------------------------------------------------------------------------------------------------
  '' Name:   tbTranslation_SelectionChanged
  '' Goal:   Depending on where we are, we want the user to be able to change the translation text
  '' History:
  '' 29-06-2010  ERK Created
  '' ------------------------------------------------------------------------------------------------------------
  'Private Sub tbTranslation_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs)
  '  Dim strTranslation As String  ' The translation for the selected line
  '  Dim intLine As Integer        ' The number of the line where we are
  '  Dim intPdeLine As Integer     ' The PDE line number to be searched
  '  Dim intPos As Integer         ' Position within string
  '  'Dim intTransSt As Integer     ' Start of translation
  '  'Dim intTransEn As Integer     ' End of translation
  '  Dim fntTrans As System.Drawing.Font = New Font("Times New Roman", 10, FontStyle.Regular)

  '  Try
  '    ' Validate
  '    If (bEdtLoad) OrElse (bTransBusy) Then Exit Sub
  '    ' Determine which line we are (not sure if this is possible...)
  '    intPos = Me.tbTranslation.SelectionStart
  '    ' Validate
  '    If (intPos = 0) Then Exit Sub
  '    ' Only proceed if we are on a BOLD letter!
  '    If (Not Me.tbTranslation.SelectionFont.Bold) OrElse (Mid(Me.tbTranslation.Text, Me.tbTranslation.SelectionStart, 1) = " ") Then
  '      ' Make sure the translation window is NOT visible!!
  '      tbPde.Visible = False
  '      Exit Sub
  '    End If
  '    If (GetTransLine(intPos, intLine)) Then
  '      '' ============= DEBUGGING =============
  '      '' Get the current start and end position of the translation 
  '      'If (GetTransPosition(intPos, intTransSt, intTransEn)) Then
  '      '  Debug.Print(Mid(Me.tbTranslation.Text, intTransSt, intTransEn - intTransSt))
  '      'End If
  '      '' ======================================
  '      ' Now get the REAL line number by dividing by 2
  '      intPdeLine = intLine \ 2
  '      ' Get the current translation for this line (in this section)
  '      strTranslation = GetPde(intPdeLine)
  '      ' Show this translation in an element
  '      With tbPde
  '        .Parent = Me.tbTranslation
  '        .Font = fntTrans
  '        .Text = strTranslation
  '        ' Set the size and location
  '        .Left = 70
  '        .Height = 40
  '        .Width = Me.tbTranslation.Width - (2 * .Left)
  '        .Top = Me.tbTranslation.GetPositionFromCharIndex(Me.tbTranslation.SelectionStart).Y + 40
  '        ' Show it
  '        .Visible = True
  '      End With
  '    End If
  '  Catch ex As Exception
  '    ' Show error
  '    HandleErr("frmMain/tbTranslation_SelectionChanged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
  '  End Try
  'End Sub

  '' ------------------------------------------------------------------------------------------------------------
  '' Name:   tbPde_KeyPress
  '' Goal:   Either disregard (ESC) or save (ENTER) the translation
  '' History:
  '' 29-06-2010  ERK Created
  '' ------------------------------------------------------------------------------------------------------------
  'Private Sub tbPde_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbPde.KeyPress
  '  Dim strTranslation As String  ' The translation for the selected line
  '  Dim intLine As Integer        ' The number of the line where we are
  '  Dim intPos As Integer         ' Insertion point where we are
  '  Dim intTransSt As Integer = 0 ' Current translation's start
  '  Dim intTransEn As Integer = 0 ' Current translation's end
  '  Dim intPdeLine As Integer     ' The PDE line number to be searched

  '  Try
  '    ' What is being pressed
  '    Select Case e.KeyChar
  '      Case ChrW(Keys.Escape)
  '        ' Disregard and exit
  '        With Me.tbPde
  '          .Text = ""
  '          .Visible = False
  '        End With
  '      Case ChrW(Keys.Enter)
  '        ' Save the text
  '        With Me.tbPde
  '          strTranslation = .Text
  '          .Text = ""
  '          .Visible = False
  '        End With
  '        ' Remove the last linefeed
  '        If (Strings.Right(strTranslation, 1) = vbLf) Then strTranslation = Strings.Left(strTranslation, strTranslation.Length - 1)
  '        ' Determine where we are
  '        intPos = Me.tbTranslation.SelectionStart
  '        If (GetTransLine(intPos, intLine)) Then
  '          ' Now get the REAL line number by dividing by 2
  '          intPdeLine = intLine \ 2
  '          ' Set the current translation for this line (in this section)
  '          SetPde(intPdeLine, strTranslation)
  '          ' Set the dirty bit
  '          SetDirty(True)
  '          ' Protect the following code from getting interrupted...
  '          bTransBusy = True
  '          ' Get the current start and end position of the translation 
  '          If (GetTransPosition(intPos, intTransSt, intTransEn)) Then
  '            ' Replace the existing translation
  '            With Me.tbTranslation
  '              .SelectionStart = intTransSt
  '              .SelectionLength = intTransEn - intTransSt
  '              .SelectedText = strTranslation
  '            End With
  '          End If
  '          ' Release protection
  '          bTransBusy = False
  '          ' Set insertion point one step further
  '          Me.tbTranslation.SelectionStart = intTransEn + 1
  '        End If
  '    End Select
  '  Catch ex As Exception
  '    ' Show error
  '    HandleErr("frmMain/tbPde_KeyPress error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
  '  End Try
  'End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbGenTitle_Leave
  ' Goal:   Process changes in the title, distributor and source bibliography
  ' History:
  ' 05-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tbGenTitle_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbGenTitle.Leave
    Dim strType As String = "title"
    If (Not DoAddFileDesc(strType, Me.tbGenTitle.Text)) Then Status("Unable to add [" & strType & "]")
    '' Validate
    'If (Not bInit) Then Exit Sub
    '' Make sure changes are processed
    'If (Not AddFileDesc(pdxCurrentFile, "title", Me.tbGenTitle.Text)) Then
    '  Status("Unable to add the title")
    'End If
    '' Set the dirty bit
    'SetDirty(True)
  End Sub
  Private Sub tbGenDistributor_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbGenDistributor.Leave
    Dim strType As String = "distributor"
    If (Not DoAddFileDesc(strType, Me.tbGenDistributor.Text)) Then Status("Unable to add [" & strType & "]")
    '' Validate
    'If (Not bInit) Then Exit Sub
    '' Make sure changes are processed
    'If (Not AddFileDesc(pdxCurrentFile, "distributor", Me.tbGenDistributor.Text)) Then
    '  Status("Unable to add the distributor")
    'End If
    '' Set the dirty bit
    'SetDirty(True)
  End Sub
  Private Sub tbGenSource_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbGenSource.Leave
    Dim strType As String = "bibl"
    If (Not DoAddFileDesc(strType, Me.tbGenSource.Text)) Then Status("Unable to add [" & strType & "]")
    '' Validate
    'If (Not bInit) Then Exit Sub
    '' Make sure changes are processed
    'If (Not AddFileDesc(pdxCurrentFile, "bibl", Me.tbGenSource.Text)) Then
    '  Status("Unable to add the source bibliography")
    'End If
    '' Set the dirty bit
    'SetDirty(True)
  End Sub
  Private Sub tbGenAuthor_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbGenAuthor.TextChanged
    Dim strType As String = "author"
    If (Not DoAddFileDesc(strType, Me.tbGenAuthor.Text)) Then Status("Unable to add [" & strType & "]")
  End Sub
  Private Sub tbGenCreaOrig_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbGenCreaOrig.TextChanged
    Dim strType As String = "original"
    If (Not DoAddFileDesc(strType, Me.tbGenCreaOrig.Text)) Then Status("Unable to add [" & strType & "]")
  End Sub
  Private Sub tbGenCreaManu_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbGenCreaManu.TextChanged
    Dim strType As String = "manuscript"
    If (Not DoAddFileDesc(strType, Me.tbGenCreaManu.Text)) Then Status("Unable to add [" & strType & "]")
  End Sub
  Private Sub tbGenSubType_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbGenSubType.TextChanged
    Dim strType As String = "subtype"
    If (Not DoAddFileDesc(strType, Me.tbGenSubType.Text)) Then Status("Unable to add [" & strType & "]")
  End Sub
  Private Sub tbGenLngId_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbGenLngId.TextChanged
    Dim strType As String = "ident"
    If (Not DoAddFileDesc(strType, Me.tbGenLngId.Text, "language")) Then Status("Unable to add [" & strType & "]")
  End Sub
  Private Sub tbGenLngName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbGenLngName.TextChanged
    Dim strType As String = "name"
    If (Not DoAddFileDesc(strType, Me.tbGenLngName.Text, "language")) Then Status("Unable to add [" & strType & "]")
  End Sub
  Private Sub tbGenEditor_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbGenEditor.TextChanged
    Dim strType As String = "editor"
    If (Not DoAddFileDesc(strType, Me.tbGenEditor.Text)) Then Status("Unable to add [" & strType & "]")
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   dgvTeiMeta_SelectionChanged
  ' Goal:   Process meta-data -- show where we are
  ' History:
  ' 12-08-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub dgvTeiMeta_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvTeiMeta.SelectionChanged
    Dim intSelId As Integer ' Selected node

    Try
      ' Get selected node
      intSelId = objMetaEd.SelectedId
      If (intSelId >= 0) Then
        ' Show the value for this selection

      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/dgvTeiMeta error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub tbTeiMetaValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbTeiMetaValue.TextChanged
    Try
      ' Is something selected?
      If (Me.dgvTeiMeta.SelectedCells.Count = 0) Then
        Status("First select a metadata item")
      Else
        If (Not TeiMetaDataEdit(Me.dgvTeiMeta.SelectedCells.Item(1).Value, Me.tbTeiMetaValue.Text)) Then Status("Unable to edit metadata")
        ' Set the dirty bit
        SetDirty(True)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/tbTeiMetaValue error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub cmdTeiMetaAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTeiMetaAdd.Click
    If (Not TeiMetaDataAdd()) Then Status("Unable to add metadata")
    ' Set the dirty bit
    SetDirty(True)
  End Sub

  Private Sub cmdTeiMetaDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTeiMetaDel.Click
    Try
      ' Is something selected?
      If (Me.dgvTeiMeta.SelectedCells.Count = 0) Then
        Status("First select a metadata item")
      Else
        If (Not TeiMetaDataDel(Me.dgvTeiMeta.SelectedCells.Item(1).Value, _
                               Me.dgvTeiMeta.SelectedCells.Item(2).Value)) Then Status("Unable to delete metadata")
        ' Set the dirty bit
        SetDirty(True)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/cmdTeiMetaDel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbSonarCollectionName_TextChanged (etc)
  ' Goal:   Process changes in the Sonar metadata information
  ' History:
  ' 19-06-2015  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tbSonarCollectionName_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbSonarCollectionName.TextChanged
    Dim strType As String = "CollectionName"
    If (Not TeiMetaDataEdit(strType, Me.tbSonarCollectionName.Text)) Then Status("Unable to edit [" & strType & "]")
  End Sub
  Private Sub tbSonarLicenseCode_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbSonarLicenseCode.TextChanged
    Dim strType As String = "LicenseCode"
    If (Not TeiMetaDataEdit(strType, Me.tbSonarLicenseCode.Text)) Then Status("Unable to edit [" & strType & "]")
  End Sub
  Private Sub tbSonarSourceName_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbSonarSourceName.TextChanged
    Dim strType As String = "SourceName"
    If (Not TeiMetaDataEdit(strType, Me.tbSonarSourceName.Text)) Then Status("Unable to edit [" & strType & "]")
  End Sub
  Private Sub tbSonarTextType_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbSonarTextType.TextChanged
    Dim strType As String = "TextType"
    If (Not TeiMetaDataEdit(strType, Me.tbSonarTextType.Text)) Then Status("Unable to edit [" & strType & "]")
  End Sub
  Private Sub tbSonarAuthorNameOrPseudonym_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbSonarAuthorNameOrPseudonym.TextChanged
    Dim strType As String = "AuthorNameOrPseudonym"
    If (Not TeiMetaDataEdit(strType, Me.tbSonarAuthorNameOrPseudonym.Text)) Then Status("Unable to edit [" & strType & "]")
  End Sub
  Private Sub tbSonarSex_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbSonarSex.TextChanged
    Dim strType As String = "Sex"
    If (Not TeiMetaDataEdit(strType, Me.tbSonarSex.Text)) Then Status("Unable to edit [" & strType & "]")
  End Sub
  Private Sub tbSonarAge_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbSonarAge.TextChanged
    Dim strType As String = "Age"
    If (Not TeiMetaDataEdit(strType, Me.tbSonarAge.Text)) Then Status("Unable to edit [" & strType & "]")
  End Sub
  Private Sub tbSonarCountry_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbSonarCountry.TextChanged
    Dim strType As String = "Country"
    If (Not TeiMetaDataEdit(strType, Me.tbSonarCountry.Text)) Then Status("Unable to edit [" & strType & "]")
  End Sub
  Private Sub tbSonarTown_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbSonarTown.TextChanged
    Dim strType As String = "Town"
    If (Not TeiMetaDataEdit(strType, Me.tbSonarTown.Text)) Then Status("Unable to edit [" & strType & "]")
  End Sub
  Private Sub tbSonarPublished_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbSonarPublished.TextChanged
    Dim strType As String = "Published"
    If (Not TeiMetaDataEdit(strType, Me.tbSonarPublished.Text)) Then Status("Unable to edit [" & strType & "]")
  End Sub
  Private Sub tbSonarPublisher_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbSonarPublisher.TextChanged
    Dim strType As String = "Publisher"
    If (Not TeiMetaDataEdit(strType, Me.tbSonarPublisher.Text)) Then Status("Unable to edit [" & strType & "]")
  End Sub
  Private Sub tbSonarPublicationName_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbSonarPublicationName.TextChanged
    Dim strType As String = "PublicationName"
    If (Not TeiMetaDataEdit(strType, Me.tbSonarPublicationName.Text)) Then Status("Unable to edit [" & strType & "]")
  End Sub
  Private Sub tbSonarPublicationPlace_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbSonarPublicationPlace.TextChanged
    Dim strType As String = "PublicationPlace"
    If (Not TeiMetaDataEdit(strType, Me.tbSonarPublicationPlace.Text)) Then Status("Unable to edit [" & strType & "]")
  End Sub
  Private Sub tbSonarTextKeyword_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbSonarTextKeyword.TextChanged
    Dim strType As String = "TextKeyword"
    If (Not TeiMetaDataEdit(strType, Me.tbSonarTextKeyword.Text)) Then Status("Unable to edit [" & strType & "]")
  End Sub
  Private Sub tbSonarTextDescription_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbSonarTextDescription.TextChanged
    Dim strType As String = "TextDescription"
    If (Not TeiMetaDataEdit(strType, Me.tbSonarTextDescription.Text)) Then Status("Unable to edit [" & strType & "]")
  End Sub
  Private Sub tbSonarTranslated_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbSonarTranslated.TextChanged
    Dim strType As String = "Translated"
    If (Not TeiMetaDataEdit(strType, Me.tbSonarTranslated.Text)) Then Status("Unable to edit [" & strType & "]")
  End Sub
  Private Sub tbSonarOriginalLanguage_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbSonarOriginalLanguage.TextChanged
    Dim strType As String = "OriginalLanguage"
    If (Not TeiMetaDataEdit(strType, Me.tbSonarOriginalLanguage.Text)) Then Status("Unable to edit [" & strType & "]")
  End Sub
  Private Sub tbSonarTranslatorName_TextChanged(sender As System.Object, e As System.EventArgs) Handles tbSonarTranslatorName.TextChanged
    Dim strType As String = "TranslatorName"
    If (Not TeiMetaDataEdit(strType, Me.tbSonarTranslatorName.Text)) Then Status("Unable to edit [" & strType & "]")
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdTeiLoadFile_Click
  ' Goal:   Load header information from a .txt or .csv file into the currently opened file
  ' History:
  ' 12-08-2014  ERK Created
  ' 03-10-2014  ERK Default use the "columns" method
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdTeiLoadFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTeiLoadFile.Click
    Dim strFile As String = ""  ' Filename
    Dim strName As String = ""  ' Name of us
    Dim bChanged As Boolean = False

    Try
      ' Ask for filename
      If (Not GetFileName(Me.OpenFileDialog1, strWorkDir, strFile, "Tab-separated text file (*.txt;*.csv)|*.txt;*.csv")) Then Status("Could not load file") : Exit Sub
      If (TeiMetaDataOneFile(strFile, "columns", bChanged)) Then
        Status("Processed " & strFile)
      Else
        Status("Could not process metadata in " & strFile)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/cmdTeiLoadFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdTeiLoadDir_Click
  ' Goal:   Load header information from a .txt or .csv files and process them in the files in one directory
  ' History:
  ' 14-08-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdTeiLoadDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTeiLoadDir.Click
    Dim strFile As String = ""    ' Filename
    Dim strDir As String = ""     ' Directory with psdx files
    Dim strMethod As String = ""  ' Desired method

    Try
      ' Elicit filename, directory and format
      With frmMetaInfo
        Select Case .ShowDialog
          Case MsgBoxResult.Cancel
            Status("Aborted")
            Exit Sub
        End Select
        ' Get the results
        strFile = .FileName
        strDir = .PsdxDir
        strMethod = .Method
      End With
      '' Ask for filename
      'If (Not GetFileName(Me.OpenFileDialog1, strWorkDir, strFile, "Tab-separated text file (*.txt;*.csv)|*.txt;*.csv")) Then Status("Could not load file") : Exit Sub
      '' Ask for directory
      'If (Not GetDirName(Me.FolderBrowserDialog1, strDir, "Directory that contains the .psdx files", strWorkDir)) Then Status("Could not open directory") : Exit Sub
      '' Find out what the format is
      ' Process all files
      If (TeiMetaDataFiles(strFile, strDir, strMethod)) Then
        Status("Processed metadata for: " & strDir)
      Else
        Status("Could not process metadata in " & strFile)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/cmdTeiLoadDir error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsMetaRead_Click
  ' Goal:   Read the <teiHeader> information from a number of files in one directory
  ' History:
  ' 26-06-2015  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsMetaRead_Click(sender As System.Object, e As System.EventArgs) Handles mnuToolsMetaRead.Click
    Dim strCsv As String = ""   ' csv table
    Dim strHtml As String = ""  ' table in html

    Try
      If (TeiGetCorpusTable(strCsv, strHtml)) Then
        ' Show the table where it can be shown
        Me.wbReport.DocumentText = strHtml
        Status("ready")
      Else
        Status("There was an error")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsMetaRead error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsMetaWrite
  ' Goal:   Write <teiHeader> information to a selected number of files in one or more directories
  ' History:
  ' 26-06-2015  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsMetaWrite_Click(sender As System.Object, e As System.EventArgs) Handles mnuToolsMetaWrite.Click
    Try
      If (TeiImportCorpusTable()) Then
        Status("ready")
      Else
        Status("There was an error")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ mnuToolsMetaWrite error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbGenTitle_TextChanged
  ' Goal:   Make sure changes are processed later?
  ' History:
  ' 05-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tbGenTitle_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbGenTitle.TextChanged
    ' Set the dirty bit
    SetDirty(True)
  End Sub
  Private Sub tbGenSource_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbGenSource.TextChanged
    ' Set the dirty bit
    SetDirty(True)
  End Sub
  Private Sub tbGenDistributor_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbGenDistributor.TextChanged
    ' Set the dirty bit
    SetDirty(True)
  End Sub
  Private Sub tbGenRevision_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbGenRevision.TextChanged
    ' Set the dirty bit
    SetDirty(True)
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbGenRevision_DoubleClick
  ' Goal:   Allow user to add a revision entry
  ' History:
  ' 05-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tbGenRevision_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbGenRevision.DoubleClick
    ' Call the appropriate routine
    AddRevision()
  End Sub
  Private Sub mnuFileRevision_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileRevision.Click
    ' Call the appropriate routine
    AddRevision()
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbGenRevision_DoubleClick
  ' Goal:   Allow user to add a revision entry
  ' History:
  ' 05-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub AddRevision()
    Dim strWho As String      ' Who made the revision
    Dim strWhen As String     ' When the revision was made
    Dim strComment As String  ' The text/comments of the revision

    Try
      ' Validate: a text has to be opened, we need to have been initialised
      If (Not bInit) OrElse (bEdtLoad) Then Exit Sub
      If (pdxCurrentFile Is Nothing) Then Exit Sub
      ' Open the appropriate dialog box
      With frmRevision
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK, Windows.Forms.DialogResult.Yes
            ' Get the who/when/comment variables
            strWho = .Who
            strWhen = .WhenMade
            strComment = .Comment
            ' Store this revision in the PSDX file
            AddRevDesc(pdxCurrentFile, strWho, strWhen, strComment)
            ' Show the adapted revision information on the General tab of the main form
            ShowRevision()
            ' Reset the revision flag
            bRevision = False
          Case Else
            ' No revision added, so exit
            Exit Sub
        End Select
      End With
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/AddRevision error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   ShowRevision
  ' Goal:   Show all the revisions on the General tab of the main form
  ' History:
  ' 05-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub ShowRevision()
    Dim intI As Integer           ' Counter
    Dim ndxRev As Xml.XmlNodeList ' List of revisions
    Dim fntNormal As System.Drawing.Font = New Font("Times New Roman", 10, FontStyle.Regular)
    Dim fntBold As System.Drawing.Font = New Font("Times New Roman", 10, FontStyle.Bold)

    Try
      ' Validate
      If (Not bInit) Then Exit Sub
      If (pdxCurrentFile Is Nothing) Then Exit Sub
      ' Get the text of the revision and split it into lines
      ndxRev = pdxCurrentFile.SelectNodes("//change")
      ' Clear the textbox
      Me.tbGenRevision.Text = ""
      ' Visit all elements (the latest revision is topmost)
      For intI = 0 To ndxRev.Count - 1
        ' Add this element
        AppendToRtb(Me.tbGenRevision, GetAttrValue(ndxRev(intI), "who") & vbTab, Color.DarkGreen, fntNormal)
        AppendToRtb(Me.tbGenRevision, GetAttrValue(ndxRev(intI), "when") & vbTab, Color.DarkBlue, fntNormal)
        AppendToRtb(Me.tbGenRevision, GetAttrValue(ndxRev(intI), "comment") & vbCrLf, Color.Black, fntNormal)
      Next intI
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ShowRevision error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbAddNewTrn_DragDrop
  ' Goal:   Handle the DragDrop event
  ' History:
  ' 25-10-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tbAddNewTrn_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles tbAddNewTrn.DragDrop
    Try
      If (e.Effect & DragDropEffects.Copy) Then
        ' Replace the content of the richtextbox
        Me.tbAddNewTrn.Text = e.Data.GetData(DataFormats.Text).ToString()
        ' Add this translation to the existing one
        AddTranslation()
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/tbAddNewTrn_DragDrop error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbAddNewTrn_DragEnter
  ' Goal:   Make sure copy works when dragenter starts
  ' History:
  ' 25-10-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tbAddNewTrn_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles tbAddNewTrn.DragEnter
    Try
      If (e.Data.GetDataPresent(DataFormats.Text)) Then
        ' Only allow COPYING, not MOVING
        e.Effect = DragDropEffects.Copy
      Else
        ' There should not be any effect...
        e.Effect = DragDropEffects.None
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/tbAddNewTrn_DragEnter error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub tbAddNewTrn_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbAddNewTrn.Enter
    Try
      ' Clear the textbox if necessary
      If (InStr(Me.tbAddNewTrn.Text, "Enter translation") > 0) Then
        Me.tbAddNewTrn.Text = ""
      End If
      'Me.tbAddNewTrn.Text = ""
      ' Change the background color
      Me.tbAddNewTrn.BackColor = Color.LightBlue
      ' Show what needs to be done
      Status("Add translation")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/tbAddNewTrn_Enter error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub tbAddNewTrn_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbAddNewTrn.Leave
    Try
      ' Clear the textbox
      'Me.tbAddNewTrn.Text = "<Enter translation here>"
      ' Change the background color
      Me.tbAddNewTrn.BackColor = Color.White
      ' Clear status
      Status("")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/tbAddNewTrn_Leave error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbAddNewTrn_KeyPress
  ' Goal:   When user presses "Enter" here, the resulting change is processed
  ' History:
  ' 25-10-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tbAddNewTrn_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbAddNewTrn.KeyPress
    Try
      ' Does user press "ENTER"?
      If (e.KeyChar = ChrW(Keys.Enter)) Then
        ' Add the translation
        AddTranslation()
      End If
      ' Add this text as translation
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/tbAddNewTrn_KeyPress error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbAddNewTrn_KeyPress
  ' Goal:   When AddTranslation presses "Enter" here, the resulting change is processed
  ' History:
  ' 25-10-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub AddTranslation()
    Dim strText As String   ' Text 
    Dim intPtc As Integer   ' Percentage

    Try
      ' Get the text as has been typed
      strText = MyTrim(Me.tbAddNewTrn.Text)
      ' Change CRLF in translation into spaces
      strText = strText.Replace(vbCrLf, " ")
      strText = strText.Replace(vbLf, " ")
      strText = strText.Replace(vbCr, " ")
      ' Some more trimming
      strText = Trim(strText)
      ' Check if anything is left
      If (strText = "") Then
        ' We don't want to add EMPTY translations
        Status("I won't add an empty translation at " & intLastTransLine + 1)
        Exit Sub
      End If
      ' Try to process it
      If (Not AddLastTrans(strText)) Then
        ' Show error
        Status("Could not process new translation")
        Exit Sub
      End If
      ' Make sure the "dirty" flag is set
      SetDirty(True)
      ' Get percentage
      intPtc = (intLastTransLine * 100) \ pdxList.Count
      ' Show success
      Status("Translation of line " & intLastTransLine & "/" & pdxList.Count & _
             " has been added " & intPtc & "%")
      ' Try go to the next line
      TransNext(True)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/AddTranslation error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   InitEditors
  ' Goal:   Initialise any editors with datagridviews and other controls
  '           that are bound to the dataset
  ' History:
  ' 16-10-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitEditors() As Boolean
    Try
      ' Do away with possible previous handlers
      DgvClear(objResEd)    ' Results database Editor
      ' Initialise the results editor (ResEd)
      InitResultsEditor()
      ' Return success
      Return True
    Catch ex As Exception
      ' Show user that there is an error
      HandleErr("frmMain/InitEditors error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitRefChain
  ' Goal:   Initialise editors used for tab page [RefWalk]
  ' History:
  ' 01-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitRefChain() As Boolean
    Try
      ' Do away with possible previous handlers
      DgvClear(objRefChainEd)   ' editor for [Chain] in [tdlRefChain]
      DgvClear(objRefItemEd)    ' editor for [Item] in [tdlRefChain]
      ' Initialise the editor for [Chain] in [tdlRefChain]
      InitRefChainEditor()
      ' Initialise the editor for [Item] in [tdlRefChain]
      InitRefItemEditor()
      ' Return success
      Return True
    Catch ex As Exception
      ' Show user that there is an error
      HandleErr("frmMain/InitRefChain error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitDbaseEditor
  ' Goal:   Initialise the editor for the results database
  ' History:
  ' 23-09-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitResultsEditor() As Boolean
    Dim intI As Integer '  Counter
    Dim lbThis As Label   ' Feature label
    Dim tbThis As TextBox ' Feature textbox

    Try
      ' Set up bindingsource for the [cboResStatus]
      With Me.bsResStatus
        .DataSource = tdlSettings
        .DataMember = "Status"
      End With
      ' Set up [cboResStatus]
      With Me.cboResStatus
        .DataSource = Me.bsResStatus
        .DisplayMember = "Name"
        .ValueMember = "Name"
      End With
      ' Initialise the Results Database DGV handle
      objResEd = New DgvHandle
      With objResEd
        .Init(Me, tdlResults, "Result", "ResId", "ResId", "ResId;Period;TextId;forestId;Cat;Select;Status", "", _
              "", "", Me.dgvDbResult, Nothing)
        .BindControl(Me.tbResNum, "ResId", "number")
        .BindControl(Me.tbResPeriod, "Period", "textbox")
        .BindControl(Me.tbResTextId, "TextId", "textbox")
        .BindControl(Me.tbResForestId, "forestId", "number")
        .BindControl(Me.tbResCat, "Cat", "textbox")
        ' The [Status] attribute is bound, but not completely...
        .BindControl(Me.cboResStatus, "Status", "combo")
        ' The following bindings WERE here previously, but they are now to be handled by a different section...
        '.BindControl(Me.tbResFile, "File", "textbox")
        '.BindControl(Me.tbResLoc, "Search", "textbox")
        '.BindControl(Me.tbResEtreeId, "eTreeId", "number")
        '.BindControl(Me.wbResText, "Text", "browser")
        '.BindControl(Me.tbResPsd, "Psd", "richtext")
        '.BindControl(Me.tbResTrans, "Pde", "richtext")
        '.BindControl(Me.tbResNotes, "Notes", "richtext")
        ' Set the parent table for the [Results] editor
        .ParentTable = ""
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Switch off multiselect
      Me.dgvDbResult.MultiSelect = False
      ' Initialise the labels for features (if present
      For intI = 1 To intNumFeaturesAllowed
        ' Get the label and the textbox controls
        lbThis = FindControl(Me.gbFeat, "lbDbFeat" & intI)
        tbThis = FindControl(Me.gbFeat, "tbDbFeat" & intI)
        ' Check
        If (Not lbThis Is Nothing) Then
          ' Make label invisible
          lbThis.Visible = False
        End If
        ' Check
        If (Not tbThis Is Nothing) Then
          tbThis.Visible = False
        End If
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitResultsEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitTeiMetaEditor
  ' Goal:   Initialise the editor for the TEI metadata
  ' History:
  ' 11-08-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function InitTeiMetaEditor() As Boolean
    Try
      ' Initialise the DGV handle
      objMetaEd = New DgvHandle
      With objMetaEd
        .Init(Me, tdlMeta, "meta", "metaId", "name", "name;value", "", _
              "", "", Me.dgvTeiMeta, Nothing)
        .BindControl(Me.tbTeiMetaValue, "value", "richtext")
        ' Set the parent table for the [meta] editor
        .ParentTable = "metadata"
        Me.tbTeiMetaValue.BackColor = Color.LightBlue
        Me.dgvTeiMeta.Columns(0).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Switch off multiselect
      Me.dgvDbResult.MultiSelect = False
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/InitTeiMetaEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   InitDepEditor
  ' Goal:   Initialise the dependency editor
  ' History:
  ' 04-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitDepEditor() As Boolean
    Dim intI As Integer '  Counter
    Dim lbThis As Label   ' Feature label
    Dim tbThis As TextBox ' Feature textbox

    Try
      ' Initialise the dependency editor's DGV handle
      objDepEd = New DgvHandle
      With objDepEd
        .Init(Me, tdlDependency, "Dep", "DepId", "forestId ASC, Id ASC", "forestId;Id;Vern;Lemma;Cpos;Pos;Hd;Rel", "", _
              "", "", Me.dgvDep, Nothing)
        .BindControl(Me.tbDepForestId, "forestId", "number")
        .BindControl(Me.tbDepId, "Id", "number")
        .BindControl(Me.tbDepVern, "Vern", "textbox")
        .BindControl(Me.tbDepLemma, "Lemma", "textbox")
        .BindControl(Me.tbDepCpos, "Cpos", "textbox")
        .BindControl(Me.tbDepPOS, "Pos", "textbox")
        .BindControl(Me.tbDepHead, "Hd", "textbox")
        .BindControl(Me.tbDepRel, "Rel", "textbox")
        ' Set the parent table for the [Dependency] editor
        .ParentTable = ""
        ' Find out where we are
        ShowCurrentDep()
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Switch off multiselect
      Me.dgvDep.MultiSelect = False
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitDepEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   dgvDep_SelectionChanged
  ' Goal:   Update whatever is needed when selection changes
  ' History:
  ' 09-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub dgvDep_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvDep.SelectionChanged
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim strGloss As String    ' Gloss
    Dim intDepId As Integer   ' selected dep ID
    Dim intForId As Integer   ' forestId of currently selected dependency item
    Dim intEtreeId As Integer ' ID of this constituent
    Dim ndxFor As XmlNode = Nothing   ' Current forest
    Dim ndxThis As XmlNode = Nothing  ' Current constituen
    Dim ndxSeg As XmlNode = Nothing   ' Segment
    Dim bDepSave As Boolean = bDepChg

    Try
      ' Validate
      If (objDepEd Is Nothing) Then Exit Sub
      ' Get ID
      intDepId = objDepEd.SelectedId
      If (intDepId < 0) Then Exit Sub
      ' Get the <forestId> 
      dtrFound = tdlDependency.Tables("Dep").Select("DepId = " & intDepId)
      If (dtrFound.Length > 0) Then
        ' Get the forest id
        intForId = CInt(dtrFound(0).Item("forestId").ToString)
        intEtreeId = CInt(dtrFound(0).Item("EtreeId").ToString)
        ' Make sure changes are not signalled
        bDepChg = True
        ' Set forest to default: current forest
        ' Update whatever is needed
        If (DepSelChange(intForId, ndxFor)) Then
          ' Adapt the messages
          ndxSeg = ndxFor.SelectSingleNode("./child::div[@lang='org']/child::seg")
          If (ndxSeg IsNot Nothing) Then
            Me.tbDepForestOrg.Text = ndxSeg.InnerText
          End If
        End If
        ndxSeg = ndxFor.SelectSingleNode("./child::div[@lang='" & strCurrentTransLang & "']/child::seg")
        If (ndxSeg IsNot Nothing) Then
          Me.tbDepForestEng.Text = ndxSeg.InnerText
        End If
        ndxThis = ndxFor.SelectSingleNode("./descendant::eTree[@Id = " & intEtreeId & "]")
        If (ndxThis IsNot Nothing) Then
          strGloss = GetFeature(ndxThis, "M", "mrph")
          Me.tbDepMorph.Text = strGloss
        End If
        bDepChg = bDepSave
      End If

    Catch ex As Exception
      ' Warn the user
      HandleErr("dgvDep_SelChg error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   cmdDepAddWordBelow_Click
  ' Goal:   Update whatever is needed when selection changes
  ' History:
  ' 16-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdDepAddWordBelow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDepAddWordBelow.Click
    Try
      If (DepAddWordBelow()) Then
        ' Make sure saving is flagged
        SetDirty(True)
        ' Tell we are okay
        Status("Ready")
      Else
        Status("Could not complete joining")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/DepAddWordBelow error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   cmdDepInsertWordBelow_Click
  ' Goal:   Insert a word below the currently selected one in the Dep Parser view
  ' History:
  ' 28-12-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdDepInsertWordBelow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDepInsertWordBelow.Click
    Try
      If (DepInsertWordBelow()) Then
        ' Make sure saving is flagged
        SetDirty(True)
        ' Tell we are okay
        Status("Ready")
      Else
        Status("Could not complete insertion")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/DepInsertWordBelow error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   cmdDepJoinWordAbove_Click
  ' Goal:   Update whatever is needed when selection changes
  ' History:
  ' 10-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdDepJoinWordAbove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDepJoinWordAbove.Click
    Try
      If (DepJoinWordAbove()) Then
        ' Make sure saving is flagged
        SetDirty(True)
        ' Tell we are okay
        Status("Ready")
      Else
        Status("Could not complete joining")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/DepJoinWordAbove error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   cmdDepSplitClauseBelow_Click
  ' Goal:   Update whatever is needed when selection changes
  ' History:
  ' 10-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdDepSplitClauseBelow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDepSplitClauseBelow.Click
    Dim ndxWork As XmlNode  ' Node
    Dim intI As Integer   ' Counter
    Dim intPtc As Integer ' Percentage

    Try
      ' Validate
      If (InStr(Me.tbDepForestEng.Text, vbLf) = 0) OrElse (InStr(Me.tbDepForestOrg.Text, vbLf) = 0) Then
        ' Warn user
        Status("You also need to split the text in the [Eng] and the [Org] textboxes")
        Exit Sub
      End If
      If (DepSplitLineBelow()) Then
        '' Re-analyze
        '' Visit all the <forest> lines
        'For intI = 0 To pdxList.Count - 1
        '  ' Where are we?
        '  intPtc = (intI + 1) * 100 \ pdxList.Count
        '  Status("Re-analyzing " & intPtc & "%", intPtc)
        '  ' Re-analyze this line
        '  If (eTreeSentence(pdxList(intI), ndxWork)) Then
        '    ' Changes have been made!
        '  End If
        'Next intI
        '' Adapt the @Id values
        'AdaptEtreeId(1)
        '' Make sure changes are process in the dependency
        'tdlDependency.AcceptChanges()
        ' Make sure saving is flagged
        SetDirty(True)
        ' Tell we are okay
        Status("Ready")
      Else
        Status("Could not complete splitting")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/cmdDepSplitClauseBelow error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   cmdDepDelWord_Click
  ' Goal:   Delete the currently selected word
  ' History:
  ' 16-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdDepDelWord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDepDelWord.Click
    Try

    Catch ex As Exception

    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   tbDepVern_TextChanged (and others below)
  ' Goal:   Update whatever is needed when selection changes
  ' History:
  ' 10-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub tbDepVern_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDepVern.TextChanged
    If (Not bDepChg) Then DepChange("Vern", Me.tbDepVern.Text) : SetDirty(True)
  End Sub
  Private Sub tbDepLemma_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDepLemma.TextChanged
    If (Not bDepChg) Then DepChange("Lemma", Me.tbDepLemma.Text) : SetDirty(True)
  End Sub
  Private Sub tbDepCpos_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDepCpos.TextChanged
    If (Not bDepChg) Then DepChange("Cpos", Me.tbDepCpos.Text) : SetDirty(True)
  End Sub
  Private Sub tbDepPOS_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDepPOS.TextChanged
    If (Not bDepChg) Then DepChange("Pos", Me.tbDepPOS.Text) : SetDirty(True)
  End Sub
  Private Sub tbDepId_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDepId.TextChanged
    If (Not bDepChg) Then DepChange("Id", Me.tbDepId.Text) : SetDirty(True)
  End Sub
  Private Sub tbDepHead_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDepHead.TextChanged
    If (Not bDepChg) Then DepChange("Head", Me.tbDepHead.Text) : SetDirty(True)
  End Sub
  Private Sub tbDepRel_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDepRel.TextChanged
    If (Not bDepChg) Then DepChange("Rel", Me.tbDepRel.Text) : SetDirty(True)
  End Sub
  Private Sub tbDepForestOrg_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDepForestOrg.TextChanged
    If (Not bDepChg) Then DepChange("Org", Me.tbDepForestOrg.Text) : SetDirty(True)
  End Sub
  Private Sub tbDepForestEng_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDepForestEng.TextChanged
    ' Can we  go on?
    If (Not bDepChg) Then
      DepChange("Eng", Me.tbDepForestEng.Text)
      SetDirty(True)
    End If
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   dgvDbResult_BeforeSelChanged
  ' Goal:   Actions to take BEFORE selection changes in [dgvDbResult]
  ' History:
  ' 23-09-2013  ERK Created
  ' 26-09-2013  ERK Shifted to XmlDocument
  ' ------------------------------------------------------------------------------------
  Private Sub dgvDbResult_BeforeSelChanged(ByVal intSelId As Integer)
    ' Dim dtrFound() As DataRow ' Result of SELECT
    Dim strNote As String     ' Value of current note

    Try
      ' Validate
      If (Not bInit) OrElse (intSelId < 0) Then Exit Sub
      ' Check if changes took place
      If (bResRecord) Then
        ' Validate
        If (ndxCurrentRes IsNot Nothing) Then
          ' Do we need to change the status to "Changed"?
          If (bResStatus) AndAlso (GetTableSetting(tdlSettings, "StatusAutoChanged") = "True") Then
            ' Set the status for this item
            AddXmlAttribute(pdxCrpDbase, ndxCurrentRes, "Status", STATUS_CHANGED)
          End If
          ' Need to add something to the notes?
          If (GetTableSetting(tdlSettings, "StatusAutoNotes") = "True") Then
            ' Get value of current note
            If (ndxCurrentRes.Attributes("Notes") Is Nothing) Then
              strNote = ""
            Else
              strNote = ndxCurrentRes.Attributes("Notes").Value
            End If
            ' Add a note to the notes
            AddXmlAttribute(pdxCrpDbase, ndxCurrentRes, "Notes", _
                            "Changed on [" & Format(Now, "G") & "] by [" & strUserName & "]" & vbCrLf & strNote)
          End If

        End If
        'dtrFound = tdlResults.Tables("Result").Select("ResId = " & intSelId)
        'If (dtrFound.Length > 0) Then
        '  ' Do we need to change the status to "Changed"?
        '  If (bResStatus) AndAlso (GetTableSetting(tdlSettings, "StatusAutoChanged") = "True") Then
        '    ' Set the status for this item
        '    dtrFound(0).Item("Status") = STATUS_CHANGED
        '  End If
        '  ' Need to add something to the notes?
        '  If (GetTableSetting(tdlSettings, "StatusAutoNotes") = "True") Then
        '    ' Add a note to the notes
        '    dtrFound(0).Item("Notes") = "Changed on [" & Format(Now, "G") & "] by [" & strUserName & "]" & vbCrLf & dtrFound(0).Item("Notes").ToString
        '  End If
        'End If
        ' Reset changes
        bResRecord = False : bResStatus = False
        ' Make sure display is refreshed
        objResEd.Refresh()
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/dgvDbResult_BeforeSelChanged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   InitRefItemEditor
  ' Goal:   Initialise the editor for the [Item] table in [tdlRefChain]
  ' History:
  ' 23-09-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitRefItemEditor() As Boolean
    Try
      ' Initialise the Results Database DGV handle
      objRefItemEd = New DgvHandle
      With objRefItemEd
        .Init(Me, tdlRefChain, "Item", "ItemId", "eTreeId ASC", "forestId;ItemId;ChainId;GrRole;Node", "", _
              "", "", Me.dgvRefItem, Nothing)
        .BindControl(Me.tbRefItemChainId, "ChainId", "number")
        .BindControl(Me.tbRefItemItemId, "ItemId", "number")
        .BindControl(Me.tbRefItemEtreeId, "eTreeId", "number")
        .BindControl(Me.tbRefItemForestId, "forestId", "number")
        .BindControl(Me.tbRefItemIPmatId, "IPmatId", "number")
        .BindControl(Me.tbRefItemIPclsId, "IPclsId", "number")
        .BindControl(Me.tbRefItemForestId, "forestId", "number")
        .BindControl(Me.tbRefItemIPnum, "IPnum", "number")
        .BindControl(Me.tbRefItemLoc, "Loc", "textbox")
        .BindControl(Me.tbRefItemIPclsLb, "IPclsLb", "textbox")
        .BindControl(Me.tbRefItemRefType, "RefType", "textbox")
        .BindControl(Me.tbRefItemSyntax, "Syntax", "textbox")
        .BindControl(Me.tbRefItemNPtype, "NPtype", "textbox")
        .BindControl(Me.tbRefItemGrRole, "GrRole", "textbox")
        .BindControl(Me.tbRefItemPGN, "PGN", "textbox")
        .BindControl(Me.tbRefItemNode, "Node", "richtext")
        ' Set the parent table for the [Results] editor
        .ParentTable = "Chain"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Switch off multiselect
      Me.dgvDbResult.MultiSelect = False
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitRefItemEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitRefChainEditor
  ' Goal:   Initialise the editor for the [Item] table in [tdlRefChain]
  ' History:
  ' 23-09-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitRefChainEditor() As Boolean
    Try
      ' Initialise the Results Database DGV handle
      objRefChainEd = New DgvHandle
      With objRefChainEd
        .Init(Me, tdlRefChain, "Chain", "ChainId", "Len DESC, ChainId ASC", "ChainId;PGN;Len;Animacy", "", _
              "", "", Me.dgvRefChain, Nothing)
        .BindControl(Me.tbRefChainId, "ChainId", "number")
        .BindControl(Me.tbRefChainPGN, "PGN", "textbox")
        .BindControl(Me.tbRefChainAnimacy, "Animacy", "textbox")
        .BindControl(Me.tbRefChainProAll, "ProAll", "textbox")
        .BindControl(Me.tbRefChainProSbj, "ProSbj", "textbox")
        .BindControl(Me.tbRefChainNmeSbj, "NmeSbj", "textbox")
        .BindControl(Me.tbRefChainSbjCount, "SbjCount", "number")
        .BindControl(Me.tbRefChainSbjIsProt, "SbjIsProt", "number")
        .BindControl(Me.tbRefChainSbjSwProtX, "SbjSwProtX", "number")
        .BindControl(Me.tbRefChainSbjZero, "SbjZero", "number")
        ' Set the parent table for the [Results] editor
        .ParentTable = "ChainList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Switch off multiselect
      Me.dgvDbResult.MultiSelect = False
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitRefChainEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbAddNewTrn_TextChanged
  ' Goal:   What to do when the text has changed
  ' History:
  ' 25-10-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tbAddNewTrn_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbAddNewTrn.TextChanged
    ' Accept changes in the correct translation place
    ' Debug.Print(Me.tbAddNewTrn.Text)
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   dgvDbResult_DoubleClick
  ' Goal:   Open the file belonging to this item, and go to the correct line
  ' History:
  ' 21-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub dgvDbResult_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvDbResult.DoubleClick
    Dim dtrThis As DataRow        ' Currently selected datarow
    Dim intLine As Integer        ' Line number to go to
    Dim strFile As String         ' File to open

    Try
      ' Validate
      If (objResEd Is Nothing) Then Exit Sub
      ' Check whether something is actually selected
      If (ndxCurrentRes Is Nothing) Then Status("First properly select one item") : Exit Sub
      If (ndxCurrentRes.Attributes("forestId") Is Nothing OrElse _
          ndxCurrentRes.Attributes("File") Is Nothing) Then Status("Sorry, cannot select this item") : Exit Sub
      ' Retrieve the forestId and the File
      intLine = ndxCurrentRes.Attributes("forestId").Value
      strFile = ndxCurrentRes.Attributes("File").Value
      '' Get the datarow of the currently shown set
      'dtrThis = objResEd.SelRow
      '' Validate
      'If (dtrThis Is Nothing) Then Exit Sub
      '' Get the necessary information
      'With dtrThis
      '  intLine = .Item("forestId")
      '  strFile = .Item("File")
      'End With
      ' Possibly adapt the file
      If (Not GetFullPsdxName(strFile)) Then
        ' Tell the user what is going on
        Status("Could not locate file [" & strFile & "] in directory [" & strWorkDir & "] or [" & Me.tbDbSrcDir.Text & "]")
        Exit Sub
      End If
      ' Should we open it?
      If (strFile <> strCurrentFile) Then
        ' Open the appropriate file
        LoadPsdx(strFile, False, intLine)
      End If
      ' Go to the appropriate line number in this file
      EditGotoLine(intLine)
      ' Switch over to the correct tabpage
      Me.TabControl1.SelectedTab = Me.tpEdit
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/dgvDbResult_DoubleClick error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   dgvDbResult_SelectionChanged
  ' Goal:   Adapt the values of the features that are being shown
  ' History:
  ' 21-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub dgvDbResult_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvDbResult.SelectionChanged
    Dim dtrThis As DataRow        ' Currently selected datarow
    'Dim dtrChildren() As DataRow  ' Children of this datarow
    Dim intI As Integer           ' Counter
    Dim lbThis As Label           ' Feature label
    Dim tbThis As TextBox         ' Feature textbox
    Dim ndxFeat As XmlNodeList    ' List of features
    Dim strStatus As String = ""  ' Status that is selected

    Try
      ' Validate
      If (objResEd Is Nothing) Then Exit Sub
      ' Make sure we are not right now still selecting
      If (objResEd.IsSelecting) OrElse (bResSel) Then Exit Sub
      ' Get the datarow of the currently shown set
      dtrThis = objResEd.SelRow
      ' Validate
      If (dtrThis Is Nothing) Then Exit Sub
      ' Perform any actions before changes
      dgvDbResult_BeforeSelChanged(intCurrentResId)
      intCurrentResId = objResEd.SelectedId
      ' Has something actually been selected?
      If (intCurrentResId < 0) Then Exit Sub
      ' Indicate selection changes
      bResSel = True
      ' Show the newest values of [Text, Psd, Pde, Status, Notes]
      ndxCurrentRes = pdxCrpDbase.SelectSingleNode("./descendant::Result[@ResId = " & intCurrentResId & "]")
      If (ndxCurrentRes Is Nothing) Then Status("Problem selecting Result number " & intCurrentResId) : bResSel = False : Exit Sub
      GetAttrValue(ndxCurrentRes, "Text")
      ' Debug.Print(Me.wbResText.Visible)
      ' wait until ready state is okay
      With Me.wbResText
        .Stop()
        While (.IsBusy)
          Application.DoEvents()
        End While
      End With
      Me.wbResText.DocumentText = "<html><body>" & GetAttrValue(ndxCurrentRes, "Text") & "</body></html>"
      Me.tbResPsd.Text = GetAttrValue(ndxCurrentRes, "Psd")
      Me.tbResTrans.Text = GetAttrValue(ndxCurrentRes, "Pde")
      Me.tbResNotes.Text = GetAttrValue(ndxCurrentRes, "Notes")
      Me.tbResEtreeId.Text = GetAttrValue(ndxCurrentRes, "eTreeId")
      Me.tbResFile.Text = GetAttrValue(ndxCurrentRes, "File")
      Me.tbResLoc.Text = GetAttrValue(ndxCurrentRes, "Search")
      ' Get selected status
      If (ndxCurrentRes.Attributes("Status") Is Nothing) Then
        strStatus = ""
      Else
        strStatus = ndxCurrentRes.Attributes("Status").Value
      End If
      Me.cboResStatus.SelectedValue = strStatus
      'Stop
      'Me.cboResStatus.SelectedItem = strStatus
      'Me.cboResStatus.SelectedText = strStatus
      ' Show the feature values belonging to this record
      ndxFeat = ndxCurrentRes.SelectNodes("./child::Feature")
      ' Adapt features allowed
      If (intNumFeaturesAllowed < ndxFeat.Count) Then
        ReDim Preserve arFtLabel(0 To ndxFeat.Count - 1)
        ReDim Preserve arFtTxtBox(0 To ndxFeat.Count - 1)
      End If
      ' Possibly make room for more features on the screen...
      For intI = intNumFeaturesAllowed To ndxFeat.Count - 1
        ' Create a new label
        arFtLabel(intI) = New Label
        Me.SplitContainer14.Panel2.Controls.Add(arFtLabel(intI))
        With arFtLabel(intI)
          .Location = New System.Drawing.Point(7, 6 + 26 * (intI - 12))
          .Size = New System.Drawing.Size(83, 13)
          .Name = "lbDbFeat" & intI + 1
        End With
        ' Create a new textbox
        arFtTxtBox(intI) = New TextBox
        Me.SplitContainer14.Panel2.Controls.Add(arFtTxtBox(intI))
        With arFtTxtBox(intI)
          .Location = New System.Drawing.Point(96, 4 + 26 * (intI - 12))
          ' .Size = New System.Drawing.Size(262, 20)
          .Size = Me.tbDbFeat13.Size
          .Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
          .Name = "tbDbFeat" & intI + 1
        End With
      Next intI
      ' Adapt features allowed
      If (intNumFeaturesAllowed < ndxFeat.Count) Then
        intNumFeaturesAllowed = ndxFeat.Count
      End If
      ' Take care of visualisation
      For intI = 0 To ndxFeat.Count - 1
        ' Get the label and the textbox controls
        lbThis = FindControl(Me.gbFeat, "lbDbFeat" & intI + 1)
        tbThis = FindControl(Me.gbFeat, "tbDbFeat" & intI + 1)
        ' Check
        If (Not lbThis Is Nothing) Then
          ' Make label visible
          lbThis.Visible = True
          lbThis.Text = ndxFeat(intI).Attributes("Name").Value
        End If
        ' Check
        If (Not tbThis Is Nothing) Then
          tbThis.Visible = True
          tbThis.Text = ndxFeat(intI).Attributes("Value").Value
        End If
      Next intI


      ' Check if [tbResNotes] is visible
      If (Not Me.tbResNotes.Visible) Then
        ' Check if there is any content in the notes
        If (Trim(Me.tbResNotes.Text) = "" OrElse Trim(Me.tbResNotes.Text) = "-") Then
          ' There are no notes in there... warn the user
          Me.tpCrpNotes.ImageIndex = 0
        Else
          Me.tpCrpNotes.ImageIndex = 1
        End If
      Else
        Me.tpCrpNotes.ImageIndex = 1
      End If
      ' Indicate selection is ready
      bResSel = False
      ' Implement current zoom factor
      If (Me.wbResText.Document.Body IsNot Nothing) Then
        ' Timing problem...
        Application.DoEvents()
        ' Now!
        Me.wbResText.Document.Body.Style = "zoom: " & intZoom & "%"
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/dgvDbResult error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   SetResDirty
  ' Goal:   Set or reset the dirty flag for the results database
  ' History:
  ' 21-12-2010  ERK Created
  ' 23-09-2013  ERK Added [bSetChanged] option
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub SetResDirty(ByVal bSel As Boolean, Optional ByVal bSetChanged As Boolean = False)
    Try
      ' Do not set "dirty" when we are selecting
      If (bResSel) Then Exit Sub
      ' Okay, we are not selecting...
      bResDirty = bSel
      ' Has something more changed?
      If (bSetChanged) Then
        ' Warn record changes
        bResRecord = True : bResStatus = True
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SetResDirty error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoFeatUrl
  ' Goal:   If the text starts with http:// or https:// then treat this as URL
  ' History:
  ' 22-12-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub DoFeatUrl(ByRef tbThis As TextBox)
    Dim strValue As String      ' The text value of the box
    Dim strHtml As String       ' What to show

    Try
      ' Validate
      If (bResSel) OrElse (objResEd Is Nothing) Then Exit Sub
      ' Get the text value
      strValue = tbThis.Text
      ' Check if this is an URL
      If (DoLike(strValue, "BT*|OED*|MED*")) Then
        ' This is a dictionary entry number - go for it
        strHtml = GetDictEntryHtml("", strValue)
        ' Evaluate the result
        If (strHtml <> "") Then
          ' Now show it
          Me.tabCorpus.SelectedTab = Me.tpWeb
          Me.wbCorpusFeature.DocumentText = strHtml
        End If
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/DoFeatUrl error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoFeatChange
  ' Goal:   Process the change of the feature value
  ' History:
  ' 21-12-2010  ERK Created
  ' 26-09-2013  ERK Adapted to working with XmlDocument and [ndxCurrentRes]
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub DoFeatChange(ByVal intNum As Integer, ByRef tbThis As TextBox)
    Dim dtrThis As DataRow        ' Currently selected datarow
    'Dim dtrChildren() As DataRow  ' Children of this datarow
    Dim ndxFeat As XmlNode        ' Feature children
    Dim strFname As String = ""   ' Name of this feature
    Dim strValue As String        ' The text value of the box
    Dim strFilter As String = ""  ' Filter string
    Dim intSelPos As Integer      ' Selection position
    Dim intSelLen As Integer      ' Length of selection
    Dim lbThis As Label           ' Label belonging to the currently selected feature

    Try
      ' Validate
      If (bResSel) OrElse (objResEd Is Nothing) Then Exit Sub
      ' Get the text value
      strValue = tbThis.Text
      ' Get the feature name
      lbThis = FindControl(Me.gbFeat, "lbDbFeat" & intNum)
      strFname = lbThis.Text
      ' Assume that we have the current [ndxCurrentRes] established in dgvDbResult.SelectionChanged
      If (ndxCurrentRes Is Nothing) Then Logging("frmMain/DoFeatChange error: [ndxCurrentRes] is not established") : Exit Sub
      ' Get the correct feature node
      ndxFeat = ndxCurrentRes.SelectSingleNode("./child::Feature[@Name = '" & strFname & "']")
      ' Change the value
      ndxFeat.Attributes("Value").Value = strValue
      ' Check whether we need to adapt [Select]
      If (Me.cboDbSelect.SelectedItem <> "") Then
        If (strFname = Me.cboDbSelect.SelectedItem) Then
          ' Get the datarow of the currently shown set
          dtrThis = objResEd.SelRow
          ' Validate
          If (dtrThis Is Nothing) Then Exit Sub
          ' Check if the filter is set
          strFilter = objResEd.Filter
          If (strFilter = "") Then
            ' We need to adapt the value
            dtrThis.BeginEdit()
            dtrThis.Item("Select") = strValue
            ' Start timer
            Me.tmeFeatChange.Enabled = True
          Else
            ' Switch off the filter and note the selection position
            intSelPos = tbThis.SelectionStart : intSelLen = tbThis.SelectionLength
            objResEd.Filter = ""
            ' We need to adapt the value
            dtrThis.BeginEdit()
            dtrThis.Item("Select") = strValue
            ' Put the filter back on
            objResEd.Filter = strFilter
            ' Put the selection length back
            tbThis.SelectionStart = intSelPos : tbThis.SelectionLength = intSelLen
          End If
          ' Debug.Print(dtrThis.Item("ResId").ToString)
        Else
          ' Find out 
        End If
      End If

      '' Check if this datarow has any children that contain features
      'dtrChildren = dtrThis.GetChildRows("Result_Feature")
      'If (dtrChildren.Length < intNum) Then Exit Sub
      ' Process the correct child
      'With dtrChildren(intNum - 1)
      '  .Item("Value") = strValue
      'End With
      '' Check whether we need to adapt [Select]
      'If (Me.cboDbSelect.SelectedItem <> "") Then
      '  If (dtrChildren(intNum - 1).Item("Name") = Me.cboDbSelect.SelectedItem) Then
      '    ' Check if the filter is set
      '    strFilter = objResEd.Filter
      '    If (strFilter = "") Then
      '      ' We need to adapt the value
      '      dtrThis.BeginEdit()
      '      dtrThis.Item("Select") = strValue
      '      ' Start timer
      '      Me.tmeFeatChange.Enabled = True
      '    Else
      '      ' Switch off the filter and note the selection position
      '      intSelPos = tbThis.SelectionStart : intSelLen = tbThis.SelectionLength
      '      objResEd.Filter = ""
      '      ' We need to adapt the value
      '      dtrThis.BeginEdit()
      '      dtrThis.Item("Select") = strValue
      '      ' Put the filter back on
      '      objResEd.Filter = strFilter
      '      ' Put the selection length back
      '      tbThis.SelectionStart = intSelPos : tbThis.SelectionLength = intSelLen
      '    End If
      '    ' Debug.Print(dtrThis.Item("ResId").ToString)
      '  Else
      '    ' Find out 
      '  End If
      'End If
      ' Make sure dirty flag is set
      SetResDirty(True, True)
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/DoFeatChange error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tmeFeatChange_Tick
  ' Goal:   Process changes due to feature change
  ' History:
  ' 11-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tmeFeatChange_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmeFeatChange.Tick
    Dim dtrThis As DataRow        ' Currently selected datarow

    Try
      ' Switch off timer
      Me.tmeFeatChange.Enabled = False
      ' Get the datarow of the currently shown set
      dtrThis = objResEd.SelRow
      ' Validate
      If (dtrThis Is Nothing) Then Exit Sub
      ' Make sure changes are processed
      dtrThis.AcceptChanges()
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/tmeFeatChange error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  Private Sub tbDbFeat1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat1.Click
    DoFeatUrl(Me.tbDbFeat1)
  End Sub
  Private Sub tbDbFeat2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat2.Click
    DoFeatUrl(Me.tbDbFeat2)
  End Sub
  Private Sub tbDbFeat3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat3.Click
    DoFeatUrl(Me.tbDbFeat3)
  End Sub
  Private Sub tbDbFeat4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat4.Click
    DoFeatUrl(Me.tbDbFeat4)
  End Sub
  Private Sub tbDbFeat5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat5.Click
    DoFeatUrl(Me.tbDbFeat5)
  End Sub
  Private Sub tbDbFeat6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat6.Click
    DoFeatUrl(Me.tbDbFeat6)
  End Sub
  Private Sub tbDbFeat7_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat7.Click
    DoFeatUrl(Me.tbDbFeat7)
  End Sub
  Private Sub tbDbFeat8_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat8.Click
    DoFeatUrl(Me.tbDbFeat8)
  End Sub
  Private Sub tbDbFeat9_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat9.Click
    DoFeatUrl(Me.tbDbFeat9)
  End Sub
  Private Sub tbDbFeat10_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat10.Click
    DoFeatUrl(Me.tbDbFeat10)
  End Sub
  Private Sub tbDbFeat11_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat11.Click
    DoFeatUrl(Me.tbDbFeat11)
  End Sub
  Private Sub tbDbFeat12_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat12.Click
    DoFeatUrl(Me.tbDbFeat12)
  End Sub
  Private Sub tbDbFeat13_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat13.Click
    DoFeatUrl(Me.tbDbFeat13)
  End Sub
  Private Sub tbDbFeat14_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat14.Click
    DoFeatUrl(Me.tbDbFeat14)
  End Sub
  Private Sub tbDbFeat15_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat15.Click
    DoFeatUrl(Me.tbDbFeat15)
  End Sub
  Private Sub tbDbFeat16_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat16.Click
    DoFeatUrl(Me.tbDbFeat16)
  End Sub
  Private Sub tbDbFeat17_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat17.Click
    DoFeatUrl(Me.tbDbFeat17)
  End Sub
  Private Sub tbDbFeat18_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat18.Click
    DoFeatUrl(Me.tbDbFeat18)
  End Sub
  Private Sub tbDbFeat19_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat19.Click
    DoFeatUrl(Me.tbDbFeat19)
  End Sub
  Private Sub tbDbFeat20_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat20.Click
    DoFeatUrl(Me.tbDbFeat20)
  End Sub
  Private Sub tbDbFeat21_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat21.Click
    DoFeatUrl(Me.tbDbFeat21)
  End Sub
  Private Sub tbDbFeat22_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat22.Click
    DoFeatUrl(Me.tbDbFeat22)
  End Sub
  Private Sub tbDbFeat23_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat23.Click
    DoFeatUrl(Me.tbDbFeat23)
  End Sub
  Private Sub tbDbFeat24_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbDbFeat24.Click
    DoFeatUrl(Me.tbDbFeat24)
  End Sub
  Private Sub tbDbFeat1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat1.TextChanged
    DoFeatChange(1, Me.tbDbFeat1)
  End Sub
  Private Sub tbDbFeat2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat2.TextChanged
    DoFeatChange(2, Me.tbDbFeat2)
  End Sub
  Private Sub tbDbFeat3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat3.TextChanged
    DoFeatChange(3, Me.tbDbFeat3)
  End Sub
  Private Sub tbDbFeat4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat4.TextChanged
    DoFeatChange(4, Me.tbDbFeat4)
  End Sub
  Private Sub tbDbFeat5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat5.TextChanged
    DoFeatChange(5, Me.tbDbFeat5)
  End Sub
  Private Sub tbDbFeat6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat6.TextChanged
    DoFeatChange(6, Me.tbDbFeat6)
  End Sub
  Private Sub tbDbFeat7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat7.TextChanged
    DoFeatChange(7, Me.tbDbFeat7)
  End Sub
  Private Sub tbDbFeat8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat8.TextChanged
    DoFeatChange(8, Me.tbDbFeat8)
  End Sub
  Private Sub tbDbFeat9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat9.TextChanged
    DoFeatChange(9, Me.tbDbFeat9)
  End Sub
  Private Sub tbDbFeat10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat10.TextChanged
    DoFeatChange(10, Me.tbDbFeat10)
  End Sub
  Private Sub tbDbFeat11_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat11.TextChanged
    DoFeatChange(11, Me.tbDbFeat11)
  End Sub
  Private Sub tbDbFeat12_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat12.TextChanged
    DoFeatChange(12, Me.tbDbFeat12)
  End Sub
  Private Sub tbDbFeat13_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat13.TextChanged
    DoFeatChange(13, Me.tbDbFeat13)
  End Sub
  Private Sub tbDbFeat14_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat14.TextChanged
    DoFeatChange(14, Me.tbDbFeat14)
  End Sub
  Private Sub tbDbFeat15_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat15.TextChanged
    DoFeatChange(15, Me.tbDbFeat15)
  End Sub
  Private Sub tbDbFeat16_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat16.TextChanged
    DoFeatChange(16, Me.tbDbFeat16)
  End Sub
  Private Sub tbDbFeat17_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat17.TextChanged
    DoFeatChange(17, Me.tbDbFeat17)
  End Sub
  Private Sub tbDbFeat18_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat18.TextChanged
    DoFeatChange(18, Me.tbDbFeat18)
  End Sub
  Private Sub tbDbFeat19_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat19.TextChanged
    DoFeatChange(19, Me.tbDbFeat19)
  End Sub
  Private Sub tbDbFeat20_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat20.TextChanged
    DoFeatChange(20, Me.tbDbFeat20)
  End Sub
  Private Sub tbDbFeat21_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat21.TextChanged
    DoFeatChange(21, Me.tbDbFeat21)
  End Sub
  Private Sub tbDbFeat22_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat22.TextChanged
    DoFeatChange(22, Me.tbDbFeat22)
  End Sub
  Private Sub tbDbFeat23_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat23.TextChanged
    DoFeatChange(23, Me.tbDbFeat23)
  End Sub
  Private Sub tbDbFeat24_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFeat24.TextChanged
    DoFeatChange(24, Me.tbDbFeat24)
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbDbNotes_TextChanged
  ' Goal:   Process the change of the "Notes" attribute
  ' History:
  ' 04-01-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tbDbNotes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbNotes.TextChanged
    Dim ndxGeneral As XmlNode = Nothing ' General section
    Dim ndxNotes As XmlNode = Nothing   ' Notes

    Try
      ' Make sure changes are processed
      tdlResults.Tables("General").Rows(0).Item("Notes") = Me.tbDbNotes.Text
      ' Check XmlDocument
      ndxGeneral = pdxCrpDbase.SelectSingleNode("./descendant::General")
      If (ndxGeneral IsNot Nothing) Then
        ndxNotes = ndxGeneral.SelectSingleNode("./child::Notes")
        If (ndxNotes Is Nothing) Then ndxNotes = AddXmlChild(ndxGeneral, "Notes")
        ndxNotes.InnerText = Me.tbDbNotes.Text
      End If
      SetResDirty(True)
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/tbDbNotes error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbDbAnalysis_TextChanged
  ' Goal:   Process the change of the "Analysis" attribute
  ' History:
  ' 04-01-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tbDbAnalysis_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbAnalysis.TextChanged
    Dim ndxGeneral As XmlNode = Nothing ' General section
    Dim ndxAna As XmlNode = Nothing     ' Analysis

    Try
      ' Make sure changes are processed
      tdlResults.Tables("General").Rows(0).Item("Analysis") = Me.tbDbAnalysis.Text
      ' Check XmlDocument
      ndxGeneral = pdxCrpDbase.SelectSingleNode("./descendant::General")
      If (ndxGeneral IsNot Nothing) Then
        ndxAna = ndxGeneral.SelectSingleNode("./child::Analysis")
        If (ndxAna Is Nothing) Then ndxAna = AddXmlChild(ndxGeneral, "Analysis")
        ndxAna.InnerText = Me.tbDbAnalysis.Text
      End If
      SetResDirty(True)
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/tbDbAnalysis error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbResNotes_TextChanged
  ' Goal:   Process the change of the "Notes" attribute
  ' History:
  ' 23-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tbResNotes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbResNotes.TextChanged
    Try
      ' Validate
      If (bInit) AndAlso (ndxCurrentRes IsNot Nothing) AndAlso (Not objResEd.IsSelecting) Then
        ' Implement the change
        AddXmlAttribute(pdxCrpDbase, ndxCurrentRes, "Notes", Me.tbResNotes.Text)
        ' Set the results dirty flag
        SetResDirty(True, True)
      End If
      '' Make sure changes are processed
      'If (CtlChanged(objResEd, Me.tbResNotes, bInit)) Then SetResDirty(True)
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/tbResNotes error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbResTrans_TextChanged
  ' Goal:   Process the change of the "Pde" attribute
  ' History:
  ' 21-06-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tbResTrans_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbResTrans.TextChanged
    Dim ndxPde As XmlNode = Nothing ' PDE node

    Try
      ' Validate
      If (bInit) AndAlso (ndxCurrentRes IsNot Nothing) AndAlso (Not objResEd.IsSelecting) Then
        ' Validate presence of PDE child
        ndxPde = ndxCurrentRes.SelectSingleNode("./child::Pde")
        If (ndxPde Is Nothing) Then
          ' Add this child
          SetXmlDocument(pdxCrpDbase)
          ndxPde = AddXmlChild(ndxCurrentRes, "Pde")
        End If
        ' Validate
        If (ndxPde Is Nothing) Then Status("Cannot process back translation") : Exit Sub
        ' Implement the change
        ndxCurrentRes.SelectSingleNode("./child::Pde").InnerText = Me.tbResTrans.Text

        ' Set the results dirty flag
        SetResDirty(True, True)
      End If
      '' Make sure changes are processed
      'If (CtlChanged(objResEd, Me.tbResTrans, bInit)) Then SetResDirty(True)
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/tbResTrans error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbResCat_TextChanged
  ' Goal:   Process the change of the "Cat" attribute
  ' History:
  ' 04-01-2010  ERK Created
  ' 26-09-2013  ERK Adapted for XmlDocument processing
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tbResCat_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbResCat.TextChanged
    Try
      ' Validate
      If (bInit) AndAlso (ndxCurrentRes IsNot Nothing) AndAlso (Not objResEd.IsSelecting) Then
        ' Make sure changes are processed
        Status("Change of category: dataview part")
        If (Not CtlChanged(objResEd, Me.tbResCat, bInit)) Then Status("Could not change category") : Exit Sub
        ' Implement the change
        Status("Change of category: database part")
        AddXmlAttribute(pdxCrpDbase, ndxCurrentRes, "Cat", Me.tbResCat.Text)
        ' Set the results dirty flag
        Status("Change of category: flag change")
        SetResDirty(True, True)
        Status("okay")
      End If
      '' Make sure changes are processed
      'If (CtlChanged(objResEd, Me.tbResCat, bInit)) Then
      '  ' Set the results dirty flag
      '  SetResDirty(True, True)
      'End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/tbResCat error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cboGenTransLang_SelectedIndexChanged
  ' Goal:   Make sure the correct language is being set globally
  ' History:
  ' 17-03-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cboGenTransLang_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboGenTransLang.SelectedIndexChanged
    Try
      ' Validate
      If (Me.cboGenTransLang.Items.Count > 0) Then
        ' Find out what the value of the selected language is
        strCurrentTransLang = Me.cboGenTransLang.SelectedItem
      Else
        strCurrentTransLang = ""
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/cboGenTransLang error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub cboDepTransLang_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDepTransLang.SelectedIndexChanged
    Try
      ' Validate
      If (Me.cboDepTransLang.Items.Count > 0) Then
        ' Find out what the value of the selected language is
        strCurrentTransLang = Me.cboDepTransLang.SelectedItem
        dgvDep_SelectionChanged(sender, e)
      Else
        strCurrentTransLang = ""
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/cboDepTransLang error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   chbDepWholeText_CheckedChanged
  ' Goal:   Show the *WHOLE* text for dependency processing
  ' History:
  ' 28-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub chbDepWholeText_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbDepWholeText.CheckedChanged
    Try
      ' Set the flag
      bDepWholeText = (Me.chbDepWholeText.Checked)
      ' Reset current dep showing
      ShowCurrentDep()
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/chbDepWholeText error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cboResStatus_SelectedIndexChanged
  ' Goal:   Process the change of the "Status" attribute
  ' History:
  ' 12-05-2011  ERK Created
  ' 23-09-2013  ERK When status is changed, then do not allow it to be set to "Changed" anymore!
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cboResStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboResStatus.SelectedIndexChanged
    Try
      ' Validate
      If (bInit) AndAlso (ndxCurrentRes IsNot Nothing) AndAlso (Not objResEd.IsSelecting) Then
        ' Has anything been selected?
        If (Me.cboResStatus.SelectedIndex >= 0) Then
          ' Make sure changes are processed
          Status("Change of status: dataview part")
          If (CtlChanged(objResEd, Me.cboResStatus, bInit)) Then SetResDirty(True) : bResStatus = False
          ' Implement the change
          Status("Change of status: database part")
          AddXmlAttribute(pdxCrpDbase, ndxCurrentRes, "Status", Me.cboResStatus.SelectedValue)
          ' Set the results dirty flag
          Status("Change of status: flag change")
          SetResDirty(True, True)
          Status("okay")
        End If
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/cboResStatus error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   ResChange
  ' Goal:   Process the change of one field's value attribute
  ' History:
  ' 23-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub ResChange(ByVal strField As String, ByVal strValue As String)
    Dim dtrThis As DataRow        ' Currently selected datarow

    Try
      ' Validate
      If (bResSel) OrElse (objResEd Is Nothing) Then Exit Sub
      ' Get the datarow of the currently shown set
      dtrThis = objResEd.SelRow
      ' Validate
      If (dtrThis Is Nothing) Then Exit Sub
      ' Adapt the value
      dtrThis.Item(strField) = strValue
      ' Make sure dirty flag is set
      SetResDirty(True)
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/ResChange error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   splAddBottom_SplitterMoved
  ' Goal:   Make sure the top and bottom split container move in sync
  ' History:
  ' 03-03-2011  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub splAddBottom_SplitterMoved(ByVal sender As System.Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles splAddBottom.SplitterMoved
    Dim bBusy As Boolean = False ' I am busy now
    Dim intPos As Integer       ' My position

    Try
      ' Validat
      If (bBusy) Then Exit Sub
      ' Tell them I am busy
      bBusy = True
      ' Get my own position
      intPos = Me.splAddBottom.SplitterDistance
      ' Set the position of the other one
      Me.splAddTop.SplitterDistance = intPos
      ' Reset busy (Is this really necessary here?)
      bBusy = False
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/splitAddBottom error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub splAddTop_SplitterMoved(ByVal sender As System.Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles splAddTop.SplitterMoved
    Dim bBusy As Boolean = False ' I am busy now
    Dim intPos As Integer       ' My position

    Try
      ' Validat
      If (bBusy) Then Exit Sub
      ' Tell them I am busy
      bBusy = True
      ' Get my own position
      intPos = Me.splAddTop.SplitterDistance
      ' Set the position of the other one
      Me.splAddBottom.SplitterDistance = intPos
      ' Reset busy (Is this really necessary here?)
      bBusy = False
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/splitAddTop error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cboDbSelect_SelectedIndexChanged
  ' Goal:   Fill [Select] with the correct values
  ' History:
  ' 15-06-2011  ERK Created
  ' 26-09-2013  ERK Adapted for work with XmlDocument
  ' 07-10-2013  ERK Bugfix for work with XmlDocument
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cboDbSelect_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDbSelect.SelectedIndexChanged
    ' Dim tblRes As DataTable   ' Results table
    Dim dtrFound() As DataRow ' Result of selecting
    ' Dim dtrChild() As DataRow ' Children
    Dim strName As String     ' Name of selected feature
    Dim strFvalue As String   ' Value of feature
    Dim ndxRes As XmlNode     ' Result node
    Dim ndxF As XmlNode       ' Feature node
    Dim intI As Integer       ' Counter
    Dim intResId As Integer   ' COrrect ID
    ' Dim intJ As Integer       ' Counter
    Dim intPtc As Integer     ' Percentage
    Dim intNum As Integer     ' Number of elements

    Try
      ' Validate
      If (Not bInit) Then Exit Sub
      If (Me.cboDbSelect.Items.Count = 0) Then Exit Sub
      If (tdlResults Is Nothing) Then Exit Sub
      ' Get the name of the selected feature
      strName = Me.cboDbSelect.SelectedItem
      ' Make sure we are not interrupted
      dgvDbResult.DataSource = Nothing
      ' Get the first Result node
      ndxRes = pdxCrpDbase.SelectSingleNode("./descendant::Result")
      intNum = ndxRes.SelectNodes("./following-sibling::Result").Count + 1 : intI = 0
      ' Walk them all
      While (ndxRes IsNot Nothing)
        ' Show where we are 
        intPtc = (intI + 1) * 100 \ intNum
        Status("Writing [Select] feature " & intPtc & "%", intPtc)
        ' Get feature value
        strFvalue = ""
        ndxF = ndxRes.SelectSingleNode("./child::Feature[@Name = '" & strName & "']")
        If (ndxF IsNot Nothing) Then
          strFvalue = ndxF.Attributes("Value").Value
        End If
        ' Get to the corresponding one in [tdlResults]
        If (ndxRes.Attributes("ResId") IsNot Nothing) Then
          intResId = ndxRes.Attributes("ResId").Value
          dtrFound = tdlResults.Tables("Result").Select("ResId=" & intResId)
          ' Got it?
          If (dtrFound.Length > 0) Then
            dtrFound(0).Item("Select") = strFvalue
          End If
        End If
        ' Go to next one
        ndxRes = ndxRes.NextSibling : intI += 1
      End While

      ' ============ Recent history:
      '' Find the correct ones
      'dtrFound = tdlResults.Tables("Feature").Select("Name = '" & strName & "'")
      'For intI = 0 To dtrFound.Length - 1
      '  ' Show where we are 
      '  intPtc = (intI + 1) * 100 \ dtrFound.Count
      '  Status("Writing [Select] feature " & intPtc & "%", intPtc)
      '  ' Change the value of the [Select] feature of my parent
      '  dtrFound(intI).GetParentRow("Result_Feature").Item("Select") = dtrFound(intI).Item("Value")
      'Next intI

      ' =============== Longer history:
      '' Set the datatable
      'tblRes = tdlResults.Tables("Result")
      '' Walk through the database
      'For intI = 0 To tblRes.Rows.Count - 1
      '  ' Show where we are 
      '  intPtc = (intI + 1) * 100 \ tblRes.Rows.Count
      '  Status("Writing [Select] feature " & intPtc & "%", intPtc)
      '  ' Go through children
      '  dtrChild = tblRes(intI).GetChildRows("Result_Feature")
      '  For intJ = 0 To dtrChild.Length - 1
      '    ' Is this the correct feature?
      '    If (dtrChild(intJ).Item("Name") = strName) Then
      '      ' Fill item with correct information
      '      tblRes(intI).Item("Select") = dtrChild(intJ).Item("Value")
      '      ' Skip other rows
      '      Exit For
      '    End If
      '  Next intJ
      'Next intI
      ' Allow interrupt again
      If (Not InitResultsEditor()) Then Status("Could not re-initialize") : Exit Sub
      ' Make sure the dirty flag is set
      bResDirty = True
      ' All fine
      Status("Ready")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/cboDbSelect error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbRefChainSelect_TextChanged
  ' Goal:   Filter what is being shown
  ' History:
  ' 01-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tbRefChainSelect_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbRefChainSelect.TextChanged
    Dim strChainId As String = ""

    Try
      ' Trim
      strChainId = Trim(Me.tbRefChainSelect.Text)
      ' Validate
      If (IsNumeric(strChainId)) Then
        objRefChainEd.Filter = "ChainId = " & strChainId
      ElseIf (strChainId = "") Then
        objRefChainEd.Filter = ""
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/tbRefChainSelect error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   dgvRefItem_DoubleClick
  ' Goal:   Show additional items for the selected [ItemId]
  ' History:
  ' 01-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub dgvRefItem_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvRefItem.DoubleClick
    Dim intEtreeId As Integer ' Currently selected [eTreeId]
    Dim ndxThis As XmlNode    ' Current node

    Try
      ' Validate
      If (Trim(Me.tbRefItemEtreeId.Text) = "") Then Exit Sub
      ' Get what is selected
      intEtreeId = CInt(Trim(Me.tbRefItemEtreeId.Text))
      ' Validate
      If (intEtreeId > 0) Then
        ' Get our node
        ndxThis = IdToNode(intEtreeId)
        ' Select the correct tab page
        Me.TabControl1.SelectedTab = Me.tpEdit
        ' Jump to this node
        GotoNode(ndxThis)
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/tbRefItemDoubleClick error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   dgvRefItem_SelectionChanged
  ' Goal:   Show additional items for the selected [ItemId]
  ' History:
  ' 01-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub dgvRefItem_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvRefItem.SelectionChanged
    Dim intEtreeId As Integer ' Currently selected [eTreeId]
    Dim strRoot As String     ' Text of rootnode
    Dim strNext As String     ' Text of next node
    Dim strContext As String  ' Context for this node (previous, current, next clause)
    Dim ndxThis As XmlNode    ' Current node
    Dim ndxNext As XmlNode    ' Working node
    Dim ndxForest As XmlNode  ' Forest

    Try
      ' Validate
      If (Trim(Me.tbRefItemEtreeId.Text) = "") Then Exit Sub
      ' Get what is selected
      intEtreeId = CInt(Trim(Me.tbRefItemEtreeId.Text))
      strContext = ""
      ' Validate
      If (intEtreeId > 0) Then
        ' Get our node
        ndxThis = IdToNode(intEtreeId)
        ' Check if we point to anything
        ndxNext = CorefDst(ndxThis)
        If (ndxNext Is Nothing) Then
          ' We don't point to anything
          strRoot = "" : strNext = ""
        Else
          ' Get the text of the next node
          strNext = NodeText(ndxNext)
          ' Get the rootnode
          strRoot = NodeText(GetRootNode(ndxThis))
        End If
        ' Show the values
        Me.tbRefItemRootNode.Text = strRoot
        Me.tbRefItemRefNode.Text = strNext
        Me.tbRefItemRefId.Text = CorefDstId(ndxThis)
        Me.tbRefItemRefType.Text = CorefAttr(ndxThis, "RefType")
        ' Get the context of the current node
        ndxForest = ndxThis.SelectSingleNode(".//ancestor::forest")
        If (Not ndxForest Is Nothing) Then
          strContext = NodeInfo(ndxForest.PreviousSibling) & vbCrLf & _
                       "*** " & NodeInfo(ndxForest) & vbCrLf & _
                       NodeInfo(ndxForest.NextSibling) & vbCrLf
        Else
          strContext = "(error)"
        End If
      End If
      Me.tbRefItemContext.Text = strContext
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/dgvRefItem error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   dgvRefChain_DoubleClick
  ' Goal:   Go to the last item on this chain
  ' History:
  ' 01-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub dgvRefChain_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvRefChain.DoubleClick
    Dim dtrFound() As DataRow     ' Result of select
    Dim intChainId As Integer     ' Selected chain
    Dim intEtreeId As Integer     ' ID of the first element in the chain
    Dim ndxThis As XmlNode        ' Current node

    Try
      ' CHeck which one is selected
      intChainId = objRefChainEd.SelectedId
      If (intChainId >= 0) Then
        ' Select items from this chain
        dtrFound = tdlRefChain.Tables("Item").Select("ChainId=" & intChainId, "ItemId ASC")
        ' Element zero has the first element of this chain
        intEtreeId = CInt(dtrFound(0).Item("eTreeId"))
        ' Get this node
        ndxThis = IdToNode(intEtreeId)
        ' Select the correct tab page
        Me.TabControl1.SelectedTab = Me.tpEdit
        ' Jump to this item
        GotoNode(ndxThis)
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/dgvRefChainDoubleClick error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   dgvRefChain_SelectionChanged
  ' Goal:   Show additional items for the selected [ChainId]
  ' History:
  ' 01-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub dgvRefChain_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvRefChain.SelectionChanged
    Dim intChainId As Integer     ' Selected chain

    Try
      ' Only come here when we are initialised
      If (Not bInit) Then Exit Sub
      ' CHeck which one is selected
      intChainId = objRefChainEd.SelectedId
      Me.tbRefChainThis.Text = GetRefChain(intChainId)
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/dgvRefChain error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbRefItemSelect_TextChanged
  ' Goal:   Adapt the filter for this item
  ' History:
  ' 01-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tbRefItemSelect_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbRefItemSelect.TextChanged
    Dim strGrRole As String = ""

    Try
      ' Trim
      strGrRole = Trim(Me.tbRefItemSelect.Text)
      ' Validate
      If (strGrRole = "") Then
        objRefItemEd.Filter = ""
      Else
        objRefItemEd.Filter = "GrRole Like '" & strGrRole & "*'"
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/tbRefItemSelect error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbRefChainAnimacy_TextChanged
  ' Goal:   Process changes in animacy
  ' History:
  ' 09-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tbRefChainAnimacy_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbRefChainAnimacy.TextChanged
    Try
      ' Are we initialized yet?
      If (Not bInit) Then Exit Sub
      ' Validate
      If (objRefChainEd Is Nothing) Then Exit Sub
      ' Make sure changes are processed in the chain file
      If (CtlChanged(objRefChainEd, Me.tbRefChainAnimacy, bInit)) Then bRefDirty = True
      ' TODO: try to let the animacy perculate into all the constituents of the selected chain
      If (AnimacyPerculate(Me.tbRefChainId.Text, Me.tbRefChainAnimacy.Text)) Then
        ' Check if we need saving
        If (bNeedSaving) Then SetDirty(True) : bNeedSaving = False
      Else
        Status("Problem with animacy perculation")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/tbRefChainAnimacy error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuRefSave_Click
  ' Goal:   Save the refchain
  ' History:
  ' 09-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuRefSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRefSave.Click
    Try
      ' Are there any data?
      If (tdlRefChain Is Nothing) Then Status("First make chains") : Exit Sub
      ' Save the refchain
      If (Not RefChainSave()) Then
        Status("Could not save the refchain")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/tbRefSave error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuRefOverlap_Click
  ' Goal:   Check for overlapping referential chains
  ' History:
  ' 09-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuRefOverlap_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRefOverlap.Click
    Try
      ' Validation
      If (strCurrentFile = "") OrElse (pdxCurrentFile Is Nothing) Then Status("First load a file") : Exit Sub
      ' Switch to the right tab-page
      Me.TabControl1.SelectedTab = Me.tpEdit
      ' Create an XML data document
      If (Not CreateDataSet("RefChain.xsd", tdlRefChain)) Then Status("Unable to create a dataset") : Exit Sub
      ' Make the dataset
      If (Not RefListXML(tdlRefChain, strCurrentFile)) Then Status("Unable to make the RefChain dataset") : Exit Sub
      ' Set the richtextbox local copy
      InitAutoCoref(Me.tbEdtMain)
      ' Start overlap checking
      With frmOverlap
        ' Start this
        .Show()
        ' Check the status
        While (Not .Ready)
          Application.DoEvents()
        End While
      End With
      ' Switch off
      FinishAutoCoref()
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/tbRefSave error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   litTree_Click
  ' Goal:   Select the indicated element when it is clicked
  ' History:
  ' 18-06-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub litTree_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles litTree.Click
    Dim intIdx As Integer       ' Index within Shapes
    Dim ndxThis As XmlNode      ' Currently selected node

    Try
      ' Find out just where we are
      intIdx = litTree.HoveredId
      ' Got anything?
      If (intIdx >= 0) Then
        If (Not NodeTreeSelect(litTree.Shapes(intIdx).NodeName, litTree.Shapes(intIdx).NodeId)) Then
          Status("Could not select this node")
        Else
          ' Initialize
          ndxThis = Nothing
          If (LitToNode(litTree.Shapes(intIdx), ndxThis)) Then
            ' Show the status of this node
            Status(eTreeStatus(ndxThis))
            ' Set this as the currently selected node
            ndxCurrentTreeNode = ndxThis
            ' Check feature show updating
            FeatShowAdapt(ndxThis)
          Else
            Status("Unknown node")
          End If
        End If
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/litTreeMouseClick error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   NodeTreeSelect
  ' Goal:   Select one particular node in the tree
  ' History:
  ' 18-06-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function NodeTreeSelect(ByVal strNodeName As String, ByVal strNodeId As String) As Boolean
    Dim objRect As Lithium.SimpleRectangle
    Dim strText As String = ""  ' Where we hit it
    Dim intI As Integer         ' Counter
    Dim ndxThis As XmlNode = Nothing    ' one node
    Dim ndxForest As XmlNode = Nothing  ' One forest
    Dim pntThis As New Point(0, 0)
    Dim evThis As MouseEventArgs = Nothing

    Try
      ' Get the current forest
      If (Not GetCurrentForest(ndxForest, ndxThis)) Then Return False
      ' Visit all the shapes
      For intI = 0 To litTree.Shapes.Count - 1
        ' Access it
        objRect = litTree.Shapes(intI)
        ' Get the node
        Select Case objRect.NodeName
          Case "eTree"
            ndxThis = ndxForest.SelectSingleNode("./descendant::eTree[@Id='" & objRect.NodeId & "']")
          Case "eLeaf"
            ndxThis = ndxForest.SelectSingleNode("./descendant::eTree[@Id='" & objRect.NodeId & "']/child::eLeaf[1]")
          Case "forest"
            ndxThis = ndxForest
          Case Else
            ' Cannot handle this kind of node
            Status("Cannot handle nodes of type: " & objRect.NodeName)
        End Select
        ' Is this the selected one?
        If (strNodeName = objRect.NodeName) AndAlso (strNodeId = objRect.NodeId) Then
          objRect.ShapeColor = NodeColor(ndxThis, True)
          objRect.IsSelected = True
          ' Set the ID of the selected shape
          litTree.SelectedId = intI
          ' scroll there: horizontally
          If (objRect.X > litTree.HorizontalScroll.Value + litTree.Width) Then
            ' We need to scroll to the right
            litTree.HorizontalScroll.Value = objRect.X + objRect.Width - litTree.Width + 50
          ElseIf (objRect.X + objRect.Width < litTree.HorizontalScroll.Value) Then
            ' Need to scroll to the left
            litTree.HorizontalScroll.Value = objRect.X - 50
          End If
          ' scroll there: vertically
          If (objRect.Y > litTree.VerticalScroll.Value + litTree.Height) Then
            ' We need to scroll into the picture
            litTree.VerticalScroll.Value = objRect.Y + objRect.Height - litTree.Height + 50
          ElseIf (objRect.Y + objRect.Height < litTree.VerticalScroll.Value) Then
            ' Need to scroll to the left
            litTree.VerticalScroll.Value = objRect.Y - 50
          End If
        Else
          objRect.ShapeColor = NodeColor(ndxThis, False)
          objRect.IsSelected = False
        End If
      Next intI
      ' Return positively
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/litTreeMouseClick error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   LitToNode
  ' Goal:   Return the proper XmlNode for the indicated lithium object
  ' History:
  ' 07-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function LitToNode(ByRef objRect As Lithium.SimpleRectangle, ByRef ndxThis As XmlNode) As Boolean
    Dim ndxForest As XmlNode = Nothing  ' One forest

    Try
      ' Get the current forest
      If (Not GetCurrentForest(ndxForest, ndxThis)) Then Return False
      ' Get the node
      Select Case objRect.NodeName
        Case "eTree"
          ndxThis = ndxForest.SelectSingleNode("./descendant::eTree[@Id='" & objRect.NodeId & "']")
        Case "eLeaf"
          ndxThis = ndxForest.SelectSingleNode("./descendant::eTree[@Id='" & objRect.NodeId & "']/child::eLeaf[1]")
        Case "forest"
          ndxThis = ndxForest
        Case Else
          ' Cannot handle this kind of node
          Status("Cannot handle nodes of type: " & objRect.NodeName)
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/LitToNode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetNewLabel
  ' Goal:   Ask user for a new label at this point
  ' History:
  ' 18-06-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function GetNewLabel(ByRef strLabel As String, ByRef pntThis As Point) As Boolean
    Try
      ' Validate
      If (pntThis.X < 0) OrElse (pntThis.Y < 0) Then Return False
      ' Show we are busy with getting a label
      bLabelSelect = True : intLabelState = Windows.Forms.DialogResult.Ignore
      ' Make and position text label
      With tbLabel
        ' Position and size the label correctly, taking into account the scrolling of the window
        .Width = 150 : .Height = 50
        .Left = pntThis.X - IIf(litTree.HorizontalScroll.Visible, litTree.HorizontalScroll.Value, 0)
        .Top = pntThis.Y - IIf(litTree.VerticalScroll.Visible, litTree.VerticalScroll.Value, 0)
        ' Debug.Print(litTree.HorizontalScroll.Visible)
        ' Put the current text of @Label in there and show it
        .Text = strLabel : .Visible = True
        ' Select the current label completely
        .Select(0, .TextLength) : .Focus()
        ' Loop and wait
        While (.Visible) : Application.DoEvents() : End While
        ' Get the result
        Select Case intLabelState
          Case Windows.Forms.DialogResult.OK
            ' Set focus back to [litTree]
            Me.litTree.Focus()
            ' Return this label
            strLabel = .Text : Return True
        End Select
      End With
      ' Set focus back to [litTree]
      Me.litTree.Focus()
      ' Return failure
      Return False
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/GetNewLabel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   litTree_KeyDown
  ' Goal:   Allow user to "step" through elements, and highlight them
  '         Stepping can be done using:
  '           arrow up    // 'u'
  '           arrow down  // 'd'
  '           arrow left  // 'l'
  '           arrow right // 'r'
  '         Functionality after the user has pressed ESCAPE:
  '           insert      // 'i'  Insert a new level between me and my parent
  '           join right  // 'j'  Join under the node right of me
  '           join left   // 'k'  Join under the node left of me
  '           delete      // 'd'  Delete current node
  '           add right   // 'a'  Add a sibling to the right of me
  '           add left    // 'f'  Add a sibling to the left of me
  '           add child   // 'c'  Create/add a child under me
  '           add eLeaf   // 'e'  Create an <eLeaf> element under me
  '           cut forest  // 't'  truncate forest from this point on
  '           glue forest // 'g'  glue two forests together
  '         Functionality WITHOUT ESCAPE activated:
  '           space       //      Add to or delete from Query preparation batch
  ' History:
  ' 18-06-2012  ERK Created
  ' 03-01-2013  ERK Added options: insert, join, delete, add
  ' 21-04-2014  ERK Added "undo" option (working on it...)
  ' ------------------------------------------------------------------------------------------------------------
  Public Function litTreeKey(ByVal e As System.Windows.Forms.KeyEventArgs) As Boolean
    Dim ndxThis As XmlNode = Nothing    ' one node
    Dim ndxForest As XmlNode = Nothing  ' One forest
    Dim ndxNew As XmlNode = Nothing     ' one node
    'Dim strNodeName As String          ' Name of selected node
    'Dim strNodeId As String            ' id of selected node
    Dim strLabel As String = ""         ' New label
    Dim pntThis As Point = Nothing      ' Current position
    Dim bHandled As Boolean = False     ' Handled changes?

    Try
      ' The key commands only work when no ALT, SHIFT or CTRL is being pressed
      If (ShiftPressed()) Or (ControlPressed()) Or (AltPressed()) Then Return True
      ' Get the current forest
      If (Not GetCurrentForest(ndxForest, ndxThis)) Then Return False
      ' Get selected constituent
      ndxThis = SelectedTreeNode(litTree, strLabel, pntThis)
      ' Anything found?
      If (ndxThis Is Nothing) Then Status("Cannot find selected node") : Return True
      ' Look for GENERAL commands
      Select Case e.KeyCode
        Case Keys.Escape
          ' Note that ESCAPE has been pressed
          e.Handled = True : e.SuppressKeyPress = True
          ' Toggle edit mode
          eTreeEditMode("toggle")
        Case Else
          ' See if we have to deal with Editing or with Moving
          If (bEditMode) Then
            ' The commands issued are for EDITING
            Select Case e.KeyCode
              Case Keys.Down  ' Step down
                Return eTreeStepDown(ndxThis)
              Case Keys.Up    ' Step up
                Return eTreeStepUp(ndxThis)
              Case Keys.Right ' Step to constituent to the right
                Return eTreeStepRight(ndxThis)
              Case Keys.Left  ' Step to constituent to the left
                Return eTreeStepLeft(ndxThis)
              Case Keys.L     ' Move current constituent to the place of the left sibling
                If (Not AddTreeStack("l (move left)")) Then Return False
                If (eTreeMoveLeft(ndxThis, ndxNew, Me.litTree)) Then
                  Return eTreeHandled(ndxNew, e)
                End If
              Case Keys.R     ' Move current constituent to the place of the right sibling
                If (Not AddTreeStack("r (move right)")) Then Return False
                If (eTreeMoveRight(ndxThis, ndxNew, Me.litTree)) Then
                  Return eTreeHandled(ndxNew, e)
                End If
              Case Keys.U     ' Move current constituent upwards
                If (Not AddTreeStack("u (move up)")) Then Return False
                If (eTreeMoveUp(ndxThis, ndxNew, Me.litTree)) Then
                  Return eTreeHandled(ndxNew, e)
                End If
              Case Keys.Space ' A space means: edit the label of this node/endnode
                If (Not AddTreeStack("(edit label)")) Then Return False
                If (eTreeLabel(ndxThis, strLabel, pntThis)) Then
                  Return eTreeHandled(ndxThis, e, True)
                End If
              Case Keys.Insert, Keys.I    ' Allow inserting a new level between me and my parent
                If (Not AddTreeStack("i (insert)")) Then Return False
                If (eTreeInsertLevel(ndxThis, ndxNew)) Then
                  Return eTreeHandled(ndxNew, e)
                End If
              Case Keys.A     ' Add sibling to the right of me
                If (Not AddTreeStack("a (add sibling right)")) Then Return False
                If (eTreeAdd(ndxThis, ndxNew, "right")) Then
                  Return eTreeHandled(ndxNew, e)
                End If
              Case Keys.F     ' Add sibling to the left of me
                If (Not AddTreeStack("f (add sibling left)")) Then Return False
                If (eTreeAdd(ndxThis, ndxNew, "left")) Then
                  Return eTreeHandled(ndxNew, e)
                End If
              Case Keys.C     ' Create <eTree> child under me
                If (Not AddTreeStack("c (create child)")) Then Return False
                If (eTreeAdd(ndxThis, ndxNew, "child")) Then
                  Return eTreeHandled(ndxNew, e)
                End If
              Case Keys.E     ' Create <eLeaf> child under me
                If (Not AddTreeStack("e (create end-node)")) Then Return False
                If (eTreeAdd(ndxThis, ndxNew, "eLeaf")) Then
                  Return eTreeHandled(ndxNew, e)
                End If
              Case Keys.K     ' Join under LEFT
                If (Not AddTreeStack("k (join under left)")) Then Return False
                If (eTreeJoinUnder(ndxThis, ndxNew, False)) Then
                  Return eTreeHandled(ndxNew, e)
                End If
              Case Keys.J     ' Join under right
                If (Not AddTreeStack("j (join under right)")) Then Return False
                If (eTreeJoinUnder(ndxThis, ndxNew, True)) Then
                  Return eTreeHandled(ndxNew, e)
                End If
              Case Keys.D     ' Delete selected node
                If (Not AddTreeStack("d (delete node)")) Then Return False
                If (eTreeDelete(ndxThis, ndxNew)) Then
                  Return eTreeHandled(ndxNew, e)
                End If
              Case Keys.S     ' Re-analyze sentence
                If (Not AddTreeStack("s (re-analyze sentence)")) Then Return False
                If (eTreeSentence(ndxThis, ndxNew)) Then
                  Return eTreeHandled(ndxNew, e)
                End If
              Case Keys.T     ' Cut sentence here
                If (Not AddTreeStack("t (truncate here)")) Then Return False
                If (eTreeCutForest(ndxThis, ndxForest, ndxNew)) Then
                  Return eTreeHandled(ndxNew, e)
                End If
              Case Keys.G     ' Glue this sentence to the one preceding me
                If (Not AddTreeStack("g (glue to preceding)")) Then Return False
                If (eTreeGlueToPrec(ndxForest, ndxNew)) Then
                  Return eTreeHandled(ndxNew, e)
                End If
              Case Keys.D2    ' The "2" key indicates doubling the current <forest>
                If (Not AddTreeStack("2 (double forest)")) Then Return False
                If (eTreeDoubleForest(ndxForest, ndxNew)) Then
                  Return eTreeHandled(ndxNew, e)
                End If
            End Select
          Else
            ' The commands issued are for MOVING around
            Select Case e.KeyCode
              Case Keys.Down, Keys.D  ' Move down
                Return eTreeStepDown(ndxThis)
              Case Keys.Up, Keys.U    ' Move up
                Return eTreeStepUp(ndxThis)
              Case Keys.Right, Keys.R ' Move to constituent to the right
                Return eTreeStepRight(ndxThis)
              Case Keys.Left, Keys.L  ' Move to constituent to the left
                Return eTreeStepLeft(ndxThis)
              Case Keys.Delete
                If (ShiftPressed()) Then
                  ' Handle DEL key: delete all selections
                  NodeTreeSelect("eTree", "")
                  ' Return success
                  Status("")
                  Return True
                Else
                  Return False
                End If
              Case Keys.Space         ' Add to or delete from the query preparation node set
                If (eTreeQueryAdd(ndxThis, pntThis)) Then e.Handled = True : Return True
            End Select
          End If
      End Select

      ' Return failure
      Return False
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/litTreeKeyDown error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   eTreeRedoText
  ' Goal:   Refresh the text that has been changed
  ' History:
  ' 30-04-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function eTreeRedoText() As Boolean
    Dim intPos As Integer ' Position

    Try
      ' Load current section again
      bEdtLoad = True
      intPos = Me.tbEdtMain.SelectionStart
      ShowSection(Me.tbEdtMain, intCurrentSection)
      bEdtLoad = False
      Me.tbEdtMain.SelectionStart = intPos
      ' Go to this line
      GoToCurrentLine(Me.tbEdtMain)
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/eTreeRedoText error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   eTreeStepDown
  ' Goal:   Move cursor to the first node down
  ' History:
  ' 05-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function eTreeStepDown(ByRef ndxThis As XmlNode) As Boolean
    Try
      ' Is control pressed or not?
      If (ControlPressed()) Then
        ' Go to the last acceptable child
        ndxThis = ndxThis.LastChild
        While (Not ndxThis Is Nothing)
          ' Try this one
          If (DoLike(ndxThis.Name, "eTree|eLeaf")) Then Exit While
          ' Go to next sibling
          ndxThis = ndxThis.PreviousSibling
        End While
      Else
        ' Go to the first acceptable child
        ndxThis = ndxThis.FirstChild
        While (Not ndxThis Is Nothing)
          ' Try this one
          If (DoLike(ndxThis.Name, "eTree|eLeaf")) Then Exit While
          ' Go to next sibling
          ndxThis = ndxThis.NextSibling
        End While
      End If
      If (ndxThis Is Nothing) Then Status("Cannot go further") : Return True
      ' Found anything?
      If (DoLike(ndxThis.Name, "eTree|eLeaf")) Then
        Select Case ndxThis.Name
          Case "eLeaf"
            NodeTreeSelect("eLeaf", ndxThis.ParentNode.Attributes("Id").Value)
          Case "eTree"
            NodeTreeSelect("eTree", ndxThis.Attributes("Id").Value)
        End Select
        ' Add more information to the message
        Status(eTreeStatus(ndxThis))
      Else
        Status("Could not go down any further")
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/eTreeStepDown error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   eTreeStatus
  ' Goal:   Give the status of the indicated node
  ' History:
  ' 07-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function eTreeStatus(ByRef ndxThis As XmlNode) As String
    Dim ndxFeat As XmlNodeList  ' List of <f> features
    Dim strMsg As String = ""   ' What we show the user
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return "(empty)"
      ' Action depends on type
      Select Case ndxThis.Name
        Case "eLeaf"
          strMsg = "eLeaf [" & ndxThis.Attributes("Type").Value & "] "
          ' Add from and to information
          strMsg &= "@from=" & ndxThis.Attributes("from").Value & " "
          strMsg &= "@to=" & ndxThis.Attributes("to").Value & " "
        Case "eTree"
          strMsg = "eTree [" & ndxThis.Attributes("Label").Value & "] "
          ' Add from and to information
          strMsg &= "@from=" & ndxThis.Attributes("from").Value & " "
          strMsg &= "@to=" & ndxThis.Attributes("to").Value & " "
        Case "forest"
          strMsg = "forest [" & ndxThis.Attributes("Location").Value & "] "
        Case Else
          strMsg = "unknown"
      End Select
      ' Check for all the features under this node
      ndxFeat = ndxThis.SelectNodes("./child::fs/child::f")
      For intI = 0 To ndxFeat.Count - 1
        With ndxFeat(intI)
          strMsg &= .Attributes("name").Value & "=" & .Attributes("value").Value & " "
        End With
      Next intI
      ' Adapt showing of features
      FeatShowAdapt(ndxThis)
      ' Return the result
      Return strMsg
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/eTreeStatus error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return "(error)"
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   eTreeShowStatus
  ' Goal:   Give the status of the currently selected node
  ' History:
  ' 07-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub eTreeShowStatus()
    Dim ndxForest As XmlNode = Nothing
    Dim ndxThis As XmlNode = Nothing

    Try
      ' Get the current forest
      If (Not GetCurrentForest(ndxForest, ndxThis)) Then Exit Sub
      ' Swho status
      Status(eTreeStatus(ndxThis))
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/eTreeShowStatus error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   eTreeStepUp
  ' Goal:   Move cursor to the first node up
  ' History:
  ' 05-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function eTreeStepUp(ByRef ndxThis As XmlNode) As Boolean
    Try
      ' GO to the parent
      ndxThis = ndxThis.ParentNode
      If (ndxThis Is Nothing) Then Status("Cannot go further") : Return True
      ' Select the parent
      Select Case ndxThis.Name
        Case "eTree"
          NodeTreeSelect("eTree", ndxThis.Attributes("Id").Value)
        Case "forest"
          NodeTreeSelect("forest", ndxThis.Attributes("forestId").Value)
      End Select
      ' Return success
      Status(eTreeStatus(ndxThis))
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/eTreeStepUp error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   eTreeStepRight
  ' Goal:   Move cursor to the first node right
  ' History:
  ' 05-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function eTreeStepRight(ByRef ndxThis As XmlNode) As Boolean
    Dim ndxNew As XmlNode = Nothing ' The new node we have selected

    Try
      ' Action depends on what I am
      Select Case ndxThis.Name
        Case "eLeaf"
          ' Find right neighbour
          ' ndxNew = ndxThis.SelectSingleNode("./following::eLeaf[1]")
          ndxNew = ndxThis.SelectSingleNode("./ancestor::forest/descendant::eLeaf[@from > " & ndxThis.Attributes("to").Value & "][1]")
          If (ndxNew Is Nothing) Then
            Status("Cannot go to the right")
          Else
            NodeTreeSelect("eLeaf", ndxNew.ParentNode.Attributes("Id").Value)
            Status(eTreeStatus(ndxNew))
          End If
        Case "eTree"
          ' Find right neighbour
          ndxNew = ndxThis.SelectSingleNode("./following-sibling::eTree[1]")
          If (ndxNew Is Nothing) Then
            Status("Cannot go to the right")
          Else
            NodeTreeSelect("eTree", ndxNew.Attributes("Id").Value)
            Status(eTreeStatus(ndxNew))
          End If
        Case "forest"
          ' Try the next line
          If (Not NextLine()) Then
            ' Impossible...
            Status("Unable to go to the next line")
          End If
          ' Make sure the cursor in the main window scrolls to the line we are now at
          GoToCurrentLine(Me.tbEdtMain)
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/eTreeStepRight error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   eTreeStepLeft
  ' Goal:   Move cursor to the first node left
  ' History:
  ' 05-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function eTreeStepLeft(ByRef ndxThis As XmlNode) As Boolean
    Dim ndxNew As XmlNode = Nothing ' The new node we have selected

    Try
      ' Action depends on what I am
      Select Case ndxThis.Name
        Case "eLeaf"
          ' Find left neighbour, if any
          ' But the neighbour must be under the same forest!!
          ndxNew = ndxThis.SelectSingleNode("./preceding::eLeaf[1][ancestor::forest/@forestId = " & _
                                            ndxThis.SelectSingleNode("./ancestor::forest[1]").Attributes("forestId").Value & "]")
          If (ndxNew Is Nothing) Then
            Status("Cannot go to the left")
          Else
            NodeTreeSelect("eLeaf", ndxNew.ParentNode.Attributes("Id").Value)
            Status(eTreeStatus(ndxNew))
          End If
        Case "eTree"
          ' Find left sibling
          ndxNew = ndxThis.SelectSingleNode("./preceding-sibling::eTree[1]")
          If (ndxNew Is Nothing) Then
            Status("Cannot go to the left")
          Else
            NodeTreeSelect("eTree", ndxNew.Attributes("Id").Value)
            Status(eTreeStatus(ndxNew))
          End If
        Case "forest"
          ' Try the next line
          If (Not PrevLine()) Then
            ' Impossible...
            Status("Unable to go to the previous line")
          End If
          ' Make sure the cursor in the main window scrolls to the line we are now at
          GoToCurrentLine(Me.tbEdtMain)
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/eTreeStepLeft error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   eTreeHandled
  ' Goal:   Standard dealing with a correct handling of an eTree editing operation
  ' History:
  ' 05-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function eTreeHandled(ByRef ndxNew As XmlNode, ByRef e As System.Windows.Forms.KeyEventArgs, _
                                Optional ByVal bRedrawOnly As Boolean = False) As Boolean
    Try
      If (Not DoEtreeHandle(ndxNew, bRedrawOnly)) Then Return False
      ' Show this is handled properly
      e.Handled = True : Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/eTreeHandled error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  Public Function DoEtreeHandle(ByRef ndxNew As XmlNode, Optional ByVal bRedrawOnly As Boolean = False) As Boolean
    Dim ndxForest As XmlNode  ' Forest
    Dim ndxThis As XmlNode    ' Working node

    Try
      ' (1) Indicate that the PSDX file now needs to be saved
      SetDirty(True)
      ' Is this just re-drawing?
      If (bRedrawOnly) Then
        Me.litTree.DrawTree()
      Else
        ' (2) Find the forest we are under
        ndxForest = ndxNew.SelectSingleNode("./ancestor-or-self::forest[1]")
        If (ndxForest Is Nothing) Then
          ' Get the first forest of me overall
          ndxForest = pdxCurrentFile.SelectSingleNode("//forest[1]")
          ' Double check
          If (ndxForest Is Nothing) Then Return False
        End If
        ' (3) The node should be the first under the forest
        ndxThis = ndxForest.SelectSingleNode("./descendant::eTree[1]")
        ' Double check
        If (ndxThis Is Nothing) Then Status("Could not find constituent under forest") : Return False
        '' (4) Adapt the <eTree>/@Id values from here on
        ' === NO: do not do this anymore, but only on demand! ===
        'AdaptEtreeId(ndxThis.Attributes("Id").Value)
        ' (5) Only now are we able to make and show the tree again
        ShowTree(Me.litTree)
        ' (6) Can we show the etree?
        If (ndxNew.Name = "eTree") Then
          ' Select this node
          NodeTreeSelect("eTree", ndxNew.Attributes("Id").Value)
        ElseIf (ndxNew.Name = "eLeaf") Then
          ' Select the eLeaf
          NodeTreeSelect("eLeaf", ndxNew.ParentNode.Attributes("Id").Value)
        End If
      End If
      ' (7) Show this is handled properly
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/DoEtreeHandled error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   litTree_KeyDown
  ' Goal:   Allow user to "step" through elements, and highlight them
  ' History:
  ' 18-06-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub litTree_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles litTree.KeyDown
    Try
      ' Are we inside?
      If (bTreeMove) OrElse (bLabelSelect) Then Exit Sub
      bTreeMove = True
      ' Handle the keydown procedure
      If (litTreeKey(e)) Then e.Handled = True
      ' Show we are not moving again
      bTreeMove = False
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/litTree_KeyDown error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   litTree_MouseMove
  ' Goal:   Keep track of mouse movements
  ' History:
  ' 18-06-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub litTree_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles litTree.MouseMove
    Dim objRect As Lithium.ShapeBase
    'Dim pntThis As Point
    Dim ndxSeg As XmlNode
    Dim strText As String = ""  ' Where we hit it
    Dim intIdx As Integer       ' Index within Shapes
    Dim intForestId As Integer  ' ID of the frest

    Try
      ' Are we inside?
      If (bTreeMove) Then Exit Sub
      bTreeMove = True
      ' Find out just where we are
      intIdx = litTree.SelectedId
      If (intIdx >= 0) Then
        For Each objRect In litTree.Shapes
          If (e.X - litTree.AutoScrollPosition.X > objRect.X) AndAlso (e.X - litTree.AutoScrollPosition.X < objRect.X + objRect.Width) AndAlso _
             (e.Y - litTree.AutoScrollPosition.Y > objRect.Y) AndAlso (e.Y - litTree.AutoScrollPosition.Y < objRect.Y + objRect.Height) Then
            With objRect
              ' Do what needs to be done here
              If (.NodeName = "forest") Then
                ' Show the translation of this verse as status
                intForestId = .NodeId
                ndxSeg = pdxCurrentFile.SelectSingleNode("//descendant::forest[@forestId='" & intForestId & "']/div[@lang='" & strCurrentTransLang & "']/seg")
                If (ndxSeg IsNot Nothing) Then
                  strText = ndxSeg.InnerText
                  Status(strText)
                  ' Me.tpTree.ToolTipText = strText
                  ' MsgBox(strText)
                End If
              End If
              Exit For
            End With
          End If
        Next
      End If
      ' Signal that we are outside
      bTreeMove = False
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/litTreeMouseMove error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   litTree_MouseUp
  ' Goal:   Note which nodes should be collapsed and which expanded
  ' History:
  ' 15-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub litTree_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles litTree.MouseUp
    Try
      If (Not CollapseChange(Me.litTree)) Then
        Status("Problem in mouse")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/litTreeMouseUp error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   litTree_OnInfo
  ' Goal:   Show necessary information
  ' History:
  ' 19-06-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub litTree_OnInfo(ByVal message As String) Handles litTree.OnInfo
    Try
      MsgBox(message)

    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/litTreeOnInfo error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbLabel_KeyDown
  ' Goal:   Handle key events on the label entry textbox
  ' History:
  ' 19-06-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tbLabel_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tbLabel.KeyDown
    Try
      Select Case e.KeyCode
        Case Keys.Enter
          ' Disappear and show I am ready
          Me.tbLabel.Visible = False
          intLabelState = Windows.Forms.DialogResult.OK
          ' Show that we handled the event?
          e.Handled = True
          e.SuppressKeyPress = True
          ' Show we are ready getting a label
          bLabelSelect = False
        Case Keys.Escape
          ' Disappear and show the result should be neglected
          Me.tbLabel.Visible = False
          intLabelState = Windows.Forms.DialogResult.Cancel
          ' Show that we handled the event?
          e.Handled = True
          e.SuppressKeyPress = True
          ' Show we are ready getting a label
          bLabelSelect = False
      End Select
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/tbLabel_Keydown error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuEditCopyTree_Click
  ' Goal:   Copy a tree to the clipboard
  ' History:
  ' 19-06-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuEditCopyTree_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditCopyTree.Click
    Dim mfThis As System.Drawing.Imaging.Metafile   ' The metafile we store our object
    Dim penThis As New Pen(Brushes.Black)           ' The pen with which we draw

    Dim grpThis As Graphics                         ' Graphics object
    Dim intI As Integer                             ' Counter
    Dim intWidth As Integer = 100                   ' Width of image
    Dim intHeight As Integer = 100                  ' Height of image
    Dim intBoxHeight As Integer = 10                ' Height of one box
    Dim strFile As String = ""                      ' File name
    Dim me_gr As Graphics = Me.CreateGraphics       ' Temporary graphics object
    Dim me_hdc As IntPtr = me_gr.GetHdc             ' Temporary handle
    Dim bounds As RectangleF

    Try
      ' Validate: Check if the right tab page is selected
      If (Not Me.TabControl1.SelectedTab Is Me.tpTree) Then Exit Sub
      ' Determine the filename
      strFile = strWorkDir & "\Tree.emf"
      ' Set the width of the pen
      penThis.Width = 1
      litTree.SaveGraphAs(strWorkDir & "\Tree.xml")
      ' Define the bounds of the image
      If (Not GetImgSizes(intWidth, intHeight, intBoxHeight)) Then Exit Sub
      bounds = New RectangleF(0, 0, intWidth, intHeight)
      mfThis = New System.Drawing.Imaging.Metafile(strFile, me_hdc, bounds, _
                                            Imaging.MetafileFrameUnit.Point, Imaging.EmfType.EmfPlusDual)
      ' Release the temporary HDC object
      me_gr.ReleaseHdc(me_hdc)
      ' Set the graphics object
      grpThis = Graphics.FromImage(mfThis)
      With grpThis
        .PageUnit = GraphicsUnit.Point
        .Clear(Color.White)
      End With
      ' Walk all the connections
      For intI = 0 To litTree.Connections.Count - 1
        ' Access this connection
        With litTree.Connections(intI)
          ' Check if this connection is under a node that is not "Expanded"
          If (IsExpanded(.From)) Then
            ' Draw it
            grpThis.DrawLine(penThis, .Start.X, .Start.Y - intBoxHeight \ 2, .End.X, .End.Y + intBoxHeight \ 2)
          End If
        End With
      Next intI
      ' Walk all the shapes
      For intI = 0 To litTree.Shapes.Count - 1
        ' Check if this shape is under a node that is not "Expanded"
        If (IsExpanded(litTree.Shapes(intI))) Then
          ' Access this connection
          With litTree.Shapes(intI)
            ' Draw it
            grpThis.DrawRectangle(penThis, .Left, .Top, .Width, .Height)
            ' Is this one selected?
            If (.IsSelected) Then
              ' Fill it with yellow
              grpThis.FillRectangle(Brushes.Yellow, .Left + 1, .Top + 1, .Width - 2, .Height - 2)
            End If
            ' Put the text there
            grpThis.DrawString(.Text, New Font("Times New Roman", 10, FontStyle.Regular), _
                               Brushes.Black, .Location.X + 5, .Location.Y + 5)
          End With
        End If
      Next intI
      ' Release them
      grpThis.Dispose()
      mfThis.Dispose()
      ' Load the object again
      mfThis = New System.Drawing.Imaging.Metafile(strFile)
      ClipboardMetafileHelper.PutEnhMetafileOnClipboard(Me.Handle, mfThis)
      ' Dispose the object again
      mfThis.Dispose()
      ' Tell the user where it is saved
      Status("Image saved at: " & strFile)
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/EditCopyTree error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   IsExpanded
  ' Goal:   Check if this shape is under an ancestor that is NOT expanded
  ' History:
  ' 21-06-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function IsExpanded(ByRef shpThis As ShapeBase) As Boolean
    Dim shpWork As ShapeBase  ' Working node

    Try
      ' Take this as working node
      shpWork = shpThis
      ' Loop through shapes
      While (Not shpWork Is Nothing)
        ' Is this one expanded?
        If (Not shpWork.Expanded) AndAlso (shpWork.ChildNodes.Count > 0) Then
          ' This node is NOT expanded
          Return False
        End If
        ' Go to next shape
        shpWork = shpWork.ParentNode
      End While
      ' Coming here means that we ARE in fact expanded
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/IsExpanded error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetImgSizes
  ' Goal:   Get the necessary width and height of the image
  '         Also get the height of boxes
  ' History:
  ' 19-06-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function GetImgSizes(ByRef intWidth As Integer, ByRef intHeight As Integer, _
                               ByRef intBoxHeight As Integer) As Boolean
    Dim intI As Integer ' Counter

    Try
      ' Validate
      If (litTree Is Nothing) Then Return False
      If (litTree.Shapes.Count = 0) Then Return False
      ' Get the box height
      intBoxHeight = litTree.Shapes(0).Height
      ' Initialise
      intWidth = 0 : intHeight = 0
      ' Walk all the shapes
      For intI = 0 To litTree.Shapes.Count - 1
        ' Is this one expanded?
        If (IsExpanded(litTree.Shapes(intI))) Then
          ' Access this shape
          With litTree.Shapes(intI)
            ' Adapt sizes if needed
            If (.Right > intWidth) Then intWidth = .Right
            If (.Bottom > intHeight) Then intHeight = .Bottom
          End With
        End If
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/GetImgSizes error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbDbFilter_TextChanged
  ' Goal:   Filter the database values
  ' History:
  ' 25-09-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tbDbFilter_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDbFilter.TextChanged
    Try
      ' Start the corpus results filter timer
      Me.tmeCrpResFilter.Enabled = True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/tbDbFilter error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub tmeCrpResFilter_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmeCrpResFilter.Tick
    Try
      ' Reset the corpus result filter
      Me.tmeCrpResFilter.Enabled = False
      ' Adjust the filter
      DoDbFilter()
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/tmeCrpResFilter_Tick error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub DoDbFilter()
    Dim strDbFilter As String = ""  ' The text of the filter

    Try
      ' Make sure we are not interrupted
      If (bDbFilter) Then Exit Sub
      bDbFilter = True
      ' Trim
      strDbFilter = Trim(Me.tbDbFilter.Text)
      '' Adaptations
      'strDbFilter = strDbFilter.Replace("'", "''")
      ' Validate
      If (IsNumeric(strDbFilter)) Then
        objResEd.Filter = "ResId = " & strDbFilter & " OR forestId = " & strDbFilter & " OR eTreeId = " & strDbFilter
      ElseIf (strDbFilter = "") Then
        objResEd.Filter = ""
      ElseIf (InStr(strDbFilter, "=") > 0) OrElse (InStr(strDbFilter.ToLower, " like ") > 0) Then
        ' Take the whole filter expression, and do not grumble of there is an error
        Try
          objResEd.Filter(True) = strDbFilter
        Catch ex As Exception
          Status("filter syntax not (yet) correct...")
        End Try
      Else
        ' Adaptations
        strDbFilter = strDbFilter.Replace("'", "''")
        Status("filtering...")
        objResEd.Filter(True) = "(Search LIKE " & "'" & strDbFilter & "*" & "'" & ")" & _
                          " OR (Cat LIKE " & "'" & strDbFilter & "*" & "'" & ")" & _
                          " OR (Select LIKE " & "'" & strDbFilter & "*" & "'" & ")" & _
                          " OR (Status LIKE " & "'" & strDbFilter & "*" & "'" & ")" & _
                          " OR (Notes LIKE " & "'" & strDbFilter & "*" & "'" & ")" & _
                          " OR (TextId LIKE " & "'" & strDbFilter & "*" & "'" & ")" & _
                          " OR (Text LIKE " & "'" & strDbFilter & "*" & "'" & ")" & _
                          " OR (Psd LIKE " & "'" & strDbFilter & "*" & "'" & ")" & _
                          " OR (Period LIKE " & "'" & strDbFilter & "*" & "'" & ")"
      End If
      ' Show how many results there are
      Status("Filtered results: " & Me.dgvDbResult.RowCount)
      ' Switch off the interrupt-prohibition flag
      bDbFilter = False
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/DoDbFilter error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoDbXpathFilter
  ' Goal:   Filter the database values by making use of an Xpath specification
  ' History:
  ' 22-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub DoDbXpathFilter()
    Try

    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/DoDbXpathFilter error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdDbSrcDir_Click
  ' Goal:   Change the value of the source directory
  ' History:
  ' 02-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdDbSrcDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDbSrcDir.Click
    Dim strSelDir As String = ""  ' The selected directory

    Try
      ' Retrieve what we have right now
      strSelDir = Me.tbDbSrcDir.Text
      ' Try get directory
      If (GetDirName(Me.FolderBrowserDialog1, strSelDir, _
                     "Select the directory where corpus texts are located", strSelDir)) Then
        ' Replace what is shown
        Me.tbDbSrcDir.Text = strSelDir
        ' Replace inside
        tdlResults.Tables("General").Rows(0).Item("SrcDir") = strSelDir
        ' Show results are dirty
        SetResDirty(True)
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/cmdDbSrcDir error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   eTreeEditMode
  ' Goal:   Issue a command pertaining to the "Edit Mode"
  ' History:
  ' 05-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub eTreeEditMode(ByVal strCommand As String)
    Try
      Select Case strCommand
        Case "toggle"
          ' Toggle the mode
          bEditMode = (Not bEditMode)
        Case "reset"
          ' Switch edit mode off
          bEditMode = False
      End Select
      ' Show what?
      If (bEditMode) Then
        Me.tbEditMode.Visible = True
      Else
        Me.tbEditMode.Visible = False
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/eTreeEditMode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   eTreeLabel
  ' Goal:   Get a new label for me
  ' History:
  ' 05-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function eTreeLabel(ByRef ndxThis As XmlNode, ByVal strLabel As String, ByVal pntThis As Point) As Boolean
    Dim bHandled As Boolean = False ' Handled or not

    Try
      ' Add the new label at this point in the tree
      Select Case ndxThis.Name
        Case "eTree", "eLeaf"
          ' Get a new label
          If (GetNewLabel(strLabel, pntThis)) Then
            ' Change the label and show we handled it
            If (Not eTreeDoLabel(ndxThis, strLabel, bHandled)) Then Return False
            ' If this is an <eLeaf>, then adapt sentence
            If (ndxThis.Name = "eLeaf") Then
              eTreeSentence(ndxThis, ndxThis)
              ' Make sure editor is set to dirty
              bEdtDirty = True
            End If
          End If
        Case "forest"
          ' Change the location??
          Select Case MsgBox("Would you really like to change the LOCATION identifier of this forest node?", MsgBoxStyle.YesNoCancel)
            Case MsgBoxResult.Yes
              ' Get a new label
              If (GetNewLabel(strLabel, pntThis)) Then
                ' Change the label and show we handled it
                If (Not eTreeDoLabel(ndxThis, strLabel, bHandled)) Then Return False
              End If
          End Select
      End Select
      ' Need handling?
      If (bHandled) Then
        ' Set the text within the shape
        litTree.Shapes(litTree.SelectedId).Text = strLabel
        ' Make sure the focus is on us again
        Me.TabControl1.Focus() : Me.litTree.Focus()
        ' Return success
        Return True
      Else
        Return False
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/eTreeLabel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  Private Sub mnuToolsTrial_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Dim strFile As String = "d:\data files\corpora\English\OEtags\OEtagPOS.txt"
    Dim strFileOut As String = "d:\data files\corpora\English\OEtags\OEtagPOS.xml"
    Dim strText As String = ""
    Dim strXsd As String = "d:\data files\corpora\English\OEtags\OEdata.xsd"
    Dim tdlMorph As DataSet = Nothing
    Dim dtrThis As DataRow = Nothing
    Dim dtrParent As DataRow = Nothing
    Dim arText() As String  ' Array of lines
    Dim arLine() As String  ' One line
    Dim intI As Integer     ' Counter

    Try
      ' Validate
      If (Not IO.File.Exists(strFile)) Then Status("Could not find input file") : Exit Sub
      Status("Working...")
      ' Read the text
      strText = IO.File.ReadAllText(strFile)
      ' Divide into array
      arText = Split(strText, vbCrLf)
      ' Create dataset
      If (Not CreateDataSet(strXsd, tdlMorph)) Then Status("Could not create dataset") : Exit Sub
      ' Create a list row
      CreateListRow(tdlMorph, "TagList")
      dtrParent = tdlMorph.Tables("TagList").Rows(0)
      ' Read data into dataset
      For intI = 0 To arText.Length - 1
        ' Validate
        If (arText(intI) <> "") Then
          ' Read line 
          arLine = Split(arText(intI), vbTab)
          dtrThis = AddOneDataRowWithParent(tdlMorph, "Tag", "TagId", dtrParent)
          With dtrThis
            .Item("POS") = arLine(0).Replace("""", "")
            .Item("Parent") = arLine(1).Replace("""", "")
            .Item("Ext") = arLine(2).Replace("""", "")
            .Item("OEtag") = arLine(3).Replace("""", "")
            .Item("Descr") = arLine(4).Replace("""", "")
          End With
        End If
      Next intI
      ' Save the result
      tdlMorph.WriteXml(strFileOut)
      ' Show we are ready
      Status("Ready")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/ToolsTrial error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxReanalyze_Click
  ' Goal:   Re-analyze all the <forest> nodes in this text: 
  '           calculate the @from, @to values and the div[@lang='org']/seg values
  ' History:
  ' 07-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxReanalyze_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxReanalyze.Click
    Dim ndxWork As XmlNode = Nothing  ' Working node
    Dim ndxPar As XmlNode             ' Parent
    Dim intI As Integer               ' Counter
    Dim intForId As Integer           ' Current forest's @forestId
    Dim strLoc As String              ' Current forest's @Location
    Dim intPtc As Integer             ' Percentage
    Dim intPos As Integer             ' Where are we?
    Dim bChanged As Boolean = False   ' Changed or not?

    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Status("First load a file") : Exit Sub
      If (pdxList.Count = 0) Then Status("First load a file") : Exit Sub
      ' clear the collapse box
      CollapseClear()
      ' Visit all the <forest> lines
      ndxPar = pdxList(0).ParentNode
      For intI = 0 To pdxList.Count - 1
        ' Where are we?
        intPtc = (intI + 1) * 100 \ pdxList.Count
        Status("Re-analyzing " & intPtc & "%", intPtc)
        ' Re-analyze this line
        If (eTreeSentence(pdxList(intI), ndxWork)) Then bChanged = True
        ' Check if there are any non-forests following
        ndxWork = pdxList(intI).NextSibling
        While (ndxWork IsNot Nothing) AndAlso (ndxWork.Name <> "forest")
          ' Delete it
          ndxWork.RemoveAll()
          ndxPar.RemoveChild(ndxWork)
          ' Try next sibling
          ndxWork = pdxList(intI).NextSibling
        End While
        ' Also check the @forestId value and the corresponding @Location value
        intForId = CInt(pdxList(intI).Attributes("forestId").Value)
        strLoc = pdxList(intI).Attributes("Location").Value.ToString
        If (intForId <> intI + 1) Then
          ' Adapt the forest id
          intForId = intI + 1
          pdxList(intI).Attributes("forestId").Value = intForId
          intPos = InStrRev(strLoc, ".")
          If (intPos > 0) Then strLoc = Strings.Left(strLoc, intPos - 1)
          strLoc &= "." & intForId
          pdxList(intI).Attributes("Location").Value = strLoc
          ' Make sure dirty flag is set
          bChanged = True
        End If
      Next intI
      ' Adapt the @Id values
      AdaptEtreeId(1)
      ' Need to set dirty or not?
      If (bChanged) Then SetDirty(True)
      ' Load current section again
      bEdtLoad = True
      intPos = Me.tbEdtMain.SelectionStart
      ShowSection(Me.tbEdtMain, intCurrentSection)
      bEdtLoad = False
      Me.tbEdtMain.SelectionStart = intPos
      ' Go to this line
      GoToCurrentLine(Me.tbEdtMain)
      '  ShowPde(Me.tbEdtMain, Me.tbMainPde, intPos)
      ' Which tab-page are we?
      If (Me.TabControl1.SelectedTab.Name = Me.tpTree.Name) Then
        ' (5) Only now are we able to make and show the tree again
        ShowTree(Me.litTree)
      End If
      ' Show ready
      Status("Ready")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/mnuSyntaxReanalyze_Click error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxEditHelp_Click
  ' Goal:   Toggle a help menu for the key-combinations that can be used to edit a tree
  ' History:
  ' 03-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxEditHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxEditHelp.Click
    Dim bShowHelp As Boolean  ' Whether help needs to be shown or not

    Try
      ' Determine the current status
      bShowHelp = (Not Me.mnuSyntaxEditHelp.Checked)
      Me.mnuSyntaxEditHelp.Checked = bShowHelp
      If (bShowHelp) Then
        ' Access it
        With Me.tbTreeHelp
          ' Fill it with correct text
          .Text = "Tree editing:" & vbCrLf & _
            " i  - Insert a new level between me and my parent" & vbCrLf & _
            " j  - Join under the node right of me" & vbCrLf & _
            " k  - Join under the node left of me" & vbCrLf & _
            " d  - Delete current node" & vbCrLf & _
            " a  - Add sibling to the right of me" & vbCrLf & _
            " f  - Add sibling to the left of me" & vbCrLf & _
            " c  - Create constituent child under me" & vbCrLf & _
            " e  - Create endnode child under me" & vbCrLf & _
            " r  - Move node to the right" & vbCrLf & _
            " l  - Move node to the left" & vbCrLf & _
            " u  - Move node to upward" & vbCrLf & _
            " t  - Truncate sentence here" & vbCrLf & _
            " g  - Glue sentence to preceding one" & vbCrLf & _
            " s -  Re-analyze sentence"
          ' We need to show it
          .Visible = True
        End With
      Else
        ' We need to hide it
        Me.tbTreeHelp.Visible = False
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxEditHelp error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxUpdate_Click
  ' Goal:   Apply automatic repair to the indicated Chechen parsed files
  ' History:
  ' 11-02-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxUpdate.Click
    Dim strDirIn As String = "" ' Input directory
    Dim strFileIn As String     ' Input file
    Dim arFile() As String      ' Input files
    Dim strSyntRep As String    ' HTML syntax report file
    Dim intI As Integer         ' Counter

    Try
      ' Initialize directory
      strDirIn = strLastSyntaxUpdate
      If (strDirIn = "") Then strDirIn = strWorkDir
      ' Get the correct input directory
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, "Directory where learn-files are located", strDirIn)) Then
        Status("Could not open input directory")
        Exit Sub
      End If
      ' Check what we returned
      If (Not IO.Directory.Exists(strDirIn)) Then Exit Sub
      ' Save this directory, if it is different from what we have
      If (strLastSyntaxUpdate <> strDirIn) Then
        strLastSyntaxUpdate = strDirIn
        SetTableSetting(tdlSettings, "LastSyntaxUpdate", strLastSyntaxUpdate)
        XmlSaveSettings(strSetFile)
      End If
      ' Get all the necessary files in this directory (plus subdirectories)
      arFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories)
      bInterrupt = False
      ' Walk the files
      For intI = 0 To arFile.Length - 1
        ' Get this file
        strFileIn = arFile(intI)
        ' Try to learn from this one
        If (Not OneUpdateSyntaxFile(strFileIn)) Then
          ' Warn the user and exit
          Status("SyntaxUpdate: Could not process file: " & strFileIn)
          Exit Sub
        End If
      Next intI
      ' Make sure we reset the currently loaded file
      pdxCurrentFile = Nothing
      ' SHow we are ready
      Status("Ready")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxUpdate error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxLearn_Click
  ' Goal:   Learn syntactic structure (what belongs to a constituent)
  ' History:
  ' 31-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxLearn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxLearn.Click
    Dim strDirIn As String = "" ' Input directory
    Dim strFileIn As String     ' Input file
    Dim arFile() As String      ' Input files
    Dim strSyntRep As String    ' HTML syntax report file
    Dim intI As Integer         ' Counter

    Try
      ' Other initialisations
      If (Not InitSyntax()) Then Status("Could not initialize syntax") : Exit Sub
      ' Check if we already did this
      If (tdlSyntax.Tables("Const").Rows.Count > 0) Then
        ' Ask user if he wants to re-do it
        Select Case MsgBox("Learn data is already available." & vbCrLf & _
                           "Would you like to learn again from scratch?", MsgBoxStyle.YesNoCancel)
          Case MsgBoxResult.Cancel, MsgBoxResult.No
            Exit Sub
        End Select
      End If
      ' Clear the <Cons> table and the <User> table
      Status("Clearing old information...")
      ClearTable(tdlSyntax.Tables("Const"))
      ClearTable(tdlSyntax.Tables("User"))
      ClearTable(tdlSyntax.Tables("Ptree"))
      ClearTable(tdlSyntax.Tables("Part"))
      tdlSyntax.AcceptChanges()
      Status("continuing...")
      ' Initialize directory
      strDirIn = strLastSyntaxLearn
      If (strDirIn = "") Then strDirIn = strWorkDir
      ' Get the correct input directory
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, "Directory where learn-files are located", strDirIn)) Then
        Status("Could not open input directory")
        Exit Sub
      End If
      ' Check what we returned
      If (Not IO.Directory.Exists(strDirIn)) Then Exit Sub
      ' Save this directory, if it is different from what we have
      If (strLastSyntaxLearn <> strDirIn) Then
        strLastSyntaxLearn = strDirIn
        SetTableSetting(tdlSettings, "LastSyntaxLearn", strLastSyntaxLearn)
        XmlSaveSettings(strSetFile)
      End If
      ' Get all the necessary files in this directory (plus subdirectories)
      arFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories)
      bInterrupt = False
      ' Walk the files
      For intI = 0 To arFile.Length - 1
        ' Get this file
        strFileIn = arFile(intI)
        ' Try to learn from this one
        If (Not OneLearnSyntaxFile(strFileIn)) Then
          ' Warn the user and exit
          Status("SyntaxLearn: Could not process file: " & strFileIn)
          Exit Sub
        End If
      Next intI
      ' Try to add rules
      If (Not DerivePartialRules()) Then Logging("Could not derive partial rules")
      ' Determine syntax report file name
      strSyntRep = GetDocDir() & "\SyntaxReport.htm"
      ' Get an overview on what we have learned
      IO.File.WriteAllText(strSyntRep, SyntaxReport)
      Me.wbReport.Navigate(strSyntRep)
      Me.TabControl1.SelectedTab = Me.tpReport
      ' Make sure we reset the currently loaded file
      pdxCurrentFile = Nothing
      ' Save the results
      SaveSyntax()
      ' SHow we are ready
      Status("Ready")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxLearn error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxSuggest_Click
  ' Goal:   Make one (or more?) suggestion(s) to build the syntax of the currently loaded file
  ' History:
  ' 31-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxSuggest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxSuggest.Click
    Dim ndxForest As XmlNode = Nothing  ' Currently selected forest
    Dim ndxThis As XmlNode = Nothing    ' Working node
    Dim ndxNew As XmlNode = Nothing     ' Newly created node
    Dim ndxList As XmlNodeList          ' List of elements
    Dim dtrNew As DataRow               ' One datarow element to play with
    Dim dtrFound() As DataRow           ' Result of select
    Dim strChildren As String = ""      ' Labels of the children currently evaluated
    Dim strLabel As String = ""         ' Constituent label (if any)
    Dim colLabel As New StringColl      ' Constituent labels (if any)
    Dim intI As Integer                 ' Counter
    Dim intJ As Integer                 ' Counter
    Dim intK As Integer                 ' Counter
    Dim intUserId As Integer            ' Selected user id
    Dim bChanged As Boolean = False     ' Did we actually change something?

    Try
      ' Validate: a file must be loaded
      If (pdxCurrentFile Is Nothing) Then Status("First load a file") : Exit Sub
      ' Validate: we must have "learned" results
      If (tdlSyntax.Tables("Const").Rows.Count = 0) Then Status("First do Syntax/Learn") : Exit Sub
      ' Get the current or the first forest of this file
      If (Not GetCurrentForest(ndxForest, ndxThis)) Then Status("Could not find first forest") : Exit Sub
      ' Open the correct tab page
      Me.TabControl1.SelectedTab = Me.tpTree
      ' Try making suggestions until we are ready
      Do
        ' Try making automatic adjustments to what we have right now
        If (Not SyntaxTryMOS(ndxForest)) Then Status("Error while trying the MOS system") : Exit Sub
        ' Make sure re-painting etc is done properly
        If (Not DoEtreeHandle(ndxForest)) Then Status("DoEtreeHandle problem") : Exit Sub
        ' Clear the list to be presented to the user
        ClearTable(tdlSyntax.Tables("User"))
        ' Get a list of all the TOP constituents of the forest: all <eTree> nodes
        ndxList = ndxForest.SelectNodes("./child::eTree")
        ' Try from maximum to minimum frame-size
        For intI = ndxList.Count To 1 Step -1
          ' Move the frame where it is possible to be moved
          For intJ = 0 To ndxList.Count - intI
            ' Allow for interrupts
            If (bInterrupt) Then bInterrupt = False : Status("Interrupted") : Exit Sub
            ' Get the frame of [intI] children
            strChildren = GetChildLabels(ndxList, intJ, intI)
            ' Check if this is somewhere on the "learned" constituents
            ' Get a list of [Labels] in decreasing likelihood
            If (SyntaxConst(strChildren, intI, colLabel)) Then
              ' Walk all the results
              For intK = 0 To colLabel.Count - 1
                ' Add the information to my growing list to be presented to the user
                dtrNew = AddOneDataRow(tdlSyntax, "User", "UserId", "UserList")
                With dtrNew
                  .Item("Label") = colLabel.Item(intK)
                  .Item("ConstId") = colLabel.Exmp(intK)
                  .Item("Children") = SelectedChildLabels(ndxList, intJ, intI)
                  .Item("Html") = SelectedChildLabelsHtml(ndxList, intJ, intI)
                  .Item("Start") = intJ
                  .Item("Len") = intI
                End With
              Next intK
            End If
          Next intJ
        Next intI
        ' Ask the user to choose one of the combinations we came up with
        ' (or perhaps make one himself?)
        With frmSyntax
          ' Show the dialog
          Select Case .ShowDialog
            Case Windows.Forms.DialogResult.Yes, Windows.Forms.DialogResult.OK
              ' Retrieve the result
              intUserId = .UserId
              dtrFound = tdlSyntax.Tables("User").Select("UserId=" & intUserId)
              If (dtrFound.Length > 0) Then
                ' Get start and length
                intJ = dtrFound(0).Item("Start")
                intI = dtrFound(0).Item("Len")
                ' Insert the "Label" node above the leftmost one of the selection
                ndxThis = ndxList(intJ)
                If (Not eTreeInsertLevel(ndxThis, ndxNew)) Then Status("Could not [Insert level]") : Exit Sub
                ' If (Not DoEtreeHandle(ndxNew)) Then Status("DoEtreeHandle problem") : Exit Sub
                ndxNew.Attributes("Label").Value = dtrFound(0).Item("Label")
                ' Walk all subsequently following selected nodes
                For intK = intJ + 1 To intJ + intI - 1
                  ' Get this node
                  ndxThis = ndxList(intK)
                  ' Put it under the node to my left
                  If (Not eTreeJoinUnder(ndxThis, ndxNew, False)) Then Status("Could not [Join under left]") : Exit Sub
                  ' If (Not DoEtreeHandle(ndxNew)) Then Status("DoEtreeHandle problem") : Exit Sub
                Next intK
                ' Re-do this sentence
                eTreeSentence(ndxThis, ndxNew)
                ' Make sure re-painting etc is done properly
                If (Not DoEtreeHandle(ndxNew)) Then Status("DoEtreeHandle problem") : Exit Sub
                ' Add to the frequency of this item
                dtrFound = tdlSyntax.Tables("Const").Select("ConstId=" & dtrFound(0).Item("ConstId"))
                If (dtrFound.Length > 0) Then
                  With dtrFound(0)
                    If (.Item("Freq").ToString <> "") Then
                      .Item("Freq") += 1
                    End If
                  End With
                End If
                ' Yes, we changed something
                bChanged = True
              End If
              ' TODO: Process the result in the current structure
              'Stop
            Case Else
              ' We need to stop
              Status("User stopped") : Exit Sub
          End Select
        End With
      Loop Until ndxList.Count = 1
      ' Show where we are
      Status("Current line has been processed")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxSuggest error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxPartialTree_Click
  ' Goal:   Find out what the best-fit partial tree is with the currently selected constituent in the center
  ' History:
  ' 14-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxPartialTree_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxPartialTree.Click
    Dim strLabel As String = ""         ' Label of currently selected nocee
    Dim strPartial As String = ""       ' Description of the match
    Dim arPartial() As String           ' Array
    Dim arLine() As String              ' Line
    Dim pntThis As Point = Nothing      ' Where we are
    Dim ndxForest As XmlNode = Nothing  ' Currently selected forest
    Dim ndxThis As XmlNode = Nothing    ' Working node
    Dim colHtml As New StringColl       ' Result
    Dim intI As Integer                 ' Counter

    Try
      ' Validate: a file must be loaded
      If (pdxCurrentFile Is Nothing) Then Status("First load a file") : Exit Sub
      ' Validate: we must have "learned" results
      If (tdlSyntax.Tables("Const").Rows.Count = 0) Then Status("First do Syntax/Learn") : Exit Sub
      ' Is the correct tab page opened?
      If (Me.TabControl1.SelectedTab.Name <> Me.tpTree.Name) Then Status("You need to select a constituent in the Tree tab page") : Exit Sub
      ' Get the current forest
      If (Not GetCurrentForest(ndxForest, ndxThis)) Then Status("Cannot find first forest") : Exit Sub
      ' Get selected constituent
      ndxThis = SelectedTreeNode(litTree, strLabel, pntThis)
      If (ndxThis Is Nothing) Then Status("First select a constituent") : Exit Sub
      If (ndxThis.Name <> "eTree") Then Status("You have to select an <eTree> constituent") : Exit Sub
      ' Get the tree's label
      strLabel = ndxThis.Attributes("Label").Value
      ' We have a constituent selected; now find all the locations where it occurs and note the tree elements it has over there
      If (Not GetPartialMatch(strLabel, strPartial)) Then Status("Could not get partial match") : Exit Sub
      ' Show the match we have
      arPartial = Split(strPartial, ";") : colHtml.Add("<html><body><table><tr><td>Path</td><td>Label</td></tr>")
      For intI = 0 To arPartial.Length - 1
        If (arPartial(intI) <> "") Then
          arLine = Split(arPartial(intI), ":")
          colHtml.Add("<tr><td>" & arLine(0) & "</td><td>" & arLine(1) & "</td></tr>")
        End If
      Next intI
      colHtml.Add("</table></body></html>")
      Me.wbReport.DocumentText = colHtml.Text
      Me.TabControl1.SelectedTab = Me.tpReport
      Status("Ready")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxPartialTree error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxPosToConllX_Click
  ' Goal:   Convert Dutch POS format to ConLLX
  ' History:
  ' 02-09-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxPosToConllX_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxPosToConllX.Click
    Dim strDirIn As String = "" ' Input directory
    Dim strFileIn As String     ' Input file
    Dim strFileOut As String    ' Output
    Dim arFile() As String      ' Input files
    Dim strLang As String = ""  ' language
    Dim strResult As String     ' One result
    Dim colTable As New StringColl ' Resulting ConLLX table
    Dim intI As Integer         ' Counter
    Dim bDoCode As Boolean      ' Do CODE nodes or not?

    Try
      ' Initialize directory
      strDirIn = strLastDepDir
      If (strDirIn = "") Then strDirIn = strWorkDir
      ' Get the correct input directory
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, "Directory containing POS-tagged 4-column Dutch text files", strDirIn)) Then
        Status("Could not open input directory")
        Exit Sub
      End If
      ' Check what we returned
      If (Not IO.Directory.Exists(strDirIn)) Then Exit Sub
      ' Save this directory, if it is different from what we have
      If (strLastDepDir <> strDirIn) Then
        strLastDepDir = strDirIn
        SetTableSetting(tdlSettings, "LastDepDir", strLastDepDir)
        XmlSaveSettings(strSetFile)
      End If
      ' Determine the language from the directory
      If (Not DeriveLanguage(strDirIn, strLang)) Then
        ' Assume this is DUTCH
        strLang = "nl"
        'Status("Cannot determine the language")
        'Exit Sub
      End If
      ' Get all the necessary files in this directory (plus subdirectories)
      arFile = IO.Directory.GetFiles(strDirIn, "*.txt", IO.SearchOption.AllDirectories)
      bInterrupt = False : bDoCode = False
      ' Walk the files
      For intI = 0 To arFile.Length - 1
        ' Get this file
        strFileIn = arFile(intI) : strResult = ""
        ' Get the dependency for this file
        If (Not OneDutchPosToConllX(strFileIn, strLang, strResult, bDoCode)) Then
          ' Warn the user and exit
          Status("Dependency: Could not process file: " & strFileIn)
          Exit Sub
        End If
        ' Determine the name for the output file
        strFileOut = IO.Path.GetDirectoryName(strFileIn) & "\" & _
          IO.Path.GetFileNameWithoutExtension(strFileIn) & ".conll"
        ' Save the result
        IO.File.WriteAllText(strFileOut, strResult)
        Logging("CONLLX table saved in: " & strFileOut)
        Status("CONLLX table saved in: " & strFileOut)
      Next intI
      ' Show we are ready
      Status("done")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxDependency error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxDependencyBatch_Click
  ' Goal:   Add dependency relations to the files in one directory
  '         Then produce one global CONLL table
  ' Note:   On request, the dependency relations are added right into the <eTree> nodes as a feature
  ' History:
  ' 24-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxDependencyBatch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxDependencyBatch.Click
    Dim strDirIn As String = "" ' Input directory
    Dim strFileIn As String     ' Input file
    Dim strFileOut As String    ' Output
    Dim arFile() As String      ' Input files
    Dim strLang As String = ""  ' language
    Dim strResult As String     ' One result
    Dim colTable As New StringColl ' Resulting ConLLX table
    Dim intI As Integer         ' Counter
    Dim bDoCode As Boolean      ' Do CODE nodes or not?
    Dim bAddInEtree As Boolean = False ' Add dependency info in <eTree> (I mean: save that info?)

    Try
      ' Initialisation
      bInterrupt = False
      ' Initialize directory
      strDirIn = strLastDepDir
      If (strDirIn = "") Then strDirIn = strWorkDir
      ' Get the correct input directory
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, "Directory containing files that need dependency marking", strDirIn)) Then
        Status("Could not open input directory")
        Exit Sub
      End If
      ' Check what we returned
      If (Not IO.Directory.Exists(strDirIn)) Then Exit Sub
      ' Save this directory, if it is different from what we have
      If (strLastDepDir <> strDirIn) Then
        strLastDepDir = strDirIn
        SetTableSetting(tdlSettings, "LastDepDir", strLastDepDir)
        XmlSaveSettings(strSetFile)
      End If
      ' Determine the language from the directory
      If (Not DeriveLanguage(strDirIn, strLang)) Then
        Status("Cannot determine the language")
        Exit Sub
      End If
      ' Ask for <bAddInEtree>
      Select Case MsgBox("Add dependency information in the .psdx files?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Cancel
          ' Leave!!
          Exit Sub
        Case MsgBoxResult.Yes
          ' Change default value
          bAddInEtree = True
      End Select
      ' Get all the necessary files in this directory (plus subdirectories)
      arFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories)
      bInterrupt = False : bDoCode = False
      ' Walk the files
      For intI = 0 To arFile.Length - 1
        ' Get this file
        strFileIn = arFile(intI) : strResult = ""
        ' Get the dependency for this file
        If (Not OneDepFile(strFileIn, strLang, strResult, bDoCode)) Then
          ' Warn the user and exit
          Status("Dependency: Could not process file: " & strFileIn)
          bInterrupt = False
          Exit Sub
        End If
        ' Need to save??
        If (bAddInEtree) Then
          ' Save the adapted psdx file
          pdxCurrentFile.Save(strFileIn)
        End If
        ' Add the result
        colTable.Add(strResult)
      Next intI
      '' Add dependency to the CURRENT file
      'If (Not DoCurrentDep()) Then Status("Problem determining dependency") : Exit Sub
      ' Save the CONLL table
      strFileOut = strLastDepDir & "\Dep_" & strLang & ".conll"
      ' Make sure CRLF is changed to LF
      IO.File.WriteAllText(strFileOut, colTable.Text)
      ' Show we are ready
      Logging("CONLLX table saved in: " & strFileOut)
      Status("CONLLX table saved in: " & strFileOut)
      ' Should we produce a new chechen.mco?
      Select Case MsgBox("Would you like me to make a new version of Chechen.mco?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.No, MsgBoxResult.Cancel
          Status("Ready")
          Exit Sub
      End Select
      ' Yes, we may make a new version...
      If (Not CreateMaltMco(strFileOut, strLang)) Then
        Status("There was a problem creating the .mco file")
      Else
        Status("Ready")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxDependencyBatch error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxRepairLezgi_Click
  ' Goal:   Take out spaces where they shouldn't be
  ' History:
  ' 19-03-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxRepairLezgi_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxRepairLezgi.Click
    Dim ndxList As XmlNodeList  ' List of forest files
    Dim intI As Integer         ' Counter
    Dim strLang As String = "lez"

    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Status("First load a psdx file") : Exit Sub
      ' Initialize the language
      If (Not DoLangInit(strLang)) Then Status("Could not initialize language [" & strLang & "]") : Exit Sub
      ' Walk all the forest files
      ndxList = pdxCurrentFile.SelectNodes("//forest")
      For intI = 0 To ndxList.Count - 1
        ' Repair this forest node
        If (Not DoRepairOneLezgi(ndxList(intI), strLang)) Then Status("Problem repairing forest node " & intI + 1) : Exit Sub
      Next intI
      ' Flag we need saving
      bDirty = True
      ' Do re-analysis
      mnuSyntaxReanalyze_Click(sender, e)
      ' Show we are ready
      Status("Ready")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxRepairLezgi error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxLezgiPos_Click
  ' Goal:   Perform POS tagging for Lezgi
  '         (1) Build a list of Vern-POS matches in this text
  '         (2) Check all constituents semi-automatically
  ' History:
  ' 19-03-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxLezgiPos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxLezgiPos.Click
    Try
      If (SyntaxPosTag("lez")) Then
        Status("Ready")
      Else
        Status("Could not POS-tag")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxLezgiPos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxLakPos_Click
  ' Goal:   Perform POS tagging for Lak
  '         (1) Do POS tagging using a <liftpos> dictionary on Cyrillic Lak
  '         (2) Build a list of Vern-POS matches in this text
  '         (3) Check all constituents semi-automatically
  ' History:
  ' 17-07-2015  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxLakPos_Click(sender As System.Object, e As System.EventArgs) Handles mnuSyntaxLakPos.Click
    Try
      If (SyntaxPosTag("lbe-Cyr", bAuto:=True, bDoAll:=True)) Then
        Status("Ready")
      Else
        Status("Could not POS-tag")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxLakPos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxLakPosBatch_Click
  ' Goal:   Perform POS tagging for Lak on a series of files
  ' History:
  ' 17-07-2015  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnySyntaxLakPosBatch_Click(sender As System.Object, e As System.EventArgs) Handles mnySyntaxLakPosBatch.Click
    Dim strDirIn As String = "d:\data files\corpora\lak\xml\PCMLBE"
    Dim arInFile() As String
    Dim intPtc As Integer

    Try
      ' Request directory
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, _
                         "Directory that contains the LAK .psdx files", strDirIn)) Then Exit Sub
      ' Double check
      If (Not IO.Directory.Exists(strDirIn)) Then Status("Could not open directory [" & strDirIn & "]") : Exit Sub
      ' Consider all *.psdx input files
      arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx", System.IO.SearchOption.AllDirectories)
      For intI = 0 To arInFile.Count - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ arInFile.Count
        Status("Processing " & intPtc & "%", intPtc)
        ' Get the input file name
        strCurrentFile = arInFile(intI)
        ' Show where we are
        Logging("Tagging " & intI + 1 & "/" & arInFile.Count & " - " & IO.Path.GetFileNameWithoutExtension(strCurrentFile))
        ' Read the file as xml
        If (ReadXmlDoc(strCurrentFile, pdxCurrentFile)) Then
          ' Perform POS tagging automatically on thisfile
          If (Not SyntaxPosTag("lbe-Cyr", bAuto:=True, bDoAll:=True)) Then
            ' SOme error
            Logging("Could not do POS tagging on: " & strCurrentFile)
          End If
          ' Save the result
          pdxCurrentFile.Save(strCurrentFile)
        Else
          Logging("Could not read file: " & strCurrentFile)
        End If
      Next intI
      ' Show we are done
      Status("Ready doing POS tagging on Lak files.")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxLakPosBatch error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxRepairDutch_Click
  ' Goal:   Repair the xml structure of Dutch files:
  '         - Remove a top "NODE"
  '         - Remove a "VP" layer
  '         - Append +1 top-nodes to first top-node
  ' History:
  ' 27-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxRepairDutch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxRepairDutch.Click
    Dim ndxList As XmlNodeList  ' List of forest files
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Status("First load a psdx file") : Exit Sub
      ' Walk all the forest files
      ndxList = pdxCurrentFile.SelectNodes("//forest")
      For intI = 0 To ndxList.Count - 1
        ' Repair this forest node
        If (Not DoRepairOneDutch(ndxList(intI))) Then Status("Problem repairing forest node " & intI + 1) : Exit Sub
      Next intI
      ' Flag we need saving
      bDirty = True
      ' Do re-analysis
      mnuSyntaxReanalyze_Click(sender, e)
      ' Show we are ready
      Status("Ready")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxDutchRepair error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxRepairDiderot_Click
  ' Goal:   Repair Diderot files for the Longdale project (January 2014)
  ' History:
  ' 10-01-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxRepairDiderot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxRepairDiderot.Click
    Dim strFile As String = ""  ' One file
    Dim arList() As String      ' Array of files
    Dim strDir As String = "d:\data files\corpora\english\longdale\diderot\xml\"
    Dim ndxList As XmlNodeList  ' List of forest files
    Dim ndxEtree As XmlNodeList ' List of nodes
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter
    Dim intK As Integer         ' Counter
    Dim intId As Integer = 1    ' First @Id value
    Dim intPtc As Integer       ' Percentage

    Try
      ' Find all psdx files in the Diderot directory
      arList = IO.Directory.GetFiles(strDir, "*.psdx", IO.SearchOption.TopDirectoryOnly)
      ' Visit them all
      For intI = 0 To arList.Length - 1
        ' Access this file
        strFile = arList(intI)
        ' Show where we are
        intPtc = (intI + 1) * 100 \ arList.Length
        Status("Diderot Repair [" & IO.Path.GetFileNameWithoutExtension(strFile) & "] " & intPtc & "%", intPtc)
        ' Read this file as the current one
        strCurrentFile = strFile : pdxCurrentFile = Nothing
        If (Not ReadXmlDoc(strCurrentFile, pdxCurrentFile)) Then Status("Could not load " & strFile) : Exit Sub
        ' Walk all the forest files
        ndxList = pdxCurrentFile.SelectNodes("//forest")
        For intJ = 0 To ndxList.Count - 1
          ' Repair this forest node
          If (Not DoRepairOneDiderot(ndxList(intJ))) Then Status("Problem repairing forest node " & intJ + 1) : Exit Sub
        Next intJ
        ' Adapt the @Id values
        intId = 1
        For intJ = 0 To ndxList.Count - 1
          ' Repair this forest node
          ndxEtree = ndxList(intJ).SelectNodes("./descendant::eTree")
          For intK = 0 To ndxEtree.Count - 1
            ' Renumber
            ndxEtree(intK).Attributes("Id").Value = intId : intId += 1
          Next intK
        Next intJ
        ' Save this file
        pdxCurrentFile.Save(strCurrentFile)
      Next intI
      ' Show we are ready
      Status("Ready")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxDiderotRepair error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxShowPOS_Click
  ' Goal:   Show the POS of this psdx text in CONLLX format
  ' History:
  ' 02-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxShowPOS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxShowPOS.Click
    Dim strInFile As String       ' Current input file
    Dim strConFile As String      ' CONLL-X output file
    Dim strLang As String = "en"  ' Language

    Try
      ' Check if a file has been loaded
      If (pdxCurrentFile Is Nothing) Then Status("First load a psdx file") : Exit Sub
      ' Reset interrupt
      bInterrupt = False
      ' Get the language
      If (Not DeriveLanguage(IO.Path.GetDirectoryName(strCurrentFile), strLang, True)) Then Exit Sub
      ' We have a clean, unparsed psdx file.
      ' (1) Convert psdx to CONLL-X format
      strInFile = strCurrentFile : strConFile = IO.Path.GetDirectoryName(strInFile) & "\" & IO.Path.GetFileNameWithoutExtension(strInFile) & ".conll"
      If (DoExport("ConLLXpos", strConFile)) Then
        ' Show the resulting [strConFile] in the report tab page
        Me.wbReport.DocumentText = ConllToHtml(strConFile)
        Me.TabControl1.SelectedTab = Me.tpReport
        ' Show we are ready
        Status("Ready")
      Else
        ' Show we are ready
        Status("There was an error")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxShowPOS error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxParse_Click
  ' Goal:   Syntactically parse the current psdx file
  '         Steps:
  ' History:
  ' 17-09-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxParse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxParse.Click
    Try
      DoSyntaxParse(True, True)
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxParse error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxDependency_Click
  ' Goal:   Convert to dependency format (held within psdx file)
  ' History:
  ' 19-06-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxDependency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxDependency.Click
    Try
      DoSyntaxParse(True, False)
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxDependency error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxConstituency_Click
  ' Goal:   Convert dependency to constituency (within psdx file)
  ' History:
  ' 19-06-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxConstituency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxConstituency.Click
    Try
      DoSyntaxParse(False, True)
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxConstituency error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxParse_Click
  ' Goal:   Syntactically parse the current psdx file
  '         Steps:
  ' History:
  ' 17-09-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub DoSyntaxParse(ByVal bDoDep As Boolean, ByVal bDepToConst As Boolean)
    Dim strInFile As String       ' Current input file
    Dim strConFile As String      ' CONLL-X output file
    Dim strLang As String = "en"  ' Language

    Try
      ' Check if a file has been loaded
      If (pdxCurrentFile Is Nothing) Then Status("First load a psdx file") : Exit Sub
      ' Reset interrupt
      bInterrupt = False
      ' Get the language
      If (Not DeriveLanguage(IO.Path.GetDirectoryName(strCurrentFile), strLang, True)) Then Status("Could not derive language") : Exit Sub
      ' Perform language initialisation
      If (Not DoLangInit(strLang)) Then Status("Could not initialize language [" & strLang & "]") : Exit Sub
      ' Do we need the first step?
      If (bDoDep) Then
        ' Check if parsing has occurred already
        If (pdxCurrentFile.SelectSingleNode("//forest/child::eTree/child::eTree") IsNot Nothing) Then
          ' Ask how to proceed
          Select Case MsgBox("This file has been (partially) parsed already." & vbCrLf & _
                             "Would you like to delete the current parsing and re-parse it?", MsgBoxStyle.YesNoCancel)
            Case MsgBoxResult.No, MsgBoxResult.Cancel
              Status("Exited without parsing")
              Exit Sub
            Case MsgBoxResult.Yes
              ' Delete the current parsing...
              If (Not ParseOneDelete(pdxCurrentFile)) Then
                ' TODO: implement this
                Status("Sorry, deleting current parse did not work out")
                Exit Sub
              End If
          End Select
        End If
        ' We have a clean, unparsed psdx file.
        ' (1) Convert psdx to CONLL-X format
        strInFile = strCurrentFile : strConFile = IO.Path.GetDirectoryName(strInFile) & "\" & IO.Path.GetFileNameWithoutExtension(strInFile) & ".conll"
        If (DoExport("ConLLXpos", strConFile)) Then
          ' Use the maltparser to get the dependencies filled in, within the same CONLL file
          If (ParseOneMalt(strConFile, strLang)) Then
            ' Process the dependencies into the psdx file
            If (ExpandOneConllToPsdx(strConFile, strLang, False)) Then
              ' Stop
              ' Okay, everything went fine
              Status("Current file has been parsed")
            Else
              ' COming here means failure
              Status("Sorry, could not parse the current psdx file")
            End If
          End If
        End If
      End If
      ' Need to do Dep to constituency?
      If (bDepToConst) Then
        If (ExpandOneConllDepToPsdx(strLang)) Then
          Status("Current file has been converted to constituency")
        Else
          ' COming here means failure
          Status("Sorry, could not parse the current psdx file")
        End If
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxParse error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxDelParse_Click
  ' Goal:   Delete syntactic parse
  ' History:
  ' 25-09-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxDelParse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxDelParse.Click
    Try
      ' Check if a file has been loaded
      If (pdxCurrentFile Is Nothing) Then Status("First load a psdx file") : Exit Sub
      ' Reset interrupt
      bInterrupt = False
      ' Attempt deleting
      If (ParseOneDelete(pdxCurrentFile)) Then
        Status("Parse has been deleted")
      Else
        Status("Could not delete current parsing")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxDelParse error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxBatchCorrectId_Click
  ' Goal:   Correct Etree @Id values for all the files in a particular directory
  ' History:
  ' 21-08-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxBatchCorrectId_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxBatchCorrectId.Click
    Dim strDirIn As String = ""   ' input directory
    Dim strDirOut As String = ""  ' output directory
    Dim strInFile As String       ' One input file
    Dim strOutFile As String      ' One output file
    Dim arInFile() As String      ' Array of input files
    Dim intI As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage
    Dim bOverwrite As Boolean = False    ' Whether we may overwrite or not

    Try
      ' Initialise the directories myself
      strDirIn = "d:\data files\corpora\chechen\xml\NMSU"
      ' Request directory
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, "Directory that contains the .psdx files", strDirIn)) Then Exit Sub
      ' Double check
      If (Not IO.Directory.Exists(strDirIn)) Then Status("Could not open directory [" & strDirIn & "]") : Exit Sub
      ' Consider all *.psdx input files
      arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx", System.IO.SearchOption.AllDirectories)
      For intI = 0 To arInFile.Count - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ arInFile.Count
        Status("Processing " & intPtc & "%", intPtc)
        ' Get the input file name
        strCurrentFile = arInFile(intI)
        ' Show where we are
        Logging("Correcting " & intI + 1 & "/" & arInFile.Count & " - " & IO.Path.GetFileNameWithoutExtension(strCurrentFile))
        ' Read the file as xml
        If (ReadXmlDoc(strCurrentFile, pdxCurrentFile)) Then
          If (Not InitCurrentFile()) Then Status("Could not initialize file " & strCurrentFile) : Exit Sub
          ' Perform the correction
          If (Not AdaptEtreeId(0)) Then Status("Could not adapt etree id's in " & strCurrentFile) : Exit Sub
          ' Save the result
          pdxCurrentFile.Save(strCurrentFile)
        Else
          Logging("Could not read file: " & strCurrentFile)
        End If
      Next intI
      ' Show we are done
      Status("Ready correcting <eTree> @Id values.")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxBatchCorrectId error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxBatchPOS_Click
  ' Goal:   Convert psdx files into a list of vernacular words and their POS-tags
  ' History:
  ' 04-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuSyntaxBatchPOS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxBatchPosSimple.Click
    Try
      MakePOSfiles("PosSimple")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxBatchPOSsimple error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub mnuSyntaxBatchPosOrg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxBatchPosOrg.Click
    Try
      MakePOSfiles("PosOrg")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxBatchPOSorg error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub mnuSyntaxBatchPosPTB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxBatchPosPTB.Click
    Try
      MakePOSfiles("PosPTB")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxBatchPosPTB error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub mnuSyntaxTokenize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSyntaxTokenize.Click
    Try
      MakePOSfiles("Token")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmMain/SyntaxTokenize error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MakePOSfiles
  ' Goal:   Convert psdx files into a list of vernacular words and their POS-tags
  ' History:
  ' 04-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub MakePOSfiles(ByVal strType As String)
    Dim strDirIn As String = ""   ' input directory
    Dim strDirOut As String = ""  ' output directory
    Dim strInFile As String       ' One input file
    Dim strOutFile As String      ' One output file
    Dim arInFile() As String      ' Array of input files
    Dim intI As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage
    Dim bOverwrite As Boolean = False    ' Whether we may overwrite or not

    Try
      ' Check if automachine is busy
      If (bAutoBusy) Then
        MsgBox("First stop auto coreference resolution (F11)")
        Exit Sub
      End If
      ' Let user select an input and output directory
      With frmDirs
        ' Initialise the directories myself
        .SrcDir = "d:\data files\corpora\chechen\xml\Mytask"
        .DstDir = .SrcDir & "\pos"
        ' If the directory does not exist, then make it
        If (Not IO.Directory.Exists(.DstDir)) Then IO.Directory.CreateDirectory(.DstDir)
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Get the source and destination directories
            strDirIn = .SrcDir
            strDirOut = .DstDir
            ' Check whether correct directories were selected
            If (IO.Directory.Exists(strDirIn)) AndAlso (IO.Directory.Exists(strDirOut)) Then
              ' Consider all *.psdx input files
              arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx")
              For intI = 0 To arInFile.Count - 1
                ' Show where we are
                intPtc = (intI + 1) * 100 \ arInFile.Count
                Status("Processing " & intPtc & "%", intPtc)
                ' Get the input file name
                strInFile = arInFile(intI)
                ' Determine the output file name
                Select Case strType
                  Case "Token"
                    strOutFile = strDirOut & "\" & IO.Path.GetFileNameWithoutExtension(strInFile) & _
                      ".tok"
                  Case Else
                    strOutFile = strDirOut & "\" & IO.Path.GetFileNameWithoutExtension(strInFile) & _
                      ".pos"
                End Select
                ' Show where we are
                Logging("Producing " & strType & " " & intI + 1 & "/" & arInFile.Count & " - " & IO.Path.GetFileNameWithoutExtension(strInFile))
                ' Check existence
                If (Not bOverwrite) AndAlso (IO.File.Exists(strOutFile)) Then
                  ' Ask if overwriting is okay
                  If (MsgBox("Can I overwrite files like: [" & strOutFile & "]?", MsgBoxStyle.YesNoCancel) <> MsgBoxResult.Yes) Then
                    ' No overwriting, so exit 
                    Exit Sub
                  End If
                  ' We can overwrite, so put to true
                  bOverwrite = True
                End If
                ' Call the correct function to produce PSD
                If (Not OnePsdFile(strInFile, strOutFile, strType)) Then
                  ' Check for interrupt
                  If (bInterrupt) Then Status("Interrupted...")
                  ' There was an error, so stop
                  Exit Sub
                End If
              Next intI
            End If
        End Select
      End With
      ' Show we are done
      Status("Ready producing POS files.")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsPsd_Click error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsMEprojDictME_Click
  ' Goal:   Convert the .txt file with the OE dictionary into a .sfm file with the correct standard
  '           format markers
  ' History:
  ' 14-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsMEprojDictME_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMEprojDictME.Click
    Try
      ' Convert the HTML dictionary into an xml one
      If (Not MorphDictHtmlToXml()) Then
        Status("Could not convert HTML dictionary to XML")
      Else
        Status("Ready")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ToolsMEprojDictME error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsMEprojMorphdict_Click
  ' Goal:   Convert the Gutenberg .xml file into a MorphDict
  ' History:
  ' 03-10-2013  ERK Created
  ' 21-01-2014  ERK Added "MED' conversion: all small sections into one
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsMEprojMorphdict_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMEprojMorphdict.Click
    Dim strLngAbbr As String = ""     ' Language abbreviation
    Dim strMdict As String = ""       ' Morph dict we need to look for
    Dim bNewMethod As Boolean = True  ' Special purpose switch
    Dim bNewMed As Boolean = False    ' Make new MED dictionary

    Try
      ' Get the language
      If (Not EnglishPeriod(strLngAbbr)) Then Exit Sub
      ' Convert the HTML dictionary into an xml one
      Select Case strLngAbbr
        Case "OEB"  ' Old English, Bossworth & Toller version
          ' TODO: implementeren!!
          If (bNewMethod) Then
            ' Set the Morphology dictionary name
            strMdict = "MorphDictOE.xml"
            ' Load the existing MorphDict
            MorphDictIni(False, strMdict)
            If (tdlMorphDict Is Nothing) Then Status("Could not load " & strLngAbbr & " dictionary") : Exit Sub
            ' Update the *verbs* in the existing MorphDict with the MED information
            If (Not MorphDictAdaptBT("Verbs")) Then
              Status("Could not adapt MorphDict with B&T information")
            Else
              Status("Ready")
            End If
          End If
        Case "ME"
          If (Not MorphDictXmlToTdl("ME")) Then Status("Could not convert XML dictionary to MorphDict") : Exit Sub
        Case "MED"
          If (bNewMethod) Then
            ' Set the Morphology dictionary name
            strMdict = "MorphDictME.xml"
            ' Do we need to make a new MED dictionary? (expecting [MEDcombinedRes.xml])
            If (bNewMed) Then
              If (Not MEDextend(Me.wbReport)) Then Status("Could not extend the MED dictionary") : Exit Sub
            End If
            ' Load the existing MorphDict
            MorphDictIni(False, strMdict)
            If (tdlMorphDict Is Nothing) Then Status("Could not load " & strLngAbbr & " dictionary") : Exit Sub
            ' Update the *verbs* in the existing MorphDict with the MED information
            If (Not MorphDictAdaptMed("Verbs")) Then
              Status("Could not adapt MorphDict with MED information")
            Else
              Status("Ready")
            End If
          Else
            ' Combine all MED dictionary parts into one...
            If (Not CombineMEDsections()) Then Status("Could not combine the MED sections") : Exit Sub
            ' Add the combinations to the <Morph> section
            If (Not MorphDictXmlToTdlMED("ME")) Then Status("Could not combine the MED into MorphDict") : Exit Sub
          End If
        Case "eModE", "LmodE"
          If (Not MorphDictWebsterToTdl()) Then Status("") : Exit Sub
          If (Not MorphDictXmlToTdl("ModE")) Then Status("Could not convert XML dictionary to MorphDict") : Exit Sub
      End Select
      Status("Ready")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ToolsMEprojMorphDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsMEprojMorphToVernPos_Click
  ' Goal:   Convert the "Morph" table to "VernPos"
  ' History:
  ' 04-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsMEprojMorphToVernPos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMEprojMorphToVernPos.Click
    Dim strLngAbbr As String = "" ' Language abbreviation
    Dim strMdict As String = ""   ' Morph dict we need to look for
    Dim intNum As Integer = 0     ' Number of removals

    Try
      ' Get the language
      If (Not EnglishPeriod(strLngAbbr)) Then Exit Sub
      ' Convert the HTML dictionary into an xml one
      Select Case strLngAbbr
        Case "OE"
          strMdict = "MorphDictOE.xml"
        Case "ME"
          strMdict = "MorphDictME.xml"
        Case "eModE", "LmodE"
          strMdict = "MorphDictModE.xml"
        Case Else
          Exit Sub
      End Select

      ' We need to read the correct [tdlMorphDict]
      MorphDictIni(False, strMdict)
      If (tdlMorphDict Is Nothing) Then Status("Could not load " & strLngAbbr & " dictionary") : Exit Sub
      ' Convert the HTML dictionary into an xml one
      If (Not MorphDictCreateVernPos()) Then
        Status("Could not create VernPos table")
      Else
        Status("Ready")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ToolsMEprojMorphToVernPos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsMEprojVposToMorph_Click
  ' Goal:   Add entries from "VernPos" into "Morph"
  ' History:
  ' 25-02-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsMEprojVposToMorph_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMEprojVposToMorph.Click
    Dim strLngAbbr As String = "" ' Language abbreviation
    Dim strMdict As String = ""   ' Morph dict we need to look for
    Dim intNum As Integer = 0     ' Number of removals

    Try
      ' Get the language
      If (Not EnglishPeriod(strLngAbbr)) Then Exit Sub
      ' Convert the HTML dictionary into an xml one
      Select Case strLngAbbr
        Case "OE"
          strMdict = "MorphDictOE.xml"
        Case "ME"
          strMdict = "MorphDictME.xml"
        Case "eModE", "LmodE"
          strMdict = "MorphDictModE.xml"
        Case Else
          Exit Sub
      End Select

      ' We need to read the correct [tdlMorphDict]
      MorphDictIni(False, strMdict)
      If (tdlMorphDict Is Nothing) Then Status("Could not load " & strLngAbbr & " dictionary") : Exit Sub
      ' Add [VernPos] entries into [Morph] where possible
      If (VernToMorph(intNum)) Then
        ' Show where we are
        Status("Added " & intNum & " entries from [VernPos] into [Morph]")
      Else
        Status("There was an error")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ToolsMEprojVernPosToMorph error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsMEunaccSpread_Click
  ' Goal:   Spread the VbType "unacc" feature
  ' Notes:
  '         Process #1: Build a list of lemma's
  '         i)	Find all verbs annotated with VbType = “unacc”
  '         ii)	If they have a lemma specified, then
  '             (1)	Add the form to a list of lemma’s
  ' 
  '         Process #2: Spread unacc feature to all forms of lemma's
  '         i)	Find all verb-forms that have a lemma specified
  '         ii)	If the lemma is in the "unacc" list of lemma's, then
  '             (1)	Change the VbType feature to “unacc”
  ' History:
  ' 13-02-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsMEunaccSpread_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMEunaccSpread.Click
    Dim strLngAbbr As String = "" ' Language abbreviation
    Dim strDirAbbr As String            ' Abbreviation for language directory
    Dim strTagRepFile As String = ""    ' YCOE tagging report file
    Dim strDirIn As String = ""         ' Input directory with YCOE files
    Dim colLemma As New StringColl      ' List of lemma's that should be marked as "unacc"

    Try
      ' Get the language
      If (Not EnglishPeriod(strLngAbbr)) Then Exit Sub
      Select Case strLngAbbr
        Case "OE"
          strDirAbbr = strLngAbbr
        Case "ME"
          strDirAbbr = strLngAbbr
        Case "eModE"
          strDirAbbr = "ModE"
        Case "LmodE"
          strDirAbbr = "MBE"
        Case Else
          Exit Sub
      End Select
      ' Let user select an input directory
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, "Give directory where unaccusative spreading for " & strLngAbbr & " should take place", _
        "d:\data files\corpora\English\xml\Adapted\" & strDirAbbr)) Then Exit Sub
      ' Process #1: build a list of unacc lemma's
      If (Not UnaccLemmaCollect(strDirIn, colLemma)) Then Status("There was an error gathering the unacc lemma's") : Exit Sub
      ' Make a report with these lemma's
      strTagRepFile = GetDocDir() & "\UnaccLemmaReport" & strLngAbbr & ".htm"
      IO.File.WriteAllText(strTagRepFile, colLemma.Html)
      ' Put the report in place, by navigating to the file
      Status("Opening the " & strLngAbbr & "-unacc-lemma report...")
      Me.wbReport.Navigate(strTagRepFile)
      ' Process #2: Spread unacc feature to all forms of lemma's
      If (Not UnaccLemmaSpread(strDirIn, colLemma)) Then Status("There was an error spreading the unacc lemma's") : Exit Sub
      ' Show we are ready
      Status("Ready")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ToolsMEunaccSpread error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsMEprojReportTagged_Click
  ' Goal:   Report how much is tagged of the ME corpus finite verbs
  ' History:
  ' 04-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsMEprojReportTagged_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMEprojReportTagged.Click
    Dim strLngAbbr As String = "" ' Language abbreviation
    Dim strDirAbbr As String            ' Abbreviation for language directory
    Dim strTagRepFile As String = ""    ' YCOE tagging report file
    Dim strDirIn As String = ""         ' Input directory with YCOE files
    Dim strMdict As String = ""
    Dim strDoLabels As String = strAnyVerb  ' Limit myself to the verbs
    Dim strIniDir As String = "d:\data files\corpora\English\xml\Adapted\"        ' Directory where we start

    Try
      ' Get the language
      If (Not EnglishPeriod(strLngAbbr)) Then Exit Sub
      Select Case MsgBox("OE: only verbs (Y) or all (N)?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Cancel
          Exit Sub
        Case MsgBoxResult.No
          strDoLabels = ""
      End Select
      Select Case strLngAbbr
        Case "OE"
          strDirAbbr = strLngAbbr : strIniDir &= "OE"
        Case "OED"
          strDirAbbr = strLngAbbr : strMdict = "MorphDictOE.xml" : strIniDir &= "OE"
          ' We need to read the [tdlMorphDict]
          MorphDictIni(False, strMdict)
          If (tdlMorphDict Is Nothing) Then Status("Could not load " & strLngAbbr & " dictionary") : Exit Sub
        Case "OEB"
          strDirAbbr = strLngAbbr : strMdict = "MorphDictOE.xml" : strIniDir &= "OE"
          ' We need to read the [tdlMorphDict]
          MorphDictIni(False, strMdict)
          If (tdlMorphDict Is Nothing) Then Status("Could not load " & strLngAbbr & " dictionary") : Exit Sub
        Case "ME"
          strDirAbbr = strLngAbbr : strIniDir &= "ME"
        Case "MED"
          strDirAbbr = strLngAbbr : strMdict = "MorphDictME.xml" : strIniDir &= "ME"
          ' We need to read the [tdlMorphDict]
          MorphDictIni(False, strMdict)
          If (tdlMorphDict Is Nothing) Then Status("Could not load " & strLngAbbr & " dictionary") : Exit Sub
        Case "eModE"
          strDirAbbr = "ModE" : strIniDir &= "eModE"
        Case "LmodE"
          strDirAbbr = "MBE" : strIniDir &= "LmodE"
        Case Else
          Exit Sub
      End Select
      ' Let user select an input directory
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, "Give directory where morphologically enriched " & strLngAbbr & " files are located", _
        strIniDir)) Then Exit Sub
      If (MakeTaggingReport(strDirIn, strTagRepFile, strDoLabels, strLngAbbr)) Then
        ' Is this MED?
        If (strLngAbbr = "MED") OrElse (strLngAbbr = "OED") Then
          ' Save
          Status("Saving: " & strMorphDictFile)
          tdlMorphDict.WriteXml(strMorphDictFile)
        End If
        ' Put the report in place, by navigating to the file
        Status("Opening the " & strLngAbbr & "-finite-verb tagging report...")
        Me.wbReport.Navigate(strTagRepFile)
        ' Show we are ready
        Status("Ready collecting " & strLngAbbr & "-finite-verb tagging information (see report)")
        Me.TabControl1.SelectedTab = Me.tpReport
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ToolsMEprojTaggedReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsMEprojReportDict_Click
  ' Goal:   Provide an HTML file and report that contains all entries in [Morph] and corresponding lemma's
  ' History:
  ' 18-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsMEprojReportDict_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMEprojReportDict.Click
    Dim strDictRepFile As String = "" ' HTML file for dictionary report

    Try
      If (MakeDictReport("ME", strDictRepFile)) Then
        '' Put the report in place, by navigating to the file
        'Status("Opening the ME dictionary report...")
        'Me.wbReport.Navigate(strDictRepFile)
        ' Show where user can open the report
        Me.wbReport.DocumentText = "<h2>Dictionary</h2><p>Report is located at:<p>" & strDictRepFile & _
          "<p><p>Open that file with Excel"
        ' Show we are ready
        Status("Ready collecting ME dictionary information (see report)")
        Logging("The ME dictionary is located at: " & strDictRepFile)
        Me.TabControl1.SelectedTab = Me.tpReport
      Else
        Status("Something went wrong")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ToolsMEprojReportDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMEprojReportVernpos_Click
  ' Goal:   Provide an HTML file and report that contains all entries in [Vernpos] and corresponding lemma's
  ' History:
  ' 23-01-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuToolsMEprojReportVernpos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMEprojReportVernpos.Click
    Dim strLngAbbr As String = ""     ' Language abbreviation
    Dim strDictRepFile As String = "" ' HTML file for dictionary report
    Dim strDirAbbr As String = ""     '
    Dim strMdict As String = ""       '

    Try
      ' Get the language
      If (Not EnglishPeriod(strLngAbbr)) Then Exit Sub
      Select Case strLngAbbr
        Case "OE"
          strDirAbbr = strLngAbbr
          strMdict = "MorphDictOE.xml"
        Case "ME"
          strDirAbbr = strLngAbbr
          strMdict = "MorphDictME.xml"
        Case "eModE"
          strDirAbbr = "ModE"
          strMdict = "MorphDictModE.xml"
        Case "LmodE"
          strDirAbbr = "MBE"
          strMdict = "MorphDictModE.xml"
        Case Else
          Status("Language not recognized: " & strLngAbbr)
          Exit Sub
      End Select
      ' Initialise morphological dictionary
      MorphDictIni(False, strMdict)
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      ' Make the vernpos dictionary...
      If (MakeVernposReport(strLngAbbr, strDictRepFile)) Then
        ' Show where user can open the report
        Me.wbReport.DocumentText = "<h2>Vernpos</h2><p>Report is located at:<p>" & strDictRepFile & _
          "<p><p>Open that file with Excel"
        ' Show we are ready
        Status("Ready collecting ME vernpos information (see report)")
        Logging("The ME vernpos is located at: " & strDictRepFile)
        Me.TabControl1.SelectedTab = Me.tpReport
      Else
        Status("Something went wrong")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ToolsMEprojReportVernpos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsMorphReportDictOE_Click
  ' Goal:   Provide an HTML file and report that contains all entries in [Morph] and corresponding lemma's
  ' History:
  ' 18-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsMorphReportDictOE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMorphReportDictOE.Click
    Dim strDictRepFile As String = "" ' HTML file for dictionary report

    Try
      If (MakeDictReport("OE", strDictRepFile)) Then
        '' Put the report in place, by navigating to the file
        'Status("Opening the OE dictionary report...")
        'Me.wbReport.Navigate(strDictRepFile)
        ' Show where user can open the report
        Me.wbReport.DocumentText = "<h2>Dictionary</h2><p>Report is located at:<p>" & strDictRepFile & _
          "<p><p>Open that file with Excel"
        ' Show we are ready
        Status("Ready collecting OE dictionary information (see report)")
        Logging("The OE dictionary is located at: " & strDictRepFile)
        Me.TabControl1.SelectedTab = Me.tpReport
      Else
        Status("Something went wrong")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ToolsMEprojReportDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsMEprojReportRemain_Click
  ' Goal:   Report how much remains to be done for the ME corpus finite verbs
  ' History:
  ' 04-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsMEprojReportRemain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMEprojReportRemain.Click
    Dim strLngAbbr As String = ""       ' Language abbreviation
    Dim strDirAbbr As String            ' Abbreviation for language directory (the last part of it)
    Dim strTagRepFile As String = ""    ' ME tagging report file
    Dim strDirIn As String = ""         ' Input directory with YCOE files
    Dim strMdict As String = ""         ' Morphological dictionary name
    Dim strDoLabels As String = strAnyVerb

    Try
      ' Get the language
      If (Not EnglishPeriod(strLngAbbr)) Then Exit Sub
      Select Case MsgBox("OE: only verbs (Y) or all (N)?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Cancel
          Exit Sub
        Case MsgBoxResult.No
          strDoLabels = ""
      End Select
      Select Case strLngAbbr
        Case "OE"
          strDirAbbr = strLngAbbr
          strMdict = "MorphDictOE.xml"
        Case "OEB"
          strDirAbbr = "OE"
          strMdict = "MorphDictOE.xml"
        Case "ME"
          strDirAbbr = strLngAbbr
          strMdict = "MorphDictME.xml"
        Case "eModE"
          strDirAbbr = "ModE"
          strMdict = "MorphDictModE.xml"
        Case "LmodE"
          strDirAbbr = "MBE"
          strMdict = "MorphDictModE.xml"
        Case Else
          Exit Sub
      End Select
      ' Let user select an input directory
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, "Give directory where morphologically enriched " & strLngAbbr & " files are located", _
        "d:\data files\corpora\English\xml\Adapted\" & strDirAbbr)) Then Exit Sub
      ' Initialise morphological dictionary
      MorphDictIni(False, strMdict)
      strMorphRemainFile = IO.Path.GetDirectoryName(strMorphDictFile) & "\MorphDictRemain" & strLngAbbr & ".xml"
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      ' Create a new dataset for remainders
      If (Not CreateDataSet("MorphDict.xsd", tdlRemain)) Then Status("Could not create Remain dataset") : Exit Sub
      '' Clear the table for remainders
      'ClearTable(tdlMorphDict.Tables("Remain"))
      '' Make sure changes are processed
      'Status("processing removal of old rows...")
      'tdlSettings.AcceptChanges()
      'Status("continuing...")
      ' Start making remain report (html)
      If (MakeRemainReport(strDirIn, strTagRepFile, strLngAbbr, strDoLabels)) Then
        ' Put the report in place, by navigating to the file
        Status("Opening the " & strLngAbbr & "-finite-verb remain report...")
        Me.wbReport.Navigate(strTagRepFile)
        ' Show we are ready
        Status("Ready collecting " & strLngAbbr & "-finite-verb remain information (see report)")
        Me.TabControl1.SelectedTab = Me.tpReport
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsMEprojRemainReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsMEprojLemmatizeByVernPos_Click
  ' Goal:   Check all the [strDoTags] in the ME corpus against the information in the "VernPos" table
  '         Add the lemma + features to those that match
  ' Note:   Change the specification of [strDoTags] to choose the @Label values being visited
  ' History:
  ' 04-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsMEprojLemmatizeByVernPos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMEprojLemmatizeByVernPos.Click
    Dim strLngAbbr As String = ""         ' Language abbreviation
    Dim strDirAbbr As String              ' Abbreviation for language directory
    Dim strDirIn As String = ""           ' Input directory
    Dim strMdict As String = ""           ' Name of morphdict
    Dim strDoTags As String = strAnyVerb  ' CHANGES in the @Label specification should be put here
    Dim bChanged As Boolean = False       ' Cny chnages?

    Try
      ' Get the language
      If (Not EnglishPeriod(strLngAbbr)) Then Exit Sub
      Select Case strLngAbbr
        Case "OEB"  ' Old English, Bossworth & Toller version
          strDirAbbr = "OE" : strMdict = "MorphDictOE.xml" : strDoTags = ""
        Case "OE"
          strDirAbbr = strLngAbbr : strMdict = "MorphDict" & strLngAbbr & ".xml"
          strMorphRemainFile = GetDocDir() & "\MorphDictRemain" & strLngAbbr & ".xml"
        Case "ME"
          strDirAbbr = strLngAbbr : strMdict = "MorphDict" & strLngAbbr & ".xml"
        Case "eModE"
          strDirAbbr = "ModE" : strMdict = "MorphDictModE.xml"
        Case "LmodE"
          strDirAbbr = "MBE" : strMdict = "MorphDictModE.xml"
        Case Else
          Exit Sub
      End Select
      ' Let user select an input directory
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, "Give directory where morphologically enriched " & strLngAbbr & " files are located", _
        "d:\data files\corpora\English\xml\Adapted\" & strDirAbbr)) Then Exit Sub
      ' We need to read the [tdlMorphDict]
      MorphDictIni(False, strMdict)
      If (tdlMorphDict Is Nothing) Then Status("Could not load " & strLngAbbr & " dictionary") : Exit Sub
      ' If this is OE, then perhaps do BT?
      If (strLngAbbr = "OE") Then
        Select Case MsgBox("Process [Remain] using BT list?", MsgBoxStyle.YesNoCancel)
          Case MsgBoxResult.Cancel
            Exit Sub
          Case MsgBoxResult.No
          Case MsgBoxResult.Yes
            If (Not MorphResolveBT(strLngAbbr, bChanged)) Then Status("Problem") : Exit Sub
            If (bChanged) Then
              ' Saving [VernPos]
              Status("Saving vernpos...")
              tdlMorphDict.WriteXml(strMorphDictFile)
            Else
              ' Ask if user wants to continue
              Select Case MsgBox("No changes made. Stop?", MsgBoxStyle.YesNo)
                Case MsgBoxResult.Yes
                  Exit Sub
              End Select
            End If
        End Select
      End If
      ' Try lemmatization using rewrite rules
      If (Not MorphSpreadVernPos(strDirIn, strDirIn, strLngAbbr, strDoTags)) Then Status("Interrupt pressed or LemmatizeRewrite met an error") : Exit Sub
      ' Any changes?
      If (MorphVernPosChanged()) Then
        ' Ask whether we should save
        Select Case MsgBox("Would you like to save the adapted [VernPos] table?", MsgBoxStyle.YesNo)
          Case MsgBoxResult.Yes
            tdlMorphDict.WriteXml(strMorphDictFile)
        End Select
      End If
      ' Show we are ready
      Status("Blessings!")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ToolsMEprojLemmatizeByVernPos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsMEprojEnrichVernPos_Click
  ' Goal:   Evaluate the forms of each verb lemma in [VernPos] and add the standard derivable forms
  ' History:
  ' 04-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsMEprojEnrichVernPos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMEprojEnrichVernPos.Click
    Dim strLngAbbr As String = "" ' Language abbreviation
    Dim strMdict As String = ""   ' Morph dict we need to look for
    Dim bChanged As Boolean = False ' Any changes made?

    Try
      ' Get the language
      If (Not EnglishPeriod(strLngAbbr)) Then Exit Sub
      ' Convert the HTML dictionary into an xml one
      Select Case strLngAbbr
        Case "OE"
          strMdict = "MorphDictOE.xml"
        Case "ME", "MED"
          strMdict = "MorphDictME.xml"
        Case "eModE", "LmodE"
          strMdict = "MorphDictModE.xml"
        Case Else
          Exit Sub
      End Select

      ' We need to read the correct [tdlMorphDict]
      MorphDictIni(False, strMdict)
      If (tdlMorphDict Is Nothing) Then Status("Could not load " & strLngAbbr & " dictionary") : Exit Sub
      ' Convert the HTML dictionary into an xml one
      If (Not MorphDictEnrichVernPos(strLngAbbr, bChanged)) Then Status("Could not enrich VernPos") : Exit Sub
      ' Any changes?
      If (bChanged) Then
        ' Ask whether we should save
        Select Case MsgBox("Would you like to save the adapted [VernPos] table?", MsgBoxStyle.YesNo)
          Case MsgBoxResult.Yes
            tdlMorphDict.WriteXml(strMorphDictFile)
        End Select
      End If
      ' Show we are ready
      Status("Blessings!")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ToolsMEprojEnrichVernPos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsMEprojMorphCategories_Click
  ' Goal:   Add the [Cat] table to the morph dict
  ' History:
  ' 08-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsMEprojMorphCategories_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMEprojMorphCategories.Click
    Dim strLngAbbr As String = ""         ' Language abbreviation
    Dim strDirAbbr As String              ' Abbreviation for language directory
    Dim strCatRepFile As String = ""    ' YCOE tagging report file
    Dim strDirIn As String = ""         ' Input directory with YCOE files
    Dim strMdict As String = ""         ' Name of the morphdict file to be downloaded

    Try
      ' Get the language
      If (Not EnglishPeriod(strLngAbbr)) Then Exit Sub
      Select Case strLngAbbr
        Case "OE"
          strDirAbbr = strLngAbbr : strMdict = "MorphDict" & strLngAbbr & ".xml"
        Case "ME"
          strDirAbbr = strLngAbbr : strMdict = "MorphDict" & strLngAbbr & ".xml"
        Case "eModE"
          strDirAbbr = "ModE" : strMdict = "MorphDictModE.xml"
        Case "LmodE"
          strDirAbbr = "MBE" : strMdict = "MorphDictModE.xml"
        Case Else
          Exit Sub
      End Select
      ' We need access to [tdlMorphDict]
      MorphDictIni(False, strMdict)
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      ' Let user select an input directory
      If (Not GetDirName(Me.FolderBrowserDialog1, strDirIn, "Give directory where morphologically enriched " & strLngAbbr & " files are located", _
        "d:\data files\corpora\English\xml\Adapted\" & strDirAbbr)) Then Exit Sub
      If (MorphCategories(strDirIn, strCatRepFile)) Then
        ' Put the report in place, by navigating to the file
        Status("Opening the categories report...")
        Me.wbReport.Navigate(strCatRepFile)
        ' Show we are ready
        Status("Ready collecting categories information (see report)")
        Me.TabControl1.SelectedTab = Me.tpReport
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsMEprojMorphCategories error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsMEprojDeriveRewrite_Click
  ' Goal:   Use the [Cat] table and other tables within the MorphDict to get suffix rewrite rules
  ' History:
  ' 08-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsMEprojDeriveRewrite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMEprojDeriveRewrite.Click
    Dim strLngAbbr As String = ""         ' Language abbreviation
    Dim strDirAbbr As String              ' Abbreviation for language directory
    Dim strMdict As String = ""         ' Name of the morphdict file to be downloaded

    Try
      ' Reset interrupt
      bInterrupt = False
      ' Get the language
      If (Not EnglishPeriod(strLngAbbr)) Then Exit Sub
      Select Case strLngAbbr
        Case "OE"
          strDirAbbr = strLngAbbr : strMdict = "MorphDict" & strLngAbbr & ".xml"
        Case "ME"
          strDirAbbr = strLngAbbr : strMdict = "MorphDict" & strLngAbbr & ".xml"
        Case "eModE"
          strDirAbbr = "ModE" : strMdict = "MorphDictModE.xml"
        Case "LmodE"
          strDirAbbr = "MBE" : strMdict = "MorphDictModE.xml"
        Case Else
          Exit Sub
      End Select
      ' We need access to [tdlMorphDict]
      MorphDictIni(False, strMdict)
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      ' Try sweeping
      If (Not MorphDeriveRules()) Then Status("Interrupt pressed or Derive-Rewrite met an error")
      ' Try sweeping
      If (Not MorphSuffixRules()) Then Status("Interrupt pressed or Suffix-Rewrite met an error")
      ' Ask whether we should save
      Select Case MsgBox("Would you like to save the [VernPos] table using the derive rules?", MsgBoxStyle.YesNo)
        Case MsgBoxResult.Yes
          ' Adapt [VernPos] with the suffix rules
          ' TODO: implement...

          ' Save the result
          tdlMorphDict.WriteXml(strMorphDictFile)
      End Select
      ' Show we are ready
      Status("Blessings!")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsOEmorphDeriveRewrite error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMEprojLemmatizeRewrite_Click
  ' Goal:   Find the best matching derived form or rewritten lemma
  ' History:
  ' 08-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsMEprojLemmatizeRewrite_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMEprojLemmatizeByDict.Click
    Dim strLngAbbr As String = "" ' Language abbreviation
    Dim strMdict As String = ""   ' Morph dict we need to look for

    Try
      ' Get the language
      If (Not EnglishPeriod(strLngAbbr)) Then Exit Sub
      ' Try lemmatization by best match
      LemmatizeBestMatch(strLngAbbr)
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsOEmorphLemmatizeByDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   LemmatizeBestMatch
  ' Goal:   Find the best matching derived form or rewritten lemma
  '         This is for Middle-English, in principle!
  ' History:
  ' 08-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub LemmatizeBestMatch(ByVal strPeriod As String)
    Dim intNum As Integer = 0   ' Number of changes

    Try
      strMorphRemainFile = GetDocDir() & "\MorphDictRemain" & strPeriod & ".xml"
      Select Case strPeriod
        Case "OE"
          ' Initialise morphological dictionary
          MorphDictIni()
        Case "ME"
          ' Initialise morphological dictionary
          MorphDictIni(False, "MorphDictME.xml")
        Case "eModE", "lModE", "LmodE"
          ' Initialise morphological dictionary
          MorphDictIni(False, "MorphDictModE.xml")
        Case Else
          Exit Sub
      End Select
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      ' Try lemmatization using rewrite rules
      If (Not MorphLemmatizeBestMatch(intNum)) Then Status("Interrupt pressed or LemmatizeRewrite met an error")
      If (intNum > 0) Then
        ' Ask whether we should save
        Select Case MsgBox("Would you like to save the adapted [VernPos] table with [" & intNum & "] changes?", MsgBoxStyle.YesNo)
          Case MsgBoxResult.Yes
            tdlMorphDict.WriteXml(strMorphDictFile)
        End Select
      End If
      ' Show we are ready
      Status("Blessings!")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/LemmatizeBestMatch error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMEprojLemmatizeRewrite_Click
  ' Goal:   Using the suffix-rewrite rules in table [Suffix],
  '           attempt to determine the lemma for words that have not yet been lemmatized
  '         Action depends on the number of different lemma's that are possible
  '         zero - leave as is
  '         one  - accept this as lemma
  '         more - add a list of lemma's, possibly extended with confidence?
  ' History:
  ' 08-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub LemmatizeRewrite(ByVal strPeriod As String)
    Dim strLangDict As String = ""  ' Name of language dictionary

    Try
      Select Case strPeriod
        Case "OE"
          ' Initialise morphological dictionary
          MorphDictIni()
          strLangDict = "OEdict_out.xml"
        Case "ME"
          ' Initialise morphological dictionary
          MorphDictIni(False, "MorphDictME.xml")
          strLangDict = "MEdict_out.xml"
        Case Else
          Exit Sub
      End Select
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      ' Try read the OE dictionary
      If (Not MorphReadOEdict(strLangDict)) Then Status("Could not read the OE dictionary database") : Exit Sub
      ' Try to extract prefixes from OEdict into MorphDict table [Prefix]
      If (Not MorphPfxOEdict()) Then Status("Could not extract prefixes from OE dictionary database") : Exit Sub
      ' Try lemmatization using rewrite rules
      If (Not MorphLemmatizeByDict()) Then Status("Interrupt pressed or LemmatizeRewrite met an error")
      ' Ask whether we should save
      Select Case MsgBox("Would you like to save the adapted [VernPos] table?", MsgBoxStyle.YesNo)
        Case MsgBoxResult.Yes
          tdlMorphDict.WriteXml(strMorphDictFile)
      End Select
      ' Show we are ready
      Status("Blessings!")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/LemmatizeRewrite error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMEprojLemmatizeByTable_Click
  ' Goal:   Process the resolved entries from table [Remain] into table [VernPos]
  '         The resolved entries are in a tab-separated text file
  ' History:
  ' 08-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsMEprojLemmatizeByTable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMEprojLemmatizeByTable.Click
    Dim strFile As String = ""          ' File with the resolve information we are looking for
    Dim strLngAbbr As String = ""       ' Language abbreviation
    Dim strDirAbbr As String            ' Abbreviation for language directory (the last part of it)
    Dim strTagRepFile As String = ""    ' ME tagging report file
    Dim strDirIn As String = ""         ' Input directory with YCOE files
    Dim strMdict As String = ""         ' Morphological dictionary name
    Dim strSrc As String = "ERK"        ' Source of information

    Try
      ' Get the language
      If (Not EnglishPeriod(strLngAbbr)) Then Exit Sub
      Select Case strLngAbbr
        Case "OE"
          strDirAbbr = strLngAbbr
          strMdict = "MorphDictOE.xml"
        Case "ME"
          strDirAbbr = strLngAbbr
          strMdict = "MorphDictME.xml"
          ' Ask for the source of this information: AvK or EK
          strSrc = InputBox("Person abbreviation", "Who is the information source?", strSrc)
        Case "eModE"
          strDirAbbr = "ModE"
          strMdict = "MorphDictModE.xml"
        Case "LmodE"
          strDirAbbr = "MBE"
          strMdict = "MorphDictModE.xml"
        Case Else
          Status("Language not recognized: " & strLngAbbr)
          Exit Sub
      End Select
      ' Initialise morphological dictionary
      MorphDictIni(False, strMdict)
      If (tdlMorphDict Is Nothing) Then Status("Could not open the morphological dictionary dataset") : Exit Sub
      ' Set lemma source
      SetLemmaSource(strSrc)
      ' Ask for the resolution tab-separated file
      If (Not GetFileName(Me.OpenFileDialog1, GetDocDir(), strFile, "tab-separated text file (*.csv)|*.csv")) Then Exit Sub
      ' Process this file
      If (Not MorphResolveTable(strFile, strLngAbbr)) Then Status("Could not process resolution information") : Exit Sub
      ' Ask whether we should save
      Select Case MsgBox("Would you like to save the adapted [VernPos] table?", MsgBoxStyle.YesNo)
        Case MsgBoxResult.Yes
          tdlMorphDict.WriteXml(strMorphDictFile)
      End Select
      ' Show we are ready
      Status("Blessings!")
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/mnuToolsOEmorphLemmatizeByTable error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuToolsMorphDictMake_Click
  ' Goal:   Convert the .txt file with the OE dictionary into a .sfm file with the correct standard
  '           format markers
  ' History:
  ' 14-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub mnuToolsMorphDictMake_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuToolsMorphDictMake.Click
    Try
      'If (Not MorphMakeDict()) Then
      '  Status("Could not make dictionary")
      'End If
      ' Now convert the SFM dictionary into an xml one
      If (Not MorphDictSfmToXml()) Then
        Status("Could not convert SFM to XML")
      Else
        Status("Ready")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/ToolsMorphDictMake error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  Private Sub mnuCorpusZoomUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusZoomUp.Click
    DoZoom(5)
  End Sub
  Private Sub mnuCorpusZoomDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCorpusZoomDown.Click
    DoZoom(-5)
    'Dim wbNew As New ZoomBrowser.MyBrowser
    'Try
    '  Me.wbResText.Document.Body.Style = "zoom: 200%"
    '  'Me.wbResText.Focus()
    '  'SendKeys.Send("^{+}")
    '  ' Attach it
    '  'wbNew.Zoom(200)
    '  'MsgBox("ok")
    '  ' DoZoom(200)
    'Catch ex As Exception

    'End Try
  End Sub
  Private Sub DoZoom(ByVal intIncr As Integer)
    Try
      ' Calculate the new zoomfactor
      intZoom += intIncr
      If (intZoom < 0) Then intZoom = 0
      If (intZoom > 400) Then intZoom = 400
      ' Implement this zoom factor
      Me.wbResText.Document.Body.Style = "zoom: " & intZoom & "%"
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/DoZoom error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

End Class
