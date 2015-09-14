Imports System.Xml
Module modMain
  ' ================================= GLOBAL VARIABLES =========================================================
  Public objCed As DgvHandle                ' Handler for Coref Type Editor dgv
  Public objSynEd As DgvHandle              ' Syntax handler
  Public objPed As DgvHandle                ' Handler for Phrase Type Editor dgv
  Public objPROed As DgvHandle              ' Handler for Pronoun editor dgv
  Public objCatEd As DgvHandle              ' Handler for Categories
  Public objNPed As DgvHandle               ' Handler for NP feature editor dgv
  Public objCnsEd As DgvHandle              ' Handler for constraint editor dgv
  Public objStaEd As DgvHandle              ' Handler for Status types (mainly for database function)
  Public objTmpEd As DgvHandle              ' Handler for template types
  Public objMosEd As DgvHandle              ' Handler for MOS types
  Public objMosAcEd As DgvHandle            ' Handler for MOS action types
  Public objTCed As DgvHandle               ' Handler for template-cell types
  Public objResEd As DgvHandle              ' Handler for results database dgv (On [frmMain])
  Public objDepEd As DgvHandle              ' Handler for dependency editor
  Public objRefItemEd As DgvHandle          ' Handler for [Item] table from [tdlRefChain]
  Public objRefChainEd As DgvHandle         ' Handler for [Chain] table from [tdlRefChain]
  Public objQryEd As DgvHandle              ' Handler for query editor (frmFindX)
  Public objQreRelEd As DgvHandle           ' Handler for query relation editor (frmFindX)\
  Public objXrelEd As DgvHandle             ' Handler for Xpath relations and functions
  Public objOEdictEd As DgvHandle           ' OE dictionary
  Public objFRed As DgvHandle               ' Find and replace handler
  Public strWorkDir As String = ""          ' Working directory where the PSDX files are located
  Public strCrpDir As String = ""           ' Directory where corpus research projects are found
  Public strCrpLast As String = ""          ' Last corpus results database file used
  Public strInputDir As String = ""         ' Directory for input files
  Public strOutputDir As String = ""        ' Directory for output files
  Public strSchemaFile As String = ""       ' XSD schema for PSDX type files
  Public strSchemaPer As String = ""        ' XSD schema for period information
  Public strSchemaCdict As String = ""      ' XSD schema for chain dictionary information
  Public strLastSyntaxLearn As String = ""  ' Last directory we used for learning syntax
  Public strLastConvert As String = ""      ' Last directory used for conversion
  Public strLastDepDir As String = ""       ' Last directory we used for adding dependency
  Public strPeriodFile As String = ""       ' Period information file name
  Public strChainDict As String = ""        ' Chain dictionary file name
  Public strResultDb As String = ""         ' Results database file name (current one)
  Public strCurrentDbFile As String = ""    ' Name of the currently loaded dgv-part of the database file (in-between)
  Public strDriveChange As String = ""      ' Automatically change files of drive X>Y
  Public strUserName As String = My.User.Name ' Who I am
  Public strCurrentFile As String = ""      ' The name of the file we are currently working on
  Public strCurrentTrans As String = ""     ' The name of the translation file we are currently working on
  Public strCurrentText As String = ""      ' The text of the file we are currently working (editing)
  Public strCurrentPeriod As String = ""    ' The period of the file we are now treating
  Public bShowCode As Boolean               ' Setting of chbShowCode
  Public bUserProPGN As Boolean             ' Setting of chbUserProPGN
  Public bShowAllAnt As Boolean             ' Whether the show all antecedents or a subset
  Public bShowCmd As Boolean                ' Setting of chbShowCmd
  Public bCheckCat As Boolean = False       ' Check new categories with new Cesax versions
  Public bCheckCns As Boolean = False       ' Check new constraints with new Cesax versions
  Public tdlPeriods As DataSet = Nothing    ' The XML document with the period definitions
  Public tdlChainDict As DataSet = Nothing  ' The XML document with the chain dictionary definitions
  Public tdlResults As DataSet = Nothing    ' Corpus Research Results database
  Public tdlDependency As DataSet = Nothing ' Dependency view database
  Public tdlRefChain As DataSet = Nothing   ' Coreferential Chain database
  Public tdlFindX As DataSet = Nothing      ' Database with results from a query
  Public tdlChart As DataSet = Nothing      ' Dataset used in charting
  Public tdlMorphTag As DataSet = Nothing   ' Dictionary with OE morphological codes
  Public tdlMorphDict As DataSet = Nothing  ' Dictionary with OE morphological codes
  Public pdxCurrentFile As Xml.XmlDocument  ' The XML document we are currently working on...
  Public pdxList As Xml.XmlNodeList         ' List of all the <forest> nodes
  Public pdxSection As Xml.XmlNodeList      ' List of all the <forest> nodes starting a section
  Public pdxCoref As Xml.XmlNodeList        ' List of all the <eTree> nodes with coreference information
  Public pdxResults As Xml.XmlDataDocument  ' XMLdocument version of [tdlResults]
  Public pdxCrpDbase As Xml.XmlDocument     ' The XML corpus database we are currently working on...
  Public ndxCurrentRes As XmlNode = Nothing ' Currently selected result
  Public intSectFirst As Integer            ' First index of this section in [pdxList]
  Public intSectLast As Integer             ' Last index of this section in [pdxList]
  Public intSectSize As Integer             ' Number of <forest> elements in this section
  Public intCurrentSection As Integer       ' The section we are now at
  Public intMaxIPdist As Integer            ' Maximum IP distance
  Public intAbsMaxIPdist As Integer         ' Cut-off maximum IP distance
  Public intProfileLevel As Integer         ' Which ones to take into account in profiling
  Public intLastTransLine As Integer = -1   ' Line of last translation (if any)
  Public intChainNum As Integer = 0         ' The number of this chain
  Public intLogBase As Integer = 2          ' The log-base used for Reference/List calculations
  Public intMaxEtreeId As Integer = 0       ' The maximum EtreeId we have
  Public tblPhrType As DataTable = Nothing  ' Table containing the Phrase Types (NP/D, PRO$/* etc)
  Public tblRefType As DataTable = Nothing  ' Table containing the Reference Types (Identity, Inferred)
  Public tblFeatDict As DataTable = Nothing ' Feature Dictionary table
  Public dtrPhrType() As DataRow            ' Correctly ordered set of <PhrType> stuff
  Public objFind As New frmFind             ' Access to the find form
  Public objFindX As New frmFindX           ' Access to the find form for Xpath purposes
  Public objFeat As New frmFeature          ' Access to the feature form
  Public arLeft() As Integer                ' Array of line positions
  Public arLeftId() As Integer              ' Array containing the <eTree Id=?> of the first <eTree> on this line
  Public conTb As New CustomContext()       ' Create an instance of custom XsltContext object.
  Public bInterrupt As Boolean = False      ' Whether we need to stop automatic operations
  Public bDebugging As Boolean = True       ' Whether we are in debugging mode or not
  Public bAutoBusy As Boolean = False       ' Whether the automatic machine is busy
  Public bColoring As Boolean = False       ' Whether we are coloring a node
  Public bNeedSaving As Boolean             ' Whether one particular file needs to be saved (psdx)
  Public strSbjSwitchDomain As String = "IP-MAT*"       ' Domain for subject switch
  Public strCrpResults As String = "CrpResult.xsd"      ' Scheme for the corpus results dataset
  Public strMorphTagScheme As String = "MorphTag.xsd"   ' Scheme for morphological tagging (OE)
  Public strMorphTagFile As String = "OEtagpos.xml.txt"
  Public strMorphDictScheme As String = "MorphDict.xsd" ' Scheme for the morphological dictionary (OE)
  Public strMorphDictFile As String = ""                ' Full location of the morphological dictionary file
  Public strMorphDictName As String = "MorphDictOE.xml" ' Name of morphological directory
  'Public bNoDgvSelChange As Boolean = False            ' Whether DGV selection change should NOT be triggered
  Public arPeriod() As String = {"MBE", "eModE", "ME", "OE"}
  Public strOEtaggedDir As String = "D:\Data files\Corpora\OE-tagged\corpus"
  Public strDefPeriodText As String = _
    "<?xml version='1.0' standalone='yes'?>" & vbCrLf & _
    "<PeriodList>" & vbCrLf & _
    "</PeriodList>" & vbCrLf
  ' =================================GLOBAL CONSTANTS===========================================================
  Public Const TREEBANK_EXTENSIONS As String = "http://www.ru.nl/letteren/xpath/tb"
  Public Const CAT_NP As String = "NP"
  Public Const CAT_NPTYPE As String = "NPtype"
  Public Const CAT_GR As String = "GrRole"
  Public Const CAT_WH As String = "NPwh"
  Public Const CAT_VB As String = "Verb"
  Public Const CAT_ADV As String = "Adverb"
  Public Const CAT_ANIM As String = "Animacy"
  Public Const CAT_COGN As String = "CognitiveState"
  Public Const CAT_PNORM As String = "PrepNorm"         ' Normalized prepositions
  Public Const NUM_FEATURES_ALLOWED As Integer = 24
  Public Const OE_TAGGED_DONE As String = "OE-tagged information added"
  Public Const STATUS_CHANGED As String = "Changed"
  Public Const STATUS_DONE As String = "Done"
  ' ================================== LOCAL VARIABLES =========================================================
  Private bInit As Boolean = False
  Private bInitXresult As Boolean = False   ' Initialisation of FindX result processing
  Private bSection As Boolean = False       ' Indicates that the section showing is busy
  Private loc_intPlus As Integer = 0        ' Number of plusses in this <seg> part so far
  Private loc_arStatusDefault() As String = {"(none)", "No status has been assigned to this category (yet)", _
                                             STATUS_DONE, "This item has been processed completely", _
                                             STATUS_CHANGED, "This item has changed and not yet been looked into afterwards"}
  Private loc_strPeriodOrder As String = "From ASC, Until ASC, PeriodId ASC"
  ' =================================LOCAL CONSTANTS============================================================
  Private SPACES As String = " " & vbTab & vbCrLf
  Private Const LABEL_DIV As String = "-+^"
  Private Const LABEL_DIV_OE As String = "-^"
  Private strLabelDivOE() As String = {"-NOM", "-GEN", "-INS", "-DAT", "-ACC", "-MAT", "-SUB"}
  ' ============================================================================================================
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  CheckSettings
  ' Goal :  Check the settings and figure out whether adjustments are needed
  ' History:
  ' 29-07-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function CheckSettings() As Boolean
    Dim strTmp As String              ' Temporary string
    Dim pdxFile As XmlDocument        ' CorpusStudio settings
    Dim ndxThis As XmlNode            ' Working node
    Dim dtrFound() As DataRow         ' Result of select
    Dim intI As Integer               ' Counter
    Dim dtrShadow() As DataRow        ' Selection within shadow set
    Dim bCopyAll As Boolean = False   ' Whether to copy all "new" categories...
    Dim bCopyThis As Boolean = False  ' Need to copy this one row or not?

    Try
      ' Check the working directory
      If (Not IO.Directory.Exists(strWorkDir)) Then
        ' Tell the user
        MsgBox("You will first need to set a valid working directory in Tools/Settings/General")
        Return False
      End If
      ' Check the period definition file
      If (Not IO.File.Exists(strPeriodFile)) Then
        ' Tell the user
        MsgBox("You will first need to set a valid period definition file in Tools/Settings/General." & _
               "This XML file is supplied with the CorpusStudio program. It can also be downloaded separately.")
        Return False
      End If
      ' Check the AbsMaxDist
      If (intAbsMaxIPdist < 10) Then
        ' Warn user
        MsgBox("You need to set the [Absolute maximum IP distance] to a value larger than 10." & vbCrLf & _
               "Go to Tools/Settings/General, and adjust this value. Recommendation: 250.")
        Return False
      End If
      ' Check if we know the corpusstudio working directory
      If (Not IO.Directory.Exists(strCrpDir)) Then
        ' Try to find and open the CorpusStudio settings file
        strTmp = GetSetting("CorpusStudio", "Settings", "Location")
        If (strTmp <> "") Then
          ' We must be able to derive the working directory
          pdxFile = Nothing
          If (ReadXmlDoc(strTmp, pdxFile)) Then
            ' Find the correct point
            ndxThis = pdxFile.SelectSingleNode("//Setting[@Name='WorkDir']")
            ' Found anything?
            If (Not ndxThis Is Nothing) Then
              ' Get the value
              strCrpDir = ndxThis.Attributes("Value").Value
              ' Save this setting
              SetTableSetting(tdlSettings, "CrpDir", strCrpDir)
            End If
          End If
        End If
      End If
      ' Have a look at the shadow settings, and make suggestions appropriately
      ' (1) Check the Category list
      If (bCheckCat) Then
        dtrShadow = tdlShadow.Tables("Category").Select("", "CatId ASC")
        For intI = 0 To dtrShadow.Length - 1
          ' Try find this row in my own settings
          dtrFound = tdlSettings.Tables("Category").Select("Name='" & dtrShadow(intI).Item("Name") & "'")
          ' Found anything?
          If (dtrFound.Length = 0) Then
            ' Assume no copying
            bCopyThis = False
            ' Check if we can copy all
            If (bCopyAll) Then
              ' Indicate we need to copy
              bCopyThis = True
            Else
              ' Suggest copying this row
              Select Case MsgBox("I found a category (see Tools/Settings/Categories) in the shadow settings, " & vbCrLf & _
                                 "which is not yet in your settings: [" & dtrShadow(intI).Item("Name") & "]" & vbCrLf & _
                                 "Shall I copy all new categories (Y), " & vbCrLf & _
                                 "- Copy this new category (N)" & vbCrLf & _
                                 "- Or would you like to skip this? (Cancel)?", MsgBoxStyle.YesNoCancel)
                Case MsgBoxResult.Yes
                  ' Set copying all
                  bCopyAll = True : bCopyThis = True
                Case MsgBoxResult.No
                  bCopyThis = True
                Case MsgBoxResult.Cancel
                  ' Leave the for-loop
                  Exit For
              End Select
            End If
            ' Do we need to copy this row?
            If (bCopyThis) Then
              ' Copy this row
              If (Not CopyTableRow(dtrShadow(intI), tdlSettings.Tables("Category"), _
                                   tdlSettings.Tables("CategoryList").Rows(0))) Then
                ' Show user that copying went wrong
                Select Case MsgBox("There was an error copying this row of categories: " & _
                                   dtrShadow(intI).Item("Name"), MsgBoxStyle.OkCancel)
                  Case MsgBoxResult.Ok
                    ' But we keep on trying!!
                  Case MsgBoxResult.Cancel
                    ' Stop
                    Exit For
                End Select
              End If
            End If
          End If
        Next intI
      End If
      ' (2) Check the constraint ranking
      If (bCheckCns) Then
        dtrShadow = tdlShadow.Tables("Constraint").Select("", "ConstraintId ASC")
        bCopyThis = False
        For intI = 0 To dtrShadow.Length - 1
          ' Try find this row in my own settings
          dtrFound = tdlSettings.Tables("Constraint").Select("Name='" & dtrShadow(intI).Item("Name") & "'")
          ' Found anything?
          If (dtrFound.Length = 0) Then
            ' This is the first place the constraints differ in ranking...
            Select Case MsgBox("Your constraint ranking differs from the one suggested in the shadow settings." & _
                               vbCrLf & "Would you like change to the new ranking?", MsgBoxStyle.YesNoCancel)
              Case MsgBoxResult.Yes
                bCopyThis = True
            End Select
          End If
        Next intI
        ' Do we need to copy the ranking?
        If (bCopyThis) Then
          ' We need to copy the shadow ranking to our own
          ' (1) Delete all rows in our own ranking
          With tdlSettings.Tables("Constraint")
            For intI = .Rows.Count - 1 To 0
              ' Delete this row
              .Rows(intI).Delete()
            Next intI
            ' Make sure this ripples through
            .AcceptChanges()
          End With
          ' (2) Copy all rows from shadow to me
          With tdlShadow.Tables("Constraint")
            For intI = 0 To .Rows.Count - 1
              ' Copy this row
              If (Not CopyTableRow(.Rows(intI), tdlSettings.Tables("Constraint"), _
                                   tdlSettings.Tables("ConstraintList").Rows(0))) Then
                ' Something went wrong
                MsgBox("There was a problem copying the constraints from the shadow settings." & vbCrLf & _
                       "Please contact your system administrator and quit the program.")
                Return False
              End If
            Next intI
          End With
        End If
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/CheckSettings error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  ReadSettings
  ' Goal :  Given initialisation and an appropriate settings file, read settings
  ' History:
  ' 29-07-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Sub ReadSettings()
    Dim tblSettings As DataTable    ' The Settings datatable
    Dim dtrThis As DataRow          ' Working row
    Dim strText As String           ' For verification
    Dim strValue As String          ' A value of a variable that MIGHT be used...
    Dim intI As Integer             ' Counter

    Try
      ' Check initialisation and settings file name
      If (strSetFile = "") Then Exit Sub
      'Dim intI As Integer
      'For intI = 0 To tdlSettings.Tables.Count - 1
      '  Debug.Print("Table #" & intI + 1 & "=" & tdlSettings.Tables(intI).TableName)
      'Next intI
      ' Get the settings part
      tblSettings = tdlSettings.Tables("Setting")
      ' Set the appropriate find attributes...
      tblSettings.PrimaryKey = New DataColumn() {tblSettings.Columns("Name")}
      ' Some Settings need a different default if they have no value
      If (GetTableSetting(tdlSettings, "CheckCns").ToString = "") Then SetTableSetting(tdlSettings, "CheckCns", "True")
      If (GetTableSetting(tdlSettings, "CheckCat").ToString = "") Then SetTableSetting(tdlSettings, "CheckCat", "True")
      If (GetTableSetting(tdlSettings, "AbsMaxIPdist").ToString = "") Then SetTableSetting(tdlSettings, "AbsMaxIPdist", "250")
      If (GetTableSetting(tdlSettings, "LogBase").ToString = "") Then SetTableSetting(tdlSettings, "LogBase", "2")
      ' Read the appropriate settings into public variables
      strWorkDir = GetSettingValue("WorkDir") & ""
      strCrpDir = GetSettingValue("CrpDir") & ""
      strPeriodFile = GetSettingValue("PeriodDefinition") & ""
      strChainDict = GetSettingValue("ChainDictionary") & ""
      strCrpLast = GetSettingValue("CrpLast") & ""
      strLastSyntaxLearn = GetSettingValue("LastSyntaxLearn") & ""
      strLastConvert = GetSettingValue("LastConvertDir") & ""
      strLastDepDir = GetSettingValue("LastDepDir") & ""
      ' If there is a last corpus results database, then make it into a menu item
      If (strCrpLast = "") Then
        ' Make sure the menu item is not visible
        frmMain.mnuCorpusRecent1.Visible = False
      Else
        frmMain.mnuCorpusRecent1.Visible = True
        ' Fill in the text of this item
        frmMain.mnuCorpusRecent1.Text = "&1 " & IO.Path.GetFileName(strCrpLast)
      End If
      ' If no chain dictionary is specified, we have to use a default one
      If (strChainDict = "") Then
        strChainDict = GetDocDir() & "\ChainDictionary.xml"
        ' strChainDict = My.Computer.FileSystem.SpecialDirectories.MyDocuments.ToString & "\ChainDictionary.xml"
        ' Save this setting
        SetTableSetting(tdlSettings, "ChainDictionary", strChainDict)
      End If
      bShowCode = (GetSettingValue("ShowCode") = "True")
      bUserProPGN = (GetSettingValue("UserProPGN") = "True")
      bShowAllAnt = (GetSettingValue("ShowAllAnt") = "True")
      bDebugging = (GetSettingValue("Debugging") = "True")
      bCheckCat = (GetSettingValue("CheckCat") = "True")
      bCheckCns = (GetSettingValue("CheckCns") = "True")
      ' Read the drivechange settings
      strDriveChange = GetSettingValue("DriveChange")
      ' Read the user's name
      strUserName = GetSettingValue("UserName")
      ' Get the maximum IP distance
      strText = GetSettingValue("MaxIPdist")
      If (IsNumeric(strText)) Then
        intMaxIPdist = CInt(strText)
      Else
        intMaxIPdist = -1
      End If
      ' Get the absolute maximum IP distance
      strText = GetSettingValue("AbsMaxIPdist")
      If (IsNumeric(strText)) Then
        intAbsMaxIPdist = CInt(strText)
      Else
        ' We need to take an appropriate value for the absolute maximum IP distance, otherwise
        '  there will be problems
        intAbsMaxIPdist = -1
      End If
      ' Get the profile level
      strText = GetSettingValue("ProfileLevel")
      If (IsNumeric(strText)) Then
        intProfileLevel = CInt(strText)
      Else
        intProfileLevel = 0
      End If
      ' Get the log base
      strText = GetSettingValue("LogBase")
      If (IsNumeric(strText)) Then intLogBase = CInt(strText) Else intLogBase = 2
      ' This is the normal file name. The directory + .txt extension will be derived "automatically"
      '  by the procedure "ReadSettings()"...
      strSchemaFile = "Psdx.xsd"
      strSchemaPer = "PeriodDef.xsd"
      strSchemaCdict = "ChainDict.xsd"
      ' Try to read the periods
      TryReadPeriods()
      ' Try to read the chain dictionary (if it exists)
      TryReadChainDict()
      ' Establish the PhrType and RefType datatables (sections of [tdlSettings])
      tblPhrType = tdlSettings.Tables("PhrType")
      tblRefType = tdlSettings.Tables("RefType")
      ' Make a correctly ordered set of <PhrType>
      dtrPhrType = tdlSettings.Tables("PhrType").Select("", "PhrTypeId ASC")
      ' Get Source include and exclude NP types
      strValue = GetSettingValue("PhraseInclude")
      If (strValue = "") Then
        SetTableSetting(tdlSettings, "PhraseInclude", strNPsourceTypes)
        ' Save the settings
        Call XmlSaveSettings(strSetFile)
      Else
        strNPsourceTypes = strValue
      End If
      strValue = GetSettingValue("PhraseExclude")
      If (strValue = "") Then
        SetTableSetting(tdlSettings, "PhraseExclude", strNPnoSourceTypes)
        ' Save the settings
        Call XmlSaveSettings(strSetFile)
      Else
        strNPnoSourceTypes = strValue
      End If
      ' Check if morphological Tag dictionary is read
      If (tdlMorphTag Is Nothing) Then
        ' Read morphological dictionary
        If (Not ReadDataset(strMorphTagScheme, strMorphTagFile, tdlMorphTag)) Then
          ' Tell the user
          Logging("Could not read morphological tag dictionary: " & strMorphTagFile)
          Status("Could not read morphological tag dictionary: " & strMorphTagFile)
          ' But return OK anyway!
        End If
      End If
      ' Check for the presence of a number of "Status" values that need to be there anyway
      For intI = 0 To loc_arStatusDefault.Length - 1 Step 2
        If (GetTableID(tdlSettings, "Status", "Name", loc_arStatusDefault(intI), "StatusId", True) < 0) Then
          ' Add this status name + description
          dtrThis = AddOneDataRow(tdlSettings, "Status", "StatusId", "StatusList")
          dtrThis.Item("Name") = loc_arStatusDefault(intI) : dtrThis.Item("Descr") = loc_arStatusDefault(intI + 1)
          ' Make sure settings are saved
          XmlSaveSettings(strSetFile)
        End If
      Next intI
      ' All is well!
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/ReadSettings error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphDictIni
  ' Goal:   Initialize the morphological dictionary
  ' History:
  ' 18-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub MorphDictIni(Optional ByVal bForce As Boolean = False, _
                          Optional ByVal strDictName As String = "")
    Dim bDoRenumber As Boolean = False ' Debugging: renumbering
    Dim dtrFound() As DataRow         ' Result of selecting
    Dim intI As Integer               ' Counter
    Dim intJ As Integer               ' Counter

    Try
      ' Check if morphological information dictionary for OE is read
      If (tdlMorphDict Is Nothing) OrElse (bForce) Then
        ' Make a name for the dictionary
        If (strDictName = "") Then
          strMorphDictFile = GetDocDir() & "\" & strMorphDictName
        Else
          strMorphDictFile = GetDocDir() & "\" & strDictName
        End If
        ' Check location of file
        If (strMorphDictFile = "") OrElse (Not IO.File.Exists(strMorphDictFile)) Then
          ' Create morphological dictionary
          If (Not CreateDataSet(strMorphDictScheme, tdlMorphDict)) Then
            ' Tell the user
            Logging("Could not create OE morphological information dictionary: " & strMorphDictFile)
            Status("Could not create OE morphological information dictionary: " & strMorphDictFile)
            ' But return OK anyway!
            Exit Sub
          End If
        Else
          ' Read morphological dictionary
          Status("Loading morphology dictionary information...")
          If (Not ReadDataset(strMorphDictScheme, strMorphDictFile, tdlMorphDict)) Then
            ' Tell the user
            Logging("Could not read OE morphological information dictionary: " & strMorphDictFile)
            Status("Could not read OE morphological information dictionary: " & strMorphDictFile)
            ' But return OK anyway!
            Exit Sub
          End If
          ' =========== DEBUGGING ===========
          'If (bDoRenumber) Then
          '  ' Find renumber section
          '  dtrFound = tdlMorphDict.Tables("VernPos").Select("Src LIKE 'LinkToMorph*'", "VernPosId ASC")
          '  For intI = 0 To dtrFound.Length - 1
          '    ' Check
          '    dtrFound(intI).Item("VernPosId") = intI + 1
          '  Next intI
          '  ' Note where we are
          '  intJ = intI
          '  ' Second half
          '  dtrFound = tdlMorphDict.Tables("VernPos").Select("Src NOT LIKE 'LinkToMorph*'", "VernPosId ASC")
          '  For intI = 0 To dtrFound.Length - 1
          '    ' Check
          '    'If (intI >= 7405) Then Stop
          '    dtrFound(intI).Item("VernPosId") = intJ + intI + 1
          '  Next intI
          '  ' Save the result
          '  tdlMorphDict.WriteXml(strMorphDictFile)
          'End If
          ' =================================
          ' Do we need to clear the <Morph> part?
          If (bForce) Then
            ClearTable(tdlMorphDict.Tables("Morph"))
            tdlMorphDict.AcceptChanges()
          End If
        End If
      End If
      If (InStr(strMorphDictFile, "OE.xml") > 0) Then
        ' Try read the OE-tagged lexicon
        If (Not MorphReadTaggedLexicon()) Then
          Logging("Could not process OE lemma's dictionary")
          Status("Could not process OE lemma's dictionary")
          Exit Sub
        End If
      End If
      Status("ready")
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/MorphDictIni error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   FeatDictInit
  ' Goal:   Initialise the table that will record the actions
  ' History:
  ' 28-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function FeatDictInit(ByVal bReset As Boolean) As Boolean
    Try
      ' Was there a previous table?
      If (tblFeatDict Is Nothing) Then
        ' Make the table
        tblFeatDict = New DataTable("FeatDict")
        ' Set columns for this table
        With tblFeatDict
          ' The vernacular entry of the dictionary
          .Columns.Add("Vern", System.Type.GetType("System.String"))
          ' The vernacular entry of the dictionary, but without $, -, *
          .Columns.Add("Bare", System.Type.GetType("System.String"))
          ' The bare entry, but with vowels deleted (for seeking)
          .Columns.Add("Seek", System.Type.GetType("System.String"))
          ' The period (in a string)
          .Columns.Add("Period", System.Type.GetType("System.String"))
          ' The category (Verb, Adverb, Noun)
          .Columns.Add("Cat", System.Type.GetType("System.String"))
          ' The type
          .Columns.Add("Type", System.Type.GetType("System.String"))
        End With
      ElseIf (bReset) Then
        ' There is a previous one already - just clear it
        ClearTable(tblFeatDict)
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/FeatDictInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MakeFeatDict
  ' Goal:   Make a dictionary of the features for NP, Adv, P or Verb
  '         1 - NP
  '         2 - Adv
  '         3 - V
  '         4 - P
  ' History:
  ' 15-02-2011  ERK Created
  ' 17-01-2013  ERK Added optional [strFile] for Vb and P
  ' ------------------------------------------------------------------------------------
  Public Function MakeFeatDict(ByVal strCategory As String, Optional ByVal strFile As String = "") As Boolean
    Dim strType As String = ""    ' Type of this category (part of the name)
    Dim strVern As String = ""    ' The vernacular lexeme
    Dim strBare As String = ""    ' The bare vernacular item
    Dim strText As String = ""    ' Text of file
    Dim strDelim As String = ""   ' Delimiter
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intK As Integer           ' Counter
    Dim dtrFound() As DataRow     ' Result of selection statement
    Dim arWord() As String        ' List of words
    Dim arLine() As String        ' Read-in dictionary

    Try
      ' Initialise the feature dictionary
      If (Not FeatDictInit(True)) Then Exit Function
      ' If this is "Adv", then we know what to do
      Select Case strCategory
        Case CAT_NP
        Case CAT_VB       ' Unaccusative verbs -- take from a list
          ' Should be a file
          If (strFile = "") Then Return False
          If (Not IO.File.Exists(strFile)) Then Return False
          ' Read file into array
          strText = IO.File.ReadAllText(strFile)
          strDelim = GetDelim(strText, vbCrLf, vbCr, vbLf)
          arLine = Split(strText, strDelim)
          ' Walk all the liens
          For intI = 0 To arLine.Length - 1
            ' Is there something?
            If (arLine(intI) <> "") Then
              ' Add this feature
              If (Not AddFeatDict(strCategory, "unacc", strCurrentPeriod, arLine(intI))) Then Return False
            End If
          Next intI
        Case CAT_PNORM    ' Normalized prepositions -- take from CSV file
          ' Should be a file
          If (strFile = "") Then Return False
          If (Not IO.File.Exists(strFile)) Then Return False
          ' Read file into array
          strText = IO.File.ReadAllText(strFile)
          strDelim = GetDelim(strText, vbCrLf, vbCr, vbLf)
          arLine = Split(strText, strDelim)
          ' Walk all the liens
          For intI = 0 To arLine.Length - 1
            ' Break this line up into ;
            arWord = Split(arLine(intI), ";")
            ' There should be 2 entries
            If (arWord.Length = 2) Then
              ' Add this feature
              If (Not AddFeatDict(strCategory, arWord(1), strCurrentPeriod, arWord(0))) Then Return False
            End If
          Next intI
        Case CAT_ADV
          dtrFound = tdlSettings.Tables("Category").Select("Name LIKE 'Adv-*'", "Type ASC")
          ' First pass: gather all the results
          For intI = 0 To dtrFound.Length - 1
            ' Access this element
            With dtrFound(intI)
              ' Get the type of this element
              strType = .Item("Type")
              ' Visit all periods
              For intJ = 0 To UBound(arPeriod)
                ' Is there anything for this period?
                If (.Item(arPeriod(intJ)).ToString <> "") Then
                  ' Get all the words stored for this period
                  arWord = Split(.Item(arPeriod(intJ)), "|")
                  ' Store all the results
                  For intK = 0 To UBound(arWord)
                    ' Add to the feature dictionary
                    If (Not AddFeatDict(strCategory, strType, arPeriod(intJ), arWord(intK))) Then Return False
                  Next intK
                End If
              Next intJ
            End With
          Next intI
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/MakeFeatDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetFeatLike
  ' Goal:   Get words like [strWord] with feature category [strCategory], i.e. NP, Adv or Verb
  '         1 - NP
  '         2 - Adv
  '         3 - V
  ' History:
  ' 15-02-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetFeatLike(ByVal strCategory As String, ByVal strWord As String) As String
    Dim strOut As String = ""   ' The resulting features
    Dim strBare As String = ""  ' The bare vernacular item
    Dim strSeek As String = ""  ' The vernacular item to be used for seeking
    Dim intI As Integer         ' Counter
    Dim dtrFound() As DataRow   ' Result of SELECT function

    Try
      ' Validate
      If (tblFeatDict Is Nothing) Then Return ""
      ' Get the bare and seek forms
      If (Not BareAndSeek(strWord, strBare, strSeek)) Then Return False
      ' Try find all that are alike
      dtrFound = tblFeatDict.Select("Seek LIKE '" & strSeek & "'", "Bare ASC")
      ' Combine them into a string
      For intI = 0 To dtrFound.Length - 1
        ' Access this element
        With dtrFound(intI)
          ' Add this item to a string, separated by CRLF
          strOut &= .Item("Vern") & " - " & .Item("Type") & " (" & .Item("Period") & ")" & vbCrLf
        End With
      Next intI
      ' Check the result
      If (strOut = "") Then strOut = "(Found no like terms)" & vbCrLf
      ' Return the result
      Return strOut
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/GetFeatLike error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetFeatEntry
  ' Goal:   Get words like [strWord] with feature category [strCategory], i.e. NP, Adv or Verb
  '         1 - Vb
  '         2 - P
  ' History:
  ' 17-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetFeatEntry(ByVal strCategory As String, ByVal strWord As String) As String
    Dim strOut As String = ""   ' The resulting features
    Dim dtrFound() As DataRow   ' Result of SELECT function

    Try
      ' Validate
      If (tblFeatDict Is Nothing) Then Return ""
      ' Action depends on category
      Select Case strCategory
        Case CAT_VB
          ' Try find this entry
          dtrFound = tblFeatDict.Select("Vern = '" & strWord.Replace("'", "''") & "'")
          ' Any result?
          If (dtrFound.Length = 0) Then
            ' It is not in the wordlist/database, so it is not unaccusative
            strOut = "other"
          Else
            ' Since it is in the list, it must be unaccusative
            strOut = "unacc"
          End If
        Case CAT_PNORM
          ' Try find this entry
          dtrFound = tblFeatDict.Select("Vern = '" & strWord.Replace("'", "''") & "'")
          ' Any result?
          If (dtrFound.Length > 0) Then
            ' Return the type (=normalized preposition) of the first result
            strOut = dtrFound(0).Item("Type")
          End If
      End Select
      ' Return the result
      Return strOut
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/GetFeatEntry error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   BareAndSeek
  ' Goal:   Get the bare and the seek form of [strVern]
  ' History:
  ' 15-02-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function BareAndSeek(ByVal strVern As String, ByRef strBare As String, ByRef strSeek As String) As Boolean
    Dim strVowel() As String = {"+", "a", "e", "i", "y", "u", "o"}
    Dim intI As Integer           ' Counter

    Try
      ' Get the bare one
      strBare = strVern.Replace("$", "")
      strBare = strBare.Replace("*", "")
      strBare = strBare.Replace("-", "")
      ' Get the seek one
      strSeek = LCase(strBare)
      For intI = 0 To UBound(strVowel)
        strSeek = strSeek.Replace(strVowel(intI), "")
      Next intI
      ' Change all "t" into "d"
      strSeek = strSeek.Replace("t", "d")
      ' Change "hr" into just "r"
      strSeek = strSeek.Replace("hr", "r")
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/BareAndSeek error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddFeatDict
  ' Goal:   Add to a dictionary of the features for NP, Adv or Verb
  '         1 - NP
  '         2 - Adv
  '         3 - V
  '         4 - P
  ' History:
  ' 15-02-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AddFeatDict(ByVal strCategory As String, ByVal strType As String, ByVal strPeriod As String, _
                               ByVal strVern As String) As Boolean
    Dim strBare As String = ""    ' The bare vernacular item
    Dim strSeek As String = ""    ' The vernacular item to be used for seeking

    Try
      ' Initialise the feature dictionary
      If (Not FeatDictInit(False)) Then Exit Function
      ' If this is "Adv", then we know what to do
      Select Case strCategory
        Case CAT_NP
        Case CAT_VB
          ' Store this word in the dictionary: Vern, Period, Cat, Type
          tblFeatDict.Rows.Add(strVern, strVern, strVern, strPeriod, strCategory, strType)
        Case CAT_PNORM  ' Normalized prepositions
          ' Store this word in the dictionary: Vern, Period, Cat, Type
          tblFeatDict.Rows.Add(strVern, strVern, strVern, strPeriod, strCategory, strType)
        Case CAT_ADV
          ' Get the bare and seek forms
          If (Not BareAndSeek(strVern, strBare, strSeek)) Then Return False
          ' Store this word in the dictionary: Vern, Period, Cat, Type
          tblFeatDict.Rows.Add(strVern, strBare, strSeek, strPeriod, strCategory, strType)
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/AddFeatDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       ReadTranslation()
  ' Goal:       Try to read the translation into our system
  ' History:
  ' 28-05-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function ReadTranslation(ByVal strFile As String) As Boolean
    Dim pdxTrans As New Xml.XmlDocument ' The XML location to store the document

    Try
      ' Show we are reading
      Status("Reading translation information...")
      ' Read the Schema file and the newly selected PSDY file
      If (ReadXmlDoc(strFile, pdxTrans)) Then
        ' Add this translation to the existing one
        '  (or replace the existing one with this translation!!)
        If (DoTransRead(pdxTrans)) Then
          ' Set the translation on the designated tabpage...
          ShowTranslation(True)
        End If
      End If
      ' Make sure we dispose of the xmldocument
      pdxTrans = Nothing
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/ReadTranslation error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       TryReadPeriods()
  ' Goal:       Try to read the periods dataset
  ' History:
  ' 28-05-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Sub TryReadPeriods()
    Try
      ' We should also try to read the period definitions...
      If (strPeriodFile <> "") Then
        ' Does the file exist?
        If (IO.File.Exists(strPeriodFile)) Then
          ' Show what we are doing
          Status("Reading period definitions...")
          ' Read the periods
          If (Not ReadDataset(strSchemaPer, strPeriodFile, tdlPeriods)) Then
            ' Log the complaint
            Logging("Unable to read valid dataset from Period Definitions file " & strPeriodFile)
          End If
          ' Show we are ready
          Status("Ready")
        Else
          ' Log the complaint
          Logging("Period definition file not found: " & strPeriodFile)
        End If
      Else
        ' Log our complaints
        Logging("No period definition file has been specified")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/TryReadPeriods error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------------------------------------------------
  ' Name:       DoMainInit()
  ' Goal:       Perform any "first time" initialisations
  '             (This was used for the initialisation of tblMatch, which is now discontinued)
  ' History:
  ' 27-04-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function DoMainInit() As Boolean
    Try
      ' Initialise the Treebank Xpath functions
      conTb.AddNamespace("tb", TREEBANK_EXTENSIONS)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/DoMainInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       Logging()
  ' Goal:       Log the message on the screen in the General section
  ' History:
  ' 23-04-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Sub Logging(ByVal strIn As String)
    'frmMain.tbGenLog.AppendText(strIn & vbCrLf)
    ' Access the right textbox
    With frmMain.tbGenLog
      ' Append the text to the LOG textbox in the General page -- add a CRLF
      .AppendText(strIn & vbCrLf)
      ' Zorgen dat we naar het einde scrollen
      .SelectionStart = .TextLength
      .ScrollToCaret()
    End With
  End Sub
  '---------------------------------------------------------------------------------------------------------
  ' Name:       LoadOneFile()
  ' Goal:       Check which Recent menu items should be displayed
  ' History:
  ' 10-03-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function LoadOneFile(ByVal strFile As String) As Boolean
    Dim intI As Integer       ' Counter
    Dim intPrev As Integer    ' Number where recent file with same name is already kept

    Try
      ' Validate
      Select Case XmlFileType(strFile)
        Case "TEI"      ' Proper
        Case "CrpOview" ' Result file or database
          Status("Use Corpus/Load to open a corpus results database") : Return False
        Case Else       ' Unknown file type
          Status("Unknown file type") : Return False
      End Select
      ' Try open this file
      frmMain.LoadPsdx(strFile)
      ' Log projtype/CRPname
      TryAppendLog("psdx", strFile)
      ' Check if this file already exists in the Recent ones
      intPrev = GetRecentNumber(strFile)
      ' Found it?
      If (intPrev > 0) Then
        ' Yes, we found it in the recent number --> Swap recent files
        For intI = intPrev To 2 Step -1
          ' Copy settings from [i-1] to [i]
          SetTableSetting(tdlSettings, "Recent" & intI, GetTableSetting(tdlSettings, "Recent" & intI - 1))
          ' Make sure changes are saved
          XmlSaveSettings(strSetFile)
        Next intI
      Else
        ' No, so shift the recent ones
        For intI = 5 To 2 Step -1
          ' Copy settings from [i-1] to [i]
          SetTableSetting(tdlSettings, "Recent" & intI, GetTableSetting(tdlSettings, "Recent" & intI - 1))
        Next intI
        ' Make sure changes are saved
        XmlSaveSettings(strSetFile)
      End If
      ' Make sure the Recent1 setting gets adapted
      If (strCurrentFile <> GetTableSetting(tdlSettings, "Recent1")) Then
        SetTableSetting(tdlSettings, "Recent1", strCurrentFile)
        ' Make sure changes are saved
        XmlSaveSettings(strSetFile)
      End If
      ' Check which menu items should be displayed
      frmMain.CheckRecent()
      ' Show we loaded this file
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/LoadOneFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       GetRecentNumber()
  ' Goal:       Get the number of the [strFile] in the list of recent files
  ' History:
  ' 15-03-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Function GetRecentNumber(ByVal strFile As String) As Integer
    Dim intI As Integer               ' Counter

    Try
      ' Check all number
      For intI = 1 To 5
        ' Is this the one?
        If (GetTableSetting(tdlSettings, "Recent" & intI) = strFile) Then
          ' It already is there
          Return intI
        End If
      Next intI
      ' Return failure
      Return -1
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/GetRecentNumber error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return -1
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       GetCrpRecNumber()
  ' Goal:       Get the number of the [strFile] in the list of recent files
  ' History:
  ' 15-03-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function GetCrpRecNumber(ByVal strFile As String) As Integer
    Dim intI As Integer               ' Counter

    Try
      ' Check all number
      For intI = 1 To 5
        ' Is this the one?
        If (GetTableSetting(tdlSettings, "CorpusRecent" & intI) = strFile) Then
          ' It already is there
          Return intI
        End If
      Next intI
      ' Return failure
      Return -1
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/GetCrpRecNumber error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return -1
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       ShowPsdx()
  ' Goal:       Show the file...
  ' History:
  ' 23-04-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function ShowPsdx(ByRef rtbThis As RichTextBox, ByRef bChanged As Boolean) As Boolean
    Dim strSeg As String = ""     ' Text of this <forest> element
    Dim intSection As Integer = 0 ' Section counter
    Dim objText As New StringColl ' The text we read

    Try
      ' Try initialize
      If (Not InitCurrentFile()) Then
        ' Switch on RichTextboxLayout again
        rtbThis.Visible = True
        ' Return failure
        Return False
      End If
      ' Insert a section here, if it is not there already
      If (Not HasStartSection()) Then
        ' Try to derive sections by looking at (CODE <heading>) entries
        If (AutoAddSections()) Then
          ' Show what has happened
          Logging("Automatically made sections based on <heading> entries.")
        Else
          If (AddStartSection()) Then
            ' Show what is happening
            Logging("There were no sections, so I added section <0>")
          End If
        End If
        ' Show we have changed
        bChanged = True
      End If
      ' Calculate which sections there are
      SectionCalculate()
      '' Go to the first section
      'ShowSection(rtbThis, 0)
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/ShowPsdx error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Switch on RichTextboxLayout again
      rtbThis.Visible = True
      ' Return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       InitCurrentFile()
  ' Goal:       Necessary initialisations: divide [pdxCurrentFile] into [pdxList] items
  '             And also calculate the values for [arLeftId()]
  ' History:
  ' 10-06-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function InitCurrentFile() As Boolean
    Try
      ' Select all the <forest> elements
      pdxList = pdxCurrentFile.SelectNodes("//forest")
      ' Anything back?
      If (pdxList Is Nothing) Then
        ' Obviously something went wrong
        MsgBox("Could not derive data...")
        Return False
      End If
      If (pdxList.Count = 0) Then
        ' There are no <forest> elements...
        MsgBox("The file you are trying to load does not contain <forest> elements")
        bInterrupt = True
        Return False
      End If
      ' Calculate the <eTree> id's starting each line
      If (Not StartIdCalculate()) Then Return False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/InitCurrentFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       InitEtreeId()
  ' Goal:       Check all the ID's of the <eTree> elements
  ' History:
  ' 10-06-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function InitEtreeId() As Boolean
    Try

    Catch ex As Exception

    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       StartIdCalculate()
  ' Goal:       Get all the Id's of the <eTree> elements starting each <forest> line
  ' History:
  ' 18-05-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Function StartIdCalculate() As Boolean
    Dim intI As Integer       ' Counter
    Dim intJ As Integer       ' Counter
    Dim intId As Integer      ' ID of the current element <eTree>
    Dim intPtc As Integer     ' Percentage
    Dim intPrevId As Integer  ' Last id of previous line
    Dim ndxThis As XmlNode    ' Id of the first <eTree> element
    Dim ndxList As XmlNodeList ' List of all the <eTree> elements in this line

    Try
      ReDim arLeftId(0 To pdxList.Count - 1) : intPrevId = -1 : intMaxEtreeId = 0
      ' Walk the tree
      For intI = 0 To pdxList.Count - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ pdxList.Count
        Status("Checking " & intPtc & "%", intPtc)
        ' initialise the element
        If (intI = 0) Then
          arLeftId(0) = -1
        Else
          ' Initialize it to the previous one
          arLeftId(intI) = arLeftId(intI - 1)
        End If
        ' Get ALL the <eTree> elements in this line
        ndxList = pdxList(intI).SelectNodes(".//eTree")
        ' Try get the first <eTree> child
        'ndxThis = pdxList(intI).Item("eTree")
        ' Got something?
        'If (Not ndxThis Is Nothing) Then
        If (ndxList.Count > 0) Then
          ' Walk all id's, making sure each is larger than the previous one
          For intJ = 0 To ndxList.Count - 1
            ' Get this node
            ndxThis = ndxList(intJ)
            ' Verify ID field
            If (ndxThis.Attributes("Id") Is Nothing) Then
              MsgBox("The PSDX code is not correct: every <eTree> should have an @Id field")
              ' Interrupt and return failure
              Status(" Interrupted!!")
              bInterrupt = True
              Return False
            End If
            ' Get this one's Id
            intId = ndxThis.Attributes("Id").Value
            ' Adapt maximum
            If (intId > intMaxEtreeId) Then intMaxEtreeId = intId
            ' Check
            If (intPrevId = -1) Then
              ' The first id must be 1 - otherwise change!
              If (intId <> 1) Then
                intId = 1
                ndxThis.Attributes("Id").Value = intId
                ' Set dirty
                frmMain.SetDirty(True)
              End If
            ElseIf (intId <> intPrevId + 1) Then
              ' Adapt
              intId = intPrevId + 1
              ndxThis.Attributes("Id").Value = intId
              ' Set dirty
              frmMain.SetDirty(True)
            End If
            ' Store previd
            intPrevId = intId
          Next intJ
          ' Get first one
          ndxThis = ndxList(0)
          ' Get its Id
          intId = ndxThis.Attributes("Id").Value
          ' Store its Id
          arLeftId(intI) = intId
        End If
      Next intI
      ' Return succes
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/StartIdCalculate error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       SectionsLoad()
  ' Goal:       Load combobox with all section locations
  ' History:
  ' 14-05-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Sub SectionsLoad(ByRef lboThis As ListBox)
    Dim intI As Integer   ' Counter
    Dim intNum As Integer ' Number of sections
    Dim strSect As String ' Identifier of this section

    Try
      ' Check validity
      If (lboThis Is Nothing) Then Exit Sub
      ' Determine the number of sections
      intNum = pdxSection.Count
      ' Initialise the combobox
      lboThis.Items.Clear()
      ' Fill the combobox
      For intI = 0 To intNum - 1
        ' Determine what we want to show for this section
        strSect = intI + 1 & ": [" & pdxSection(intI).Attributes("Location").Value & "] " & _
          ForestText(pdxSection(intI))
        ' Add this to the combobox
        lboThis.Items.Add(strSect)
      Next intI
      ' Select the first section
      lboThis.SelectedIndex = 0
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/SectionsLoad error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ''---------------------------------------------------------------------------------------------------------
  '' Name:       ShowTranslation()
  '' Goal:       Load and show the english translation indicated section of the whole XML document
  '' History:
  '' 28-06-2010  ERK Created
  ''---------------------------------------------------------------------------------------------------------
  'Public Function ShowTranslation(ByRef rtbThis As RichTextBox, ByVal intSection As Integer) As Boolean
  '  Dim strLoc As String          ' Location of this <forest> element
  '  Dim strSeg As String = ""     ' Text of this <forest> element
  '  Dim strLine As String         ' Text of line to be added
  '  Dim intI As Integer           ' A counter
  '  Dim intPtc As Integer         ' Percentage
  '  Dim objText As New StringColl ' The text we read
  '  Dim objTrans As New StringColl ' The translation
  '  Dim ndxThis As XmlNode        ' One node
  '  Dim ndxForest As XmlNode      ' This line's forest
  '  Dim bLeafy As Boolean         ' Whether this is a direct parent of an <eLeaf>
  '  Dim fntNormal As System.Drawing.Font = New Font("Courier New", 10, FontStyle.Regular)
  '  Dim fntBold As System.Drawing.Font = New Font("Courier New", 10, FontStyle.Bold)
  '  Dim fntTrans As System.Drawing.Font = New Font("Times New Roman", 10, FontStyle.Regular)

  '  Try
  '    ' Does this section exist?
  '    If (bSection) OrElse (rtbThis Is Nothing) OrElse (intSection < 0) OrElse _
  '       (intSection > pdxSection.Count - 1) Then
  '      ' Return failure
  '      Return False
  '    End If
  '    ' Clear the richtextbox
  '    rtbThis.Clear()
  '    ' Walk the list
  '    For intI = 0 To intSectSize - 1
  '      ' Show where we are
  '      intPtc = (100 * intI) \ intSectSize
  '      Status("Reading translation page " & intSection & " " & intPtc & "%", intPtc)
  '      ' Make sure we are interruptable
  '      Application.DoEvents()
  '      ' Get the correct forest node
  '      ndxForest = GetForest(intI)
  '      If (ndxForest Is Nothing) Then Return Nothing
  '      ' Access this element
  '      With ndxForest
  '        ' Check this one
  '        If (.Attributes("Location") Is Nothing) OrElse (.Attributes("Location").Value = "") Then
  '          ' Show that this line is kind of empty
  '          strLine = "(code)"
  '          ' Add a new Match element, if there are children (the first child will be an <eTree> element)
  '          If (.ChildNodes.Count > 0) Then
  '            ' Get the first childnode, which should be an <eTree> element
  '            ndxThis = .FirstChild
  '            ' Determine whether this is a direct parent of an <eLeaf>
  '            bLeafy = (ndxThis.FirstChild.Name = "eLeaf")
  '          End If
  '        Else
  '          ' Retrieve details of this element
  '          strLoc = .Attributes("Location").Value
  '          ' Add the location
  '          AppendToRtb(rtbThis, "[" & strLoc & "] ", Color.Black, fntBold)
  '          ' Add the original text
  '          AppendToRtb(rtbThis, ForestText(ndxForest, "org") & vbCrLf, Color.DarkBlue, fntNormal)
  '          ' Get the forest node translation
  '          strLine = ForestText(ndxForest, "eng")
  '          ' Add spaces the size of the location
  '          AppendToRtb(rtbThis, StrDup(strLoc.Length + 3, " "), Color.Black, fntBold)
  '          ' Does a translation exist?
  '          If (strLine = "") Then strLine = "(No translation)"
  '          ' Add the translation line
  '          AppendToRtb(rtbThis, strLine & vbCrLf, Color.DarkGreen, fntTrans)
  '        End If
  '        ' Add the line
  '        objText.Add(strLine)
  '      End With
  '    Next intI
  '    ' Show we are ready
  '    Status("Ready")
  '    ' Return success
  '    Return True
  '  Catch ex As Exception
  '    ' Show error
  '    HandleErr("modMain/ShowTranslation error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
  '    ' Return failure
  '    Return False
  '  End Try
  'End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       ShowTranslation()
  ' Goal:       Load and show the english translation, letting the currently selected line stand out
  ' History:
  ' 07-03-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function ShowTranslation(ByVal bShowHtml As Boolean) As Boolean
    Dim intLine As Integer  ' Line we need to go to
    Dim strHtml As String   ' HTML content

    Try
      ' Validate
      If (pdxList.Count = 0) Then Exit Function
      ' Get the line number
      intLine = GetEdtLine()
      ' Go to the indicated line
      frmMain.GotoTransLine(intLine)
      ' Need to show the HTML?
      If (bShowHtml) Then
        ' Also get and show the translation in HTML
        strHtml = GetTransHtml(frmMain.mnuTransSection.Checked, intLine)
        frmMain.wbReport.DocumentText = strHtml
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/ShowTranslation error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MakeChart
  ' Goal:   Produce a "chart" (for the current section??) of the text
  ' History:
  ' 12-03-2012  ERK Created
  ' 09-07-2013  ERK Added option [bAllSc]
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MakeChart(ByVal bDoSub As Boolean, ByVal bAllSc As Boolean, ByVal strOutType As String, _
                            ByVal intV1ptc As Integer, ByVal intV2ptc As Integer, ByVal intSptc As Integer) As Boolean
    Dim intI As Integer               ' Counter
    Dim intPtc As Integer             ' Percentage
    Dim intCount As Integer           ' Number of <forest> lines
    Dim intSct As Integer = 0         ' Section number
    Dim intSlots As Integer = 0       ' Number of chart slots found
    Dim intV1 As Integer = 0          ' Slot used for V1
    Dim intV2 As Integer = 0          ' Slot used for V2
    Dim strLoc As String = ""         ' Location ID
    Dim strSeg As String = ""         ' Text of one line
    Dim strVern As String = ""        ' Vernacular of this line
    Dim strFile As String = ""        ' Name of the file to be stored
    Dim strText As String = ""        ' Textof the file
    Dim strMsg As String = ""         ' Message report
    Dim ndxThis As XmlNode = Nothing  ' <forest> node
    Dim ndxLast As XmlNode = Nothing  ' Node of last forest
    Dim ndxTop As XmlNode = Nothing   ' Top-level node
    Dim dtrThis As DataRow = Nothing  ' One datarow to work with
    Dim colBack As New StringColl
    Dim colWorder As New StringColl   ' Word order container

    Try
      ' Try creating a dataset for the elements
      If (Not InitChart()) Then Status("Could not initialize [Chart]") : Return False
      ' Find first forest element - depending on the project type
      If (bAllSc) Then
        ' Get the first <forest> node
        ndxThis = pdxList(0)
        If (ndxThis Is Nothing) Then Status("MakeChart error: cannot find first line") : Return False
        ' Get the last <forest> node
        ndxLast = pdxList(pdxList.Count - 1)
      Else
        ndxThis = pdxSection(intCurrentSection)
        If (ndxThis Is Nothing) Then Status("MakeChart error: cannot find first line") : Return False
        ' Try find the "last" forest node\
        If (intCurrentSection = pdxSection.Count - 1) Then
          ndxLast = ndxThis.SelectSingleNode("./following-sibling::forest[last()]")
        Else
          ' ndxLast = pdxSection(intCurrentSection + 1).SelectSingleNode("./preceding-sibling::forest")
          ndxLast = pdxSection(intCurrentSection + 1)
          ndxLast = ndxLast.PreviousSibling
        End If
      End If
      ' Double Check
      If (ndxLast Is Nothing) Then
        ' We cannot figure out what the count is
        intCount = pdxList.Count
      Else
        ' Get number of nodes
        intCount = ndxLast.Attributes("forestId").Value - ndxThis.Attributes("forestId").Value + 1
      End If
      ' Start html output
      colBack.Add("<html><head><style type=" & """" & "Text/css" & """" & ">" & vbCrLf & _
        "tr.top td { border-top: 1px solid black; }" & vbCrLf & _
        "tr.bottom td { border-bottom: 1px solid black; }" & vbCrLf & _
        "tr.row td:first-child { border-left: thin solid black; }" & vbCrLf & _
        "tr.row td:last-child { border-right: thin solid black; }" & vbCrLf & _
        "td.verb { background: lightyellow; }" & vbCrLf & _
        "</style></head><body>")

      ' Start a table
      colWorder.Add("<table><tr><td>Line</td><td>Const</td></td>Text</td></tr>")
      ' Step through them
      intI = 1
      While (Not ndxThis Is Nothing) AndAlso (intI <= intCount)
        ' Show where we are
        intPtc = intI * 100 \ intCount
        Status("Charting " & intPtc & "%", intPtc)
        ' Get the current location
        If (Not GetForestLoc(ndxThis, strLoc)) Then Status("MakeChart error: cannot find location of line " & intI) : Return False
        ' Find the English translation of this node
        strSeg = GetSeg(ndxThis, "eng")
        ' Get the vernacular of this node
        strVern = GetSeg(ndxThis, "org")
        ' Start the line with the vernacular
        colWorder.Add("<tr><td>" & strLoc & "</td><td>")
        ' Add an entry for this line in the chart data structure
        dtrThis = AddOneDataRow(tdlChart, "Clause", "ClauseId", "")
        dtrThis.Item("Vern") = strVern : dtrThis.Item("Eng") = strSeg : dtrThis.Item("LocId") = strLoc
        dtrThis.Item("Type") = "Full"
        ' Walk through all the top-level <eTree> nodes of this one single forest
        ndxTop = ndxThis.FirstChild
        While (Not ndxTop Is Nothing)
          ' Check the type
          If (ndxTop.Name = "eTree") Then
            ' Process this top-level node
            ' (1) What is the IP-type of this node?
            If (DoLike(ndxTop.Attributes("Label").Value, "IP-MAT*|S|CP-QUE*|CP-EXL*")) Then
              ' This is a main-clause IP
              ' Add an entry for this line in the chart data structure
              dtrThis = AddOneDataRow(tdlChart, "Clause", "ClauseId", "")
              dtrThis.Item("Vern") = NodeText(ndxTop) : dtrThis.Item("Eng") = "" : dtrThis.Item("LocId") = strLoc
              ' =============== DEBUG ============
              ' If (InStr(strLoc, ".42") > 0) Then Stop
              ' ==================================
              ' Type depends on the kind of label
              If (DoLike(ndxTop.Attributes("Label").Value, "CP-QUE*|CP-EXL*")) Then
                dtrThis.Item("Type") = "MainWh"
              Else
                dtrThis.Item("Type") = "Main"
              End If
              ' Process the main clause
              colWorder.Add(ChartNode(ndxTop, dtrThis, bDoSub, False))
            ElseIf (DoLike(ndxTop.Attributes("Label").Value, "IP*|SBAR")) Then
              ' THis is an IP, but not a main clause
              ' Add an entry for this line in the chart data structure
              dtrThis = AddOneDataRow(tdlChart, "Clause", "ClauseId", "")
              dtrThis.Item("Vern") = NodeText(ndxTop) : dtrThis.Item("Eng") = "" : dtrThis.Item("LocId") = strLoc
              ' Type depends on the kind of label
              If (DoLike(ndxTop.Attributes("Label").Value, "IP-IMP*")) Then
                dtrThis.Item("Type") = "MainImp"
              Else
                dtrThis.Item("Type") = "Sub"
              End If
              ' Process the subordinate clause
              If (strOutType = "Chart") Then colWorder.Add(ChartNode(ndxTop, dtrThis, bDoSub, (dtrThis.Item("Type") = "Sub")))
            ElseIf (ndxTop.Attributes("Label").Value <> "CODE") Then
              ' This is not a main clause -- add it in brackets?
              colWorder.Add("(" & ndxTop.Attributes("Label").Value & " " & NodeText(ndxTop) & ")")
            End If
          End If
          ' Go to next one
          ndxTop = ndxTop.NextSibling
        End While
        ' Finish this line
        colWorder.Add("</td><td>" & strVern & "</td></tr>")
        ' Is there a back translation?
        If (strSeg <> "") Then
          ' Add the back translation in a separate line
          colWorder.Add("<tr><td colspan=2>-</td><td>" & strSeg & "</td></tr>")
        End If
        '' Add line to collection
        'colBack.Add("<font color='blue' size='1'>[" & strLoc & "] </font>" & strSeg & " ")
        ' GO to next forest
        ndxThis = ndxThis.NextSibling : intI += 1
      End While
      ' Finish this table
      colWorder.Add("</table>")
      ' Action depends on output type
      Select Case strOutType
        Case "WordOrder"
          colBack.Add(colWorder.Text)
        Case "Chart"
          colBack.Add("<h1>Slotting</h1>")
          ' Try determining the slots for the constituents
          If (Not DoChartSlots(intSlots, intV1ptc, intV2ptc, intSptc, intV1, intV2, strMsg)) Then Logging("Problem determining chart slots")
          ' Add the message
          colBack.Add(strMsg)
          ' TODO: determine (a) slot usage and (b) slot entropy
          colBack.Add(DoChartQuant(intSlots, intV1, intV2))
          ' Try produce a nice slotting table 
          colBack.Add(DoChartTable(intSlots, intV1, intV2))
      End Select
      ' Finish HTML output
      colBack.Add("</body></html>")
      ' Save the html output
      strText = colBack.Text
      ' Determine the file name
      strFile = IO.Path.GetDirectoryName(strCurrentFile) & "\" & IO.Path.GetFileNameWithoutExtension(strCurrentFile) & _
        "_" & strOutType & ".htm"
      ' Write the file name
      IO.File.WriteAllText(strFile, strText, System.Text.Encoding.UTF8)
      ' Show the HTML output
      frmMain.wbReport.DocumentText = strText
      Status("File has been saved: " & strFile)
      ' Make a file name for the chart datastructure
      strFile = IO.Path.GetDirectoryName(strCurrentFile) & "\" & IO.Path.GetFileNameWithoutExtension(strCurrentFile) & _
       "_chart.xml"
      ' Write the datastructure there
      tdlChart.WriteXml(strFile)
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/MakeChart error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoChartQuant
  ' Goal:   Take the [tdlChart] dataset, and calculate measures
  ' History:
  ' 26-09-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function DoChartQuant(ByVal intSlotNumber As Integer, ByVal intV1 As Integer, ByVal intV2 As Integer) As String
    Dim dtrFound() As DataRow     ' Result of selecting
    Dim dtrSlot() As DataRow      ' One particular slot
    Dim colBack As New StringColl ' COllection of html string
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intK As Integer           ' Counter
    Dim intFilled As Integer      ' Sparseness
    Dim dblEntr As Double         ' Entropy
    Dim dblUsage As Double        ' Sparseness
    Dim dblFit As Double          ' Goodness of fit
    Dim colEnt As New StringColl  ' Entropy collection
    Dim intPtc As Integer         ' Percentage
    Dim intClauseId As Integer    ' ID of current clause

    Try
      ' Validate
      If (tdlChart Is Nothing) Then Return False
      ' There must be at least some content
      If (tdlChart.Tables("Clause").Rows.Count = 0) Then Return False
      If (tdlChart.Tables("Const").Rows.Count = 0) Then Return False
      ' Start the HTML output
      colBack.Add("<h3>Chart quantities</h3><table><tr><td><b>Slot</b></td>" & _
                  "<td><b>Entropy</b></td><td><b>Usage</b></td><td><b>Fit</b></td></tr>")
      ' Select the lines in order
      dtrFound = tdlChart.Tables("Clause").Select("", "ClauseId ASC")
      If (dtrFound.Length = 0) Then Return ""
      ' Walk all the slots
      For intI = 1 To intSlotNumber
        ' Where are we?
        intPtc = 100 * (intI) \ intSlotNumber
        Status("Quantities " & intPtc & "%", intPtc)
        ' Initialise for this slot
        intFilled = 0 : colEnt.Clear()
        ' Walk all the lines
        For intJ = 0 To dtrFound.Length - 1
          ' Get the element for this clause/slot
          dtrSlot = tdlChart.Tables("Const").Select("ClauseId=" & dtrFound(intJ).Item("ClauseId") & _
                                                        " AND SlotNo=" & intI)
          If (dtrSlot.Length > 0) Then
            ' Keep track of filled slots
            intFilled += 1
            ' Keep track of slot type
            colEnt.AddUnique(dtrSlot(0).Item("Label").ToString)
          End If
        Next intJ
        ' Calculation...
        ' (1) Entropy = amount of chaos. This runs from 1 (low) to 100 (total chaos)
        If (intFilled > 1) Then
          dblEntr = 100 * ((colEnt.Count / intFilled))
        Else
          dblEntr = -1
        End If
        ' (2) Usage: how well is a slot being used? This runs from 1 (low) to 100 (completely used)
        dblUsage = 100 * ((intFilled) / dtrFound.Length)
        ' (3) Fit: how good does this particular slot fare? Runs from -n (bad fit) to 0 (moderate) to +n (good fit)
        dblFit = Math.Log(dblUsage / dblEntr)
        ' dblEntr = 100 * colEnt.Count / dtrFound.Length
        'dblUsage = 100 * (dtrFound.Length - intFilled) / dtrFound.Length
        ' Give the counts for this slot number
        colBack.AddUnique("<tr><td>" & intI & "</td>" & _
                          "<td>" & Format(dblEntr, "0.00") & "</td>" & _
                          "<td>" & Format(dblUsage, "0.00") & "</td>" & _
                          "<td>" & Format(dblFit, "0.00") & "</td>" & _
                          "</tr>")
      Next intI
      ' Finish table
      colBack.Add("</table>")
      ' Return the table
      Return colBack.Text
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/DoChartQuant error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoChartTable
  ' Goal:   Take the [tdlChart] dataset, and transform it into a nice table that can be copied to EXCEL
  ' History:
  ' 26-09-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function DoChartTable(ByRef intSlotNumber As Integer, ByVal intV1 As Integer, ByVal intV2 As Integer) As String
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intK As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage
    Dim intClauseId As Integer    ' ID of current clause
    Dim strText As String         ' Text in one slot
    Dim strLabel As String        ' Label(s) in one slot
    Dim colBack As New StringColl ' COllection of html string
    Dim dtrFound() As DataRow     ' Result of selecting
    Dim dtrChildren() As DataRow  ' All children of me
    Dim dtrSlot() As DataRow      ' All elements for one particular slot

    Try
      ' Validate
      If (tdlChart Is Nothing) Then Return False
      ' There must be at least some content
      If (tdlChart.Tables("Clause").Rows.Count = 0) Then Return False
      If (tdlChart.Tables("Const").Rows.Count = 0) Then Return False
      ' Start the HTML output
      colBack.Add("<table><tr class=" & """" & "bottom" & """" & "><td><b>LocId</b></td>")
      For intI = 1 To intSlotNumber
        ' Is this a special slot?
        If (intI = 1) Then
          colBack.Add("<td><b>Conj</b></td>")
        ElseIf (intI = intV1) Then
          colBack.Add("<td><b>V1</b></td>")
        ElseIf (intI = intV2) Then
          colBack.Add("<td><b>V2</b></td>")
        Else
          colBack.Add("<td><b>Slot" & intI & "</b></td>")
        End If
      Next intI
      ' Finish the first row with column headings
      colBack.Add("</tr>")
      ' Get all the clause rows that are going to be shown in the table
      dtrFound = tdlChart.Tables("Clause").Select("", "ClauseId ASC")
      For intI = 0 To dtrFound.Length - 1
        ' Where are we?
        intPtc = 100 * (intI + 1) \ dtrFound.Length
        Status("Charting " & intPtc & "%", intPtc)
        ' Are there any children here?
        dtrChildren = dtrFound(intI).GetChildRows("Clause_Const")
        If (dtrChildren.Length > 0) Then
          ' Get the Clause ID
          intClauseId = dtrFound(intI).Item("ClauseId")
          ' Start the first row and fill in the LocId
          colBack.Add("<tr><td>" & dtrFound(intI).Item("LocId") & "</td>")
          ' Visit all possible slots of this row
          For intJ = 1 To intSlotNumber
            ' Find out what is in this slot
            dtrSlot = tdlChart.Tables("Const").Select("ClauseId=" & intClauseId & " AND SlotNo='" & intJ & "'")
            'dtrSlot = tdlChart.Tables("Const").Select("ClauseId=" & intClauseId & " AND SlotNo=10")
            strText = ""
            ' Are there more than one candidates for one slot?
            If (dtrSlot.Length = 1) Then
              ' Only one candidate -- simple
              strText = dtrSlot(0).Item("Text")
            Else
              ' More than one candidate!!
              For intK = 0 To dtrSlot.Length - 1
                If (strText <> "") Then strText &= "<br>"
                strText &= "(" & dtrSlot(intK).Item("Label") & " " & dtrSlot(intK).Item("Text") & ")"
              Next intK
            End If
            ' output this slot
            If (intJ = intV1 OrElse intJ = intV2) Then
              colBack.Add("<td class=" & """" & "verb" & """" & ">" & strText & "</td>")
            Else
              colBack.Add("<td>" & strText & "</td>")
            End If
          Next intJ
          ' Finish the first row
          colBack.Add("</tr>")
          ' Start the second row and fill in the LocId
          If (dtrFound(intI).Item("Type") = "Main") Then
            ' Just skip this cell
            colBack.Add("<tr class=" & """" & "bottom" & """" & "><td>&nbsp;</td>")
          Else
            ' Give the special clause type in red
            colBack.Add("<tr class=" & """" & "bottom" & """" & "><td><font color='red' size='1'>" & dtrFound(intI).Item("Type") & "</font></td>")
          End If
          ' Visit all possible slots of this row
          For intJ = 1 To intSlotNumber
            ' Find out what is in this slot
            dtrSlot = tdlChart.Tables("Const").Select("ClauseId=" & intClauseId & " AND SlotNo='" & intJ & "'")
            strLabel = "<font color='blue' size='1'>"
            For intK = 0 To dtrSlot.Length - 1
              If (strLabel <> "") Then strLabel &= " "
              strLabel &= dtrSlot(intK).Item("Label")
              ' Check for empty labels
              If (dtrSlot(intK).Item("Label").ToString = "") Then strLabel &= "&nbsp;"
            Next intK
            If (dtrSlot.Length = 0) Then strLabel &= "&nbsp;"
            ' output this slot, and give borders, depending on which cell it is
            If (intJ = intV1 OrElse intJ = intV2) Then
              colBack.Add("<td class=" & """" & "verb" & """" & ">" & strLabel & "</font></td>")
            Else
              colBack.Add("<td>" & strLabel & "</font></td>")
            End If
          Next intJ
          ' Finish the second row
          colBack.Add("</tr>")
        End If
      Next intI
      ' Finish table
      colBack.Add("</table>")
      ' Return the table
      Return colBack.Text
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/DoChartTable error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoChartSlots
  ' Goal:   Try figuring out which constituent goes into which slot:
  '         (1) First slot is "Conj"
  '         (2) Appositives combine with their predecessor
  '         (3) Same slots combine into one
  '         (4) Similar slots (AP, PP etc) before the V1: combine into one
  '         (5) Determine the position of the V1 slot (must cover >[intV1perc]% of the data)
  '         (6) Shift all constituents with V1 occurring in a slot before the one found in (5)
  '         (7) Determine the position of the V2 slot (must cover >[intV2perc]% of the data)
  '         (8) Shift all constituents with V2 occurring in a slot before the one found in (7)
  '         (9) Determine the position of the S slot (must cover >[intSperc]% of the data)
  '             (Make sure the S slot does NOT overlap with the V1 or V2 slot)
  '         (10)Shift clause-final [AdvCl, SubCl, InfCl] completely to the last slot
  '         (11)Do the subclause test (if there are any subordinate clauses with complementizer)
  '             a. Check for occurrances of [V1 XP V2], where "XP" is not a "NEG" element
  '             b. If (a) yields results, then the "C" position is [intV1]
  '                otherwise the "C" position is the best possible on the left side
  ' History:
  ' 26-09-2012  ERK Created
  ' 09-07-2013  ERK Make sure no other material appears in the V2-slot...
  ' ------------------------------------------------------------------------------------------------------------
  Private Function DoChartSlots(ByRef intSlotNumber As Integer, ByVal intV1perc As Integer, _
        ByVal intV2perc As Integer, ByVal intSperc As Integer, ByRef intV1 As Integer, _
        ByRef intV2 As Integer, ByRef strMsg As String) As Boolean
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intK As Integer           ' Counter
    Dim intSlot As Integer        ' Current slot number
    Dim intNextSlotNo As Integer  ' SLot number of the next constituent
    Dim intFound As Integer       ' Slot number that we found
    Dim intAdd As Integer         ' Value to be added to the V1 slot
    Dim intClauseId As Integer    ' ID of current clause
    Dim intPtc As Integer         ' Percentage
    Dim intSbj As Integer         ' Slot calculated for the subject
    Dim intV1xpV2 As Integer      ' Number of occurrences of [V1 XP V2] with XP not equal to "NEG"
    Dim intV1V2 As Integer        ' Number of occurrences of [V1 ... V2] 
    Dim intPos1 As Integer        ' Position of one slot
    Dim intPos2 As Integer        ' Position of one slot
    Dim dtrFound() As DataRow     ' Result of selecting
    Dim dtrChild As DataRow       ' One child row
    Dim dtrPrev As DataRow        ' Previous child row
    Dim dtrChildren() As DataRow  ' All children of me
    Dim colPos As New Collection  ' Collection of numbers

    Try
      ' Validate
      If (tdlChart Is Nothing) Then Return False
      ' There must be at least some content
      If (tdlChart.Tables("Clause").Rows.Count = 0) Then Return False
      If (tdlChart.Tables("Const").Rows.Count = 0) Then Return False
      ' Set the total number of slots to at least #2
      intSlotNumber = 2
      ' All clause-initial conjunctions go into slot #1
      dtrFound = tdlChart.Tables("Clause").Select("Type <> 'Full'", "ClauseId ASC")
      For intI = 0 To dtrFound.Length - 1
        ' Where are we?
        intPtc = 100 * (intI + 1) \ dtrFound.Length
        Status("Conjunctions " & intPtc & "%", intPtc)
        ' ============= DEBUG ===============
        'intClauseId = dtrFound(intI).Item("ClauseId")
        'If (intClauseId = 65) Then Stop
        ' ===================================
        ' Get the first child [Const] row of this one
        dtrChildren = dtrFound(intI).GetChildRows("Clause_Const")
        ' Initialise the slot number
        intSlot = 1
        ' Are there any children?
        If (dtrChildren.Length > 0) Then
          ' Treat the first child
          dtrChild = dtrChildren(0)
          If (Not dtrChild Is Nothing) Then
            ' Get the label of this constituent
            If (dtrChild.Item("Label") = "Conj") Then
              ' Assign it to the first slot
              dtrChild.Item("SlotNo") = 1
            Else
              ' Otherwise it must be at least in slot 2
              dtrChild.Item("SlotNo") = 2 : intSlot = 2
            End If
          End If
          ' Treat the other children -- starting at slot #2
          dtrPrev = dtrChild
          For intJ = 1 To dtrChildren.Length - 1
            dtrChild = dtrChildren(intJ)
            ' Do we have a previous element?
            If (dtrPrev Is Nothing) Then
              ' No previous element: set the slot to the consecutive number
              intSlot += 1 : dtrChild.Item("SlotNo") = intSlot
              If (intSlot > intSlotNumber) Then intSlotNumber = intSlot
            Else
              ' If we are an [Appos], we belong to the previous one
              If (dtrChild.Item("Label") = "Appos") Then
                dtrChild.Item("SlotNo") = intSlot   ' Do NOT increment
              ElseIf (dtrChild.Item("Label") = dtrPrev.Item("Label")) Then
                ' We are the same as the previous --> assign same slot number
                dtrChild.Item("SlotNo") = intSlot   ' Do NOT increment
              Else
                ' Increment slot number
                intSlot += 1 : dtrChild.Item("SlotNo") = intSlot
                If (intSlot > intSlotNumber) Then intSlotNumber = intSlot
              End If
            End If
            ' Switch to next row
            dtrPrev = dtrChild
          Next intJ
        End If
      Next intI
      ' Adapt the material preceding V1: conflate AP, PP, PpCl, SubCl into one slot
      ' (This works for main and subordinate clauses)
      dtrFound = tdlChart.Tables("Clause").Select("", "ClauseId ASC")
      For intI = 0 To dtrFound.Length - 1
        ' Where are we?
        intPtc = 100 * (intI + 1) \ dtrFound.Length
        Status("Conflating " & intPtc & "%", intPtc)
        ' Check if this clause has a V1 slot
        intClauseId = dtrFound(intI).Item("ClauseId")
        ' ============= DEBUG ===============
        'If (intClauseId = 65) Then Stop
        ' ===================================
        If (tdlChart.Tables("Const").Select("ClauseId = " & intClauseId & " AND Label='V1'").Length > 0) Then
          ' There is at least 1 child with V1 --> get all the children
          dtrChildren = dtrFound(intI).GetChildRows("Clause_Const")
          intSlot = -1
          For intJ = 0 To dtrChildren.Length - 1
            ' Is this the V1 slot already?
            Select Case dtrChildren(intJ).Item("Label")
              Case "V1"
                ' This is the V1 slot, so there is nothing more we can do at this point
                Exit For
              Case "AP", "PP", "PpCl", "SubCl", "AdjP"
                ' Is this the first occurrence?
                If (intSlot < 0) Then
                  intSlot = dtrChildren(intJ).Item("SlotNo")
                Else
                  ' Assign this one the right slot number
                  dtrChildren(intJ).Item("SlotNo") = intSlot
                  ' Shift all the following slot numbers down
                  For intK = intJ + 1 To dtrChildren.Length - 1
                    dtrChildren(intK).Item("SlotNo") -= 1
                  Next intK
                End If
            End Select
          Next intJ
        End If
      Next intI
      ' Figure out what the best slot is going to be for V1
      strMsg = ""
      If (Not GetChartSlot("V1", intV1perc, 2, intV1, strMsg)) Then Status("Could not get V1 slot") : Return False
      ' Shift everything that has a V1 in a slot LOWER than the V1 slot we have determined
      ' (N.B: we leave intact the instances where the V1 is HIGHER than the slot we determined...)
      dtrFound = tdlChart.Tables("Clause").Select("", "ClauseId ASC")
      For intI = 0 To dtrFound.Length - 1
        ' Where are we?
        intPtc = 100 * (intI + 1) \ dtrFound.Length
        Status("Shifting V1 " & intPtc & "%", intPtc)
        ' Check if this clause has a V1 slot
        intClauseId = dtrFound(intI).Item("ClauseId")
        ' ============= DEBUG ===============
        'If (intClauseId = 65) Then Stop
        ' ===================================
        If (tdlChart.Tables("Const").Select("ClauseId = " & intClauseId & " AND Label='V1'").Length > 0) Then
          ' There is at least 1 child with V1 --> get all the children
          dtrChildren = dtrFound(intI).GetChildRows("Clause_Const")
          intSlot = -1
          For intJ = 0 To dtrChildren.Length - 1
            ' Is this the V1 slot?
            If (dtrChildren(intJ).Item("Label") = "V1") Then
              ' Check if the slot value is okay
              If (dtrChildren(intJ).Item("SlotNo") >= intV1) Then
                ' Slot value is okay: leave the for loop
                Exit For
              Else
                ' Slot value for V1 is not okay: get the difference
                intAdd = intV1 - dtrChildren(intJ).Item("SlotNo")
                ' Add this value to this and subsequent slots
                For intK = intJ To dtrChildren.Length - 1
                  ' Add it
                  dtrChildren(intK).Item("SlotNo") += intAdd
                Next intK
                ' Possibly adapt the maximum slot number
                If (dtrChildren(intK - 1).Item("SlotNo") > intSlotNumber) Then intSlotNumber = dtrChildren(intK - 1).Item("SlotNo")
                ' Exit the for loop
                Exit For
              End If
            End If
          Next intJ
        End If
      Next intI
      ' Figure out what the best slot is going to be for V2
      If (Not GetChartSlot("V2", intV2perc, intV1, intV2, strMsg)) Then Status("Could not get V2 slot") : Return False

      ' Shift everything that has a V2 in a slot LOWER than the V2 slot we have determined
      ' (N.B: we leave intact the instances where the V2 is HIGHER than the slot we determined...)
      dtrFound = tdlChart.Tables("Clause").Select("", "ClauseId ASC")
      For intI = 0 To dtrFound.Length - 1
        ' Where are we?
        intPtc = 100 * (intI + 1) \ dtrFound.Length
        Status("Shifting V2 " & intPtc & "%", intPtc)
        ' Check if this clause has a V2 slot
        intClauseId = dtrFound(intI).Item("ClauseId")
        ' ============= DEBUG ===============
        'If (intClauseId = 65) Then Stop
        ' ===================================
        If (tdlChart.Tables("Const").Select("ClauseId = " & intClauseId & " AND Label='V2'").Length > 0) Then
          ' There is at least 1 child with V2 --> get all the children
          dtrChildren = dtrFound(intI).GetChildRows("Clause_Const")
          intSlot = -1
          For intJ = 0 To dtrChildren.Length - 1
            ' Is this the V2 slot?
            If (dtrChildren(intJ).Item("Label") = "V2") Then
              ' Check if the slot value is okay
              If (dtrChildren(intJ).Item("SlotNo") >= intV2) Then
                ' Slot value is okay: leave the for loop
                Exit For
              Else
                ' Slot value for V2 is not okay: get the difference
                intAdd = intV2 - dtrChildren(intJ).Item("SlotNo")
                ' Add this value to this and subsequent slots
                For intK = intJ To dtrChildren.Length - 1
                  ' Add it
                  dtrChildren(intK).Item("SlotNo") += intAdd
                Next intK
                ' Possibly adapt the maximum slot number
                If (dtrChildren(intK - 1).Item("SlotNo") > intSlotNumber) Then intSlotNumber = dtrChildren(intK - 1).Item("SlotNo")
                ' Exit the for loop
                Exit For
              End If
            End If
          Next intJ
        Else
          ' There is no V2 element in this clause, but we have to shift if there is something in the V2 slot...
          dtrChildren = tdlChart.Tables("Const").Select("ClauseId = " & intClauseId & " AND SlotNo >= " & intV2, "SlotNo ASC")
          For intJ = 0 To dtrChildren.Length - 1
            ' Shift this one up
            dtrChild = dtrChildren(intJ)
            dtrChild.Item("SlotNo") += 1
          Next intJ
        End If
      Next intI
      ' Try to determine the subject slot position
      If (Not GetChartSlot("S", intSperc, 2, intSbj, strMsg)) Then Status("Could not get Sbj slot") : Return False
      ' Double check the S slot: it must not coincide with the V1 or V2 slot
      If (intV2 = intSbj) AndAlso (intSbj > 2) Then
        ' English language specific: move the Sbj slot to the left
        intSbj -= 1
        strMsg &= "<br>Sbj slot shifted left to avoid collision with V2 slot" & vbCrLf
      End If
      If (intV1 = intSbj) AndAlso (intSbj > 2) Then
        ' English language specific: move the Sbj slot to the left
        intSbj -= 1
        strMsg &= "<br>Sbj slot shifted left to avoid collision with V1 slot" & vbCrLf
      End If
      If (intV1 = intSbj OrElse intV2 = intSbj) Then
        ' There still is collision - unavoidable
        strMsg &= "<br>Sbj slot collides with V1 or V2 slot, but this is unavoidable here" & vbCrLf
      End If
      ' Subject shifting: whenever the S slot is lower than [intSbj] we found...
      ' (1) if S before [intV1], then move S as far RIGHT to [intV1-1] as possible (provided slots are empty)
      ' (2) if S after [intV1] (but before [intSbj]): shift rightward starting from S
      ' (3) if S after [intSbj] --> leave
      dtrFound = tdlChart.Tables("Clause").Select("", "ClauseId ASC")
      For intI = 0 To dtrFound.Length - 1
        ' Where are we?
        intPtc = 100 * (intI + 1) \ dtrFound.Length
        Status("Shifting S " & intPtc & "%", intPtc)
        ' Check if this clause has an S slot
        intClauseId = dtrFound(intI).Item("ClauseId")
        ' ============= DEBUG ===============
        'If (intClauseId = 65) Then Stop
        ' ===================================
        If (tdlChart.Tables("Const").Select("ClauseId = " & intClauseId & " AND Label='S'").Length > 0) Then
          ' There is at least 1 child with S --> get all the children
          dtrChildren = dtrFound(intI).GetChildRows("Clause_Const")
          intSlot = -1
          For intJ = 0 To dtrChildren.Length - 1
            ' Is this the S slot?
            If (dtrChildren(intJ).Item("Label") = "S") Then
              ' Get the slot number
              intFound = dtrChildren(intJ).Item("SlotNo")
              ' Check if the slot value is okay
              If (intFound >= intSbj) Then
                ' Slot value is okay: leave the for loop
                Exit For
              Else
                ' Slot value for S is not okay: [intFound] < [intSbj]
                ' (1) is the subject before [intV1]?
                If (intFound < intV1 - 1) Then
                  ' Try to shift the subject as far right as possible, but keep before [intV1]
                  ' (a) are there any right neighbours?
                  If ((intJ + 1) < (dtrChildren.Length - 1)) Then
                    ' (a) what is the slot number of the next constituent?
                    intNextSlotNo = dtrChildren(intJ + 1).Item("SlotNo")
                    ' (b) Can we shift the S then?
                    If (intFound < intNextSlotNo - 1) Then
                      ' (c) Shift the S to the position immediately before the next occupied slot
                      dtrChildren(intJ).Item("SlotNo") = intNextSlotNo - 1
                    End If
                  Else
                    ' Too little space: we can adapt straightforwardly
                    dtrChildren(intJ).Item("SlotNo") = intV1 - 1
                  End If
                  ' Exit the for loop
                  Exit For
                ElseIf (intFound > intV1) Then
                  ' Are we still before [intV2]?
                  Stop
                End If
              End If
            End If
          Next intJ
        End If
      Next intI
      ' Shift the clause-final [AdvCl, SubCl, InfCl] to [intSlotNumber]
      dtrFound = tdlChart.Tables("Clause").Select("", "ClauseId ASC")
      For intI = 0 To dtrFound.Length - 1
        ' Where are we?
        intPtc = 100 * (intI + 1) \ dtrFound.Length
        Status("Shifting final clauses " & intPtc & "%", intPtc)
        ' Get all children
        dtrChildren = dtrFound(intI).GetChildRows("Clause_Const")
        ' Check the final child
        If (dtrChildren.Length > 0) Then
          intSlot = dtrChildren.Length - 1
          If (DoLike(dtrChildren(intSlot).Item("Label"), "AdvCl|InfCl|SubCl")) Then
            ' Fix the slot to the last one
            dtrChildren(intSlot).Item("SlotNo") = intSlotNumber
          End If
        End If
      Next intI
      ' Try finding [V1 XP V2] occurrences in subordinate clauses
      intV1xpV2 = 0 : intV1V2 = 0
      dtrFound = tdlChart.Tables("Clause").Select("Type LIKE 'Sub*'", "ClauseId ASC")
      For intI = 0 To dtrFound.Length - 1
        ' Where are we?
        intPtc = 100 * (intI + 1) \ dtrFound.Length
        Status("Middlefield " & intPtc & "%", intPtc)
        ' Get all children
        dtrChildren = dtrFound(intI).GetChildRows("Clause_Const")
        intPos1 = -1 : intPos2 = -1
        ' Walk the children
        For intJ = 0 To dtrChildren.Length - 1
          ' Look for V1 and V2 children
          If (dtrChildren(intJ).Item("Label") = "V1") Then
            intPos1 = intJ
          ElseIf (dtrChildren(intJ).Item("Label") = "V2") Then
            intPos2 = intJ : Exit For
          End If
        Next intJ
        ' Do we have V1 and V2?
        If (intPos1 > 0) AndAlso (intPos2 > 0) Then
          ' There is a V1 ... V2
          intV1V2 += 1
          ' See if there is any intervenor
          If (intPos1 + 2 = intPos2) Then
            ' If the intervenor is not a NEG, then increment V1-XP-V2
            If (Not DoLike(dtrChildren(intPos1 + 1).Item("Label"), "NEG*|Neg*|AP|FP|ALSO")) Then
              intV1xpV2 += 1
              ' Show the clause in red
              strMsg &= "<br><font color='red'>" & intV1V2 & " " & ChartClause(dtrChildren) & "</font>" & vbCrLf
            End If
          ElseIf (intPos1 + 2 < intPos2) Then
            ' Ther must be an intervenor other than NEG
            intV1xpV2 += 1
            ' Show the clause in red
            strMsg &= "<br><font color='red'>" & intV1V2 & " " & ChartClause(dtrChildren) & "</font>" & vbCrLf
          Else
            ' Show the clause
            strMsg &= "<br>" & intV1V2 & " " & ChartClause(dtrChildren) & vbCrLf
          End If
        End If
      Next intI
      ' Give the results
      strMsg &= "<br>V1...V2 = " & intV1V2 & " while V1-XP-V2 = " & intV1xpV2 & vbCrLf

      ' Add the total number of slots needed
      strMsg = "<br>Number of slots = " & intSlotNumber & "<br><br>" & strMsg & vbCrLf
      ' Show we are ready
      Status("")
      ' =========== Debugging =============
      ' MsgBox("Number of slots = " & intSlotNumber & vbCrLf & vbCrLf & strMsg)
      ' ===================================
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/DoChartSlots error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   ChartClause
  ' Goal:   Give a string of all the constituents in [dtrChildren]
  ' History:
  ' 2296-09-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function ChartClause(ByRef dtrChildren() As DataRow) As String
    Dim intI As Integer           ' Counter
    Dim strBack As String = ""    ' What we return

    Try
      ' Validate
      If (dtrChildren Is Nothing) Then Return ""
      ' Walk these children
      For intI = 0 To dtrChildren.Length - 1
        ' Space?
        If (strBack <> "") Then strBack &= " "
        ' Get the label and the constituent
        strBack &= "[" & dtrChildren(intI).Item("Label") & " " & VernToEnglish(dtrChildren(intI).Item("Text")) & "]"
      Next intI
      ' Return the result
      Return strBack
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/ChartClause error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetChartSlot
  ' Goal:   Figure out the most likely chart slot for the indicated label [strLabel]
  '         This slot must cover more than [dblPerc] % of the data
  '         The position of the slot must be higher than [intMinSlot]
  ' History:
  ' 26-09-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function GetChartSlot(ByVal strLabel As String, ByVal dblPerc As Double, _
        ByVal intMinSlot As Integer, ByRef intSlot As Integer, ByRef strMsg As String) As Boolean
    Dim intI As Integer           ' Counter
    Dim intMin As Integer         ' Minimal frequency
    Dim intMax As Integer         ' Maximum frequency
    Dim arFreq() As Integer       ' Array with frequencies
    Dim dblTotal As Double        ' Total number
    Dim dblFreq As Double         ' Frequency
    Dim dtrFound() As DataRow     ' Result of selecting
    Dim strSelect As String = "(Parent.Type = 'Main') AND Label='"

    Try
      ' Figure out what the best slot is going to be for [strLabel]
      '  (Only take into account MAIN clauses)
      dtrFound = tdlChart.Tables("Const").Select(strSelect & strLabel & "'", "SlotNo ASC")
      If (dtrFound.Length > 0) Then
        ' Get the total number (for the frequency)
        dblTotal = dtrFound.Length
        ' Get the minimum and maximum number for the slot size
        intMin = dtrFound(0).Item("SlotNo")
        intMax = dtrFound(dtrFound.Length - 1).Item("SlotNo")
        ' Initialise [intSlot] as one further than the [intMinSlot] slot
        intSlot = intMinSlot + 1
        ' Make room for messages
        strMsg &= vbCrLf & "<br>Frequencies for " & strLabel & ":" & vbCrLf
        ' Is there a difference?
        If (intMin = intMax) Then
          ' We have found the slot number for [strLabel], and no further work is needed
          intSlot = intMin
          strMsg &= "<br>(minimum and maximum slot both are " & intMin & ")" & vbCrLf
        Else
          ' Start table
          strMsg &= "<table><tr><td>Position</td><td>Frequency</td><td>Percentage</td></tr>" & vbCrLf
          ' Fill an array with frequencies
          ReDim arFreq(0 To intMax - intMin) : dblFreq = 0
          For intI = intMin To intMax
            ' Get the amount of constituents with [strLabel] in this slot
            arFreq(intI - intMin) = tdlChart.Tables("Const").Select(strSelect & strLabel & "' AND SlotNo = " & intI).Count
            dblFreq += arFreq(intI - intMin)
            strMsg &= "<tr><td>" & intI & "</td><td>" & arFreq(intI - intMin) & _
                  "</td><td>" & Format((100 * dblFreq / dblTotal), "0.0") & "%</td></tr>" & vbCrLf
          Next intI
          ' Finish table 
          strMsg &= "</table>" & vbcrlf
          ' Determine the [dblPerc]% best slot
          dblFreq = 0
          For intI = 0 To intMax - intMin
            ' Add this frequency
            dblFreq += arFreq(intI)
            ' Are we above the [dblPerc]%?
            If ((100 * dblFreq / dblTotal) > dblPerc) Then
              intSlot = intI + intMin
              Exit For
            End If
          Next intI
        End If
        strMsg &= "<br>Cut-off = " & Format(dblPerc, "0") & " > " & strLabel & " slot = " & intSlot & "<p></p>" & vbCrLf
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/DoChartSlots error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   AddChartConst
  ' Goal:   Add one constituent to the chart
  ' History:
  ' 26-09-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function AddChartConst(ByRef dtrParent As DataRow, ByVal strLabel As String, ByVal strText As String) As Boolean
    Dim dtrConst As DataRow = Nothing   ' One constituent

    Try
      ' Validate
      If (dtrParent Is Nothing) Then Return False
      ' Add a constituent row
      dtrConst = AddOneDataRowWithParent(tdlChart, "Const", "ConstId", dtrParent)
      ' Fill in the elements
      With dtrConst
        .Item("Label") = strLabel
        .Item("Text") = strText
        .Item("SlotNo") = -1    ' Signals: unassigned
        .Item("ClauseId") = dtrParent.Item("ClauseId")
      End With
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/MakeChart error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   ChartNode
  ' Goal:   Get the HTML text of ONE sentence
  '         But the sentence may consist of several sub-clauses, each with their own row in the chart!!
  '         Subordinate clause:
  '         - Only those with C <> '0'
  '         - Include also IP-MAT [ IP-MAT CONJP ] instances
  ' History:
  ' 12-03-2012  ERK Created
  ' 28-09-2012  ERK Expanded
  ' ------------------------------------------------------------------------------------------------------------
  Public Function ChartNode(ByRef ndxThis As XmlNode, ByRef dtrParent As DataRow, ByVal bDoSub As Boolean, _
                            ByVal bIsSubCl As Boolean) As String
    Dim ndxWork As XmlNode = Nothing  ' Working node
    Dim ndxConst As XmlNodeList       ' List of child nodes of the top-level one
    Dim ndxChild As XmlNode = Nothing ' Child node
    Dim dtrSub As DataRow = Nothing   ' Subclause datarow
    Dim strLabel As String            ' Value of current node's label
    Dim strBack As String = ""        ' What we return
    Dim strNode As String             ' The text of the node
    Dim colConst As New StringColl    ' Constituent types
    Dim colBack As New StringColl     ' What we will return
    Dim colSub As New StringColl      ' Result of subclause
    Dim intLeft As Integer            ' Position of the left verb
    Dim intRight As Integer           ' Position of the right verb
    Dim intI As Integer               ' Counter
    Dim bSub As Boolean = False       ' Do we have a subject already?

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return ""
      If (ndxThis.Name <> "eTree") Then Return ""
      ' Are we into a subclause?
      If (bIsSubCl) AndAlso (Not ndxThis.SelectSingleNode("./child::eTree[@Label = 'C']") Is Nothing) Then
        ' Get the immediate children, which are (FP) C IP-SUB
        ndxConst = ndxThis.SelectNodes("./child::eTree")
        ' Process these children
        For intI = 0 To ndxConst.Count - 1
          ' Determine the type of this constituent
          strLabel = ndxConst(intI).Attributes("Label").Value
          ' Get the text of this node
          strNode = VernToEnglish(NodeText(ndxConst(intI)))
          ' Action depends on type
          If (strLabel = "C") Then
            colConst.Add("C") : AddChartConst(dtrParent, "C", strNode)
          ElseIf (strLabel Like "FP*") Then
            colConst.Add("FP") : AddChartConst(dtrParent, "FP", strNode)
          ElseIf (strLabel Like "IP*") Then
            ' Get my children
            ndxConst = ndxConst(intI).SelectNodes("./child::eTree")
            ' Exit the loop
            Exit For
          Else
            colConst.Add("Other[" & strLabel & "]") : AddChartConst(dtrParent, "Other[" & strLabel & "]", strNode)
          End If
        Next intI
      ElseIf (bIsSubCl) AndAlso (ndxThis.Attributes("Label").Value Like "CONJP*") Then
        ' Get the immediate children, which are (FP) C IP-SUB
        ndxConst = ndxThis.SelectNodes("./child::eTree")
        ' Process these children
        For intI = 0 To ndxConst.Count - 1
          ' Determine the type of this constituent
          strLabel = ndxConst(intI).Attributes("Label").Value
          ' Get the text of this node
          strNode = VernToEnglish(NodeText(ndxConst(intI)))
          ' Action depends on type
          If (strLabel = "CONJ") Then
            colConst.Add("Conj") : AddChartConst(dtrParent, "Conj", strNode)
          ElseIf (strLabel Like "IP*") Then
            ' Get my children
            ndxConst = ndxConst(intI).SelectNodes("./child::eTree")
            ' Exit the loop
            Exit For
          Else
            colConst.Add("Other[" & strLabel & "]") : AddChartConst(dtrParent, "Other[" & strLabel & "]", strNode)
          End If
        Next intI
      Else
        ' Break this up into constituents
        ndxConst = ndxThis.SelectNodes("./child::eTree")
      End If
      ' Initialize the non-finite verb and the finite-verb slot locations
      intLeft = -1 : intRight = -1
      ' Walk through constituents, and give each of them a type
      For intI = 0 To ndxConst.Count - 1
        ' Check for empty constituent
        If (ndxConst(intI).SelectSingleNode("./child::eLeaf[@Type='Star']") Is Nothing) Then
          ' Determine the type of this constituent
          strLabel = ndxConst(intI).Attributes("Label").Value
          ' Get the text of this node
          strNode = VernToEnglish(NodeText(ndxConst(intI)))
          ' =========== debug ===========
          ' If (strLabel = "CONJP") Then Stop
          ' =============================
          If (DoLike(strLabel, strFiniteVerb)) Then
            colConst.Add("V1") : AddChartConst(dtrParent, "V1", strNode)
            intLeft = intI
          ElseIf (DoLike(strLabel, strAnyVerb)) Then
            colConst.Add("V2") : AddChartConst(dtrParent, "V2", strNode)
            intRight = intI
          ElseIf (DoLike(strLabel, "*NOM*|*SBJ*")) AndAlso (Not strLabel Like "*PRD*") Then
            ' Check for empty subjects --> they have to be filled in
            If (strNode = "") Then strNode = "___"
            ' This could be a subject, but not necessarily
            If (bSub) Then
              colConst.Add("O1") : AddChartConst(dtrParent, "O1", strNode)
              ' colConst.Add("S") : AddChartConst(dtrParent, "S", strNode)
            Else
              ' Not a subject yet
              colConst.Add("S") : AddChartConst(dtrParent, "S", strNode) : bSub = True
            End If
          ElseIf (DoLike(strLabel, "NP*OB1*|NP*PRD*|NP-ACC*")) Then
            colConst.Add("O1") : AddChartConst(dtrParent, "O1", strNode)
          ElseIf (strLabel Like "NP*OB2*|NP-DAT*|NP-INS*") Then
            colConst.Add("O2") : AddChartConst(dtrParent, "O2", strNode)
          ElseIf (strLabel Like "NEG") Then
            colConst.Add("Neg") : AddChartConst(dtrParent, "Neg", strNode)
          ElseIf (strLabel Like "PP*") Then
            ' Check the child nodes
            ndxChild = ndxConst(intI).SelectSingleNode("./child::eTree[tb:Like(string(@Label), 'CP*')]", conTb)
            If (ndxChild Is Nothing) Then
              colConst.Add("PP") : AddChartConst(dtrParent, "PP", strNode)
            Else
              colConst.Add("PpCl") : AddChartConst(dtrParent, "PpCl", strNode)
            End If
          ElseIf (DoLike(strLabel, "NP*|ADV*|PTP*|INTJ|FP|ALSO")) Then
            colConst.Add("AP") : AddChartConst(dtrParent, "AP", strNode)
          ElseIf (strLabel Like "*PRN*") Then
            colConst.Add("Appos") : AddChartConst(dtrParent, "Appos", strNode)
          ElseIf (DoLike(strLabel, "CONJ|CONJ-*")) Then
            colConst.Add("Conj") : AddChartConst(dtrParent, "Conj", strNode)
          ElseIf (strLabel Like "ADJP*") Then
            colConst.Add("AdjP") : AddChartConst(dtrParent, "AdjP", strNode)
            ' This could be the right bracket
            If (intRight < 0) Then intRight = intI
          ElseIf (DoLike(strLabel, "CP-REL*|CP-CLF*")) Then
            ' This is a relative clause, probably extraposed, since it does not occurr under an IP-SUB
            colConst.Add("RC") : AddChartConst(dtrParent, "RC", strNode)
          ElseIf (DoLike(strLabel, "IP-MAT*|CP-THT*|CP-CAR*|CP-EXL*|CONJP*")) Then
            ' This is a subordinate clause, which deserves its own treatment
            If (bDoSub) AndAlso (IsSubClause(ndxConst(intI)) OrElse (DoLike(strLabel, "IP-MAT*|CONJP*"))) Then
              ' Create a subclause datarow
              dtrSub = AddOneDataRow(tdlChart, "Clause", "ClauseId", "")
              dtrSub.Item("LocId") = dtrParent.Item("LocId") : dtrSub.Item("Type") = "Sub"
              dtrSub.Item("Eng") = "" : dtrSub.Item("Vern") = VernToEnglish(NodeText(ndxConst(intI)))
              ' Process the subclause
              colConst.Add("SubCl[" & ChartNode(ndxConst(intI), dtrSub, bDoSub, True) & "]")
            Else
              colConst.Add("SubCl") : AddChartConst(dtrParent, "SubCl", strNode)
            End If
          ElseIf (DoLike(strLabel, "CP-ADV*|IP-PPL*")) Then
            ' This is a subordinate clause, which deserves its own treatment
            If (bDoSub) Then
              ' Create a subclause datarow
              dtrSub = AddOneDataRow(tdlChart, "Clause", "ClauseId", "")
              dtrSub.Item("LocId") = dtrParent.Item("LocId") : dtrSub.Item("Type") = "Sub"
              dtrSub.Item("Eng") = "" : dtrSub.Item("Vern") = VernToEnglish(NodeText(ndxConst(intI)))
              ' Process the subclause
              colConst.Add("AdvCl[" & ChartNode(ndxConst(intI), dtrSub, bDoSub, False) & "]")
            Else
              colConst.Add("AdvCl") : AddChartConst(dtrParent, "AdvCl", strNode)
            End If
          ElseIf (strLabel Like "CP-INF*|IP-INF*") Then
            ' This is a subordinate clause, which deserves its own treatment
            If (bDoSub) Then
              ' Create a subclause datarow
              dtrSub = AddOneDataRow(tdlChart, "Clause", "ClauseId", "")
              dtrSub.Item("LocId") = dtrParent.Item("LocId") : dtrSub.Item("Type") = "Sub"
              dtrSub.Item("Eng") = "" : dtrSub.Item("Vern") = VernToEnglish(NodeText(ndxConst(intI)))
              ' Process the subclause
              colConst.Add("InfCl[" & ChartNode(ndxConst(intI), dtrSub, bDoSub, False) & "]")
            Else
              colConst.Add("InfCl") : AddChartConst(dtrParent, "InfCl", strNode)
            End If
          ElseIf (strLabel Like "CP-QUE*") Then
            ' This is a subordinate clause, which deserves its own treatment
            If (bDoSub) Then
              ' Create a subclause datarow
              dtrSub = AddOneDataRow(tdlChart, "Clause", "ClauseId", "")
              dtrSub.Item("LocId") = dtrParent.Item("LocId") : dtrSub.Item("Type") = "Sub"
              dtrSub.Item("Eng") = "" : dtrSub.Item("Vern") = VernToEnglish(NodeText(ndxConst(intI)))
              ' Process the subclause
              colConst.Add("WhCl[" & ChartNode(ndxConst(intI), dtrSub, bDoSub, False) & "]")
            Else
              colConst.Add("WhCl") : AddChartConst(dtrParent, "WhCl", strNode)
            End If
          ElseIf (DoLike(strLabel, ",|.|'|:|;|[?]|!|CODE|E_S|" & """")) Then
            ' This is punctuation -- we skip it
            ' colConst.Add("Punct")
            strLabel = strLabel
          Else
            colConst.Add("Other[" & strLabel & "]") : AddChartConst(dtrParent, "Other[" & strLabel & "]", strNode)
          End If
        End If
      Next intI
      ' Check for several different situations
      If (intLeft >= 0 AndAlso intRight >= 0) Then
        ' This contains a finite and a non-finite verb
        ' (a) How many have we got on the left?
        Select Case intLeft
          Case 0  ' Verb initial clause...
          Case 1  ' Verb is preceded by 1 element
            ' Check the nature of the first element
          Case 2  ' Verb is preceded by 2 elements
          Case Else ' Finite verb, but too far away...
        End Select

      ElseIf (intLeft >= 0) Then
        ' THis only contains a finite verb
        Select Case intLeft
          Case 0  ' Verb initial clause...
          Case 1  ' Verb is preceded by 1 element
          Case 2  ' Verb is preceded by 2 elements
          Case Else ' Finite verb, but too far away...
        End Select

      ElseIf (intRight >= 0) Then
        ' This only contains a non-finite verb
        Select Case intRight
          Case 0  ' Verb initial clause...
          Case 1  ' Verb is preceded by 1 element
          Case 2  ' Verb is preceded by 2 elements
          Case Else ' Finite verb, but too far away...
        End Select

      Else
        ' There is no verb what soever --> cannot make anything of this!
        colBack.Add("")
      End If
      ' =========================================================================
      ' 25/Sep/2012: preliminary solution --> Add the constituent's global names
      ' =========================================================================
      If (colConst.Count > 0) Then
        strBack = colConst.Item(0)
        For intI = 1 To colConst.Count - 1
          If (intI = intLeft) Then
            strBack &= " [" & colConst.Item(intI)
          ElseIf (intI = intRight) Then
            strBack &= " " & colConst.Item(intI) & "]"
          Else
            strBack &= " " & colConst.Item(intI)
          End If
        Next intI
      Else
        strBack = ""
      End If
      ' Return the result
      Return strBack
      ' Return colBack.Text
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/ChartNode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  IsSubClause
  ' Goal :  Check if the node is the start of a proper subclause
  ' History:
  ' 29-09-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Function IsSubClause(ByRef ndxThis As XmlNode) As Boolean
    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' There should be a "C" labelled child
      If (Not ndxThis.SelectSingleNode("./child::eTree[@Label = 'C']") Is Nothing) Then
        ' The child of the "C" should NOT be "zero"
        Return (ndxThis.SelectSingleNode("./child::eTree[@Label = 'C' and child::eLeaf[@Text = '0']]") Is Nothing)
      End If
      ' This is not a proper subclause
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/IsSubClause error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  ViewFileContent
  ' Goal :  Load the indicates PSDX file and show it in the web browser window
  ' History:
  ' 15-04-2010  ERK Created
  ' 27-11-2012  ERK Added possibility to view a gloss (in brackets)
  ' ---------------------------------------------------------------------------------------------------------
  Public Sub ViewFileContent(ByVal strLang As String, Optional ByVal strGloss As String = "")
    Dim intI As Integer               ' Counter
    Dim intPtc As Integer             ' Percentage
    Dim intCount As Integer           ' Number of <forest> lines
    Dim intSct As Integer = 0         ' Section number
    Dim strLoc As String = ""         ' Location ID
    Dim strSeg As String = ""         ' Text of one line
    Dim strGls As String = ""         ' Gloss of one line
    Dim strFile As String = ""        ' Name of the file to be stored
    Dim strText As String = ""        'Textof the file
    Dim ndxThis As XmlNode = Nothing  ' <forest> node
    Dim colBack As New StringColl

    Try
      ' Adapt the file to the actual one
      strFile = IO.Path.GetDirectoryName(strCurrentFile) & "\" & _
        IO.Path.GetFileNameWithoutExtension(strCurrentFile) & "-" & strLang & ".html"
      ' Find first forest element - depending on the project type
      If (Not GetFirstForest(pdxCurrentFile, ndxThis)) Then Status("ViewFile error: cannot find first line") : Exit Sub
      ' Get number of nodes
      intCount = ndxThis.ParentNode.ChildNodes.Count
      ' Start html output
      colBack.Add("<html><body>")
      ' Step through them
      intI = 1
      While (Not ndxThis Is Nothing)
        ' Show where we are
        intPtc = intI * 100 \ intCount
        Status("Loading " & strFile & " " & intPtc & "%", intPtc)
        ' Get the current location
        If (Not GetForestLoc(ndxThis, strLoc)) Then Status("ViewFile error: cannot find location of line " & intI) : Exit Sub
        ' Do we need to add a section break?
        If (Not ndxThis.Attributes("Section") Is Nothing) Then
          ' Increment section number
          intSct += 1
          ' Add a section break
          colBack.Add("<p> --- Section " & intSct & " --- <p>")
        End If
        ' Find the text of this node
        strSeg = GetSeg(ndxThis, strLang)
        ' Need a gloss?
        If (strGloss = "") Then
          ' Add line to collection
          colBack.Add("<font color='blue' size='1'>[" & strLoc & "] </font>" & strSeg & " ")
        Else
          ' Get the gloss segment
          strGls = GetSeg(ndxThis, strGloss)
          ' Add line and the gloss to collection
          colBack.Add("<font color='blue' size='1'>[" & strLoc & "] </font>" & strSeg & " " & _
                      "<i>(" & strGls & ")</i> ")
        End If
        ' GO to next forest
        ndxThis = ndxThis.NextSibling : intI += 1
      End While
      ' Finish HTML output
      colBack.Add("</body></html>")
      ' Save the html output
      strText = colBack.Text
      IO.File.WriteAllText(strFile, strText, System.Text.Encoding.UTF8)
      ' Show the HTML output
      frmMain.wbReport.DocumentText = strText
      Status("File has been saved: " & strFile)
    Catch ex As Exception
      ' Warn user
      HandleErr("modViewer/ViewLineLoad error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Reset the initialisation
      bInit = False
    End Try
  End Sub
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  GetForestLoc
  ' Goal :  Get the TextId component of the forest [ndxThis]
  ' History:
  ' 02-07-2011  ERK Created 
  ' ----------------------------------------------------------------------------------------------------------
  Public Function GetForestLoc(ByRef ndxThis As XmlNode, ByRef strLoc As String) As Boolean
    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Validate
      If (ndxThis.Name <> "forest") Then Return False
      ' Get the forest location
      strLoc = ndxThis.Attributes("Location").Value
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn user
      HandleErr("modMain/GetForestLoc error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  GetSeg
  ' Goal :  Get the "original" text of the current sentence node
  '         Psdx:   Get the stuff that is between <seg>...</seg> 
  '         Negra:  Combine all the terminal nodes in the correct order
  ' History:
  ' 20-02-2010  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Public Function GetSeg(ByRef ndThis As Xml.XmlNode, ByVal strLang As String) As String
    Dim ndOrg As Xml.XmlNode    ' The div/seg node we are looking for
    Dim strBack As String = ""  ' What we return
    Dim strEndPunct = ".,;:!?"

    Try
      ' Does the node exist?
      If (ndThis Is Nothing) Then Return ""
      ' Do we have a <seg> section?
      ndOrg = ndThis.SelectSingleNode("./div[@lang='" & strLang & "']/seg")
      ' Validity
      If (ndOrg Is Nothing) Then Return ""
      ' Return the appropriate value
      ' We do, so return the [innerText] of it
      Return ndOrg.InnerText
    Catch ex As Exception
      ' Warn the user
      HandleErr("modParse/GetSeg error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return empty
      Return ""
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       LineToSection()
  ' Goal:       Given a line number, determine which section it is in
  ' History:
  ' 13-06-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function LineToSection(ByVal intLine As Integer) As Integer
    Dim intI As Integer   ' Counter

    Try
      ' Validate
      If (intLine < 0) Then Return 0
      If (pdxSection Is Nothing) Then Return 0
      If (pdxSection.Count = 0) Then Return 0
      ' Walk all sections
      For intI = 0 To pdxSection.Count - 1
        ' What is the forest node id starting this section?
        If (pdxSection(intI).Attributes("forestId").Value > intLine) Then
          ' Our line number must be in the previous section
          Return intI - 1
        End If
      Next intI
      ' Our line number must be in the last section
      Return pdxSection.Count - 1
    Catch ex As Exception
      ' Warn the user
      HandleErr("modParse/LineToSection error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return default
      Return 0
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       GotoSection()
  ' Goal:       Go to the indicated section -- but this does not actually LOAD the section!
  ' History:
  ' 12-07-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function GotoSection(ByVal intSection As Integer) As Boolean
    Try
      ' Does this section exist?
      If (intSection < 0) OrElse _
         (intSection > pdxSection.Count - 1) Then
        ' Return failure
        Return False
      End If
      ' Set the global variable to indicate which section we are at
      intCurrentSection = intSection
      ' Determine the first and last index of this section in [pdxList]
      intSectFirst = pdxSection(intSection).Attributes("forestId").Value - 1
      If (intSection = pdxSection.Count - 1) Then
        ' This is the last section
        intSectLast = pdxList.Count - 1
      Else
        ' This is an in-between section
        intSectLast = pdxSection(intSection + 1).Attributes("forestId").Value - 1
      End If
      ' Determine the number of <forest> elements in this section
      intSectSize = intSectLast - intSectFirst + 1
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/GotoSection error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       ShowSection()
  ' Goal:       Load and show the indicated section of the whole XML document
  ' History:
  ' 14-05-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function ShowSection(ByRef rtbThis As RichTextBox, ByVal intSection As Integer) As Boolean
    Dim strLoc As String          ' Location of this <forest> element
    Dim strSeg As String = ""     ' Text of this <forest> element
    Dim strLine As String         ' Text of line to be added
    Dim intI As Integer           ' A counter
    Dim intJ As Integer           ' Second counter
    Dim intPtc As Integer         ' Percentage
    Dim objText As New StringColl ' The text we read
    Dim ndxThis As XmlNode        ' One node
    Dim ndxCoref As XmlNodeList   ' List of nodes with coref information
    Dim ndxForest As XmlNode      ' This line's forest
    Dim bLeafy As Boolean         ' Whether this is a direct parent of an <eLeaf>

    Try
      ' Does this section exist?
      If (bSection) OrElse (rtbThis Is Nothing) Then
        ' Return failure
        Return False
      End If
      ' Indicate that we are busy
      bSection = True
      ' Show what section we are now in (if [pdxSection] exists)
      If (pdxSection Is Nothing) Then
        frmMain.SectionText("Section " & intSection + 1)
      Else
        frmMain.SectionText("Section " & intSection + 1 & "/" & pdxSection.Count)
      End If
      ' Try goto the indicated section
      If (Not GotoSection(intSection)) Then Return False
      ' Switch off changing of the richtextbox
      rtbThis.Visible = False
      ' Clear the file as it is right now
      objText.Clear()
      ' Make the correct size for the Left Line array
      ReDim arLeft(0 To intSectSize) : arLeft(0) = 0
      ' Walk the list
      For intI = 0 To intSectSize - 1
        ' ============= DEBUG =============
        If (intI = intSectSize) Then Stop
        ' =================================
        ' Show where we are
        intPtc = (100 * intI) \ intSectSize
        Status("Reading section " & intSection & " " & intPtc & "%", intPtc)
        ' Make sure we are interruptable
        Application.DoEvents()
        ' Get the correct forest node
        ndxForest = GetForest(intI)
        If (ndxForest Is Nothing) Then Return Nothing
        ' Access this element
        With ndxForest
          ' Check this one
          If (.Attributes("Location") Is Nothing) OrElse (.Attributes("Location").Value = "") Then
            ' Show that this line is kind of empty
            strLine = "(code)"
            ' Add a new Match element, if there are children (the first child will be an <eTree> element)
            If (.ChildNodes.Count > 0) Then
              ' Get the first childnode, which should be an <eTree> element
              ndxThis = .FirstChild
              ' Determine whether this is a direct parent of an <eLeaf>
              bLeafy = (ndxThis.FirstChild.Name = "eLeaf")
            End If
          Else
            ' Retrieve details of this element
            strLoc = .Attributes("Location").Value
            ' Get the forest node text
            strLine = ForestText(ndxForest)
          End If
          ' Add the line
          objText.Add(strLine)
          ' Keep track of the position
          If (intI < intSectSize) Then
            arLeft(intI + 1) = arLeft(intI) + 1 + Len(strLine)
          End If
        End With
      Next intI
      ' Set the font color correctly
      rtbThis.Clear()
      rtbThis.SelectionColor = Color.Black
      ' Copy file to the richtextbox
      rtbThis.Text = objText.Text
      ' Reset richtextbox stuff to REGULAR
      rtbThis.Font = New Font(rtbThis.Font, FontStyle.Regular)
      rtbThis.SelectionFont = rtbThis.Font
      ' Walk through it again
      For intI = 0 To intSectSize - 1
        ' Show where we are
        intPtc = (100 * intI) \ intSectSize
        Status("Processing coreference " & intPtc & "%", intPtc)
        ' Make sure we are interruptable
        Application.DoEvents()
        ' Get the correct forest node
        ndxForest = GetForest(intI)
        If (ndxForest Is Nothing) Then Return Nothing
        ' Get the list of nodes with coref information (only local descendants of this line!!)
        ndxCoref = ndxForest.SelectNodes("./descendant::fs[@type='coref']")
        ' Found anything?
        If (Not ndxCoref Is Nothing) AndAlso (ndxCoref.Count > 0) Then
          ' Walk through the results
          For intJ = 0 To ndxCoref.Count - 1
            ' Show this one element
            ShowOneCoref(rtbThis, ndxCoref(intJ).ParentNode)
          Next intJ
        End If
        ' Get a list of nodes with foreign words (FW ...)
        ndxCoref = ndxForest.SelectNodes("./descendant::eTree[@Label='FW']")
        ' Found anything?
        If (Not ndxCoref Is Nothing) AndAlso (ndxCoref.Count > 0) Then
          ' Walk through the results
          For intJ = 0 To ndxCoref.Count - 1
            ' Show this one element
            ShowOneEtree(rtbThis, ndxCoref(intJ), Color.DarkGray)
          Next intJ
        End If
        ' Get a list of nodes with META information (META ...)
        ndxCoref = ndxForest.SelectNodes("./descendant::eTree[@Label='META']")
        rtbThis.BackColor = Color.White
        ' Found anything?
        If (Not ndxCoref Is Nothing) AndAlso (ndxCoref.Count > 0) Then
          ' Walk through the results
          For intJ = 0 To ndxCoref.Count - 1
            ' Show this one element
            ShowOneEtree(rtbThis, ndxCoref(intJ), Color.White)
          Next intJ
        End If
        ' Determine whether the current line is a section start
        If (ndxForest.Attributes("Section") Is Nothing) Then
          ' Reset richtextbox stuff to REGULAR
          LineItalic(rtbThis, intI, False)
        Else
          ' This is a section start, so make it Italic
          LineItalic(rtbThis, intI, True)
        End If
      Next intI
      ' Set the position to the start of this section
      rtbThis.SelectionStart = 0 : rtbThis.SelectionLength = 0
      ' Show we are ready
      Status("Section has been read")
      ' Switch on RichTextboxLayout again
      rtbThis.Visible = True
      ' Indicate that we are nolonger busy
      bSection = False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/ShowSection error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Switch on RichTextboxLayout again
      rtbThis.Visible = True
      ' Return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       ForestText()
  ' Goal:       Return the text associated with this forest node
  ' History:
  ' 14-05-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function ForestText(ByRef ndxForest As XmlNode, Optional ByVal strLanguage As String = "org") As String
    Dim strSeg As String    ' Text of the segment
    Dim strLine As String   ' Line text to return
    Dim ndxThis As XmlNode  ' Child of the forest node

    Try
      ' Check validity
      If (ndxForest Is Nothing) Then Return ""
      ' Get the appropriate child
      ndxThis = ndxForest.SelectSingleNode("./div[@lang='" & strLanguage & "']/seg")
      If (Not ndxThis Is Nothing) Then
        strSeg = ndxThis.InnerText
        ' Show the <org/seg> element in the current rich text box
        strLine = strSeg
      Else
        strLine = "(empty line)"
      End If
      ' Return the value
      Return strLine
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/ForestText error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       GetForestNode()
  ' Goal:       Get the <forest> node above this node
  ' History:
  ' 28-05-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function GetForestNode(ByRef ndxThis As XmlNode, ByRef ndxForest As XmlNode) As Boolean
    Dim bFound As Boolean = False ' Whether we found the result
    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Start by setting the forest to myself
      ndxForest = ndxThis
      ' Get initial result
      bFound = (ndxForest.Name = "forest")
      ' Loop until we find it
      While (Not ndxForest Is Nothing) AndAlso (Not bFound)
        ' Go one step higher if possible
        ndxForest = ndxForest.ParentNode
        ' Have we now found it?
        bFound = (ndxForest.Name = "forest")
      End While
      ' Did we find it?
      Return bFound
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/GetForestNode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       GetOrgText()
  ' Goal:       Try get the original text of the <forest> node above this [ndxThis]
  ' History:
  ' 28-05-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function GetOrgText(ByRef ndxThis As XmlNode, ByRef strLoc As String) As String
    Dim ndxForest As XmlNode = Nothing ' The associated forest node

    Try
      ' Try get the forest
      If (GetForestNode(ndxThis, ndxForest)) Then
        ' Get the proper location
        strLoc = ndxForest.Attributes("TextId").Value & "," & ndxForest.Attributes("Location").Value
        ' Return the text
        Return ForestText(ndxForest)
      Else
        ' Empty location
        strLoc = ""
        ' Return failure
        Return ""
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/GetOrgText error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function

  '---------------------------------------------------------------------------------------------------------
  ' Name:       ShowOneCoref()
  ' Goal:       Show the coreference information of the indicated node
  ' History:
  ' 10-05-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function ShowOneCoref(ByRef rtfThis As RichTextBox, ByRef ndxEtree As XmlNode) As Boolean
    Dim strRefType As String  ' The type of coreference
    Dim intLine As Integer    ' The current line
    Dim intFrom As Integer    ' From
    Dim intTo As Integer      ' To
    Dim intSelStart As Integer ' Previous selection start
    Dim intSelLen As Integer  ' Previous selection length
    Dim ndxThis As XmlNode    ' One node

    Try
      ' Validate
      If (rtfThis Is Nothing) OrElse (ndxEtree Is Nothing) Then Return False
      ' Determine the coreference type
      strRefType = CorefAttr(ndxEtree, "RefType")
      ' Access this element
      With ndxEtree
        ' Determine the from and the to
        intFrom = .Attributes("from").Value
        intTo = .Attributes("to").Value
        ' Determine the current line
        ndxThis = .SelectSingleNode("./ancestor::forest[1]")
        intLine = ndxThis.Attributes("forestId").Value - intSectFirst
        ' Make sure we are not interrupted
        bColoring = True
        ' Set the correct selection with the correct colour
        With rtfThis
          ' Get previous selection
          intSelStart = .SelectionStart : intSelLen = .SelectionLength
          ' Set new selection + colour
          .SelectionStart = arLeft(intLine - 1) + intFrom - 1
          .SelectionLength = intTo - intFrom + 1
          .SelectionColor = GetRefColor(strRefType)
          ' Restore previous selection
          .SelectionStart = intSelStart
          .SelectionLength = intSelLen
        End With
        ' Restore interruption possibilities
        bColoring = False
      End With
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/ShowOneCoref error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       ShowOneEtree()
  ' Goal:       Show one <eTree> element in a particular color
  ' History:
  ' 10-05-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function ShowOneEtree(ByRef rtfThis As RichTextBox, ByRef ndxEtree As XmlNode, ByVal colThis As Color) As Boolean
    Dim intLine As Integer    ' The current line
    Dim intFrom As Integer    ' From
    Dim intTo As Integer      ' To
    Dim ndxThis As XmlNode    ' One node
    Dim fntReg As New Font(rtfThis.Font, FontStyle.Regular)
    Dim fntBold As New Font(rtfThis.Font, FontStyle.Bold)

    Try
      ' Validate
      If (rtfThis Is Nothing) OrElse (ndxEtree Is Nothing) Then Return False
      ' Access this element
      With ndxEtree
        ' Double check: if there is no 'from' and 'to', we cannot draw it; but we don't complain...
        If (.Attributes("from") Is Nothing OrElse .Attributes("to") Is Nothing) Then Return True
        ' Determine the from and the to
        intFrom = .Attributes("from").Value
        intTo = .Attributes("to").Value
        ' Determine the current line
        ndxThis = .SelectSingleNode("./ancestor::forest[1]")
        intLine = ndxThis.Attributes("forestId").Value - intSectFirst
        ' Set the correct selection with the correct colour
        With rtfThis
          .SelectionStart = arLeft(intLine - 1) + intFrom - 1
          .SelectionLength = intTo - intFrom + 1
          If (colThis = Color.White) Then
            .SelectionColor = Color.DarkBlue
            .SelectionBackColor = Color.LightYellow
          Else
            .SelectionColor = colThis
          End If
        End With
      End With
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/ShowOneEtree error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       GetRefColor()
  ' Goal:       Get the color belonging to this reference type
  ' History:
  ' 10-05-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function GetRefColor(ByVal strName As String) As Color
    Dim strCol As String        ' The name of the color
    Dim dtrFound() As DataRow   ' Result of select statement

    Try
      ' If the name is empty, the color is plain black:
      If (strName = "") Then
        strCol = "Black"
      Else
        ' Find this color
        dtrFound = tblRefType.Select("Name='" & strName & "'")
        ' Any results?
        If (dtrFound Is Nothing) OrElse (dtrFound.Length = 0) Then
          ' no result so supply the default ERROR colour
          strCol = "Purple"
        Else
          ' Determine the color name for this reference type
          strCol = dtrFound(0).Item("Color")
        End If
      End If
      ' Get the named color in RGB
      Return Color.FromName(strCol)
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/GetRefColor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return default colour: black
      Return Color.Black
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetPde
  ' Goal:   Get the PDE translation of the currently selected line
  ' History:
  ' 29-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function GetPde(ByVal intLine As Integer) As String
    Dim ndxForest As XmlNode  ' The forest node that is selected

    Try
      ' Validate
      If (intLine < 0) OrElse (intLine >= intSectSize) Then Return ""
      ' Get the forest node
      ndxForest = pdxList(intLine + intSectFirst)
      ' Get the translation
      Return ndxForest.SelectSingleNode("./div[@lang='eng']/seg").InnerText
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/GetPde error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return default colour: black
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   SetPde
  ' Goal:   Set the PDE translation of the currently selected line
  ' History:
  ' 29-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Sub SetPde(ByVal intLine As Integer, ByVal strPde As String)
    Dim ndxForest As XmlNode  ' The forest node that is selected

    Try
      ' Validate
      If (intLine < 0) OrElse (intLine >= intSectSize) Then Exit Sub
      ' Get the forest node
      ndxForest = pdxList(intLine + intSectFirst)
      ' Set the translation
      ndxForest.SelectSingleNode("./div[@lang='eng']/seg").InnerText = strPde
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/SetPde error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '' ------------------------------------------------------------------------------------------------------------
  '' Name:   GetTransPosition
  '' Goal:   Get the start and end position of the translated text where we are...
  '' History:
  '' 29-06-2010  ERK Created
  '' ------------------------------------------------------------------------------------------------------------
  'Public Function GetTransPosition(ByVal intPos As Integer, ByRef intTransStart As Integer, _
  '                                 ByRef intTransEnd As Integer) As Boolean
  '  Dim intPos2 As Integer    ' Position within string
  '  Dim intNum As Integer = 0 ' Number of lines
  '  Dim strSpace As String = " " & vbCrLf & vbTab

  '  Try
  '    ' Validate
  '    If (intPos < 0) OrElse (intPos > frmMain.tbTranslation.TextLength) Then Return False
  '    ' Count number of LF from the start
  '    intPos2 = InStr(frmMain.tbTranslation.Text, vbLf)
  '    ' Loop 
  '    While (intPos2 > 0) AndAlso (intPos2 < intPos)
  '      ' Increment number of LFs
  '      intNum += 1
  '      ' Save the previous one
  '      intTransStart = intPos2
  '      ' Go to next one
  '      intPos2 = InStr(intPos2 + 1, frmMain.tbTranslation.Text, vbLf)
  '    End While
  '    ' Transfer
  '    intPos2 = intTransStart
  '    ' If the number of LFs is even, then we have to go one LF further
  '    If (intNum Mod 2 = 0) Then
  '      ' Go to the next line
  '      intPos2 = InStr(intPos2 + 1, frmMain.tbTranslation.Text, vbLf)
  '    End If
  '    ' This could be the correct start
  '    intTransStart = intPos2
  '    ' We should skip spaces
  '    While (InStr(strSpace, Mid(frmMain.tbTranslation.Text, intPos2, 1)) > 0) And (intPos2 < frmMain.tbTranslation.TextLength)
  '      ' Keep track of start
  '      intTransStart = intPos2
  '      ' Try next position
  '      intPos2 += 1
  '    End While
  '    ' Find the next LF (if any)
  '    intPos2 = InStr(intPos2 + 1, frmMain.tbTranslation.Text, vbLf)
  '    ' Did we find one?
  '    If (intPos2 = 0) Then
  '      ' We are at the end
  '      intPos2 = frmMain.tbTranslation.TextLength
  '    Else
  '      intPos2 -= 1
  '    End If
  '    ' Put the position in [intTransEnd]
  '    intTransEnd = intPos2
  '    Return True
  '  Catch ex As Exception
  '    ' Show error
  '    HandleErr("modMain/GetTransPosition error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
  '    ' Return failre
  '    Return False
  '  End Try
  'End Function
  '' ------------------------------------------------------------------------------------------------------------
  '' Name:   GetTransLine
  '' Goal:   Get the line number within the translated text where we are...
  '' History:
  '' 29-06-2010  ERK Created
  '' ------------------------------------------------------------------------------------------------------------
  'Public Function GetTransLine(ByVal intPos As Integer, ByRef intLineNum As Integer) As Boolean
  '  Dim intPos2 As Integer    ' Position within string
  '  Dim intNum As Integer = 0 ' Number of lines

  '  Try
  '    ' Validate
  '    If (intPos < 0) OrElse (intPos > frmMain.tbTranslation.TextLength) Then Return False
  '    ' Count number of LF from the start
  '    intPos2 = InStr(frmMain.tbTranslation.Text, vbLf)
  '    ' Loop 
  '    While (intPos2 > 0) AndAlso (intPos2 < intPos)
  '      ' Increment number of LFs
  '      intNum += 1
  '      ' Go to next one
  '      intPos2 = InStr(intPos2 + 1, frmMain.tbTranslation.Text, vbLf)
  '    End While
  '    ' Return the number of linefeeds found
  '    intLineNum = intNum
  '    Return True
  '  Catch ex As Exception
  '    ' Show error
  '    HandleErr("modMain/GetTransLine error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
  '    ' Return failre
  '    Return False
  '  End Try
  'End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetRtfLineNumber
  ' Goal:   Determine the line number
  ' History:
  ' 01-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetRtfLineNumber(ByRef rtfThis As RichTextBox, ByVal intPos As Integer, _
                                   ByRef intLineNum As Integer) As Boolean
    Dim intMin As Integer     ' Minimum
    Dim intMax As Integer     ' Maximum

    Try
      ' Validate
      If (intPos < 0) Then Return False
      ' The following should not happen, but apparently it does..
      If (arLeft Is Nothing) Then Return False
      ' Special case: intpos=1
      If (intPos = 0) Then
        intLineNum = 0
        Return True
      End If
      ' Determine where we are with respect to [arLeft()]
      intMin = 0 : intMax = UBound(arLeft)
      intLineNum = GetSectionLineNum(intPos, intMin, intMax)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/GetRtfPosition error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetRtfPosition
  ' Goal:   Determine the line number and the character position
  ' History:
  ' 26-04-2010  ERK Created
  ' 18-05-2010  ERK Adapted to work with [arLeft()]
  ' ------------------------------------------------------------------------------------
  Public Function GetRtfPosition(ByVal intPos As Integer, ByRef intLineNum As Integer, _
      ByRef intCharNum As Integer, ByRef intMyLeft As Integer) As Boolean
    Dim intMin As Integer     ' Minimum
    Dim intMax As Integer     ' Maximum

    Try
      ' Are we starting from a correct position?
      If (intPos < 0) Then Return False
      ' Special case: intpos = 1
      If (intPos = 0) Then
        ' We are completely at the beginning of the file!!
        intLineNum = 0
        intCharNum = 0
        Return True
      End If
      ' Determine where we are with respect to [arLeft()]
      intMin = 0 : intMax = UBound(arLeft)
      intLineNum = GetSectionLineNum(intPos, intMin, intMax)
      ' Save this first position to the left
      intMyLeft = arLeft(intLineNum)
      ' Determine the character number
      intCharNum = intPos - intMyLeft
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/GetRtfPosition error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  IsStarred
  ' Goal :  Check if this node has a "Starred" <eLeaf>
  ' History:
  ' 29-11-2010  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function IsStarred(ByRef ndxThis As XmlNode) As Boolean
    Dim ndxWork As XmlNode  ' Working node

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Visit all my children
      ndxWork = ndxThis.FirstChild
      While (Not ndxWork Is Nothing)
        ' Check name of child
        If (ndxWork.Name = "eLeaf") Then
          ' Is this of type "Star"?
          Return (ndxWork.Attributes("Type").Value = "Star")
        End If
        ' Go to next one
        ndxWork = ndxWork.NextSibling
      End While
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/IsStarred error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  LeafText
  ' Goal :  Get the text of the first <eLeaf> child of the node ndxThis if it exists
  '         If this child does not exist, then return an empty string
  ' History:
  ' 21-05-2010  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function LeafText(ByRef ndxThis As XmlNode) As String
    Dim ndxLeaf As XmlNode  ' The first <eLeaf> child

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return ""
      ' Get the child (which may be of ANY <eLeaf> type!!!)
      ndxLeaf = ndxThis.SelectSingleNode("./child::eLeaf[1]")
      ' Did we get a child?
      If (ndxLeaf Is Nothing) Then
        ' Return empty
        Return ""
      Else
        ' Return the text of the child
        Return ndxLeaf.Attributes("Text").Value
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/LeafText error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  LabelType
  ' Goal :  Get the type of label, possibly depending on children
  ' History:
  ' 05-01-2009  ERK Created for Cesac
  ' 29-04-2010  ERK Adapted for Cesax
  ' ----------------------------------------------------------------------------------------------------------
  Public Function LabelType(ByRef ndxThis As XmlNode) As String
    Dim strLabel As String    ' The <Label> value of the constituend numbered [intId]
    Dim strLeaf As String     ' The lexeme of the selected constituent
    Dim intI As Integer       ' Counter in datarow set [dtrPhrType]
    Dim ndxChild As XmlNode   ' First (legitimate) child of me
    Dim bFound As Boolean     ' Flag

    Try
      ' Check validity
      If (ndxThis Is Nothing) Then Return ""
      ' Get label value
      strLabel = ndxThis.Attributes("Label").Value
      ' GO through all possible label matchings
      For intI = 0 To dtrPhrType.Length - 1
        ' Access the information of this label
        With dtrPhrType(intI)
          ' See if this match works
          If (LabelMatch(.Item("Node"), strLabel)) Then
            ' Initially we are okay
            bFound = True
            ' Should the parent also match?
            If (.Item("Child") <> "") Then
              ' Get the first <eTree> child that is not a CODE one
              ndxChild = ndxThis.SelectSingleNode("./eTree[not(@Label='CODE')]")
              ' Check if this is the correct first child - CODE children should be skipped
              If (ndxChild Is Nothing) Then
                ' There is no correct child -- check if the LEXEME (in the <eLeaf>) meets the specs
                strLeaf = LeafText(ndxThis)
                ' Was an appropriate hild found?
                If (strLeaf = "") Then
                  ' The leaf is empty, so this is not a good match
                  bFound = False
                Else
                  ' There was a child - check it
                  If (Not LabelMatch(.Item("Child"), strLeaf)) Then
                    ' If a child should match, and it does not, this is not valid
                    bFound = False
                  End If
                End If
              Else
                ' Try to match the FIRST child with the label or else with the node
                If (Not LabelMatch(.Item("Child"), ndxChild.Attributes("Label").Value)) Then
                  ' Match is not complete, so try next arLabel element
                  bFound = False
                End If
              End If
            End If
            ' Are we okay?
            If (bFound) Then
              ' No parent match is needed
              Select Case .Item("Type")
                Case "Must"
                  ' The labeltype is "Must"
                  Return "Must"
                Case "Can"
                  ' Label type depends on Target type
                  Select Case .Item("Target")
                    Case "Any"
                      Return "MayAny"
                    Case "Dst"
                      Return "MayDst"
                  End Select
              End Select
            End If
          End If
        End With
      Next intI
      ' Return failure
      Return ""
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/LabelType error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  LabelMatch
  ' Goal :  See if the label [strLabel] matches with the input string [strIn]
  ' Note :  It is possible to have a wildcard Left and/or Right 
  '           and also somewhere in the Middle
  ' History:
  ' 05-01-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function LabelMatch(ByVal strLabel As String, ByVal strIn As String, _
      Optional ByVal bTakeMainPart As Boolean = True) As Boolean
    Dim intLen As Integer     ' Store the length
    Dim intPos As Integer     ' Position of wildcard somewhere within the string

    Try
      ' Is a wildcard attached to the label?
      If (Right(strLabel, 1) = "*") Then
        ' What follows after the label doesn't matter
        strLabel = Left(strLabel, Len(strLabel) - 1)
        intLen = Len(strLabel)
      Else
        ' Complete matching
        intLen = Len(strLabel)
        ' If [strLabel] does not contain any LabelDivision like -, +, ^, then take the main part of [strIn]
        If (InStr(strLabel, LABEL_DIV) = 0) AndAlso (bTakeMainPart) Then
          ' Take the main part of [strIn]
          strIn = Labelmain(strIn)
        End If
        ' But the maximum length should count for complete matching!
        If (Len(strIn) > intLen) Then intLen = Len(strIn)
      End If
      ' Should the first character also be skipped?
      If (Left(strLabel, 1) = "*") Then
        ' Adjust the label
        strLabel = Mid(strLabel, 2)
        ' Do partial matching anywhere within [strIn]
        LabelMatch = (InStr(strIn, strLabel) > 0)
      ElseIf (InStr(strLabel, "*") > 0) Then
        ' There is a wildcard somewhere within the label, but for the remainder it is left-aligned
        intPos = InStr(strLabel, "*")
        Return ((Left(strLabel, intPos - 1) = Left(strIn, intPos - 1)) And _
                      (InStr(intPos + 1, strIn, Mid(strLabel, intPos + 1)) > 0))
      Else
        ' Do left-aligned matching
        Return (Left(strLabel, intLen) = Left(strIn, intLen))
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/LabelMatch error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  LabelmainOE
  ' Goal :  Get the main part of the label, the part before a LabelDivider like + ^ -
  '         Include things like -GEN, -MAT etc
  ' History:
  ' 18-02-2013  ERK Derived from LabelMain()
  ' ----------------------------------------------------------------------------------------------------------
  Public Function LabelmainOE(ByVal strIn As String, Optional ByVal strDiv As String = LABEL_DIV_OE) As String
    'Dim strDiv As String = LABEL_DIV  ' Label divider possibilities
    Dim intPos As Integer = 0         ' Position of divider
    Dim intI As Integer               ' Counter

    Try
      ' First try the [strLabelDivOE] values
      For intI = 0 To strLabelDivOE.Length - 1
        ' Is this one present within [strIn]?
        intPos = InStr(strIn, strLabelDivOE(intI))
        If (intPos > 0) Then
          ' Strip it off from the end
          Return Left(strIn, intPos + Len(strLabelDivOE(intI)) - 1)
        End If
      Next intI
      ' Start from the left of [strIn]
      intPos = 1
      While (InStr(strDiv, Mid(strIn, intPos, 1)) = 0) AndAlso (intPos <= Len(strIn))
        ' Try next position
        intPos += 1
      End While
      ' Return what is left
      Return Left(strIn, intPos - 1)
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/LabelMainOE error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return strIn
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  LabelLemmaOE
  ' Goal :  Try to get the POS of the lemma belonging to this word
  ' History:
  ' 18-02-2013  ERK Derived from LabelMain()
  ' ----------------------------------------------------------------------------------------------------------
  Public Function LabelLemmaOE(ByVal strIn As String) As String
    Try
      ' First apply LabelMainOE
      strIn = LabelmainOE(strIn)
      ' Examine the result
      Select Case strIn
        Case "NS"
          Return "N"
        Case "ADJ", "ADV", "C", "CONJ", "N", "NR", "NUM", "P", "PRO", "PRO$", "D", "Q", "RP", "WADV", "WADJ", "WPRO", "INTJ"
          Return strIn
      End Select
      ' Some more comparisons
      If (strIn Like "V*") Then Return "VB"
      If (strIn Like "DO*") Then Return "DO"
      If (strIn Like "B*") Then Return "BE"
      If (strIn Like "ADJ*") Then Return "ADJ"
      If (strIn Like "ADV*") Then Return "ADV"
      If (strIn Like "ADV*") Then Return "ADV"
      If (strIn Like "HV*") Then Return "HV"
      If (strIn Like "NEG*") Then Return "NEG"
      If (strIn Like "Q*") Then Return "Q"
      If (strIn Like "RP*") Then Return "RP"
      If (strIn Like "*+ADV") Then Return "ADV"
      If (strIn Like "P21") Then Return "P"
      If (strIn Like "PP") Then Return "P"
      If (strIn Like "UTP") Then Return "RP"
      If (strIn Like "WQ") Then Return "WPRO"
      If (strIn Like "WADV+*") Then Return "WADV"
      If (strIn Like "X*") Then Return "X"
      If (strIn Like "MAN*") Then Return "MAN"
      If (strIn Like "NP*") Then Return "N"
      Stop
      ' Return what is left
      Return strIn
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/LabelMainOE error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return strIn
    End Try
  End Function

  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Labelmain
  ' Goal :  Get the main part of the label, the part before a LabelDivider like + ^ -
  ' History:
  ' 29-06-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function Labelmain(ByVal strIn As String) As String
    Dim strDiv As String = LABEL_DIV  ' Label divider possibilities
    Dim intPos As Integer = 0         ' Position of divider

    Try
      ' Start from the left of [strIn]
      intPos = 1
      While (InStr(strDiv, Mid(strIn, intPos, 1)) = 0) AndAlso (intPos <= Len(strIn))
        ' Try next position
        intPos += 1
      End While
      ' Return what is left
      Return Left(strIn, intPos - 1)
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/LabelMain error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return strIn
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DoTransRead
  ' Goal:   Read the translation information from [pdxThis], and 
  '           store it in the existing [pdxCurrentFile]
  ' History:
  ' 29-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function DoTransRead(ByRef pdxThis As XmlDocument) As Boolean
    Dim pdxSrc As XmlNodeList   ' List of source elements
    Dim pdxDst As XmlNodeList   ' List of destination elements
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (pdxThis Is Nothing) Then Return False
      ' Select all the <div lang='eng'> elements in the source and target
      ' WAS: pdxSrc = pdxThis.SelectNodes("//div[@lang='eng']")
      pdxSrc = pdxThis.SelectNodes("//seg[parent::div[@lang='eng']]")
      pdxDst = pdxCurrentFile.SelectNodes("//seg[parent::div[@lang='eng']]")
      ' Check their numbers
      If (pdxSrc.Count <> pdxDst.Count) Then
        ' Show user 
        MsgBox("modMain/DoTransRead: cannot read the file, since the numbers don't match the currently loaded file")
        Return False
      End If
      ' Loop through them all
      For intI = 0 To pdxSrc.Count - 1
        ' Double checking
        If (pdxSrc(intI).InnerText <> "") AndAlso (pdxSrc(intI).InnerText <> pdxDst(intI).InnerText) Then
          ' Set the value of the destination
          pdxDst(intI).InnerText = pdxSrc(intI).InnerText
        End If
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Show message to user
      HandleErr("modMain/DoTransSave error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       GetFirstTrans()
  ' Goal:       Get the first line that needs translation
  ' History:
  ' 25-10-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function GetFirstTrans() As Integer
    Dim colOrg As New StringColl  ' Original text
    Dim colTrn As New StringColl  ' Translation
    Dim strOrg As String = ""     ' Original
    Dim strTrn As String = ""     ' Translation
    Dim strText As String = ""    ' Text to be translated
    Dim strLoc As String = ""     ' Location of this forest element
    Dim ndxThis As XmlNode        ' This forest node
    Dim intI As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage

    Try
      ' Validate - do we have read anything?
      If (pdxList Is Nothing) Then Return ""
      ' Go through the whole list
      For intI = 0 To pdxList.Count - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ pdxList.Count
        Status("Translation part 1 (" & intPtc & "%)", intPtc)
        ' Get this forest node
        ndxThis = pdxList(intI)
        ' Check if this needs translation
        If (ndxThis.SelectSingleNode("./div[@lang='eng']").InnerText = "") Then
          ' This needs translation
          Return intI
        End If
      Next intI
      ' Return failure
      Return -1
    Catch ex As Exception
      ' Show message to user
      HandleErr("modMain/GetFirstTrans error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       GetTextTrans()
  ' Goal:       Get the text and its translation, centering around [intLine]
  ' Note:       The [bSection] indicates that we should only show the current section
  ' History:
  ' 03-03-2011  ERK Created
  ' 25-03-2011  ERK Experimentally added listview
  '---------------------------------------------------------------------------------------------------------
  Public Function GetTextTrans(ByVal intLine As Integer, ByVal bSection As Boolean, _
                               ByRef strPrevOrg As String, _
        ByRef strPrevTrn As String, ByRef strThisOrg As String, ByRef strThisTrn As String, _
        ByRef strNextOrg As String, ByRef strNextTrn As String) As String
    Dim colOrg As New StringColl  ' Original text
    Dim colTrn As New StringColl  ' Translation
    Dim strBare As String = ""    ' Bare translation, without location
    Dim strOrg As String = ""     ' Original
    Dim strTrn As String = ""     ' Translation
    Dim strText As String = ""    ' Text to be translated
    Dim strLoc As String = ""     ' Location of this forest element
    Dim strBack As String = ""    ' What we return
    Dim ndxThis As XmlNode        ' This forest node
    Dim intI As Integer           ' Counter
    Dim intStart As Integer       ' Starting point
    Dim intEnd As Integer         ' Ending point
    Dim intPtc As Integer         ' Percentage
    'Dim lvItem As ListViewItem    ' Current item
    Dim arOne(0 To 2) As String   ' List view items

    Try
      ' Validate - do we have read anything?
      If (pdxList Is Nothing) Then Return ""
      ' Determine the start and end
      If (bSection) Then
        ' Just this section
        intStart = intSectFirst : intEnd = intSectLast
      Else
        ' The whole file
        intStart = 0 : intEnd = pdxList.Count - 1
      End If
      '' ========== ListView =====
      '' Initialise listview
      'lvThis.Items.Clear()
      '' =========================
      ' Go through the whole list
      For intI = intStart To intEnd
        ' Show where we are
        intPtc = (intI + 1 - intStart) * 100 \ (intEnd - intStart + 1)
        Status("Reading text (" & intPtc & "%)", intPtc)
        ' Get this forest node
        ndxThis = pdxList(intI)
        ' Get the location
        strLoc = ndxThis.Attributes("Location").Value
        ' Get the Original and the English
        strOrg = ndxThis.SelectSingleNode("./div[@lang='org']").InnerText
        strBare = ndxThis.SelectSingleNode("./div[@lang='eng']").InnerText
        '' ========== ListView =====
        '' Store items in listview
        'lvItem = lvThis.Items.Add("line_" & CStr(intI))
        'arOne(0) = strLoc : arOne(1) = strOrg : arOne(2) = strBare
        'lvItem.SubItems.AddRange(arOne)
        '' =========================
        strOrg = strLoc & ": " & strOrg
        strTrn = strLoc & ": " & strBare
        ' Are we before or at+after the [intLine]?
        If (intI < intLine) Then
          ' We are before the line: Add the original
          colOrg.Add(strOrg)
          ' Add the translation
          colTrn.Add(strTrn)
        ElseIf (intI = intLine) Then
          ' We are AT the line
          intLastTransLine = intLine
          ' This is what we return
          strBack = strBare
          ' Store these last ones
          strThisOrg = strOrg
          strThisTrn = strTrn
          ' Prepare [prevOrg] and [prevTrn]
          strPrevOrg = colOrg.Text
          strPrevTrn = colTrn.Text
          ' Reset [colOrg] and [colTrn] so that they can contain what comes from now on
          colOrg.Clear()
          colTrn.Clear()
        Else
          ' We are after the line: Add the original
          colOrg.Add(strOrg)
          ' Add the translation
          colTrn.Add(strTrn)
        End If
      Next intI
      '' ========== ListView =====
      '' Set the correct selection
      'lvThis.Items(intLine - intStart).Selected = True
      '' =========================
      ' Return the Original and the Translation completely
      strNextOrg = colOrg.Text
      strNextTrn = colTrn.Text
      ' Show we are ready
      Status("Translation ready")
      ' Return the result
      Return strBack
    Catch ex As Exception
      ' Show message to user
      HandleErr("modMain/GetTextTrans error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       GetTransHtml()
  ' Goal:       Get an HTML translation of the complete text
  ' Note:       The [bSection] indicates that we should only show the current section
  '             The [intLine] optionally says which line to highlight
  ' History:
  ' 04-03-2011  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function GetTransHtml(ByVal bSection As Boolean, Optional ByVal intLine As Integer = -1) As String
    Dim colHtml As New StringColl ' Where we gather the result
    Dim strBare As String = ""    ' Bare translation, without location
    Dim strOrg As String = ""     ' Original
    Dim strTrn As String = ""     ' Translation
    Dim strLoc As String = ""     ' Location of this forest element
    Dim ndxThis As XmlNode        ' This forest node
    Dim intI As Integer           ' Counter
    Dim intStart As Integer       ' Starting point
    Dim intEnd As Integer         ' Ending point
    Dim intPtc As Integer         ' Percentage

    Try
      ' Validate - do we have read anything?
      If (pdxList Is Nothing) Then Return ""
      ' Determine the start and end
      If (bSection) Then
        ' Just this section
        intStart = intSectFirst : intEnd = intSectLast
      Else
        ' The whole file
        intStart = 0 : intEnd = pdxList.Count - 1
      End If
      ' Start the html
      colHtml.Add("<html><body><p><h2>" & IO.Path.GetFileNameWithoutExtension(strCurrentFile) & "</h2><p><table>")
      ' Go through the whole list
      For intI = intStart To intEnd
        ' Show where we are
        intPtc = (intI + 1 - intStart) * 100 \ (intEnd - intStart + 1)
        Status("Reading text (" & intPtc & "%)", intPtc)
        ' Get this forest node
        ndxThis = pdxList(intI)
        ' Get the location
        strLoc = ndxThis.Attributes("Location").Value
        ' Get the Original and the English
        strOrg = ndxThis.SelectSingleNode("./div[@lang='org']").InnerText
        strTrn = ndxThis.SelectSingleNode("./div[@lang='eng']").InnerText
        ' Put this in HTML
        If (intLine = intI) Then
          colHtml.Add("<tr><td  valign=top>" & strLoc & _
                      "</td><td  valign=top><font color='red'>" & strOrg & "</font>" & _
                      "</td><td  valign=top><font color='blue'>" & strTrn & "</font></td></tr>")
        Else
          colHtml.Add("<tr><td  valign=top>" & strLoc & "</td><td  valign=top>" & strOrg & _
                      "</td><td  valign=top>" & strTrn & "</td></tr>")
        End If
      Next intI
      ' Finish the html
      colHtml.Add("</table></body></html>")
      ' Show the translation is also available in HTML
      Status("The translation also is available in the [Report] tabpage")
      ' Return the result
      Return colHtml.Text
    Catch ex As Exception
      ' Show message to user
      HandleErr("modMain/GetTransHtml error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       AddLastTrans()
  ' Goal:       Add the [strText] as the translation of the last line 
  ' History:
  ' 25-10-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function AddLastTrans(ByVal strText As String) As Boolean
    Dim ndxNext As XmlNode  ' Working node

    Try
      ' Validate
      If (pdxList Is Nothing) Then Return False
      If (intLastTransLine >= pdxList.Count) Then Return False
      ' Get to the last translation line
      ndxNext = pdxList(intLastTransLine)
      If (Not ndxNext Is Nothing) Then
        ' Go through the children
        ndxNext = ndxNext.FirstChild
        While (Not ndxNext Is Nothing)
          ' Action depends on the kind of child
          If (ndxNext.Name = "div") Then
            ' Check which it is
            Select Case ndxNext.Attributes("lang").Value
              Case "eng"
                ' Set the English translation
                ndxNext.FirstChild.InnerText = strText
                ' Return success
                Return True
            End Select
          End If
          ' Go to next child
          ndxNext = ndxNext.NextSibling
        End While
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show message to user
      HandleErr("modMain/GetFirstTrans error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DoTransSave
  ' Goal:   Perform the actual saving of the translation part of a file
  ' History:
  ' 29-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function DoTransSave(ByVal strFile As String) As Boolean
    Dim pdxTransFile As New XmlDocument ' This is where we copy the translation to
    Dim pdxMyList As XmlNodeList        ' List of elements
    Dim ndxChild As XmlNode             ' One child
    Dim intI As Integer                 ' Counter
    Dim intJ As Integer                 ' Second counter

    Try
      ' Show status
      Status("Copying the translation file...")
      ' Copy the translation part to the [pdxTransFile]
      pdxTransFile = pdxCurrentFile.Clone
      Status("Saving the translation file...")
      ' Select all the <forest> elements
      pdxMyList = pdxTransFile.SelectNodes("//forest")
      ' Walk this list
      For intI = pdxMyList.Count - 1 To 0 Step -1
        ' Delete all the <eTree> children
        For intJ = pdxMyList(intI).ChildNodes.Count - 1 To 0 Step -1
          ' Get the child
          ndxChild = pdxMyList(intI).ChildNodes(intJ)
          ' Is this an <eTree> child?
          If (ndxChild.Name = "eTree") Then
            ' Delete this child
            pdxMyList(intI).RemoveChild(ndxChild)
          End If
        Next intJ
      Next intI
      ' Save the CRP information to the selected filename as XML
      pdxTransFile.Save(strFile)
      ' Show status
      Status("Translation file has been saved: " & strFile & " (" & Now & ")")
      ' Remove what I had made
      pdxTransFile = Nothing
      ' Return success
      Return True
    Catch ex As Exception
      ' Show message to user
      HandleErr("modMain/DoTransSave error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MyTrim
  ' Goal:   Trim leading and trailing spaces, CR, LF, tab etc.
  ' History:
  ' 25-10-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MyTrim(ByVal strIn As String) As String
    Dim strSkip As String = vbCrLf & vbTab & " "  ' What to skip
    Dim intPos As Integer = 1       ' Position in string

    Try
      ' Skip CRLF in the beginning
      While (InStr(strSkip, Mid(strIn, intPos, 1)) > 0) AndAlso (intPos < strIn.Length)
        intPos += 1
      End While
      ' Cut off from here
      strIn = Mid(strIn, intPos)
      ' Look backwards
      intPos = strIn.Length
      While (intPos > 1) AndAlso (InStr(strSkip, Mid(strIn, intPos - 1, 1)) > 0) AndAlso (intPos <= strIn.Length)
        intPos -= 1
      End While
      ' Cut off
      strIn = Left(strIn, intPos)
      ' Return actual trim
      Return (Trim(strIn))
    Catch ex As Exception
      ' Show message to user
      HandleErr("modMain/MyTrim error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return strIn
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   TrySaveResDb
  ' Goal:   Process the change of the feature value
  ' History:
  ' 21-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function TrySaveResDb(ByVal strFile As String) As Boolean
    Try
      ' Validate
      If (tdlResults Is Nothing) Then Return False
      ' Save dataset to the location
      ' tdlResults.WriteXml(strFile)
      pdxCrpDbase.Save(strFile)
      ' Show we saved it
      Status("Saved results database " & strResultDb & " at: " & Format(Now, "g"))
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/TrySaveResDb error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetResFieldValues
  ' Goal:   Get the periods of the results
  ' History:
  ' 04-01-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetResFieldValues(ByRef dtrFound() As DataRow, ByVal strField As String, _
                                    Optional ByVal bSort As Boolean = False) As String()
    Dim intI As Integer       ' Counter
    Dim strPer As String = "" ' CSV string
    Dim arBack() As String    ' Return array

    Try
      ' Validate
      If (dtrFound Is Nothing) Then Return Nothing
      If (dtrFound.Length = 0) Then Return Nothing
      ' Go through all the rows
      For intI = 0 To dtrFound.Length - 1
        ' Add the period from this row
        AddSemiStack(strPer, dtrFound(intI).Item(strField).ToString.Replace(";", "_"), True)
      Next intI
      ' Get the array
      arBack = Split(strPer, ";")
      ' Optionally sort the array
      If (bSort) Then Array.Sort(arBack)
      ' Replace _ again
      For intI = 0 To UBound(arBack)
        arBack(intI) = arBack(intI).Replace("_", ";")
      Next intI
      ' Return the result
      Return arBack
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/GetResFieldValues error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetPerVal
  ' Goal:   Get the position of [strPeriod] within [arPeriod]
  ' History:
  ' 04-01-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetPerVal(ByRef arPeriod() As String, ByVal strPeriod As String) As Integer
    Dim intI As Integer       ' Counter

    Try
      ' Validate
      If (arPeriod Is Nothing) Then Return -1
      If (strPeriod = "") Then Return -1
      ' Find the place
      For intI = 0 To UBound(arPeriod)
        If (arPeriod(intI) = strPeriod) Then Return intI
      Next intI
      ' Return failure
      Return -1
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/GetPerVal error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuCorpusReorder_Click
  ' Goal:   Load the corpus database, re-order it, and save it again
  ' Note:   This re-orders the corpus based on the order in which the XML elements are available
  ' History:
  ' 23-05-2011  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function CorpusReOrder(ByVal strFile As String) As Boolean
    Dim pdxDoc As XmlDocument = Nothing ' Where we load to
    Dim ndxThis As XmlNode              ' One node
    Dim intResId As Integer = 1         ' ResId variable

    Try
      ' Validate
      If (strFile = "") Then Return False
      If (Not IO.File.Exists(strFile)) Then Return False
      ' Load it as xml document
      If (Not ReadXmlDoc(strFile, pdxDoc)) Then Return False
      ' Walk through all <Result> records
      ndxThis = pdxDoc.SelectSingleNode("//Result")
      While (Not ndxThis Is Nothing)
        ' Set the ID
        ndxThis.Attributes("ResId").Value = intResId
        intResId += 1
        ' Go to next one
        ndxThis = ndxThis.NextSibling
      End While
      ' Save result again
      pdxDoc.Save(strFile)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/CorpusReOrder error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuCorpusOrderBy_Click
  ' Goal:   Load the corpus database, re-order it, and save it again
  ' Note:   This re-orders the corpus based on:
  '         (1) Period @until, @from
  '         (2) Name of the file
  ' History:
  ' 23-05-2011  ERK Created
  ' 27-09-2013  ERK This needs re-working, because we now switched to an XmlDocument
  ' ------------------------------------------------------------------------------------------------------------
  Public Function CorpusOrderBy() As Boolean
    'Dim dtrFound() As DataRow     ' Result of re-ordering
    Dim intResId As Integer = 1   ' ResId variable
    Dim ndxThis As XmlNode        ' One node
    'Dim ndxWork As XmlNode        ' Working node
    'Dim intI As Integer           ' Counter
    'Dim intPtc As Integer         ' Percentage

    Try
      ' Validate
      If (tdlResults Is Nothing) OrElse (tdlResults.Tables("Result") Is Nothing) Then Return False
      If (tdlResults.Tables("Result").Rows.Count = 0) Then Return False
      '' Determine the period number for all elements
      'With tdlResults.Tables("Result")
      '  For intI = 0 To .Rows.Count - 1
      '    With .Rows(intI)
      '      .Item("PeriodNo") = GetPeriodId(.Item("Period"))
      '    End With
      '  Next intI
      'End With
      ' Determine the period number for all elements
      ndxThis = pdxCrpDbase.SelectSingleNode("./descendant::Result")
      While (ndxThis IsNot Nothing)
        AddXmlAttribute(pdxCrpDbase, ndxThis, "PeriodNo", GetPeriodId(GetAttrValue(ndxThis, "Period")))
        ndxThis = ndxThis.NextSibling
      End While
      '' Reorder correctly
      'dtrFound = tdlResults.Tables("Result").Select("", "PeriodNo ASC, TextId ASC, forestId ASC")
      '' Give new ResId numbers
      'For intI = 0 To dtrFound.Length - 1
      '  ' Show where we are
      '  intPtc = (intI + 1) * 100 \ dtrFound.Length
      '  Status("Revising ResId numbers " & intPtc & "%", intPtc)
      '  '' Get the XmlDocument entry with the indicated
      '  'dtrFound(intI).Item("ResId") = intResId : intResId += 1
      'Next intI
      ' Save to temporary file and re-order
      Status("Sorting XmlDocument...")
      If (Not SortResultsXml(pdxCrpDbase)) Then Return False
      ' Walk through the results and assign new ResId numbers
      ndxThis = pdxCrpDbase.SelectSingleNode("./descendant::Result") : intResId = 1
      While (ndxThis IsNot Nothing)
        ndxThis.Attributes("ResId").Value = intResId : intResId += 1
        ndxThis = ndxThis.NextSibling
      End While
      ' Return success
      Status("Ready")
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/CorpusOrderBy error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   SortResultsXml
  ' Goal:   Sort the Results database XmlDocument in-line
  ' History:
  ' 27-09-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function SortResultsXml(ByRef pdxIn As XmlDocument) As Boolean
    Try
      If (Not SortResultsXml(pdxIn.DocumentElement)) Then Return False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/SortResultsXml error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  Private Function SortResultsXml(ByRef ndRoot As XmlNode) As Boolean
    Dim ndChild As XmlNode

    Try
      ' Sort this level
      If (Not SortResultElements(ndRoot)) Then Return False
      ' Sort my children
      For Each ndChild In ndRoot.ChildNodes
        If (Not SortResultsXml(ndChild)) Then Return False
      Next ndChild
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/SortResultsXml error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   SortResultElements
  ' Goal:   Sort the Results database XmlDocument in-line, making use of Bubble Sort
  '         Comparison order: "PeriodNo ASC, TextId ASC, forestId ASC"
  ' History:
  ' 27-09-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function SortResultElements(ByRef ndThis As XmlNode) As Boolean
    Dim intI As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage
    Dim intCount As Integer       ' Number of changes
    Dim intChanged As Integer     ' Number of previous changes
    Dim ndxCurr As XmlNode        ' Current node
    Dim ndxPrev As XmlNode        ' Previous
    Dim ndxWork As XmlNode        ' Working node
    Dim ndxLast As XmlNode        ' Last one
    Dim changed As Boolean = True ' Flag

    Try
      ' Loop until ready
      While changed
        changed = False : intChanged = intCount : intCount = 0
        For intI = 1 To ndThis.ChildNodes.Count - 1
          ' Show where we are
          intPtc = intI * 100 \ ndThis.ChildNodes.Count
          Status("Sorting after [" & intCount & "] changes " & intPtc & "%", intPtc)
          If (bInterrupt) Then Return False
          ' Get current and previous
          ndxCurr = ndThis.ChildNodes(intI) : ndxPrev = ndThis.ChildNodes(intI - 1)
          ' Is this a <Result> node?
          If (ndxCurr.Name = "Result") AndAlso (ndxPrev.Name = "Result") Then
            ' Compare the attributes PeriodNo, TextId and forestId
            If (SortResultCompare(ndxCurr, ndxPrev)) Then
              ' Find the first node before which I have to go
              ndxWork = ndxPrev.PreviousSibling : ndxLast = ndxWork
              While (ndxWork IsNot Nothing) AndAlso (SortResultCompare(ndxCurr, ndxWork))
                ndxLast = ndxWork : ndxWork = ndxWork.PreviousSibling
              End While
              ' Move myself ahead
              ndThis.InsertBefore(ndxCurr, ndxLast)
              changed = True : intCount += 1
            End If
          End If
        Next
      End While
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/SortResultElements error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   SortResultCompare
  ' Goal:   Compare the two items, current and previous
  '         Comparison order: "PeriodNo ASC, TextId ASC, forestId ASC"
  ' History:
  ' 27-09-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function SortResultCompare(ByRef ndxCurr As XmlNode, ByRef ndxPrev As XmlNode) As Boolean
    Dim bNeed As Boolean = False  ' Needs adaptation

    Try
      If (CInt(GetAttrValue(ndxCurr, "PeriodNo")) < CInt(GetAttrValue(ndxPrev, "PeriodNo"))) Then
        bNeed = True
      ElseIf (CInt(GetAttrValue(ndxCurr, "PeriodNo")) = CInt(GetAttrValue(ndxPrev, "PeriodNo"))) Then
        ' The periods are equal -- check for text id
        If (GetAttrValue(ndxCurr, "TextId") < GetAttrValue(ndxPrev, "TextId")) Then
          bNeed = True
        ElseIf (GetAttrValue(ndxCurr, "TextId") = GetAttrValue(ndxPrev, "TextId")) Then
          If (CInt(GetAttrValue(ndxCurr, "forestId")) < CInt(GetAttrValue(ndxPrev, "forestId"))) Then
            bNeed = True
          End If
        End If
      End If
      ' Return the need
      Return bNeed
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/SortResultCompare error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuCorpusAddFile
  ' Goal:   Add the indicated file to the already existing results
  ' Notes:  Only consider those with status "Done"
  '         Change these statuses into "DoneImp"
  ' History:
  ' 10-10-2011  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function CorpusAddFile(ByVal strFile As String) As Boolean
    Dim pdxDoc As XmlDocument = Nothing ' Where we load to
    Dim ndxThis As XmlNode              ' One node
    Dim dtrFound() As DataRow           ' Result of SELECT
    Dim intResId As Integer = 1         ' ResId variable
    Dim intPtc As Integer               ' Progress
    Dim intAll As Integer               ' Total number of results
    Dim intI As Integer                 ' Counter

    Try
      ' Validate
      If (strFile = "") Then Return False
      If (tdlResults Is Nothing) OrElse (tdlResults.Tables("Result").Rows.Count = 0) Then Return False
      If (Not IO.File.Exists(strFile)) Then Return False
      ' Load it as xml document
      If (Not ReadXmlDoc(strFile, pdxDoc)) Then Return False
      ' Walk through all <Result> records
      ndxThis = pdxDoc.SelectSingleNode("//Result")
      intAll = ndxThis.ParentNode.ChildNodes.Count
      intI = 1
      While (Not ndxThis Is Nothing)
        ' Show where we are
        intPtc = (intI * 100) \ intAll : intI += 1
        Status("Adding [" & ndxThis.Attributes("Period").Value & "-" & ndxThis.Attributes("TextId").Value _
               & "] " & intPtc & "%", intPtc)
        ' See if this line needs to be added
        If (ndxThis.Attributes("Status") Is Nothing) OrElse _
           (ndxThis.Attributes("Status").Value <> "Remove") Then
          ' Yes, this line needs to be added
          If (Not CorpusAddLine(ndxThis)) Then Return False
        End If
        ' Go to next one
        ndxThis = ndxThis.NextSibling
      End While
      ' Do the ResId numbers again
      dtrFound = tdlResults.Tables("Result").Select("", "ResId ASC")
      For intI = 0 To dtrFound.Length - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ dtrFound.Length
        Status("Renumbering " & intPtc & "%", intPtc)
        ' Set this ResId
        dtrFound(intI).Item("ResId") = intI + 1
      Next intI
      ' Return success
      Status("Done")
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/CorpusAddFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   CorpusAddLine
  ' Goal:   Add the contents of this one line to the existing Results database
  '           on the most appropriate location!!
  '         Then readjust the ResultId numbers that follow
  ' History:
  ' 10-10-2011  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function CorpusAddLine(ByRef ndxThis As XmlNode) As Boolean
    Dim strPeriod As String         ' Period to be added
    Dim intPeriod As Integer        ' The numerical number for this period
    Dim strName As String           ' Name of the file that needs adding
    Dim strFtName As String         ' Name of feature
    Dim strFtVal As String          ' Feature value
    Dim intForestId As Integer      ' The ForestId of the new addition
    Dim intResId As Integer = 1     ' ResId variable
    Dim intI As Integer             ' Counter
    Dim intJ As Integer             ' Counter
    Dim intResPer As Integer        ' Period of current result id
    Dim bFound As Boolean = False   ' Flag when we have found and processed our target
    Dim dtrNew As DataRow = Nothing ' New [Result] datarow
    Dim dtrFt As DataRow = Nothing  ' New [Feature] datarow
    Dim dtrParent As DataRow = Nothing  ' The parent CrpOview row
    Dim ndxFeat As XmlNode          ' Feature node
    Dim arAttr() As String = {"OviewId", "File", "Search", "TextId", "Cat", "forestId", "eTreeId", "Period", _
                              "Notes", "Select", "Status"}
    Dim arInner() As String = {"Text", "Psd"}

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      If (tdlResults Is Nothing) OrElse (tdlResults.Tables("Result").Rows.Count = 0) Then Return False
      ' Get the period, forestId and the filename of the new addition
      strPeriod = ndxThis.Attributes("Period").Value
      intPeriod = GetPeriodId(strPeriod) : If (intPeriod < 0) Then Status("Cannot find period") : Return False
      strName = ndxThis.Attributes("TextId").Value
      intForestId = ndxThis.Attributes("forestId").Value
      ' ====================== DEBUG =============
      'If (strName = "cmrolltr") Then Stop
      ' ==========================================
      ' Go through the database from the start!!
      With tdlResults.Tables("Result")
        ' There should at least be one row
        If (.Rows.Count = 0) Then
          ' Set parent row from the table
          dtrParent = tdlResults.Tables("CrpOview").Rows(0)
        Else
          ' Get the parent row
          dtrParent = .Rows(0).GetParentRow("CrpOview_Result")
        End If
        For intI = 0 To .Rows.Count - 1
          ' Access this element
          With .Rows(intI)
            ' Check for the period
            intResPer = GetPeriodId(.Item("Period"))
            If (intResPer >= intPeriod) Then
              ' Is this the exact period?
              If (intResPer = intPeriod) Then
                ' The period is exact, so now look for the correct location
                If (.Item("TextId") > strName) Then
                  ' Exit the for loop with this IntI in mind
                  Exit For
                End If
              Else
                ' The period is not exact, so exit the for-loop
                Exit For
              End If
            End If
          End With
        Next intI
        ' Check if the row is too large
        If (intI >= .Rows.Count) Then
          ' We need to add a new row
        End If
        '' We have come here with a correct intI, I hope!
        'With .Rows(intI)
        'End With
        ' Our Item should be inserted somewhere here...
        If (Not CreateNewRow(tdlResults, "Result", "ResId", intResId, dtrNew)) Then Return False
        ' ============== DEBUG ===============
        ' Debug.Print(tdlResults.Tables("Result").Rows(intI).Item("ResId"))
        ' ====================================
        ' Set the parent of this row
        dtrNew.SetParentRow(dtrParent)
        ' Copy the attributes
        For intJ = 0 To UBound(arAttr)
          If (ndxThis.Attributes(arAttr(intJ)) Is Nothing) Then
            dtrNew.Item(arAttr(intJ)) = ""
          Else
            dtrNew.Item(arAttr(intJ)) = ndxThis.Attributes(arAttr(intJ)).Value
          End If
        Next intJ
        ' Copy the inner text
        For intJ = 0 To UBound(arInner)
          dtrNew.Item(arInner(intJ)) = ndxThis.SelectSingleNode("./" & arInner(intJ)).InnerText
        Next intJ
        ' Set the status to imported
        If (dtrNew.Item("Status").ToString = "Done") Then
          ' Change the status from "Done" into "DoneImported" (DoneImp)
          dtrNew.Item("Status") = "DoneImp"
        End If
        ' Set the correct resultId if we can
        If (intI < .Rows.Count) Then
          ' Yes, we can set the result id
          intResId = .Rows(intI).Item("ResId") - 1
        Else
          ' No, take the biggest number we can
          intResId = .Rows.Count
        End If
        dtrNew.Item("ResId") = intResId
        ' Produce and make and copy all features
        ndxFeat = ndxThis.FirstChild
        While (Not ndxFeat Is Nothing)
          ' Is this the correct child?
          If (ndxFeat.Name = "Feature") Then
            ' Also check on the VALUE of this feature
            If (ndxFeat.Attributes("Value") Is Nothing) Then
              strFtVal = ""
            Else
              strFtVal = ndxFeat.Attributes("Value").Value
            End If
            ' The name should be there
            strFtName = ndxFeat.Attributes("Name").Value
            ' Copy feature name and value
            '   (NOTE: [intResId] functions as dummy -- it is not actually used)
            If (Not CreateNewRow(tdlResults, "Feature", "", intResId, dtrFt, _
                                 strFtName, strFtVal)) Then Return False
            ' Set the parent table
            dtrFt.SetParentRow(dtrNew)
          End If
          ' Go to next feature
          ndxFeat = ndxFeat.NextSibling
        End While
        ' Make sure results are processed
        tdlResults.AcceptChanges()
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/CorpusAddLine error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetPeriodId
  ' Goal:   Get an ID value of the indicated period (for comparing the order of periods)
  ' Note:   This assumes that the periods have been sorted correctly!!
  ' History:
  ' 10-10-2011  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function GetPeriodId(ByVal strPeriod As String) As Integer
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim intI As Integer       ' Counter

    Try
      ' Validate
      If (tdlPeriods Is Nothing) OrElse (tdlPeriods.Tables("Period").Rows.Count = 0) Then Return False
      ' Position the periods in the correct order
      dtrFound = tdlPeriods.Tables("Period").Select("", loc_strPeriodOrder)
      ' Now find the correct period
      For intI = 0 To dtrFound.Length - 1
        ' Check this period
        If (dtrFound(intI).Item("Name") = strPeriod) Then
          ' Return it
          Return intI + 1
        End If
      Next intI
      ' Return failure
      Return -1
      '' Get to the correct row
      'dtrFound = tdlPeriods.Tables("Period").Select("Name='" & strPeriod & "'")
      '' Get result
      'If (dtrFound.Length = 0) Then
      '  Return -1
      'Else
      '  Return dtrFound(0).Item("PeriodId")
      'End If
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/GetPeriodId error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   RefListStat
  ' Goal:   Perform statistical measures on the reference list stored in [tdlThis]
  ' History:
  ' 17-08-2011  ERK Created and statistics derived from earlier RefListHTML
  ' 09-09-2011  ERK Added a measure for "subject-switch" 
  ' ------------------------------------------------------------------------------------
  Public Function RefListStat(ByRef tdlThis As DataSet, ByVal strFile As String) As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim strName As String     ' Name of statistical operation
    Dim intI As Integer       ' counter
    Dim intNoTr As Integer    ' Total amount of items that are NOT a trace
    Dim intLen As Integer     ' Length indicator

    Try
      ' Validate
      If (tdlThis Is Nothing) OrElse (strFile = "") Then Return False
      If (tdlThis.Tables("ChainList").Rows.Count = 0) Then Return False
      If (tdlThis.Tables("Chain").Rows.Count = 0) Then Return False
      ' If there are any previous rows, then delete them
      If (Not tdlThis.Tables("Stat") Is Nothing) Then
        dtrFound = tdlThis.Tables("Stat").Select("")
        For intI = dtrFound.Length - 1 To 0 Step -1
          ' Delete this row
          dtrFound(intI).Delete()
        Next intI
        ' Make sure changes are accepted
        tdlThis.AcceptChanges()
      End If
      ' ========================================
      ' Chain length numbers WITH traces
      strName = "Length (with Trace)"
      Status("Chain length...")
      ' Order chains according to their length
      dtrFound = tdlThis.Tables("Chain").Select("", "Len ASC") : intLen = 0
      For intI = 0 To dtrFound.Length - 1
        ' Check for interrupt
        If (bInterrupt) Then Return False
        ' What is the current length?
        If (dtrFound(intI).Item("Len") > intLen) Then
          ' Get this new length
          intLen = dtrFound(intI).Item("Len")
          ' Store the number of occurrances for this new length
          If (Not AddStatRow(tdlThis, strFile, 0, intLen, intLen, strName, _
                  tdlThis.Tables("Chain").Select("Len=" & intLen).Length, dtrFound.Length, "One")) Then Return False
        End If
      Next intI
      ' ========================================
      ' Chain length numbers WITHOUT traces
      strName = "Length (no Trace)"
      ' Order chains according to their length
      dtrFound = tdlThis.Tables("Chain").Select("NoTraceLen>0", "NoTraceLen ASC") : intLen = 0
      For intI = 0 To dtrFound.Length - 1
        ' Check for interrupt
        If (bInterrupt) Then Return False
        ' What is the current length?
        If (dtrFound(intI).Item("NoTraceLen") > intLen) Then
          ' Get this new length
          intLen = dtrFound(intI).Item("NoTraceLen")
          ' Store the number of occurrances for this new length
          If (Not AddStatRow(tdlThis, strFile, 0, intLen, intLen, strName, _
                  tdlThis.Tables("Chain").Select("NoTraceLen=" & intLen).Length, dtrFound.Length, "One")) Then Return False
        End If
      Next intI
      ' ========================================
      ' Chain length distribution (intervals) WITH traces
      Status("Chain length distribution...")
      If (Not ChainStatFreq(tdlThis, strFile, 1, "Distribution (with Trace)", True)) Then Return False
      ' ========================================
      ' Chain length distribution (intervals) WITHOUT traces
      If (Not ChainStatFreq(tdlThis, strFile, 2, "Distribution (no Trace)", False)) Then Return False
      ' ========================================
      ' Subject percentages WITH traces
      Status("Subject percentages...")
      If (Not ChainStatSbj(tdlThis, strFile, 3, "Subject (with Trace)", True)) Then Return False
      ' ========================================
      ' Subject percentages WITHOUT traces
      If (Not ChainStatSbj(tdlThis, strFile, 4, "Subject (no Trace)", False)) Then Return False
      ' ========================================
      ' Pronoun percentages WITHOUT traces
      If (Not ChainStatPro(tdlThis, strFile, 5, "Pronoun (no Trace)")) Then Return False
      ' ========================================
      ' Pronoun antecedent distance distribution, maximum and average
      Status("Pronoun antecedent distribution...")
      If (Not ChainStatProDist(tdlThis, strFile, 6, "ProDist3s-", "IPnum")) Then Return False
      ' ========================================
      ' Pronoun antecedent distance distribution, maximum and average
      If (Not ChainStatProDist(tdlThis, strFile, 7, "ProDist3s-", "Forest")) Then Return False
      ' ========================================
      ' Several different global counts
      ' (1) First calculate the total amount of chain items that are NOT a trace
      intNoTr = tdlThis.Tables("Item").Select("NPtype <> 'Trace'").Length
      ' (2) Ellipsis: get the total amount of chain items that are stored as ZeroSbj
      Status("Ellipsis...")
      intI = tdlThis.Tables("Item").Select("NPtype = 'ZeroSbj'").Length
      If (Not AddStatRow(tdlThis, strFile, 0, 0, 0, "Ellipsis", intI, intNoTr, "Global")) Then Return False
      ' (3) Non-empty subjects: Get the total amount of subject chain items that are not traces and not ZeroSbj
      Status("Subject chains...")
      intI = tdlThis.Tables("Item").Select("NPtype <> 'ZeroSbj' AND NPtype <> 'Trace' AND GrRole = 'Subject'").Length
      If (Not AddStatRow(tdlThis, strFile, 0, 0, 0, "Subject (not empty)", intI, intNoTr, "Global")) Then Return False
      If (Not AddStatRow(tdlThis, strFile, 0, 0, 0, "SubjectChainSwitch", intI, intNoTr, "Global")) Then Return False
      ' (4) Improved Subject Switch
      '     Get the total number of Referent(sbj(line[i])) <> Referent(sbj(line[i+1]))
      '   (the relevant total number is calculated and returned in [intNoTr])
      Status("Subject referent switch...")
      If (Not DoSbjRefSwitch(tdlThis, intI, "ipcls", intNoTr)) Then Return False
      If (Not AddStatRow(tdlThis, strFile, 0, 0, 0, "SubjectRefSwitch[ipcls]", intI, intNoTr, "Global")) Then Return False
      ' The same, but now for the <IP-MAT> domain
      If (Not DoSbjRefSwitch(tdlThis, intI, "ipmat", intNoTr)) Then Return False
      If (Not AddStatRow(tdlThis, strFile, 0, 0, 0, "SubjectRefSwitch[ipmat]", intI, intNoTr, "Global")) Then Return False
      ' The same, but now for the <IP-MAT> domain and only 3rd person referents
      If (Not DoSbjRefSwitch(tdlThis, intI, "ipmat3", intNoTr)) Then Return False
      If (Not AddStatRow(tdlThis, strFile, 0, 0, 0, "SubjectRefSwitch[ipmat3]", intI, intNoTr, "Global")) Then Return False
      ' The same, but now for the <forest> domain
      If (Not DoSbjRefSwitch(tdlThis, intI, "forest", intNoTr)) Then Return False
      If (Not AddStatRow(tdlThis, strFile, 0, 0, 0, "SubjectRefSwitch[forest]", intI, intNoTr, "Global")) Then Return False
      ' The same, but now for the <chunk> domain
      If (Not DoSbjRefSwitch(tdlThis, intI, "chunk", intNoTr)) Then Return False
      If (Not AddStatRow(tdlThis, strFile, 0, 0, 0, "SubjectRefSwitch[chunk]", intI, intNoTr, "Global")) Then Return False
      ' (4b) Subject referent counting
      Status("Subject referent counting...")
      '   For the <IP-MAT> domain
      If (Not DoSbjRefSwitch(tdlThis, intI, "CountIpMat", intNoTr)) Then Return False
      If (Not AddStatRow(tdlThis, strFile, 0, 0, 0, "SubjectRefCount[ipmat]", intI, intNoTr, "Global")) Then Return False
      '   For the <IP-MAT> domain and only 3rd person referents
      If (Not DoSbjRefSwitch(tdlThis, intI, "CountIpMat3", intNoTr)) Then Return False
      If (Not AddStatRow(tdlThis, strFile, 0, 0, 0, "SubjectRefCount[ipmat3]", intI, intNoTr, "Global")) Then Return False
      ' (4c) Number of different referents that occur as subject
      If (Not DoSbjRefSwitch(tdlThis, intI, "CountIpAll", intNoTr)) Then Return False
      If (Not AddStatRow(tdlThis, strFile, 0, 0, 0, "SubjectRefCount[all]", intI, intNoTr, "Global")) Then Return False
      ' (5) Subject-switch (not really true)
      '     Get the total amount of chain items that switch FROM subject, where the subject is NOT a trace
      If (Not DoChainSbjSwitch(tdlThis, intI)) Then Return False
      ' (6) Protagonist Subject Ratio (= #time subject=protagonist[x] / #subjects during protagonist[x])
      Status("Protagonist subject ratio...")
      If (Not DoProtSbjRatio(tdlThis, intI)) Then Return False
      Status("Statistics ready")
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/RefListStat error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddStatRow
  ' Goal:   Add one row with statistical information that was derived
  ' History:
  ' 17-08-2011  ERK Created and statistics derived from earlier RefListHTML
  ' ------------------------------------------------------------------------------------
  Private Function AddStatRow(ByRef tdlThis As DataSet, ByVal strFile As String, ByVal intColumn As Integer, _
          ByVal intFrom As Integer, ByVal intUntil As Integer, ByVal strName As String, ByVal intCount As Integer, _
          ByVal intTotal As Integer, ByVal strType As String) As Boolean
    Dim dtrParent As DataRow = Nothing  ' The parent [StatList] row
    Dim dtrStat As DataRow = Nothing    ' One datarow with statistical counts
    Dim intId As Integer                ' ID

    Try
      ' Validate
      If (tdlThis Is Nothing) Then Return False
      If (strFile = "") OrElse (strName = "") Then Return False
      ' Get the parent
      If (tdlThis.Tables("StatList") Is Nothing) OrElse (tdlThis.Tables("StatList").Rows.Count = 0) Then
        ' Start up a Stat List row in the dataset
        If (Not CreateNewRow(tdlThis, "StatList", "", intId, dtrParent)) Then Return False
      Else
        ' Get the parent row
        dtrParent = tdlThis.Tables("StatList").Rows(0)
      End If

      ' Store the XML information
      If (Not CreateNewRow(tdlThis, "Stat", "StatId", intId, dtrStat)) Then Return False
      With dtrStat
        .SetParentRow(dtrParent)
        .Item("StatId") = intId
        .Item("File") = IO.Path.GetFileNameWithoutExtension(strFile)
        .Item("Name") = strName
        .Item("From") = intFrom
        .Item("Until") = intUntil
        .Item("Count") = intCount
        .Item("Total") = intTotal
        .Item("Type") = strType
        .Item("Column") = intColumn
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/RefListStat error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   RefListXML
  ' Goal:   Make an XML overview of ALL coreferential chains in this document
  ' History:
  ' 10-08-2011  ERK Created
  ' 16-08-2011  ERK Chains of length "1" (New/Assumed) should also be added
  ' ------------------------------------------------------------------------------------
  Public Function RefListXML(ByRef tdlThis As DataSet, ByVal strFile As String) As Boolean
    Dim ndxThis As XmlNode            ' Current node
    Dim colDone As New NodeColl       ' Collection of nodes we have treated already
    Dim dtrThis As DataRow = Nothing  ' General purpose datarow
    Dim intPtc As Integer             ' Where we are
    Dim intI As Integer               ' counter
    Dim intJ As Integer               ' counter
    Dim intChainId As Integer         ' ID of this chain
    Dim intForNum As Integer          ' Number of forest nodes
    Dim dtrChain() As DataRow         ' Selection of <chain> elements
    Dim dtrItem() As DataRow          ' Selection of <item> elements
    Dim strPGN As String              ' The PGN value of a whole chain
    Dim strThis As String             ' PGN of this item
    Dim strPer As String = ""         ' Period
    Dim colBack As New StringColl     ' Messages about chains with the same chainroot
    Dim arSyntax() As String          ' Where we store [Syntax] fields

    Try
      ' Validate: check if a text has been loaded
      If (pdxCurrentFile Is Nothing) Then Return False
      ' Initialise current file
      If (Not InitCurrentFile()) Then Return False
      ' Add general information to the dataset
      SetTableSetting(tdlThis, "File", strFile)
      SetTableSetting(tdlThis, "Date", Format(Now, "g"))
      SetTableSetting(tdlThis, "LogBase", intLogBase)
      ' Try get period
      If (HasPeriod(strFile, strPer)) Then
        SetTableSetting(tdlThis, "Period", strPer)
      Else
        SetTableSetting(tdlThis, "Period", "(unknown)")
      End If
      ' Start up a Chain List row in the dataset
      If (Not CreateNewRow(tdlThis, "ChainList", "", intI, dtrThis)) Then Return False
      ' Other initialisations
      intChainNum = 0             ' Start with chain number zero
      intForNum = pdxList.Count   ' The number of <forest> nodes
      bNeedSaving = False         ' Initially we do not need to save here
      ' Walk backwards through the forest node list
      For intI = intForNum - 1 To 0 Step -1
        ' Check for interrupt
        If (bInterrupt) Then Return False
        ' Show where we are in terms of number of forests
        intPtc = (intForNum - intI) * 100 \ intForNum
        Status("Processing chainlist [" & IO.Path.GetFileNameWithoutExtension(strFile) & "] " & intPtc & "%", intPtc)
        ' Get the forest node
        ndxThis = pdxList(intI)
        ' Perform action recursively backwards
        If (Not TravEtreeBack(ndxThis, colDone, tdlThis)) Then Return False
      Next intI
      ' Has anything changed in this file?
      If (bNeedSaving) Then
        ' Save the PSDX file, which may have been updated with @coref/@ChainId features
        pdxCurrentFile.Save(strFile)
      End If
      ' Find all the Chains with NoTraceLen larger than 1
      dtrChain = tdlThis.Tables("Chain").Select("NoTraceLen>1")
      ' Walk all these chains
      For intI = 0 To dtrChain.Length - 1
        ' Get the chain id
        intChainId = dtrChain(intI).Item("ChainId")
        ' Get all the items on this chain
        dtrItem = tdlThis.Tables("Item").Select("ChainId=" & intChainId, "ItemId ASC")
        ' Walk all the items, except for the last one
        For intJ = 0 To dtrItem.Length - 2
          dtrItem(intJ).Item("DistIPnum") = dtrItem(intJ).Item("IPnum") - dtrItem(intJ + 1).Item("IPnum")
          dtrItem(intJ).Item("DistForest") = dtrItem(intJ).Item("forestId") - dtrItem(intJ + 1).Item("forestId")
        Next intJ
      Next intI
      ' Start double chain error list
      colBack.Add("<h1>Overlapping chains</h1><table>" & _
                  "<tr><td>Higher chain</td><td>Len</td><td>Overlap</td><td>Len</td><td>Root</td></tr>")
      ' Find ALL the chains to determine their PGN
      dtrChain = tdlThis.Tables("Chain").Select("")
      ' Make room for the syntax
      ReDim arSyntax(0 To dtrChain.Length - 1)
      ' Check if the current root is already present in a previous item
      Status("Checking for overlapping chains...")
      ' Walk all these chains
      For intI = 0 To dtrChain.Length - 1
        '' Initialise the animacy
        'If (dtrChain(intI).Item("Animacy").ToString = "") Then dtrChain(intI).Item("Animacy") = "-"
        'dtrChain(intI).Item("Animacy") = "(unknown)"
        ' Get the chain id
        intChainId = dtrChain(intI).Item("ChainId")
        ' Get all the items on this chain
        dtrItem = tdlThis.Tables("Item").Select("ChainId=" & intChainId, "ItemId ASC")
        ' Initialise the PGN value
        strPGN = ""
        ' Store the first item's syntax
        arSyntax(intI) = dtrItem(0).Item("Syntax")
        ' Walk all the items, except for the last one
        For intJ = 0 To dtrItem.Length - 1
          ' Get this item's PGN
          strThis = dtrItem(intJ).Item("PGN")
          ' See if we can improve on the PGN
          If (Not DoLike(strPGN, "1s|1p|3p|3fs|3ms|3ns|2s|2p")) Then
            ' The PGN could become MORE specific, so ...
            ' ... check if the new item's PGN is more specific than mine
            If (MoreSpecific(strThis, strPGN)) Then strPGN = strThis
          End If
          ' Need animacy?
          If (dtrChain(intI).Item("Animacy") = "-") Then
            ' See if we can set the animacy
            If (DoLike(strThis, "3fs|3ms|1p|1s")) Then
              ' it must be animate
              dtrChain(intI).Item("Animacy") = "a"
            ElseIf (DoLike(strThis, "3ns")) Then
              ' It must be inanimate
              dtrChain(intI).Item("Animacy") = "i"
            End If
          End If
        Next intJ
        ' Set the PGN of this chain
        dtrChain(intI).Item("PGN") = strPGN
      Next intI
      ' Place the list with double chains somewhere
      colBack.Add("</table>")
      frmMain.wbError.DocumentText = colBack.Text
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/RefListXML error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AnimacyPerculate
  ' Goal:   Let animacy perculate!
  ' History:
  ' 16-04-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AnimacyPerculate(ByVal intChainId As Integer, ByVal strAnimValue As String) As Boolean
    Dim ndxList As XmlNodeList  ' Result of SELECT
    Dim strValue As String      ' Current value
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (intChainId <= 0) Then Return False
      If (Not DoLike(strAnimValue, "-|a|i")) Then Return False
      ' Find all the chain elements with this id
      ndxList = pdxCurrentFile.SelectNodes("//eTree[child::fs/child::f[@name='ChainId' and @value='" & intChainId & "']]")
      ' Process them
      For intI = 0 To ndxList.Count - 1
        ' Check if we need processing
        strValue = GetFeature(ndxList(intI), "NP", "anim")
        If (strValue = "") OrElse ((strValue = "-") AndAlso (strAnimValue <> "-")) OrElse (strValue <> strAnimValue) Then
          ' Add this value
          If (Not AddFeature(pdxCurrentFile, ndxList(intI), "NP", "anim", strAnimValue)) Then Return False
          ' Indicate that saving is needed
          bNeedSaving = True
        End If
      Next intI
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/AnimacyPerculate error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CheckOverlap
  ' Goal:   Check for overlap between chains
  ' History:
  ' 09-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function CheckOverlap(ByVal bAuto As Boolean) As Boolean
    Dim dtrChain() As DataRow     ' Selection of <chain> elements
    Dim dtrItem() As DataRow      ' Selection of <item> elements
    Dim dtrHigh() As DataRow      ' Selection of <item> elements
    Dim ndxHigh As XmlNode        ' First element on the "higher" chain
    Dim ndxSrc As XmlNode         ' Source node
    Dim ndxDst As XmlNode         ' Destination node
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intK As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage progress
    Dim intChainId As Integer     ' Chain's id
    Dim intRootEtreeId As Integer ' Current chain's root id
    Dim intTargetId As Integer    ' eTreeId of the target
    Dim intSourceId As Integer    ' eTreeId of the source
    Dim strRefType As String      ' Reference type
    Dim strSrcIp As String
    Dim strDstIp As String

    Try
      ' Validate
      If (tdlRefChain Is Nothing) Then Return False
      ' Select ALL the chains in increasing order
      dtrChain = tdlRefChain.Tables("Chain").Select("", "ChainId ASC")
      ' Walk all the chains
      For intI = 0 To dtrChain.Length - 1
        ' Show progress
        intPtc = (intI + 1) * 100 \ dtrChain.Length
        Status("Overlap " & intPtc, intPtc)
        ' Get values from this chain
        intChainId = dtrChain(intI).Item("ChainId")
        intRootEtreeId = dtrChain(intI).Item("RootEtreeId")
        ' Initialise
        intSourceId = -1
        ' Get the first node of this chain
        dtrItem = tdlRefChain.Tables("Item").Select("ChainId=" & intChainId, "eTreeId ASC")
        ' Do we have something?
        If (dtrItem.Length > 0) Then
          ' Double check
          If (Not dtrItem(0).Item("Syntax") Like "*PRN*") Then
            ' Look within all "higher" chains
            For intJ = intI + 1 To dtrChain.Length - 1
              ' We already know that the ChainId value is higher (due to the sorting of [dtrChain])
              ' Check identity of the root id
              If (intRootEtreeId = dtrChain(intJ).Item("RootEtreeId")) Then
                ' Get the first node on this chain
                dtrHigh = tdlRefChain.Tables("Item").Select("ChainId=" & dtrChain(intJ).Item("ChainId"), "ItemId ASC")
                ' Check the @Syntax field
                If (dtrHigh.Length > 0) AndAlso (Not dtrHigh(0).Item("Syntax") Like "*PRN*") Then
                  ' Get the first node of this "higher" chain -- this is the target we suggest
                  intTargetId = dtrHigh(0).Item("eTreeId")
                  ndxHigh = IdToNode(intTargetId)
                  ' Get the first node within [dtrItem] that has a HIGHER eTreeId than [ndxHigh] has
                  For intK = 0 To dtrItem.Length - 1
                    If (dtrItem(intK).Item("eTreeId") > intTargetId) Then
                      intSourceId = dtrItem(intK).Item("eTreeId")
                      Exit For
                    End If
                  Next intK
                  ' Validate
                  If (intSourceId >= 0) Then
                    ' Are we in automatic mode?
                    If (bAuto) Then
                      ' Get source and destination
                      ndxSrc = IdToNode(intSourceId)
                      ndxDst = IdToNode(intTargetId)
                      strSrcIp = GetIpAncestor(ndxSrc).Attributes("Label").Value
                      strDstIp = GetIpAncestor(ndxDst).Attributes("Label").Value
                      ' Determine reference type
                      If ((InStr(strSrcIp, "SPE") > 0) AndAlso (InStr(strDstIp, "SPE") = 0)) OrElse _
                         ((InStr(strSrcIp, "SPE") = 0) AndAlso (InStr(strDstIp, "SPE") > 0)) Then
                        strRefType = strRefCross
                      Else
                        strRefType = strRefIdentity
                      End If
                      If (CorefFromTo(ndxSrc, ndxDst, strRefType)) Then
                        ' Copy features from antecedent to me
                        If (Not CopyFeaturesNode(ndxSrc, ndxDst, strRefType)) Then
                          ' Well, it is a pity that we cannot copy the features, but not a deadly sin!
                        End If
                        ' Copy the ChainId feature from source to antecedent and further down
                        If (Not CopyChainId(ndxSrc, ndxDst, strRefType)) Then
                          ' It's a pity if this doesn't work -- user will have to recalculate
                        End If
                        ' Add the history
                        If (Not AddHistory(ndxSrc, "coref", "OverlapRepair")) Then Return False
                        ' Show we have changed
                        frmMain.SetDirty(True)
                      Else
                        ' Show error information on error page
                        frmMain.wbError.DocumentText = "Error<p>Source: " & NodeInfo(ndxSrc) & "<p>Target: " & NodeInfo(ndxDst)
                        ' There is a problem
                        Return False
                      End If
                    Else
                      ' Suggest this link to the user
                      SelectNode(IdToNode(intSourceId), "source")
                      SelectNode(IdToNode(intTargetId), "target")
                      ' Show question to use
                      Status("Resolve overlap here, then restart overlap checking")
                      Return True
                    End If
                  End If
                End If
              End If
            Next intJ
          End If
        End If
      Next intI
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/CheckOverlap error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   OverlapRepair
  ' Goal:   Check for overlap between chains and automatically repair it
  ' History:
  ' 10-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function OverlapRepair(ByVal bAuto As Boolean) As Boolean
    'Dim dtrChain() As DataRow     ' Selection of <chain> elements
    'Dim dtrItem() As DataRow      ' Selection of <item> elements
    'Dim dtrHigh() As DataRow      ' Selection of <item> elements
    'Dim ndxThis As XmlNode        ' One element
    'Dim ndxHigh As XmlNode        ' First element on the "higher" chain
    'Dim ndxSrc As XmlNode         ' Source node
    'Dim ndxDst As XmlNode         ' Destination node
    'Dim intI As Integer           ' Counter
    'Dim intJ As Integer           ' Counter
    'Dim intK As Integer           ' Counter
    'Dim intPtc As Integer         ' Percentage progress
    'Dim intChainId As Integer     ' Chain's id
    'Dim intRootEtreeId As Integer ' Current chain's root id
    'Dim intTargetId As Integer    ' eTreeId of the target
    'Dim intSourceId As Integer    ' eTreeId of the source
    'Dim strRefType As String      ' Reference type
    'Dim strSrcIp As String
    'Dim strDstIp As String

    Try
      ' Validate

      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/OverlapRepair error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   MoreSpecific
  ' Goal:   Check if [strPgnNew] is more specific than [strPgnOld]
  ' History:
  ' 18-08-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function MoreSpecific(ByVal strPgnNew As String, ByVal strPgnOld As String) As Boolean
    Dim strOldP As String = ""  ' Person of old
    Dim strOldG As String = ""  ' Old gender
    Dim strOldN As String = ""  ' Number old
    Dim strNewP As String = ""  ' New Person
    Dim strNewG As String = ""  ' New gender
    Dim strNewN As String = ""  ' New Number

    Try
      ' If the new one contains a semicolon - don't even consider it
      If (InStr(strPgnNew, ";") > 0) Then Return False
      ' First check the old one
      If (strPgnOld = "unknown") OrElse (strPgnOld = "empty") OrElse (strPgnOld = "") Then Return True
      ' Now split the agreement of the old and the new
      If (Not SplitAgree(strPgnNew, strNewP, strNewG, strNewN)) Then Return False
      If (Not SplitAgree(strPgnOld, strOldP, strOldG, strOldN)) Then Return False
      ' Check if the number is more specific
      If (strOldN = "") Then Return True
      ' Check if the gender is more specific
      If (strOldG = "") Then Return True
      ' Check if the person is more specific
      If (strOldP = "") Then Return True
      ' Otherwise return false
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/MoreSpecific error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   TravEtreeBack
  ' Goal:   Recursive function to go backwards, finding <eTree> elements and acting upon them
  ' History:
  ' 16-08-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function TravEtreeBack(ByRef ndxThis As XmlNode, ByRef colDone As NodeColl, ByRef tdlThis As DataSet) As Boolean
    Dim ndxChild As XmlNode ' Child node

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Check for interrupt
      If (bInterrupt) Then Return False
      ' First check all its children - last to first
      ndxChild = ndxThis.LastChild
      While (Not ndxChild Is Nothing)
        ' Process this child
        If (Not TravEtreeBack(ndxChild, colDone, tdlThis)) Then Return False
        ndxChild = ndxChild.PreviousSibling
      End While
      ' Check what we have
      Select Case ndxThis.Name
        Case "eTree"
          ' Perform the action for this item
          If (Not MakeOneChain(ndxThis, colDone, tdlThis)) Then Return False
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/TravEtreeBack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   RefListGlobal
  ' Goal:   Get table rows for global reflist measures
  ' History:
  ' 14-09-2011  ERK Derived from DoRefList (12-07-2011)
  ' 31-01-2012  ERK Added protagonist counting
  ' ------------------------------------------------------------------------------------
  Public Function RefListGlobal(ByRef tdlThis As DataSet, ByVal strFile As String, _
                                ByRef colThis As StringColl, ByRef colProt As StringColl) As Boolean
    Dim dtrFound() As DataRow     ' Result of ordering with SELECT
    Dim dtrItem() As DataRow      ' All items of this chain
    ' Dim colThis As New StringColl ' Where we gather the values
    Dim strPer As String = ""     ' Period for this file
    Dim strLine As String = ""    ' The informatino for one line
    Dim intI As Integer           ' Counter

    Try
      ' Validate
      If (tdlThis Is Nothing) OrElse (strFile = "") Then Return False
      ' Determine the period for this file
      If (Not HasPeriod(strFile, strPer, False)) Then Return False
      ' Determine the short file name
      strFile = IO.Path.GetFileNameWithoutExtension(strFile)
      ' Get the correct Global information
      dtrFound = tdlThis.Tables("Stat").Select("Type='Global'", "Name DESC")
      '' (1b) Make the header
      'colThis.Add("<tr><td>Period</td><td>File</td><td>Measure</td><td align='middle'>Count</td><td align='middle'>Total</td><td align='middle'>Percentage</td></tr>")
      ' Check if we need to make new headers
      If (colThis.Count < dtrFound.Length) Then
        ' Make new headers
        For intI = 0 To dtrFound.Length - 1
          colThis.Add("<h4>" & dtrFound(intI).Item("Name") & "</h4>" & _
                      "<table><tr><td>Period</td><td>File</td><td align='middle'>Count</td><td align='middle'>Total</td><td align='middle'>Percentage</td></tr>")
        Next intI
        ' Make one header for the protagonist counting
        colProt.Add("<h4>Protagonists</h4>" & _
                    "<table><tr><td>Period</td><td>File</td><td>ChainId</td><td>ProtIsSbj</td>" & _
                    "<td>ProtSwSbj</td><td>SbjCount</td><td>Pro</td><td>SbjPro</td><td>SbjNme</td>" & _
                    "<td>SbjZero</td><td>Length</td><td>Protagonist</td></tr>")
      End If
      ' (1c) Visist all rows with a global operation
      For intI = 0 To dtrFound.Length - 1
        ' (1d) Output statistics for this row
        With dtrFound(intI)
          strLine = "<tr><td>" & strPer & "</td><td>" & strFile & _
                      "</td><td align='middle'>" & .Item("Count") & _
                      "</td><td align='middle'>" & .Item("Total") & "</td><td align='middle'>" & _
                      Format(.Item("Count") / .Item("Total"), "p") & "</td></tr>"
          ' Append the information at this line
          colThis.Item(intI) &= strLine & vbCrLf
        End With
      Next intI
      ' Find the TEN (10) largest 3s protagonist chains in this file
      dtrFound = tdlThis.Tables("Chain").Select("PGN='3s' OR PGN='3fs' OR PGN='3ms' OR PGN='3ns' OR PGN='3'", _
                                                "Len DESC")
      For intI = 0 To Math.Min(10, dtrFound.Length - 1)
        With dtrFound(intI)
          ' Get all items of this chain
          dtrItem = tdlThis.Tables("Item").Select("ChainId=" & .Item("ChainId"), "eTreeId ASC")
          ' Process this line
          strLine = "<tr><td>" & strPer & "</td><td>" & strFile & _
                      "</td><td align='middle'>" & .Item("ChainId") & _
                      "</td><td align='middle'>" & .Item("SbjIsProt") & _
                      "</td><td align='middle'>" & .Item("SbjSwProtX") & _
                      "</td><td align='middle'>" & .Item("SbjCount") & _
                      "</td><td align='middle'>" & .Item("ProAll") & _
                      "</td><td align='middle'>" & .Item("ProSbj") & _
                      "</td><td align='middle'>" & .Item("NmeSbj") & _
                      "</td><td align='middle'>" & .Item("SbjZero") & _
                      "</td><td align='middle'>" & .Item("Len") & _
                      "</td><td align='left'>" & dtrItem(0).Item("Node") & "</td></tr>"
          ' Append the information at this line
          colProt.Add(strLine)
        End With
      Next intI
      '' Add results to [strGlobal]
      'strGlobal = colThis.Text
      ' REturn success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/RefListGlobal error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   RefListHTML
  ' Goal:   Give an overview of ALL coreferential chains in this document in HTML format
  '         Assume all necessary information is in [tdlThis] (scheme is RefChain.xsd)
  ' History:
  ' 10-08-2011  ERK Derived from DoRefList (12-07-2011)
  ' ------------------------------------------------------------------------------------
  Public Function RefListHTML(ByRef colThis As StringColl, ByRef tdlThis As DataSet) As Boolean
    Dim dtrFound() As DataRow         ' Result of ordering with SELECT
    Dim colDone As New NodeColl       ' Collection of nodes we have treated already
    Dim colStat As New StringColl     ' Statistics collection
    Dim colChain As New StringColl    ' 
    Dim colIndex As New StringColl    ' Index of chain number // chain length
    Dim dtrThis As DataRow = Nothing  ' General purpose datarow
    Dim strPre As String              ' Preamble
    Dim strName As String = ""        ' Name of the statistics measure
    Dim strRange As String = ""       ' Range of values
    Dim intJ As Integer               ' counter
    Dim intI As Integer               ' Counter
    Dim intPtc As Integer             ' Percentage
    Dim intLen As Integer = 0         ' Length of a chain
    Dim intCount As Integer           ' Intermediate count
    Dim intRng As Integer             ' Number of elements in this range
    Dim intFrom As Integer = 0        ' From range
    Dim intUntil As Integer = 0       ' Until value of range
    Dim intLast As Integer = 0        ' Value of the largest column

    Try
      ' Validate: check if we have the dataset
      If (tdlThis Is Nothing) Then Return False
      If (tdlThis.Tables("ChainList").Rows.Count = 0) Then Return False
      If (tdlThis.Tables("Chain").Rows.Count = 0) Then Return False
      ' Start producing HTML output
      strPre = "<html><body><h3>Coreferential chain list</h3><table>" & _
                  "<tr><td>File:</td><td>" & IO.Path.GetFileNameWithoutExtension(GetTableSetting(tdlThis, "File")) & "</td></tr>" & _
                  "<tr><td>Date:</td><td>" & GetTableSetting(tdlThis, "Date") & "</td></tr>" & _
                  "<tr><td>Period:</td><td>" & GetTableSetting(tdlThis, "Period") & "</td></tr>" & _
                  "</table><p>"
      colThis.Add(strPre)
      ' Other initialisations
      intChainNum = 0
      ' (1) First show the GENERAL statistics
      colThis.Add("<h4>General statistics</h4><table>")
      dtrFound = tdlThis.Tables("Stat").Select("Type='Global'", "Name ASC")
      ' (1b) Make the header
      colThis.Add("<tr><td>Name</td><td align='middle'>Count</td><td align='middle'>Total</td><td align='middle'>Percentage</td></tr>")
      ' (1c) Visist all rows with a global operation
      For intI = 0 To dtrFound.Length - 1
        ' (1d) Output statistics for this row
        With dtrFound(intI)
          colThis.Add("<tr><td>" & .Item("Name") & "</td><td align='middle'>" & .Item("Count") & _
                      "</td><td align='middle'>" & .Item("Total") & "</td><td align='middle'>" & _
                      Format(.Item("Count") / .Item("Total"), "p") & "</td></tr>")
        End With
      Next intI
      ' (1e) Finish the table
      colThis.Add("</table>")
      ' (2) Show the INTERVAL based statistics
      colThis.Add("<h4>Statistics over intervals (logbase=" & intLogBase & _
                  ")</h4><table><tr><td></td><td></td>")
      ' (2a) Header for statistics divided over intervals
      dtrFound = tdlThis.Tables("Stat").Select("Type='Interval'", "Column ASC, StatId ASC")
      For intI = 0 To dtrFound.Length - 1
        ' Is this a new section?
        With dtrFound(intI)
          If (.Item("Name") <> strName) Then
            ' Start new section
            strName = .Item("Name")
            colThis.Add("<td align='middle' colspan='3'>" & strName & "</td>")
          End If
        End With
      Next intI
      ' (2b) Finish the first header row, start second header row
      colThis.Add("</tr><tr><td>From</td><td>Until</td>") : strName = ""
      For intI = 0 To dtrFound.Length - 1
        ' Is this a new section?
        With dtrFound(intI)
          If (.Item("Name") <> strName) Then
            ' Start new section
            strName = .Item("Name")
            colThis.Add("<td align='middle'>Count</td><td align='middle'>Total</td><td align='middle'>Perc</td>")
          End If
        End With
      Next intI
      colThis.Add("</tr>")
      ' (2c) Determine the maximum column value
      dtrFound = tdlThis.Tables("Stat").Select("Type='Interval'", "Column DESC")
      intLast = dtrFound(0).Item("Column")
      ' (2d) Gather information per section
      '     Get the total number we should be aware of 
      intLen = tdlThis.Tables("Stat").Select("Type='Interval'").Length
      intCount = 0 : intFrom = 1 : intUntil = 1 : intRng = 0
      ' (2e) Walk all possible ranges
      Do
        ' Get rows within this range
        dtrFound = tdlThis.Tables("Stat").Select("Type='Interval' AND From >= " & intFrom & _
                                                 " AND Until <= " & intUntil, "StatId ASC")
        ' Determine how many chains there are with this length
        intRng = dtrFound.Length
        ' Output the range values for this line
        colThis.Add("<tr><td>" & intFrom & "</td><td>" & intUntil & "</td>")
        ' Visit all the columns
        intJ = 0
        For intI = 1 To intLast
          ' Check if this item has the correct column
          If (intJ < dtrFound.Length) AndAlso (dtrFound(intJ).Item("Column") = intI) Then
            ' Fill in the values for this column
            With dtrFound(intJ)
              colThis.Add("<td align='middle'>" & .Item("Count") & "</td><td align='middle'>" & .Item("Total") & _
                          "</td><td align='middle'>" & Format(.Item("Count") / .Item("Total"), "p") & "</td>")
            End With
            ' Go to the next item
            intJ += 1
          Else
            ' Fill in empty values
            colThis.Add("<td align='middle'>-</td><td align='middle'>-</td><td align='middle'>-</td>")
          End If
        Next intI
        ' Finish this row
        colThis.Add("</tr>")
        ' Adapt the total count of elements we had
        intCount += intRng
        ' Calculate the next range
        intUntil = intUntil * intLogBase
        intFrom = (intUntil \ intLogBase) + 1
      Loop While (intCount < intLen)

      ' Finish this table
      colThis.Add("</table>")
      ' (3) Next show the specific statistics
      dtrFound = tdlThis.Tables("Stat").Select("Type='One'", "Name ASC, StatId ASC")
      For intI = 0 To dtrFound.Length - 1
        With dtrFound(intI)
          ' Is this a new section?
          If (.Item("Name") <> strName) Then
            ' Need to finish the old section?
            If (strName <> "") Then colThis.Add("</table>")
            ' Start the new section
            strName = .Item("Name")
            colThis.Add("<h4>" & strName & "</h4><table><tr><td>From</td><td>Until</td><td>Count</td><td>Total</td><td>Percentage</td></tr>")
          End If
          ' Add this row with data
          colThis.Add("<tr><td>" & .Item("From") & "</td><td>" & .Item("Until") & "</td><td>" & .Item("Count") & "</td><td>" & _
                      .Item("Total") & "</td><td>" & Format(.Item("Count") / .Item("Total"), "p") & "</td></tr>")
        End With
      Next intI
      ' Finish the last table
      colThis.Add("</table>")
      ' (3) Collect all the chains -- in order of decreasing length
      dtrFound = tdlThis.Tables("Chain").Select("", "Len DESC, ChainId ASC")
      ' Walk through all elements of the table
      For intI = 0 To dtrFound.Length - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ dtrFound.Length
        Status("RefListHtml " & intPtc, intPtc)
        ' Check for interrupt
        If (bInterrupt) Then Return False
        ' Show the information for this chain
        If (Not ShowOneChain(tdlThis, dtrFound(intI).Item("ChainId"), colChain, colIndex)) Then Return False
      Next intI
      ' (4) Show the list of chains
      colThis.Add("<h4>List of chains</h4>")
      colThis.Add(colChain.Text)
      ' Add the index of chains
      colThis.Add("<h4>Index</h4><table><tr><td>Length</td><td>Chain number</td></tr>")
      For intI = 0 To colIndex.Count - 1
        colThis.Add("<tr><td>" & colIndex.Item(intI) & "</td><td>" & colIndex.Exmp(intI) & "</td></tr>")
      Next intI
      ' Finish this table
      colThis.Add("</table>")
      ' (5) Show a list of the different subject referents
      colThis.Add("<p>Here are several lists of subject-referents in main clauses." & vbCrLf & _
                  "<p>The first list is for <i>all</i> subjects, " & _
                  "<br>The second list is for all <i>main</i> clause subjects, " & _
                  "<br>the third one for <i>3rd person</i> main clause subjects." & _
                  "<p>Each list only gives the <i>first</i> occurrance of a subject-referent.")
      ' (5a) First *all* subject referents of all clauses
      colThis.Add(SbjRefList(tdlThis, "SbjRefAll"))
      ' (5b) Then *all* subject referents of all *main* clauses
      colThis.Add(SbjRefList(tdlThis, "SbjRefMat"))
      ' (5c) Now only the 3rd person singular subject referents of all main clauses
      colThis.Add(SbjRefList(tdlThis, "SbjRefMat3"))
      ' Finish the html file
      colThis.Add("</body></html>")
      ' Return positively
      Status("Ready")
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/RefListHTML error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SbjRefList
  ' Goal:   List of all different [strType]-clause subject referents
  ' History:
  ' 31-01-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function SbjRefList(ByRef tdlThis As DataSet, ByVal strType As String) As String
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim colBack As New StringColl ' What we return
    Dim strRoot As String         ' Text of root node
    Dim intI As Integer           ' Counter

    Try
      ' Validate
      If (tdlThis Is Nothing) Then Return ""
      ' Get the correct information and get it in order
      dtrFound = tdlThis.Tables("Item").Select("GrRole='Subject' AND " & strType & "='True'", _
                                               "forestId ASC")
      ' Initialise
      colBack.Add("<h4>" & strType & "</h4><table><tr><td>Loc</td><td>ChainId</td><td>PGN</td><td>RefType</td>" & _
                  "<td>NPtype</td><td>Node</td></tr>")
      ' Walk the results
      For intI = 0 To dtrFound.Length - 1
        With dtrFound(intI)
          ' Get root node
          strRoot = NodeText(GetRootNode(IdToNode(dtrFound(intI).Item("eTreeId"))), False)
          ' Provide  output for this row
          colBack.Add("<tr><td>" & .Item("Loc") & "</td>" & _
                      "<td>" & .Item("ChainId") & "</td>" & _
                      "<td>" & .Item("PGN") & "</td>" & _
                      "<td>" & .Item("RefType") & "</td>" & _
                      "<td>" & .Item("NPtype") & "</td>" & _
                      "<td>" & strRoot & "</td></tr>")
        End With
      Next intI
      ' Finish table
      colBack.Add("</table>")
      ' Return the result
      Return colBack.Text
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/SbjRefList error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ChainStatSbj
  ' Goal:   Calculate the percentage of chain items with GrRole="Subject"
  '         Exclude chains with NPtype="Trace"
  '         Do this for chains with lengths in groups of 1, 2, 3-4, 5-8, 9-16 etc.
  ' History:
  ' 10-08-2011  ERK Created
  ' 17-08-2011  ERK Worked around for <StatList>
  ' ------------------------------------------------------------------------------------
  Private Function ChainStatSbj(ByRef tdlThis As DataSet, ByVal strFile As String, _
         ByVal intColumn As Integer, ByVal strName As String, ByVal bTrace As Boolean) As Boolean
    Dim dtrFound() As DataRow   ' Result of SELECT
    Dim intI As Integer         ' Counter
    Dim intChainId As Integer   ' ID of this chain
    Dim intLen As Integer       ' Length we are currently working with
    Dim intMin As Integer       ' Minimal length
    Dim intMax As Integer       ' Maximal length
    Dim intCount As Integer = 0 ' Number of chains done
    Dim intSbj As Integer       ' Number of subjects
    Dim intRng As Integer       ' Number of elements in this range
    Dim intBase As Integer      ' The base number (all elements that are not a "Trace")
    Dim intSbjTot As Integer    ' Total subject
    Dim intBaseTot As Integer   ' Total base

    Try
      ' Validate
      If (tdlThis Is Nothing) Then Return False
      ' Set initial lengths
      intMin = 1 : intMax = 1 : intLen = tdlThis.Tables("Item").Rows.Count
      intSbjTot = 0 : intBaseTot = 0
      Do
        ' Get rows within this range
        dtrFound = tdlThis.Tables("Chain").Select("Len >= " & intMin & " AND Len <= " & intMax)
        ' Count all subjects and all elements in this range
        intRng = 0 : intSbj = 0 : intBase = 0
        For intI = 0 To dtrFound.Length - 1
          ' Adapt the number of items
          intRng += dtrFound(intI).Item("Len")
          ' Get my own Chain Id
          intChainId = dtrFound(intI).Item("ChainId")
          ' Calculate the base number
          If (bTrace) Then
            intBase += tdlThis.Tables("Item").Select("ChainId=" & intChainId).Length
          Else
            intBase += tdlThis.Tables("Item").Select("ChainId=" & intChainId & _
                        " AND NPtype<>'Trace'").Length
          End If
          ' Get all Chain Items for this Chain
          If (bTrace) Then
            intSbj += tdlThis.Tables("Item").Select("ChainId=" & intChainId & " AND GrRole='Subject'").Length
          Else
            intSbj += tdlThis.Tables("Item").Select("ChainId=" & intChainId & _
                                                    " AND GrRole='Subject' AND NPtype<>'Trace'").Length
          End If
        Next intI
        ' Adapt grand totals
        intSbjTot += intSbj
        intBaseTot += intBase
        ' Store the results for this range in the dataset
        If (Not AddStatRow(tdlThis, strFile, intColumn, intMin, intMax, strName, intSbj, intBase, "Interval")) Then Return False
        ' Adapt the total count of elements we had
        intCount += intRng
        ' Calculate the next range
        intMax = intMax * intLogBase
        intMin = (intMax \ intLogBase) + 1
      Loop While (intCount < intLen)
      ' Add the grand total
      If (Not AddStatRow(tdlThis, strFile, 0, 0, 0, strName, intSbjTot, intBaseTot, "Global")) Then Return False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/ChainStatSbj error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ChainStatPro
  ' Goal:   Calculate the percentage of chain items with NPtype="Pro|PossPro|ZeroSbj|Dem"
  '         Exclude chains with NPtype="Trace"
  '         Do this for chains with lengths in groups of 1, 2, 3-4, 5-8, 9-16 etc.
  ' History:
  ' 10-08-2011  ERK Created
  ' 17-08-2011  ERK Adapted for <StatList>
  ' ------------------------------------------------------------------------------------
  Private Function ChainStatPro(ByRef tdlThis As DataSet, ByVal strFile As String, _
           ByVal intColumn As Integer, ByVal strName As String) As Boolean
    Dim dtrFound() As DataRow   ' Result of SELECT
    Dim intI As Integer         ' Counter
    Dim intChainId As Integer   ' ID of this chain
    Dim intLen As Integer       ' Length we are currently working with
    Dim intMin As Integer       ' Minimal length
    Dim intMax As Integer       ' Maximal length
    Dim intCount As Integer = 0 ' Number of chains done
    Dim intPro As Integer       ' Number of items with a pronominal element
    Dim intRng As Integer       ' Number of elements in this range
    Dim intBase As Integer      ' The base number (all elements that are not a "Trace")
    Dim intProTot As Integer    ' Total pronoun
    Dim intBaseTot As Integer   ' Total base

    Try
      ' Validate
      If (tdlThis Is Nothing) Then Return False
      ' Set initial lengths
      intMin = 1 : intMax = 1 : intLen = tdlThis.Tables("Item").Rows.Count
      intProTot = 0 : intBaseTot = 0
      Do
        ' Get rows within this range
        dtrFound = tdlThis.Tables("Chain").Select("Len >= " & intMin & " AND Len <= " & intMax)
        ' Count all subjects and all elements in this range
        intRng = 0 : intPro = 0 : intBase = 0
        For intI = 0 To dtrFound.Length - 1
          ' Adapt the number of items
          intRng += dtrFound(intI).Item("Len")
          ' Get my own Chain Id
          intChainId = dtrFound(intI).Item("ChainId")
          ' Calculate the base number
          intBase += tdlThis.Tables("Item").Select("ChainId=" & intChainId & _
                      " AND NPtype<>'Trace'").Length
          ' Get all Chain Items for this Chain
          intPro += tdlThis.Tables("Item").Select("ChainId=" & intChainId & _
                      " AND (NPtype='Pro' OR NPtype='PossPro' OR NPtype='ZeroSbj' OR NPtype='Dem')").Length
        Next intI
        ' Adapt grand totals
        intProTot += intPro
        intBaseTot += intBase
        ' Store the results for this range in the dataset
        If (Not AddStatRow(tdlThis, strFile, intColumn, intMin, intMax, strName, intPro, intBase, "Interval")) Then Return False
        ' Adapt the total count of elements we had
        intCount += intRng
        ' Calculate the next range
        intMax = intMax * intLogBase
        intMin = (intMax \ intLogBase) + 1
      Loop While (intCount < tdlThis.Tables("Item").Rows.Count)
      ' Add the grand total
      If (Not AddStatRow(tdlThis, strFile, 0, 0, 0, strName, intProTot, intBaseTot, "Global")) Then Return False
      ' Check the grand total for just pronomina
      intProTot = tdlThis.Tables("Item").Select("NPtype='Pro' OR NPtype='PossPro' OR NPtype='Dem'").Length
      If (Not AddStatRow(tdlThis, strFile, 0, 0, 0, "NonZero" & strName, intProTot, intBaseTot, "Global")) Then Return False
      ' Check the grand total for just pronomina, but only when they are subject
      intProTot = tdlThis.Tables("Item").Select("(NPtype='Pro' OR NPtype='PossPro' OR NPtype='Dem') AND GrRole='Subject'").Length
      intBaseTot = tdlThis.Tables("Item").Select("NPtype<>'Trace' AND GrRole='Subject'").Length
      If (Not AddStatRow(tdlThis, strFile, 0, 0, 0, "NonZeroSbj" & strName, intProTot, intBaseTot, "Global")) Then Return False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/ChainStatPro error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ChainStatProDist
  ' Goal:   Calculate the antecedent distances of chain items with NPtype="Pro|PossPro|ZeroSbj|Dem"
  '         Exclude chain items with NPtype="Trace"
  '         Do this for length groups of 1, 2, 3-4, 5-8, 9-16 etc.
  ' History:
  ' 10-08-2011  ERK Created
  ' 17-08-2011  ERK Adapted for <StatList>
  ' 09-02-2012  ERK Added 3rd person singular SUBJECT pronomina ARD measure
  ' ------------------------------------------------------------------------------------
  Private Function ChainStatProDist(ByRef tdlThis As DataSet, ByVal strFile As String, _
          ByVal intColumn As Integer, ByVal strName As String, ByVal strDistType As String) As Boolean
    Dim dtrFound() As DataRow   ' Result of SELECT
    Dim intMin As Integer       ' Minimal length
    Dim intMax As Integer       ' Maximal length
    Dim intI As Integer         ' Counter
    Dim intCount As Integer     ' Number of chains done
    Dim intPro As Integer       ' Number of items with a pronominal element
    Dim intBase As Integer      ' The base number (all elements that are not a "Trace")
    Dim intProTot As Integer    ' Total pronoun
    Dim strDefPro As String = "(NPtype='Pro' OR NPtype='PossPro' OR NPtype='ZeroSbj' OR NPtype='Dem')" & _
                              " AND (PGN='3s' OR PGN='3ms' OR PGN='3fs' OR PGN='3ns')"

    Try
      ' Validate
      If (tdlThis Is Nothing) Then Return False
      ' Set initial lengths
      intMin = 1 : intMax = 1 : intPro = 0 : intCount = 0
      ' Determine the base number
      dtrFound = tdlThis.Tables("Item").Select("Dist" & strDistType & ">0" & " AND " & strDefPro)
      intBase = dtrFound.Length
      Do
        ' Get items with a @DistForest within this range
        dtrFound = tdlThis.Tables("Item").Select("Dist" & strDistType & " >= " & intMin & _
                      " AND Dist" & strDistType & " <= " & intMax & " AND " & strDefPro, _
                      "Dist" & strDistType & " ASC")
        intPro = dtrFound.Length
        ' Adapt maximum
        intCount += intPro
        ' Store the results for this range in the dataset
        If (Not AddStatRow(tdlThis, strFile, intColumn, intMin, intMax, strName & strDistType, intPro, intBase, "Interval")) Then Return False
        ' Calculate the next range
        intMax = intMax * intLogBase
        intMin = (intMax \ intLogBase) + 1
      Loop While (intCount < intBase)
      ' Do we have a maximum?
      If (intPro > 0) Then
        ' Get the most maximum distance
        intProTot = dtrFound(intPro - 1).Item("Dist" & strDistType)
        ' Get the Chain Id of this guy
        intBase = dtrFound(intPro - 1).Item("ChainId")
      Else
        ' There is no maximum
        intProTot = 0
        intBase = 1
      End If
      ' Add the maximum found
      If (Not AddStatRow(tdlThis, strFile, 0, 0, 0, strName & strDistType, intProTot, intBase, "Global")) Then Return False
      ' We also want to know the average
      ' (1) Initialise
      dtrFound = tdlThis.Tables("Item").Select("Dist" & strDistType & ">0" & " AND " & strDefPro)
      intCount = 0 : intBase = dtrFound.Length
      ' (2) Loop through all items
      For intI = 0 To dtrFound.Length - 1
        ' Add this distance
        intCount += dtrFound(intI).Item("Dist" & strDistType)
      Next intI
      ' Process these numbers in a global statistics row
      If (Not AddStatRow(tdlThis, strFile, 0, 0, 0, "Avg" & strName & strDistType, intCount, intBase, "Global")) Then Return False
      ' We also want to know the average
      ' (1) Initialise
      dtrFound = tdlThis.Tables("Item").Select("Dist" & strDistType & ">0" & " AND " & strDefPro & " AND GrRole='Subject'")
      intCount = 0 : intBase = dtrFound.Length
      ' (2) Loop through all items
      For intI = 0 To dtrFound.Length - 1
        ' Add this distance
        intCount += dtrFound(intI).Item("Dist" & strDistType)
      Next intI
      ' Process these numbers in a global statistics row
      If (Not AddStatRow(tdlThis, strFile, 0, 0, 0, "AvgSbj" & strName & strDistType, intCount, intBase, "Global")) Then Return False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/ChainStatProDist error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DoChainSbjSwitch
  ' Goal:   Calculate the number of times a Subject becomes something else
  '         Exclude cases where the subject is a Trace
  ' History:
  ' 12-08-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function DoChainSbjSwitch(ByRef tdlThis As DataSet, ByRef intCount As Integer) As Boolean
    Dim dtrFound() As DataRow   ' Result of SELECT
    Dim dtrItem() As DataRow    ' Set of chain items
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter
    Dim intChainId As Integer   ' ID of this chain

    Try
      ' Validate
      If (tdlThis Is Nothing) Then Return False
      ' Initialisations
      intCount = 0
      ' Walk ALL chains larger than size 1 (there is no switch in those of size 1
      dtrFound = tdlThis.Tables("Chain").Select("Len > 1")
      For intI = 0 To dtrFound.Length - 1
        ' Get this chain's number
        intChainId = dtrFound(intI).Item("ChainId")
        ' Consider all chain items (in order) where the NP type is not Trace
        dtrItem = tdlThis.Tables("Item").Select("ChainId=" & intChainId & " AND NPtype <> 'Trace'", "ItemId ASC")
        For intJ = 1 To dtrItem.Length - 1
          ' Check if THIS item is NOT a subject, but the PREVIOUS was
          If (dtrItem(intJ).Item("GrRole") <> "Subject") AndAlso (dtrItem(intJ - 1).Item("GrRole") = "Subject") Then
            ' Subject switch counter
            intCount += 1
          End If
        Next intJ
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/DoChainSbjSwitch error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DoProtSbjRatio
  ' Goal:   Calculate the number of times protagonist[x] is the subject
  '         divided by the number of subjects occurring during the lifespan of protagonist[x]
  '
  '         Also count the following:
  '           ProAll = Number of pronouns per protagonist
  '           ProSbj = Number of subject pronouns per protagonist
  '         Also determine the Protagonist Subject Lifespan Disruptance =
  '           Number of subject referent switches during lifespan of protagonist[x]
  '           divided by total number of subjects during the lifespan of protagonist[x]
  ' History:
  ' 12-08-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function DoProtSbjRatio(ByRef tdlThis As DataSet, ByRef intCount As Integer) As Boolean
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim dtrItem() As DataRow      ' Set of chain items
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intChainId As Integer     ' ID of this chain
    Dim intSbjIsProtX As Integer  ' Number of times protagonist[x] is the subject 
    Dim intSbjNumX As Integer     ' Number of subjects during lifespan protagonist[x]
    Dim intSbjSwProtX As Integer  ' Number of subject referent switches during lifespan prot[x]
    Dim intLifeBeg As Integer     ' Starting line number of the lifespan of protagonist[x]
    Dim intLifeEnd As Integer     ' Finishing line number of the lifespan of protagonist[x]
    Dim intSbjId As Integer       ' Chain ID of the subject
    Dim dblSbjIsProt As Double    ' Total
    Dim dblSbjCount As Double     ' Total
    Dim intProAll As Integer      ' Number of pronouns per protagonist
    Dim intProSbj As Integer      ' Number of subject-pronouns per protagonist
    Dim intNmeSbj As Integer      ' Number of subject-proper names per protagonist
    Dim intSbjZero As Integer     ' Number of zero subjects (Ellipsis)

    Try
      ' Validate
      If (tdlThis Is Nothing) Then Return False
      ' Initialisations
      intCount = 0 : dblSbjIsProt = 0 : dblSbjCount = 0
      ' Walk ALL chains larger than size 1 (Don't count chains of size 0)
      dtrFound = tdlThis.Tables("Chain").Select("NoTraceLen > 1")
      For intI = 0 To dtrFound.Length - 1
        ' Get this chain's number
        intChainId = dtrFound(intI).Item("ChainId")
        ' ============ DEBUGGING ============
        ' If (intChainId = 293) Then Stop
        ' ===================================
        ' Consider all chain items (in order) where the NP type is not Trace
        dtrItem = tdlThis.Tables("Item").Select("ChainId=" & intChainId, "ItemId ASC")
        ' Check the "lifespan"
        intLifeEnd = dtrItem(0).Item("forestId")
        intLifeBeg = dtrItem(dtrItem.Length - 1).Item("forestId")
        ' Get the number of times this protagonist is a grammatical subject
        intSbjIsProtX = tdlThis.Tables("Item").Select("ChainId=" & intChainId & " AND GrRole='Subject'").Length
        ' Get the number of subjects during the lifespan of protagonist[x]
        intSbjNumX = tdlThis.Tables("Item").Select("forestId>=" & intLifeBeg & _
          " AND forestId<=" & intLifeEnd & " AND GrRole='Subject'").Length
        ' Consider all forests in the lifespan of this chain item
        dtrItem = tdlThis.Tables("Item").Select("forestId>=" & intLifeBeg & _
            " AND forestId<=" & intLifeEnd & " AND GrRole='Subject'", "forestId ASC")
        ' Initialise
        intSbjSwProtX = 0
        ' Check if there are any items
        If (dtrItem.Length > 0) Then
          ' Get the chain Id of the subject
          intSbjId = dtrItem(0).Item("ChainId")
          ' Walk the lifespan of this chain Id
          For intJ = 1 To dtrItem.Length - 1
            ' Is this a different chain id?
            If (intSbjId <> dtrItem(intJ).Item("ChainId")) Then
              ' Keep track of changes
              intSbjId = dtrItem(intJ).Item("ChainId")
              intSbjSwProtX += 1
            End If
          Next intJ
        End If
        ' Get the number of pronouns
        intProAll = tdlThis.Tables("Item").Select("ChainId=" & intChainId & _
            " AND (NPtype='Pro' OR NPtype='PossPro')").Length
        ' Get the number of subject pronouns
        intProSbj = tdlThis.Tables("Item").Select("ChainId=" & intChainId & _
            " AND GrRole='Subject' AND NPtype='Pro'").Length
        ' Get the number of subject proper names
        intNmeSbj = tdlThis.Tables("Item").Select("ChainId=" & intChainId & _
            " AND GrRole='Subject' AND NPtype='Proper'").Length
        ' Get the number of zero subjects
        intSbjZero = tdlThis.Tables("Item").Select("ChainId=" & intChainId & _
            " AND NPtype='ZeroSbj'").Length
        ' Store these numbers
        With dtrFound(intI)
          .Item("SbjIsProt") = intSbjIsProtX
          .Item("SbjCount") = intSbjNumX
          .Item("SbjSwProtX") = intSbjSwProtX
          .Item("ProAll") = intProAll
          .Item("ProSbj") = intProSbj
          .Item("NmeSbj") = intNmeSbj
          .Item("SbjZero") = intSbjZero
        End With
        ' Keep track of totals
        dblSbjIsProt += intSbjIsProtX : dblSbjCount += intSbjNumX
        ' Increment the counter
        intCount += 1
      Next intI
      ' TODO: Possibly store grand totals??
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/DoProtSbjRatio error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DoSbjRefSwitch
  ' Goal:   Get the total number of:
  '           Referent(sbj(line[i])) <> 
  '           Referent(sbj(line[i+1]))
  ' Note:   The domain can be any of the following:
  '           chunk  - One chunk is an IP-SUB, IP-MAT or a part of it
  '           forest - One forest is a line in the text
  '           ipcls  - A whole IP-MAT or IP-SUB, but not if the parent is CP-REL
  '           ipmat  - Only IP-MAT
  '           ipmat3 - Only IP-MAT and only 3rd person (sg/pl) referents
  '           CountIpMat  - Number of different subject referents in main clauses
  '           CountIpMat3 - As above, but for 3rd person
  '           CountIpAll  - Number of different subject referents overall
  '         The [intMaxId] should return the maximum number of domain elements
  ' History:
  ' 09-09-2011  ERK Created
  ' 15-09-2011  ERK Added "ipcls" and "ipmat" types
  ' 16-09-2011  ERK Added "ipmat3" type
  ' 31-01-2012  ERK Added "CountIpMat" and "CountIpMat3" types
  ' 02-02-2012  ERK Added "CountIpAll"
  ' ------------------------------------------------------------------------------------
  Private Function DoSbjRefSwitch(ByRef tdlThis As DataSet, ByRef intCount As Integer, _
                                  ByVal strDomain As String, ByRef intMaxId As Integer) As Boolean
    Dim dtrItem() As DataRow        ' Set of chain items
    Dim dtrParent As DataRow        ' Perent row
    Dim strPgn As String            ' The PGN of this chain
    Dim intI As Integer             ' Counter
    Dim intChainId As Integer       ' ID of this chain
    Dim intSbjChainId As Integer    ' The ChainID of the current subject
    Dim intIPmatNum As Integer = 0  ' Number of IP-MAT items
    Dim colSbjId As New StringColl  ' The different subject referents

    Try
      ' Validate
      If (tdlThis Is Nothing) Then Return False
      ' Initialisations
      intCount = 0 : intSbjChainId = -1
      Select Case strDomain
        Case "chunk"  ' Look at increasing [IPnum]
          ' Find out what the maximum IPnum is
          intMaxId = tdlThis.Tables("Item").Select("", "IPnum DESC")(0).Item("IPnum")
          ' Walk through all the IPnums
          For intI = 1 To intMaxId
            ' Check if this IPnum has a subject
            dtrItem = tdlThis.Tables("Item").Select("IPnum=" & intI & " AND GrRole='Subject'")
            If (dtrItem.Length > 0) Then
              ' There is a subject-- get its ID
              intChainId = dtrItem(0).Item("ChainId")
              ' See this differs from the first subject
              If (intSbjChainId < 0) Then
                ' This is the first time, so there is no switch here
                intSbjChainId = intChainId
              ElseIf (intSbjChainId <> intChainId) Then
                ' There is switch of subject: go to the new subject, and keep track of the number of switches
                intSbjChainId = intChainId
                intCount += 1
              End If
            End If
          Next intI
        Case "forest" ' Look at increasing @forestId
          ' Find out what the maximum forestId is
          intMaxId = tdlThis.Tables("Item").Select("", "forestId DESC")(0).Item("forestId")
          ' Walk through all the @forestId
          For intI = 1 To intMaxId
            ' Check if this @forestId has a subject
            dtrItem = tdlThis.Tables("Item").Select("forestId=" & intI & " AND GrRole='Subject'")
            If (dtrItem.Length > 0) Then
              ' Possibly there are more subjects in one <forest>. 
              '   - Get the ChainId of the first one
              intChainId = dtrItem(0).Item("ChainId")
              ' See this differs from the current subject
              If (intSbjChainId < 0) Then
                ' This is the first time, so there is no switch here
                intSbjChainId = intChainId
              ElseIf (intSbjChainId <> intChainId) Then
                ' There is switch of subject: go to the new subject, and keep track of the number of switches
                intSbjChainId = intChainId
                intCount += 1
              End If
            End If
          Next intI
        Case "ipcls" ' Look at increasing @IPclsId
          ' Find out what the maximum @IPmatIds is
          intMaxId = tdlThis.Tables("Item").Select("", "IPclsId DESC")(0).Item("IPclsId")
          ' Walk through all the @forestId
          For intI = 1 To intMaxId
            ' Check if this @IPclsId has a subject
            dtrItem = tdlThis.Tables("Item").Select("IPclsId=" & intI & " AND GrRole='Subject'")
            If (dtrItem.Length > 0) Then
              ' Increment the total number of IPmats
              intIPmatNum += 1
              ' Possibly there are more subjects in one <forest>. 
              '   - Get the ChainId of the first one
              intChainId = dtrItem(0).Item("ChainId")
              ' See this differs from the current subject
              If (intSbjChainId < 0) Then
                ' This is the first time, so there is no switch here
                intSbjChainId = intChainId
              ElseIf (intSbjChainId <> intChainId) AndAlso (intChainId >= 0) Then
                ' There is switch of subject: go to the new subject, and keep track of the number of switches
                intSbjChainId = intChainId
                intCount += 1
              End If
            End If
          Next intI
          ' Adapt the NUMBER of different IP-MAT/IP-SUB items
          intMaxId = intIPmatNum
        Case "ipmat" ' Look at increasing @IPmatId
          ' Find out what the maximum @IPmatIds is
          intMaxId = tdlThis.Tables("Item").Select("", "IPmatId DESC")(0).Item("IPmatId")
          ' Walk through all the @forestId
          For intI = 1 To intMaxId
            ' Check if this @IPmatId has a subject
            dtrItem = tdlThis.Tables("Item").Select("IPmatId=" & intI & " AND GrRole='Subject'")
            If (dtrItem.Length > 0) Then
              ' Increment the total number of IPmats
              intIPmatNum += 1
              ' Possibly there are more subjects in one <forest>. 
              '   - Get the ChainId of the first one
              intChainId = dtrItem(0).Item("ChainId")
              ' See this differs from the current subject
              If (intSbjChainId < 0) Then
                ' This is the first time, so there is no switch here
                intSbjChainId = intChainId
              ElseIf (intSbjChainId <> intChainId) AndAlso (intChainId >= 0) Then
                ' There is switch of subject: go to the new subject, and keep track of the number of switches
                intSbjChainId = intChainId
                intCount += 1
              End If
            End If
          Next intI
          ' Adapt the NUMBER of different IP-MAT items
          intMaxId = intIPmatNum
        Case "ipmat3" ' Look at increasing @IPmatId, but only if the subject referent is 3rd person
          ' Find out what the maximum @IPmatIds is
          intMaxId = tdlThis.Tables("Item").Select("", "IPmatId DESC")(0).Item("IPmatId")
          ' Walk through all the @forestId
          For intI = 1 To intMaxId
            ' Check if this @IPmatId has a subject
            dtrItem = tdlThis.Tables("Item").Select("IPmatId=" & intI & " AND GrRole='Subject'")
            If (dtrItem.Length > 0) Then
              ' Get the parent of <Item>
              dtrParent = dtrItem(0).GetParentRow("Chain_Item")
              ' Check the PGN of the chain in which this belongs
              strPgn = dtrParent.Item("PGN").ToString
              If (strPgn <> "") AndAlso (Left(strPgn, 1) = "3") Then
                ' Increment the total number of IPmats
                intIPmatNum += 1
                ' Possibly there are more subjects in one <forest>. 
                '   - Get the ChainId of the first one
                intChainId = dtrItem(0).Item("ChainId")
                ' See this differs from the current subject
                If (intSbjChainId < 0) Then
                  ' This is the first time, so there is no switch here
                  intSbjChainId = intChainId
                ElseIf (intSbjChainId <> intChainId) AndAlso (intChainId >= 0) Then
                  ' There is switch of subject: go to the new subject, and keep track of the number of switches
                  intSbjChainId = intChainId
                  intCount += 1
                End If
              End If
            End If
          Next intI
          ' Adapt the NUMBER of different IP-MAT items
          intMaxId = intIPmatNum
        Case "CountIpAll" ' Calculate number of different subject referents
          ' Find out what the maximum @IPmatIds is
          dtrItem = tdlThis.Tables("Item").Select("GrRole='Subject'", "IPmatId DESC")
          ' Set the total number of main clauses with a subject
          intIPmatNum = dtrItem.Length
          ' Walk through the results
          For intI = 0 To dtrItem.Length - 1
            ' Get the first subject's chain Id
            intChainId = dtrItem(intI).Item("ChainId")
            ' See if we stored this one already
            If (Not colSbjId.Exists(intChainId)) Then
              ' Store each unique subject-referent, with the chain root
              colSbjId.Add(dtrItem(intI).Item("ChainId"))
              ' Set the [SbjRefMat] field of this item
              dtrItem(intI).Item("SbjRefAll") = "True"
            End If
          Next intI
          ' Get the total count
          intCount = colSbjId.Count
          ' Adapt the NUMBER of different IP-MAT items
          intMaxId = intIPmatNum
        Case "CountIpMat" ' Look at increasing @IPmatId
          ' Find out what the maximum @IPmatIds is
          dtrItem = tdlThis.Tables("Item").Select("GrRole='Subject' AND IPclsLb LIKE 'IP-MAT*'", "IPmatId DESC")
          ' Set the total number of main clauses with a asubject
          intIPmatNum = dtrItem.Length
          ' Walk through the results
          For intI = 0 To dtrItem.Length - 1
            ' Get the first subject's chain Id
            intChainId = dtrItem(intI).Item("ChainId")
            ' See if we stored this one already
            If (Not colSbjId.Exists(intChainId)) Then
              ' Store each unique subject-referent, with the chain root
              colSbjId.Add(dtrItem(intI).Item("ChainId"))
              ' Set the [SbjRefMat] field of this item
              dtrItem(intI).Item("SbjRefMat") = "True"
            End If
          Next intI
          ' Get the total count
          intCount = colSbjId.Count
          ' Adapt the NUMBER of different IP-MAT items
          intMaxId = intIPmatNum
        Case "CountIpMat3" ' Look at increasing @IPmatId
          ' Find out what the maximum @IPmatIds is
          dtrItem = tdlThis.Tables("Item").Select("GrRole='Subject' AND IPclsLb LIKE 'IP-MAT*'", "IPmatId DESC")
          ' Set the total number of main clauses with a asubject
          intIPmatNum = dtrItem.Length
          ' Walk through the results
          For intI = 0 To dtrItem.Length - 1
            ' Get the first subject's chain Id
            intChainId = dtrItem(intI).Item("ChainId")
            ' Get the parent of <Item>
            dtrParent = dtrItem(intI).GetParentRow("Chain_Item")
            ' Check the PGN of the chain in which this belongs
            strPgn = dtrParent.Item("PGN").ToString
            '' Debugging
            'If (intChainId <> 5) Then Stop
            If (strPgn <> "") AndAlso (Left(strPgn, 1) = "3") Then
              ' See if we stored this one already
              If (Not colSbjId.Exists(intChainId)) Then
                ' Store each unique subject-referent, with the chain root
                colSbjId.Add(dtrItem(intI).Item("ChainId"))
                ' Set the [SbjRefMat] field of this item
                dtrItem(intI).Item("SbjRefMat3") = "True"
              End If
              ' Increment the total number of main clauses with a subjct
              intIPmatNum += 1
            End If
          Next intI
          ' Get the total count
          intCount = colSbjId.Count
          ' Adapt the NUMBER of different IP-MAT items
          intMaxId = intIPmatNum
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/DoSbjRefSwitch error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ChainStatFreq
  ' Goal:   Calculate the number of chains with a particular length (in intervals)
  ' History:
  ' 10-08-2011  ERK Created
  ' 17-08-2011  ERK Adapted for <StatList> idea
  ' ------------------------------------------------------------------------------------
  Private Function ChainStatFreq(ByRef tdlThis As DataSet, ByVal strFile As String, _
        ByVal intColumn As Integer, ByVal strName As String, Optional ByVal bTraces As Boolean = True) As Boolean
    Dim dtrFound() As DataRow   ' Result of SELECT
    Dim intLen As Integer       ' Length we are currently working with
    Dim intMin As Integer       ' Minimal length
    Dim intMax As Integer       ' Maximal length
    Dim intCount As Integer = 0 ' Number of chains done
    Dim intRng As Integer       ' Number of elements in this range
    Dim intProTot As Integer    ' Total pronoun
    Dim intBaseTot As Integer   ' Total base

    Try
      ' Validate
      If (tdlThis Is Nothing) Then Return False
      ' Set initial lengths
      intMin = 1 : intMax = 1
      intProTot = 0 : intBaseTot = 0
      ' Get the total number of chains
      If (bTraces) Then
        ' Traces are allowed, so we look at ALL chains
        intLen = tdlThis.Tables("Chain").Rows.Count
      Else
        ' Traces are not allowed -- consider all chains with a "no-trace-length" larger than zero
        intLen = tdlThis.Tables("Chain").Select("NoTraceLen > 0").Length
      End If
      Do
        ' Get rows within this range
        If (bTraces) Then
          dtrFound = tdlThis.Tables("Chain").Select("Len >= " & intMin & " AND Len <= " & intMax)
        Else
          dtrFound = tdlThis.Tables("Chain").Select("NoTraceLen >= " & intMin & " AND NoTraceLen <= " & intMax)
        End If
        ' Determine how many chains there are with this length
        intRng = dtrFound.Length
        ' Store the results for this range in the dataset
        If (Not AddStatRow(tdlThis, strFile, intColumn, intMin, intMax, strName, intRng, intLen, "Interval")) Then Return False
        ' Adapt the total count of elements we had
        intCount += intRng
        ' Calculate the next range
        intMax = intMax * intLogBase
        intMin = (intMax \ intLogBase) + 1
      Loop While (intCount < intLen)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/ChainStatFreq error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetIpMatOrSubId
  ' Goal:   Get the @Id number of the IP-MAT or IP-SUB I am part of
  '         Only incorporate IP-SUB, if it is not part of a CP-REL (but of another CP)
  ' History:
  ' 14-09-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetIpMatOrSubId(ByRef ndxThis As XmlNode, ByVal bMatOnly As Boolean) As Integer
    Dim ndxWork As XmlNode  ' Working node
    Dim ndxCp As XmlNode    ' The parent CP of a subordinate clause
    Dim strLabel As String  ' The label of the node we are now considering

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return -1
      If (ndxThis.Name <> "eTree") Then Return -1
      ' Go upwards one by one
      ndxWork = ndxThis.ParentNode
      While (Not ndxWork Is Nothing) AndAlso (ndxWork.Name = "eTree")
        ' Get this one's label
        strLabel = ndxWork.Attributes("Label").Value
        ' Check if this is an IP-MAT
        If (strLabel Like "IP-MAT*") Then
          ' Okay, found my parent
          Return ndxWork.Attributes("Id").Value
        ElseIf (Not bMatOnly) AndAlso (strLabel Like "IP-SUB*") Then
          ' THis is a subordinate clause -- check for the ancestor CP
          ndxCp = ndxWork.SelectSingleNode("./ancestor::eTree[starts-with(@Label, 'CP')]")
          If (ndxCp Is Nothing) Then
            ' We cannot determine what the CP parent of us is--be on the safe side and don't count it in
            ' Stop
          Else
            ' Check the kind of CP parent we have
            If (Not ndxCp.Attributes("Label").Value Like "CP-REL*") Then
              ' This is a subordinate clause, but not a relative clause --> okay
              Return ndxWork.Attributes("Id").Value
            End If
          End If
        End If
        ' Go up one more level
        ndxWork = ndxWork.ParentNode
      End While
      ' Return failure
      Return -1
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/GetIpMatOrSubId error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MakeOneChain
  ' Goal:   See if this is a chain, and add it to [tdlThis]
  ' History:
  ' 10-08-2011  ERK Created
  ' 09-02-2011  ERK Added RootEtreeId in <Chain>
  ' ------------------------------------------------------------------------------------
  Private Function MakeOneChain(ByRef ndxThis As XmlNode, ByRef colDone As NodeColl, _
                              ByRef tdlThis As DataSet) As Boolean
    Dim ndxWork As XmlNode              ' Working node
    Dim ndxForest As XmlNode = Nothing  ' Forest node
    Dim intLen As Integer = 0           ' Length of this chain
    Dim intNoTraceLen As Integer = 0    ' Length of this chain excluding traces
    Dim strLastText As String = ""      ' Last node text
    Dim strRefType As String            ' The kind of reference we have here
    Dim dtrChain As DataRow = Nothing   ' Datarow for this chain
    Dim strAnim As String = ""          ' Animacy of this chain
    Dim intChainId As Integer = 0       ' ID for this chain
    Dim dtrItem As DataRow = Nothing    ' One datarow for an item
    Dim intItemId As Integer = 0        ' ID for this item
    Dim intIPmatId As Integer           ' the @Id of the IP under which we are 
    Dim intIPclsId As Integer           ' the @Id of the IPmat under which we are 
    Dim bStop As Boolean = False        ' Flag to stop the chain

    Try
      ' Validate
      If (ndxThis Is Nothing) OrElse (colDone Is Nothing) Then Return False
      ' Action depends on the reference type of this node\
      strRefType = CorefAttr(ndxThis, "RefType")
      Select Case strRefType
        Case "NewVar", "Inert"
          ' These are purely hypothetical, theoretical nodes, so we exclude them
        Case "New", "Assumed", "Identity", "CrossSpeech", "Inferred"
          ' Check if this node has already been processed
          If (colDone.Exists(ndxThis)) Then Return True
          ' This node starts a chain!
          intChainNum += 1
          ' (1) Start XML output for this chain
          dtrChain = AddOneDataRow(tdlThis, "Chain", "ChainId", "ChainList")
          intChainId = dtrChain.Item("ChainId")
          ' See if we have the animacy feature for this chain
          strAnim = GetFeature(ndxThis, "NP", "anim")
          dtrChain.Item("Animacy") = IIf(strAnim = "", "-", strAnim)
          ' Go into a loop
          ndxWork = ndxThis
          While (Not ((ndxWork Is Nothing) OrElse (colDone.Exists(ndxWork)) OrElse bStop))
            ' Increment the length of this chain
            intLen += 1
            ' Give information about this one
            strLastText = NodeText(ndxWork, True)
            ' Get the forestnode of me
            If (Not GetForestNode(ndxWork, ndxForest)) Then Return False
            ' Get the @Id of the IPmat under which we are 
            intIPclsId = GetIpMatOrSubId(ndxWork, False)
            intIPmatId = GetIpMatOrSubId(ndxWork, True)
            ' Store the XML information
            If (Not CreateNewRow(tdlThis, "Item", "ItemId", intItemId, dtrItem)) Then Return False
            With dtrItem
              .SetParentRow(dtrChain)
              .Item("ChainId") = intChainId
              .Item("eTreeId") = ndxWork.Attributes("Id").Value
              .Item("forestId") = ndxForest.Attributes("forestId").Value
              .Item("IPmatId") = intIPmatId
              .Item("IPclsId") = intIPclsId
              .Item("IPclsLb") = NodeLabel(IdToNode(intIPclsId))
              If (ndxWork.Attributes("IPnum") Is Nothing) Then
                .Item("IPnum") = .Item("forestId")
              Else
                .Item("IPnum") = ndxWork.Attributes("IPnum").Value
              End If
              .Item("SbjRefAll") = "False"
              .Item("SbjRefMat") = "False"
              .Item("SbjRefMat3") = "False"
              .Item("Loc") = NodeLocation(ndxWork)
              .Item("RefType") = CorefAttr(ndxWork, "RefType")
              .Item("Syntax") = ndxWork.Attributes("Label").Value
              .Item("NPtype") = GetFeature(ndxWork, "NP", "NPtype")
              .Item("GrRole") = GetFeature(ndxWork, "NP", "GrRole")
              .Item("Node") = strLastText
              ' Set initial (default) values of distances
              .Item("DistIPnum") = 0 : .Item("DistForest") = 0
              ' Set PGN value, if known
              .Item("PGN") = GetFeature(ndxWork, "NP", "PGN")
              ' Adapt the NoTraceLen
              If (.Item("NPtype") <> "Trace") Then intNoTraceLen += 1
            End With
            ' File this node
            colDone.AddUnique(ndxWork)
            ' Note our current Reference type
            strRefType = CorefAttr(ndxWork, "RefType")
            ' Calculate whether we need to stop
            bStop = (Not DoLike(strRefType, "Identity|CrossSpeech"))
            ' Check if the Chain number is present as feature
            If (HasFeature(ndxWork, "coref", "ChainId")) Then
              ' Check the value of the feature
              If (GetFeature(ndxWork, "coref", "ChainId") <> intChainId) Then
                ' Change the value of the feature
                If (Not AddFeature(pdxCurrentFile, ndxWork, "coref", "ChainId", intChainId)) Then Return False
                ' Indicate we have changed
                bNeedSaving = True
              End If
            Else
              ' Add the chain number as a Coref type feature
              If (Not AddFeature(pdxCurrentFile, ndxWork, "coref", "ChainId", intChainId)) Then Return False
              ' Indicate we have changed
              bNeedSaving = True
            End If
            ' Go to the next antecedent
            ndxWork = CorefDst(ndxWork)
          End While
          ' Add the length of this chain
          dtrChain.Item("Len") = intLen
          ' Add the length of this chain excluding traces
          dtrChain.Item("NoTraceLen") = intNoTraceLen
          ' Add the chain's root
          ' ======== DEBUG ==========
          ' If (dtrChain.Item("ChainId") = 13) Then Stop
          ' =========================
          ndxWork = GetRootNode(ndxThis)
          If (ndxWork Is Nothing) Then
            dtrChain.Item("RootEtreeId") = -1
          Else
            dtrChain.Item("RootEtreeId") = ndxWork.Attributes("Id").Value
          End If
        Case ""
          ' No need to take action!!
        Case Else
          ' This should not occur, but may happen with other files -- just signal this
          Logging("modMain/MakeOneChain unknown reference type = " & strRefType & " in: " & _
                  IO.Path.GetFileNameWithoutExtension(GetTableSetting(tdlThis, "File")))
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/MakeOneChain error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ShowOneChain
  ' Goal:   Put a description of this chain into [colThis], and keep track of the [colIndex] too
  ' History:
  ' 10-08-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function ShowOneChain(ByRef tdlThis As DataSet, ByVal intChainId As Integer, ByRef colThis As StringColl, _
                                ByRef colIndex As StringColl) As Boolean
    Dim intLen As Integer = 0       ' Length of this chain
    Dim colTmp As New StringColl    ' Temporary collection
    Dim strLastText As String = ""  ' Last node text
    Dim dtrItem() As DataRow          ' One datarow for an item
    Dim intI As Integer             ' Counter

    Try
      ' Validate
      If (tdlThis Is Nothing) OrElse (colThis Is Nothing) Then Return False
      ' (1) Start html output for this chain
      colTmp.Add("<tr><td nowrap>Loc:</td><td nowrap>RefType</td>" & _
                  "<td nowrap>Syntax</td><td nowrap>NPtype</td><td nowrap>GrRole</td><td nowrap>Node</td></tr>")
      ' Walk through all items of this chain
      dtrItem = tdlThis.Tables("Item").Select("ChainId=" & intChainId)
      ' Get my length
      intLen = dtrItem.Length
      For intI = 0 To intLen - 1
        ' Process this item
        With dtrItem(intI)
          ' Keep track of the last text
          strLastText = .Item("Node")
          ' Store the HTML information
          colTmp.Add("<tr><td nowrap>" & .Item("Loc") & "</td>" & _
                      "<td nowrap>" & .Item("RefType") & "</td>" & _
                      "<td nowrap>" & .Item("Syntax") & "</td>" & _
                      "<td nowrap>" & .Item("NPtype") & "</td>" & _
                      "<td nowrap>" & .Item("GrRole") & "</td>" & _
                      "<td>" & strLastText & "</td>")
        End With
        ' Are we referring further?
        If (intI >= intLen - 1) Then
          colTmp.Add("<td></td></tr>")
        Else
          colTmp.Add("<td nowrap>points to...</td></tr>")
        End If
      Next intI
      ' Finish this table
      colTmp.Add("</table>")
      ' Complete the output for this chain
      colThis.Add("<p><a name='" & GetAnchor(intChainId) & "'><b>Antecedent chain #" & intChainId & _
                  " </b></a> (len=" & intLen & ")")
      ' Should we add other numbers?
      If (intLen > 1) Then
        dtrItem = tdlThis.Tables("Chain").Select("ChainId=" & intChainId)
        ' Are there any results?
        If (dtrItem.Length > 0) Then
          ' Output these results
          colThis.Add(" ProtIsSbj=" & dtrItem(0).Item("SbjIsProt") & _
                      " ProtSbjSw=" & dtrItem(0).Item("SbjSwProtX") & _
                      " NumSbj=" & dtrItem(0).Item("SbjCount") & _
                      " ProAll=" & dtrItem(0).Item("ProAll") & _
                      " ProSbj=" & dtrItem(0).Item("ProSbj") & _
                      " NmeSbj=" & dtrItem(0).Item("NmeSbj") & _
                      " SbjZero=" & dtrItem(0).Item("SbjZero"))
        End If
      End If
      colThis.Add("<br><table>")
      colThis.Add(colTmp.Text)
      ' Add an index for this particular chain
      colIndex.Add(intLen, "<a href='#" & GetAnchor(intChainId) & "'>" & intChainId & "</a> (" & strLastText & ")")
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/ShowOneChain error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetAnchor
  ' Goal:   Return a unique anchor for this number
  ' History:
  ' 09-08-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetAnchor(ByVal intNum As Integer) As String
    Try
      Return "a" & Format(intNum, "0000")
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/GetAnchor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  GetName
  ' Goal :  Get a name for the new element you want to add
  ' History:
  ' 02-06-2010  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function GetName(ByVal strQuestion As String) As String
    Dim strName As String = ""  ' The name to be returned
    Try
      ' Get a name for the query
      With frmGetName
        ' Set the caption
        .Text = strQuestion
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.Cancel
            ' Don't continue
            Return ""
          Case Windows.Forms.DialogResult.OK
            ' Retrieve the query name etc.
            strName = .ElementName
        End Select
      End With
      ' Do trim the name...
      strName = Trim(strName)
      ' Return this name
      Return strName
    Catch ex As Exception
      ' Warn the user
      HandleErr("modMain/GetName error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   RefChainSave
  ' Goal:   Save the referential chains to the proper place
  ' History:
  ' 09-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function RefChainSave() As Boolean
    Dim strFile As String   ' Current file's name
    Dim strOutXml As String ' Target to save

    Try
      ' Get the current file's name
      strFile = strCurrentFile
      ' Validate
      If (strFile = "") Then Return False
      If (tdlRefChain Is Nothing) Then Return False
      ' Calculate the XML file for the chain information
      strOutXml = IO.Path.GetDirectoryName(strFile) & "\" & _
        IO.Path.GetFileNameWithoutExtension(strFile) & "_Chains.xml"
      ' Store the XML output file
      tdlRefChain.WriteXml(strOutXml)
      ' Show where we save it
      Status("Saved: " & strOutXml)
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modMain/RefChainSave error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetRefChain
  ' Goal:   Get all the elements in the referential chain with [@intChainId]
  ' History:
  ' 10-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetRefChain(ByVal intChainId As Integer) As String
    Dim strRoot As String = ""    ' Text of rootnode
    Dim colThis As New StringColl ' Collection of node information
    Dim dtrFound() As DataRow     ' Result of select
    Dim ndxRoot As XmlNode        ' Chain root
    Dim intI As Integer           ' Counter
    Dim bStop As Boolean = False  ' Stop flag

    Try
      ' Validate
      If (tdlRefChain Is Nothing) Then Return ""
      If (intChainId >= 0) Then
        ' Select items from this chain
        dtrFound = tdlRefChain.Tables("Item").Select("ChainId=" & intChainId, "ItemId ASC")
        ' Do calculate the chain root
        intI = dtrFound(dtrFound.Length - 1).Item("eTreeid")
        If (intI >= 0) Then
          ndxRoot = IdToNode(intI)
          colThis.Add("ROOT--[" & NodeText(GetRootNode(ndxRoot)) & "]")
        End If
        ' Walk all the elements
        For intI = 0 To dtrFound.Length - 1
          With dtrFound(intI)
            ' Add this node's information
            colThis.Add(.Item("Loc") & " [" & .Item("Syntax") & " " & .Item("Node") & "]")
          End With
        Next intI
        ' Store the chain text
        strRoot = colThis.Text
      End If
      ' Return the results
      Return strRoot
    Catch ex As Exception
      ' Warn the user
      HandleErr("modMain/GetRefChain error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetRefChain
  ' Goal:   Get all the elements in the referential chain with [@intChainId]
  ' History:
  ' 10-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function InitXresult() As Boolean
    Try
      ' Validate
      If (bInitXresult) Then Return True
      ' Create a dataset

      ' Set the init flag
      bInitXresult = True
      ' Return positively
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modMain/InitXresult error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitChart
  ' Goal:   Initialise creation of a chart
  ' History:
  ' 26-09-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function InitChart() As Boolean
    Try
      ' Check if the dataset already exists
      If (tdlChart Is Nothing) Then
        ' Create a dataset
        If (Not CreateDataSet("Chart.xsd", tdlChart)) Then Return False
      Else
        ' Empty the rows of constituents and clauses
        ClearTable(tdlChart.Tables("Const"))
        ClearTable(tdlChart.Tables("Clause"))
      End If
      ' return positively
      Return True
    Catch ex As Exception
      ' Show user that there is an error
      HandleErr("modMain/InitChart error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   FindControl
  ' Goal:   Find a control with name [strName] hierarchically within [ctlThis]
  ' History:
  ' 18-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function FindControl(ByRef ctlThis As Control, ByVal strName As String) As Control
    Dim ctlWork As Control  ' Working control
    Dim ctlTry As Control   ' An attempt

    Try
      ' Look at this level
      For Each ctlWork In ctlThis.Controls
        ' ========== DEBUG ========
        ' Debug.Print("FindControl -- checking [" & ctlWork.Name & "] against [" & strName & "]")
        ' =========================
        ' Check this one's name
        If (strName = ctlWork.Name) Then Return ctlWork
        ' Check this one's children
        If (ctlThis.HasChildren) Then
          ctlTry = FindControl(ctlWork, strName)
          If (ctlTry IsNot Nothing) Then Return ctlTry
        End If
      Next ctlWork
      ' Return empty
      Return Nothing
    Catch ex As Exception
      ' Show user that there is an error
      HandleErr("modMain/FindControl error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       XmlFileType()
  ' Goal:       Check the XML file type by looking for the first two lines
  ' History:
  ' 22-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function XmlFileType(ByVal strFile As String) As String
    Dim strCheck As String          ' String
    Dim rdThis As IO.StreamReader
    Dim intPos As Integer           ' Position in string
    Dim intPos2 As Integer          ' Closing

    Try
      ' Validate: file exists?
      If (Not IO.File.Exists(strFile)) Then Return ""
      ' Validate: read first two lines and look for <TEI> line
      rdThis = New IO.StreamReader(strFile)
      strCheck = rdThis.ReadLine()
      ' Is this a <xml> line?
      If (InStr(strCheck, "<?xml") > 0) Then
        ' Read one more line
        strCheck = rdThis.ReadLine
      End If
      rdThis.Close()
      ' Get the XML code from here
      intPos = InStr(strCheck, "<")
      If (intPos = 0) Then Return ""
      intPos2 = InStr(strCheck, ">")
      If (intPos2 = 0) Then Return ""
      Return Mid(strCheck, intPos + 1, intPos2 - intPos - 1)
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/XmlFileType error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return ""
    End Try
  End Function

End Module
