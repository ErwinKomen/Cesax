Imports System.Xml
Module modAuto
  ' ========================================= LOCAL VARIABLES ===========================================
  Private loc_ndxAntec As XmlNodeList         ' All possible antecedents
  Private loc_pdxIp As XmlNodeList            ' Selection of all IP's
  Private loc_colIPstack As New StringColl    ' The IP stack
  Private loc_tblDstStack() As DataTable      ' A stack of destination tables
  Private loc_intAskWidth As Integer = 0      ' The width of the Ask form
  Private loc_intAskHeight As Integer = 0     ' The heigth of the Ask form
  Private loc_intSelNodeId As Integer = -1    ' ID of the selected node
  Private loc_intMult As Integer = 1          ' Multiplication factor for constraint evaluation
  Private loc_intAntSize As Integer = 50      ' Maximum number of antecedents in stack
  Private loc_intMaxIPdist As Integer = 15    ' Maximum IP distance we allow
  Private loc_intAbsMaxIPdist As Integer = 100 ' Absolut maximum IP distance, after which we cut off
  Private loc_strLastCns As String = ""       ' The last constraint that was used
  Private loc_dblNPcount As Integer = 0       ' Number of NPs treated
  Private loc_strPgnList As String = ""       ' List of PGN values
  Private loc_bLocalRefl As Boolean = False   ' Local type: reflexive
  Private loc_bLocalApp As Boolean = False    ' Local type: Appositive
  Private loc_bLocalBare As Boolean = False   ' Local type: bare
  Private tblAction As DataTable = Nothing    ' Table containing essential information about actions taken
  Private tblSrcStack As DataTable = Nothing  ' Stack of constituents that could be sources of coreference
  Private tblAntStack As DataTable = Nothing  ' Stack of potential antecedent constituents
  Private loc_colNPstack As New StringColl    ' All the <eTree> ID's of those we've had
  'Private arConstraint() As String = _
  '  {"AgrGenderNumber", "Disjoint", "EqualHead", "AgrGenderSrc", "IPdist", "Ambiguity", _
  '   "GrRoleDst", "NPtypeDst", "NoCrossAgrPerson", "NoCrossEqSubject", "AgrPerson", "ProTop", "FamDef", "Cohere"}
  Private tblConstraint As DataTable = Nothing  ' Table with constraints
  Private strAutoStatScheme As String = "AutoStat.xsd"
  Private strHtmlHead As String = "<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8' /></head>"
  Private colCurrentRefLearnTest As New StringColl  ' Currently being trained STATE
  Private colCurrentRefLearnTrain As New StringColl ' Currently being trained STATE
  Private colCurrentTrainingState As New StringColl ' Currently being trained STATE
  Private colCurrentTrainingCoref As New StringColl ' Currently being trained COREF
  Private colHeadBefore As New StringColl           ' Heads that have occurred before me
  Private arFnameSmall() As String = {"Label", "HeadCat", "Anchor", "PostModCat", "NPtype", "HeadBefore"}
  Private arFnameLarge() As String = {"Label", "NPtype", "GrRole", "Starred", "HasFreeRel", "HasCPrel", "Ch1_Label", "Ch1_WordType", "Ch1WordText", _
                                      "Ch2_Label", "Ch3_Label", "Ch4_Label", "BeSister", "SbjSister", "VbSister", "CpSister", "SbjType", _
                                      "SbjText", "HeadBefore", "HeadText", "PGN", "Len", "HasNeg", "Mood", "IsSpeech", "HasPremod"}
  ' ========================================= GLOBAL VARIABLES ==========================================
  Public strNounTypes As String = "N|NS|NR|NRS|NPR|NPRS|N^*|NS^*|NR^*|NRS^*|N$|NS$|NR$|NRS$|NPR$|NPRS$|*+N$|*+NS$|*+NR$|*+NRS$|*+NPR$|*+NPRS$|*+NS|*+N"
  Public strNPsourceTypes As String = "NP|NP-*|PRO[$]*|NPR$*|WNP*"
  Public strNPnoSourceTypes As String = "NP-MSR*|NP-ADV*"
  Public strNPsourceTypesOrg As String = "@Label='NP' or starts-with(@Label,'NP-') " & _
    "or starts-with(@Label,'PRO$') or starts-with(@Label, 'NPR$')"
  Public strIPtypes As String = "QTP|IP-MAT|IP-MAT-*|IP-ABS|IP-ABS-*|IP-IMP|IP-IMP-*|IP-INF|IP-INF-*|IP-SUB|IP-SUB[-=]*|IP-SMC*|IP-PPL*"
  Public strIPstackType As String = "QTP|IP-MAT|IP-MAT-*|IP-ABS|IP-ABS-*|IP-IMP|IP-IMP-*|IP-SUB|IP-SUB[-=]*|IP-PPL*"
  Public strNounTypesTb As String = "tb:Like(string(@Label), '" & strNounTypes & "')"
  Public strHdNounTypesTb As String = "tb:Like(string(@Label), '" & strNounTypes & "|ADJ*|Q|Q-*|FW" & "')"
  Public strNPsourceTypesTb As String = "tb:Like(string(@Label), '" & strNPsourceTypes & "')"
  Public strIPtypesTb As String = "tb:Like(string(@Label), '" & strIPtypes & "')"
  Public strFiniteBe As String = "BEP*|BED*|NEG+BEP*|NEG+BED*"
  Public strVerbBe As String = "BE*|BA*|*+BE*|*+BA*"
  Public strVerbHv As String = "HV*|*+HV*"
  Public strVerbDo As String = "DO*|*+DO*|DA[GN]*|*+DA[GN]*"
  'Public strFiniteAux As String = "BE[IPD]*|UTP|*HV[IPD]*|*AX[IPD]*|*MD|*DO[IPD]*|" & _
  '  "NEG+B[IPD]*|NEG+AX[IPD]*|NEG+*MD"
  Public strFiniteAux As String = "BE[IPD]*|UTP|*HV[IPD]*|*AX[IPD]*|*DO[IPD]*|" & _
    "NEG+B[IPD]*|NEG+AX[IPD]*"
  Public strFiniteVerb As String = "BEI|BEP*|BED*|UTP|*HVI|*HVP*|*HVD*|*AXI|*AXP*|*AXD*|*MD|*MDP*|*MDD*|VBI|" & _
    "*VBP*|*VBD*|*DOI|*DOP*|*DOD*|NEG+BEI|NEG+BEP*|NEG+BED*|NEG+AXI|NEG+*AXP*|NEG+*AXD*|NEG+*MD|" & _
    "NEG+MDP*|NEG+VBI|NEG+*VBP*|NEG+*VBD"
  Public strAnyVerb As String = "V[AB]*|B[AE]*|HV*|AX*|MD*|D[AO]*|*+V[AB]*|*+B[AE]*|*+HV*|*+AX*|*+MD*|*+D[OA]*"
  Public strVerbAffirm As String = "VA*|VB*|BA*|BE*|HV*|AX*|MD*|DO*|DA*"
  Public strVerbImv As String = "AXI*|*HVI*|*BEI*|*VBI*"
  Public strVerbSubj As String = "*VBDS*|*VBPS*"
  Public strHumanAnt As String = "human|man|men|person|people"
  Public strSpeakIntro As String = "ask*|say*|sai*|quo*|answ*"
  Public strNearDem As String = "*is*|*ys*|*as*|*es*|*os*"
  Public strSubjectType As String = "NP*SBJ*|NP*NOM|NP-NOM-RSP"
  Public strVariableCP As String = "CP-FRL*|CP-REL*|CP-QUE*"
  Public strSmartCP As String = "CP-ADV*|CP-*LFD"   ' CP types that are immediately recognized as Left-First
  Public strSmartXP As String = "PP|PP-*|ADVP*"     ' XP types that are candidates for Left-First
  Public strSkipTypes As String = "CODE|,|.|META*|FW"
  Public strLexicalNPhead As String = "N*|PRO*"
  Public bIsFullyAuto As Boolean = False            ' Whether we are doing fully automatic coreferencing
  Public bDoOverwrite As Boolean = False            ' Whether existing output may be overwritten
  Public bMaxEntTraining As Boolean = False         ' Are we training for maximum entropy?
  Public bMaxEntUsing As Boolean = False            ' Are we using maximum entropy for resolution?
  Public tdlAutoStat As DataSet = Nothing           ' Statistics of fully automatic processing
  Public intCnsSet As Integer = 0                   ' The set of constraints we are using
  Public colIPpos As New StringColl                 ' Collection of IP positions
  Public bTraining As Boolean = False               ' Are we training or not?
  Public strCurrentTrainingCoref As String = ""     ' Name of the currently being trained COREF file
  Public strCurrentTrainingState As String = ""     ' Name of the currently being trained STATE file
  Public strCurrentTrainingRefLearn As String = ""  ' Name of the currently being trained STATE file
  Public colCrpResults As New StringColl            ' Collection holding [tdlResults]
  Public intCrpResultsId As Integer                 ' Global results ID counter
  ' =====================================================================================================
  ' ------------------------------------------------------------------------------------
  ' Name:   GetNPcount
  ' Goal:   Return the total number of NP's treated for this file
  ' History:
  ' 14-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetNPcount() As Integer
    Return loc_dblNPcount
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   OneFullAutoFile
  ' Goal:   Perform fully automatic coreferencing from [strInFile]
  '           to the resulting [strOutFile]
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function OneFullAutoFile(ByVal strInFile As String, ByVal strOutFile As String) As Boolean
    Dim intSection As Integer     ' Section counter
    Dim strToronto As String      ' Toronto ID number
    Dim strStatFile As String     ' Name of the statistics file

    Try
      ' Validate
      If (Not IO.File.Exists(strInFile)) Then
        Status("Could not find file " & strInFile)
        Return False
      End If
      ' Show we are reading
      Status("Reading file " & strInFile)
      ' ============= PROFILING =============
      ProfileStart("OneFullAutoFile", 1)
      ' =====================================
      ' Try read this file into an XML structure
      If (ReadXmlDoc(strInFile, pdxCurrentFile)) Then
        ' Make sure the flag indicates that we are doing fully automatic coreferencing
        bIsFullyAuto = True
        ' Make sure the correct file is set
        strCurrentFile = strInFile
        ' Get the correct period of this file
        strCurrentPeriod = GetPeriod(pdxCurrentFile)
        ' Try initialize
        If (Not InitCurrentFile()) Then Return False
        ' Set the currentfile
        strCurrentFile = strInFile
        ' Delete sections if this is YCOE -- somehow they got muddled up...
        If (IsYcoe(pdxCurrentFile)) Then DelAllSections()
        ' Insert a section here, if it is not there already
        If (Not HasStartSection()) Then
          ' Try to derive sections by looking at (CODE <heading>) entries
          ' For YCOE we look at (CODE <Tnnnnn...>) entries, and when there is a 
          '   change in "nnnnn", then this means a change in Toronto file, and 
          '   we interpret this to be a section break in the text
          If (AutoAddSections()) Then
            ' Show what has happened
            If (IsYcoe(pdxCurrentFile)) Then
              Logging("Automatically made " & pdxSection.Count & " section(s) based on Toronto information.")
            Else
              Logging("Automatically made " & pdxSection.Count & " section(s) based on <heading> entries.")
            End If
          Else
            If (AddStartSection()) Then
              ' Show what is happening
              Logging("There were no sections, so I added section <1>")
            End If
          End If
          ' Add changes to the input file
          pdxCurrentFile.Save(strInFile)
        End If
        ' Initialise the NP count
        loc_dblNPcount = 0
        ' Calculate which sections there are
        Status("Calculating sections...")
        SectionCalculate()
        ' Walk through all sections
        For intSection = 0 To pdxSection.Count - 1
          ' Initialise the statistics report for this file/section combination
          strStatFile = AutoStatInit(strInFile, intSection)
          ' Go to this section
          If (Not GotoSection(intSection)) Then Return False
          ' Perform autocoreferencing (including adding NP features if needed)
          If (Not DoAutoCoref()) Then Return False
          ' Write the result to the destination
          Status("Saving results of section " & intSection & " from " & IO.Path.GetFileNameWithoutExtension(strOutFile))
          pdxCurrentFile.Save(strOutFile)
          Status("Saving statistics")
          If (Not AutoStatFinish(strStatFile)) Then
            Logging("Could not save autostat report at " & strStatFile)
          End If
          ' Possibly get toronto ID
          strToronto = GetTorontoId(pdxSection(intSection).Item("eTree"))
          ' Show what we have done
          If (strToronto = "") Then
            Logging("   Section " & intSection + 1 & " of " & pdxSection.Count)
          Else
            Logging("   Section " & intSection + 1 & " of " & pdxSection.Count & " toronto: " & strToronto & ">")
          End If
        Next intSection
        ' Store this revision in the PSDX file
        AddRevDesc(pdxCurrentFile, strUserName, Format(Now, "g"), "Added with File/FullAuto")
        ' Save the final result
        pdxCurrentFile.Save(strOutFile)
        ' Switch off the fully automatic flag
        bIsFullyAuto = False
        ' ============= PROFILING =============
        ProfileEnd("OneFullAutoFile", 1)
        ' =====================================
        ' Return success
        Return True
      Else
        ' Could not read the file, so return failure
        Status("Could not read the XML file: " & strInFile)
        ' Switch off the fully automatic flag
        bIsFullyAuto = False
        ' ============= PROFILING =============
        ProfileEnd("OneFullAutoFile")
        ' =====================================
        Return False
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OneFullAutoFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Switch off the fully automatic flag
      bIsFullyAuto = False
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DoAutoCoref
  ' Goal:   Automatically resolve coreferences as far as possible
  '         This is only done within [pdxList], starting from [intSectFirst] until [intSectLast]
  ' History:
  ' 26-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function DoAutoCoref() As Boolean
    Dim ndxForest As XmlNode    ' The forest we are currently examining
    Dim strBack As String = ""  ' The answer...
    ' Dim arFeatDef(0) As String  ' Feature definitions dummy
    Dim intLine As Integer      ' The line within [pdxList]
    Dim intPtc As Integer       ' Percentage
    Dim bNeedsNP As Boolean     ' Whether NP features need to be done

    Try
      ' Validate
      If (pdxList Is Nothing) OrElse (intSectFirst < 0) OrElse (intSectLast >= pdxList.Count) Then Return False
      ' Another validation: check whether this file has been provided with NP features
      If (pdxCurrentFile.SelectSingleNode("//fs[@type='NP']") Is Nothing) Then
        ' Do we need to ask?
        If (bIsFullyAuto) Then
          bNeedsNP = True
        Else
          ' We need to ask
          ' This is not the correct file!
          Select Case MsgBox("The file you are trying to process has NOT been supplied with NP features." & _
                             vbCrLf & "Would you like to add these now?", MsgBoxStyle.YesNoCancel)
            Case MsgBoxResult.Yes
              ' Yes we should
              bNeedsNP = True
            Case MsgBoxResult.No, MsgBoxResult.Cancel
              ' Return failure
              Return False
          End Select
        End If
        ' Do we need to do NP features?
        If (bNeedsNP) Then
          ' Initialise missing pronoun report
          InitMissing()
          ' Initialise pronoun tracking
          InitCatTrack()
          ' Try to get the NP features
          If (Not OneDoFeaturesPdx(pdxCurrentFile, CAT_NP)) Then
            ' Tell the user
            MsgBox("NP features have, unfortunately, not been determined.")
            Return False
          End If
          ' Try get the pronoun tracking report
          If (HasCatTrack()) Then
            ' Calculate this report and put it in place
            frmMain.wbReport.DocumentText = CatTrackHtml(CAT_NP)
            '' Switch to this tabpage (unless we are in fully auto mode)
            'If (Not bIsFullyAuto) Then
            '  frmMain.TabControl1.SelectedTab = frmMain.tpReport
            'End If
          End If
          ' Try get the missing pronouns report
          If (HasMissing()) Then
            ' Calculate the report and put it in place
            frmMain.wbError.DocumentText = MissingHtml("NP")
            ' Show there are some errors
            Status("Ready adapting NP features. There are missing pronouns reported in the <Errors> page.")
          Else
            ' Show we are done
            Status("Ready adapting NP features.")
          End If
        End If
      End If
      ' Initialise whatever needs to be done for autocoreferencing...
      loc_colIPstack.Clear()
      tblSrcStack = tdlSettings.Tables("StackEl").Clone
      ClearTable(tblSrcStack)
      tblAntStack = tdlSettings.Tables("StackEl").Clone
      ClearTable(tblAntStack)
      ' Initialise the constraint table
      tblConstraint = tdlSettings.Tables("Constraint")
      ' TODO: INITIALISATIONS
      AutoReportInit()
      ' Check the maximum IP distance
      If (intMaxIPdist < 0) Then
        ' If nothing is specified, we take our default one
        intMaxIPdist = loc_intMaxIPdist
      End If
      ' Go through all the necessary <forest> elements in [pdxList]
      For intLine = intSectFirst To intSectLast
        ' Show where we are (depending on the mode we are in)
        intPtc = (100 * (intLine - intSectFirst + 1)) \ (intSectLast - intSectFirst + 1)
        If (bIsFullyAuto) Then
          ' ============== DEBUGGING =============
          'If (intPtc > 100) Then Stop
          ' ======================================
          Status("FullyAuto section " & intCurrentSection + 1 & "/" & pdxSection.Count & _
                 " " & intPtc & "%", intPtc)
        Else
          ' Do give a status overview if possible!
          Status("Semi-automatic " & intPtc & "%", intPtc)
          CorefProgress("stack=" & tblAntStack.Rows.Count)
        End If
        ' Get this line's forest
        ndxForest = pdxList(intLine)
        ' Show the corresponding PDE translation
        ' ============== DEBUGGING =============
        ' If (Not ndxForest.SelectSingleNode("./descendant::eTree[@Id='4615']") Is Nothing) Then Stop
        ' ======================================

        ' Go through the antecedent stack to take off the "IsLocal" marked constituents
        If (Not RemoveLocalAntecedents(ndxForest)) Then Exit For
        ' Prepare IP position measures
        colIPpos.Clear()
        If (Not PrepareForest(ndxForest)) Then
          CorefProgress("")
          Return False
        End If
        ' Clear the source collection table
        ClearTable(tblSrcStack)
        ' Process discourse NEW, appositive and inert elements
        '   (Possibly put them into the antecedent's collection)
        If (Not TraverseNodes(ndxForest, "NewAndLocal", strBack)) Then
          CorefProgress("")
          Return False
        End If
        ' ============== DEBUG ========
        ' Initialise the NP collection
        loc_colNPstack.Clear()
        ' =============================
        ' ============== DEBUG ========
        ' If (ndxForest.Attributes("forestId").Value = 71) Then Stop
        ' =============================
        ' Walk this forest in chunks (smart)
        If (Not ChunkWalk(ndxForest)) Then
          ' Allow interrupting
          Application.DoEvents()
          ' Check if interrupt was pressed
          If (bInterrupt) Then
            CorefProgress("")
            Return False
          End If
        End If
        ' ============== DEBUG ========
        ' Check if all NPs of [ndxForest] have been treated
        strBack = ""
        If (Not CheckForest(ndxForest, strBack)) Then
          ' Try it one more time
          If (Not ChunkWalk(ndxForest)) Then
            ' Allow interrupting
            Application.DoEvents()
            ' Check if interrupt was pressed
            If (bInterrupt) Then
              CorefProgress("")
              Return False
            End If
          End If
          ' Test it one more time
          If (Not CheckForest(ndxForest, strBack)) Then
            ' Are we in fully auto mode?
            If (Not bIsFullyAuto) Then
              ' Warn the user
              MsgBox("DoAutoCoref: The following NPs have NOT been dealt with: " & strBack & vbCrLf & _
                     "forestId = " & ndxForest.Attributes("forestId").Value & vbCrLf & _
                     "Fully auto processing will now be interrupted. Sorry...")
              ' Make sure we stop
              bInterrupt = True
              Return False
            End If
          End If
        End If
      Next intLine
      ' Release the source stack
      tblSrcStack = Nothing
      ' Clear the progress
      CorefProgress("")
      ' TODO: any clearing up
      AutoReportFinish()
      ' Default: return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/DoAutoCoref error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CheckForest
  ' Goal:   Check if all NP's from the forest have been dealt with
  ' History:
  ' 01-10-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CheckForest(ByRef ndxForest As XmlNode, ByRef strBack As String) As Boolean
    Dim ndxNP As XmlNodeList  ' Result of select query
    Dim intId As Integer      ' Id of <eTree> element
    Dim intI As Integer       ' COunter
    Dim bOkay As Boolean = True ' Initially we think we are okay

    Try
      ' Validate
      If (ndxForest Is Nothing) Then Return False
      ' Select all relevant NPs
      ndxNP = ndxForest.SelectNodes(".//eTree[" & strNPsourceTypesTb & "]", conTb)
      ' Check them all
      For intI = 0 To ndxNP.Count - 1
        ' Should we check this one?
        If (Not DoLike(ndxNP(intI).Attributes("Label").Value, strNPnoSourceTypes)) Then
          ' Get the Id
          intId = ndxNP(intI).Attributes("Id").Value
          ' Verify if the Id is present
          If (loc_colNPstack.Find(CStr(intId)) < 0) Then
            ' It is not present, so add to the [strBack] stack
            AddSemiStack(strBack, CStr(intId))
            ' Make sure we are not okay
            bOkay = False
          End If
        End If
      Next intI
      ' Return the state of [bOkay]
      Return bOkay
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CheckForest error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   RemoveLocalAntecedents
  ' Goal:   Remove the antecedents from the stack that are marked as "IsLocal"
  '         Exception: if the next <forest> (which is in [ndxThis]) starts with
  '           a CONJ element, then don't.
  ' History:
  ' 08-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function RemoveLocalAntecedents(ByRef ndxThis As XmlNode) As Boolean
    Dim intI As Integer     ' Counter
    Dim dtrThis As DataRow  ' One element
    Dim ndxNext As XmlNode  ' First <eTree> child

    Try
      ' Validate
      If (tblAntStack Is Nothing) Then Return False
      ' Check next forest
      If (Not ndxThis Is Nothing) Then
        ' There is a next forest, so check the first <eTree> child
        ndxNext = ndxThis.SelectSingleNode("./eTree[not(@Label='CODE')]")
        If (Not ndxNext Is Nothing) Then
          ' The first child should be an IP*
          If (ndxNext.Attributes("Label").Value Like "IP*") Then
            ' Now get the first child of this IP
            ndxNext = ndxNext.SelectSingleNode("./eTree[not(@Label='CODE')]")
            ' Got something
            If (Not ndxNext Is Nothing) Then
              ' Is it a CONJ?
              If (ndxNext.Attributes("Label").Value = "CONJ") Then
                ' The next forest starts with a CONJ, so don't remove local antecedents yet
                Return True
              End If
            End If
          End If
        End If
      End If
      ' Make sure everything is updated
      For intI = tblAntStack.Rows.Count - 1 To 0 Step -1
        ' Get this one
        dtrThis = tblAntStack.Rows(intI)
        ' Validate
        If (Not dtrThis Is Nothing) Then
          ' Look at this one
          If (dtrThis.Item("IsLocal") = "True") Then
            ' ============== DEBUGGING =============
            ' If (dtrThis.Item("Id") = 4615) OrElse ((dtrThis.Item("Id") = 4617)) Then Stop
            ' ======================================
            dtrThis.Delete()
          End If
        End If
      Next intI
      ' Return succes
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/RemoveLocalAntecedents error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AutoStatInit
  ' Goal:   Initialise the table that will record the statistics for fully automatic processing
  ' History:
  ' 13-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AutoStatInit(ByVal strFile As String, ByVal intSection As Integer, _
                               Optional ByVal bNewStatList As Boolean = True) As String
    Dim strStatFile As String   ' Where to load statistics from/to
    Dim strName As String       ' Name of file without extensions
    Dim strDirStat As String    ' Directory for statistics
    Dim dtrThis As DataRow      ' One <StatList> row

    Try
      ' (1) Get the name and the statistics directory
      strName = IO.Path.GetFileNameWithoutExtension(strFile)
      ' Directory depends on what we're doing: fully auto or not
      If (bIsFullyAuto) Then
        strDirStat = IO.Path.GetDirectoryName(strFile) & "\FullAuto\stat"
      Else
        strDirStat = IO.Path.GetDirectoryName(strFile) & "\stat"
      End If
      If (Not IO.Directory.Exists(strDirStat)) Then IO.Directory.CreateDirectory(strDirStat)
      ' Construct the Autostat file
      strStatFile = strDirStat & "\" & strName & "-" & intSection + 1 & "-stat.xml"
      ' Check if this file already exists
      If (IO.File.Exists(strStatFile)) Then
        ' It already exists, so load the already existing information
        If (Not ReadDataset(strAutoStatScheme, strStatFile, tdlAutoStat)) Then Return ""
      Else
        ' The dataset does not yet exist, so make it
        ' Does the dataset exist already?
        If (Not tdlAutoStat Is Nothing) Then
          ' Delete the previous one
          tdlAutoStat = Nothing
        End If
        ' Create a new dataset
        If (Not CreateDataSet(strAutoStatScheme, tdlAutoStat)) Then
          ' Something went wrong
          MsgBox("modAuto/AutoStatInit fatal error: unable to create AutoStat dataset")
          Return ""
        End If
      End If
      ' Are we supposed to create a new row?
      If bNewStatList Then
        ' Create a <StatList> row with the correct filename
        dtrThis = tdlAutoStat.Tables("StatList").NewRow
        With dtrThis
          .Item("File") = strFile
          .Item("NPcount") = 0
        End With
        tdlAutoStat.Tables("StatList").Rows.Add(dtrThis)
      End If
      ' Reset the counting of NPs
      loc_dblNPcount = 0
      ' Return with the correct filename
      Return strStatFile
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AutoStatInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AutoStatAdd
  ' Goal:   Add an entry with the indicated reference type and constraint
  ' History:
  ' 13-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub AutoStatAdd(ByVal intNodeId As Integer, ByVal strInfo As String, ByVal strRefType As String, _
                          ByVal strConstraint As String, ByVal strLinkType As String, ByVal intIPdist As Integer)
    Dim intStatId As Integer = -1     ' ID for current entry
    Dim dtrItem As DataRow = Nothing  ' New element
    Dim dtrParent As DataRow          ' The parent
    Dim strFile As String             ' Short file name
    Dim bIsNew As Boolean = True      ' Whether we are working with a new row

    Try
      ' Validate
      If (strCurrentFile = "") OrElse (tdlAutoStat Is Nothing) Then Exit Sub
      ' Get the correct parent row
      If (tdlAutoStat.Tables("StatList") Is Nothing) Then
        dtrParent = Nothing
      Else
        ' Get the last <StatList> row
        With tdlAutoStat.Tables("StatList")
          dtrParent = .Rows(.Rows.Count - 1)
        End With
      End If
      ' Get the name of the current file
      strFile = IO.Path.GetFileNameWithoutExtension(strCurrentFile)
      ' Create a new entry
      With tdlAutoStat.Tables("Stat")
        ' Do we already have an entry?
        If (intNodeId >= 0) Then
          If (.Select("NodeId=" & intNodeId).Length > 0) Then
            ' The entry already exists -- this means there is a correction, and it needs to be redone
            dtrItem = .Select("NodeId=" & intNodeId)(0)
            ' Get the correct ID
            intStatId = dtrItem("StatId")
            ' Show this is old
            bIsNew = False
            ' Exit Sub
          End If
        End If
        ' Create a new entry if needed
        If (bIsNew) Then
          intStatId = .Rows.Count + 1
          dtrItem = .NewRow
        End If
        With dtrItem
          .Item("StatId") = intStatId
          ' Possibly add the node id - otherwise leave it empty!!
          If (intNodeId >= 0) Then
            .Item("NodeId") = intNodeId
          End If
          .Item("File") = strFile
          .Item("Info") = strInfo
          .Item("Reftype") = strRefType
          .Item("Constraint") = strConstraint
          .Item("LinkType") = strLinkType
          .Item("IPdist") = intIPdist
          ' Also set the parent row
          .SetParentRow(dtrParent)
        End With
        ' Add this row if it is new
        If (bIsNew) Then .Rows.Add(dtrItem)
      End With
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AutoStatAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   AutoStatFinish
  ' Goal:   Save the autostat report to the indicated file
  ' History:
  ' 13-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AutoStatFinish(ByVal strOutFile As String) As Boolean
    Dim dtrStatList As DataRow  ' The place to put the NPcount

    Try
      ' Validate
      If (tdlAutoStat Is Nothing) Then Return False
      If (strOutFile = "") Then Return False
      ' Get the <StatList> line
      If (tdlAutoStat.Tables("StatList") Is Nothing) Then Return False
      With tdlAutoStat.Tables("StatList")
        ' Get the LAST <StatList> element...
        dtrStatList = .Rows(.Rows.Count - 1)
      End With
      ' Save some more information
      With dtrStatList
        .Item("Date") = Format(Now, "f")
        .Item("User") = strUserName
        .Item("NPcount") = loc_dblNPcount ' The amount of NPs that was put into the source stack
      End With
      ' Save the information
      tdlAutoStat.WriteXml(strOutFile)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AutoStatFinish error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ActionInit
  ' Goal:   Initialise the table that will record the actions
  ' History:
  ' 28-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub ActionInit()
    Try
      ' Was there a previous table?
      If (tblAction Is Nothing) Then
        ' Make the table
        tblAction = New DataTable("Action")
        ' Set columns for this table
        With tblAction
          ' ID within this table of actions
          .Columns.Add("Id", System.Type.GetType("System.Int32"))
          ' Id of the source NP's <eTree>
          .Columns.Add("SrcId", System.Type.GetType("System.Int32"))
          ' Id where the reference goes to (only if it goes somewhere)
          .Columns.Add("RefId", System.Type.GetType("System.Int32"))
          ' Type link made (automatic or user choice)
          .Columns.Add("LinkType", System.Type.GetType("System.String"))
          ' Type of coreference relation (including New, Assumed etc)
          .Columns.Add("RefType", System.Type.GetType("System.String"))
          ' The location of the source's <forest>
          .Columns.Add("Location", System.Type.GetType("System.String"))
          ' The complete source IP, coded in HTML, with the source NP red
          .Columns.Add("SrcIP", System.Type.GetType("System.String"))
          ' The complete IP of the antecedent, coded in HTML, with the antecedent NP in red
          .Columns.Add("AntIP", System.Type.GetType("System.String"))
          ' The complete IP that was rejected, coded in HTML, with the antecedent NP in red
          .Columns.Add("AutoIP", System.Type.GetType("System.String"))
          ' The constraints and their values that made the choice Ant
          .Columns.Add("Constraints", System.Type.GetType("System.String"))
          ' The constraints and their values that made the choice of Auto
          .Columns.Add("AutoCons", System.Type.GetType("System.String"))
          ' Room to report a possible reason for the choice made
          .Columns.Add("Reason", System.Type.GetType("System.String"))
        End With
      Else
        ' There is a previous one already - just clear it
        ClearTable(tblAction)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/ActionInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   ActionAdd
  ' Goal:   Add one line to the table of actions
  ' History:
  ' 28-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub ActionAdd(ByVal intSrcId As Integer, ByVal intRefId As Integer, ByVal strLinkType As String, _
      ByVal strRefType As String, ByVal strLocation As String, ByVal strScrIP As String, ByVal strAntIP As String, _
      ByVal strAutoIP As String, ByVal strConstraints As String, ByVal strAutoCons As String, _
      ByVal strReason As String)
    Dim intMyId As Integer    ' My own ID
    Dim dtrAction As DataRow  ' New row to add

    Try
      ' Validate
      If (tblAction Is Nothing) Then Exit Sub
      ' Determine the new ID for this line
      intMyId = tblAction.Rows.Count + 1
      ' Make a new row
      dtrAction = tblAction.NewRow
      ' Fill the row
      With dtrAction
        .Item("Id") = intMyId
        .Item("SrcId") = intSrcId
        .Item("RefId") = intRefId
        .Item("LinkType") = strLinkType
        .Item("RefType") = strRefType
        .Item("Location") = strLocation
        .Item("SrcIP") = strScrIP
        .Item("AntIP") = strAntIP
        .Item("AutoIP") = strAutoIP
        .Item("Constraints") = strConstraints
        .Item("AutoCons") = strAutoCons
        .Item("Reason") = strReason
      End With
      ' Add the row
      tblAction.Rows.Add(dtrAction)
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/ActionAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   IsAction
  ' Goal:   Tell us whether any action has taken place or not
  ' History:
  ' 16-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function IsAction() As Boolean
    Try
      ' If there is no table, no action took place
      If (tblAction Is Nothing) Then Return False
      ' If there is a table, see how many rows there are
      Return (tblAction.Rows.Count > 0)
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/IsAction error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ActionReport
  ' Goal:   Report on the automatically made coreferences
  '         The output string gives the HTML report
  ' History:
  ' 28-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function ActionReport() As String
    Dim objResult As StringColl ' The result of the report
    Dim dtrCns() As DataRow     ' Ordered set of constraints
    Dim intAuto As Integer = 0  ' Number of automatically made links
    Dim intUser As Integer = 0  ' Number of links made by user
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (tblAction Is Nothing) Then Return "(No results available)"
      If (tblConstraint Is Nothing) Then Return "(No constraint table available)"
      ' Initialise the result
      objResult = New StringColl
      ' Start making the report
      objResult.Add("<html>" & strHtmlHead & "<body><h1>CESAX autocoreference report</h1><p><table>" & _
                    "<tr><td>Time:</td><td>" & Format(Now, "F") & "</td></tr>" & _
                    "<tr><td>Text:</td><td>" & strCurrentFile & "</td></tr>" & _
                    "</table><br> <p>" & vbCrLf)
      ' Give an overview of the constraints
      objResult.Add("<h2>Constraint description</h2>")
      dtrCns = tblConstraint.Select("", "Level ASC, Name ASC")
      objResult.Add("<table><tr><td>Constraint</td><td>Level</td><td>Description</td></tr>")
      For intI = 0 To dtrCns.Length - 1
        ' Add this constraint
        objResult.Add("<tr><td>" & dtrCns(intI).Item("Name") & "</td><td>" & _
                      dtrCns(intI).Item("Level") & "</td><td>" & _
                      dtrCns(intI).Item("Descr") & "</td></tr>")
      Next intI
      ' Finish constraint table
      objResult.Add("</table><p>")
      ' Give an overview of each decision made by the user and the program
      objResult.Add("<h2>Overview of added coreference links</h2>")
      ' Visit all elements of the table inorder
      For intI = 0 To tblAction.Rows.Count - 1
        ' Attend to this element
        With tblAction(intI)
          ' Track statistics
          If (.Item("LinkType") = "Auto") Then intAuto += 1 Else intUser += 1
          ' Output HTML code
          Select Case .Item("RefType")
            Case strRefNew, strRefAssumed, strRefNewVar, strRefInert
              objResult.Add("<table><tr><td>" & .Item("Id") & ":</td><td><font size='1'><b>" & .Item("Location") & _
                            "</b></font></td><td>" & .Item("RefType") & " <i>(" & .Item("LinkType") & ")</i></td></tr>")
              objResult.Add("<tr><td></td><td></td><td>" & CStr(.Item("Constraints")).Replace(";", "; ") & "</td></tr>")
              ' Do we have a reason?
              If (.Item("Reason") & "" <> "") Then
                ' Add the reason
                objResult.Add("<tr><td></td><td></td><td>Reason: " & .Item("Reason") & "</td></tr>")
              End If
              objResult.Add("<tr><td></td><td></td><td>" & .Item("SrcIP") & "</td></tr></table><br> <p>")
            Case strRefIdentity, strRefInferred, strRefCross
              objResult.Add("<table><tr><td>" & .Item("Id") & ":</td><td><font size='1'><b>" & .Item("Location") & _
                            "</b></font></td><td>" & .Item("RefType") & " <i>(" & .Item("LinkType") & ")</i></td></tr>")
              ' Do we have a reason?
              If (.Item("Reason") & "" <> "") Then
                ' Add the reason
                objResult.Add("<tr><td></td><td></td><td>Reason: " & .Item("Reason") & "</td></tr>")
              End If
              ' Do we have a rejected element?
              If (.Item("AutoIP") & "" <> "") Then
                objResult.Add("<tr><td></td><td></td><td>" & CStr(.Item("AutoCons")).Replace(";", "; ") & "</td></tr>")
                objResult.Add("<tr><td></td><td>to-auto:</td><td>" & .Item("AutoIP") & "</td></tr>")
              End If
              objResult.Add("<tr><td></td><td></td><td>" & CStr(.Item("Constraints")).Replace(";", "; ") & "</td></tr>")
              objResult.Add("<tr><td></td><td>to-user:</td><td>" & .Item("AntIP") & "</td></tr>")
              objResult.Add("<tr><td></td><td>from:</td><td>" & .Item("SrcIP") & "</td></tr></table><br> <p>")
            Case Else
              ' SOmething is wrong here!!
              MsgBox("modAuto/ActionReport: unknown RefType " & .Item("RefType"))
          End Select
        End With
      Next intI
      ' Give statistics
      objResult.Add("<table><tr><td><b>Automatically added:</b></td><td>" & intAuto & "</td></tr>" & _
                    "<tr><td><b>User interaction:</b></td><td>" & intUser & "</td></tr>" & _
                    "<tr><td><b>Total actions:</b></td><td>" & tblAction.Rows.Count & "</td></tr></table><p>" & vbCrLf)
      ' Finish the report
      objResult.Add("</body></html>" & vbCrLf)
      ' Return this report
      Return objResult.Text
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/ActionReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   VerifyReport
  ' Goal:   Give an overview of what points to what by what type
  '         The output string gives the HTML report
  ' History:
  ' 21-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function VerifyReport() As String
    Dim ndxForest As XmlNode    ' Current forest
    Dim ndxEtree As XmlNodeList ' A list of all <eTree> elements within this <forest>
    Dim ndxThis As XmlNode      ' One <eTree> element
    Dim ndxAnt As XmlNode       ' The antecedent <eTree>
    Dim objResult As StringColl ' The result of the report
    Dim strRefType As String    ' The kind of reference made
    Dim strSrc As String        ' Text of the source clause
    Dim strAnt As String        ' Text of the antecedent clause
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter
    Dim intPtc As Integer       ' Percentage
    Dim intNum As Integer = 0   ' Number of coref elements made

    Try
      ' Initialise the result
      objResult = New StringColl
      ' Start making the report
      objResult.Add("<html>" & strHtmlHead & "<body><h1>CESAX verification report</h1><p><table>" & _
                    "<tr><td>Time:</td><td>" & Format(Now, "F") & "</td></tr>" & _
                    "<tr><td>Text:</td><td>" & strCurrentFile & "</td></tr>" & _
                    "</table><br> <p>" & vbCrLf)
      ' Give an overview of each decision made by the user and the program
      objResult.Add("<h2>Overview of coreference links</h2>")
      ' Visit all the <forest>'s we have made
      For intI = 0 To pdxList.Count - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ pdxList.Count
        Status("Verification report " & intPtc & "%", intPtc)
        ' Get this forest
        ndxForest = pdxList(intI)
        ' Look at all <eTree> elements
        ndxEtree = ndxForest.SelectNodes(".//eTree")
        For intJ = 0 To ndxEtree.Count - 1
          ' Consider this <eTree> element
          ndxThis = ndxEtree(intJ)
          ' Check if this has a coreference
          If (HasCoref(ndxThis)) Then
            ' Keep track of where we are
            intNum += 1
            ' Get the reference type
            strRefType = CorefAttr(ndxThis, "RefType")
            ' Start the output for this element
            objResult.Add("<table><tr><td>" & intNum & "</td><td><font size='1'><b>" & _
                          ndxForest.Attributes("Location").Value & _
                          "</b></font>: <font color='green'>" & strRefType & "</font></td></tr>")
            ' Get the text of the source 
            strSrc = GetIPfromConstituent(ndxThis)
            ' Find out where it points to
            ndxAnt = CorefDst(ndxThis)
            ' Action depends on whether we have an antecedent or not
            If (ndxAnt Is Nothing) Then
              ' There is no antecedent - just add a row for the source
              objResult.Add("<tr><td></td><td>" & strSrc & "</td></tr></table>")
            Else
              ' There is an antecedent - get its text
              strAnt = GetIPfromConstituent(ndxAnt)
              ' Add the antecedent and the source (in that order)
              objResult.Add("<tr><td>Ant:</td><td>" & strAnt & "</td></tr>")
              objResult.Add("<tr><td>Src:</td><td>" & strSrc & "</td></tr></table>")
            End If
            ' Add an additional paragraph in XHTML
            objResult.Add("<br>")
          End If
        Next intJ
      Next intI
      ' Combine the output
      objResult.Add("</body></html>" & vbCrLf)
      ' Show we are ready
      Status("Ready")
      ' Return this report
      Return objResult.Text
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/VerifyReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   PrepareForest
  ' Goal:   Calculate IP positions and put all the IPs on the antecedent's stack
  ' History:
  ' 31-08-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function PrepareForest(ByRef ndxThis As XmlNode) As Boolean
    Dim ndxIp As XmlNodeList          ' List of all IP nodes of the forest (to be put on antecedent's stack)
    Dim strIPposition As String       ' The absolute position of an IP

    Try
      ' As soon as possible put all the IP nodes on the antecedent's stack
      ndxIp = ndxThis.SelectNodes("./descendant::eTree[" & strIPtypesTb & "]", conTb)
      For intI = 0 To ndxIp.Count - 1
        ' The IP position for each NP within this IP is: forestID/intI+1/ndxIp.Count
        strIPposition = ndxThis.Attributes("forestId").Value & ";" & intI + 1 & ";" & ndxIp.Count
        colIPpos.Add(ndxIp(intI).Attributes("Id").Value, strIPposition)
        ' Put this one IP on the antecedent's stack
        AddIPtoStack(ndxIp(intI))
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/PrepareForest error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '' ------------------------------------------------------------------------------------
  '' Name:   WalkSmart
  '' Goal:   Walk this IP in Left/Breadth/Depth order
  ''         (1) All IP-containing constituents at the left edge
  ''         (2) The constituent itself (breadth)
  ''         (3) All (remaining) children of the constituent (depth)
  '' History:
  '' 18-06-2010  ERK Created
  '' ------------------------------------------------------------------------------------
  'Private Function WalkSmart(ByRef ndxThis As XmlNode) As Boolean
  '  Dim ndxChild As XmlNodeList       ' List of child <eTree> nodes
  '  '   Dim ndxIp As XmlNodeList          ' List of all IP nodes of the forest (to be put on antecedent's stack)
  '  Dim intI As Integer               ' Counter
  '  Dim intStart As Integer           ' Where to start the "proper" children
  '  Dim ndxNext As XmlNode = Nothing  ' Next node
  '  Dim strMsg As String = ""         ' Message
  '  '   Dim strIPposition As String       ' The absolute position of an IP

  '  Try
  '    ' Validate
  '    If (ndxThis Is Nothing) Then Return True
  '    ' Check interrupt
  '    If (bInterrupt) Then Return False
  '    ' Get a list of correct children, EXCLUDING Conjunctions..., and EXCLUDING CODE nodes
  '    '   Also exclusing "," nodes
  '    ' N.B: we have to skip CONJ actually here at this level already!!!
  '    ' N.B2: but we should NOT skip CONJP...
  '    ndxChild = ndxThis.SelectNodes("./eTree[not(tb:Like(string(@Label),'CONJ|CODE|,'))]", conTb)
  '    ' Validate
  '    If (ndxChild.Count = 0) Then Return True
  '    ' Initialise the start to element 0
  '    intStart = 0
  '    ' Are we dealing with an <eTree>?
  '    If (ndxThis.Name = "eTree") Then
  '      ' ============= DEBUGGING ===========
  '      'If (ndxThis.Attributes("Id").Value = 20633) Then Stop
  '      ' ===================================
  '      ' (1) Left-first: 
  '      ' There is 1 or more child, so check out these children
  '      '   Whether they are IP-containing PPs or CP-ADV's
  '      While (IsPPorCPwithIP(ndxThis, ndxChild(intStart)))
  '        ' This child should be treated first
  '        If (Not WalkSmart(ndxChild(intStart))) Then
  '          ' Return failure
  '          Return False
  '        End If
  '        ' The point to start should be incremented
  '        intStart += 1
  '      End While
  '      ' (2) Breadth:
  '      ' If the correct condition holds, then we ourselves need to be dealt with
  '      If (DoLike(ndxThis.Attributes("Label").Value, strIPtypes)) Then
  '        '' As soon as we have an IP, it should be put on the antecedents stack
  '        'AddIPtoStack(ndxThis)
  '        ' This is an IP, so perform the necessary action on it
  '        CoreferenceIP(ndxThis)
  '        ' Push this IP's ID on the "stack" which contains the last X number of IPs
  '        ' (Right now we'll just work with an ordered set of ALL IPs for simplicity)
  '        ' N.B: we are NOT yet doing something with this IP stack...
  '        loc_colIPstack.Add(ndxThis.Attributes("Id").Value)
  '      End If
  '    ElseIf (ndxThis.Name = "forest") Then
  '      '' As soon as possible put all the IP nodes on the antecedent's stack
  '      'ndxIp = ndxThis.SelectNodes("./descendant::eTree[" & strIPtypesTb & "]", conTb)
  '      'For intI = 0 To ndxIp.Count - 1
  '      '  ' The IP position for each NP within this IP is: forestID/intI+1/ndxIp.Count
  '      '  strIPposition = ndxThis.Attributes("forestId").Value & ";" & intI + 1 & ";" & ndxIp.Count
  '      '  colIPpos.Add(ndxIp(intI).Attributes("Id").Value, strIPposition)
  '      '  ' Put this one IP on the antecedent's stack
  '      '  AddIPtoStack(ndxIp(intI))
  '      'Next intI
  '      ' What happens if a forest has MORE THAN 1 child?
  '      ' ==> next child will be treated further on, I suppose?
  '      If (ndxChild.Count > 1) Then
  '        ' Make a message
  '        strMsg = "modAuto/WalkSmart: a forest has more than 1 child." & vbCrLf & _
  '               "This may be unproblematic, but do check the results of line [" & _
  '               ndxThis.Attributes("Location").Value & "]"
  '        ' Show and log this message
  '        If (Not bIsFullyAuto) Then MsgBox(strMsg)
  '        Logging(strMsg)
  '      End If
  '      ' We should really walk the first child too
  '      If (Not WalkSmart(ndxChild(0))) Then
  '        ' Return failure
  '        Return False
  '      End If
  '      ' The point to start should be incremented
  '      intStart += 1
  '    End If
  '    ' (3) Depth:
  '    ' Now all remaining children should be checked out
  '    For intI = intStart To ndxChild.Count - 1
  '      ' Check out this child
  '      If (Not WalkSmart(ndxChild(intI))) Then
  '        ' Something bad happened, so return
  '        Return False
  '      End If
  '    Next intI
  '    ' Return success
  '    Return True
  '  Catch ex As Exception
  '    ' Show error
  '    MsgBox("modAuto/WalkSmart error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
  '    ' Return failure
  '    Return False
  '  End Try
  'End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ConstPrecedes
  ' Goal:   Find out if [ndxOne] linearly precedes [ndxTwo]
  '         True if:
  '           @to(One) < @to(Two)
  '         or else
  '           @from(One) < @from(Two)
  ' History:
  ' 18-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function ConstPrecedes(ByRef ndxOne As XmlNode, ByRef ndxTwo As XmlNode) As Boolean
    Try
      ' Validate
      If (ndxOne Is Nothing) OrElse (ndxTwo Is Nothing) Then Return False
      If (ndxOne.Name <> "eTree") OrElse (ndxTwo.Name <> "eTree") Then Return False
      ' Compare first match
      If (ndxOne.Attributes("to").Value < ndxTwo.Attributes("to").Value) OrElse _
         (ndxOne.Attributes("from").Value < ndxTwo.Attributes("from").Value) Then
        ' One precedes Two
        Return True
      Else
        ' It is the other way around
        Return False
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/ConstPrecedes error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsPPorCPwithIP
  ' Goal:   Find out if [ndxLow] is a PP/CP/IP within a higher IP [ndxHigh]
  '         True if:
  '         (1)  High is an IP
  '         (2)  a. Low is an IP, or
  '              b. Low is a PP or ADVP containing an IP descendant somewhere
  '              c. Low is a CP-ADV
  '              d. Low is a CP... LFD
  ' History:
  ' 21-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function IsPPorCPwithIP(ByRef ndxHigh As XmlNode, ByRef ndxLow As XmlNode) As Boolean
    Dim strLowLabel As String   ' The category of the Low IP
    Dim strHighLabel As String  ' The category of the high IP
    Dim ndxNext As XmlNode      ' Node to probe for something

    Try
      ' Validate
      If (ndxHigh Is Nothing) OrElse (ndxLow Is Nothing) Then Return False
      If (ndxHigh.Name <> "eTree") OrElse (ndxLow.Name <> "eTree") Then Return False
      ' Check if High is an IP
      strHighLabel = ndxHigh.Attributes("Label").Value
      If (Not DoLike(strHighLabel, strIPtypes)) Then
        ' Condition (1) is not met, so return failure
        Return False
      End If
      ' Condition (1) is met, and condition (2) is met before the caller came here (crucial!!)
      ' Make a working copy
      ndxNext = ndxLow
      ' Now check (2a)
      strLowLabel = ndxNext.Attributes("Label").Value
      If (DoLike(strLowLabel, strIPtypes)) Then
        ' Condition (2a) is met, so return success
        Return True
      ElseIf (DoLike(strLowLabel, strSmartCP)) Then
        ' Condition (2c) is met
        Return True
      Else
        ' Start checking condition (2b)
        If (DoLike(strLowLabel, strSmartXP)) Then
          ' This PP should have an IP descendant somewhere
          ndxNext = ndxNext.SelectSingleNode("./descendant::eTree[starts-with(@Label,'IP')]")
          '' Is there such a CP child?
          'If (Not ndxNext Is Nothing) Then
          '  ' Then this one should have an IP child
          '  ndxNext = ndxNext.SelectSingleNode("./child::eTree[starts-with(@Label, 'IP')]")
          ' Any success?
          If (Not ndxNext Is Nothing) Then
            ' Yes, we have met condition (3b)
            Return True
          End If
          'End If
        End If
      End If
      ' Conditions are not met, so return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/IsPPorCPwithIP error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SortIPs
  ' Goal:   Sort this list of IP's in order of increasing [from] and [to] attributes
  ' History:
  ' 18-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function SortIPs(ByRef pdxThis As XmlNodeList, ByRef intIdx() As Integer, ByRef intLen As Integer) As Boolean
    Dim intI As Integer       ' First counter
    Dim intTemp As Integer    ' Temporary index
    Dim bChanged As Boolean   ' Whether any changes have been made

    Try
      ' First make a default array
      ReDim intIdx(0) : intLen = 0
      ' Validate
      If (pdxThis Is Nothing) Then Return True
      If (pdxThis.Count = 0) Then Return True
      ' Make an appropriate list
      ReDim intIdx(0 To pdxThis.Count - 1)
      ' Put the length in there
      intLen = intIdx.Length
      ' Set the indices to this list
      For intI = 0 To pdxThis.Count - 1
        intIdx(intI) = intI
      Next
      ' Continue until no changes have been made
      Do
        ' Initialise no changes
        bChanged = False
        ' Go through all the elements
        For intI = 1 To pdxThis.Count - 1
          ' Compare the two
          If (ConstPrecedes(pdxThis(intIdx(intI)), pdxThis(intIdx(intI - 1)))) Then
            ' Move this one point up
            intTemp = intIdx(intI - 1)
            intIdx(intI - 1) = intIdx(intI)
            intIdx(intI) = intTemp
            ' Indicate we have changed
            bChanged = True
          End If
        Next intI
      Loop While (bChanged)
      ' Now we have an array with indices of how to walk the [pdxThis] list
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/SortIPs error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   UniqueChildOfCorefXP
  ' Goal:   Check if this XP is the unique <eTree> child of an YP that already has 
  '           coreference information
  '         If so, return the antecedent of the parent in [ndxAnt]
  ' History:
  ' 22-11-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function UniqueChildOfCorefXP(ByRef ndxThis As XmlNode, ByRef ndxAnt As XmlNode) As Boolean
    Dim ndxParent As XmlNode  ' My father 
    Dim ndxNext As XmlNode    ' Working node
    Dim intNum As Integer = 0 ' Number of children

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      If (ndxThis.Name <> "eTree") Then Return False
      ' Get my father
      ndxParent = ndxThis.ParentNode
      ' Validate parent
      If (ndxParent Is Nothing) Then Return False
      If (ndxParent.Name <> "eTree") Then Return False
      ' Check number of <eTree> children of my parent
      ndxNext = ndxParent.FirstChild
      While (Not ndxNext Is Nothing)
        ' See if this child needs counting
        If (ndxNext.Name = "eTree") Then
          ' Check if this is not a CODE node
          If (Not DoLike(ndxNext.Attributes("Label").Value, "CODE|META*")) Then
            intNum += 1
          End If
        End If
        ' Go to next child
        ndxNext = ndxNext.NextSibling
      End While
      ' Check the number of children
      If (intNum = 1) Then
        ' Now check if the parent already has coreference information
        If (HasCoref(ndxParent)) Then
          ' Set the antecedent
          ndxAnt = CorefDst(ndxParent)
          Return True
        End If
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/UniqueChildOfCorefXP error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   CoreferenceSrcColl
  ' Goal:   Try resolve coreference for the elements of the source collection
  '         This assumes that the source collection has been filled from elsewhere
  ' History:
  ' 01-09-2010  ERK Created: derived from CoreferenceIP
  ' ------------------------------------------------------------------------------------
  Private Sub CoreferenceSrcColl()
    Dim dtrThis As DataRow        ' One element from the table
    Dim dtrAnt As DataRow         ' The antecedent to the currently evaluated constituent
    Dim dtrSrcStack() As DataRow  ' The source stack in ordered and tweaked form
    Dim ndxAnt As XmlNode         ' The antecedent node
    Dim ndxSrc As XmlNode         ' Source node
    Dim ndxPar As XmlNode         ' Parent node
    Dim strType As String = ""    ' Type of resolvable situation
    Dim strRefType As String = "" ' The kind of coreference relation
    Dim strReason As String = ""  ' Reason for needing to ask the user
    Dim strCnsEval As String = "" ' Evaluation of constraints
    Dim strBack As String = ""    ' Result (empty)
    Dim strMsg As String = ""     ' Message
    Dim strBestAnt As String = "" ' Return from BestAnt()
    Dim bNoAnt As Boolean         ' Whether something should not be considered an antecedent
    Dim bIsInferred As Boolean    ' Whether SHIFT was pressed, implying an INFERRED relation
    Dim intBest As Integer        ' The index of the best evaluated antecedent in the collection
    Dim intDstId As Integer       ' Id of the destination node
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter

    Try
      ' Validate
      If (tblSrcStack.Rows.Count = 0) Then Exit Sub
      ' ============= PROFILING =============
      ProfileStart("CoreferenceSrcColl", 1)
      ' =====================================
      ' Prepare the list of potential antecedents...
      ' (1) Visit all potential antecedents, and increment their IP distance (salience measure)
      For intI = tblAntStack.Rows.Count - 1 To 0 Step -1
        ' Take this element
        dtrThis = tblAntStack(intI)
        ' Calculate the IP distance -- this is valid for ALL source elements in the stack!
        dtrThis.Item("IPdist") = CnsIPdist(tblSrcStack.Rows(0), dtrThis)
        ' If we are in fully automatic mode, then check whether the distance does not become too large
        If (bIsFullyAuto) Then
          ' Check the distance
          If (dtrThis.Item("IPdist") > intMaxIPdist) Then
            ' ============== DEBUGGING =============
            ' If (dtrThis.Item("Id") = 4615) OrElse ((dtrThis.Item("Id") = 4617)) Then Stop
            ' ======================================
            ' Delete this item, since it is too far away for Fully Automatic mode
            dtrThis.Delete()
          End If
        Else
          ' If we are not in fully automatic mode, there still is a cut-off point...
          ' Check the distance
          If (dtrThis.Item("IPdist") > intAbsMaxIPdist) Then
            ' ============== DEBUGGING =============
            ' If (dtrThis.Item("Id") = 4615) OrElse ((dtrThis.Item("Id") = 4617)) Then Stop
            ' ======================================
            ' Delete this item, since it is too far away for Semi-Automatic mode
            dtrThis.Delete()
          End If
        End If
      Next intI
      ' Incrementing of NPcount is already done in ChunkWalk!!
      '' Add the source stack number to the total amount of NPs
      'loc_dblNPcount += tblSrcStack.Rows.Count
      ' ============= PROFILING =============
      ProfileStart("CoreferenceSrcColl/Local", 2)
      ' =====================================
      ' Recursively go through this IP, and resolve depth-first:
      ' (1) appositions
      ' (2) reflexives
      ' (3) bare nominal complements of finite BE clauses and of PPs
      ' These elements are taken OUT of the source stack as soon as they are linked!!!
      ' (NOTE: the "local" action should FOLLOW the "srcstack" action!!!)
      loc_bLocalApp = False : loc_bLocalBare = False : loc_bLocalRefl = False
      'If (Not TraverseIpTree(ndxThisIp, "Local", strBack)) Then
      '  ' Something went wrong! -- Programmer's help needed
      '  MsgBox("modAuto/CoreferenceSrcColl: something went wrong with TraverseIpTree[Local]")
      '  If (bDebugging) Then Stop
      '  ' Get out of here!!
      '  Exit Sub
      'End If
      ' ============= PROFILING =============
      ProfileEnd("CoreferenceSrcColl/Local", 2)
      ' =====================================
      ' ============= PROFILING =============
      ProfileStart("CoreferenceSrcColl/PrefNum", 2)
      ' =====================================
      ' ============================================
      ' Get the local preferential numbers for the SOURCE stack
      For intI = 0 To tblSrcStack.Rows.Count - 1
        ' Access this element
        dtrThis = tblSrcStack(intI)
        '' ==================== DEBUGGING =============
        'If (dtrThis.Item("Id") = 1346) Then Stop
        '' ============================================
        ' Make a preferential number for this element
        dtrThis.Item("PrefNum") = GetPrefNum(dtrThis)
        If (dtrThis.Item("PrefNum") < 0) Then
          ' Need to see what happens here
          ' ================ DEBUGGING ================
          'Stop
          ' ============================================
        End If
        '' ====== DEBUG =====
        'Debug.Print(dtrThis.Item("PrefNum") & vbTab & dtrThis.Item("NPtype") & vbTab & _
        '            dtrThis.Item("GrRole") & vbTab & dtrThis.Item("Vern"))
        '' ==================
      Next intI
      ' ============= PROFILING =============
      ProfileEnd("CoreferenceSrcColl/PrefNum", 2)
      ' =====================================
      ' Derive a source stack ordered in likelyhood of being referential
      dtrSrcStack = tblSrcStack.Select("", "PrefNum ASC")
      ' There are still unresolved constituents, so visit all elements
      For intI = 0 To dtrSrcStack.Length - 1
        ' Make sure we are indeed interruptable
        Application.DoEvents()
        ' See if user pressed interrupt
        If (bInterrupt) Then Exit Sub
        ' Access this element and initialise the antecedent to NOTHING
        dtrThis = dtrSrcStack(intI) : dtrAnt = Nothing : ndxSrc = IdToNode(dtrThis.Item("Id"))
        ' ==================== DEBUGGING =============
        ' If (dtrThis.Item("Id") = 2233) Then Stop
        If (dtrThis.Item("PrefNum") < 0) Then Stop
        ' ============================================
        ' Validate this element's preferential number
        If (dtrThis.Item("PrefNum") >= 0) Then
          ' Assume this CAN be an antecedent
          bNoAnt = False : bIsInferred = False : strReason = "" : ndxAnt = Nothing
          ' ============== DEBUGGING =====================
          'Debug.Print("Node=[" & ndxSrc.SelectSingleNode("./ancestor::forest").Attributes("forestId").Value & _
          '            ":" & ndxSrc.Attributes("Id").Value & "] " & _
          '            " srcStack=" & tblSrcStack.Rows.Count & _
          '            " antStack=" & tblAntStack.Rows.Count & _
          '            " refState=" & CorefAttr(ndxSrc, "RefType"))
          ' ==============================================
          ' See if it really needs resolving
          If (NeedsResolving(dtrThis)) Then
            ' First check to see if this is an "it" pronoun, for which reference is already coded...
            If (IsCataphoricNp(dtrThis, ndxSrc)) Then
              ' It cannot function as antecedent!
              bNoAnt = True
            ElseIf (UniqueChildOfCorefXP(ndxSrc, ndxAnt)) Then
              ' This is the unique <eTree> child of an XP that already has coreference
              ' So we need to copy the coreference information of the parent
              '  (We cannot use [MakeIdentityLink], because the destination may not be in the antecedent stack)
              ndxPar = ndxSrc.ParentNode
              strRefType = CorefAttr(ndxPar, "RefType")
              If (Not CorefFromTo(ndxSrc, ndxAnt, strRefType)) Then
                ' Return with failure
                bInterrupt = True
                Exit Sub
              End If
              ' Add the history
              AddHistory(ndxSrc, "coref", "AutoLocal")
              ' We should report that this connection has been made!!
              AutoReport("Made " & strRefType & " link from " & NodeInfo(ndxSrc) & " to " & NodeInfo(ndxAnt))
              ' Delete the source node from the source stack
              RemoveFromStack(tblSrcStack, ndxSrc)
              ' And then the [ndxSrc] should NOT go into the antecedent stack
              bNoAnt = True
            Else
              ' Try to resolve it with reference to the stack of PRECEDING constituents
              'intBest = BestAnt(dtrThis, tblAntStack, dtrAnt, strReason)
              strBestAnt = BestAnt(dtrThis, tblAntStack, dtrAnt, strReason, intBest)
              '' The evaluation should at least be ZERO
              Select Case strBestAnt
                Case "Error"
                  ' There was an error, so leave this one
                  Exit For
                Case "Ambiguous"
                  ' Should we bother asking at all?
                  If (bIsFullyAuto) Then
                    ' Fully automatic, so make an identity link by default
                    If (Not MakeIdentityLink(dtrThis, dtrAnt, "AutoSusp", dtrAnt, bIsInferred, "")) Then
                      ' Something went wrong here - now what?
                      If (Not bInterrupt) Then
                        MsgBox("modAuto/CoreferenceSrcColl: problem making an identity link (Suspicious/Unable to recover antecedent)")
                        bInterrupt = True
                      End If
                      bInterrupt = True
                      Exit Sub
                    End If
                  Else
                    ' There is more than one best match, so the user should decide
                    If (dtrAnt Is Nothing) Then
                      ndxAnt = Nothing
                      strCnsEval = ""
                    Else
                      ndxAnt = IdToNode(dtrAnt.Item("Id"))
                      strCnsEval = ""
                    End If
                    ' TODO: ask the user
                    If (AskUserDecision(IdToNode(dtrThis.Item("Id")), ndxAnt, _
                        "More than one best match. Please decide what the correct antecedent is." & strCnsEval, bIsInferred)) Then
                      ' Check whether there is an antecedent
                      If (Not ndxAnt Is Nothing) Then
                        ' Find out which element from the antecedent stack this is
                        intJ = GetAntStack(ndxAnt)
                        ' Validate
                        If (intJ >= 0) Then
                          ' Make an identity relation to this antecedent
                          If (Not MakeIdentityLink(dtrThis, tblAntStack(intJ), "User", dtrAnt, bIsInferred, "")) Then
                            ' There is a problem here - now what?
                            If (Not bInterrupt) Then
                              MsgBox("modAuto/CoreferenceSrcColl: problem making an identity link (Ambiguous/Antecedent problem)")
                              bInterrupt = True
                            End If
                            bInterrupt = True
                            Exit Sub
                          End If
                        End If
                      End If
                    Else
                      ' SOmething went wrong
                      If (Not bInterrupt) Then
                        MsgBox("modAuto/CoreferenceSrcColl: problem making an identity link (Ambiguous/AskUser problem)")
                        bInterrupt = True
                      End If
                      bInterrupt = True
                      Exit Sub
                    End If
                  End If
                Case "Suspicious", "BadBestMatch"
                  ' Should we bother asking at all?
                  If (bIsFullyAuto) Then
                    ' Fully automatic, so make an identity link by default
                    If (Not MakeIdentityLink(dtrThis, dtrAnt, "AutoSusp", dtrAnt, bIsInferred, "")) Then
                      ' We just continue, because we are in fully automatic
                    End If
                  Else
                    ' There is a best match (in [dtrAnt]), but there is too much ambiguity, so let user decide
                    ndxAnt = IdToNode(dtrAnt.Item("Id"))
                    strCnsEval = ""
                    ' TODO: ask the user
                    If (AskUserDecision(IdToNode(dtrThis.Item("Id")), ndxAnt, _
                        strReason & vbCrLf & "Please decide what the correct antecedent is." & _
                        vbCrLf & strCnsEval, bIsInferred)) Then
                      ' Check whether there is an antecedent
                      If (Not ndxAnt Is Nothing) Then
                        ' Find out which element from the antecedent stack this is
                        intJ = GetAntStack(ndxAnt)
                        ' Check if the antecedent is in the stack
                        If (intJ < 0) Then
                          ' It is cataphoric: copy the antecedent into the stack
                          If (AddToStack(ndxAnt, tblAntStack)) Then
                            ' Retrial!
                            intJ = GetAntStack(ndxAnt)
                          End If
                        End If
                        ' Validate
                        If (intJ >= 0) Then
                          ' Make an identity relation to this antecedent
                          If (Not MakeIdentityLink(dtrThis, tblAntStack(intJ), "User", dtrAnt, bIsInferred, "")) Then
                            ' Something went wrong here - now what?
                            If (Not bInterrupt) Then
                              MsgBox("modAuto/CoreferenceSrcColl: problem making an identity link (Suspicious/Unable to recover antecedent)")
                              bInterrupt = True
                            End If
                            bInterrupt = True
                            Exit Sub
                          End If
                        End If
                      End If
                    Else
                      ' SOmething went wrong
                      If (Not bInterrupt) Then
                        MsgBox("modAuto/CoreferenceSrcColl: problem making an identity link (Suspicious/AskUser problem)")
                        bInterrupt = True
                      End If
                      Exit Sub
                    End If
                  End If
                Case "NoAntecedent"
                  ' This means there is no antecedent in the collection, or the best match that we found
                  '   does not fulfill necessary criteria. So it is either new or assumed
                  If (Not MakeUnlinkedNP(dtrThis, "Auto", strReason)) Then
                    ' SOmething went wrong
                    If (Not bInterrupt) Then
                      MsgBox("modAuto/CoreferenceSrcColl: problem making an unlinked NP (NoAnt/BadBestMatch)")
                      bInterrupt = True
                    End If
                    bInterrupt = True
                    Exit Sub
                  End If
                Case "NoSuspicion"
                  ' Make an identity link
                  If (Not MakeIdentityLink(dtrThis, dtrAnt, "Auto", Nothing, False, strReason)) Then
                    ' SOmething went wrong
                    If (Not bInterrupt) Then
                      MsgBox("modAuto/CoreferenceSrcColl: problem making an identity link (else)")
                      bInterrupt = True
                    End If
                    bInterrupt = True
                    Exit Sub
                  End If
                Case Else
                  ' This should not happen, so warn the user
                  MsgBox("modAuto/CoreferenceSrcColl: BestAnt returns unknown case [" & strBestAnt & "]")
              End Select
            End If
          ElseIf (dtrThis.Item("Coref") & "" <> "") Then
            ' This element already HAS a coreference relation.
            ' If we are training, then add the data to the trainingsmodel that is being built up
            If (bTraining) Then
              ' Process the training data
              If (Not AddTrainingData(ndxSrc, dtrThis, tblAntStack)) Then
                Logging("modAuto/CoreferenceSrcColl: error adding trainings data")
                bInterrupt = True
                Exit Sub
              End If
            End If
            ' Now antecedent checking and feature copying should take place!
            ' (0) What relation does it have?
            strRefType = CorefAttr(ndxSrc, "RefType")
            ' (1) Does this element have an antecedent?
            intDstId = dtrThis.Item("Ref")
            If (intDstId >= 0) Then
              ' (2) Check whether the existing antecedent is correct
              ndxAnt = IdToNode(intDstId)
              If (IsInAntChain(ndxAnt, intJ)) Then
                ' The antecedent is in the CHAIN of an element in the antecedent chain
                ' Repair this reference!!
                intDstId = tblAntStack(intJ).Item("Id")
                If (Not CorefDstSetRef(IdToNode(dtrThis.Item("Id")), intDstId)) Then
                  ' Warn the user
                  MsgBox("modAuto/CoreferenceSrcColl: the reference from id=" & dtrThis.Item("Id") & _
                         " was incorrect, but I am not able to reset it to ref=" & intDstId)
                End If
                ' Determine antecedent freshly
                ndxAnt = IdToNode(intDstId)
              End If
              ' (3) Determine the antecedent of this element
              dtrAnt = tblAntStack(GetAntStack(ndxAnt))
              ' Did we really get an antecedent?
              ' Check whether it is not:
              ' (a) cataphoric
              ' (b) in the source stack
              If (dtrAnt Is Nothing) AndAlso (intDstId < dtrThis.Item("Id")) AndAlso (GetSrcStack(ndxAnt) < 0) Then
                ' Perhaps this is referring back to an anaphoric pronoun?
                If (IsReflexivePronoun(ndxAnt)) Then
                  ' No wonder this is not on the stack: it is an anaphoric pronoun!
                  ' Repair this by moving the reference from the reflexive to its antecedent
                  ndxAnt = CorefDst(ndxAnt)
                  ' Double check
                  If (ndxAnt Is Nothing) Then
                    ' Trouble!
                    Select Case MsgBox("modAuto/CoreferenceSrcColl: " & _
                        "unable to relocate reflexive antecedent." & _
                        vbCrLf & "Currently link is from [" & dtrThis.Item("Loc") & "].[" & dtrThis.Item("Vern") & "] to [" & _
                        GetLocation(ndxAnt) & "].[" & NodeText(ndxAnt, True) & "]" _
                        , MsgBoxStyle.OkCancel)
                      Case MsgBoxResult.Ok
                        ' Just continue...
                      Case MsgBoxResult.Cancel
                        ' Exit with interrupt set
                        bInterrupt = True
                        Exit Sub
                    End Select
                  End If
                  ' We get here, meaning that we are not too much in trouble... Continue!
                  intDstId = ndxAnt.Attributes("Id").Value
                  ' Change the link from [dtrThis] backwards
                  If (Not CorefDstSetRef(IdToNode(dtrThis.Item("Id")), intDstId)) Then
                    ' Only warn  if this is not fully autocoref
                    If (Not bIsFullyAuto) Then
                      ' Warn the user
                      MsgBox("modAuto/CoreferenceSrcColl: the reference from id=" & dtrThis.Item("Id") & _
                             " was incorrect, but I am not able to reset it to ref=" & intDstId)
                    End If
                  End If
                  ' We should report that this connection has been changed!!
                  AutoReport("Changed link from " & NodeInfo(IdToNode(dtrThis.Item("Id"))) & _
                            " to " & NodeInfo(IdToNode(intDstId)))
                ElseIf (GetAntStack(ndxAnt) < 0) Then
                  ' Only warn if we are not fully auto
                  If (Not bIsFullyAuto) Then
                    ' Only log this problem
                    Logging("unable to find anaphoric antecedent of already annotated constituent in source or antecedent collection." & _
                        vbCrLf & "from [" & dtrThis.Item("Loc") & "].[" & dtrThis.Item("Vern") & "] to [" & _
                        GetLocation(ndxAnt) & "].[" & NodeText(ndxAnt, True) & "]")
                  End If
                End If
              Else
                ' ======= DEBUGGING ========
                'If (InStr(NodeInfo(ndxAnt), "French") > 0) Then Stop
                ' ==========================
                ' Determine the nodes
                ndxSrc = IdToNode(dtrThis.Item("Id"))
                If (dtrAnt Is Nothing) Then
                  ndxAnt = Nothing
                Else
                  ndxAnt = IdToNode(dtrAnt.Item("Id"))
                  ' Copy some features
                  If (Not CopyFeaturesNode(ndxSrc, ndxAnt, strRefType)) Then
                    Logging("unable to copy features." & _
                        vbCrLf & "from [" & dtrThis.Item("Loc") & "].[" & dtrThis.Item("Vern") & "] to [" & _
                        GetLocation(ndxAnt) & "].[" & NodeText(ndxAnt, True) & "]")
                  End If
                End If
                '' Determine whether this is inferred or not
                'bIsInferred = (strRefType = strRefInferred)
                '' (4) Copy relevant features from the antecedent to myself (if there is an antecedent)
                'If (Not dtrAnt Is Nothing) AndAlso (Not bIsInferred) AndAlso (strRefType <> strRefCross) Then
                '  ' Only now features may actually be copied!
                '  CopyFeatures(dtrThis, dtrAnt, bIsInferred)
                'End If
                ' We should report that this connection has been made!!
                AutoReport("Kept " & strRefType & " link from " & NodeInfo(ndxSrc) & _
                          " to " & NodeInfo(ndxAnt))
                ' (5) Delete the antecedent from the stack (if there is one)
                If (Not dtrAnt Is Nothing) AndAlso (strRefType <> strRefInferred) Then
                  ' =========== DEBUGGING ============
                  ' If (dtrAnt.Item("Id") = 4615) OrElse (dtrAnt.Item("Id") = 4617) Then Stop
                  ' ==================================
                  ' Check if the source is a *PRN* element
                  If (ndxSrc.Attributes("Label").Value Like "*PRN*") Then
                    ' We have to remove the source rather than the destination
                    ' This will probably be done automatically??
                  Else
                    ' Yes, remove this element from the antecedent's stack, since something else is referring to it
                    dtrAnt.Delete()
                  End If
                End If
              End If
            Else
              ' This is a non-referring coreference relation (new, assumed etc)
              ' Report that we kept the existing coreference type
              AutoReport("Kept existing " & strRefType & " at " & NodeInfo(IdToNode(dtrThis.Item("Id"))))
            End If
          Else
            ' This element does NOT need resolving, which means that it should be classified as NEW
            If (Not MakeUnlinkedNP(dtrThis, "Auto", "Element does not need resolving")) Then
              ' Make a message
              strMsg = "modAuto/CoreferenceSrcColl problem. Unable to make an unlinked NP." & vbCrLf & _
                NodeInfo(IdToNode(dtrThis.Item("Id"))) & vbCrLf
              ' Show and/or log this message
              If (Not bIsFullyAuto) Then MsgBox(strMsg)
              Logging(strMsg)
            End If
          End If
          ' Check if we should copy to the antecedent's stack
          If (Not bNoAnt) Then
            ' Now we've dealt with it, copy this element to the antecedent's stack
            CopyToAntStack(dtrSrcStack(intI))
          End If
        End If
        ' Note: don't copy elements to the antecedent's stack, whose prefnumbers are zero (which?)
        ' ========== DEBUGGING ======
        If (bDebugging) Then Debug.Print("Antecedent stack = " & tblAntStack.Rows.Count)
        ' ===========================
      Next intI
      ' ============= PROFILING =============
      ProfileEnd("CoreferenceSrcColl", 1)
      ' =====================================
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CoreferenceSrcColl error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   AddIPtoStack
  ' Goal:   Add the [ndxIp] to the stack of antecedents
  ' History:
  ' 30-06-2010  ERK Created
  ' 29-08-2010  ERK Added IPposition
  ' ------------------------------------------------------------------------------------
  Private Sub AddIPtoStack(ByRef ndxThis As XmlNode)
    Dim strLabel As String = ""   ' Label we are considering
    Dim strChLabel As String = "" ' The child's labe
    Dim strNPtype As String = ""  ' Type of NP
    Dim strGrRole As String = ""  ' Grammatical role
    Dim strPGN As String = ""     ' Person/Gender/Number for agreement
    Dim strIPlabel As String = "" ' The label of the IP under which the NP finds itself
    Dim strCoref As String = ""   ' The kind of coreference (if it exists)
    Dim strVern As String = ""    ' The vernacular text of this constituent
    Dim strHdNoun As String = ""  ' The head noun (if existing) of this NP
    Dim strLocType As String = "" ' Kind of local relation to be made
    Dim bIsSpeech As Boolean      ' Whether the IP is a speech IP
    Dim intDstId As Integer       ' The ID of a possible destination (if it exists)
    Dim intNodeId As Integer      ' the ID of this node
    Dim intStackId As Integer     ' the ID of this stack element
    Dim intIpNum As Integer = -1  ' IPnum of this one
    Dim dtrThis As DataRow        ' One new element

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Exit Sub
      ' Check if it is already in here
      If (GetAntStack(ndxThis) >= 0) Then Exit Sub
      ' Get the ID of this element
      intNodeId = ndxThis.Attributes("Id").Value
      ' Get the label of this element
      strLabel = ndxThis.Attributes("Label").Value
      ' Only certain IP's may end up on the stack
      ' (E.g. small clause IPs and infinitieve IPs don't need to end up on the stack?)
      If (Not DoLike(strLabel, strIPstackType)) Then Exit Sub
      ' Since this is an IP,the IPlabael is the same
      strIPlabel = strLabel
      ' =============== DEBUG =============
      'If (intNodeId = "25481") Then Stop
      ' ===================================
      ' Try get values for all relevant feature values
      strNPtype = "IP"
      strGrRole = "Clause"
      ' Get the IP number of this node
      If (Not GetIpNumber(ndxThis, intIpNum)) Then bInterrupt = True : Exit Sub
      ' An IP is 3rd person neuter singular
      strPGN = "3ns"
      ' Okay, we can get the vernacular text of an IP...
      strVern = NodeText(ndxThis, True)
      ' IP's don't have a head noun
      strHdNoun = ""
      ' Determine if this is a speech IP
      bIsSpeech = IsSpeechIp(ndxThis)
      ' IP's may not refer to something else
      strCoref = "" : intDstId = -1
      ' Get the ID for this stack element
      intStackId = tblAntStack.Rows.Count + 1
      ' Make this new datarow
      dtrThis = tblAntStack.NewRow
      ' Fill this element
      With dtrThis
        .Item("StackElId") = intStackId : .Item("Id") = intNodeId
        .Item("Label") = strLabel : .Item("Vern") = strVern
        .Item("Head") = strHdNoun : .Item("Equal") = ""
        .Item("NPtype") = strNPtype : .Item("PGN") = strPGN
        .Item("GrRole") = strGrRole : .Item("Coref") = strCoref
        .Item("ChainNum") = 0 : .Item("IPdist") = 0
        .Item("PrefNum") = 0 : .Item("Ref") = intDstId
        .Item("IPlabel") = strIPlabel : .Item("Eval") = 0
        .Item("ChainIds") = "" : .Item("Level") = 1
        .Item("Loc") = GetLocation(ndxThis)
        .Item("IsLocal") = "False"
        .Item("IsSpeech") = IIf(bIsSpeech, "True", "False")
        .Item("IPpos") = GetIpPosition(ndxThis)
        .Item("HasCoref") = "0"
        .Item("IPances") = strIPlabel
        .Item("Person") = "3" : .Item("Gender") = "n" : .Item("Number") = "s"
        .Item("First") = GetFirstWord(ndxThis)
        .Item("IPnum") = intIpNum
      End With
      ' Add this element to the source stack, with IPdist set to "0"
      tblAntStack.Rows.Add(dtrThis)
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AddIPtoStack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   HasNonWhTrace
  ' Goal:   Check whether the node [ndxThis] is an NP containing a CP with *ICH*-n content.
  '         This is a forward (cataphoric) reference, which we can try to resolve
  '         Return the number "n" of the trace in [intId]
  ' History:
  ' 21-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function HasNonWhTrace(ByRef ndxThis As XmlNode, ByRef intId As Integer) As Boolean
    Dim ndxNext As XmlNode      ' working node
    Dim ndxLeaf As XmlNode      ' Node for the <eLeaf>
    Dim strText As String = ""  ' Text of ELEAF
    Dim intPos As Integer       ' Position in string

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Get the children in succession
      ndxNext = ndxThis.FirstChild
      While (Not ndxNext Is Nothing)
        ' Is this <eTree>?
        If (ndxNext.Name = "eTree") Then
          ' Check if it has a CP label
          If (ndxNext.Attributes("Label").Value Like "CP*") Then
            ' Get its <eLeaf>
            ndxLeaf = ndxNext.Item("eLeaf")
            If (Not ndxLeaf Is Nothing) Then
              ' Get the text of the eLeaf
              strText = ndxLeaf.Attributes("Text").Value
              ' Is this what we are looking for?
              If (strText Like "?ICH?-*") Then
                ' Yes, we are looking for this - get the number
                intPos = InStr(strText, "-")
                If (intPos > 0) Then
                  ' Get the number
                  intId = CInt(Mid(strText, intPos + 1))
                  ' Return success
                  Return True
                End If
              End If
            End If
          End If
        End If
        ' Try the next node
        ndxNext = ndxNext.NextSibling
      End While
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/HasNonWhTrace error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   HasNPcataphore
  ' Goal:   Check whether the node [ndxThis] is an NP with *ICH*-n leaf.
  '         This is a forward (cataphoric) reference, which we can try to resolve
  '         Return the number "n" of the trace in [intId]
  ' History:
  ' 26-10-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function HasNPcataphore(ByRef ndxThis As XmlNode, ByRef intId As Integer) As Boolean
    Dim ndxNext As XmlNode      ' working node
    Dim ndxLeaf As XmlNode      ' Node for the <eLeaf>
    Dim strText As String = ""  ' Text of ELEAF
    Dim intPos As Integer       ' Position in string

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Get the children in succession
      ndxNext = ndxThis.FirstChild
      While (Not ndxNext Is Nothing)
        ' Is this <eLeaf>?
        If (ndxNext.Name = "eLeaf") Then
          ' Get its <eLeaf>
          ndxLeaf = ndxNext
          ' Get the text of the eLeaf
          strText = ndxLeaf.Attributes("Text").Value
          ' Is this what we are looking for?
          If (strText Like "?ICH?-*") Then
            ' Yes, we are looking for this - get the number
            intPos = InStr(strText, "-")
            If (intPos > 0) Then
              ' Get the number
              intId = CInt(Mid(strText, intPos + 1))
              ' Return success
              Return True
            End If
          End If
        End If
        ' Try the next node
        ndxNext = ndxNext.NextSibling
      End While
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/HasNPcataphore error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsCataphoricIt
  ' Goal:   Check and see if this is a cataphoric "it/hit" pronoun,
  '           coreference of which is already coded in the treebank format
  ' History:
  ' 28-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function IsCataphoricNp(ByRef dtrSrc As DataRow, ByRef ndxThis As XmlNode) As Boolean
    Dim ndxNext As XmlNode              ' working node
    Dim ndxForest As XmlNode = Nothing  ' Our own forest's node
    Dim strLabel As String              ' Label of this node
    Dim strSelect As String             ' The Xpath selection criterion
    Dim intId As Integer                ' My own ID 
    Dim intRefId As Integer             ' The reference ID attached to this node

    Try
      ' Validate
      If (dtrSrc Is Nothing) Then Return False
      ' Get the label of the node
      strLabel = dtrSrc.Item("Label")
      ' Validate
      If (strLabel = "") Then Return False
      ' Get this node
      intId = dtrSrc.Item("Id")
      ' Only continue for pronouns and bare demonstratives
      If (DoLike(dtrSrc.Item("NPtype"), "Pro|Dem")) Then
        ' Check if it is a cataphoric it
        ' Does it end with a number?
        If (strLabel Like "*-[0-9]") Then
          ' Yes, we have something!
          ' ========= OLD ============
          'ndxThis = IdToNode(intId)
          '' Validate once more
          'If (ndxThis Is Nothing) Then Return False
          ' ==========================
          ' Evaluate the first <eLeaf> child we can find (since it is a pronoun)
          ndxNext = ndxThis.SelectSingleNode("./descendant::eLeaf")
          ' Validate
          If (ndxNext Is Nothing) Then Return False
          ' Action depends on our own PGN, which should be 3ns (or at least 3s)
          ' WAS: If (DoLike(ndxNext.Attributes("Text").Value, "it|hit|hyt|yt")) Then
          If (DoLike(GetFeature(ndxThis, "NP", "PGN"), "3ns|3s")) Then
            ' We probably have what we are looking for...
            intRefId = CInt(Mid(strLabel, InStrRev(strLabel, "-") + 1))
            ' Get the forest
            If (Not GetForestNode(ndxThis, ndxForest)) Then Return False
            ' Get the antecedent of our pronoun
            strSelect = "./descendant::eTree[@Id>" & intId & " and tb:Like(string(@Label), '*-" & intRefId & "')]"
            ndxNext = ndxForest.SelectSingleNode(strSelect, conTb)
            ' Did we get something?
            If (ndxNext Is Nothing) Then Return False
            ' Yes, we got something!!!
            ' Double check that this indeed refers to something with content
            If (NodeText(ndxThis) = "") Then Return False
            ' Make the link!!!
            If (AddCorefLink(ndxThis, ndxNext, "Auto", "Identity", "(Treebank cataphoric it)", Nothing, "", _
                             "Accepting treebank coding as a cataphoric <it>")) Then
              ' Return success
              Return True
            End If
          End If
        ElseIf (HasNonWhTrace(ndxThis, intRefId)) Then
          ' This is a Pro or Dem containing an *ICH*-n element
          ' Get the antecedent of our pronoun:
          ' (1) Get the forest node we are in
          If (Not GetForestNode(ndxThis, ndxForest)) Then Return False
          ' Determine what we are looking for
          strSelect = "./descendant::eTree[@Id>" & intId & " and tb:Like(string(@Label), '*-" & intRefId & "')]"
          ' Try to get it
          ndxNext = ndxForest.SelectSingleNode(strSelect, conTb)
          ' Did we get something?
          If (ndxNext Is Nothing) Then Return False
          ' Yes. Now get the first IP under this node
          ndxNext = ndxNext.SelectSingleNode("./descendant::eTree[starts-with(@Label, 'IP')]")
          ' Got something?
          If (ndxNext Is Nothing) Then Return False
          ' Make the link!!!
          If (AddCorefLink(ndxThis, ndxNext, "Auto", "Identity", "(Treebank cataphoric Dem/Pro)", Nothing, "", _
                           "Accepting treebank coding as a cataphoric Dem/Pro")) Then
            ' Return success
            Return True
          End If
        End If
      Else
        If (HasNPcataphore(ndxThis, intRefId)) Then
          ' This is an NP containing an *ICH*-n element
          ' Get the antecedent of our NP:
          ' (1) Get the forest node we are in
          If (Not GetForestNode(ndxThis, ndxForest)) Then Return False
          ' Determine what we are looking for
          strSelect = "./descendant::eTree[@Id>" & intId & " and tb:Like(string(@Label), '*-" & intRefId & "')]"
          ' Try to get it
          ndxNext = ndxForest.SelectSingleNode(strSelect, conTb)
          ' Did we get something?
          If (ndxNext Is Nothing) Then Return False
          '' Yes. Now get the first IP under this node
          'ndxNext = ndxNext.SelectSingleNode("./descendant::eTree[starts-with(@Label, 'IP')]")
          '' Got something?
          'If (ndxNext Is Nothing) Then Return False
          ' Make the link!!!
          If (AddCorefLink(ndxThis, ndxNext, "Auto", "Identity", "(Treebank cataphoric NP)", Nothing, "", _
                           "Accepting treebank coding as a cataphoric NP")) Then
            ' Return success
            Return True
          End If
        End If
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/IsCataphoricIt error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsInAntStack
  ' Goal:   Check if the indicated node is already in the antecedent's collection
  ' History:
  ' 31-08-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function IsInAntStack(ByRef ndxThis As XmlNode) As Boolean
    Dim intId As Integer      ' ID of this node
    Dim dtrFound() As DataRow ' Result of SELECT statement

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      If (ndxThis.Name <> "eTree") Then Return False
      ' Get the ID of this node
      intId = ndxThis.Attributes("Id").Value
      ' Find it in the antecedent's collection
      dtrFound = tblAntStack.Select("Id=" & intId)
      ' Return the result
      Return (dtrFound.Length > 0)
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/IsInAntStack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CopyToAntStack
  ' Goal:   Copy the indicated element to the antecedent's stack
  ' History:
  ' 23-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub CopyToAntStack(ByRef dtrSrc As DataRow)
    Dim dtrThis As DataRow    ' The new element for the antecedent's stack
    Dim dtrFound() As DataRow ' Result of SELECT statement
    Dim bIsLocal As Boolean = False ' Whether this antecedent should ONLY be available for the current <forest>

    Try
      ' Validate
      If (dtrSrc Is Nothing) Then Exit Sub
      ' ============= PROFILING =============
      ProfileStart("CopyToAntStack", 1)
      ' =====================================
      ' Check whether this constituent CAN actually be referred to or not
      If (IsPossibleAntecedent(dtrSrc, bIsLocal)) Then
        ' ============= DEBUGGING ==================
        ' Debug.Print("CopyToAntStack size = " & tblAntStack.Rows.Count)
        ' ==========================================
        ' Check if this item is not actually already present in the antecedent's stack
        dtrFound = tblAntStack.Select("Id=" & dtrSrc.Item("Id"))
        ' Only process one that is not yet in the stack!
        If (dtrFound.Length = 0) Then
          ' =========== debugging============
          'If (dtrSrc.Item("Id") = 25477) Then Stop
          ' =================================
          ' Make a new datarow
          dtrThis = tblAntStack.NewRow
          ' Fill this row with the correct values
          For intJ = 0 To tblAntStack.Columns.Count - 1
            ' Copy this item's value
            dtrThis.Item(intJ) = dtrSrc.Item(intJ)
          Next intJ
          ' Add it to the antecedents stack
          tblAntStack.Rows.Add(dtrThis)
        Else
          ' This constituent is already on the stack.  Get its datarow
          dtrThis = dtrFound(0)
        End If
        ' Make sure certain features are copied
        ' Set the "Level" depending on the clause type
        Select Case Left(dtrThis.Item("Label"), 2)
          Case "IP"
            dtrThis.Item("Level") = 1
          Case Else
            dtrThis.Item("Level") = 0
        End Select
        ' Make sure the IP distance is initialized
        dtrThis.Item("IPdist") = 0
        ' Pass on the [IsLocal] parameter -- if it is set
        If (bIsLocal) Then
          dtrThis.Item("IsLocal") = "True"
          ' N.B: don't put it to "False" here!!!
        End If
      End If
      ' ============= PROFILING =============
      ProfileEnd("CopyToAntStack", 1)
      ' =====================================
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CopyToAntStack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   MakeUnlinkedNPnode
  ' Goal:   Make a "New" or "Assumed" coreference relation, depending on the NP type
  ' History:
  ' 23-06-2010  ERK Created
  ' 01-09-2010  ERK This takes the XML node instead of a source collection element
  ' ------------------------------------------------------------------------------------
  Public Function MakeUnlinkedNPnode(ByRef ndxThis As XmlNode, ByVal strLinkType As String, _
      ByVal strReason As String, Optional ByVal strRefType As String = "") As Boolean
    Dim bIsNew As Boolean     ' The type of link we are going to make
    Dim strIPsrc As String    ' THe source's IP text in HTML
    Dim strLocation As String ' The location of the source
    Dim strInfo As String     ' Necessary information
    Dim strNPtype As String   ' The type of NP
    Dim strVern As String     ' The vernacular text of me
    Dim intId As Integer      ' This node's ID

    Try
      ' ============= PROFILING =============
      ProfileStart("MakeUnlinkedIP", 1)
      ' =====================================
      ' Determine the NP type
      strNPtype = GetFeature(ndxThis, "NP", "NPtype")
      ' Determine the vernacular text, including the stars
      strVern = NodeText(ndxThis, True)
      ' Determine my ID
      intId = ndxThis.Attributes("Id").Value
      ' Are we forced to choose one kind of unlinked NP type?
      If (strRefType = "") Then
        ' First check to see if this is not a Trace (a variable)
        If (strNPtype = "Trace") AndAlso (Left(strVern, 3) = "*T*") Then
          ' This is a variable, which should get a different reference type
          strRefType = strRefNewVar
          bIsNew = False
        Else
          ' The type depend on other matters
          bIsNew = (strNPtype <> "DemNP")
          ' Determine what the actual ref type is
          strRefType = IIf(bIsNew, strRefNew, strRefAssumed)
        End If
        ' ========== DEBUGGING ================
        'If (Not bIsNew) Then Stop
        ' =====================================
      Else
        ' We are forced to choose one type
        bIsNew = (strRefType = strRefNew)
      End If
      strInfo = NodeInfo(ndxThis)
      ' Since this is an UNLINKED coreference, delete any possible antecedent reference
      DelOneCoref(ndxThis)
      ' What kind of relation are we making?
      If (bIsNew) Then
        ' Make a new relation
        If (CorefFromTo(ndxThis, Nothing, strRefNew)) Then
          ' Add the history
          If (Not AddHistory(ndxThis, "coref", strLinkType)) Then Return False
          ' We should report that this connection has been made!!
          AutoReport("Established as 'New': " & strInfo)
          ' Get the source IP clause
          strIPsrc = GetIPfromConstituent(ndxThis)
          ' Get the location
          strLocation = GetLocation(ndxThis)
          ' Store this link in the AutoReport table
          ActionAdd(intId, -1, strLinkType, strRefNew, strLocation, strIPsrc, "", "", "", "", strReason)
          ' Keep track of statistics
          AutoStatAdd(intId, strInfo, strRefType, loc_strLastCns, strLinkType, 0)
          ' Make sure the dirty flag in the main form is set
          frmMain.SetDirty(True)
        Else
          ' There is a problem
          Logging("modAuto/MakeUnlinkedNPnode: unable to mark " & intId & _
                  " as discourse-new")
          ' ============= PROFILING =============
          ProfileEnd("MakeUnlinkedIP", 1)
          ' =====================================
          ' We can still try to find the next one
          Return False
        End If
      Else
        ' Make an assumed relation
        If (CorefFromTo(ndxThis, Nothing, strRefType)) Then
          ' We should report that this connection has been made!!
          AutoReport("Established as '" & strRefType & "': " & strInfo)
          ' Get the source IP clause
          strIPsrc = GetIPfromConstituent(ndxThis)
          ' Get the location
          strLocation = GetLocation(ndxThis)
          ' Store this link in the AutoReport table
          ActionAdd(intId, -1, strLinkType, strRefType, strLocation, strIPsrc, "", "", "", "", strReason)
          ' Keep track of statistics
          AutoStatAdd(intId, strInfo, strRefType, loc_strLastCns, strLinkType, 0)
          ' Make sure the dirty flag in the main form is set
          frmMain.SetDirty(True)
        Else
          ' There is a problem
          Logging("modAuto/MakeUnlinkedNPnode: unable to mark " & intId & _
                  " as " & strRefType)
          ' ============= PROFILING =============
          ProfileEnd("MakeUnlinkedIP", 1)
          ' =====================================
          ' We can still try to find the next one
          Return False
        End If
      End If
      ' ============= PROFILING =============
      ProfileEnd("MakeUnlinkedIP", 1)
      ' =====================================
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/MakeUnlinkedNPnode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MakeUnlinkedNP
  ' Goal:   Make a "New" or "Assumed" coreference relation, depending on the NP type
  ' History:
  ' 23-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function MakeUnlinkedNP(ByRef dtrThis As DataRow, ByVal strLinkType As String, _
      ByVal strReason As String, Optional ByVal strRefType As String = "") As Boolean
    Dim ndxThis As XmlNode    ' The node we are talking about
    Dim bIsNew As Boolean     ' The type of link we are going to make
    Dim strIPsrc As String    ' THe source's IP text in HTML
    Dim strLocation As String ' The location of the source
    Dim strInfo As String     ' Necessary information

    Try
      ' Validate
      If (dtrThis Is Nothing) Then Return False
      ' ============== DEBUGGING ==================
      'If (dtrThis.Item("Id") = 21) Then Stop
      ' ===========================================
      ' ============= PROFILING =============
      ProfileStart("MakeUnlinkedIP", 1)
      ' =====================================
      ndxThis = IdToNode(dtrThis.Item("Id"))
      ' Are we forced to choose one kind of unlinked NP type?
      If (strRefType = "") Then
        ' First check to see if this is not a Trace (a variable)
        If (dtrThis.Item("NPtype") = "Trace") AndAlso (Left(dtrThis.Item("Vern"), 3) = "*T*") Then
          ' This is a variable, which should get a different reference type
          strRefType = strRefNewVar
          bIsNew = False
        Else
          ' The type depend on other matters
          bIsNew = (dtrThis.Item("NPtype") <> "DemNP")
          ' Determine what the actual ref type is
          strRefType = IIf(bIsNew, strRefNew, strRefAssumed)
        End If
        ' ========== DEBUGGING ================
        'If (Not bIsNew) Then Stop
        ' =====================================
      Else
        ' We are forced to choose one type
        bIsNew = (strRefType = strRefNew)
      End If
      ' Make sure the Reference Type gets put into the datarow
      dtrThis.Item("Coref") = strRefType
      strInfo = NodeInfo(ndxThis)
      ' What kind of relation are we making?
      If (bIsNew) Then
        ' Make a new relation
        If (CorefFromTo(ndxThis, Nothing, strRefNew)) Then
          ' Add the history
          If (Not AddHistory(ndxThis, "coref", strLinkType)) Then Return False
          ' We should report that this connection has been made!!
          AutoReport("Established as 'New': " & strInfo)
          ' Get the source IP clause
          strIPsrc = GetIPfromConstituent(ndxThis)
          ' Get the location
          strLocation = GetLocation(ndxThis)
          ' Store this link in the AutoReport table
          ActionAdd(dtrThis.Item("Id"), -1, strLinkType, "New", strLocation, strIPsrc, "", "", "", "", strReason)
          ' Keep track of statistics
          AutoStatAdd(dtrThis.Item("Id"), strInfo, strRefType, loc_strLastCns, strLinkType, 0)
          ' Make sure the dirty flag in the main form is set
          frmMain.SetDirty(True)
        Else
          ' There is a problem
          Logging("modAuto/MakeUnlinkedNP: unable to mark " & dtrThis.Item("Id") & _
                  " as discourse-new")
          ' ============= PROFILING =============
          ProfileEnd("MakeUnlinkedIP", 1)
          ' =====================================
          ' We can still try to find the next one
          Return False
        End If
      Else
        ' Make an assumed relation
        If (CorefFromTo(ndxThis, Nothing, strRefType)) Then
          ' We should report that this connection has been made!!
          AutoReport("Established as '" & strRefType & "': " & strInfo)
          ' Get the source IP clause
          strIPsrc = GetIPfromConstituent(ndxThis)
          ' Get the location
          strLocation = GetLocation(ndxThis)
          ' Store this link in the AutoReport table
          ActionAdd(dtrThis.Item("Id"), -1, strLinkType, strRefType, strLocation, strIPsrc, "", "", "", "", strReason)
          ' Keep track of statistics
          AutoStatAdd(dtrThis.Item("Id"), strInfo, strRefType, loc_strLastCns, strLinkType, 0)
          ' Make sure the dirty flag in the main form is set
          frmMain.SetDirty(True)
        Else
          ' There is a problem
          Logging("modAuto/MakeUnlinkedNP: unable to mark " & dtrThis.Item("Id") & _
                  " as " & strRefType)
          ' ============= PROFILING =============
          ProfileEnd("MakeUnlinkedIP", 1)
          ' =====================================
          ' We can still try to find the next one
          Return False
        End If
      End If
      ' ============= PROFILING =============
      ProfileEnd("MakeUnlinkedIP", 1)
      ' =====================================
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/MakeUnlinkedNP error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsInAntChain
  ' Goal:   Check whether the node [ndxThis] is stored somewhere in a CHAIN element 
  '           of an item in the antecedent's stack.
  '         If this is so, then get the index of the node [ndxThis] in the antecedents' stack
  ' History:
  ' 30-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function IsInAntChain(ByRef ndxThis As XmlNode, ByRef intAntStackRow As Integer) As Boolean
    Dim intI As Integer       ' Coutner
    Dim intThisId As Integer  ' the ID of the element

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return -1
      ' Initialise
      intAntStackRow = -1 : intThisId = ndxThis.Attributes("Id").Value
      ' Check all
      For intI = 0 To tblAntStack.Rows.Count - 1
        ' ======== DEBUGGING ===========
        'Debug.Print("Stack[" & intI & "] chain=" & tblAntStack.Rows(intI).Item("ChainIds"))
        ' ==============================
        ' Is it this row?
        If (IsInSemiStack(tblAntStack.Rows(intI).Item("ChainIds") & "", intThisId)) Then
          ' Got it
          intAntStackRow = intI
          Return True
        End If
      Next intI
      ' Found nothing
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/IsInAntChain error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetAntStack
  ' Goal:   Get the index of the node [ndxThis] in the antecedents' stack
  ' History:
  ' 22-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetAntStack(ByRef ndxThis As XmlNode) As Integer
    Dim intI As Integer  ' Coutner

    Try
      ' Validate
      If (ndxThis Is Nothing) OrElse (tblAntStack Is Nothing) Then Return -1
      ' Check all
      For intI = 0 To tblAntStack.Rows.Count - 1
        ' Is it this row?
        If (ndxThis.Attributes("Id").Value = tblAntStack.Rows(intI).Item("Id")) Then
          ' Got it
          Return intI
        End If
      Next intI
      ' Found nothing
      Return -1
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetAntStack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetSrcStack
  ' Goal:   Get the index of the node [ndxThis] in the source stack
  ' History:
  ' 22-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetSrcStack(ByRef ndxThis As XmlNode) As Integer
    Dim intI As Integer  ' Coutner

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return -1
      If (tblSrcStack Is Nothing) Then Return -1
      ' Check all
      For intI = 0 To tblSrcStack.Rows.Count - 1
        ' Is it this row?
        If (ndxThis.Attributes("Id").Value = tblSrcStack.Rows(intI).Item("Id")) Then
          ' Got it
          Return intI
        End If
      Next intI
      ' Found nothing
      Return -1
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetSrcStack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsSpeechIp
  ' Goal:   Determine whether this IP is (in)direct speech or not
  '         I am direct speech if:
  '         (1) this IP is marked as IP*SPE, or
  '         I am indirect speech, if:
  '         (1) my grandfather IP has a speech intro verb, or
  '         (2) my grandfather IP is (in)direct speech
  ' History:
  ' 04-10-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function IsSpeechIp(ByRef ndxThis As XmlNode) As Boolean
    Dim ndxNext As XmlNode = Nothing  ' Next node
    Dim intId As Integer              ' Index in collection

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Should be an <eTree> type
      If (ndxThis.Name <> "eTree") Then Return False
      ' It should have an IP label
      If (Not ndxThis.Attributes("Label").Value Like "IP*") Then Return False
      ' Check the label
      If (ndxThis.Attributes("Label").Value Like "*SPE*") Then Return True
      ' Get the IP parent
      ndxNext = GetIpAncestor(ndxThis)
      If (ndxNext Is Nothing) Then Return False
      ' See if this IP ancestor is in the source collection
      intId = GetSrcStack(ndxNext)
      If (intId >= 0) Then
        ' The IP is in the source collection
        If (tblSrcStack(intId).Item("IsSpeech").ToString Like "True") Then Return True
      End If
      ' See if this IP ancestor is in the antecedent collection
      intId = GetAntStack(ndxNext)
      If (intId >= 0) Then
        ' The IP is in the source collection
        If (tblAntStack(intId).Item("IsSpeech").ToString Like "True") Then Return True
      End If
      ' Try to get a verb in the immediate children
      If (SomeChild(ndxThis, "V*", ndxNext)) Then
        ' We have the first "verbal" child -- get its leaf
        ndxNext = ndxNext.SelectSingleNode(".//eLeaf")
        ' Got any result?
        If (Not ndxNext Is Nothing) Then
          ' compare the resulting text
          If (IsAmongSpeechIntroVerbs(ndxNext.Attributes("Text").Value)) Then
            ' Return success
            Return True
          End If
        End If
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/IsSpeechIp error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsAmongSpeechIntroVerbs
  ' Goal:   Check if the given input string is among the listed speech intro verbs
  ' History:
  ' 04-10-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function IsAmongSpeechIntroVerbs(ByVal strIn As String) As Boolean
    Dim strVariants As String ' Variants for this entry
    Dim intI As Integer       ' Counter
    Dim intK As Integer       ' Variant counter (in order of decreasing periods)
    Dim bPeriodOk As Boolean  ' Whether this period should be checked or not
    Dim strResult As String = ""  ' The resulting period
    Dim dtrFound() As DataRow     ' Result of SELECT function

    Try
      ' Validate
      If (strIn = "") Then Return False
      ' Select all SpeechIntro verbs
      dtrFound = tdlSettings.Tables("Category").Select("Name='SpeechIntro'")
      ' Initialise period OK flag
      bPeriodOk = False
      ' Visit all periods in increasing order (that is, starting from MBE > ME > OE)
      For intK = 0 To UBound(arPeriod)
        ' See if this period is ok
        If (arPeriod(intK) = strCurrentPeriod) Then bPeriodOk = True
        ' Only start processing if this period is okay
        If (bPeriodOk) Then
          ' Try find the correct entry
          For intI = 0 To dtrFound.Length - 1
            ' Get the variants for this entry and for this period
            strVariants = dtrFound(intI).Item(arPeriod(intK)) & ""
            If (strVariants <> "") Then
              ' Check if it is in here
              If (DoLike(strIn, strVariants)) Then
                ' We have found the verb to be among the speech verbs - return tru
                Return True
              End If
            End If
          Next intI
        End If
      Next intK
      ' Default: return false
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/IsAmongSpeechIntroVerbs error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SomeChild
  ' Goal:   Get the first child fulfilling the conditions
  ' History:
  ' 04-10-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function SomeChild(ByRef ndxThis As XmlNode, ByVal strFilter As String, ByRef ndxChild As XmlNode) As Boolean
    Dim ndxNext As XmlNode  ' Next node

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Should be an <eTree> type
      If (Not DoLike(ndxThis.Name, "forest|eTree")) Then Return False
      ' Start from the first child
      ndxNext = ndxThis.FirstChild
      While (Not ndxNext Is Nothing)
        ' Is this an <eTree>
        If (ndxNext.Name = "eTree") Then
          ' Check the label
          If (DoLike(ndxNext.Attributes("Label").Value, strFilter)) Then
            ' Found the correct child
            ndxChild = ndxNext
            Return True
          End If
        End If
        ' Try next child
        ndxNext = ndxNext.NextSibling
      End While
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/SomeChild error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   FirstChildImm
  ' Goal:   Get the immediately first child if it fulfills the conditions
  ' History:
  ' 29-11-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function FirstChildImm(ByRef ndxThis As XmlNode, ByVal strFilter As String, _
                                 ByRef ndxChild As XmlNode, ByVal strSkip As String) As Boolean
    Dim ndxNext As XmlNode  ' Next node
    Dim strLabel As String  ' The label

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Should be an <eTree> type
      If (ndxThis.Name <> "eTree") Then Return False
      ' Start from the first child
      ndxNext = ndxThis.FirstChild
      While (Not ndxNext Is Nothing)
        ' Is this an <eTree>
        If (ndxNext.Name = "eTree") Then
          ' Check the label
          strLabel = ndxNext.Attributes("Label").Value
          If (DoLike(strLabel, strFilter)) Then
            ' Found the correct child
            ndxChild = ndxNext
            Return True
          ElseIf (Not DoLike(strLabel, strSkip)) Then
            ' Conditions are not fulfilled, so return failure
            Return False
          End If
        End If
        ' Try next child
        ndxNext = ndxNext.NextSibling
      End While
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/FirstChildImm error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MakeIdentityLink
  ' Goal:   Establish an identity link from the node referred to in [dtrThis] to the
  '           node referred to in [dtrAnt]
  '         Then remove the antecedent from the stack of antecedents
  ' Notes:  The [dtrAuto] element, if not empty, contains the best choice made by the 
  '           algorithm, which has been rejected by the user
  ' History:
  ' 22-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function MakeIdentityLink(ByRef dtrThis As DataRow, ByRef dtrAnt As DataRow, _
      ByVal strLinkType As String, ByRef dtrAuto As DataRow, ByVal bIsInferred As Boolean, _
      ByVal strReason As String) As Boolean
    Dim ndxThis As XmlNode    ' The source node
    Dim ndxAnt As XmlNode     ' The possible antecedent
    Dim ndxAuto As XmlNode    ' The antecedent that has been rejected
    Dim strType As String     ' The type of coreference link to be made
    Dim strAutoCon As String  ' Constraints for the automatically determined one
    Dim bIPsrc As Boolean     ' Whether the source node is part of a SPEECH IP
    Dim bIPant As Boolean     ' Whether the antecedent node is part of a speech IP

    Try
      ' Validate
      If (dtrThis Is Nothing) OrElse (dtrAnt Is Nothing) Then Return False
      ' ============= PROFILING =============
      ProfileStart("MakeIdentityLink", 1)
      ' =====================================
      ' Determine the source and antecedent nodes
      ndxThis = IdToNode(dtrThis.Item("Id"))
      ndxAnt = IdToNode(dtrAnt.Item("Id"))
      If (dtrAuto Is Nothing) Then
        ndxAuto = Nothing : strAutoCon = ""
      ElseIf (dtrAuto.Item("Id") <> dtrAnt.Item("Id")) Then
        ndxAuto = IdToNode(dtrAuto.Item("Id"))
        ' Get the constraints for the automatically determined one
        strAutoCon = dtrAuto.Item("Constraints")
      Else
        ' The automatically proposed node is equal to the antecedent we want to have
        'WAS: ndxAuto = Nothing : strAutoCon = ""
        ' TODO: check if this is okay
        ndxAuto = IdToNode(dtrAuto.Item("Id"))
        ' Get the constraints for the [User-Agree] variant
        strAutoCon = dtrAuto.Item("Constraints")
      End If
      '' ================ DEBUG =================
      'If (dtrThis.Item("Id") = 20454) Then Stop
      'If (dtrThis.Item("Vern") = "he") AndAlso (dtrAnt.Item("Vern") = "we") Then Stop
      ' if (bDebugging) then Debug.Print(dtrAnt.Item("Constraints"))
      '' ========================================
      ' First check whether control was pressed
      If (bIsInferred) Then
        ' The coreference type is inferred
        strType = "Inferred"
      Else
        ' Determine the type of coreference link
        bIPsrc = (dtrThis.Item("IsSpeech") Like "True")
        bIPant = (dtrAnt.Item("IsSpeech") Like "True")
        'bIPsrc = (dtrThis.Item("IPlabel") Like "*SPE")
        'bIPant = (dtrAnt.Item("IPlabel") Like "*SPE")
        If (bIPant And Not bIPsrc) OrElse (bIPsrc And Not bIPant) Then
          ' There is crossspeech
          strType = "CrossSpeech"
        Else
          strType = "Identity"
        End If
      End If
      ' Set a coreference relation
      If (AddCorefLink(ndxThis, ndxAnt, strLinkType, strType, dtrAnt.Item("Constraints"), _
                       ndxAuto, strAutoCon, strReason)) Then
        ' Make sure the source gets its HasCoref feature checked
        dtrThis.Item("HasCoref") = "1"
        '' Copy some of the the antecedent's properties to ME
        'If (Not CopyFeatures(dtrThis, dtrAnt, bIsInferred)) Then Return False
        ' If the source NP has a head noun, and there are head nouns in the antecedent chain, we want to save them
        If (Not bIsInferred) Then
          ' Obviously only if the relation is not "inferred"
          ChainDictProcess(dtrThis.Item("Head"), dtrAnt.Item("Head"), dtrAnt.Item("Equal"))
        End If
        ' Deletion from the antecedent's stack should NOT take place under certain circumstances
        ' (1) an indefinite NP may stay on the stack, because others can take it as antecedent
        ' (2) the antecedent of an "Inferred" relation should stay on the stack
        If (Not DoLike(dtrAnt.Item("NPtype"), "IndefNP|Proper")) AndAlso (Not bIsInferred) Then
          ' Delete the antecedent from the stack of antecedents
          dtrAnt.Delete()
        End If
        ' ============= PROFILING =============
        ProfileEnd("MakeIdentityLink", 1)
        ' =====================================
        ' Return success
        Return True
      Else
        ' There is a problem
        Logging("modAuto/MakeIdentityLink: unable to connect from " & dtrThis.Item("Id") & _
                " to " & ndxAnt.Attributes("Id").Value)
        ' ============= PROFILING =============
        ProfileEnd("MakeIdentityLink", 1)
        ' =====================================
        ' Return failure
        Return False
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/MakeIdentityLink error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddHistory
  ' Goal:   Add a history <f> feature to the [strTYpe] section of <fs>
  ' History:
  ' 05-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AddHistory(ByRef ndxSrc As XmlNode, ByVal strType As String, ByVal strMode As String) As Boolean
    Dim strHistory As String = "" ' Current history for this Type

    Try
      ' Validate
      If (ndxSrc Is Nothing) Then Return False
      ' Get the history
      strHistory = GetFeature(ndxSrc, "Coref", "History")
      ' Add this new history element
      AddSemiStack(strHistory, "" & strUserName & ":" & strMode & "(" & Format(Now, "d") & ")")
      ' Add this feature on the correct place
      If (Not AddFeature(pdxCurrentFile, ndxSrc, strType, "history", strHistory)) Then Return False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AddHistory error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddCorefLink
  ' Goal:   Add a coreference link based upon the XML nodes
  ' History:
  ' 28-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AddCorefLink(ByRef ndxThis As XmlNode, ByRef ndxAnt As XmlNode, ByVal strLinkType As String, _
      ByVal strRefType As String, ByVal strConstraints As String, ByRef ndxAuto As XmlNode, _
      ByVal strAutoCons As String, ByVal strReason As String) As Boolean
    Dim strLocation As String ' The <forest> location of the source
    Dim strIPsrc As String    ' HTML of the source clause
    Dim strIPant As String    ' HTML of the antecedent clause
    Dim strIPauto As String   ' HTML of the rejected antecedent's clause
    Dim strInfo As String     ' Information
    Dim dtrThis As DataRow    ' Antecedent stack item
    Dim intAntId As Integer   ' Antecedent stack number
    Dim intIPdist As Integer = -1 ' IP depth

    Try
      ' Validate
      If (ndxThis Is Nothing) OrElse (ndxAnt Is Nothing) Then Return False
      ' Set a coreference relation with the subject
      If (CorefFromTo(ndxThis, ndxAnt, strRefType)) Then
        ' Copy features from antecedent to me
        If (Not CopyFeaturesNode(ndxThis, ndxAnt, strRefType)) Then
          ' Well, it is a pity that we cannot copy the features, but not a deadly sin!
        End If
        ' Copy the ChainId feature from source to antecedent and further down
        If (Not CopyChainId(ndxThis, ndxAnt, strRefType)) Then
          ' It's a pity if this doesn't work -- user will have to recalculate
        End If
        ' Add the history
        If (Not AddHistory(ndxThis, "coref", strLinkType)) Then Return False
        ' We should report that this connection has been made!!
        strInfo = "from " & NodeInfo(ndxThis) & " to " & NodeInfo(ndxAnt)
        AutoReport("Made " & strRefType & " link " & strInfo & " Constraints: " & strConstraints)
        ' Get the source IP clause
        strIPsrc = GetIPfromConstituent(ndxThis)
        strIPant = GetIPfromConstituent(ndxAnt)
        strIPauto = GetIPfromConstituent(ndxAuto)
        ' Get the location
        strLocation = GetLocation(ndxThis)
        ' Is this a user link?
        If (strLinkType = "User") Then
          ' Adapt the link type, depending on agreement or disagreement between user and us
          If (EqualNodes(ndxAuto, ndxAnt)) Then
            strLinkType &= "-Agree"
          Else
            strLinkType &= "-Differ"
          End If
        End If
        ' Store this link in the AutoReport table
        ActionAdd(ndxThis.Attributes("Id").Value, ndxAnt.Attributes("Id").Value, strLinkType, strRefType, _
                  strLocation, strIPsrc, strIPant, strIPauto, strConstraints, strAutoCons, strReason)
        ' Keep track of statistics
        intAntId = GetAntStack(ndxAnt)
        If (intAntId >= 0) Then
          dtrThis = tblAntStack(intAntId)
          If (Not dtrThis Is Nothing) Then
            intIPdist = dtrThis.Item("IPdist")
          End If
        End If
        AutoStatAdd(ndxThis.Attributes("Id").Value, strInfo, strRefType, strConstraints, strLinkType, intIPdist)
        ' Make sure the dirty flag in the main form is set
        frmMain.SetDirty(True)
        ' Return success
        Return True
      Else
        Return False
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AddCorefLink error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetIPfromConstituent
  ' Goal:   Get the <forest> IP of the indicated node in HTML form
  ' History:
  ' 28-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetIPfromConstituent(ByRef ndxThis As XmlNode) As String
    Dim ndxForest As XmlNode = Nothing  ' The forest node of the source
    Dim ndxBefore As XmlNode = Nothing  ' The preceding forest node (if any)
    Dim strContext As String = ""       ' The preceding context (if any)
    Dim strIPtext As String = ""        ' Text of the IP in HTML
    Dim strTrans As String = ""         ' Translation of last line (if available)

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return ""
      ' ============= PROFILING =============
      ProfileStart("GetIPfromConst", 1)
      ' =====================================
      ' Get the forest node
      If (Not GetForestNode(ndxThis, ndxForest)) Then
        ' Report error
        MsgBox("modAuto/GetIPfromConstituent: cannot find the <forest> of the antecedent")
        Return ""
      End If
      ' Also try to get the preceding forest node
      ndxBefore = ndxForest.PreviousSibling
      If (Not ndxBefore Is Nothing) Then
        loc_intSelNodeId = -1
        If (TraverseNodes(ndxBefore, "Syntax", strContext)) Then
          ' Now we have context!!!
          strContext &= "<br>" & vbCrLf
        End If
      End If
      ' Store the link of the selected ID
      loc_intSelNodeId = ndxThis.Attributes("Id").Value
      ' Recursively make the IP text in HTML
      If (TraverseNodes(ndxForest, "Syntax", strIPtext)) Then
        ' REturn the text as we found it
        strContext &= strIPtext
      End If
      ' Try to get a translation of this last sentence
      strTrans = ForestText(ndxForest, "eng")
      ' Do we have a translation?
      If (strTrans <> "") Then
        ' Add this in blue
        strContext &= "<p><font color='blue' size='2'>" & strTrans & "<font>" & vbCrLf
      End If
      ' ============= PROFILING =============
      ProfileEnd("GetIPfromConst", 1)
      ' =====================================
      ' Return the overall result
      Return strContext
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetIPfromConstituent error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetPGNlist
  ' Goal:   Get a list of all the PGN's in the settings
  ' History:
  ' 16-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetPGNlist() As String
    Dim dtrFound() As DataRow   ' Result of select
    Dim strList As String = ""  ' List of the results
    Dim intI As Integer         ' Counter

    Try
      ' Check if we already have a copy
      If (loc_strPgnList = "") Then
        ' validate
        If (tdlSettings Is Nothing) Then Return ""
        If (tdlSettings.Tables("Pronoun") Is Nothing) Then Return ""
        dtrFound = tdlSettings.Tables("Pronoun").Select("PGN <> 'unknown'", "PGN ASC")
        ' Make a list of the results
        For intI = 0 To UBound(dtrFound)
          ' Add item to the list
          AddSemiStack(strList, dtrFound(intI).Item("PGN"), True)
        Next intI
        ' Store the results
        loc_strPgnList = strList
      End If
      ' Return the result
      Return loc_strPgnList
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetPGNlist error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   UserAdaptPGN
  ' Goal:   Ask user to adapt the PGN, and if all goes well, then adapt it
  ' History:
  ' 16-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function UserAdaptPGN(ByRef ndxThis As XmlNode, ByRef strPgn As String, _
                               Optional ByVal bOnlyOne As Boolean = True) As Boolean
    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Get the current PGN
      strPgn = GetFeature(ndxThis, "NP", "PGN")
      ' Double check
      If (strPgn = "") Then strPgn = GetPGNlist()
      ' Ask user
      With frmPGN
        ' Set the default values
        .PGN = strPgn
        .Clause = GetIPfromConstituent(ndxThis)
        ' Look at the dialog result
        Select Case .ShowDialog
          Case DialogResult.OK
            ' Get the PGN result
            strPgn = .PGN
            ' Is this a good result?
            If (strPgn <> "") Then
              ' Do we allow only one?
              If (InStr(strPgn, ";") = 0) OrElse (Not bOnlyOne) Then
                ' Validate
                If (ndxThis Is Nothing) Then Return False
                ' Store this feature in the XML structure
                If (Not AddFeature(pdxCurrentFile, ndxThis, "NP", "PGN", strPgn)) Then Return False
                ' When we have come here, this means success
                Return True
              End If
            End If
        End Select
      End With
      ' In all other cases but success: return failure
      bInterrupt = True
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/UserAdaptPGN error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetLocation
  ' Goal:   Get the Location parameter of the <forest> IP of the indicated node
  ' History:
  ' 28-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetLocation(ByRef ndxThis As XmlNode) As String
    Dim ndxForest As XmlNode = Nothing  ' The forest node of the source
    Dim strIPtext As String = ""        ' Text of the IP in HTML

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return ""
      ' Get the forest node
      If (Not GetForestNode(ndxThis, ndxForest)) Then
        ' No real report is needed - there just are those that don't have a forest node I suppose...
        ' ================= DEBUGGING ==================
        If (bDebugging) Then Debug.Print("modAuto/GetLocation: cannot find the <forest> of the antecedent")
        ' ==============================================
        Return ""
      End If
      ' Does this forest have the correct attribute?
      If (ndxForest.Attributes("Location") Is Nothing) Then Return ""
      ' Return the appropriate value
      Return ndxForest.Attributes("Location").Value
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetLocation error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CopyFeatures
  ' Goal:   Copy relevant features from the antecedent [dtrAnt] to myself [dtrThis]
  ' History:
  ' 23-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CopyFeatures(ByRef dtrThis As DataRow, ByRef dtrAnt As DataRow, _
                                Optional ByVal bInferred As Boolean = False) As Boolean
    Dim strPGNsrc As String       ' Person/Number/Gender of the source
    Dim strPGNdst As String       ' Person/Number/Gender of the destination
    Dim strSrc(0 To 2) As String  ' Person, number and gender split up
    Dim strDst(0 To 2) As String  ' Person, number and gender split up
    Dim intI As Integer           ' Counter
    Dim bCopy As Boolean = True   ' Copy PGN features, unless there is a prohibition

    Try
      ' Validate
      If (dtrThis Is Nothing) OrElse (dtrAnt Is Nothing) Then Return False
      ' =========== DEBUGGING ==============
      'If (dtrAnt.Item("Id") = 25481) Then Stop
      ' ====================================
      ' Some copying can only be done if the relation is NOT inferred
      If (Not bInferred) Then
        ' 1 - Copy the [Equal] value of the antecedent to me
        AddSemiStack(dtrThis.Item("Equal"), dtrAnt.Item("Equal"), True)
        ' 2 - Copy the [Head] value of the antecedent to me
        AddSemiStack(dtrThis.Item("Equal"), dtrAnt.Item("Head"), True)
        ' 3 - Copy the [Id] value of the antecedent to me
        '     But ONLY if the relation is not "Inferred"
        AddSemiStack(dtrThis.Item("ChainIds"), dtrAnt.Item("Id"))
        ' 4 - Increment the "chain depth" of me
        dtrThis.Item("ChainNum") += 1
      End If
      ' 5 - If my own PGN is "empty" or less specific than I am, then copy the
      '      antecedent's PGN to me
      Select Case dtrThis.Item("PGN")
        Case "empty", "unknown"
          ' Copy antecedent's PGN
          ' ============= DEBUG ===========
          'If (dtrThis.Item("Id") = 3911) Then Stop
          ' ===============================
          dtrThis.Item("PGN") = dtrAnt.Item("PGN")
        Case Else
          ' Get source and destination agreement
          strPGNsrc = dtrThis.Item("PGN")
          strPGNdst = dtrAnt.Item("PGN")
          ' Distentangle the source
          If (Not SplitAgree(strPGNsrc, strSrc(0), strSrc(1), strSrc(2))) Then
            ' Something went wrong
            Return False
          End If
          ' Distentangle the antecedent
          If (Not SplitAgree(strPGNdst, strDst(0), strDst(1), strDst(2))) Then
            ' Something went wrong
            Return False
          End If
          ' Check for disagreement and so forth
          For intI = 0 To 2
            ' Check this PGN item
            If (strSrc(intI) <> strDst(intI)) Then
              ' There is disagreement, is this okay?
              If (Not DoLike(strSrc(intI), "empty|unknown")) Then
                bCopy = False
              End If
            End If
          Next intI
          ' Only copy if there are no prohibitions
          If (bCopy) Then
            ' Copy antecedent's PGN
            ' ============= DEBUG ===========
            'If (dtrThis.Item("Id") = 3911) Then Stop
            ' ===============================
            dtrThis.Item("PGN") = dtrAnt.Item("PGN")
          Else
            ' ============= DEBUGGING =========
            Stop
            ' ===============================
          End If
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CopyFeatures error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CopyChainId
  ' Goal:   Copy the ChainId feature from myself [ndxThis] to the antecedent [ndxAnt]
  ' History:
  ' 09-02-2012  ERK Created (derived from CopyFeatures)
  ' ------------------------------------------------------------------------------------
  Public Function CopyChainId(ByRef ndxThis As XmlNode, ByRef ndxAnt As XmlNode, _
                                ByVal strRefType As String) As Boolean
    Dim intChainId As Integer   ' Chain Id
    Dim strFeat As String       ' String value of feature
    Dim ndxWork As XmlNode      ' Working node
    Dim dtrItem() As DataRow    ' Result of select
    Dim dtrParent As DataRow    ' The chain with this id

    Try
      ' Get the ChainId of the source
      strFeat = CorefAttr(ndxThis, "ChainId")
      If (strFeat = "") Then Return False
      intChainId = CInt(strFeat)
      If (intChainId < 0) Then Return False
      ' Get the parent chain id
      If (tdlRefChain Is Nothing) Then
        dtrParent = Nothing
      Else
        dtrParent = tdlRefChain.Tables("Chain").Select("ChainId=" & intChainId)(0)
      End If
      ' Walk the antecedent chain
      Status("Chain id perculation...")
      ndxWork = ndxThis
      While (Not ndxWork Is Nothing) AndAlso _
            (DoLike(CorefAttr(ndxWork, "RefType"), strRefIdentity & "|" & strRefCross))
        ' Check the chain id feature
        If (CorefAttr(ndxWork, "ChainId") <> strFeat) Then
          ' Set the chain id
          AddFeature(pdxCurrentFile, ndxWork, "coref", "ChainId", strFeat)
          ' Check if we need to move the <Item> in the reference chain list
          If (Not tdlRefChain Is Nothing) Then
            ' Find the correct item
            dtrItem = tdlRefChain.Tables("Item").Select("eTreeId=" & ndxWork.Attributes("Id").Value)
            ' Check result
            If (dtrItem.Length > 0) Then
              ' Take the first one
              With dtrItem(0)
                ' Change the ChainId we belong to
                .Item("ChainId") = intChainId
                ' Set the parent
                .SetParentRow(dtrParent)
              End With
            End If
          End If
        End If
        ' Go to the next antecedent
        ndxWork = CorefDst(ndxWork)
      End While
      Status("done")
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CopyChainId error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   CopyFeaturesNode
  ' Goal:   Copy relevant features from the antecedent [ndxAnt] to myself [ndxThis]
  ' History:
  ' 25-11-2010  ERK Created (derived from CopyFeatures)
  ' ------------------------------------------------------------------------------------
  Public Function CopyFeaturesNode(ByRef ndxThis As XmlNode, ByRef ndxAnt As XmlNode, _
                                ByVal strRefType As String) As Boolean
    Dim dtrThis As DataRow        ' Myself
    Dim dtrAnt As DataRow         ' ANtecedent
    Dim strPGNsrc As String       ' Person/Number/Gender of the source
    Dim strPGNdst As String       ' Person/Number/Gender of the destination
    Dim strSrc(0 To 2) As String  ' Person, number and gender split up
    Dim strDst(0 To 2) As String  ' Person, number and gender split up
    Dim intI As Integer           ' Counter
    Dim intId As Integer          ' Index in stack
    Dim bCopy As Boolean = True   ' Copy PGN features, unless there is a prohibition

    Try
      ' Validate
      If (ndxThis Is Nothing) OrElse (ndxAnt Is Nothing) Then Return False
      ' Determine the correct source datarow
      intId = GetSrcStack(ndxThis)
      If (intId < 0) Then Return False
      dtrThis = tblSrcStack(intId)
      ' Determine the correct antecedent datarow
      intId = GetAntStack(ndxAnt)
      If (intId < 0) Then
        ' Okay, this is probably a manual addition -- so we can return TRUE, but not continue!
        Return True
      End If
      dtrAnt = tblAntStack(intId)
      ' =========== DEBUGGING ==============
      'If (dtrAnt.Item("Id") = 25481) Then Stop
      ' ====================================
      ' Some copying can only be done if the relation is NOT inferred
      If (strRefType = strRefIdentity) OrElse (strRefType = strRefCross) Then
        ' 1 - Copy the [Equal] value of the antecedent to me
        AddSemiStack(dtrThis.Item("Equal"), dtrAnt.Item("Equal"), True)
        ' 2 - Copy the [Head] value of the antecedent to me
        AddSemiStack(dtrThis.Item("Equal"), dtrAnt.Item("Head"), True)
        ' 3 - Copy the [Id] value of the antecedent to me
        '     But ONLY if the relation is not "Inferred"
        AddSemiStack(dtrThis.Item("ChainIds"), dtrAnt.Item("Id"))
        ' 4 - Increment the "chain depth" of me
        dtrThis.Item("ChainNum") += 1
      End If
      ' PGN copying may only be done if there is identity
      If (strRefType = strRefIdentity) Then
        ' 5 - If my own PGN is "empty" or less specific than I am, then copy the
        '      antecedent's PGN to me
        Select Case dtrThis.Item("PGN")
          Case "empty", "unknown"
            ' Copy antecedent's PGN
            ' ============= DEBUG ===========
            'If (dtrThis.Item("Id") = 3911) Then Stop
            ' ===============================
            dtrThis.Item("PGN") = dtrAnt.Item("PGN")
          Case Else
            ' Get source and destination agreement
            strPGNsrc = dtrThis.Item("PGN")
            strPGNdst = dtrAnt.Item("PGN")
            ' Check the source
            If (InStr(strPGNsrc, ";") > 0) Then
              ' Is this fully auto?
              If (bIsFullyAuto) Then
                ' Take the first one
                strPGNsrc = Split(strPGNsrc, ";")(0)
              Else
                ' Ask user to choose
                If (Not UserAdaptPGN(ndxThis, strPGNsrc)) Then Return False
              End If
              ' Adapt it in the dtr
              dtrThis.Item("PGN") = strPGNsrc
            End If
            ' Distentangle the source
            If (Not SplitAgree(strPGNsrc, strSrc(0), strSrc(1), strSrc(2))) Then
              ' Something went wrong
              Return False
            End If
            ' Check the antecedent
            If (InStr(strPGNdst, ";") > 0) Then
              ' Is this fully auto?
              If (bIsFullyAuto) Then
                ' Take the first one
                strPGNdst = Split(strPGNdst, ";")(0)
              Else
                ' Ask user to choose
                If (Not UserAdaptPGN(ndxAnt, strPGNdst)) Then Return False
              End If
              ' Adapt it in the dtr
              dtrThis.Item("PGN") = strPGNdst
            End If
            ' Distentangle the antecedent
            If (Not SplitAgree(strPGNdst, strDst(0), strDst(1), strDst(2))) Then
              ' Something went wrong
              Return False
            End If
            ' Check for disagreement and so forth
            For intI = 0 To 2
              ' Check this PGN item
              If (strSrc(intI) <> strDst(intI)) Then
                ' There is disagreement, is this okay?
                If (Not DoLike(strSrc(intI), "empty|unknown")) Then
                  bCopy = False
                End If
              End If
            Next intI
            ' Only copy if there are no prohibitions
            If (bCopy) Then
              ' Copy antecedent's PGN
              ' ============= DEBUG ===========
              'If (dtrThis.Item("Id") = 3911) Then Stop
              ' ===============================
              dtrThis.Item("PGN") = dtrAnt.Item("PGN")
            Else
              ' ============= DEBUGGING =========
              'Stop
              ' ===============================
            End If
        End Select
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CopyFeaturesNode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AutoReport
  ' Goal:   Add a line to the report made in the PDE editor's window of frmMain
  ' History:
  ' 17-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub AutoReport(ByVal strText As String)
    Dim bShowAutoLog As Boolean = frmMain.mnuViewAuto.Checked

    ' Check visibility of autolog textbox itself
    frmMain.tbAutoLog.Visible = bShowAutoLog
    ' Check visibility
    If (Not frmMain.panAutoLog.Visible) AndAlso (bShowAutoLog) Then frmMain.panAutoLog.Visible = True
    ' Add this line to the lines in [EdtPde]
    frmMain.tbAutoLog.AppendText(strText & vbCrLf)
    ' Make sure scrolling happens
    frmMain.tbAutoLog.ScrollToCaret()
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   AutoReportInit
  ' Goal:   Start the autocoreferencing report...
  ' History:
  ' 17-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub AutoReportInit()
    Dim bShowAutoLog As Boolean = frmMain.mnuViewAuto.Checked

    ' Check visibility of autolog textbox itself
    frmMain.tbAutoLog.Visible = bShowAutoLog
    ' Check visibility
    If (Not frmMain.panAutoLog.Visible) AndAlso (bShowAutoLog) Then frmMain.panAutoLog.Visible = True
    ' Add this line to the lines in [EdtPde]
    frmMain.tbAutoLog.Text = "Auto coreferencing on: " & Today & vbCrLf
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   AutoReportFinish
  ' Goal:   Finish the auto report
  ' History:
  ' 18-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub AutoReportFinish()
    ' Make invisible
    frmMain.panAutoLog.Visible = False
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   NodeInfo
  ' Goal:   Give relevant node information for the AutoReport
  ' History:
  ' 17-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function NodeInfo(ByRef ndxThis As XmlNode) As String
    Dim ndxForest As XmlNode = Nothing ' The forest's parent node

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return ""
      Select Case ndxThis.Name
        Case "eTree"
          ' Get the forest's parent
          If (Not GetForestNode(ndxThis, ndxForest)) Then
            ' Problem!
            Return ""
          End If
          ' Return the appropriate information
          Return ndxForest.Attributes("Location").Value & "[" & ndxThis.Attributes("Label").Value & _
            " " & NodeText(ndxThis, True) & "]"
        Case "forest"
          Return "[" & ndxThis.Attributes("Location").Value & "] " & ndxThis.SelectSingleNode("./div[@lang='org']").InnerText
        Case Else
          Return ""
      End Select
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/NodeInfo error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   NodeInfoPlusChildren
  ' Goal:   Give relevant node information including that of direct children
  '         Do this in HTML form
  ' History:
  ' 16-04-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function NodeInfoPlusChildren(ByRef ndxThis As XmlNode, Optional ByVal strHighL As String = "") As String
    Dim ndxForest As XmlNode = Nothing  ' The forest's parent node
    Dim ndxChild As XmlNode             ' Child node
    Dim colBack As New StringColl       ' What we return
    Dim strBack As String               ' What we return
    Dim intPos As Integer               ' pOsition in string

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return ""
      Select Case ndxThis.Name
        Case "eTree"
          ' Get the forest's parent
          If (Not GetForestNode(ndxThis, ndxForest)) Then
            ' Problem!
            Return ""
          End If
          ' Return the appropriate information
          colBack.Add(ndxForest.Attributes("Location").Value & "[<font color='blue' size='1'>" & _
                      ndxThis.Attributes("Label").Value & " </font>")
        Case "forest"
          colBack.Add("[" & ndxThis.Attributes("Location").Value & "] ")
      End Select
      ' Walk the direct children
      ndxChild = ndxThis.FirstChild
      While (Not ndxChild Is Nothing)
        ' Process this child
        If (ndxChild.Name = "eTree") Then
          ' Return the appropriate information
          colBack.Add(" [<font color='blue' size='1'>" & ndxChild.Attributes("Label").Value & _
                      " </font>" & VernToEnglish(NodeText(ndxChild, False)) & "] ")
        ElseIf (ndxChild.Name = "eLeaf") Then
          colBack.Add(" " & VernToEnglish(ndxChild.Attributes("Text").Value) & " ")
        End If
        ' Go to next child
        ndxChild = ndxChild.NextSibling
      End While
      Select Case ndxThis.Name
        Case "eTree"
          colBack.Add("] ")
        Case "forest"
          colBack.Add("] ")
      End Select
      ' Combine the result
      strBack = colBack.Text
      ' Do we have highlighting?
      If (strHighL <> "") Then
        ' Find the highlighted element
        intPos = InStr(strBack, strHighL)
        If (intPos > 0) Then
          ' Highlight it
          strBack = Left(strBack, intPos - 1) & "<b>" & strHighL & "</b>" & Mid(strBack, intPos + strHighL.Length)
        End If
      End If
      ' Return the result
      Return strBack
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/NodeInfoPlusChildren error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsPossibleAntecedent
  ' Goal:   Determine whether this NP can be referred to
  ' Note:   NP's that can NOT be referred to are:
  '         - appositive NP's
  '         - quantificational NP's from a preceding IP (which this one will be)
  ' History:
  ' 09-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function IsPossibleAntecedent(ByRef dtrThis As DataRow, ByRef bIsLocal As Boolean) As Boolean
    Dim ndxThis As XmlNode = Nothing  ' This node's representation in XML

    Try
      ' Initialise [bIsLocal]
      bIsLocal = False
      ' Validate
      If (dtrThis Is Nothing) Then Return False
      ' Check for quantificational NP -- But they do seem to be able to serve as antecedent!
      'If (dtrThis.Item("NPtype") Like "Q*") Then
      '  ' This is an NP that is available locally within this <forest>
      '  bIsLocal = True
      '  Return True
      'End If
      ' Traces should remain local!!
      If (dtrThis.Item("NPtype") = "Trace") AndAlso (Left(dtrThis.Item("Vern"), 3) = "*T*") Then
        ' Make sure this stays within one <forest>
        bIsLocal = True
      ElseIf (dtrThis.Item("Coref") = "NewVar") Then
        ' ANything marked as "newvar" should also stay locally
        bIsLocal = True
      End If
      ' Check for appositive NP
      If (dtrThis.Item("Label") Like "*PRN*") Then Return False
      ' Find out which node I am
      ndxThis = IdToNode(dtrThis.Item("Id"))
      ' Double check
      If (ndxThis Is Nothing) Then
        MsgBox("modAuto/IsPossibleAntecedent: could not determine antecedent of node " & dtrThis.Item("Id"))
        Return False
      End If
      ' Check if this has received an "Inert" coreference label
      If (CorefAttr(ndxThis, "RefType") Like strRefInert) Then Return False
      'If (dtrThis.Item("Coref") Like strRefInert) Then Return False
      ' By default: we are okay
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/IsPossibleAntecedent error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   GetSrcFeatVectorSmall
  ' Goal:   Get the values for the source-node feature vector as used for ICHL21
  '         Source node:
  '       	1   Label:      function part of ndxThis.Label
  '         2.  HeadCat:    phrasal category of head
  '     	  3.  Anchor:     does this NP have an anchor or not?
  '     	  4.  PMcat:      phrasal category of post-modifier (if any)
  '         5.  NPtype:     my own NP type
  '         Other:
  '         6.  HeadBefore: ndxThis.Head has occurred before
  ' History:
  ' 01-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetSrcFeatVectorSmall(ByRef ndxThis As XmlNode, ByRef arSrcFeat() As String) As Boolean
    Dim ndxWork As XmlNode = Nothing  ' Working node
    Dim ndxHead As XmlNode = Nothing  ' Head node
    Dim ndxPost As XmlNode = Nothing  ' Postmodifier
    Dim ndxLeaf As XmlNode            ' Leaf node
    Dim strHead As String             ' Head noun
    Dim intVectorSize As Integer = 6

    Try
      ' Validate
      If (ndxThis Is Nothing) OrElse (ndxThis.Name <> "eTree") Then Return False
      ' ====================== DEBUG
      ' If (ndxThis.Attributes("Id").Value = "212") Then Stop
      ' ========================================
      ' Get the head of this item
      If (Not GetNPhead(ndxThis, ndxHead, False)) Then Return False
      ' Get the head noun
      ndxLeaf = ndxHead.SelectSingleNode("./descendant::eLeaf[@Type='Vern']")
      If (ndxLeaf Is Nothing) Then
        strHead = "-"
      ElseIf (DoLike(ndxHead.Attributes("Label").Value, strLexicalNPhead)) Then
        strHead = ndxLeaf.Attributes("Text").Value
      Else
        strHead = "-"
      End If
      ' Make room
      ReDim arSrcFeat(0 To intVectorSize - 1)
      ' Start filling with features:
      ' (1) Functional part of my own label
      arSrcFeat(0) = LabelCategory(ndxThis, "function")
      ' (2) Phrasal category of head
      arSrcFeat(1) = LabelCategory(ndxHead, "phrase")
      ' (3) Presence of anchor
      If (ndxThis.SelectSingleNode("child::eTree[1][tb:matches(@Label, 'PRO$|NPR$|NPRS$')]", conTb) Is Nothing) Then
        arSrcFeat(2) = "-"
      Else
        arSrcFeat(2) = "A"
      End If
      ' (4) Phrasal category of post-modifier (if any)
      ndxPost = ndxHead.SelectSingleNode("following-sibling::eTree[not(child::eLeaf/@Type = 'Star') and not(@Label = 'CODE')][1]")
      If (ndxPost Is Nothing) Then
        arSrcFeat(3) = "-"
      Else
        arSrcFeat(3) = LabelCategory(ndxPost, "phrase")
      End If
      ' (5) NPtype of myself
      arSrcFeat(4) = GetFeature(ndxThis, "NP", "NPtype")
      If (arSrcFeat(4) = "") Then arSrcFeat(4) = "-"
      ' (6) Head noun has occurred before
      If (strHead = "-") OrElse Not (colHeadBefore.Exists(strHead)) Then
        arSrcFeat(5) = "0"
      Else
        arSrcFeat(5) = "1"
      End If
      ' Add the current head to the collection
      If (strHead <> "") Then colHeadBefore.AddUnique(strHead)
      ' By default: we are okay
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetSrcFeatVectorSmall error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetSrcFeatVector
  ' Goal:   Get the values for the source-node feature vector
  '         Source node:
  '       	1   Label:      ndxThis.Label
  '     	  2.  NPtype
  '     	  3.  GrRole
  '     	  4.  Starred:    IsStarred(ndxThis)
  '         Some child:
  '     	  5.  HasFreeRel: SomeChild(ndxThis, “CP-FRL*”)
  '     	  6.  HasCPrel:   SomeChild(ndxThis, “CP-REL*”)
  '         First child:
  '     	  7.  FirstChild:     ndxThis.Child(1).Label
  '     	  8.  FirstEleafType: ndxThis.Eleaf(1).Type
  '     	  9.  FirstEleafText: ndxThis.Eleaf(1).Text
  '         Other children
  '         10. Child2:     ndxThis.Child(2).Label
  '         11. Child3:     ndxThis.Child(3).Label
  '         12. Child4:     ndxThis.Child(4).Label
  '         Some sister:
  '     	  13. BeSister:  HasSister(ndxThis, strFiniteBe)
  '     	  14. SbjSister: HasSister(ndxThis, strSubjectType)
  '     	  15. VbSister:  HasSister(ndxThis, ‘V*’)
  '     	  16. CpSister:  HasSister(ndxThis, ‘CP*’)
  '         Subject sister:
  '     	  17. SbjType:   NPtype(GetSister(ndxThis,strSubjectType))
  '     	  18. SbjText:   GetSister(ndxThis,strSubjecType).FirstEleaf.Text
  '         Other:
  '         19. HeadBefore: ndxThis.Head has occurred before
  '         20. HeadText:   ndxThis.Head.FirstEleaf.Text
  '         21. PGN:        ndxThis.PGN
  '         22. Len:        number of words in the NP
  '         23. Neg:        presence of negator in NP
  '         24. Mood:       mood of clause in which I am (interrogative, imperative, declarative, subjunctive)
  '         25. Speech:     NP is part of direct speech
  '         26. HasPreM:    NP has a pre-modifier (non-D element before the head)
  ' History:
  ' 01-04-2013  ERK Created
  ' 15-04-2013  ERK Added [bSmall] option
  ' ------------------------------------------------------------------------------------
  Private Function GetSrcFeatVector(ByRef ndxThis As XmlNode, ByRef arSrcFeat() As String) As Boolean
    Dim ndxWork As XmlNode = Nothing    ' Working node
    Dim ndxHead As XmlNode = Nothing    ' Head node
    Dim ndxLeaf As XmlNode              ' Leaf node
    Dim ndxClause As XmlNode = Nothing  ' Clausal head (nearest IP or CP)
    Dim strHead As String               ' Head noun
    Dim intVectorSize As Integer = 26

    Try
      ' Validate
      If (ndxThis Is Nothing) OrElse (ndxThis.Name <> "eTree") Then Return False
      ' Get the head of this item
      ' Get the head of this item
      If (Not GetNPhead(ndxThis, ndxHead, False)) Then Return False
      ' Get the head noun
      ndxLeaf = ndxHead.SelectSingleNode("./descendant::eLeaf[@Type='Vern']")
      If (ndxLeaf Is Nothing) Then
        strHead = "-"
      ElseIf (DoLike(ndxHead.Attributes("Label").Value, strLexicalNPhead)) Then
        strHead = ndxLeaf.Attributes("Text").Value
      Else
        strHead = "-"
      End If
      ' Make room
      ReDim arSrcFeat(0 To intVectorSize - 1)
      ' Start filling with features:
      ' (1) My own label
      arSrcFeat(0) = ndxThis.Attributes("Label").Value
      ' (2) NPtype
      arSrcFeat(1) = GetFeature(ndxThis, "NP", "NPtype")
      If (arSrcFeat(1) = "") Then
        If (DoLike(ndxThis.Attributes("Label").Value, "NPR*|NR*")) Then
          arSrcFeat(1) = "Proper"
        Else
          arSrcFeat(1) = "-"
        End If
      End If
      ' =========== DEBUG ===============
      If (arSrcFeat(1) = "") Then
        Stop
        arSrcFeat(1) = GetFeature(ndxThis, "NP", "NPtype")
      End If
      ' =================================
      ' (3) GrRole
      arSrcFeat(2) = GetFeature(ndxThis, "NP", "GrRole")
      '' =========== DEBUG ===============
      'If (arSrcFeat(2) = "") Then
      '  Stop
      '  arSrcFeat(2) = GetFeature(ndxThis, "NP", "GrRole")
      'End If
      '' =================================
      If (arSrcFeat(2) = "") Then
        arSrcFeat(2) = "-"
      End If
      ' (4) Starred (1) or not (0)?
      arSrcFeat(3) = IIf(IsStarred(ndxThis), "1", "0")
      ' (5) Free relative (1) or not(0)?
      arSrcFeat(4) = IIf(SomeChild(ndxThis, "CP-FRL*", ndxWork), "1", "0")
      ' (6) CP relative (1) or not(0)?
      arSrcFeat(5) = IIf(SomeChild(ndxThis, "CP-REL*", ndxWork), "1", "0")
      ' Get first child
      ndxWork = ndxThis.SelectSingleNode("./child::eTree[not(tb:matches(string(@Label), 'CODE|,|.'))]", conTb)
      If (ndxWork Is Nothing) Then
        arSrcFeat(6) = "0"
        ' Is first child an eleaf?
        ndxWork = ndxThis.SelectSingleNode("./child::eLeaf[1]")
        If (ndxWork Is Nothing) Then
          arSrcFeat(7) = "0"
          arSrcFeat(8) = "0"
        Else
          arSrcFeat(7) = ndxWork.Attributes("Type").Value
          arSrcFeat(8) = LCase(ndxWork.Attributes("Text").Value)
          ' Adapt leaf-text if needed
          If (Left(arSrcFeat(8), 3) = "*T*") Then arSrcFeat(8) = "*"
        End If
      Else
        arSrcFeat(6) = ndxWork.Attributes("Label").Value
        'Get the first <eLeaf> descendant in this NP
        ndxWork = ndxThis.SelectSingleNode("./descendant::eLeaf[1]")
        If (ndxWork Is Nothing) Then
          arSrcFeat(7) = "0"
          arSrcFeat(8) = "0"
        Else
          arSrcFeat(7) = ndxWork.Attributes("Type").Value
          arSrcFeat(8) = LCase(ndxWork.Attributes("Text").Value)
          ' Adapt leaf-text if needed
          If (Left(arSrcFeat(8), 3) = "*T*") Then arSrcFeat(8) = "*"
        End If
      End If
      ' Get second child
      ndxWork = ndxThis.SelectSingleNode("./child::eTree[not(tb:matches(string(@Label), 'CODE|,|.'))][2]", conTb)
      If (ndxWork Is Nothing) Then
        arSrcFeat(9) = "0"
      Else
        arSrcFeat(9) = ndxWork.Attributes("Label").Value
      End If
      ' Get third child
      ndxWork = ndxThis.SelectSingleNode("./child::eTree[not(tb:matches(string(@Label), 'CODE|,|.'))][3]", conTb)
      If (ndxWork Is Nothing) Then
        arSrcFeat(10) = "0"
      Else
        arSrcFeat(10) = ndxWork.Attributes("Label").Value
      End If
      ' Get fourth child
      ndxWork = ndxThis.SelectSingleNode("./child::eTree[not(tb:matches(string(@Label), 'CODE|,|.'))][4]", conTb)
      If (ndxWork Is Nothing) Then
        arSrcFeat(11) = "0"
      Else
        arSrcFeat(11) = ndxWork.Attributes("Label").Value
      End If
      ' (13) Sister of type be
      arSrcFeat(12) = IIf(HasSister(ndxThis, strFiniteBe), "1", "0")
      ' (14) Sister of type be
      arSrcFeat(13) = IIf(HasSister(ndxThis, strSubjectType), "1", "0")
      ' (15) Sister of type be
      arSrcFeat(14) = IIf(HasSister(ndxThis, "V*"), "1", "0")
      ' (16) Sister of any CP-kind --> should actually exclude CP-FRL and CP-REL...
      arSrcFeat(15) = IIf(HasSister(ndxThis, "CP*"), "1", "0")
      ' (17) Get subject sister
      If (GetSister(ndxThis, strSubjectType, ndxWork)) Then
        arSrcFeat(16) = GetFeature(ndxWork, "NP", "NPtype")
        If (ndxWork.SelectSingleNode("./descendant::eLeaf") Is Nothing) Then
          arSrcFeat(17) = "0"
        Else
          arSrcFeat(17) = ndxWork.SelectSingleNode("./descendant::eLeaf").Attributes("Text").Value
        End If
      Else
        arSrcFeat(16) = "0" : arSrcFeat(17) = "0"
      End If
      ' (19) Head noun has occurred before
      If (strHead = "-") OrElse Not (colHeadBefore.Exists(strHead)) Then
        arSrcFeat(18) = "0"
      Else
        arSrcFeat(18) = "1"
      End If
      ' (20) Text of the head 
      arSrcFeat(19) = strHead
      ' (21) PGN of the NP
      arSrcFeat(20) = GetFeature(ndxThis, "NP", "PGN")
      If (arSrcFeat(20) = "") Then arSrcFeat(20) = "-"
      ' (22) Len = number of words
      arSrcFeat(21) = ndxThis.SelectNodes("./descendant::eLeaf[@Type = 'Vern']").Count
      ' (23) Presence of negator child
      arSrcFeat(22) = IIf(SomeChild(ndxThis, "NEG*", ndxWork), "1", "0")
      ' Get the clausal head in preparation of (24) mood and (25) speech
      ndxClause = GetIpAncestor(ndxThis)
      If (ndxClause Is Nothing) Then
        ' There is no clausal head
        arSrcFeat(23) = "0" : arSrcFeat(24) = "0"
      Else
        ' We have a clausal head
        ' (24) Get the mood 
        arSrcFeat(23) = GetClauseMood(ndxClause)
        If (arSrcFeat(23) = "") Then arSrcFeat(23) = "-"
        ' (25) Check if this is direct speech (simple check)
        arSrcFeat(24) = IIf(ndxClause.Attributes("Label").Value Like "*SPE*", "1", "0")
      End If
      ' (26) Check for pre-modifiers within the NP (non "D"-like constituents preceding the head, excluding skippables)
      If (ndxHead.SelectSingleNode("./preceding-sibling::eTree[not(tb:matches(string(@Label), 'CODE|,|.|D|D^*'))]", conTb) Is Nothing) Then
        arSrcFeat(25) = "0"
      Else
        arSrcFeat(25) = "1"
      End If
      ' Add the current head to the collection
      If (strHead <> "") Then colHeadBefore.AddUnique(strHead)
      ' By default: we are okay
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetSrcFeatVector error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetBasicFeatVector
  ' Goal:   Get basic values for this node
  '       	1   Label:      ndxThis.Label
  '     	  2.  NPtype
  '     	  3.  GrRole
  '         4.  PGN
  '     	  5.  FirstEleafText: ndxThis.Eleaf(1).Text
  ' History:
  ' 02-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetBasicFeatVector(ByRef ndxThis As XmlNode, ByRef arBasicFeat() As String) As Boolean
    Dim ndxWork As XmlNode = Nothing  ' Working node
    Dim intVectorSize As Integer = 5

    Try
      ' Validate
      If (ndxThis Is Nothing) OrElse (ndxThis.Name <> "eTree") Then Return False
      ' Make room
      ReDim arBasicFeat(0 To intVectorSize - 1)
      ' Start filling with features:
      ' (1) My own label
      arBasicFeat(0) = ndxThis.Attributes("Label").Value
      ' (2) NPtype
      arBasicFeat(1) = GetFeature(ndxThis, "NP", "NPtype")
      ' (3) GrRole
      arBasicFeat(2) = GetFeature(ndxThis, "NP", "GrRole")
      ' (4) Person-Gender-Number
      arBasicFeat(3) = GetFeature(ndxThis, "NP", "PGN")
      ' (5) First text of <eLeaf>
      ndxWork = ndxThis.SelectSingleNode("./descendant::eLeaf[@Type='Vern'][1]")
      If (ndxWork Is Nothing) Then
        arBasicFeat(4) = "-"
      Else
        arBasicFeat(4) = ndxWork.Attributes("Text").Value
      End If
      ' By default: we are okay
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetBasicFeatVector error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetStackFeatVector
  ' Goal:   Get all necessary feature vector values from the [StackEl] item
  ' History:
  ' 02-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetStackFeatVector(ByRef dtrThis As DataRow, ByRef arBasicFeat() As String) As Boolean
    Dim ndxWork As XmlNode = Nothing  ' Working node
    Dim intVectorSize As Integer = 11

    Try
      ' Validate
      If (dtrThis Is Nothing) Then Return False
      ' Make room
      ReDim arBasicFeat(0 To intVectorSize - 1)
      ' Start filling with features:
      With dtrThis
        arBasicFeat(0) = .Item("Person")
        arBasicFeat(1) = .Item("Gender")
        arBasicFeat(2) = .Item("Number")
        arBasicFeat(3) = .Item("Label")
        arBasicFeat(4) = .Item("GrRole")
        arBasicFeat(5) = .Item("NPtype")
        arBasicFeat(6) = .Item("IsSpeech")
        arBasicFeat(7) = .Item("IPances")
        arBasicFeat(8) = .Item("Head")
        arBasicFeat(9) = .Item("First")
        arBasicFeat(10) = .Item("HasCoref")
      End With
      ' (1) My own label
      ' By default: we are okay
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetStackFeatVector error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddTrainingData
  ' Goal:   Add trainingsdata for this combination of:
  '         (1) source node + feature vector
  '         (2) set of antecedent nodes with their feature vectors
  ' History:
  ' 01-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function AddTrainingData(ByRef ndxSrc As XmlNode, ByRef dtrSrc As DataRow, ByRef tblAnt As DataTable) As Boolean
    Dim strCns As String        ' Name of constraint
    Dim strState As String      ' Referential state
    Dim dtrCns() As DataRow     ' Ordered set of constraints
    Dim arSrcFeat(0) As String  ' Array of source-node features
    Dim arFeat() As String      ' Feature vector
    Dim arSrcBasic(0) As String ' Basic features of source
    Dim arAntBasic(0) As String ' Basic features of antecedent
    Dim ndxAnt As XmlNode       ' Antecedent node
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter
    Dim intDef As Integer = -1  ' Default value for constraint, signifying it has NO value
    Dim intAntId As Integer     ' ID of the antecedent (if any)
    Dim arCns() As String = {"Disjoint", "EqHead", "IPdist", "Ambiguity", "NoCataphore"}

    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (tblAnt Is Nothing) OrElse (tblAnt.Rows.Count = 0) Then Return False
      ' Get source-node features into the array
      If (Not GetSrcFeatVector(ndxSrc, arSrcFeat)) Then Return False
      ' Get basic source node feature values for coreference resolution
      If (Not GetStackFeatVector(dtrSrc, arSrcBasic)) Then Return False
      ' Get the referential state of this item
      strState = GetFeature(ndxSrc, "coref", "RefType")
      If (strState = "CrossSpeech") Then strState = "Identity"
      ' Get the antecedent ID (if any)
      intAntId = CorefDstId(ndxSrc)
      ' Action depends on state
      Select Case strState
        Case "Identity", "Inferred"
          ' We should output the feature vector and the resolution
          AppendTrainingState(arSrcFeat, "Linked")
          ' Initialise the constraints
          dtrCns = tblConstraint.Select("", "Level ASC, Name ASC")
          ' Make space for feature vector
          ReDim arFeat(0 To arCns.Length - 1)
          ' Walk all possible antecedents
          For intI = 0 To tblAnt.Rows.Count - 1
            ' Get the antecedent node
            ndxAnt = IdToNode(tblAnt.Rows(intI).Item("Id"))
            ' Get basic features of antecedent for coreference resolution
            If (Not GetStackFeatVector(tblAnt.Rows(intI), arAntBasic)) Then Return False
            ' Get several constraint values
            For intJ = 0 To arCns.Length - 1
              ' Get the name of the current constraint
              strCns = arCns(intJ)
              ' Get the value for this constraint
              If (Not ConstraintValue(strCns, dtrSrc, tblAnt.Rows(intI), arFeat(intJ))) Then Return False
            Next intJ
            '' Walk all the constraints and calculate their values
            'For intJ = 0 To dtrCns.Count - 1
            '  ' Get the name of the current constraint
            '  strCns = dtrCns(intJ).Item("Name")
            '  ' Get the value for this constraint
            '  If (Not ConstraintValue(strCns, dtrSrc, tblAnt.Rows(intI), arFeat(intJ))) Then Return False
            'Next intJ
            ' Is this antecedent the correct one?
            If (intAntId > 0) AndAlso (tblAnt.Rows(intI).Item("Id") = intAntId) Then
              ' Positive ID
              AppendTrainingCoref(arSrcBasic, arAntBasic, arFeat, strState)
            Else
              ' Default negative ID
              AppendTrainingCoref(arSrcBasic, arAntBasic, arFeat, "none")
            End If
          Next intI
        Case "New", "NewVar", "Assumed", "Inert"
          ' We should output the feature vector and the resolution
          AppendTrainingState(arSrcFeat, strState)
        Case Else
          ' Find out what the state is, and what we need to do with it
          Stop
          HandleErr("modAuto/AddTrainingsData: state [" & strState & "] is unknown")
          Return False
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AddTrainingData error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   OneRefLearnFile
  ' Goal:   Train referential state determining from one Psdx file
  ' History:
  ' 09-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function OneRefLearnFile(ByVal strFile As String, ByVal bSmall As Boolean) As Boolean
    Dim ndxFor As XmlNode = Nothing   ' Forest node
    Dim ndxThis As XmlNode = Nothing  ' Current node
    Dim ndxList As XmlNodeList        ' List of <eLeaf> nodes
    Dim strLabel As String = ""       ' Output label of a constituent
    Dim strChildren As String = ""    ' Children of a constituent in semi-stack
    Dim strPartial As String = ""     ' Partial tree seen from me as center
    Dim intI As Integer               ' Counter
    Dim intPtc As Integer             ' Percentage
    Dim intCount As Integer           ' Number of child nodes
    'Dim intNum As Integer             ' Number of child nodes retrieved
    Dim pdxFile As XmlDocument = Nothing  ' Document we are working on

    Try
      ' Validate
      If (Not IO.File.Exists(strFile)) Then Return False
      ' Read the file as an XML document
      Status("Reading [" & IO.Path.GetFileNameWithoutExtension(strFile) & "]...")
      ' Try read this file into an XML structure
      If (Not ReadXmlDoc(strFile, pdxFile)) Then
        ' There was an error
        Status("Unable to read file: " & strFile)
        Return False
      End If
      ' Walk all the forests in the file
      ' Set the current file
      pdxCurrentFile = pdxFile
      ' Get the correct period of this file
      strCurrentPeriod = GetPeriod(pdxCurrentFile)
      ' Look for the first <forest> element
      If (Not GetFirstForest(pdxFile, ndxFor)) Then Logging("Cannot find first forest") : Return False
      ' Determine how many elements there are potentially
      If (ndxFor Is Nothing) Then
        intCount = 0
      ElseIf (ndxFor.ParentNode Is Nothing) Then
        intCount = 0
      Else
        intCount = ndxFor.ParentNode.ChildNodes.Count
      End If
      ' Other initialisations: clear the collection heads that have occurred "before"
      colHeadBefore.Clear()
      ' Walk through all the <forest> elements FORWARD
      While (Not ndxFor Is Nothing)
        ' Check if we are not being interrupted
        If (bInterrupt) Then Return False
        ' Show where we are (how do we KNOW where we are?)
        intPtc = (ndxFor.Attributes("forestId").Value) * 100 \ intCount
        Status("Learning [" & IO.Path.GetFileNameWithoutExtension(strFile) & "] " & intPtc & "%", intPtc)
        ' Make sure this is a forest and not something else
        If (ndxFor.Name = "forest") Then
          ' Check if this is the first forest of a section
          If (ndxFor.Attributes("Section") IsNot Nothing) Then
            ' Reset the HeadBefore collection
            colHeadBefore.Clear()
          End If
          ' Visit all the <eTree> constituents of this forest that have a referential state
          ndxList = ndxFor.SelectNodes("./descendant::eTree[not(tb:matches(@Label, 'ADV*|NP-MSR*|NP-ADV*')) " & _
                                       "and (child::fs/child::f/@name='RefType') " & _
                                       "and not(child::eLeaf[@Type = 'Star']) ]", conTb)
          For intI = 0 To ndxList.Count - 1
            ' Produce a feature vector for this constituent
            If (Not AddTrainingState(ndxList(intI), bSmall)) Then Return False
          Next intI
        End If
        ' Go to the next forst
        ndxFor = ndxFor.NextSibling
      End While
      ' Return success 
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modAuto/OneRefLearnFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddTrainingState
  ' Goal:   Add trainingsdata of type "state" for this:
  '         (1) source node + feature vector
  ' History:
  ' 09-04-2013  ERK Created
  ' 12-08-2013  ERK Add learned output to the tdlResults dataset
  ' ------------------------------------------------------------------------------------
  Private Function AddTrainingState(ByRef ndxSrc As XmlNode, ByVal bSmall As Boolean) As Boolean
    Dim strState As String      ' Referential state
    Dim arSrcFeat(0) As String  ' Array of source-node features
    Dim arFname() As String     ' The array with the feature names
    Dim intDef As Integer = -1  ' Default value for constraint, signifying it has NO value

    Try
      ' Validate
      If (ndxSrc Is Nothing) Then Return False
      ' Get source-node features into the array
      If (bSmall) Then
        If (Not GetSrcFeatVectorSmall(ndxSrc, arSrcFeat)) Then Return False
        arFname = arFnameSmall
      Else
        If (Not GetSrcFeatVector(ndxSrc, arSrcFeat)) Then Return False
        arFname = arFnameLarge
      End If
      ' Get the referential state of this item
      strState = GetFeature(ndxSrc, "coref", "RefType")
      If (strState = "CrossSpeech") OrElse (strState = "BoundAnaphor") Then strState = "Identity"
      ' Action depends on state
      Select Case strState
        Case "Identity", "Inferred"
          ' We should output the feature vector and the resolution
          ' OLD: AppendRefLearn(arSrcFeat, "Linked")
          AppendRefLearn(arSrcFeat, strState)
          If (Not AppendRefStateToResults(arFname, arSrcFeat, strState, ndxSrc)) Then Return False
        Case "New", "NewVar", "Assumed", "Inert"
          ' We should output the feature vector and the resolution
          AppendRefLearn(arSrcFeat, strState)
          If (Not AppendRefStateToResults(arFname, arSrcFeat, strState, ndxSrc)) Then Return False
        Case Else
          ' Find out what the state is, and what we need to do with it
          Stop
          HandleErr("modAuto/AddTrainingState: state [" & strState & "] is unknown")
          Return False
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AddTrainingState error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AppendRefStateToResults
  ' Goal:   Append referential state learning data to the tdlResults database
  ' History:
  ' 12-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function AppendRefStateToResults(ByRef arFname() As String, ByRef arFeat() As String, ByVal strState As String, _
                                           ByRef ndxThis As XmlNode) As Boolean
    'Dim dtrNew As DataRow   ' One row
    'Dim dtrFeat As DataRow  ' One feature datarow
    Dim ndxFor As XmlNode   ' The forest
    Dim intI As Integer     ' Counter

    Try
      ' Validate
      ' If (tdlResults Is Nothing) Then Return False
      If (ndxThis Is Nothing) Then Return False
      If (ndxThis.Name <> "eTree") Then Return False
      ' Get my forest
      ndxFor = ndxThis.SelectSingleNode(".//ancestor::forest[1]")
      ' Start adding the results row
      intCrpResultsId += 1
      colCrpResults.Add(" <Result ResId=" & """" & intCrpResultsId & """" & _
                        " File=" & """" & ndxFor.Attributes("File").Value & ".psdx" & """" & _
                        " TextId=" & """" & ndxFor.Attributes("TextId").Value & """" & _
                        " Search=" & """" & ndxFor.Attributes("Location").Value & """" & _
                        " Cat=" & """" & strState & """" & _
                        " forestId=" & """" & ndxFor.Attributes("forestId").Value & """" & _
                        " eTreeId=" & """" & ndxThis.Attributes("Id").Value & """" & _
                        " Notes=" & """" & "-" & """" & _
                        " >")
      colCrpResults.Add("  <Text>" & XmlEscape(NodeText(ndxThis)) & "</Text>")
      colCrpResults.Add("  <Psd>" & XmlEscape(NodeSyntax(ndxThis)) & "</Psd>")
      ' Add all the features available in [arFeat]
      For intI = 0 To arFeat.Length - 1
        ' Add this feature name + value
        colCrpResults.Add("  <Feature Name=" & """" & arFname(intI) & """" & _
                          " Value=" & """" & XmlEscape(arFeat(intI)).Replace("'", "''") & """" & " />")
      Next intI
      ' Finish this line
      colCrpResults.Add(" </Result>")

      '' Create a new row for results
      'dtrNew = AddOneDataRow(tdlResults, "Result", "ResId", "")
      '' Add the information I know for this particular entry
      'With dtrNew
      '  .Item("File") = ndxFor.Attributes("File").Value & ".psdx"
      '  .Item("TextId") = ndxFor.Attributes("TextId").Value
      '  .Item("Search") = ndxFor.Attributes("Location").Value
      '  .Item("Cat") = strState
      '  .Item("forestId") = ndxFor.Attributes("forestId").Value
      '  .Item("eTreeId") = ndxThis.Attributes("Id").Value
      '  .Item("Notes") = "-"
      '  .Item("Psd") = NodeSyntax(ndxThis)
      '  .Item("Text") = NodeText(ndxThis)
      'End With
      '' Add all the features available in [arFeat]
      'For intI = 0 To arFeat.Length - 1
      '  ' Add this feature name + value
      '  dtrFeat = AddOneDataRowWithParent(tdlResults, "Feature", "", dtrNew)
      '  With dtrFeat
      '    .Item("Name") = arFname(intI)
      '    .Item("Value") = arFeat(intI)
      '  End With
      'Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AppendRefStateToResults error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitTraining
  ' Goal:   Initialise training data output
  ' History:
  ' 01-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function InitTraining(ByVal strType As String) As Boolean
    ' Dim strFile As String   ' Name of output file

    Try
      ' Check which one needs to be initialized
      Select Case strType
        Case "reflearn"
          ' Determine the name of the output file
          strCurrentTrainingRefLearn = GetDocDir() & "\RefStateDir"
          ' Clear this file
          IO.File.WriteAllText(strCurrentTrainingRefLearn & ".train", "")
          IO.File.WriteAllText(strCurrentTrainingRefLearn & ".test", "")
          ' Clear stacks
          colCurrentRefLearnTest.Clear()
          colCurrentRefLearnTrain.Clear()
        Case "state"
          ' Determine the name of the output file
          strCurrentTrainingState = GetDocDir() & "\" & IO.Path.GetFileNameWithoutExtension(strCurrentFile) & _
            "_s" & intCurrentSection + 1 & "_State.txt"
          ' Clear this file
          IO.File.WriteAllText(strCurrentTrainingState, "")
          ' Clear stacks
          colCurrentTrainingState.Clear()
        Case "coref"
          ' Determine the name of the output file
          strCurrentTrainingCoref = GetDocDir() & "\" & IO.Path.GetFileNameWithoutExtension(strCurrentFile) & _
            "_s" & intCurrentSection + 1 & "_Coref.txt"
          ' Clear this file
          IO.File.WriteAllText(strCurrentTrainingCoref, "")
          ' Clear stacks
          colCurrentTrainingCoref.Clear()
      End Select
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/InitTraining error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AppendRefLearn
  ' Goal:   Append feature vector + strState resolution to the trainings data
  '           which serves to identify what kind of non-coreferring state a constituent has
  ' History:
  ' 01-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function AppendRefLearn(ByRef arFeatVect() As String, ByVal strState As String) As Boolean
    Dim strLine As String
    Dim intPtcTrain As Integer = 70

    Try
      ' Determine the line
      strLine = strCurrentPeriod & vbTab & Join(arFeatVect, vbTab) & vbTab & strState
      ' Training or testing?
      If (Rnd() * 100 < intPtcTrain) Then
        ' Training
        colCurrentRefLearnTrain.Add(strLine)
      Else
        ' Testing
        colCurrentRefLearnTest.Add(strLine)
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AppendRefLearn error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AppendTrainingState
  ' Goal:   Append feature vector + strState resolution to the trainings data
  '           which serves to identify what kind of non-coreferring state a constituent has
  ' History:
  ' 01-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function AppendTrainingState(ByRef arFeatVect() As String, ByVal strState As String) As Boolean
    Try
      ' Add one line of STATE learning
      colCurrentTrainingState.Add(strCurrentPeriod & vbTab & Join(arFeatVect, vbTab) & vbTab & strState)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AppendTrainingState error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AppendTrainingCoref
  ' Goal:   Append feature vector + strState resolution to the trainings data
  '           which serves to identify what kind of non-coreferring state a constituent has
  ' History:
  ' 01-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function AppendTrainingCoref(ByRef arFeatVect() As String, ByRef arFeatVect2() As String, _
             ByRef arFeatVect3() As String, ByVal strState As String) As Boolean
    Try
      ' Append the information
      colCurrentTrainingCoref.Add(strCurrentPeriod & vbTab & Join(arFeatVect, vbTab) & vbTab & Join(arFeatVect2, vbTab) & vbTab & _
                                  Join(arFeatVect3, vbTab) & vbTab & strState)
      ' colCurrentTrainingCoref.Add(Join(arFeatVect2, vbTab) & vbTab & strState)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AppendTrainingCoref error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SaveTrainingCoref
  ' Goal:   Save -STATE and -COREF training files
  ' History:
  ' 01-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function SaveTrainingCoref(ByVal strType As String) As Boolean
    Try
      ' Action depends on [strType]
      Select Case strType
        Case "reflearn"
          Status("Saving <reflearn> training data ...")
          IO.File.WriteAllText(strCurrentTrainingRefLearn & ".train", colCurrentRefLearnTrain.Text)
          Status("Saving <reflearn> TEST data ...")
          IO.File.WriteAllText(strCurrentTrainingRefLearn & ".test", colCurrentRefLearnTest.Text)
        Case "state"
          Status("Saving <state> training data ...")
          IO.File.WriteAllText(strCurrentTrainingState, colCurrentTrainingState.Text)
        Case "coref"
          Status("Saving <coref> training data ...")
          IO.File.WriteAllText(strCurrentTrainingCoref, colCurrentTrainingCoref.Text)
      End Select
      Status("Done")
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AppendTrainingCoref error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   BestAnt
  ' Goal:   Find the best antecedent for a coreference relation, that:
  '         - starts from [dtrSrc]
  '         - goes to one element in [tblAnt] or in [tblSrc]
  '         Return this best match in [dtrAnt]
  '         Return an integer number indicating the likelyhood that this is a good match
  '           (numbers below a threshold need to be manually evaluated)
  ' History:
  ' 03-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function BestAnt(ByRef dtrSrc As DataRow, ByRef tblAnt As DataTable, _
                           ByRef dtrAnt As DataRow, ByRef strReason As String, ByRef intBest As Integer) As String
    Dim intI As Integer       ' Counter
    Dim intJ As Integer       ' Counter
    'Dim intBest As Integer   ' Index of the best match
    Dim intEval As Integer    ' The evaluation number of the best one
    Dim intNum As Integer     ' Number of best matches
    Dim intValue As Integer   ' The value given to one particular constraint
    Dim intMin As Integer     ' The evaluation value of the best antecedent right now
    Dim intMaxId As Integer   ' Id of the nearest by antecedent
    Dim strAction As String   ' The kidn of action we need to take
    Dim strCns As String      ' Name of current constraint
    Dim dtrCns() As DataRow   ' Ordered set of constraints

    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (tblAnt Is Nothing) Then Return -1
      ' ============= PROFILING =============
      ProfileStart("BestAnt", 3)
      ' =====================================
      ' Are there any antecedents? If there is NO antecedent, then return -4
      If (tblAnt.Rows.Count = 0) Then Return "NoAntecedent"
      ' Reset the evaluation of the antecedents
      For intI = 0 To tblAnt.Rows.Count - 1
        tblAnt.Rows(intI).Item("Eval") = 0
        tblAnt.Rows(intI).Item("Constraints") = ""
      Next intI
      ' Initialise the reason
      strReason = ""
      ' ============= PROFILING =============
      ProfileStart("BestAnt/Selecting", 4)
      ' =====================================
      ' Initialise the constraints
      dtrCns = tblConstraint.Select("", "Level ASC, Name ASC")
      ' ============= PROFILING =============
      ProfileEnd("BestAnt/Selecting", 4)
      ' =====================================
      ' Initialise the multiplication factor
      loc_intMult = 1
      ' ============== DEBUGGING =================
      'If (dtrSrc.Item("Id") = 497) Then Stop
      If (bDebugging) Then Debug.Print("BestAnt of " & dtrSrc.Item("Id") & " = " & dtrSrc.Item("Vern"))
      ' ==========================================
      ' Initialise the minimum
      intMin = -1
      '' Initialise the maximum number of antecedents
      'intAntMax = Math.Max(tblAnt.Rows.Count - loc_intAntSize, 0)
      ' Go through all the constraints until we have one best candidate left
      For intI = 0 To dtrCns.Count - 1
        ' Keep track of interrupt
        Application.DoEvents()
        If (bInterrupt) Then Return "Error"
        ' Get the name of the current constraint
        strCns = dtrCns(intI).Item("Name")
        ' Keep track of the last constraint that was used
        loc_strLastCns = strCns
        ' Get the multiplication factor of this constraint
        loc_intMult = dtrCns(intI).Item("Mult")
        ' ============= PROFILING =============
        ProfileStart("BestAnt/AllConstraints", 4)
        ' =====================================
        ' Evaluate all antecedents
        For intJ = 0 To tblAnt.Rows.Count - 1
          ' ============= PROFILING =============
          ProfileStart("Evaluate [" & strCns & "]", 5)
          ' =====================================
          ' Add the evaluation for this particular constraint
          If (EvaluateConstraint(strCns, dtrSrc, tblAnt.Rows(intJ), intValue, intMin)) Then
            ' ============= PROFILING =============
            ProfileEnd("Evaluate [" & strCns & "]", 5)
            ' =====================================
            ' Add the constraint's name and value...
            AddSemiStack(tblAnt.Rows(intJ).Item("Constraints"), strCns & "[" & intValue & "]")
            '' =========== DEBUGGING ===========
            'Debug.Print("Constraint=" & strCns & " Antecedent " & intJ + 1 & _
            '            " Eval=" & tblAnt.Rows(intJ).Item("Eval") & _
            '            " =[" & _
            '            NodeInfo(IdToNode(tblAnt.Rows(intJ).Item("Id"))) & "]")
            '' =================================
          Else
            ' Something went wrong
            AutoReport("Could not evaluate constraint [" & strCns & "] for (" & dtrSrc.Item("Id") & _
                       "," & tblAnt.Rows(intJ).Item("Id") & ")")
            ' Return failure
            Return "Error"
          End If
        Next intJ
        ' ============= PROFILING =============
        ProfileEnd("BestAnt/AllConstraints", 4)
        ' =====================================
        ' Are we in fully automatic?
        If (bIsFullyAuto) Then
          ' ============= PROFILING =============
          ProfileStart("BestAnt/MinimumEval", 4)
          ' =====================================
          ' Calculate the minimal value
          If (Not MinimumEvaluation(intMin)) Then Return "Error"
          ' ============= PROFILING =============
          ProfileEnd("BestAnt/MinimumEval", 4)
          ' =====================================
        End If
        ' I used to have this here, but the results were not in order. Multiplication factor is 
        '   now determined BEFORE actual application of the constraint
        '' Get the multiplication factor of this constraint
        'loc_intMult = dtrCns(intI).Item("Mult")
        ' See if ONE winner is already left over
        If (NumBest(tblAnt, intBest, intEval) = 1) Then
          ' One is left -- determine the datarow
          dtrAnt = tblAnt(intBest)
          ' ================= DEBUG ==================
          ' Check an alternative antecedent (a better one) for CmKentse:497
          'For intJ = 0 To tblAnt.Rows.Count - 1
          '  If (tblAnt.Rows(intJ).Item("Id") = 471) Then
          '    ' This is the better one
          '    Debug.Print("Constraints of winner (" & dtrAnt.Item("Id") & "): " & dtrAnt.Item("Constraints"))
          '    Debug.Print("Constraints of loser (471): " & tblAnt.Rows(intJ).Item("Constraints"))
          '  End If
          'Next intJ
          ' ==========================================
          ' We found one "winner" 
          strAction = ""
          ' First check if this is a suspicious combination for some reason
          If (IsSuspicious(dtrSrc, dtrAnt, strReason, strAction)) Then
            ' What action are we to take?
            Select Case strAction
              Case "MakeNew"
                ' Indicate that a new relation should be made
                ' ============= PROFILING =============
                ProfileEnd("BestAnt", 3)
                ' =====================================
                Return "BadBestMatch"
              Case Else
                ' We should ask the user
                strReason = "Suspicious combination: " & strReason
                ' The link is regarded as suspicious, so should be checked by hand
                ' ============= DEBUGGING ================
                If (bDebugging) Then
                  ' Give a fuller account
                  Debug.Print(strReason & vbCrLf & _
                              "The proposed link is from [" & _
                              dtrSrc.Item("Vern") & "] > [" & dtrAnt.Item("Vern") & "]")
                End If
                ' =========================================
                ' ============= PROFILING =============
                ProfileEnd("BestAnt", 3)
                ' =====================================
                Return "Suspicious"
            End Select
            ' Check whether there is ambiguity...
          ElseIf (CnsAmbiguity(dtrSrc, dtrAnt) > 1) Then
            ' There is too much ambiguity: let user decide
            strReason = "There is 1 best antecedent, but there is more than 1 full NP to be chosen from."
            ' ============= DEBUGGING ================
            ' Show what is the source and the destination
            If (bDebugging) Then
              ' Give a fuller account
              Debug.Print(strReason & vbCrLf & _
                          "The proposed link is from [" & _
                          dtrSrc.Item("Vern") & "] > [" & dtrAnt.Item("Vern") & "]")
            End If
            ' =========================================
            ' ============= PROFILING =============
            ProfileEnd("BestAnt", 3)
            ' =====================================
            Return "Suspicious"
          Else
            ' No, there is no ambiguity, so let's make the link
            ' ============= Print this evaluation (if we are debugging)
            If (bDebugging) Then
              ' Print the evaluation of the last 10 antecedents
              intJ = tblAnt.Rows.Count - 1
              While (intJ >= 0) And (intJ > tblAnt.Rows.Count - 10)
                Debug.Print("Candidate [" & tblAnt.Rows(intJ).Item("Vern") & "] = " & _
                    tblAnt(intJ).Item("Eval"))
                ' Go to the next candidate
                intJ -= 1
              End While
            End If
            ' =====================================
            ' ============= PROFILING =============
            ProfileEnd("BestAnt", 3)
            ' =====================================
            Return "NoSuspicion"
          End If
        End If
      Next intI
      ' ============== DEBUG =================
      If (bDebugging) Then
        intNum = NumBest(tblAnt, intBest, intEval)
        Debug.Print("Number of matches was: " & intNum)
      End If
      ' ============= PROFILING =============
      ProfileEnd("BestAnt", 3)
      ' =====================================
      ' ======================================
      ' We have not been able to get ONE best match, so return THE ONE NEAREST BY
      ' (1) Get the ID number of the [intBest]
      intMaxId = tblAnt(intBest).Item("Id")
      For intI = 0 To tblAnt.Rows.Count - 1
        ' Check if this has the same evaluation
        With tblAnt(intI)
          If (.Item("Eval") = intEval) Then
            ' Check if it has a higher antecedent number
            If (.Item("Id") > intMaxId) Then
              ' Adapt the index of the one nearest by
              intBest = intI
              ' Adapt the highest ID
              intMaxId = .Item("Id")
            End If
          End If
        End With
      Next intI
      ' But do put one of the best matches where it should go
      dtrAnt = tblAnt(intBest)
      Return "Ambiguous"
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/BestAnt error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MinimumEvaluation
  ' Goal:   Determine what the minimum "Eval" value is of all candidates in the antecedent's stack
  ' Return: The actual value
  ' History:
  ' 13-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function MinimumEvaluation(ByRef intEval As Integer) As Boolean
    Dim intI As Integer     ' Counter
    Dim intThis As Integer  ' Value of current one

    Try
      ' Validate
      If (tblAntStack Is Nothing) Then Return False
      If (tblAntStack.Rows.Count = 0) Then Return False
      ' Get start value of the evaluation
      intEval = tblAntStack.Rows(0).Item("Eval")
      For intI = 1 To tblAntStack.Rows.Count - 1
        ' Get this value
        intThis = tblAntStack.Rows(intI).Item("Eval")
        ' Is this one lower?
        If (intThis < intEval) Then intEval = intThis
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/MinimumEvaluation error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ConstraintValue
  ' Goal:   Evaluate one constraint between source and destination
  ' Return: The actual value is also returned in [intValue]
  ' History:
  ' 24-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function ConstraintValue(ByVal strConstraint As String, ByRef dtrSrc As DataRow, ByRef dtrDst As DataRow, _
                                    ByRef intValue As Integer) As Boolean
    Try
      ' Validate
      If (strConstraint = "") OrElse (dtrSrc Is Nothing) OrElse (dtrDst Is Nothing) Then Return False
      ' Action depends on which constraint this is
      Select Case strConstraint
        Case "Agree"
          ' One violation when PGN does not agree
          intValue = CnsAgree(dtrSrc, dtrDst)
        Case "AgrPerson"
          ' One violation when PGN does not agree
          intValue = CnsAgrPerson(dtrSrc, dtrDst)
        Case "AgrClause"
          ' Restrictions on a clausal antecedent
          intValue = CnsAgrClause(dtrSrc, dtrDst)
        Case "NoClause"
          ' One violation when the antecedent is a clause
          intValue = CnsNoClause(dtrSrc, dtrDst)
        Case "AgrGenderNumber"
          ' One violation when PGN does not agree
          intValue = CnsAgrGenderNumber(dtrSrc, dtrDst)
        Case "NoCrossAgrPerson"
          ' One violation when there is agreement in person at a cross speech boundary
          intValue = CnsNoCrossAgrPerson(dtrSrc, dtrDst)
        Case "NoCrossEqSubject"
          ' One violation when there is agreement in person at a cross speech boundary
          intValue = CnsNoCrossEqSubject(dtrSrc, dtrDst)
        Case "Disjoint"
          ' One violation when src+target are in the same IP MAT/SUB/SMC
          intValue = CnsDisjoint(dtrSrc, dtrDst)
        Case "EqualHead"
          ' One violation when the src head noun does not agree with any of the head nouns
          '   in the chain of the target
          intValue = CnsEqHead(dtrSrc, dtrDst)
        Case "AgrGenderSrc"
          ' One violation when the src head noun does not agree with any of the head nouns
          '   in the chain of the target
          intValue = CnsAgrGenderSrc(dtrSrc, dtrDst)
        Case "IPdist"
          ' One violation for every IP between Src and Target
          intValue = CnsIPdist(dtrSrc, dtrDst)
        Case "Ambiguity"
          ' One violation for every full NP in the IP of the target
          intValue = CnsAmbiguity(dtrSrc, dtrDst)
        Case "NPtypeRel"
          ' The likelihood that the source NP's type matches that of the target
          intValue = CnsNPtypeRel(dtrSrc, dtrDst)
        Case "GrRoleRel"
          ' The likelihood that the source NP's GrRole matches that of target
          intValue = CnsGrRoleRel(dtrSrc, dtrDst)
        Case "NPtypeDst"
          ' The likelihood that the source NP's type matches that of the target
          intValue = CnsNPtypeDst(dtrSrc, dtrDst)
        Case "GrRoleDst"
          ' The likelihood that the source NP's GrRole matches that of target
          intValue = CnsGrRoleDst(dtrSrc, dtrDst)
        Case "NoCataphore"
          ' The likelihood that the source NP's GrRole matches that of target
          intValue = CnsNoCataphore(dtrSrc, dtrDst)
        Case "NearDem"
          ' Special requirements of the near demonstrative
          intValue = CnsNearDem(dtrSrc, dtrDst)
        Case "ProTop"
        Case "FamDef"
        Case "Cohere"
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/ConstraintValue error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   EvaluateConstraint
  ' Goal:   Evaluate one constraint between source and destination
  ' Return: The actual value is also returned in [intValue]
  ' History:
  ' 24-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function EvaluateConstraint(ByVal strConstraint As String, ByRef dtrSrc As DataRow, _
      ByRef dtrDst As DataRow, ByRef intValue As Integer, ByRef intMin As Integer) As Boolean
    Try
      ' Should we process this constraint for this antecedent?
      If (bIsFullyAuto) AndAlso (intMin >= 0) AndAlso (dtrDst.Item("Eval") > intMin) Then
        ' We are in fully automatic mode, a minimum has been determined, and we are less
        '   harmonic than this minimum: don't evaluate!
        dtrDst.Item("Eval") *= loc_intMult
        intValue = 0
        Return True
      End If
      ' Multiply constraint with the correct factor
      dtrDst.Item("Eval") *= loc_intMult
      ' Get the constraint's value
      If (Not ConstraintValue(strConstraint, dtrSrc, dtrDst, intValue)) Then Return False
      '' Action depends on which constraint this is
      'Select Case strConstraint
      '  Case "Agree"
      '    ' One violation when PGN does not agree
      '    intValue = CnsAgree(dtrSrc, dtrDst)
      '  Case "AgrPerson"
      '    ' One violation when PGN does not agree
      '    intValue = CnsAgrPerson(dtrSrc, dtrDst)
      '  Case "AgrClause"
      '    ' Restrictions on a clausal antecedent
      '    intValue = CnsAgrClause(dtrSrc, dtrDst)
      '  Case "NoClause"
      '    ' One violation when the antecedent is a clause
      '    intValue = CnsNoClause(dtrSrc, dtrDst)
      '  Case "AgrGenderNumber"
      '    ' One violation when PGN does not agree
      '    intValue = CnsAgrGenderNumber(dtrSrc, dtrDst)
      '  Case "NoCrossAgrPerson"
      '    ' One violation when there is agreement in person at a cross speech boundary
      '    intValue = CnsNoCrossAgrPerson(dtrSrc, dtrDst)
      '  Case "NoCrossEqSubject"
      '    ' One violation when there is agreement in person at a cross speech boundary
      '    intValue = CnsNoCrossEqSubject(dtrSrc, dtrDst)
      '  Case "Disjoint"
      '    ' One violation when src+target are in the same IP MAT/SUB/SMC
      '    intValue = CnsDisjoint(dtrSrc, dtrDst)
      '  Case "EqualHead"
      '    ' One violation when the src head noun does not agree with any of the head nouns
      '    '   in the chain of the target
      '    intValue = CnsEqHead(dtrSrc, dtrDst)
      '  Case "AgrGenderSrc"
      '    ' One violation when the src head noun does not agree with any of the head nouns
      '    '   in the chain of the target
      '    intValue = CnsAgrGenderSrc(dtrSrc, dtrDst)
      '  Case "IPdist"
      '    ' One violation for every IP between Src and Target
      '    intValue = CnsIPdist(dtrSrc, dtrDst)
      '  Case "Ambiguity"
      '    ' One violation for every full NP in the IP of the target
      '    intValue = CnsAmbiguity(dtrSrc, dtrDst)
      '  Case "NPtypeRel"
      '    ' The likelihood that the source NP's type matches that of the target
      '    intValue = CnsNPtypeRel(dtrSrc, dtrDst)
      '  Case "GrRoleRel"
      '    ' The likelihood that the source NP's GrRole matches that of target
      '    intValue = CnsGrRoleRel(dtrSrc, dtrDst)
      '  Case "NPtypeDst"
      '    ' The likelihood that the source NP's type matches that of the target
      '    intValue = CnsNPtypeDst(dtrSrc, dtrDst)
      '  Case "GrRoleDst"
      '    ' The likelihood that the source NP's GrRole matches that of target
      '    intValue = CnsGrRoleDst(dtrSrc, dtrDst)
      '  Case "NoCataphore"
      '    ' The likelihood that the source NP's GrRole matches that of target
      '    intValue = CnsNoCataphore(dtrSrc, dtrDst)
      '  Case "NearDem"
      '    ' Special requirements of the near demonstrative
      '    intValue = CnsNearDem(dtrSrc, dtrDst)
      '  Case "ProTop"
      '  Case "FamDef"
      '  Case "Cohere"
      'End Select
      ' Check for errors
      If (intValue < 0) Then
        ' There is an error, so return failure
        Return False
      Else
        ' Nothing wrong, so adapt the evaluation with the new value
        dtrDst.Item("Eval") += intValue
        ' Return success
        Return True
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/EvaluateConstraint error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CorefProgress
  ' Goal:   Set the text of the [tsProgress] element
  ' History:
  ' 09-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub CorefProgress(ByVal strText As String)
    Try
      ' Find the correct component and set the text
      frmMain.tsProgress.Text = strText
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CorefProgress error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   GetAntMax
  ' Goal:   Get the maximum number of antecedents allowed
  ' History:
  ' 30-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetAntMax() As Integer
    ' This number is just stored locally here in [modAuto]
    Return loc_intAntSize
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsCrossSpeech
  ' Goal:   Check whether there is a crossspeech boundary between source and antecedent
  ' History:
  ' 11-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function IsCrossSpeech(ByRef dtrSrc As DataRow, ByRef dtrDst As DataRow) As Boolean
    Dim bIPsrc As Boolean         ' Whether the source constituent is a speech IP
    Dim bIPant As Boolean         ' Whether the antecedent is a speech IP

    Try
      ' Check if there is a cross speech boundary
      bIPsrc = (dtrSrc.Item("IsSpeech") Like "True")
      bIPant = (dtrDst.Item("IsSpeech") Like "True")
      ' One of them needs to have SPE, and the other not
      Return ((bIPant And Not bIPsrc) OrElse (Not bIPant And bIPsrc))
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/IsCrossSpeech error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsProNoHuman
  ' Goal:   Check whether a pronominal source links to a nominal antecedent
  '           which is not human
  ' History:
  ' 04-10-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function IsProNoHuman(ByRef dtrSrc As DataRow, ByRef dtrDst As DataRow) As Boolean
    Dim strHead As String   ' The head of the antecedent

    Try
      ' Check if the source is a PRO
      If (dtrSrc.Item("NPtype") = "Pro") Then
        ' Now get the head of the antecedent
        strHead = LCase(dtrDst.Item("Head"))
        ' Do we have something?
        If (strHead <> "") Then
          ' Check for known human antecedents
          Return (Not DoLike(strHead, strHumanAnt))
        End If
      End If
      ' Return false by default
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/ProNoHuman error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsSuspicious
  ' Goal:   Check if the proposed link is suspicious and in need of manual verification
  '         Recognize several situations that are suspicious, as e.g:
  '         (1) The source is a bare noun
  '         (2) If either source or destination has a head noun (or chain), they don't match
  '         (x) etcetera!
  ' History:
  ' 11-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function IsSuspicious(ByRef dtrSrc As DataRow, ByRef dtrDst As DataRow, ByRef strReason As String, _
                                ByRef strAction As String) As Boolean
    Dim intSrcType As Integer ' Source NP's evaluation type

    Try
      ' ============= PROFILING =============
      ProfileStart("IsSuspicious]", 1)
      ' =====================================
      '' ================ DEBUGGING ==========
      'If (dtrSrc.Item("Id") = 2209) Then Stop
      '' =====================================
      ' Normally the action will be "AskUser"
      strAction = "AskUser"
      ' Non-agreement is suspicious (actually perhaps agreement should be obligatory)
      If (CnsAgrGenderNumber(dtrSrc, dtrDst) <> 0) Then
        ' Give the reason for being suspicious
        strReason = "There is no agreement in Gender or Number." & vbCrLf & _
          "Source=" & dtrSrc.Item("PGN") & " Antecedent=" & dtrDst.Item("PGN") & "."
        ' This must be a new element, because there is no agreement
        strAction = "MakeNew"
        ' ============= PROFILING =============
        ProfileEnd("IsSuspicious]", 1)
        ' =====================================
        ' Return showing that we are suspicious
        Return True
      End If
      ' Is this cross speech or not?
      If (IsCrossSpeech(dtrSrc, dtrDst)) Then
        ' There must not be person agreement
        If (CnsAgrPerson(dtrSrc, dtrDst) = 0) Then
          ' Give the reason for being suspicious
          strReason = "There is agreement in Person, which is not allowed at cross speech." & vbCrLf & _
            "Source=" & dtrSrc.Item("PGN") & " Antecedent=" & dtrDst.Item("PGN") & "."
          ' ============= PROFILING =============
          ProfileEnd("IsSuspicious]", 1)
          ' =====================================
          ' Return showing that we are suspicious
          Return True
        End If
        ' There is a subject restriction
        If (CnsNoCrossEqSubject(dtrSrc, dtrDst) > 0) Then
          ' Give the reason for being suspicious
          strReason = "There are restrictions on a direct speech NP pointing back to a subject." & vbCrLf & _
            "Source=" & dtrSrc.Item("GrRole") & " Antecedent=" & dtrDst.Item("GrRole") & "."
          ' ============= PROFILING =============
          ProfileEnd("IsSuspicious]", 1)
          ' =====================================
          ' Return showing that we are suspicious
          Return True
        End If
        ' There is some person restriction: 3 -> 1 should be questioned...
        If (AgrPerson(dtrSrc, dtrDst, 3, 1)) Then
          ' A third person points to a first person: check with user
          strReason = "A 3rd person points to a 1st person across direct speech." & vbCrLf & _
           "Source=" & dtrSrc.Item("GrRole") & " Antecedent=" & dtrDst.Item("GrRole") & "."
          ' ============= PROFILING =============
          ProfileEnd("IsSuspicious]", 1)
          ' =====================================
          ' Return showing that we are suspicious
          Return True
        End If
      Else
        ' There should be person agreement
        If (CnsAgrPerson(dtrSrc, dtrDst) <> 0) Then
          ' Give the reason for being suspicious
          strReason = "There is no agreement in Person, where it should be." & vbCrLf & _
            "Source=" & dtrSrc.Item("PGN") & " Antecedent=" & dtrDst.Item("PGN") & "."
          ' ============= PROFILING =============
          ProfileEnd("IsSuspicious]", 1)
          ' =====================================
          ' Return showing that we are suspicious
          Return True
        End If
      End If
      ' Check the source NP type
      If (CnsEqHead(dtrSrc, dtrDst) <> 0) Then
        ' If we are dealing with a definite/demonstrative NP, then it can still be correct
        If (DoLike(dtrSrc.Item("NPtype"), "DemNP|DefNP")) Then
          ' Give the reason for being suspicious
          strReason = "Definite source NP's head noun does not match antecedent's head(or chain)." & vbCrLf & _
            "Source=" & dtrSrc.Item("Head") & " Antecedent=" & dtrDst.Item("Head") & "(" & dtrDst.Item("Equal") & ")."
          ' Perhaps we'd better ask the user (default action)
        ElseIf (dtrSrc.Item("NPtype") = "Proper") AndAlso (dtrDst.Item("NPtype") = "Proper") Then
          ' Give the reason for being suspicious
          strReason = "Source's proper name does not match antecedent's head(or chain)." & vbCrLf & _
            "Source=" & dtrSrc.Item("Head") & " Antecedent=" & dtrDst.Item("Head") & "(" & dtrDst.Item("Equal") & ")."
          ' Perhaps we'd better ask the user (default action)
        ElseIf (dtrSrc.Item("NPtype") = "AnchoredNP") Then
          ' Give the reason for being suspicious
          strReason = "Source NP is anchored, and best antecedent does not match." & vbCrLf & _
            "Source=" & dtrSrc.Item("Head") & " Antecedent=" & dtrDst.Item("Head") & "(" & dtrDst.Item("Equal") & ")."
          ' This must be a new element, because there is no agreement
          strAction = "MakeNew"
        Else
          ' Give the reason for being suspicious
          strReason = "Mismatch between source NP's head noun and antecedent chain ." & vbCrLf & _
            "Source=" & dtrSrc.Item("Head") & " Antecedent=" & dtrDst.Item("Head") & "(" & dtrDst.Item("Equal") & ")."
          ' This must be a new element, because there is no agreement
          strAction = "MakeNew"
        End If
        ' ============= PROFILING =============
        ProfileEnd("IsSuspicious]", 1)
        ' =====================================
        ' Return showing that we are suspicious
        Return True
      End If
      ' Check whether there is agreement in gender between source and antecedent
      If (CnsAgrGenderSrc(dtrSrc, dtrDst) <> 0) Then
        ' Give the reason for being suspicious
        strReason = "The gender of antecedent or source is less specific than the other." & vbCrLf & _
          "Source=" & dtrSrc.Item("PGN") & " Antecedent=" & dtrDst.Item("PGN") & "."
        ' ============= PROFILING =============
        ProfileEnd("IsSuspicious]", 1)
        ' =====================================
        ' Return showing that we are suspicious
        Return True
      End If
      ' Check whether the source is likely enough to be a coreference source
      intSrcType = CnsNPtypeSrc(dtrSrc, dtrDst)
      ' Action depends on the numerical value
      Select Case intSrcType
        Case Is >= 6
          ' It seems this situation NEVER arises! (filtered out by NeedsResolving)
          ' Give the reason for being suspicious
          strReason = "The source NP cannot be a coreference source." & vbCrLf & _
            "Source=" & dtrSrc.Item("Head") & " Antecedent=" & dtrDst.Item("Head") & "(" & dtrDst.Item("Equal") & ")."
          ' We should ask the user to decide!
          strAction = "MakeNew"
          ' ============= PROFILING =============
          ProfileEnd("IsSuspicious]", 1)
          ' =====================================
          ' Return showing that we are suspicious
          Return True
        Case 5
          ' Give the reason for being suspicious
          strReason = "The source NP is not a very likely coreference source." & vbCrLf & _
            "Source=" & dtrSrc.Item("Head") & " Antecedent=" & dtrDst.Item("Head") & "(" & dtrDst.Item("Equal") & ")."
          ' ============= PROFILING =============
          ProfileEnd("IsSuspicious]", 1)
          ' =====================================
          ' Return showing that we are suspicious
          Return True
      End Select
      ' Disjoint members are suspicious too, I suppose (perhaps under certain circumstances?)
      If (CnsDisjoint(dtrSrc, dtrDst)) Then
        ' Give the reason for being suspicious
        strReason = "Source and antecedent are in the same syntactic domain."
        ' ============= PROFILING =============
        ProfileEnd("IsSuspicious]", 1)
        ' =====================================
        ' Return showing that we are suspicious
        Return True
      End If
      ' Check whether the IP distance is not too large
      If (dtrDst.Item("IPdist") > loc_intMaxIPdist) Then
        ' Give the reason for being suspicious
        strReason = "The antecedent is further from the source than defined in Tools/Settings (" & _
          loc_intMaxIPdist & ")." & vbCrLf & _
          " IP distance=" & dtrDst.Item("IPdist") & "."
        ' ============= PROFILING =============
        ProfileEnd("IsSuspicious]", 1)
        ' =====================================
        ' Return showing that we are suspicious
        Return True
      End If
      ' Check whether there are not two candidates only differing in IP distance 
      '   between 1-3. E.g. one candidate has IP distance "1", the other "2" or "3"
      If (NumCloseCandidates(dtrDst)) > 1 Then
        ' Give the reason for being suspicious
        strReason = "There is more than 1 very likely candidate within close vicinity."
        ' ============= PROFILING =============
        ProfileEnd("IsSuspicious]", 1)
        ' =====================================
        ' Return showing that we are suspicious
        Return True
      End If
      ' Check whether the nominal antecedent of a pronoun is human
      If (IsProNoHuman(dtrSrc, dtrDst)) Then
        ' Give the reason for being suspicious
        strReason = "A pronoun refers back to a possibly non-human antecedent."
        ' ============= PROFILING =============
        ProfileEnd("IsSuspicious]", 1)
        ' =====================================
        ' Return showing that we are suspicious
        Return True
      End If
      ' ============= PROFILING =============
      ProfileEnd("IsSuspicious]", 1)
      ' =====================================
      ' There is no suspicion
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/IsSuspicious error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Make sure suspicion is aroused
      Return True
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   NumCloseCandidates
  ' Goal:   Show the constraints for this item in the antecedent's stack
  ' History:
  ' 22-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function NumCloseCandidates(ByRef dtrThis As DataRow) As Integer
    Dim intNum As Integer = 0 ' Number of close enough candidates
    Dim arDst() As String     ' Constraint values of the present destination
    Dim arAlt() As String     ' Constraint values of the alternative antecedent
    Dim intI As Integer       ' Counter
    Dim intJ As Integer       ' Second counter
    Dim intMax As Integer     ' Number of constraints to compare
    Dim bEqual As Boolean     ' Whether two constraint evaluations are equal

    Try
      ' Validate
      If (dtrThis Is Nothing) Then Return 0
      ' Get the constraint values of myself
      arDst = Split(dtrThis.Item("Constraints"), ";")
      ' Visit all the antecedents with IP distance <=3
      For intI = 0 To tblAntStack.Rows.Count - 1
        ' Access this one
        With tblAntStack.Rows(intI)
          ' Is it close enough?
          If (.Item("IPdist") <= 4) Then
            ' Get its constraints
            arAlt = Split(.Item("Constraints"), ";")
            ' Compare the constraints (except for the IPdist)
            intMax = Math.Max(UBound(arDst), UBound(arAlt))
            ' Assume equality
            bEqual = True
            For intJ = 0 To intMax
              ' Compare this constraint
              If (arDst(intJ) <> arAlt(intJ)) Then
                ' Double check wither this is not IPdist
                If (Not arDst(intJ) Like "IPdist*") Then
                  ' We really found a difference!!
                  bEqual = False
                  Exit For
                End If
              End If
            Next intJ
            ' Adapt the total count of close candidates
            If (bEqual) Then intNum += 1
          End If
        End With
      Next intI
      ' REturn the total number of candidates
      Return intNum
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/NumCloseCandidates error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Make sure suspicion is aroused
      Return 0
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ShowConstraints
  ' Goal:   Show the constraints for this item in the antecedent's stack
  ' History:
  ' 25-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function ShowConstraints(ByVal intId As Integer) As String
    Dim intI As Integer           ' Index inthe antecedent's stack
    Dim strConstraints As String  ' The constraints

    Try
      ' Validate
      If (intId < 0) Then Return ""
      ' Get index of antecedent's stack
      intI = GetAntStack(IdToNode(intId))
      ' Validate
      If (intI < 0) Then Return ""
      ' We have something valid then...
      strConstraints = tblAntStack(intI).Item("Constraints")
      ' Anything?
      If (Trim(strConstraints) <> "") Then
        ' Return the result
        Return strConstraints
      End If
      ' Failure
      Return ""
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/ShowConstraints error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SplitAgree
  ' Goal:   Split agreement into person, gender and number for separate evaluation
  ' History:
  ' 24-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function SplitAgree(ByVal strPGN As String, ByRef strPerson As String, ByRef strGender As String, _
                               ByRef strNumber As String) As Boolean
    ' Dim arPgn() As String ' Array of different PGNs

    Try
      ' Validate
      If (strPGN = "") Then Return False
      ' Set default values
      strPerson = "empty" : strGender = "empty" : strNumber = "empty"
      ' If the length is exactly 3, there is person, gender and number
      If (strPGN.Length = 3) Then
        strPerson = Left(strPGN, 1)
        strGender = Mid(strPGN, 2, 1)
        strNumber = Right(strPGN, 1)
      Else
        ' Try special cases
        Select Case strPGN
          Case "unknown"
            ' Copy the unknown part...
            strPerson = "unknown" : strGender = "unknown" : strNumber = "unknown"
          Case "empty"
            ' All default values apply
          Case "3", "2", "1"
            ' Person and Gender remain "empty" - to be filled in
            strPerson = strPGN
          Case "s", "p"
            ' Copy the number, but keep the rest empty
            strNumber = strPGN
          Case "f", "m", "n"
            ' Copy the gender, but keep the rest empty
            strGender = strPGN
          Case "3m", "3f", "3n", "2m", "2f", "2n", "1m", "1f", "1n"
            ' Copy person and gender
            strPerson = Left(strPGN, 1)
            strGender = Right(strPGN, 1)
          Case "3s", "3p", "2s", "2p", "1s", "1p"
            ' Copy person and number
            strPerson = Left(strPGN, 1)
            strNumber = Right(strPGN, 1)
          Case "3ms;3fs"
            strPerson = "3" : strNumber = "s" : strGender = "unknown"
          Case Else
            ' Warn user
            MsgBox("modAuto/SplitAgree: don't know what to do with [" & strPGN & "]")
            Return False
        End Select
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/SplitAgree error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetCommonPGN
  ' Goal:   If the PGN is not certain, we should ask the user, and if he is
  '           not able to supply it, we return FALSE
  ' History:
  ' 14-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetCommonPGN(ByVal strPGN As String) As String
    Dim arPgn() As String   ' The different PGNs
    Dim intI As Integer     ' Counter
    Dim strPerson As String ' The person
    Dim strGender As String ' The gender
    Dim strNumber As String ' The number
    Dim strThis(0 To 2) As String ' PGN divided

    Try
      ' Validate
      If (strPGN = "") Then Return "unknown"
      ' Split the array
      arPgn = Split(strPGN, ";")
      ' ====== DEBUG =======
      'If (strPGN = "2p;2s") OrElse (strPGN = "2s;2p") Then Stop
      ' ====================
      ' Derive the pgn of the first element
      If (Not SplitAgree(arPgn(0), strThis(0), strThis(1), strThis(2))) Then
        ' Something went wrong
        Return "unknown"
      End If
      ' Assign values to pgn
      strPerson = strThis(0) : strGender = strThis(1) : strNumber = strThis(2)
      ' Check the other values
      For intI = 1 To UBound(arPgn)
        ' Get the PGN of this element
        If (Not SplitAgree(arPgn(intI), strThis(0), strThis(1), strThis(2))) Then
          ' Something went wrong
          Return "unknown"
        End If
        ' Compare the PGN's and generalize
        If (strPerson <> strThis(0)) Then strPerson = ""
        If (strGender <> strThis(1)) Then strGender = ""
        If (strNumber <> strThis(2)) Then strNumber = ""
      Next intI
      ' Don't allow "empty" here
      If (strPerson = "empty") Then strPerson = ""
      If (strGender = "empty") Then strGender = ""
      If (strNumber = "empty") Then strNumber = ""
      ' Combine the three
      strPGN = strPerson & strGender & strNumber
      If (strPGN = "") Then strPGN = "unknown"
      Return strPGN
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetCommonPGN error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Failure
      Return "unknown"
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MakeCertainPgn
  ' Goal:   If the PGN is not certain, we should ask the user, and if he is
  '           not able to supply it, we return FALSE
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function MakeCertainPgn(ByRef ndxThis As XmlNode) As Boolean
    Dim strPgn As String = ""         ' PGN feature value

    Try
      ' Do we need to ask?
      strPgn = GetFeature(ndxThis, "NP", "PGN")
      If (InStr(strPgn, ";") > 0) Then
        ' Are we in fully auto mode?
        If (bIsFullyAuto) Then
          ' Just choose the first possibility
          strPgn = Split(strPgn, ";")(0)
          If (Not AddFeature(pdxCurrentFile, ndxThis, "NP", "PGN", strPgn)) Then Return False
        Else
          ' Get the node
          If (Not UserAdaptPGN(ndxThis, strPgn)) Then
            ' Some error occurred, so return failure
            Return False
          End If
        End If
      End If
      ' Normally we return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/MakeCertainPgn error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsCertainPgn
  ' Goal:   If the PGN is not certain, we should ask the user, and if he is
  '           not able to supply it, we return FALSE
  ' History:
  ' 12-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function IsCertainPgn(ByRef dtrThis As DataRow) As Boolean
    Dim strPgn As String = ""         ' PGN feature value
    Dim ndxThis As XmlNode = Nothing  ' Current node

    Try
      ' Do we need to ask?
      strPgn = dtrThis.Item("PGN")
      If (InStr(strPgn, ";") > 0) Then
        ' Are we in fully automatic mode?
        If (bIsFullyAuto) Then
          ' Get the greatest common divider in PGN
          dtrThis.Item("PGN") = GetCommonPGN(strPgn)
          '' ============= DEBUG ========
          'If (dtrThis.Item("PGN") = "2empty") Then Stop
          '' =============================
          ' Return okay
          Return True
        End If
        ' Get the node
        ndxThis = IdToNode(dtrThis.Item("Id"))
        If (UserAdaptPGN(ndxThis, strPgn)) Then
          ' Store this feature in the datarow
          dtrThis.Item("PGN") = strPgn
        End If
      End If
      ' Normally we return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/IsCertainPgn error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   CnsAgrPerson
  ' Goal:   Apply the constraint "AgreePerson"
  '         Is there possible agreement between the source and the destination?
  ' Return: Return 1 for violation, 0 for no violation, and -1 for error
  ' History:
  ' 24-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CnsAgrPerson(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    Dim strPGNsrc As String ' Person/Number/Gender of the source
    Dim strPGNdst As String ' Person/Number/Gender of the destination
    'Dim bIPsrc As Boolean         ' Whether the source constituent is a speech IP
    'Dim bIPant As Boolean         ' Whether the antecedent is a speech IP
    Dim strSrc(0 To 2) As String  ' Person, number and gender split up
    Dim strDst(0 To 2) As String  ' Person, number and gender split up

    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (dtrAnt Is Nothing) Then Return -1
      ' Check whether there is uncertainty in PGN
      If (Not IsCertainPgn(dtrSrc)) Then Return False
      If (Not IsCertainPgn(dtrAnt)) Then Return False
      ' Get source and destination agreement
      strPGNsrc = dtrSrc.Item("PGN")
      strPGNdst = dtrAnt.Item("PGN")
      ' ============ DEBUG ============
      ' If (strPGNsrc = "3ms" AndAlso strPGNdst = "3s") Then Stop
      ' If (dtrAnt.Item("Id") = 162) Then Stop
      ' ===============================
      'Debug.Print(GetFeature(IdToNode(dtrAnt.Item("Id")), "NP", "PGN"))
      ' Distentangle the source
      If (Not SplitAgree(strPGNsrc, strSrc(0), strSrc(1), strSrc(2))) Then
        ' Something went wrong
        Return False
      End If
      ' Distentangle the antecedent
      If (Not SplitAgree(strPGNdst, strDst(0), strDst(1), strDst(2))) Then
        ' Something went wrong
        Return False
      End If
      ' Compare source and antecedent PERSON
      If (strSrc(0) = strDst(0)) Then
        ' No violation whatsoever
        Return 0
      ElseIf (strSrc(0) = "unknown") OrElse (strDst(0) = "unknown") Then
        ' When either is unknown, then user-checking needs to take place, so return a violation
        Return 1
      ElseIf (strSrc(0) = "empty") Then
        ' An "empty" source NP can refer to ANY preceding NP as far as PGN is concerned
        Return 0
      ElseIf (strDst(0) = "empty") Then
        ' This must be severely punished: it should be impossible to point to an empty DESTINATION
        Return 1
      Else
        ' They are not equal, and there is no special case, so: violation
        Return 1
        '' Check if there is a cross speech boundary
        'bIPsrc = (dtrSrc.Item("IPlabel") Like "*SPE*")
        'bIPant = (dtrAnt.Item("IPlabel") Like "*SPE*")
        '' Check if they do NOT agree
        'If (bIPant Xor bIPsrc) Then
        '  ' There is a speech boundary transition, so give them the benefit of the doubt
        '  Return 0
        'Else
        'End If
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsAgrPerson error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AgrPerson
  ' Goal:   Check if the agreement from person [intSrc] to person [intDst] is there
  ' Return: Return True or False
  ' History:
  ' 15-03-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function AgrPerson(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow, ByVal intSrc As Integer, _
                             ByVal intDst As Integer) As Boolean
    Dim strPGNsrc As String ' Person/Number/Gender of the source
    Dim strPGNdst As String ' Person/Number/Gender of the destination
    Dim strSrc(0 To 2) As String  ' Person, number and gender split up
    Dim strDst(0 To 2) As String  ' Person, number and gender split up

    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (dtrAnt Is Nothing) Then Return False
      ' Check whether there is uncertainty in PGN
      If (Not IsCertainPgn(dtrSrc)) Then Return False
      If (Not IsCertainPgn(dtrAnt)) Then Return False
      ' Get source and destination agreement
      strPGNsrc = dtrSrc.Item("PGN")
      strPGNdst = dtrAnt.Item("PGN")
      ' Distentangle the source
      If (Not SplitAgree(strPGNsrc, strSrc(0), strSrc(1), strSrc(2))) Then
        ' Something went wrong
        Return False
      End If
      ' Distentangle the antecedent
      If (Not SplitAgree(strPGNdst, strDst(0), strDst(1), strDst(2))) Then
        ' Something went wrong
        Return False
      End If
      ' Compare source and antecedent PERSON
      Return ((strSrc(0) = intSrc) AndAlso (strDst(0) = intDst))
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AgrPerson error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CnsNoCrossEqSubject
  ' Goal:   Apply the constraint "NoCrossEqSubject"
  '         You can NOT point back from a subject inside a SPEECH IP to the subject 
  '           of the preceding non-speech IP.
  ' Return: Return 1 for violation, 0 for no violation, and -1 for error
  ' History:
  ' 25-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CnsNoCrossEqSubject(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    Dim strRoleSrc As String      ' Grammatical role of the source
    Dim strRoleAnt As String      ' Grammatical role of the antecedent
    Dim bIPsrc As Boolean         ' Whether the source constituent is a speech IP
    Dim bIPant As Boolean         ' Whether the antecedent is a speech IP
    Dim strSrc(0 To 2) As String  ' Person, number and gender split up
    Dim strDst(0 To 2) As String  ' Person, number and gender split up

    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (dtrAnt Is Nothing) Then Return -1
      ' Check if there is a cross speech boundary
      bIPsrc = (dtrSrc.Item("IsSpeech") Like "True")
      bIPant = (dtrAnt.Item("IsSpeech") Like "True")
      If (bIPant And bIPsrc) OrElse (Not bIPant And Not bIPsrc) Then
        ' Both are either speech IP or no speech IP, which means there is no boundary crossing
        Return 0
      End If
      ' There is a crossspeech boundary, so now we need to check the person agreement
      ' Get source and destination grammatical role
      strRoleSrc = dtrSrc.Item("GrRole")
      strRoleAnt = dtrAnt.Item("GrRole")
      ' If both are subject, then the proposed link is bad
      If (strRoleSrc = "Subject") AndAlso (strRoleAnt = "Subject") Then Return 1
      ' If the direct speech is imperative, then there is an additional requirement
      If (dtrSrc.Item("IPlabel") Like "*IMV*") Then
        ' Yes, the source IP is imperative. Now the antecedent may not be subject if the source is an argument
        If (strRoleSrc = "Argument") AndAlso (strRoleAnt = "Subject") Then Return 1
      End If
      ' Otherwise there are no problems (I think)
      Return 0
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsNoCrossEqSubject error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CnsNoCrossAgrPerson
  ' Goal:   Apply the constraint "NoCrossAgrPerson"
  '         There should be disagreement in PERSON at a crossspeech boundary
  ' Return: Return 1 for violation, 0 for no violation, and -1 for error
  ' History:
  ' 24-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CnsNoCrossAgrPerson(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    Dim strPGNsrc As String ' Person/Number/Gender of the source
    Dim strPGNdst As String ' Person/Number/Gender of the destination
    Dim bIPsrc As Boolean         ' Whether the source constituent is a speech IP
    Dim bIPant As Boolean         ' Whether the antecedent is a speech IP
    Dim strSrc(0 To 2) As String  ' Person, number and gender split up
    Dim strDst(0 To 2) As String  ' Person, number and gender split up

    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (dtrAnt Is Nothing) Then Return -1
      ' Check whether there is uncertainty in PGN
      If (Not IsCertainPgn(dtrSrc)) Then Return False
      If (Not IsCertainPgn(dtrAnt)) Then Return False
      ' =============== DEBUGGING ====================
      'If (dtrSrc.Item("Id") = 497) AndAlso ((dtrAnt.Item("Id") = 471) OrElse (dtrAnt.Item("Id") = 418)) Then
      '  Stop
      'End If
      ' ==============================================
      ' Check if there is a cross speech boundary
      bIPsrc = (dtrSrc.Item("IPlabel") Like "*SPE*")
      bIPant = (dtrAnt.Item("IPlabel") Like "*SPE*")
      If (bIPant And bIPsrc) OrElse (Not bIPant And Not bIPsrc) Then
        ' Both are either speech IP or no speech IP, which means there is no boundary crossing
        Return 0
      End If
      ' There is a crossspeech boundary, so now we need to check the person agreement
      ' Get source and destination agreement
      strPGNsrc = dtrSrc.Item("PGN")
      strPGNdst = dtrAnt.Item("PGN")
      ' Distentangle the source
      If (Not SplitAgree(strPGNsrc, strSrc(0), strSrc(1), strSrc(2))) Then
        ' Something went wrong
        Return False
      End If
      ' Distentangle the antecedent
      If (Not SplitAgree(strPGNdst, strDst(0), strDst(1), strDst(2))) Then
        ' Something went wrong
        Return False
      End If
      ' Compare source and antecedent PERSON
      If (strSrc(0) = strDst(0)) Then
        ' There is violation, because they DO agree, but they should not!
        Return 1
      ElseIf (strSrc(0) = "unknown") OrElse (strDst(0) = "unknown") Then
        ' When either is unknown, then user-checking needs to take place, so return a violation
        Return 1
      ElseIf (strSrc(0) = "empty") Then
        ' An "empty" source NP can refer to ANY preceding NP as far as PGN is concerned
        Return 0
      ElseIf (strDst(0) = "empty") Then
        ' This must be severely punished: it should be impossible to point to an empty DESTINATION
        Return 1
      Else
        ' They are not equal, and there is no special case, so: NO violation
        Return 0
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsNoCrossAgrPerson error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CnsAgrGenderNumber
  ' Goal:   Apply the constraint "AgreeGenderNumber"
  '         Is there possible agreement between the source and the destination?
  ' Return: Return 1 for violation, 0 for no violation, and -1 for error
  ' History:
  ' 24-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CnsAgrGenderNumber(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    Dim strPGNsrc As String ' Person/Number/Gender of the source
    Dim strPGNdst As String ' Person/Number/Gender of the destination
    Dim strSrc(0 To 2) As String  ' Person, number and gender split up
    Dim strDst(0 To 2) As String  ' Person, number and gender split up

    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (dtrAnt Is Nothing) Then Return -1
      ' ============= DEBUGGING =================
      'If (dtrSrc.Item("Id") = 10922) Then Stop
      ' =========================================
      ' Check whether there is uncertainty in PGN
      If (Not IsCertainPgn(dtrSrc)) Then Return -1
      If (Not IsCertainPgn(dtrAnt)) Then Return -1
      ' Get source and destination agreement
      strPGNsrc = dtrSrc.Item("PGN")
      strPGNdst = dtrAnt.Item("PGN")
      ' =========== DEBUGGING ==============
      'If (strPGNsrc = "3ms") AndAlso (strPGNdst = "3s") Then Stop
      'If (dtrSrc.Item("Vern") = "his") AndAlso (dtrAnt.Item("Vern") = "hi") Then Stop
      ' ====================================
      ' Distentangle the source
      If (Not SplitAgree(strPGNsrc, strSrc(0), strSrc(1), strSrc(2))) Then
        ' Something went wrong
        Return -1
      End If
      ' Distentangle the antecedent
      If (Not SplitAgree(strPGNdst, strDst(0), strDst(1), strDst(2))) Then
        ' Something went wrong
        Return -1
      End If
      ' Compare source and antecedent Gender and Number
      If (strSrc(1) = strDst(1)) AndAlso (strSrc(2) = strDst(2)) Then
        ' No violation whatsoever
        Return 0
      ElseIf (strSrc(1) = "unknown") OrElse (strDst(1) = "unknown") OrElse _
             (strSrc(2) = "unknown") OrElse (strDst(2) = "unknown") Then
        ' When either is unknown, then user-checking needs to take place, so return a violation
        Return 1
      ElseIf (strSrc(1) = "empty") Then
        ' There is person agreement -- what about number?
        If (strSrc(2) = strDst(2)) OrElse (strSrc(2) = "empty") Then
          ' There is possible agreement
          Return 0
        Else
          ' No real agreement
          Return 1
        End If
      ElseIf ((strDst(1) = "empty") AndAlso (strSrc(2) = strDst(2))) OrElse ((strDst(2) = "empty") AndAlso (strSrc(1) = strDst(1))) Then
        ' One of the destinations is unspecified, but the other matches
        Return 0
      ElseIf (strDst(1) = "empty") AndAlso (strDst(2) = "empty") Then
        ' Both destinations are not specified, so there is vacuous agreement
        Return 0
      Else
        ' They are not equal, and there is no special case, so: violation
        Return 1
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsAgrGenderNumber error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AgreePGN
  ' Goal:   Check agreement in Person, Gender and Number between two nodes
  ' Return: Return 1 for violation, 0 for no violation, and -1 for error
  ' History:
  ' 03-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function AgreePGN(ByRef ndxSrc As XmlNode, ByRef ndxAnt As XmlNode) As Boolean
    Dim strPGNsrc As String       ' Person/Gender/Number of the source
    Dim strPGNdst As String       ' Person/Gender/Number of the destination
    Dim strSrc(0 To 2) As String  ' Person, gender and number split up
    Dim strDst(0 To 2) As String  ' Person, gender and number split up
    Dim intI As Integer           ' Counter

    Try
      ' Validate
      If (ndxSrc Is Nothing) OrElse (ndxAnt Is Nothing) Then Return -1
      ' Check whether there is uncertainty in PGN
      If (Not MakeCertainPgn(ndxSrc)) Then Return -1
      If (Not MakeCertainPgn(ndxAnt)) Then Return -1
      ' Get source and destination agreement
      strPGNsrc = GetFeature(ndxSrc, "NP", "PGN")
      strPGNdst = GetFeature(ndxAnt, "NP", "PGN")
      ' Distentangle the source
      If (Not SplitAgree(strPGNsrc, strSrc(0), strSrc(1), strSrc(2))) Then
        ' Something went wrong
        Return False
      End If
      ' Distentangle the antecedent
      If (Not SplitAgree(strPGNdst, strDst(0), strDst(1), strDst(2))) Then
        ' Something went wrong
        Return False
      End If
      ' Check Person, Gender and Number individually...
      For intI = 0 To 2
        ' Are they the same?
        If (strSrc(intI) <> strDst(intI)) Then
          ' They are not the same - is one of them "unknown" or "empty"?
          If (Not DoLike(strSrc(intI), "unknown|empty") AndAlso Not DoLike(strDst(intI), "unknown|empty")) Then
            ' There is no real agreement, where there should be
            Return False
          End If
        End If
      Next intI
      ' No objections have been found, so return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AgreePGN error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AgreeGN
  ' Goal:   Check agreement in Gender and Number between two nodes
  '         (Don't count "Person")
  ' Return: Return 1 for violation, 0 for no violation, and -1 for error
  ' History:
  ' 23-11-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function AgreeGN(ByRef ndxSrc As XmlNode, ByRef ndxAnt As XmlNode) As Boolean
    Dim strPGNsrc As String       ' Person/Gender/Number of the source
    Dim strPGNdst As String       ' Person/Gender/Number of the destination
    Dim strSrc(0 To 2) As String  ' Person, gender and number split up
    Dim strDst(0 To 2) As String  ' Person, gender and number split up
    Dim intI As Integer           ' Counter

    Try
      ' Validate
      If (ndxSrc Is Nothing) OrElse (ndxAnt Is Nothing) Then Return -1
      ' Check whether there is uncertainty in PGN
      If (Not MakeCertainPgn(ndxSrc)) Then Return -1
      If (Not MakeCertainPgn(ndxAnt)) Then Return -1
      ' Get source and destination agreement
      strPGNsrc = GetFeature(ndxSrc, "NP", "PGN")
      strPGNdst = GetFeature(ndxAnt, "NP", "PGN")
      ' Distentangle the source
      If (Not SplitAgree(strPGNsrc, strSrc(0), strSrc(1), strSrc(2))) Then
        ' Something went wrong
        Return False
      End If
      ' Distentangle the antecedent
      If (Not SplitAgree(strPGNdst, strDst(0), strDst(1), strDst(2))) Then
        ' Something went wrong
        Return False
      End If
      ' Check Gender and Number individually...
      For intI = 1 To 2
        ' Are they the same?
        If (strSrc(intI) <> strDst(intI)) Then
          ' There should be complete agreement in gender and number
          '' They are not the same - is one of them "unknown" or "empty"?
          'If (Not DoLike(strSrc(intI), "unknown|empty") AndAlso Not DoLike(strDst(intI), "unknown|empty")) Then
          '  ' There is no real agreement, where there should be
          Return False
          'End If
        End If
      Next intI
      ' No objections have been found, so return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AgreeGN error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MismatchPGN
  ' Goal:   Check agreement in Person, Gender and Number between two nodes
  ' Return: Return TRUE if there is any mismatch (quite strict!)
  ' History:
  ' 03-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function MismatchPGN(ByRef ndxSrc As XmlNode, ByRef ndxAnt As XmlNode) As Boolean
    Dim strPGNsrc As String       ' Person/Gender/Number of the source
    Dim strPGNdst As String       ' Person/Gender/Number of the destination
    Dim strSrc(0 To 2) As String  ' Person, gender and number split up
    Dim strDst(0 To 2) As String  ' Person, gender and number split up
    Dim intI As Integer           ' Counter

    Try
      ' Validate
      If (ndxSrc Is Nothing) OrElse (ndxAnt Is Nothing) Then Return -1
      ' Check whether there is uncertainty in PGN
      If (Not MakeCertainPgn(ndxSrc)) Then Return -1
      If (Not MakeCertainPgn(ndxAnt)) Then Return -1
      ' Get source and destination agreement
      strPGNsrc = GetFeature(ndxSrc, "NP", "PGN")
      strPGNdst = GetFeature(ndxAnt, "NP", "PGN")
      ' Distentangle the source
      If (Not SplitAgree(strPGNsrc, strSrc(0), strSrc(1), strSrc(2))) Then
        ' Something went wrong
        Return True
      End If
      ' Distentangle the antecedent
      If (Not SplitAgree(strPGNdst, strDst(0), strDst(1), strDst(2))) Then
        ' Something went wrong
        Return True
      End If
      ' Check Person, Gender and Number individually...
      For intI = 0 To 2
        ' Are they the same?
        If (strSrc(intI) <> strDst(intI)) Then
          ' There is a mismatch, so we return negative
          Return True
        End If
      Next intI
      ' No mismatches have been found, so return false
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/MismatchPGN error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return True
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CnsAgrGenderSrc
  ' Goal:   Apply the constraint "AgrGenderSource"
  '         When the source has a gender (m,f,n) specified, then the antecedent should
  '           have the same gender. It may not be "unknown" nor "empty"
  ' Return: Return 1 for violation, 0 for no violation, and -1 for error
  ' History:
  ' 30-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CnsAgrGenderSrc(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    Dim strPGNsrc As String ' Person/Number/Gender of the source
    Dim strPGNdst As String ' Person/Number/Gender of the destination
    Dim strSrc(0 To 2) As String  ' Person, number and gender split up
    Dim strDst(0 To 2) As String  ' Person, number and gender split up

    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (dtrAnt Is Nothing) Then Return -1
      ' Check whether there is uncertainty in PGN
      If (Not IsCertainPgn(dtrSrc)) Then Return False
      If (Not IsCertainPgn(dtrAnt)) Then Return False
      ' Get source and destination agreement
      strPGNsrc = dtrSrc.Item("PGN")
      strPGNdst = dtrAnt.Item("PGN")
      ' Distentangle the source
      If (Not SplitAgree(strPGNsrc, strSrc(0), strSrc(1), strSrc(2))) Then
        ' Something went wrong
        Return -1
      End If
      ' Distentangle the antecedent
      If (Not SplitAgree(strPGNdst, strDst(0), strDst(1), strDst(2))) Then
        ' Something went wrong
        Return -1
      End If
      ' Compare source and antecedent Gender
      If (strSrc(1) = strDst(1)) Then
        ' No violation whatsoever
        Return 0
      ElseIf (strSrc(1).Length = 1) Then
        ' The source has a specific gender, but apparently the destination doesn't!
        Return 1
      ElseIf (strDst(1).Length = 1) Then
        ' The destination has a specific gender, but apparently the source doesn't!
        Return 1
      Else
        ' The source does not have a specific gender, so no violation
        Return 0
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsAgrGenderSrc error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CnsAgrClause
  ' Goal:   Apply the constraint "AgrClause"
  '         When the antecedent is an IP (having PGN = 3ns),then the source
  '           should either have 3s or 3ns
  ' Return: Return 1 for violation, 0 for no violation, and -1 for error
  ' History:
  ' 30-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CnsAgrClause(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    Dim strPGNsrc As String ' Person/Number/Gender of the source
    Dim strPGNdst As String ' Person/Number/Gender of the destination
    Dim strSrc(0 To 2) As String  ' Person, number and gender split up
    Dim strDst(0 To 2) As String  ' Person, number and gender split up

    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (dtrAnt Is Nothing) Then Return -1
      ' Get source and destination agreement
      strPGNsrc = dtrSrc.Item("PGN")
      strPGNdst = dtrAnt.Item("PGN")
      ' Check the condition
      If (strPGNdst = "3ns") Then
        ' Check one step further
        If (dtrAnt.Item("Label") Like "IP-*") Then
          ' The antecedent is an IP, so the source must match
          If (DoLike(strPGNsrc, "3s|3ns")) Then Return 0 Else Return 1
        End If
      End If
      ' Conditions are not met, so there is no violation
      Return 0
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsAgrClause error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CnsNoClause
  ' Goal:   Apply the constraint "NoClause"
  '         The antecedent may not be an IP
  ' Return: Return 1 for violation, 0 for no violation, and -1 for error
  ' History:
  ' 30-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CnsNoClause(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (dtrAnt Is Nothing) Then Return -1
      ' Check one step further
      If (dtrAnt.Item("Label") Like "IP-*") Then
        ' The antecedent is a clause
        Return 1
      Else
        Return 0
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsNoClause error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CnsAgree
  ' Goal:   Apply the constraint "Agree"
  '         Is there possible agreement between the source and the destination?
  ' Return: Return 1 for violation, 0 for no violation, and -1 for error
  ' History:
  ' 03-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CnsAgree(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    Dim strPGNsrc As String ' Person/Number/Gender of the source
    Dim strPGNdst As String ' Person/Number/Gender of the destination

    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (dtrAnt Is Nothing) Then Return -1
      ' Check whether there is uncertainty in PGN
      If (Not IsCertainPgn(dtrSrc)) Then Return False
      If (Not IsCertainPgn(dtrAnt)) Then Return False
      ' Get source and destination agreement
      strPGNsrc = dtrSrc.Item("PGN")
      strPGNdst = dtrAnt.Item("PGN")
      ' Compare the two
      If (strPGNdst = strPGNsrc) Then
        ' No violation whatsoever
        Return 0
      ElseIf (Left(strPGNdst, 1) <> Left(strPGNsrc, 1)) Then
        ' There is disagreement in PERSON, so no possible match
        Return 1
      ElseIf (strPGNdst = "unknown") OrElse (strPGNsrc = "unknown") Then
        ' When either is unknown, then user-checking needs to take place, so return a violation
        Return 1
      ElseIf (strPGNsrc = "empty") Then
        ' An "empty" source NP can refer to ANY preceding NP as far as PGN is concerned
        Return 0
      ElseIf (strPGNdst = "empty") Then
        ' This must be severely punished: it should be impossible to point to an empty DESTINATION
        Return 1
      Else
        ' Take away the person agreement
        strPGNdst = Mid(strPGNdst, 2)
        strPGNsrc = Mid(strPGNsrc, 2)
        ' Check if we go from specific to generic
        Select Case strPGNdst
          Case "ms", "fs", "ns"
            ' If the source is just "singular", then it is okay
            If (strPGNsrc = "s") Then
              ' Only little mismatch
              Return 1
            Else
              ' There is gender mismatch
              Return 1
            End If
          Case Else
            ' There are other mismatches
            Return 1
        End Select
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsAgree error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CnsDisjoint
  ' Goal:   Apply the constraint "Disjoint"
  '         Do both constituents have the same IP-MAT/SUB/INF above them?
  ' Return: 0 = they don't have the same IP-MAT/SUB/INF, or the source constituent is a possessive
  '         1 = they do have the same IP-MAT/SUB/INF
  '         -1 = something went wrong
  ' Note:   As soon as the NP is part of a PP, then it will not be disjoint with an antecedent
  '           due to c-command rules
  ' History:
  ' 03-06-2010  ERK Created
  ' 03-09-2010  ERK Added PP-not-disjoint condition
  ' 29-11-2010  ERK If the antecedent is LFD, it should not be disjoint
  ' ------------------------------------------------------------------------------------
  Private Function CnsDisjoint(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    Dim ndxSrc As XmlNode   ' The source node
    Dim ndxAnt As XmlNode   ' The antecedent node

    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (dtrAnt Is Nothing) Then Return -1
      ' Get source and destination nodes
      ndxSrc = IdToNode(dtrSrc.Item("Id"))
      ndxAnt = IdToNode(dtrAnt.Item("Id"))
      ' The real work is done by another one...
      If (IsDisjoint(ndxSrc, ndxAnt)) Then
        Return 1
      Else
        Return 0
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsDisjoint error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsDisjoint
  ' Goal:   Check if nodes are "Disjoint"
  '         Do both constituents have the same IP-MAT/SUB/INF above them?
  ' Return: True if they are disjoint or if there is an error
  '         False if they are really not disjoint
  ' Note:   As soon as the NP is part of a PP, then it will not be disjoint with an antecedent
  '           due to c-command rules
  ' History:
  ' 03-06-2010  ERK Created
  ' 03-09-2010  ERK Added PP-not-disjoint condition
  ' 29-11-2010  ERK If the antecedent is LFD, it should not be disjoint
  ' ------------------------------------------------------------------------------------
  Public Function IsDisjoint(ByRef ndxSrc As XmlNode, ByRef ndxAnt As XmlNode) As Boolean
    Dim ndxPP As XmlNode    ' Possible PP parent of the source
    Dim ndxPPant As XmlNode ' Possible PP parent of the antecedent
    Dim ndxSrcPar As XmlNode  ' Parent of source
    Dim ndxAntPar As XmlNode  ' Parent of antecedent
    Dim strSrcLabel As String ' Label of source
    Dim strAntLabel As String '  Label of antecedent

    Try
      ' Validate
      If (ndxSrc Is Nothing) OrElse (ndxAnt Is Nothing) Then Return True
      ' See if the source is part of a PP
      ndxPP = GetParentNode(ndxSrc, "PP*", strIPtypes)
      ' Do we have a PP parent?
      If (Not ndxPP Is Nothing) Then
        ' Check if the antecedent is part of the same PP
        ndxPPant = GetParentNode(ndxAnt, "PP*", strIPtypes)
        ' Check if they have the same PP antecedent
        Return (CompareNodes(ndxPP, ndxPPant) = 1)
      End If
      ' Get labels
      strSrcLabel = ndxSrc.Attributes("Label").Value
      strAntLabel = ndxAnt.Attributes("Label").Value
      ' Check whether the source node is a possessive
      If (DoLike(strSrcLabel, "PRO$*|NPR$*")) Then
        ' This is a possessive, for which there are other disjoint rules
        ' The source and target may not have the same NP kind
        ndxSrcPar = GetParentNode(ndxSrc, "NP|NP-*", strIPtypes)
        ndxAntPar = GetParentNode(ndxAnt, "NP|NP-*", strIPtypes)
        ' Compare the two
        Return (CompareNodes(ndxSrcPar, ndxAntPar) = 1)
      ElseIf (DoLike(strAntLabel, "*LFD*")) Then
        ' If the antecedent is a left dislocated element, there is no disjoint violation
        Return False
      Else
        ' Find out what the nearest parent IP-MAT is of the source
        ndxSrcPar = GetParentNode(ndxSrc, strIPtypes)
        ndxAntPar = GetParentNode(ndxAnt, strIPtypes)
        ' Compare the two
        Return (CompareNodes(ndxSrcPar, ndxAntPar) = 1)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/IsDisjoint error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return True
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CompareNodes
  ' Goal:   Compare whether two nodes are the same or not for CnsDisjoint
  ' History:
  ' 25-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CompareNodes(ByRef ndxSrc As XmlNode, ByRef ndxAnt As XmlNode) As Integer
    Try
      ' Compare the two
      If (ndxSrc Is Nothing) Then
        ' Check the antecedent
        If (ndxAnt Is Nothing) Then
          ' They both don't have an IP-MAT, so they belong to the same clause
          Return 1
        Else
          ' They cannot belong tot he same clause
          Return 0
        End If
      Else
        ' Check the antecedent
        If (ndxAnt Is Nothing) Then
          ' They cannot belong to the same clause
          Return 0
        Else
          ' See if they have the same ID
          If (ndxSrc.Attributes("Id").Value = ndxAnt.Attributes("Id").Value) Then
            ' They are the same
            Return 1
          Else
            ' They are different
            Return 0
          End If
        End If
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CompareNodes error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CnsNPtypeRel
  ' Goal:   Check whether the NP source type and NP target type match
  '         Principle: you cannot link from an indefinite source 
  ' Return: Return 1 for violation, 0 for no violation, and -1 for error
  ' History:
  ' 10-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CnsNPtypeRel(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    Dim strNPtypeSrc As String  ' The NP type of the source
    Dim strNPtypeDst As String  ' The NP type of the destination
    Dim strRelation As String   ' The relation between NP source and destination

    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (dtrAnt Is Nothing) Then Return -1
      ' Get source and destination NP type
      strNPtypeSrc = dtrSrc.Item("NPtype")
      strNPtypeDst = dtrAnt.Item("NPtype")
      ' See if either of these is "unknown"
      ' WAS: If (strNPtypeDst = "unknown") OrElse (strNPtypeSrc = "unknown") Then Return 0
      ' Calculate the relation
      strRelation = strNPtypeSrc & " > " & strNPtypeDst
      ' Number of violation marks depends on the relation
      Select Case strRelation
        Case "ZeroSbj > ZeroSbj", "ZeroSbj > Pro", "Proper > Proper"    ' Proper name to name
          Return 0
        Case "Pro > Pro"          ' Pronoun refers back to a pronoun
          Return 1
        Case "DefNP > DefNP", "Pro > DefNP"
          Return 3
        Case "DefNP > Pro", "DemNP > DemNP", "Pro > Proper", "DefNP > Proper", _
             "DemNP > Proper", "Pro > DemNP", "Pro > ZeroSbj", _
             "Proper > DefNP", "Proper > Pro"
          Return 4
        Case Else
          ' All other references are possible, but receive highest violation marks
          Return 5
      End Select
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsNPtypeRel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   CnsGrRoleRel
  ' Goal:   Check whether the source GrRole and target role match
  '         Principle: I don't know...
  ' Return: Return 1 for violation, 0 for no violation, and -1 for error
  ' History:
  ' 10-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CnsGrRoleRel(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    Dim strGrRoleSrc As String  ' The grammatical role of the source
    Dim strGrRoleDst As String  ' The grammatical role of the destination
    Dim strRelation As String   ' The relation between NP source and destination

    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (dtrAnt Is Nothing) Then Return -1
      ' Get source and destination NP type
      strGrRoleSrc = dtrSrc.Item("GrRole")
      strGrRoleDst = dtrAnt.Item("GrRole")
      ' See if either of these is "unknown"
      ' WAS: If (strGrRoleDst = "unknown") OrElse (strGrRoleSrc = "unknown") Then Return 0
      ' Calculate the relation
      strRelation = strGrRoleSrc & " > " & strGrRoleDst
      ' Action depends on the relation
      Select Case strRelation
        Case "Subject > Subject"
          ' This is the most occurring relation
          Return 0
        Case "PossDet > Subject"
          ' A possessive determiner mostly is a PRO$
          Return 3
        Case "PPobject > PPobject", "Subject > PossDet", "Subject > Argument", "Subject > PPobject", _
             "PPobject > Subject", "Argument > Subject", "PossDet > PossDet", "Argument > Argument", _
             "PossDet > PPobject", "PossDet > Argument"
          ' These relations are still quite likely
          Return 4
        Case Else
          ' All other references are possible, but are punished most severely
          Return 5
      End Select
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsGrRoleRel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CnsGrRoleSrc
  ' Goal:   Assign a preferential number to the grammatical role of an NP
  ' Return: Return number between 0..5, and -1 for error
  ' History:
  ' 12-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CnsGrRoleSrc(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    Dim strGrRoleSrc As String  ' The grammatical role of the source

    Try
      ' Validate
      If (dtrSrc Is Nothing) Then Return -1
      ' Get source and destination NP type
      strGrRoleSrc = dtrSrc.Item("GrRole")
      ' Action depends on the source GrRole
      Select Case strGrRoleSrc
        Case "Subject"
          ' This is the most occurring relation
          Return 0
        Case "PossDet", "PPobject", "Argument"
          ' Less likely
          Return 3
        Case Else
          ' All other references are possible, but are punished most severely
          Return 5
      End Select
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsGrRoleSrc error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CnsGrRoleDst
  ' Goal:   Assign a preferential number to the grammatical role of an NP
  ' Return: Return number between 0..5, and -1 for error
  ' History:
  ' 23-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CnsGrRoleDst(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    Dim strGrRoleSrc As String  ' The grammatical role of the source

    Try
      ' Validate
      If (dtrAnt Is Nothing) Then Return -1
      ' Get source and destination NP type
      strGrRoleSrc = dtrAnt.Item("GrRole")
      ' Action depends on the source GrRole
      Select Case strGrRoleSrc
        Case "Subject"
          ' This is the most occurring relation
          Return 0
        Case "PossDet", "Argument", "LeftDis"
          ' This is trial: it seems that Possessive Determiners are a bit more likely
          '   than PPobjects. I don't know about arguments, though
          Return 1
        Case "PPobject"
          ' Less likely
          Return 2
        Case Else
          ' All other references are possible, but are punished most severely
          Return 3
      End Select
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsGrRoleDst error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CnsNPtypeSrc
  ' Goal:   Assign a preferential number to the NP type of the source
  ' Return: Return number between 0..5, and -1 for error
  ' History:
  ' 12-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CnsNPtypeSrc(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    Dim strNPtypeSrc As String  ' The grammatical role of the source

    Try
      ' Validate
      If (dtrSrc Is Nothing) Then Return -1
      ' Get source and destination NP type
      strNPtypeSrc = dtrSrc.Item("NPtype")
      ' Action depends on the source NPtype
      Select Case strNPtypeSrc
        Case "ZeroSbj"
          ' This SHOULD receive the most attention
          Return 0
        Case "Pro", "Dem", "PossPro"
          ' This is the most occurring relation
          Return 1
        Case "Proper"
          Return 2
        Case "DefNP", "AnchoredNP"
          Return 3
        Case "DemNP"
          ' Less likely
          Return 4
        Case "IndefNP"
          ' An indefinite NP can never function as the source of a coreference relation!
          Return 6
        Case Else
          ' All other references are possible, but are punished most severely
          ' This includes: FUllNP, unknown, Bare, BareWithPP, QuantNP
          Return 5
      End Select
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsNPtypeSrc error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CnsNPtypeDst
  ' Goal:   Assign a preferential number to the NP type of the antecedent
  ' Return: Return number between 0..5, and -1 for error
  ' History:
  ' 23-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CnsNPtypeDst(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    Dim strNPtypeAnt As String  ' The grammatical role of the antecedent

    Try
      ' Validate
      If (dtrAnt Is Nothing) Then Return -1
      ' Get source and destination NP type
      strNPtypeAnt = dtrAnt.Item("NPtype")
      ' Action depends on the source NPtype
      Select Case strNPtypeAnt
        Case "ZeroSbj"
          ' This SHOULD receive the most attention
          Return 0
        Case "Pro"
          ' PRO is the most occurring relation
          Return 1
        Case "Proper"
          Return 2
        Case "DefNP", "AnchoredNP"
          Return 3
        Case "DemNP"
          ' Less likely
          Return 4
        Case Else
          ' All other references are possible, but are punished most severely
          Return 5
      End Select
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsNPtypeDst error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsNearDem
  ' Goal:   Check whether this is a near demonstrative
  ' History:
  ' 06-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function IsNearDem(ByRef dtrThis As DataRow) As Boolean
    Dim strNPtype As String  ' The grammatical role of the source

    Try
      ' Validate
      If (dtrThis Is Nothing) Then Return -1
      ' Get source and destination NP type
      strNPtype = dtrThis.Item("NPtype")
      ' Check if source type is a demonstrative
      Return (strNPtype = "Dem") AndAlso (DoLike(dtrThis.Item("Vern"), strNearDem))
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/IsNearDem error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CnsNearDem
  ' Goal:   Source must be a near demonstrative!
  '         One violation for an antecedent that already has a coreference, 
  '           unless the antecedent NP also contains a near demonstrative
  ' History:
  ' 06-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CnsNearDem(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (dtrAnt Is Nothing) Then Return -1
      ' =============== DEBUGGING ===============
      'If (dtrSrc.Item("Id") = 13304) Then
      '  If (dtrAnt.Item("Id") = 13299) Then
      '    Stop
      '  End If
      'End If
      ' =========================================
      ' Check source
      If (IsNearDem(dtrSrc)) Then
        ' Check antecedent
        If (Not IsNearDem(dtrAnt)) Then
          ' Only non-near demonstratives are allowed, if the antecedent has no coreference
          If (HasCoref(IdToNode(dtrAnt.Item("Id")))) Then
            Return 1
          End If
        End If
      End If
      ' No violation
      Return 0
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsNearDem error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CnsEqHead
  ' Goal:   One violation when the source NP has a head noun, and this head noun
  '           does not agree with any of the head nouns found in the chain of the target
  '         There is no violation if either the source or the target don't have a head noun
  ' Return: Return 0 for agreement, 1 for disagreement, and -1 for error
  ' History:
  ' 18-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CnsEqHead(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    Dim strHeadSrc As String  ' The head noun of the source

    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (dtrAnt Is Nothing) Then Return -1
      '' ============= DEBUGGING
      'If (dtrSrc.Item("Id") = 46) Then Stop
      '' ========================================
      ' Get the head noun of the source
      strHeadSrc = dtrSrc.Item("Head")
      ' Does the source have a head?
      If (strHeadSrc = "") Then Return 0
      ' Check if this head noun equals the target's head noun
      If (dtrAnt.Item("Head").ToString Like strHeadSrc) Then
        ' There is equality, so return success
        Return 0
      End If
      ' Check if this head noun appears in the target's chain
      If (IsInSemiStack(dtrAnt.Item("Equal") & "", strHeadSrc)) Then
        ' There is equality, so return success
        Return 0
      End If
      ' Otherwise check if this head noun could point to the antecedent according
      '   to the information saved in the chain dictionary
      If (ChainDictHas(strHeadSrc, dtrAnt.Item("Head"), dtrAnt.Item("Equal"))) Then
        ' There is a possible link, so return success
        Return 0
      End If
      ' No equality has been reached, so return a violation mark
      Return 1
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsEqHead error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetPrefNum
  ' Goal:   Get a preferential number for the indicated element
  '         The smaller the number, the higher the likelyhood that the source is referential
  ' History:
  ' 12-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetPrefNum(ByRef dtrSrc As DataRow) As Integer
    Dim intNPtype As Integer  ' Number for the NP type
    Dim intGrRole As Integer  ' Number for the grammatical role

    Try
      ' Validate
      If (dtrSrc Is Nothing) Then Return -1
      ' Get other numbers
      intNPtype = CnsNPtypeSrc(dtrSrc, Nothing)
      intGrRole = CnsGrRoleSrc(dtrSrc, Nothing)
      ' Validate
      If (intNPtype < 0) OrElse (intGrRole < 0) Then Return -1
      ' Return the preferential number
      Return intNPtype * 10 + intGrRole
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetPrefNum error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CnsIPdist
  ' Goal:   Count number of IP's between the source and the target
  ' Return: Return  -1 for error
  '         Otherwise return the number of IP's between me and antecedent
  ' History:
  ' 11-06-2010  ERK Created
  ' 17-12-2010  ERK Changed method of calculating the IP distance
  ' ------------------------------------------------------------------------------------
  Private Function CnsIPdist(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    'Dim ndxSrc As XmlNode   ' The source node
    'Dim ndxAnt As XmlNode   ' The antecedent node
    Dim intDist As Integer  ' The distance
    Dim intIpSrc As Integer ' Number of the source IP
    Dim intIpDst As Integer ' Number of the antecedent IP

    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (dtrAnt Is Nothing) Then Return -1
      '' Get the source and antecedent
      'ndxSrc = IdToNode(dtrSrc.Item("Id"))
      'ndxAnt = IdToNode(dtrAnt.Item("Id"))
      '' Get IP numbers
      'If (Not GetIpNumber(ndxSrc, intIpSrc)) Then Return -1
      'If (Not GetIpNumber(ndxAnt, intIpDst)) Then Return -1
      ' Get IP numbers directly from the [StackEl] entries
      intIpSrc = dtrSrc.Item("IPnum")
      intIpDst = dtrAnt.Item("IPnum")
      ' Calculate the distance
      If (intIpSrc < 0) OrElse (intIpDst < 0) Then
        intDist = 0
      Else
        ' Take the absolute value for the distance
        ' intDist = Math.Abs(intIpSrc - intIpDst)
        If (intIpSrc > intIpDst) Then
          intDist = intIpSrc - intIpDst
        Else
          intDist = intIpDst - intIpSrc
        End If
      End If
      ' Return the IP distance of the target
      Return intDist
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsIPdist error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CnsNoCataphore
  ' Goal:   One violation mark for an antecedent if it follows rather than preceds me
  ' Return: Return  -1 for error
  '         Otherwise return the number of IP's between me and antecedent
  ' History:
  ' 11-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function CnsNoCataphore(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (dtrAnt Is Nothing) Then Return -1
      ' Return the IP distance of the target
      Return IIf(dtrAnt.Item("Id") > dtrSrc.Item("Id"), 1, 0)
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsNoCataphore error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CnsAmbiguity
  ' Goal:   Count the number of full NP's in the antecedent's IP, to which no reference is 
  '           made from the current IP we're in.
  '         These full NP's must have the same IP ancestor as the antecedent we're looking at
  ' Return: Return  -1 for error
  '         Otherwise return the number of full NP's (not pronouns) in the IP
  '           of the antecedent [dtrAnt]
  ' History:
  ' 11-06-2010  ERK Created
  ' 04-10-2010  ERK Restrict ambiguity to antecedents with the PGN that is compatible with source
  ' ------------------------------------------------------------------------------------
  Private Function CnsAmbiguity(ByRef dtrSrc As DataRow, ByRef dtrAnt As DataRow) As Integer
    Dim dtrFound() As DataRow   ' Result of SELECT
    Dim ndxAntIp As XmlNode     ' The antecedent's IP node
    Dim intIPdist As Integer    ' The IPdistance of the target's IP
    Dim intNum As Integer = 0   ' Number of full NP's in the IP's target
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (dtrAnt Is Nothing) Then Return -1
      ' ============= PROFILING =============
      ProfileStart("CnsAmbiguity", 2)
      ' =====================================
      ' Get the IP distance of the target
      intIPdist = dtrAnt.Item("IPdist")
      ' Get the antecedent's IP node
      ndxAntIp = GetIpAncestor(IdToNode(dtrAnt.Item("Id")))
      ' Retrieve all NP's with this depth from the table
      dtrFound = tblAntStack.Select("IPdist=" & intIPdist)
      ' Count the number of full NP's 
      For intI = 0 To dtrFound.Length - 1
        ' Check if this is a full NP.
        ' Also check if it is a Dem, Pro (4/oct/2010)
        Select Case dtrFound(intI).Item("NPtype")
          Case "Proper", "DefNP", "DemNP", "Bare", "IndefNP", "unknown", "FullNP", "Dem", "Pro"
            ' Check to see if any constituent from the source stack makes a reference to this one
            If (Not HasLink(tblSrcStack, dtrFound(intI).Item("Id"))) Then
              ' Check compatibility of PGN
              If (CnsAgrGenderNumber(dtrSrc, dtrFound(intI)) = 0) Then
                ' Do we have an IP parent?
                If (ndxAntIp Is Nothing) Then
                  ' Increment the number of full NPs
                  intNum += 1
                Else
                  ' Double check to see that this antecedent has the same IP ancestor
                  ' (so an Argument and PP object, both full NP, in the same IP, are ambiguous too)
                  If (GetIpAncestor(IdToNode(dtrFound(intI).Item("Id"))) Is ndxAntIp) Then
                    ' Only now can we increment!
                    intNum += 1
                  End If
                End If
              End If
            End If
          Case "ZeroSbj", "Trace", "QuantNP"
            ' No need to count a pronoun or an empty category
            ' Besides: it is impossible to refer back to a QuantNP from another IP
            ' N.B: the category PRO$ is one that is filled in since it has not been coded properly...
        End Select
      Next intI
      ' ============= PROFILING =============
      ProfileEnd("CnsAmbiguity", 2)
      ' =====================================
      ' =========== DEBUGGING ========
      'If (intNum > 0) Then Stop
      ' ==============================
      ' Return the number of full NP's found
      Return intNum
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/CnsAmbiguity error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   HasLink
  ' Goal:   See if any of the elements in the stack [tblThis] contains a "Ref" pointer
  '           to the element with [intId]
  ' History:
  ' 11-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function HasLink(ByRef tblThis As DataTable, ByVal intId As Integer) As Boolean
    Dim intI As Integer   ' Counter

    Try
      ' validate
      If (tblThis Is Nothing) OrElse (intId < 0) Then Return False
      ' Check the stack
      For intI = 0 To tblThis.Rows.Count - 1
        ' Is there any reference?
        If (tblThis.Rows(intI).Item("Ref") & "" <> "") Then
          ' Check this element
          If (tblThis.Rows(intI).Item("Ref") = intId) Then
            ' Okay there is a link, so return true
            Return True
          End If
        End If
      Next intI
      ' Found nothing, so return false
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/HasLink error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return falure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetParentNode
  ' Goal:   Pass back the number of (equally) best matches
  ' Note:   The [strPattern] may consist of several patterns separated by |
  '         The [strPattern] label should be met before a label matching [strLimit] is reached!!
  ' History:
  ' 03-06-2010  ERK Created
  ' 06-09-2010  ERK Added [strLimit]
  ' ------------------------------------------------------------------------------------
  Private Function GetParentNode(ByRef ndxThis As XmlNode, ByVal strPattern As String, _
                                 Optional ByVal strLimit As String = "") As XmlNode
    Dim ndxNext As XmlNode  ' The next higher node
    Dim bFound As Boolean   ' Flag to match the result

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return Nothing
      ' Start with myself
      ndxNext = ndxThis
      ' Find out if we match
      If (ndxNext.Name = "eTree") Then
        bFound = DoLike(ndxNext.Attributes("Label").Value, strPattern)
      Else
        bFound = False
      End If
      ' Go into a loop
      While (Not bFound) AndAlso (Not ndxNext Is Nothing)
        ' Go one step higher
        ndxNext = ndxNext.ParentNode
        ' Find out if we match
        If (Not ndxNext Is Nothing) Then
          If (ndxNext.Name = "eTree") Then
            ' First check if we are not passing the [strLimit], if that is defined
            If (strLimit <> "") Then
              If DoLike(ndxNext.Attributes("Label").Value, strLimit) Then
                ' We have reached our limit, so retreat
                Return Nothing
              End If
            End If
            ' Now check if we are matching the pattern
            bFound = DoLike(ndxNext.Attributes("Label").Value, strPattern)
          End If
        End If
      End While
      ' Return the parent node
      Return ndxNext
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetParentNode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   NumBest
  ' Goal:   Pass back the number of (equally) best matches
  ' History:
  ' 03-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function NumBest(ByRef tblAnt As DataTable, _
                           ByRef intBest As Integer, ByRef intEval As Integer) As Integer
    Dim intNum As Integer = 0     ' Number of elements with the same evaluation as the best one
    Dim intI As Integer           ' Counter

    Try
      ' Validate
      If (tblAnt Is Nothing) Then Return 0
      If (tblAnt.Rows.Count = 0) Then Return 0
      ' ============= PROFILING =============
      ProfileStart("NumBest", 2)
      ' =====================================
      ' Get the evaluation of the FIRST element
      intEval = tblAnt(0).Item("Eval") : intBest = 0 : intNum = 1
      ' See which others improve
      For intI = 1 To tblAnt.Rows.Count - 1
        ' Is this an improvement?
        If (tblAnt(intI).Item("Eval") < intEval) Then
          ' Store this element
          intBest = intI
          ' Store this evaluation
          intEval = tblAnt(intI).Item("Eval")
          ' Reset the number of elements with this evaluation
          intNum = 1
        ElseIf (tblAnt(intI).Item("Eval") = intEval) Then
          ' Increment the number of elements with the best evaluation
          intNum += 1
        End If
      Next intI
      ' ============= PROFILING =============
      ProfileEnd("NumBest", 2)
      ' =====================================
      ' Return the number of best matches
      Return intNum
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/NumBest error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return 0
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsReflexivePronoun
  ' Goal:   Is this current NP a reflexive pronoun?
  ' History:
  ' 03-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function IsReflexivePronoun(ByRef ndxThis As XmlNode) As Boolean
    Dim ndxChild As XmlNode ' The child
    Dim strText As String   ' The text of the potential reflexive pronoun

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Is this an NP?
      If (ndxThis.Attributes("Label").Value Like "NP*") Then
        ' Get the first child
        ndxChild = ndxThis.Item("eTree")
        ' Check if this is valid
        If (ndxChild Is Nothing) Then Return False
      Else
        ' Take this to be the child
        ndxChild = ndxThis
      End If
      ' This element should have the label PRO*RFL
      If (ndxChild.Attributes("Label").Value Like "PRO*") Then
        ' It could be overtly marked as reflexive
        If (ndxChild.ParentNode.Attributes("Label").Value Like "*RFL*") Then
          ' This is a reflexive 
          Return True
        Else
          ' Okay, now we should get the text of this pronoun
          strText = NodeText(ndxChild)
          ' See if this fulfills the criteria
          Return (strText Like "*s*l*[fv]*")
        End If
      End If
      ' Unable to resolve it to a reflexive
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/IsReflexivePronoun error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsLocalType
  ' Goal:   An NP is locally resolvable, if any of the following holds:
  '         - it is a reflexive pronoun
  '         - it is an appositive NP
  '         - it is a zero of type *con* (subject elided under conjunction) or *pro* (same)
  ' Return: TRUE if it is resolvable, and [strType] gives the situation:
  '         "appositive"
  '         "reflexive"
  '         "zero"
  '         "barecompl"
  '         "complement"
  '         "posspro"
  '         "vocative"
  '         "another"   - NPs with "another" are new or contrastive (inferred)
  '         "freerel"   - Free relatives are to be taken as new
  '         "d-rel"     - NP = D + CP-REL --> Usually new like "freerel"
  ' History:
  ' 03-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function IsLocalType(ByRef ndxThis As XmlNode, ByRef strType As String) As Boolean
    Dim ndxChild As XmlNode   ' The child
    Dim ndxSister As XmlNode  ' Sister
    Dim strText As String     ' The text of the potential reflexive pronoun
    Dim strLabel As String    ' My own Label value
    Dim strChLabel As String  ' Label value of my child
    Dim strGrRole As String   ' Grammatical role

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' ========== DEBUGGING ==============
      'If (ndxThis.Attributes("Id").Value = 42) Then Stop
      ' ===================================
      ' Get my label
      strLabel = ndxThis.Attributes("Label").Value
      ' Is this an appositive NP?
      If (strLabel Like "NP*PRN*") Then
        ' This is an appositive NP
        strType = "appositive"
        Return True
      End If
      ' See if this is a free relative
      ndxChild = Nothing : ndxSister = Nothing
      If (SomeChild(ndxThis, "CP-FRL*", ndxChild)) Then
        ' This is a free relative
        strType = "freerel"
        Return True
      End If
      ' See if this is a left dislocated NP
      If (strLabel Like "NP*LFD*") Then
        ' Double check to see if this is not a Trace left dislocated element
        If (Not IsStarred(ndxThis)) Then
          ' This is a left dislocated NP, which is like a free-relative or d-relative: should be new
          strType = "lfd"
          Return True
        End If
      End If
      ' See if this is a D-Relative: D + CP-REL
      If (FirstChildImm(ndxThis, "D|D^*", ndxChild, "CODE|,|.")) Then
        ' See if there is a relative clause child
        If (SomeChild(ndxThis, "CP-REL*", ndxChild)) Then
          ' This is a d-relative
          strType = "d-rel"
          Return True
        End If
      End If
      ' See if this is "another"
      If (FirstChildImm(ndxThis, "D+OTHER", ndxChild, "CODE|,|.")) Then
        ' This is an NP headed by "another"
        strType = "another"
        Return True
      End If
      ' ============= DEBUGGING ============
      'If (ndxThis.Attributes("Label").Value = "PRO+N") Then Stop
      ' ====================================
      ' Is this perhaps a zero subject?
      If (strLabel Like "NP*SBJ*") Then
        ' Does it have a straight first <eLeaf> child?
        ndxChild = ndxThis.Item("eLeaf")
        If (Not ndxChild Is Nothing) Then
          ' There is a first <eLeaf> child - check its status
          If (ndxChild.Attributes("Type").Value = "Star") AndAlso _
             ((ndxChild.Attributes("Text").Value = "*con*") OrElse _
              (ndxChild.Attributes("Text").Value = "*pro*")) Then
            ' We got a zero subject *con*
            strType = "zero"
            Return True
          End If
        End If
      End If
      ' Is this an NP?
      If (strLabel Like "NP*") Then
        ' Get the first child
        ndxChild = ndxThis.Item("eTree")
        ' Check if this is valid
        If (ndxChild Is Nothing) Then Return False
      Else
        ' Take this to be the child
        ndxChild = ndxThis
      End If
      ' Get my child label
      strChLabel = ndxChild.Attributes("Label").Value
      ' This element should have the label PRO*RFL
      If (strChLabel Like "PRO*") Then
        ' Perhaps it is a possessive?
        If (strChLabel Like "PRO$*") Then
          ' Found a possessive pronoun
          strType = "posspro"
          Return True
        End If
        ' Okay, now we should get the text of this pronoun
        strText = NodeText(ndxChild)
        ' See if this fulfills the criteria
        If ((strText Like "*s*l*[fv]*") OrElse (ndxChild.ParentNode.Attributes("Label").Value Like "*RFL*")) AndAlso _
           (ndxChild.Attributes("Id").Value <> ndxThis.Attributes("Id").Value) Then
          strType = "reflexive"
          Return True
        End If
      End If
      ' Looking for bare NP complements. It must be an NP
      '   ADDITION: make sure it is NOT a proper noun
      If (strLabel Like "NP*") AndAlso (Not strLabel Like "*VOC*") Then
        ' Find out if it is a bare NP by looking at its feature
        If (GetFeature(ndxThis, "NP", "NPtype") Like "Bare*") Then
          ' Get the grammatical role
          strGrRole = GetFeature(ndxThis, "NP", "GrRole")
          ' It is a bare NP. Now see if it is a PPobject or what
          Select Case strGrRole
            Case "PPobject"
              ' It is a bare NP object of a PP!
              strType = "barecompl"
              Return True
            Case "Argument", "Oblique"
              ' Now we should check two more things to determine whether this is an NP complement of a be clause
              ' (1) Do we have a finite form of "BE" (or any other auxiliary) as sister node?
              If (HasSister(ndxThis, strFiniteBe)) Then
                ' (2) Is there a sister node that is subject?
                If (HasSister(ndxThis, strSubjectType)) Then
                  ' Yes, we have a subject sister
                  strType = "barecompl"
                  Return True
                End If
              End If
          End Select
        End If
      End If
      ' Get vocatives
      If (strLabel Like "*VOC*") Then
        ' We have a vocative NP that is not a bare noun, so for instance a proper name
        strType = "vocative"
        Return True
      End If
      ' The NP complement of a [Sbj Be/Have Ob] phrase must be selected out too,
      '   unless this complement is followed by a CP
      If (strLabel Like "NP-OB*") Then
        ' We are dealing with an NP object
        ' But we want to EXCLUDE certain categories
        If (Not DoLike(strChLabel, "NPR*")) Then
          ' does it have a finite Aux sister?
          If (HasSister(ndxThis, strFiniteBe)) Then
            ' It should NOT have a finite verb sister
            If (Not HasSister(ndxThis, "V*")) Then
              ' It should NOT have a CP sister
              If (Not HasSister(ndxThis, "CP*")) Then
                ' Does it have a subject sister?
                If (GetSister(ndxThis, strSubjectType, ndxSister)) Then
                  '' Then in all likelihood this is a complement NP
                  'strType = "complement"
                  'Return True
                  ' (3) The subject must not be an expletive
                  If (Not SomeChild(ndxSister, "EX", ndxChild)) Then
                    ' (4) The subject must not be a pronoun like "it", "hit"
                    If (SomeChild(ndxSister, "PRO|PRO^*|PRO-*", ndxChild)) Then
                      ' Check what pronoun it is
                      If (Not DoLike(LeafText(ndxChild), "hit|it")) Then
                        ' We conclude it is a complement
                        strType = "complement"
                        Return True
                      End If
                    Else
                      ' We conclude it is a complement
                      strType = "complement"
                      Return True
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
      ' Unable to resolve it to a special type
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/IsLocalType error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   HasSister
  ' Goal:   See if this node has a sister fulfilling the [strSisterForm] criteria
  ' History:
  ' 08-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function HasSister(ByRef ndxThis As XmlNode, ByVal strSisterForm As String) As Boolean
    Dim ndxNext As XmlNode  ' Next or previous node

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Try preceding sisters
      ndxNext = ndxThis.PreviousSibling
      While (Not ndxNext Is Nothing)
        ' Check this sister
        If (ndxNext.Name = "eTree") Then
          ' Check the label
          If (DoLike(ndxNext.Attributes("Label").Value, strSisterForm)) Then
            ' Found the correct element
            Return True
          End If
        End If
        ' Go to the previous sibling
        ndxNext = ndxNext.PreviousSibling
      End While
      ' Try following sisters
      ndxNext = ndxThis.NextSibling
      While (Not ndxNext Is Nothing)
        ' Check this sister
        If (ndxNext.Name = "eTree") Then
          ' Check the label
          If (DoLike(ndxNext.Attributes("Label").Value, strSisterForm)) Then
            ' Found the correct element
            Return True
          End If
        End If
        ' Go to the next sibling
        ndxNext = ndxNext.NextSibling
      End While
      ' Found nothing so return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/HasSister error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetSister
  ' Goal:   Get a sister fulfilling the [strSisterForm] criteria
  ' History:
  ' 08-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetSister(ByRef ndxThis As XmlNode, ByVal strSisterForm As String, ByRef ndxNext As XmlNode) As Boolean
    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Try preceding sisters
      ndxNext = ndxThis.PreviousSibling
      While (Not ndxNext Is Nothing)
        ' Check this sister
        If (ndxNext.Name = "eTree") Then
          ' Check the label
          If (DoLike(ndxNext.Attributes("Label").Value, strSisterForm)) Then
            ' Found the correct element
            Return True
          End If
        End If
        ' Go to the previous sibling
        ndxNext = ndxNext.PreviousSibling
      End While
      ' Try following sisters
      ndxNext = ndxThis.NextSibling
      While (Not ndxNext Is Nothing)
        ' Check this sister
        If (ndxNext.Name = "eTree") Then
          ' Check the label
          If (DoLike(ndxNext.Attributes("Label").Value, strSisterForm)) Then
            ' Found the correct element
            Return True
          End If
        End If
        ' Go to the next sibling
        ndxNext = ndxNext.NextSibling
      End While
      ' Found nothing so return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetSister error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetPossProAnt
  ' Goal:   Get the antecedent of a possessive pronoun, if possible!
  '         There are three possibilities:
  '         (1): part of argument NP --> point to subject NP
  '         (2): part of oblique PP --> point to subject NP
  '         (3): part of conjunctive NP --> point to conj. NP head
  ' History:
  ' 03-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetPossProAnt(ByRef ndxThis As XmlNode, ByRef ndxAnt As XmlNode) As Boolean
    Dim ndxNext As XmlNode    ' Working node
    Dim ndxNPsrc As XmlNode   ' The source node -- parent of the PRO$
    Dim strLabel As String    ' Value of one label
    Dim strBack As String     ' Found state

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' =================== DEBUGGING ==================
      'If (ndxThis.Attributes("Id").Value = 8009) Then Stop
      ' ================================================
      ' Get the parent NP
      ndxNPsrc = ndxThis.ParentNode : If (ndxNPsrc Is Nothing) Then Return False
      ' Get the label
      strLabel = ndxNPsrc.Attributes("Label").Value
      ' Validate again
      If (Not strLabel Like "NP*") Then Return False
      ' Initialise
      ndxAnt = Nothing
      ' The NP source should also not be a subject!
      If (InStr(strLabel, "SBJ") > 0) Then Return False
      ' Okay, we have established a non-subject NP parent of the PRO$...
      ' Determine what kind of parent we have...
      ndxNext = ndxNPsrc.ParentNode : If (ndxNext Is Nothing) Then Return False
      strLabel = ndxNext.Attributes("Label").Value
      ' There are three possible parents we can continue with
      If (strLabel Like "CONJP*") Then
        ' Situation (3)? We are part of a conjunctive NP: NP[ NP CONJP[ CONJ NP[ PRO$ ...]]]
        ' Check whether this CONJP is directly under an NP
        ndxNext = ndxNext.ParentNode : If (ndxNext Is Nothing) Then Return False
        ' Check if this is not a forest
        If (ndxNext.Name <> "eTree") Then Return False
        ' Get the label
        strLabel = ndxNext.Attributes("Label").Value
        If (strLabel Like "NP*") Then
          ' Get the first NP child -- that is the antecedent
          If (GetFirstNPchild(ndxNext, ndxAnt)) Then
            ' Double check: it should not be a trace
            If (Not IsStarred(ndxAnt)) Then Return True
          End If
        End If
      ElseIf (strLabel Like "PP*") Then
        ' Situation (2)? We are part of a PP: IP[ SBJ .. PP[ P NP[PRO$ ...]]]
        ' Check whether this PP is directly under an IP
        ndxNext = ndxNext.ParentNode : If (ndxNext Is Nothing) Then Return False
        ' Check if this is not a forest
        If (ndxNext.Name <> "eTree") Then Return False
        ' Get the label
        strLabel = ndxNext.Attributes("Label").Value
        If (strLabel Like "IP*") AndAlso (Not strLabel Like "*-INF*") Then
          ' Get the subject of this IP (we are not the subject)
          If (GetIpSubjectChild(ndxNext, ndxAnt)) Then
            ' Check mismatches in person/number/gender
            If (Not MismatchPGN(ndxThis, ndxAnt)) AndAlso (Not IsStarred(ndxAnt)) Then
              ' Return success
              Return True
            End If
          End If
        End If
      ElseIf (strLabel Like "IP*") AndAlso (Not strLabel Like "*-INF*") Then
        ' Situation (1): we are an argument of the IP
        ' Get the subject of this IP (we are not the subject)
        If (GetIpSubjectChild(ndxNext, ndxAnt)) Then
          ' Check to see if there is an NP that PRECEDES the subject of the IP
          strBack = ""
          If (Not SeekInterveningNP(ndxAnt.ParentNode, ndxAnt, strBack)) Then Return False
          ' See if an intervener was found
          If (strBack = "NotFound") AndAlso (Not IsStarred(ndxAnt)) Then
            ' No intervening NP was found, so we are safe
            Return True
          End If
        End If
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetPossProAnt error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetVocAnt
  ' Goal:   Get the antecedent of a vocative, if possible!
  '         There is, right now, one possibility:
  '         (1) point to subject NP of the IP we are in
  ' History:
  ' 19-11-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetVocAnt(ByRef ndxThis As XmlNode, ByRef ndxAnt As XmlNode) As Boolean
    Dim ndxNext As XmlNode    ' Working node
    Dim strLabel As String    ' Value of one label

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Initialise
      ndxAnt = Nothing
      ' Determine what kind of parent we have...
      ndxNext = ndxThis.ParentNode : If (ndxNext Is Nothing) Then Return False
      strLabel = ndxNext.Attributes("Label").Value
      ' There are three possible parents we can continue with
      If (strLabel Like "IP*") AndAlso (Not strLabel Like "*-INF*") Then
        ' Situation (1): we are an argument of the IP
        ' Get the subject of this IP (we are not the subject)
        If (GetIpSubjectChild(ndxNext, ndxAnt)) Then
          ' Vocative should agree in gender at least
          If (AgreeGN(ndxThis, ndxAnt)) Then
            ' We have found the correct antecedent
            Return True
          End If
        End If
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetVocAnt error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SeekInterveningNP
  ' Goal:   Function to recursively walk the <eTree> elements until either:
  '         (1) the node <ndxAnt> is reached, or
  '         (2) a node with label NP*/PRO$ is reached
  ' Note:   Success is stored in the bFound parameter
  ' History:
  ' 06-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function SeekInterveningNP(ByRef ndxThis As XmlNode, ByRef ndxAnt As XmlNode, ByRef strBack As String) _
                                      As Boolean
    Dim ndxList As XmlNodeList  ' List of child nodes
    Dim ndxChild As XmlNode     ' One child
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' First test this node
      If (ndxThis.Name = "eTree") Then
        ' ============== debugging
        'If (ndxThis.Attributes("Id").Value = 13299) Then Stop
        ' ==================
        ' Check if they are equal
        If (EqualNodes(ndxThis, ndxAnt)) Then
          ' We have reached the antecedent before an intervener is found
          strBack = "NotFound"
          Return True
        ElseIf (DoLike(ndxThis.Attributes("Label").Value, strNPsourceTypes)) Then
          ' An intervening node of NP type is found
          strBack = "Found"
          Return True
        End If
      End If
      ' Nothing is found, so go depth first
      ' Visit the children
      ndxList = ndxThis.ChildNodes
      For intI = 0 To ndxList.Count - 1
        ' Get the current node
        ndxChild = ndxList(intI)
        ' Perform the action on this child recursively
        If (Not SeekInterveningNP(ndxChild, ndxAnt, strBack)) Then
          ' There is an error, so return
          Return False
        End If
        ' If something has been found, we need to stop
        If (strBack <> "") Then Return True
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/SeekInterveningNP error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   HasEtreeChild
  ' Goal:   Does this node have any children of type <eTree>?
  ' History:
  ' 06-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function HasEtreeChild(ByRef ndxThis As XmlNode) As Boolean
    Dim ndxNext As XmlNode  ' Working node

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Go to first node
      ndxNext = ndxThis.FirstChild
      ' Try in loop
      While (Not ndxNext Is Nothing)
        ' Is this an [eTree]?
        If (ndxNext.Name = "eTree") Then Return True
        ' Go to next child
        ndxNext = ndxNext.NextSibling
      End While
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/HasEtreeChild error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   HasEtreeSibling
  ' Goal:   Does this node have any siblings of type <eTree>?
  ' History:
  ' 06-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function HasEtreeSibling(ByRef ndxThis As XmlNode) As Boolean
    Dim ndxNext As XmlNode  ' Working node

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Go to first node
      ndxNext = ndxThis.NextSibling
      ' Try in loop
      While (Not ndxNext Is Nothing)
        ' Is this an [eTree]?
        If (ndxNext.Name = "eTree") Then Return True
        ' Go to next child
        ndxNext = ndxNext.NextSibling
      End While
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/HasEtreeSibling error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetEtreeFirstChild
  ' Goal:   Get the first <eTree> child of this node
  ' History:
  ' 06-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetEtreeFirstChild(ByRef ndxThis As XmlNode) As XmlNode
    Dim ndxNext As XmlNode  ' Working node

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return Nothing
      ' Go to first node
      ndxNext = ndxThis.FirstChild
      ' Try in loop
      While (Not ndxNext Is Nothing)
        ' Is this an [eTree]?
        If (ndxNext.Name = "eTree") Then
          ' Return this node
          Return ndxNext
        End If
        ' Go to next child
        ndxNext = ndxNext.NextSibling
      End While
      ' Return failure
      Return Nothing
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetEtreeFirstChild error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetEtreeNextSibling
  ' Goal:   Get the next <eTree> sibling of this node
  ' History:
  ' 06-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetEtreeNextSibling(ByRef ndxThis As XmlNode) As XmlNode
    Dim ndxNext As XmlNode  ' Working node

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return Nothing
      ' Go to first node
      ndxNext = ndxThis.NextSibling
      ' Try in loop
      While (Not ndxNext Is Nothing)
        ' Is this an [eTree]?
        If (ndxNext.Name = "eTree") Then
          ' Return this node
          Return ndxNext
        End If
        ' Go to next child
        ndxNext = ndxNext.NextSibling
      End While
      ' Return failure
      Return Nothing
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetEtreeNextSibling error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   EqualNodes
  ' Goal:   See if the nodes are equal [eTree] ones
  ' History:
  ' 06-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function EqualNodes(ByRef ndxThis As XmlNode, ByRef ndxAnt As XmlNode) As Boolean
    Try
      ' Validate
      If (ndxAnt Is Nothing) OrElse (ndxThis Is Nothing) Then Return False
      If (ndxThis.Name <> "eTree") OrElse (ndxAnt.Name <> "eTree") Then Return False
      ' Compare the IDs
      Return (ndxThis.Attributes("Id").Value = ndxAnt.Attributes("Id").Value)
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/EqualNodes error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetFirstNPchild
  ' Goal:   Get the first child that is an NP
  '         Input to this function should not be empty
  ' History:
  ' 03-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetFirstNPchild(ByRef ndxThis As XmlNode, ByRef ndxAnt As XmlNode) As Boolean
    Dim ndxNext As XmlNode  ' Working node

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' NO NEED to check the label
      ' Start from first child
      ndxNext = ndxThis.FirstChild
      ' Go in a loop
      While (Not ndxNext Is Nothing)
        ' Try this element
        If (ndxNext.Name = "eTree") Then
          ' Is it an NP
          If (ndxNext.Attributes("Label").Value Like "NP*") Then
            ' Got you!
            ndxAnt = ndxNext
            Return True
          End If
        End If
        ' Get the following element
        ndxNext = ndxNext.NextSibling
      End While
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetFirstNPchild error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetIpSubjectChild
  ' Goal:   Get the first child labeled as having grammatical role "Subject" 
  '         Input to this function should be an IP node
  ' History:
  ' 03-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetIpSubjectChild(ByRef ndxThis As XmlNode, ByRef ndxSubj As XmlNode) As Boolean
    Dim ndxNext As XmlNode  ' Working node
    Dim ndxPrev As XmlNode  ' Previous node (working)

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Check the label
      If (Not ndxThis.Attributes("Label").Value Like "IP*") Then Return False
      ' Start from first child
      ndxNext = ndxThis.FirstChild : ndxPrev = Nothing
      ' Go in a loop
      While (Not ndxNext Is Nothing)
        ' Try this element
        If (ndxNext.Name = "eTree") Then
          ' Is it a subject?
          If (GetFeature(ndxNext, "NP", "GrRole") = "Subject") Then
            ' Got you!
            ndxSubj = ndxNext
            Return True
          End If
        End If
        ' Keep trak of previous
        ndxPrev = ndxNext
        ' Get the following element
        ndxNext = ndxNext.NextSibling
        ' Have we got a next one?
        If (ndxNext Is Nothing) Then
          ' Try going upwards
          ndxNext = ndxPrev.ParentNode : If (ndxNext Is Nothing) Then Return False
          ' One more upwards
          ndxNext = ndxNext.ParentNode : If (ndxNext Is Nothing) Then Return False
          ' Now take the first child
          ndxNext = ndxNext.FirstChild
        End If
      End While
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetIpSubjectChild error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetIpSubjectNode
  ' Goal:   Get the first constituent labeled as having grammatical role "Subject" 
  '         Do this by going left or up in the tree starting from [ndxThis]
  ' History:
  ' 01-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetIpSubjectNode(ByRef ndxThis As XmlNode, ByRef ndxSubj As XmlNode) As Boolean
    Dim ndxNext As XmlNode  ' Working node

    Try
      ' Start from where we are
      ndxNext = ndxThis
      ' Go in a loop
      While (Not ndxNext Is Nothing)
        ' Try this element
        If (ndxNext.Name = "eTree") Then
          ' Is it a subject?
          If (GetFeature(ndxNext, "NP", "GrRole") = "Subject") Then
            ' Got you!
            ndxSubj = ndxNext
            Return True
          End If
        End If
        ' Get the following element
        ndxNext = IIf(ndxNext.PreviousSibling Is Nothing, ndxNext.ParentNode, ndxNext.PreviousSibling)
      End While
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetIpSubjectNode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetIpSubject
  ' Goal:   Get the first constituent labeled as having grammatical role "Subject" 
  ' History:
  ' 03-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetIpSubject(ByRef tblThis As DataTable, ByRef tblAnt As DataTable, _
                                ByRef ndxSubj As XmlNode) As Boolean
    Dim intI As Integer     ' Counter

    Try
      ' Validate
      If (tblThis Is Nothing) Then Return False
      ' Walk all elements in the current IP's table
      For intI = 0 To tblThis.Rows.Count - 1
        ' Does this one have the correct grammatical role?
        If (tblThis.Rows(intI).Item("GrRole") = "Subject") Then
          ' We found the subject, so return it
          ndxSubj = IdToNode(tblThis.Rows(intI).Item("Id"))
          Return True
        End If
      Next intI
      ' Try the antecedents in BACKWARDS direction
      For intI = tblAnt.Rows.Count - 1 To 0 Step -1
        ' Does this one have the correct grammatical role?
        If (tblAnt.Rows(intI).Item("GrRole") = "Subject") Then
          ' We found the subject, so return it
          ndxSubj = IdToNode(tblAnt.Rows(intI).Item("Id"))
          Return True
        End If
      Next intI
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetIpSubject error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetPrecSubject
  ' Goal:   Get the constituent labeled as having grammatical role "Subject" in the preceding IP
  ' History:
  ' 09-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetPrecSubject(ByRef tblThis As DataTable, ByRef tblAnt As DataTable, _
                                ByRef ndxSubj As XmlNode) As Boolean
    Dim intI As Integer     ' Counter

    Try
      ' Validate
      If (tblThis Is Nothing) OrElse (tblAnt Is Nothing) Then Return False
      ' Try the antecedents in BACKWARDS direction
      For intI = tblAnt.Rows.Count - 1 To 0 Step -1
        ' Does this one have the correct grammatical role?
        If (tblAnt.Rows(intI).Item("GrRole") = "Subject") Then
          ' We found the subject, so return it
          ndxSubj = IdToNode(tblAnt.Rows(intI).Item("Id"))
          Return True
        End If
      Next intI
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetPrecSubject error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   Unresolved
  ' Goal:   Get the number of unresolved coreferences
  ' History:
  ' 03-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function Unresolved(ByRef tblThis As DataTable) As Integer
    Dim intI As Integer       ' Counter
    Dim intNum As Integer = 0 ' Number of unresolved ones

    Try
      ' Validate
      If (tblThis Is Nothing) Then Return 0
      ' Walk all the rows
      For intI = 0 To tblThis.Rows.Count - 1
        ' If this needs resolving, then increment the number
        If (NeedsResolving(tblThis(intI))) Then intNum += 1
      Next intI
      ' Return the total number of unresolved coreferences
      Return intNum
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/Unresolved error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return 0
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   NeedsResolving
  ' Goal:   Does this datarow show an element that needs resolving?
  ' Notes:  Elements that DON'T need resolving are:
  '         - NPtype = Indef
  '         - NPtype = Q    (quantifiers don't refer to something)
  '         - An NP with a coref relation that is already established
  ' History:
  ' 03-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function NeedsResolving(ByRef dtrThis As DataRow) As Boolean
    Try
      ' Validate
      If (dtrThis Is Nothing) Then Return False
      ' If this is an indefinite, then we don't need to resolve it
      Select Case dtrThis.Item("NPtype")
        Case "IndefNP", "QuantNP"
          ' There is no need to resolve this
          Return False
        Case "Trace"
          ' Traces don't need to get something
          Return False
        Case Else
          ' It needs resolving - see if this one is still unresolved
          If (dtrThis.Item("Coref") & "" = "") Then
            ' No coreference relation is specified, but it should be there
            Return True
          End If
      End Select
      ' Return no need by default
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/NeedsResolving error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   AllowedCorefType
  ' Goal:   Check whether this NP is allowed to serve as either source or target
  '           of a coreference relation
  ' History:
  ' 25-06-2010  ERK Created
  ' 12-07-2010  ERK Added allowing WH-traces
  ' ------------------------------------------------------------------------------------
  Private Function AllowedCorefType(ByRef ndxThis As XmlNode, ByRef strLabel As String) As Boolean
    Dim ndxChild As XmlNode   ' The current node's first <eLeaf> or <eTree> child

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Get the label of the string
      strLabel = ndxThis.Attributes("Label").Value
      ' =========== DEBUG =================
      'If (strLabel = "PRO+N") Then Stop
      ' ===================================
      ' First test: is it an allowed NP?
      If (DoLike(strLabel, strNPsourceTypes)) Then
        ' Second test: look at its first child
        ndxChild = ndxThis.SelectSingleNode("./eLeaf | ./eTree")
        ' Check this child
        If (ndxChild.Name = "eLeaf") Then
          ' It is an <eLeaf>, now check whether it is of type STAR
          If (ndxChild.Attributes("Type").Value = "Star") Then
            ' It is of type Star, now check if it is *con*
            If (ndxChild.Attributes("Text").Value = "*con*") OrElse (ndxChild.Attributes("Text").Value = "*pro*") Then
              ' This is okay
              Return True
            ElseIf (DoLike(ndxChild.Attributes("Text").Value, "[*]T[*]|[*]T[*]-*")) Then
              ' According to the YCOE documentation, for instance, these are WH-traces
              Return True
            End If
          Else
            ' No worries: 
            Return True
          End If
        Else
          ' This is not an <eLeaf>, but an <eTree>
          ' There is at least one <eTree> first child that may never serve as src/target of coref: expletive
          If (ndxChild.Attributes("Label").Value <> "EX") Then
            ' Return success
            Return True
          End If
        End If
      End If
      ' Inappropriate: return false
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AllowedCorefType error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetIpPosition
  ' Goal:   Derive the absolute IP position if which we are part
  ' History:
  '29-08-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetIpPosition(ByRef ndxThis As XmlNode) As String
    Dim ndxNext As XmlNode = Nothing  ' Working element
    Dim intI As Integer               ' Counter

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return ""
      ' Start out with myself
      ndxNext = ndxThis
      Do
        ' Get the IP I am part of
        ndxNext = GetIpAncestor(ndxNext)
        ' Any result
        If (ndxNext Is Nothing) Then Return ""
        ' Find this IP on the collection
        intI = colIPpos.Find(ndxNext.Attributes("Id").Value)
        ' Get node one higher
        ndxNext = ndxNext.ParentNode
      Loop Until (intI >= 0) OrElse (ndxNext Is Nothing)
      ' Result?
      If (intI < 0) Then Return ""
      ' Get the IP position
      Return colIPpos.Exmp(intI)
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetIpPosition error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetFirstWord
  ' Goal:   Get the first word in the constitiuent in [ndxThis]
  ' History:
  ' 08-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetFirstWord(ByRef ndxThis As XmlNode) As String
    Dim ndxWork As XmlNode  ' Workingnode

    Try
      ' Validate
      If (ndxThis.Name <> "eTree") Then Return ""
      ' Get first <eLeaf> that contains vernacular
      ndxWork = ndxThis.SelectSingleNode("./descendant::eLeaf[@Type='Vern'][1]")
      If (ndxWork Is Nothing) Then
        Return "-"
      Else
        Return ndxWork.Attributes("Text").Value
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetFirstWord error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddToStack
  ' Goal:   Add the NP [ndxThis] to the indicated stack
  ' History:
  ' 08-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function AddToStack(ByRef ndxThis As XmlNode, ByRef tblStack As DataTable) As Boolean
    Dim strLabel As String = ""   ' Label we are considering
    Dim strChLabel As String = "" ' The child's labe
    Dim strNPtype As String = ""  ' Type of NP
    Dim strGrRole As String = ""  ' Grammatical role
    Dim strPGN As String = ""     ' Person/Gender/Number for agreement
    Dim strPerson As String = ""  ' Person
    Dim strGender As String = ""  ' Gender
    Dim strNumber As String = ""  ' Number
    Dim strIPlabel As String = "" ' The label of the IP under which the NP finds itself
    Dim strIPances As String = "" ' The label of the nearest IP under which NP is
    Dim strCoref As String = ""   ' The kind of coreference (if it exists)
    Dim strVern As String = ""    ' The vernacular text of this constituent
    Dim strFirst As String = ""   ' First vernacular word in this line
    Dim strHdNoun As String = ""  ' The head noun (if existing) of this NP
    Dim strLocType As String = "" ' Kind of local relation to be made
    Dim strMsg As String = ""     ' Message text
    Dim strIPpos As String = ""   ' The absolute IP position of which we are part
    Dim strIsLocal As String = "" ' Whether needs to stay in one <forest>
    Dim intDstId As Integer       ' The ID of a possible destination (if it exists)
    Dim intNodeId As Integer      ' the ID of this node
    Dim intStackId As Integer     ' the ID of this stack element
    Dim intIpNum As Integer = -1  ' IPnum of this one
    Dim bIsSpeech As Boolean      ' Whether the IP is a speech IP
    Dim ndxParent As XmlNode      ' An appropriate parent
    Dim dtrThis As DataRow        ' One new element

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Check if it is already in here
      If (GetAntStack(ndxThis) >= 0) Then Return True
      ' Check whether this constituent belongs to the group of constituents that CAN serve as source stack
      If (AllowedCorefType(ndxThis, strLabel)) Then
        ' Get the ID of this element
        intNodeId = ndxThis.Attributes("Id").Value
        ' Determine the IP position
        strIPpos = GetIpPosition(ndxThis)
        ' =============== DEBUG =============
        'If (intNodeId = "829") Then Stop
        ' ===================================
        ' ======= DEBUGGING ========
        'If (InStr(NodeInfo(ndxThis), "French") > 0) Then Stop
        ' ==========================
        strVern = NodeText(ndxThis, True)
        ' Try get values for all relevant feature values
        strNPtype = GetFeature(ndxThis, "NP", "NPtype")
        strGrRole = GetFeature(ndxThis, "NP", "GrRole")
        strPGN = GetFeature(ndxThis, "NP", "PGN")
        ' Check if the PGN is not "unknown", and if we should ask the user or not...
        If (DoLike(strLabel, "PRO|PRO$|PRO^*")) AndAlso (strPGN = "unknown") AndAlso (bUserProPGN) _
           AndAlso (Not bIsFullyAuto) Then
          ' Ask user to adapt the PGN
          If (UserAdaptPGN(ndxThis, strPGN)) Then
            ' The adapted value is in [strPGN], but what about the tdlSettings?
            If (Not AddPronounPGN(strNPtype, strPGN, strVern)) Then
              ' Don't do anything in case of failure -- nothing goes wrong basically...
            End If
          End If
        End If
        ndxParent = GetParentNode(ndxThis, strIPtypes)
        If (ndxParent Is Nothing) Then
          strIPlabel = "IP"
          bIsSpeech = False
        Else
          strIPlabel = ndxParent.Attributes("Label").Value
          bIsSpeech = IsSpeechIp(ndxParent)
        End If
        ' Get ancestor IP
        ndxParent = ndxThis.SelectSingleNode("./ancestor::eTree[tb:matches(string(@Label), 'IP|IP-*')]", conTb)
        If (ndxParent Is Nothing) Then
          strIPances = "-"
        Else
          strIPances = ndxParent.Attributes("Label").Value
        End If
        ' Split person/gender/number
        If (Not SplitAgree(strPGN, strPerson, strGender, strNumber)) Then Logging("modAuto/AddToStack: bad PGN split") : Return False
        ' Get head noun
        strHdNoun = GetHeadNoun(ndxThis)
        ' If the head noun is empty, but this is a pronoun, try to retrieve it
        If (strHdNoun = "") Then
          ' Check if it has coref
          If (CorefDstId(ndxThis) >= 0) Then
            ' ============== DEBUG ==============
            'If (ndxThis.Attributes("Id").Value = 10768) Then Stop
            ' ===================================
            ' Yes, try to retrieve it
            strHdNoun = GetRootHead(ndxThis)
            '' Double check
            'If (strHdNoun = "") AndAlso (Not DoLike(strPGN, "3ns|1p|2")) AndAlso (strNPtype <> "Dem") Then
            '  Stop
            'End If
          ElseIf (DoLike(strNPtype, "Pro|PossPro|ZeroSbj")) AndAlso (Not DoLike(strPGN, "3ns|1p|2")) Then
            ' This element should have had a root...
            'Stop
          End If
        End If
        strCoref = CorefAttr(ndxThis, "RefType") & ""
        intDstId = CorefDstId(ndxThis)
        ' Get the IP number of this node
        If (Not GetIpNumber(ndxThis, intIpNum)) Then bInterrupt = True : Exit Function
        ' Get the ID for this stack element
        intStackId = tblStack.Rows.Count + 1
        ' Check if this element should stay local or not
        If (strNPtype = "Trace") AndAlso (Left(strVern, 3) = "*T*") Then
          strIsLocal = "True"
        ElseIf (strCoref = strRefNewVar) Then
          strIsLocal = "True"
        Else
          strIsLocal = "False"
        End If
        ' Make this new datarow
        dtrThis = tblStack.NewRow
        ' Fill this element
        With dtrThis
          .Item("StackElId") = intStackId : .Item("Id") = intNodeId
          .Item("Label") = strLabel : .Item("Vern") = strVern
          .Item("Head") = strHdNoun : .Item("Equal") = ""
          .Item("NPtype") = strNPtype : .Item("PGN") = strPGN
          .Item("GrRole") = strGrRole : .Item("Coref") = strCoref
          .Item("ChainNum") = 0 : .Item("IPdist") = 0
          .Item("PrefNum") = 0 : .Item("Ref") = intDstId
          .Item("IPlabel") = strIPlabel : .Item("Eval") = 0
          .Item("ChainIds") = "" : .Item("Loc") = GetLocation(ndxThis)
          .Item("IsLocal") = strIsLocal : .Item("Constraints") = "(none)"
          .Item("IPpos") = strIPpos : .Item("Level") = 0
          .Item("IsSpeech") = IIf(bIsSpeech, "True", "False")
          .Item("HasCoref") = IIf(HasCoref(ndxThis), "1", "0")
          .Item("IPances") = strIPances
          .Item("Person") = strPerson : .Item("Gender") = strGender : .Item("Number") = strNumber
          .Item("First") = GetFirstWord(ndxThis)
          .Item("IPnum") = intIpNum
        End With
        ' Add this element to the source stack, with IPdist set to "0"
        tblStack.Rows.Add(dtrThis)
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AddToStack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ResolveLocally
  ' Goal:   Resolve coreference for locally bound situations:
  '         (1) appositives, 
  '         (2) reflexives 
  '         (3) bare nouns
  '         (4) complements of finite BE clauses: Sbj is/was NP
  '         (5) possessive pronouns in a conjunction: [NP and his/her/their NP] --> Doen??
  '         (6) vocatives --> bind to the subject of the clause if possible
  ' History:
  ' 19-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function ResolveLocally(ByRef ndxThis As XmlNode) As Boolean
    Dim strCoref As String = ""         ' The kind of coreference (if it exists)
    Dim strLocType As String = ""       ' Kind of local relation to be made
    Dim strChLabel As String = ""       ' The child's labe
    Dim strMsg As String = ""           ' Message text
    Dim strBare As String = ""          ' Whether this is a BARE or NON-BARE complement
    Dim strRefType As String            ' Kind of reference
    Dim intThisId As Integer            ' This node's ID
    Dim ndxAnt As XmlNode               ' Antecedent we found
    Dim bIsInferred As Boolean = False  ' Whether control was pressed
    Dim bOkay As Boolean                ' Checkmark
    Dim bHadAnt As Boolean = False      ' Whether there already was an antecedent

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return True
      ' Try to resolve the coreferencing LOCALLY
      If (IsLocalType(ndxThis, strLocType)) Then
        ' Initialise the antecedent to the currently existing reference (if any)
        ndxAnt = CorefDst(ndxThis) : bHadAnt = (Not ndxAnt Is Nothing)
        ' WAS: ndxAnt = Nothing
        Select Case strLocType
          Case "freerel"    ' A free relative [what I don't know] --> must alway be NEW
            ' Add this one as discourse-new
            If (MakeUnlinkedNPnode(ndxThis, "AutoLocal", "Free relative", strRefNew)) Then
              ' Remove the source node from the stack
              RemoveFromStack(tblSrcStack, ndxThis)
            End If
          Case "lfd", "d-rel", "another"  ' Different options --> usually new
            ' If this is automatic mode, we cannot ask...
            If (bIsFullyAuto) Then Return True
            ' Set the message
            Select Case strLocType
              Case "lfd"
                strMsg = "Please indicate whether this left dislocated NP is New"
              Case "d-rel"
                strMsg = "Please indicate whether this D-relative is New"
              Case "another"
                strMsg = "Please select the contrastive antecedent for [another x]"
            End Select
            ' Ask user and suggest this to be "New"
            ndxAnt = Nothing
            If (Not AskUserDecision(ndxThis, ndxAnt, strMsg, bIsInferred)) Then
              ' There was an error, so return
              Return False
            End If
            ' Did the user select an antecedent after all?
            If (Not ndxAnt Is Nothing) Then
              ' Determine the reference type
              strRefType = IIf(bIsInferred, strRefInferred, strRefIdentity)
              ' There is a new antecedent -- make a connection
              If (Not AddCorefLink(ndxThis, ndxAnt, "User-Differ", strRefType, _
                                  "(Locally determined lfd/d-rel/another)", Nothing, "", _
                                   "Locally determined lfd/d-rel/another")) Then
                ' There is a problem
                Logging("modAuto/ResolveLocally: unable to connect locally from " & intThisId & _
                        " to " & ndxAnt.Attributes("Id").Value)
                ' Something went wrong
                Return False
              End If
              ' First add the source to the antecedent's collection...
              AddToStack(ndxThis, tblAntStack)
              ' Remove the source node from the source collection
              RemoveFromStack(tblSrcStack, ndxThis)
            End If
          Case "vocative"   ' Vocative should point to the subject of the clause it is in
            ' ============================================================================================
            ' No, we cannot deal with vocatives in a local manner...
            ' ============================================================================================
            '  ' Get the ID of ourselves
            '  intThisId = ndxThis.Attributes("Id").Value
            '  ' Determine the antecedent of this possessive pronoun, and return it in [ndxAnt]
            '  If (Not GetVocAnt(ndxThis, ndxAnt)) Then
            '    ' Okay, it is quite possible that we cannot find an antecedent for a PossPro!
            '    ' Leave the SELECT statement (but continue in the [intI] loop)
            '    Exit Select
            '  End If
            '  ' There is an antecedent -- is this connection already there?
            '  If (CorefDstId(ndxThis) = ndxAnt.Attributes("Id").Value) Then
            '    ' No need to redo everything!
            '  ElseIf (HasCoref(ndxThis)) Then
            '    ' Really no need to improve upon ourselves!
            '  Else
            '    ' There is a new antecedent -- make a connection
            '    If (Not AddCorefLink(ndxThis, ndxAnt, "AutoLocal", "Identity", "(Locally determined vocative)", Nothing, "", _
            '                         "Locally determined vocative")) Then
            '      ' There is a problem
            '      Logging("modAuto/ResolveLocally: unable to connect locally from " & intThisId & _
            '              " to " & ndxAnt.Attributes("Id").Value)
            '      ' Something went wrong
            '      Return False
            '    End If
            '  End If
            '  ' (N.B: should NOT add the vocative to the antecedent's collection!!)
            '  '' First add the source to the antecedent's collection...
            '  'AddToStack(ndxThis, tblAntStack)
            '  ' Remove the source node from the source collection
            '  RemoveFromStack(tblSrcStack, ndxThis)
          Case "posspro"    ' Some possessive pronouns have a syntactically determined antecedent
            ' Get the ID of ourselves
            intThisId = ndxThis.Attributes("Id").Value
            ' Determine the antecedent of this possessive pronoun, and return it in [ndxAnt]
            If (Not GetPossProAnt(ndxThis, ndxAnt)) Then
              ' Okay, it is quite possible that we cannot find an antecedent for a PossPro!
              ' Leave the SELECT statement (but continue in the [intI] loop)
              Exit Select
            End If
            ' Now we should check agreement in PGN!!!
            If (Not AgreePGN(ndxThis, ndxAnt)) Then
              ' There is no agreement in either person, gender or number
              Exit Select
            End If
            ' We should also check disjointness - they are suspicious otherwise
            If (IsDisjoint(ndxThis, ndxAnt)) Then
              ' They are disjoint, so that is suspicious
              Exit Select
            End If
            ' There is an antecedent -- is this connection already there?
            If (CorefDstId(ndxThis) = ndxAnt.Attributes("Id").Value) Then
              ' No need to redo everything!
            ElseIf (HasCoref(ndxThis)) Then
              ' Really no need to improve upon ourselves!
            Else
              '' =========== DEBUGGING ===============
              'Debug.Print("Make link from " & NodeInfo(ndxThis) & " to " & NodeInfo(ndxAnt))
              '' =====================================
              ' There is a new antecedent -- make a connection
              If (Not AddCorefLink(ndxThis, ndxAnt, "AutoLocal", "Identity", "(Locally determined reflexive)", Nothing, "", _
                                   "Locally determined reflexive")) Then
                ' There is a problem
                Logging("modAuto/ResolveLocally: unable to connect locally from " & intThisId & _
                        " to " & ndxAnt.Attributes("Id").Value)
                ' Something went wrong
                Return False
              End If
            End If
            ' First add the source to the antecedent's collection...
            AddToStack(ndxThis, tblAntStack)
            ' Remove the source node from the source collection
            RemoveFromStack(tblSrcStack, ndxThis)
          Case "appositive" ' An appositive should be linked to the nearest higher NP
            loc_bLocalApp = True
            ' Get the nearest preceding or higher NP of type "NP-*"
            If (Not GetAppAntecedent(ndxThis, ndxAnt)) And (Not bHadAnt) Then
              ' When we are in fully automatic, we are not allowed to ask!
              If (bIsFullyAuto) Then Return True
              ' No antecedent has been found, so check with the user
              If (Not AskUserDecision(ndxThis, ndxAnt, _
                      "Please select the antecedent for the appositive NP", bIsInferred)) Then
                ' There was an error, so return
                Return False
              End If
              ' Double check whether an antecedent has now actually been provided or not
              If (ndxAnt Is Nothing) Then
                ' No antecedent is provided, so do what??
                ' Just exit the procedure nicely...
                Return True
              End If
            End If
            ' There is an antecedent -- is this connection already there?
            If (CorefDstId(ndxThis) = ndxAnt.Attributes("Id").Value) Then
              ' No need to redo everything!
            ElseIf (Not AgreeGN(ndxThis, ndxAnt)) Then
              ' No agreement, so exit nicely and let user provide reference
              Return True
            Else
              ' There is a new antecedent -- make a connection
              If (Not AddCorefLink(ndxThis, ndxAnt, "AutoLocal", "Identity", "(Locally determined appositive)", Nothing, "", _
                                   "Locally determined apposition")) Then
                ' There is a problem
                Logging("modAuto/ResolveLocally: unable to make apposition connection from " & intThisId & _
                        " to " & ndxAnt.Attributes("Id").Value)
                ' Something went wrong
                Return False
              End If
            End If
            ' Remove the source node from the stack
            RemoveFromStack(tblSrcStack, ndxThis)
          Case "reflexive"  ' A reflexive should be linked to the nearest subject
            loc_bLocalRefl = True
            ' Get the ID of ourselves
            intThisId = ndxThis.Attributes("Id").Value
            ' Determine the subject of this IP, and return it in [ndxAnt]
            If (Not GetIpSubjectNode(ndxThis, ndxAnt)) Then
              ' Note there is a problem...
              Logging("modAuto/ResolveLocally: unable to determine the subject of " & intThisId)
              ' Leave the SELECT statement (but continue in the [intI] loop)
              Exit Select
            End If
            ' There is an antecedent -- is this connection already there?
            If (CorefDstId(ndxThis) = ndxAnt.Attributes("Id").Value) Then
              ' No need to redo everything!
            Else
              ' There is a new antecedent -- make a connection
              If (Not AddCorefLink(ndxThis, ndxAnt, "AutoLocal", "Identity", "(Locally determined reflexive)", Nothing, "", _
                                   "Locally determined reflexive")) Then
                ' There is a problem
                Logging("modAuto/ResolveLocally: unable to connect locally from " & intThisId & _
                        " to " & ndxAnt.Attributes("Id").Value)
                ' Something went wrong
                Return False
              End If
            End If
            ' Remove the source node from the stack
            RemoveFromStack(tblSrcStack, ndxThis)
          Case "barecompl", "complement"  ' The NP does not refer, and cannot be referred to
            loc_bLocalBare = True
            strBare = IIf(strLocType = "barecompl", "bare", "non-bare")
            ' Start out without success...
            bOkay = False
            ' Check if this already had a link
            strCoref = CorefAttr(ndxThis, "RefType")
            If (strCoref = "") Then
              ' Make the link
              If (MakeUnlinkedNPnode(ndxThis, "AutoLocal", "Locally determined " & strBare & _
                                 " NP complement", "Inert")) Then
                ' Remove the source node from the stack
                RemoveFromStack(tblSrcStack, ndxThis)
                ' Indicate success
                bOkay = True
              End If
            ElseIf (strCoref <> "Inert") Then
              strMsg = "Locally determined " & strBare & " NP complement (replaces [" & strCoref & "])"
              ' Only make a link if it does not exist already!!
              If (MakeUnlinkedNPnode(ndxThis, "AutoLocal", strMsg, "Inert")) Then
                ' Remove the source node from the stack
                RemoveFromStack(tblSrcStack, ndxThis)
                ' Indicate success
                bOkay = True
              Else
                ' We were unable to make this "Inert" link, but that is okay...
                bOkay = True
              End If
            Else
              ' There already is an Inert relation -- Remove the source node from the stack
              RemoveFromStack(tblSrcStack, ndxThis)
              ' Indicate success
              bOkay = True
            End If
            ' Did we succeed?
            If (Not bOkay) Then
              ' There is a problem
              Logging("modAuto/ResolveLocally: unable to deal with " & strBare & " complement Id=" & intThisId)
              ' Something went wrong
              MsgBox("modAuto/ResolveLocally: cannot get a source stack reference to a constituent")
              Return False
            End If
        End Select
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/ResolveLocally error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '' ------------------------------------------------------------------------------------
  '' Name:   TraverseIpTree
  '' Goal:   Traverse a tree with a particular action in mind
  ''         SrcStack - add elements to the source stack if they should go there
  ''         Local    - Local pronoun resolution attempt
  ''         Syntax   - produce HTML output of the indicated node
  ''                    (the ID of the selected node is in loc_intSelNodeId)
  '' History:
  '' 03-06-2010  ERK Created
  '' ------------------------------------------------------------------------------------
  'Private Function TraverseIpTree(ByRef ndxThis As XmlNode, ByVal strAction As String, _
  '                                ByRef strBack As String) As Boolean
  '  Dim strChLabel As String = "" ' The child's labe
  '  Dim strMsg As String = ""     ' Message text
  '  Dim ndxList As XmlNodeList    ' List of child nodes
  '  Dim ndxChild As XmlNode       ' The current node
  '  Dim intI As Integer           ' Counter

  '  Try
  '    ' Action BEFORE going down
  '    Select Case strAction
  '      Case "SrcStack" ' Add elements to source stack if they should go there
  '        ' Check if this NP is not already in the Antecedent collection
  '        If (Not IsInAntStack(ndxThis)) Then
  '          ' Copy to the collection of source NPs
  '          If (Not AddToStack(ndxThis, tblSrcStack)) Then Return False
  '        End If
  '    End Select
  '    ' Visit the children
  '    ndxList = ndxThis.SelectNodes("./eTree")
  '    For intI = 0 To ndxList.Count - 1
  '      ' Get the current node
  '      ndxChild = ndxList(intI)
  '      ' Get this label
  '      strChLabel = ndxChild.Attributes("Label").Value
  '      ' Children nodes should NOT start with "IP"
  '      If (Not DoLike(strChLabel, strIPtypes)) Then
  '        ' Perform the action on this child recursively
  '        If (Not TraverseIpTree(ndxChild, strAction, strBack)) Then
  '          ' There is an error, so return
  '          Return False
  '        End If
  '      End If
  '    Next intI
  '    ' Action AFTER going down
  '    Select Case strAction
  '      Case "Local"    ' Try to resolve local pronoun referencing, including "Inert" situations
  '        If (Not ResolveLocally(ndxThis)) Then Return False
  '    End Select
  '    ' (nothing right now)
  '    ' Return success
  '    Return True
  '  Catch ex As Exception
  '    ' Show error
  '    HandleErr("modAuto/TraverseIpTree error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
  '    ' Return failure
  '    Return False
  '  End Try
  'End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsSelectedNode
  ' Goal:   Check if node [ndxThis] equals the selected node
  ' History:
  ' 28-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function IsSelectedNode(ByRef ndxThis As XmlNode) As Boolean
    Try
      ' Validate
      If (ndxThis Is Nothing) OrElse (loc_intSelNodeId < 0) Then Return False
      ' Is this an <eTree>?
      If (ndxThis.Name = "eTree") Then
        ' Do the check
        Return (ndxThis.Attributes("Id").Value = loc_intSelNodeId)
      Else
        Return False
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/IsSelectedNode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsInSelectedNode
  ' Goal:   Check if node [ndxThis] is part of the selected node
  ' History:
  ' 28-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function IsInSelectedNode(ByRef ndxThis As XmlNode) As Boolean
    Dim ndxNext As XmlNode = Nothing  ' parent nodes

    Try
      ' Validate
      If (ndxThis Is Nothing) OrElse (loc_intSelNodeId < 0) Then Return False
      ' Walk until topmost parent
      ndxNext = ndxThis
      While (Not ndxNext Is Nothing)
        ' Verify
        If (ndxNext.Name = "eTree") Then
          If (ndxNext.Attributes("Id").Value = loc_intSelNodeId) Then Return True
        End If
        ' Try next one
        ndxNext = ndxNext.ParentNode
      End While
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/IsInSelectedNode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   TraverseNodes
  ' Goal:   Traverse a tree with a particular action in mind. Visit ALL (!!!) nodes
  '         Syntax   - produce HTML output of the indicated node
  '                    (the ID of the selected node is in loc_intSelNodeId)
  '                    N.B: all <CODE> nodes are skipped
  '         NewAndLocal  - try to get all discourse NEW elements as well as resolving local
  '                        coreference situations
  ' History:
  ' 28-06-2010  ERK Created
  ' 29-08-2010  ERK Added processing of discourse-new elements
  ' ------------------------------------------------------------------------------------
  Private Function TraverseNodes(ByRef ndxThis As XmlNode, ByVal strAction As String, _
                                  ByRef strBack As String) As Boolean
    Dim strLabel As String = ""   ' Label we are considering
    Dim strVern As String = ""    ' Vernacular text
    Dim ndxList As XmlNodeList    ' List of child nodes
    Dim ndxChild As XmlNode       ' The current node
    Dim intI As Integer           ' Counter
    Dim bIsSelNode As Boolean     ' Whether this is the selected node
    Dim bTextSelNode As Boolean   ' Whether this text belongs to the selected node
    Dim bCondition As Boolean     ' The condition to go recursively down
    Dim bZero As Boolean = False  ' To flag bad "zero" leafs

    Try
      ' Do we have a selected node?
      If (loc_intSelNodeId < 0) Then
        ' We can never be inside the selected node
        bIsSelNode = False : bTextSelNode = False
      Else
        ' Is this part of the selected node?
        bIsSelNode = IsSelectedNode(ndxThis)
        bTextSelNode = IsInSelectedNode(ndxThis)
        ' The following code was used, but is too slow
        ''bIsSelNode = (Not ndxThis.SelectSingleNode("./self::eTree[@Id=" & loc_intSelNodeId & "]") Is Nothing)
        ''bTextSelNode = (Not ndxThis.SelectSingleNode("./ancestor-or-self::eTree[@Id=" & loc_intSelNodeId & "]") Is Nothing)
      End If
      ' Action BEFORE going down
      Select Case strAction
        Case "Syntax"
          Select Case ndxThis.Name
            Case "eLeaf"
              ' What type of <eLeaf> is this?
              bZero = ((ndxThis.Attributes("Type").Value = "Star") AndAlso _
                       (ndxThis.Attributes("Text").Value <> "*con*")) OrElse _
                      (ndxThis.Attributes("Type").Value = "Zero")
              ' Only output the correct leaves
              If (Not bZero) Then
                ' Get the vernacular text
                strVern = VernToEnglish(ndxThis.Attributes("Text").Value)
                ' Is this the selected Id?
                If (bTextSelNode) Then
                  ' The text of the selected Id is set in red.
                  strBack &= "<font color='red' size='2'><b>" & strVern & "</b><font>" & vbCrLf
                Else
                  ' Just add the text of this node in black
                  strBack &= "<font color='black' size='2'>" & strVern & "<font>" & vbCrLf
                End If
              End If
            Case "eTree"
              ' Only do something if this is the selected node
              If (bIsSelNode) Then
                ' Action depends on what kind of label we have
                strLabel = ndxThis.Attributes("Label").Value
                If (strLabel <> "CODE") Then
                  ' First give the bracket
                  strBack &= "<font color='black' size='2'>[<font>" & vbCrLf
                  ' Give the label of this node followed by a space
                  strBack &= "<font color='green' size='1'>" & ndxThis.Attributes("Label").Value & "<font> " & vbCrLf
                End If
              End If
            Case "forest"
              ' Give the forest's location in brackets
              strBack &= "<font color='black' size='2'>[" & ndxThis.Attributes("Location").Value & "] "
          End Select
        Case "NewAndLocal"
          ' Is this one of the nodes we are looking for?
          If (ndxThis.Name = "eTree") Then
            ' DEBUGGING =======
            'If (ndxThis.Attributes("Id").Value = 2992) Then Stop
            ' ===================
            ' Is this node a noun phrase?
            If (DoLike(ndxThis.Attributes("Label").Value, strNPsourceTypes)) Then
              ' Check for discourse-newness
              If (IsDiscourseNew(ndxThis)) Then
                ' It is discourse-new, so put it in the antecedent's collection
                If (Not AddToStack(ndxThis, tblAntStack)) Then Return False
                ' Make sure it is marked as discourse new
                If (Not MakeUnlinkedNPnode(ndxThis, "AutoLocal", "Discourse-New")) Then Return False
              End If
              '' Try to resolve any other local coreferencing
              'If (Not ResolveLocally(ndxThis)) Then Return False
            End If
          End If
      End Select
      ' Visit the children
      ndxList = ndxThis.ChildNodes
      For intI = 0 To ndxList.Count - 1
        ' Get the current node
        ndxChild = ndxList(intI)
        ' Always go down - initialy
        bCondition = True
        ' Going recursively down depends on what action we are performing
        Select Case strAction
          Case "Syntax"
            ' Only go down if this is not a particular kind
            If (ndxThis.Name = "eTree") Then
              ' Action depends on what kind of label we have
              strLabel = ndxThis.Attributes("Label").Value
              If (strLabel = "CODE") Then
                bCondition = False
              End If
            End If
        End Select
        ' Is the condition met?
        If (bCondition) Then
          ' Perform the action on this child recursively
          If (Not TraverseNodes(ndxChild, strAction, strBack)) Then
            ' There is an error, so return
            Return False
          End If
        End If
      Next intI
      ' Action AFTER going down
      Select Case strAction
        Case "Syntax"
          Select Case ndxThis.Name
            Case "eTree"
              ' Only do something if this is the selected node
              If (bIsSelNode) Then
                ' Action depends on what kind of label we have
                strLabel = ndxThis.Attributes("Label").Value
                If (strLabel <> "CODE") Then
                  ' Give a closing bracket
                  strBack &= "<font color='black' size='2'>]<font>" & vbCrLf
                End If
              End If
            Case "forest"
              ' Finish with a linebreak
              strBack &= vbCrLf
          End Select
      End Select
      ' (nothing right now)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/TraverseNodes error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DoChunk
  ' Goal:   Process one chunk of NPs, as have been put in the source collection
  ' History:
  ' 01-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function DoChunk() As Boolean
    Try
      ' Validate: is there something in the source collection?
      If (tblSrcStack.Rows.Count = 0) Then Return True
      ' Check interrupt
      If (bInterrupt) Then Return False
      ' Process this chunk
      CoreferenceSrcColl()
      ' ================ DEBUGGING =============
      ' If (tblSrcStack.Rows.Count > 0) Then Stop
      ' ========================================
      ' Make sure the source collection is cleared
      ClearTable(tblSrcStack)
      ' Check interrupt
      If (bInterrupt) Then Return False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/DoChunk error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ChunkWalk
  ' Goal:   Traverse a tree in IP chunks for coreference processing
  ' History:
  ' 01-09-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function ChunkWalk(ByRef ndxThis As XmlNode) As Boolean
    Dim strLabel As String = ""   ' Label we are considering
    Dim strVern As String = ""    ' Vernacular text
    Dim ndxList As XmlNodeList    ' List of child nodes
    Dim ndxChild As XmlNode       ' The current node
    Dim intI As Integer           ' Counter
    Dim bIP As Boolean = False    ' This flags that we have started on an IP

    Try
      ' Action BEFORE going down
      If (ndxThis.Name = "eTree") Then
        ' ========= DEBUGGING =============
        'If (ndxThis.Attributes("Id").Value = 1346) Then Stop
        ' =================================
        ' Get the label of this node
        strLabel = ndxThis.Attributes("Label").Value
        ' Is this one of the approved NP or IP types?
        If (DoLike(strLabel, strNPsourceTypes)) AndAlso (Not DoLike(strLabel, strNPnoSourceTypes)) Then
          ' Do keep track of all the source NPs
          loc_dblNPcount += 1
          ' ============== Debugging ================
          ' Indicate that we have treated it
          ' =========================================
          loc_colNPstack.Add(ndxThis.Attributes("Id").Value)
          ' Add it to the source stack
          If (Not AddToStack(ndxThis, tblSrcStack)) Then Return False
          ' Don't check local coreferencing if this already was done
          If (Not HasCoref(ndxThis)) Then
            ' Try to resolve any other local coreferencing
            If (Not ResolveLocally(ndxThis)) Then Return False
          End If
          ' Check if this NP is the root-node of this forest
          If (ndxThis.ParentNode.Name = "forest") Then
            ' This counts as the beginning of an IP, so first process the source stack as we now have it
            If (Not DoChunk()) Then Return False
            ' Also mark that we have are now starting to process an IP on this level
            bIP = True
          End If
        ElseIf (DoLike(strLabel, strIPtypes)) OrElse (ndxThis.ParentNode.Name = "forest") Then
          ' ================== DEBUGGING ===================
          'If (ndxThis.Attributes("Id").Value = 16796) Then Stop
          ' ================================================
          ' It is the beginning of an IP, so first process the source stack as we now have it
          If (Not DoChunk()) Then Return False
          ' Also mark that we have are now starting to process an IP on this level
          bIP = True
        Else
          ' This node is being skipped; is that alright? --> Keep track of it
          ' TODO: add code to keep track of which labels + those in clauses under them are being skipped
        End If
      End If
      ' Visit the children
      ndxList = ndxThis.ChildNodes
      For intI = 0 To ndxList.Count - 1
        ' Get the current node
        ndxChild = ndxList(intI)
        ' Only continue with <eTree> children!
        If (ndxChild.Name = "eTree") Then
          ' Perform the action on this child recursively
          If (Not ChunkWalk(ndxChild)) Then
            ' There is an error, so return
            Return False
          End If
          ' ========== DEBUG ===================
          ' Debug.Print("source stack size=" & tblSrcStack.Rows.Count)
          ' ====================================
        End If
      Next intI
      'Debug.Print(ndxList.Count)
      ' Action AFTER going down
      If (bIP) Then
        ' ================== DEBUGGING ===================
        'If (ndxThis.Attributes("Id").Value = 16796) Then Stop
        'If (tblSrcStack.Rows.Count > 0) Then
        '  ' Show how many items there are in the source collection
        '  Debug.Print("ChunkWalk source stack before second DoChunk = " & tblSrcStack.Rows.Count)
        'End If
        ' ================================================
        ' We have FINISHED processing one IP, so now we need to process the remaining chunk (if any)
        If (Not DoChunk()) Then Return False
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/ChunkWalk error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsDiscourseNew
  ' Goal:   Check if this node is an NP that is discourse new
  ' History:
  ' 31-08-2010  ERK Created
  ' 19-03-2012  ERK Added "WNP*" as a category that is always discourse-new
  ' ------------------------------------------------------------------------------------
  Private Function IsDiscourseNew(ByRef ndxThis As XmlNode) As Boolean
    Dim strNPtype As String = ""  ' The kind of NP we are dealing with
    Dim ndxNext As XmlNode        ' Working node

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      If (ndxThis.Name <> "eTree") Then Return False
      ' Make sure the label starts with 'NP'
      If (ndxThis.Attributes("Label").Value Like "NP*") Then
        ' Check whether this NP already has a coreference type
        If (GetFeature(ndxThis, "coref", "RefType") <> "") Then Return False
        ' Ascertain the type of NP this is (which may already be enough)
        strNPtype = GetFeature(ndxThis, "NP", "NPtype")
        Select Case strNPtype
          Case "DefNP"
            ' If this is a definite NP, but there is a possible antecedent with the same
            '   head noun, then we may *not* have a new one. So: check the head noun

            ' Consider the following situations:
            ' (1) the + (adj) + N + PP/CP (restrictive postmodification)
            ' (2) Possessive + (adj) + N  (anchored NP)
            ' [DELETED: (4) Sbj + is/was + NP[cmpl] (complement of a copula construction)
            '     --> This is now labelled as "Inert"]
            ' FIRST step: find out if there is a definite determiner part of the NP
            ndxNext = FirstChild(ndxThis, "D|D^*")
            ' Any results?
            If (ndxNext Is Nothing) Then
              ' No determiner -- try situation (2)
              ndxNext = FirstChild(ndxThis, "NP*POS|NP*GEN|PRO$|NPR$|NPRS$")
              ' Any results?
              If (Not ndxNext Is Nothing) Then
                ' Go to the next one
                ndxNext = ndxNext.NextSibling
                If (Not ndxNext Is Nothing) Then
                  ' Possibly skip an adjective
                  If (DoLike(ndxNext.Attributes("Label").Value, "ADJ*|QS")) Then
                    ' Go to the next one
                    ndxNext = ndxNext.NextSibling : If (ndxNext Is Nothing) Then Return False
                  End If
                  ' If (2) holds, then there should be a Noun coming now
                  If (DoLike(ndxNext.Attributes("Label").Value, "N|NS")) Then Return True
                End If
              End If
            Else
              ' Yes, a determiner. Look for the next sibling
              ndxNext = ndxNext.NextSibling
              If (Not ndxNext Is Nothing) Then
                ' Possibly skip an adjective
                If (DoLike(ndxNext.Attributes("Label").Value, "ADJ*|QS")) Then
                  ' Go to the next one
                  ndxNext = ndxNext.NextSibling : If (ndxNext Is Nothing) Then Return False
                End If
                ' Check out this label
                Select Case ndxNext.Attributes("Label").Value
                  Case "N", "NS"
                    ' Okay, this could be (1): postmodification...
                    ndxNext = ndxNext.NextSibling
                    If (Not ndxNext Is Nothing) Then
                      ' If (1) holds, this should be a PP
                      If (DoLike(ndxNext.Attributes("Label").Value, "PP*|CP*")) Then Return True
                    End If
                  Case "NPR", "NPRS"
                    ' Okay, this could be (3): proper noun premodification
                    ndxNext = ndxNext.NextSibling
                    If (Not ndxNext Is Nothing) Then
                      ' If (3) holds, then there should be a Noun coming now
                      If (DoLike(ndxNext.Attributes("Label").Value, "N|NS")) Then Return True
                    End If
                End Select
              End If
            End If
            ' Check if situation (4) holds
            ' NOTE: situation (4) is already checked in ResolveLocally()
            '' A: we should NOT be the subject NP
            'If (GetFeature(ndxThis, "NP", "GrRole") <> "Subject") Then
            '  ' I should have a finite BE sibling
            '  If (Not FirstChild(ndxThis.ParentNode, strFiniteBe) Is Nothing) Then
            '    ' We can assume that there is a subject somewhere
            '    Return True
            '  End If
            'End If
          Case "IndefNP"
            ' Indefinite NPs are always regarded as discourse new
            Return True
          Case "QuantNP"
            ' Quantificational NPs? I think they are always regarded as discourse new, but am not sure...
            Return True
        End Select
      ElseIf (ndxThis.Attributes("Label").Value Like "WNP*") Then
        ' Question NPs are always regarded as discourse new
        Return True
      End If
      ' We haven't made it...
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/IsDiscourseNew error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   FirstChild
  ' Goal:   Get the first child with the indicated label
  ' History:
  ' 31-08-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function FirstChild(ByRef ndxThis As XmlNode, ByVal strPattern As String) As XmlNode
    Dim ndxNext As XmlNode        ' Working node

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return Nothing
      If (ndxThis.Name <> "eTree") Then Return Nothing
      ' Look for the first child
      ndxNext = ndxThis.FirstChild
      ' Look for the correct sibling
      While (Not ndxNext Is Nothing)
        ' Check this one
        If (ndxNext.Name = "eTree") Then
          ' Check the label
          If (DoLike(ndxNext.Attributes("Label").Value, strPattern)) Then
            ' Got him!
            Return ndxNext
          End If
        End If
        ' Go to the next one
        ndxNext = ndxNext.NextSibling
      End While
      ' Return failure
      Return Nothing
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/FirstChild error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DoLike
  ' Goal:   Perform the "Like" function using the pattern (or patterns) stored in [strPattern]
  '         There can be more than 1 pattern in [strPattern], which must be separated
  '         by a vertical bar: |
  ' History:
  ' 17-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function DoLike(ByVal strText As String, ByVal strPattern As String) As Boolean
    Dim arPattern() As String   ' Array of patterns
    Dim intI As Integer         ' Counter

    Try
      ' Reduce the [strPattern]
      strPattern = Trim(strPattern)
      ' SPlit the [strPattern] into different ones
      arPattern = Split(strPattern, "|")
      ' Perform the "Like" operation for all needed patterns
      For intI = 0 To UBound(arPattern)
        ' See if something positive comes out of this comparison
        If (strText Like arPattern(intI)) Then
          ' Return success
          Return True
        End If
      Next intI
      ' No match has happened, so return false
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/DoLike error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   RemoveFromStack
  ' Goal:   Remove the indicated node from the indicated stack (if there)
  ' History:
  ' 15-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub RemoveFromStack(ByRef tblStack As DataTable, ByRef ndxThis As XmlNode)
    Dim intI As Integer       ' Counter
    Dim intThisId As Integer  ' This node's ID
    Dim dtrThis As DataRow    ' Current row

    Try
      ' Validate
      If (tblStack Is Nothing) OrElse (ndxThis Is Nothing) Then Exit Sub
      ' Get this node's ID
      intThisId = ndxThis.Attributes("Id").Value
      ' ============== DEBUGGING =============
      ' If (intThisId = 4615) OrElse (intThisId = 4617) Then Stop
      ' ======================================
      ' Walk the stack
      For intI = 0 To tblStack.Rows.Count - 1
        ' Consider this element
        dtrThis = tblStack.Rows(intI)
        ' Is this the correct row?
        If (dtrThis.Item("Id") = intThisId) Then
          ' This is the one - delete it
          dtrThis.Delete()
          ' And exit
          Exit Sub
        End If
      Next intI
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/RemoveFromStack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   AskUserDecision
  ' Goal:   Ask the user to supply an antecedent to [ndxSrc].
  '           The caller of this subroutine may propose a best match in [ndxDst]
  '           If we get an antecedent, return it in [ndxDst]
  ' History:
  ' 15-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function AskUserDecision(ByRef ndxSrc As XmlNode, ByRef ndxDst As XmlNode, ByVal strQuestion As String, _
                                  ByRef bIsInferred As Boolean) As Boolean
    Dim intDstId As Integer           ' The ID of the target
    Dim intReply As DialogResult      ' Result of replying
    Dim ndxAnt As XmlNode = Nothing   ' The possible antecedent
    Dim bError As Boolean = False     ' An error occurred
    Dim strLinkType As String         ' The kind of link

    Try
      ' Validate
      If (ndxSrc Is Nothing) Then Logging("modAuto/AskUserDecision: empty source node") : Return False
      ' See if there is a proposal by the caller
      If (Not ndxDst Is Nothing) Then
        ' Select this as potential target
        SelectNode(ndxDst, "target")
        ' Keep the potential antecedent
        ndxAnt = ndxDst
      End If
      ' Select this constituent in the window
      SelectNode(ndxSrc, "source")
      ' Fill in the ask panel's matters
      intReply = frmMain.DoAskListView(strQuestion, ndxSrc, ndxAnt, tblSrcStack, tblAntStack, intDstId, bIsInferred)
      Select Case intReply
        Case DialogResult.Yes, DialogResult.OK
          ' The target is in [intDstId], but is it valid?
          Select Case intDstId
            Case -1   ' There's been an error
              bError = True
              Logging("modAuto/AskUserDecision: bad destination Id")
            Case -2   ' Establish a NEW element
              ' Determine the linktype
              If (ndxAnt Is Nothing) Then strLinkType = "User-Agree" Else strLinkType = "User-Differ"
              ' Make the relation
              If (MakeUnlinkedNP(tblSrcStack(GetSrcStack(ndxSrc)), strLinkType, "", strRefNew)) Then
                ' Wipe out the potential antecedent
                ndxDst = Nothing
              Else
                ' There is a problem
                Logging("modAuto/AskUserDecision: unable to mark " & ndxSrc.Attributes("Id").Value & _
                        " as New")
                ' We can still try to find the next one
              End If
            Case -3   ' Establish an ASSUMED relation
              ' Determine the linktype
              If (ndxAnt Is Nothing) Then strLinkType = "User-Agree" Else strLinkType = "User-Differ"
              ' Make the relation
              If (MakeUnlinkedNP(tblSrcStack(GetSrcStack(ndxSrc)), strLinkType, "", strRefAssumed)) Then
                ' Wipe out the potential antecedent
                ndxDst = Nothing
              Else
                ' There is a problem
                Logging("modAuto/AskUserDecision: unable to mark " & ndxSrc.Attributes("Id").Value & _
                        " as Assumed")
                ' We can still try to find the next one
              End If
            Case -4   ' Skip this one
              ' Make sure no potential antecedent is returned
              ndxDst = Nothing
            Case -5   ' Establish an INERT relation
              ' Determine the linktype
              If (ndxAnt Is Nothing) Then strLinkType = "User-Agree" Else strLinkType = "User-Differ"
              ' Make the relation
              If (MakeUnlinkedNP(tblSrcStack(GetSrcStack(ndxSrc)), strLinkType, "", strRefInert)) Then
                ' Wipe out the potential antecedent
                ndxDst = Nothing
              Else
                ' There is a problem
                Logging("modAuto/AskUserDecision: unable to mark " & ndxSrc.Attributes("Id").Value & _
                        " as Inert")
                ' We can still try to find the next one
              End If
            Case Else
              ' Find the node that belongs to this one
              ndxDst = IdToNode(intDstId)
              ' This node will be returned by reference, and acted upon
          End Select
        Case Else
          ' Indicate that there is an error
          bError = True
          Logging("modAuto/AskUserDecision: user pressed cancel or abort (reply=" & intReply.ToString & ")")
      End Select
      ' Do make the nodes blank again
      If (Not ndxAnt Is Nothing) Then SelectNode(ndxAnt, "blank")
      SelectNode(ndxSrc, "blank")
      ' Return success or failure, depending on the error state
      Return (Not bError)
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AskUserDecision error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetAppAntecedent
  ' Goal:   Find the appositive antecedent:
  '         - If the PRN ends with a number (e.g: NP-PRN-1), then find preceding *ICH*-1
  '         - Try nearest NP-parent
  '         - Try nearest preceding NP-sibling
  '         - Otherwise we cannot find it
  ' Note:   There should be agreement in gender/number (not person)
  ' History:
  ' 15-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function GetAppAntecedent(ByRef ndxSrc As XmlNode, ByRef ndxAnt As XmlNode) As Boolean
    Dim ndxNext As XmlNode = Nothing  ' Node to consider
    Dim strLabel As String            ' Label of me
    Dim strNum As String              ' Number
    Dim intPos As Integer             ' Position within string

    Try
      ' Validate
      If (ndxSrc Is Nothing) Then Return False
      ' Get my label
      strLabel = ndxSrc.Attributes("Label").Value
      ' Do we end with a number?
      intPos = InStrRev(strLabel, "-")
      If (intPos > 0) Then
        ' Check for number
        strNum = Mid(strLabel, intPos + 1)
        If (IsNumeric(strNum)) Then
          ' Yes, we have a NP-PRN-n item -- Look for preceding *CH*-n
          If (GetForestNode(ndxSrc, ndxNext)) Then
            ' We have a forest node -- look downward
            ndxAnt = ndxNext.SelectSingleNode("./descendant::eLeaf[tb:Like(string(@Text),'?ICH?-" & _
                                              strNum & "')]", conTb)
            If (Not ndxAnt Is Nothing) Then
              ' We need to have the parent of this [eLeaf]
              ndxAnt = ndxAnt.ParentNode
              ' We got the antecedent
              Return True
            End If
          End If
        End If
      End If
      ' Consider the parents
      ndxNext = ndxSrc
      While (Not ndxNext.ParentNode Is Nothing)
        ' Move to the parent
        ndxNext = ndxNext.ParentNode
        ' Is this <eTree>?
        If (ndxNext.Name = "eTree") Then
          ' Check this parent
          If (DoLike(ndxNext.Attributes("Label").Value, "NP|NP-*")) Then
            ' We found the antecedent
            ndxAnt = ndxNext
            Return True
          End If
        End If
      End While
      ' Consider preceding siblings
      ndxNext = ndxSrc
      While (Not ndxNext.PreviousSibling Is Nothing)
        ' Move to that sibling
        ndxNext = ndxNext.PreviousSibling
        ' Is this <eTree>?
        If (ndxNext.Name = "eTree") Then
          ' Check this sibling
          If (DoLike(ndxNext.Attributes("Label").Value, "NP|NP-*")) Then
            ' We found the antecedent
            ndxAnt = ndxNext
            Return True
          End If
        End If
      End While
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetAppAntecedent error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetConjpHeadNoun
  ' Goal:   Get the head noun of a NP containing an NP and CONJP
  ' History:
  ' 25-11-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetConjpHeadNoun(ByRef ndxParent As XmlNode) As XmlNode
    Dim ndxNext As XmlNode            ' Working node
    Dim ndxList As XmlNodeList        ' Working list
    Dim bHasConjP As Boolean = False  ' Initially we have not found CONJP

    Try
      ' Verify
      If (ndxParent Is Nothing) Then Return Nothing
      ' Start with the last child
      ndxNext = ndxParent.LastChild
      ' Work backwards
      While (Not ndxNext Is Nothing)
        ' Is this an <eTree>?
        If (ndxNext.Name = "eTree") Then
          ' Check if we had ConjP already
          If (bHasConjP) Then
            ' We've had CONJP, now THIS is the preceding one
            ' Make a list of this one's children
            ndxList = ndxNext.SelectNodes("./child::eTree[" & strHdNounTypesTb & "]", conTb)
            ' Any results?
            If (ndxList.Count > 0) Then
              ' Return the last item
              Return ndxList.Item(ndxList.Count - 1)
            ElseIf (DoLike(ndxNext.Attributes("Label").Value, "NP|NP-*")) Then
              ' The element preceding the CONJP is a kind of NP
              ndxList = ndxNext.SelectNodes("./child::eTree[" & strHdNounTypesTb & "]", conTb)
              ' Any results?
              If (ndxList.Count > 0) Then
                ' Return the last item
                Return ndxList.Item(ndxList.Count - 1)
              End If
            End If
          Else
            ' Check if this is CONJP?
            If (ndxNext.Attributes("Label").Value = "CONJP") Then
              ' Now we need to have the immediately preceding <eTree>
              bHasConjP = True
            End If
          End If
        End If
        ' Go to preceding sibling
        ndxNext = ndxNext.PreviousSibling
      End While
      ' Return failure
      Return Nothing
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetConjpHeadNoun error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetRootNode
  ' Goal:   Retrieve the chain root of [ndxThis]
  '         Only links of type Identity or CrossSpeech are allowed here
  ' History:
  ' 26-11-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetRootNode(ByRef ndxThis As XmlNode) As XmlNode
    Dim ndxNext As XmlNode      ' Working node
    Dim colId As New StringColl ' Collection of IDs
    Dim intId As Integer        ' Id we are looking at

    Try
      ' Validate 
      If (ndxThis Is Nothing) Then Return Nothing
      If (ndxThis.Name <> "eTree") Then Return Nothing
      ' Start with myself
      ndxNext = ndxThis
      ' Loop deep down...
      Do
        ' Check what ID I have
        intId = ndxNext.Attributes("Id").Value
        ' Check circularity
        If (colId.Exists(intId)) Then
          ' There is circularity, so return nothing
          Return Nothing
        End If
        ' Keep track of the ID's we've had
        colId.Add(intId)
        ' Check if this one has no more coref destination
        If (CorefDstId(ndxNext) < 0) OrElse _
           (Not DoLike(CorefAttr(ndxNext, "RefType"), strRefIdentity & "|" & strRefCross)) Then
          ' THis is the last one!
          Return ndxNext
        End If
        ' Go to the next one
        ndxNext = CorefDst(ndxNext)
      Loop Until (ndxNext Is Nothing)
      ' Return failure
      Return Nothing
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetRootNode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   GetRootHead
  ' Goal:   Retrieve the first head noun in the coreference link of [ndxThis]
  ' History:
  ' 26-11-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetRootHead(ByRef ndxThis As XmlNode) As String
    Dim ndxNext As XmlNode      ' Working node
    Dim ndxWork As XmlNode      ' Other workind node (for traces)
    Dim ndxForest As XmlNode    ' The forest we are in
    Dim strHdNoun As String     ' Head noun we find here
    Dim strLabel As String      ' Label of me
    Dim strSelect As String     ' Select statement to find trace in forest
    Dim intId As Integer        ' Id we are looking at
    Dim intRefId As Integer     ' ID of the trace
    Dim colId As New StringColl ' Collection of IDs

    Try
      ' Validate 
      If (ndxThis Is Nothing) Then Return ""
      If (ndxThis.Name <> "eTree") Then Return ""
      ' Initialisations
      ndxForest = Nothing
      ' Start with myself
      ndxNext = ndxThis
      colId.Add(ndxNext.Attributes("Id").Value)
      ' Try to walk the chain
      Do
        ' Get the next one
        ndxNext = CorefDst(ndxNext)
        ' Process the ID if possible
        If (Not ndxNext Is Nothing) Then
          intId = ndxNext.Attributes("Id").Value
          ' =========== DEBUGGING ===========
          'If (intId = 155) Then Stop
          ' =================================
          ' Check circularity
          If (colId.Exists(intId)) Then Return ""
          colId.Add(intId)
          ' Check this reference
          strHdNoun = GetHeadNoun(ndxNext)
          If (strHdNoun <> "") Then
            ' Return this one
            Return strHdNoun
          End If
          ' Check if this node is a trace
          strLabel = LeafText(ndxNext)
          If (Left(strLabel, 3) = "*T*") Then
            ' THis is a trace
            intRefId = CInt(Mid(strLabel, InStrRev(strLabel, "-") + 1))
            ' Get the forest
            If (GetForestNode(ndxThis, ndxForest)) Then
              ' Get the antecedent of our pronoun
              strSelect = "./descendant::eTree[@Id<" & intId & " and tb:Like(string(@Label), '*-" & intRefId & "')]"
              ndxWork = ndxForest.SelectSingleNode(strSelect, conTb)
              ' Did we get something?
              If (Not ndxWork Is Nothing) Then
                ' We should go to the parent of this node
                ndxWork = ndxWork.ParentNode
                ' Is this really a CP node?
                If (Not ndxWork Is Nothing) Then
                  ' Check label
                  If (ndxWork.Attributes("Label").Value Like "CP*") Then
                    ' The parent is a CP node, so now we need to get the head noun of its parent
                    strHdNoun = GetHeadNoun(ndxWork.ParentNode)
                    If (strHdNoun <> "") Then
                      ' Return this one
                      Return strHdNoun
                    End If
                  End If
                End If
              End If
            End If
          ElseIf (IsFreeOrDrelative(ndxNext)) Then
            ' THis is a D-relative or a Free relative, so return the whole NP as head noun
            Return NodeText(ndxNext)
          End If
        End If
      Loop Until (ndxNext Is Nothing)
      ' Return failure
      Return ""
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetRootNode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IsFreeOrDrelative
  ' Goal:   See if this is a free relative or a D-relative
  ' History:
  ' 29-11-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function IsFreeOrDrelative(ByRef ndxThis As XmlNode) As Boolean
    Dim ndxWork As XmlNode = Nothing  ' Working node

    Try
      ' Validate 
      If (ndxThis Is Nothing) Then Return False
      ' Is this a D-relative?
      If (FirstChildImm(ndxThis, "D|D^*", ndxWork, strSkipTypes)) Then
        ' Check whether there is a CP-REL child
        If (SomeChild(ndxThis, "CP-REL*", ndxWork)) Then
          ' Found a D-relative
          Return True
        End If
      ElseIf (FirstChildImm(ndxThis, "CP-FRL*", ndxWork, strSkipTypes)) Then
        ' Found a free relative
        Return True
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/IsFreeOrDrelative error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetHeadNoun
  ' Goal:   See if this NP has a head noun. If so, return it
  '         The head noun is defined as the last [strNounTypes] labelled <eTree> child of the NP
  ' History:
  ' 10-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetHeadNoun(ByRef ndxThis As XmlNode) As String
    Dim ndxList As XmlNodeList = Nothing  ' The children with label "N" or "N^*" of this NP
    Dim ndxNoun As XmlNode = Nothing      ' The children with label "N" or "N^*" of this NP
    Dim strLabel As String                ' The node's label value
    Dim strLeaf As String                 ' The leaf text

    Try
      ' Validate 
      If (ndxThis Is Nothing) Then Return ""
      ' This should be an <eTree>
      If (ndxThis.Name <> "eTree") Then Return ""
      ' Get the node's label
      strLabel = ndxThis.Attributes("Label").Value
      If (Not strLabel Like "NP*") Then Return ""
      ' Special treatment for NPR$
      If (strLabel = "NPR$") Then
        ' There only is 1 child, and we need to return it
        Return LeafText(ndxThis)
      End If
      ' Look for the last appropriate child
      ndxList = ndxThis.SelectNodes("./child::eTree[" & strHdNounTypesTb & "]", conTb)
      ' Did we get something?
      If (ndxList Is Nothing) Then Return ""
      If (ndxList.Count = 0) Then
        ' Perhaps we have a sequence of NP ... CONJP?
        ndxNoun = GetConjpHeadNoun(ndxThis)
        ' Any result?
        If (ndxNoun Is Nothing) Then
          ' Find out what the list of <eTree> children is
          ndxList = ndxThis.SelectNodes("./child::eTree")
          If (ndxList.Count > 0) Then
            ' Return failure
            Return ""
          Else
            ' Really don't know...
            Return ""
          End If
        End If
      Else
        ' We have something, so get to the last element
        ndxNoun = ndxList.Item(ndxList.Count - 1)
      End If
      ' Get the leaf text of this noun
      strLeaf = LeafText(ndxNoun)
      ' Is this is a possessive?
      If (Right(ndxNoun.Attributes("Label").Value, 1) = "$") Then
        ' This is a possessive - try to get rid of the last "s" or "'s"
        If (Right(strLeaf, 2) = "'s") Then
          strLeaf = Left(strLeaf, strLeaf.Length - 2)
        ElseIf (Right(strLeaf, 1) = "s") Then
          strLeaf = Left(strLeaf, strLeaf.Length - 1)
        End If
      End If
      ' Return the text of this noun, and replace [$] with nothing
      Return strLeaf.Replace("$", "")
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetHeadNoun error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetHeadNoun
  ' Goal:   Get the head noun as a string
  ' History:
  ' 16-04-2012  ERK Derived from [GetHeadNoun]
  ' ------------------------------------------------------------------------------------
  Public Function GetHeadNoun(ByRef ndxThis As XmlNode, ByRef strBack As String) As Boolean
    Dim ndxHead As XmlNode  ' The node of the head

    Try
      ' Get the head noun
      ndxHead = GetHeadNounNode(ndxThis)
      ' Initialise
      strBack = ""
      ' Check
      If (Not ndxHead Is Nothing) Then
        ' Get the head noun string
        If (ndxHead.Name = "eLeaf") Then
          strBack = ndxHead.Attributes("Text").Value
        ElseIf (ndxHead.Attributes("Id").Value <> ndxThis.Attributes("Id").Value) Then
          strBack = NodeText(ndxHead)
        Else
          ' This is a possessive
          strBack = NodeText(ndxHead)
        End If
      End If
      ' Return success
      Return (strBack <> "")
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetHeadNoun error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetHeadNounNode
  ' Goal:   See if this NP has a head noun. If so, return it
  '         The head noun is defined as the last [strNounTypes] labelled <eTree> child of the NP
  ' History:
  ' 10-06-2010  ERK Derived from [GetHeadNoun]
  ' ------------------------------------------------------------------------------------
  Public Function GetHeadNounNode(ByRef ndxThis As XmlNode) As XmlNode
    Dim ndxList As XmlNodeList = Nothing  ' The children with label "N" or "N^*" of this NP
    Dim ndxNoun As XmlNode = Nothing      ' The children with label "N" or "N^*" of this NP
    Dim strLabel As String                ' The node's label value

    Try
      ' Validate 
      If (ndxThis Is Nothing) Then Return Nothing
      ' If this is a forest...
      If (ndxThis.Name = "forest") Then
        ' Take the first <eTree> under it
        ndxThis = ndxThis.SelectSingleNode("./descendant::eTree")
        ' Validate 
        If (ndxThis Is Nothing) Then Return Nothing
      End If
      ' Get the node's label
      strLabel = ndxThis.Attributes("Label").Value
      If (Not strLabel Like "NP*") Then Return ndxThis
      ' Special treatment for NPR$
      If (strLabel = "NPR$") Then
        ' There only is 1 child, and we need to return it
        Return ndxThis
      End If
      ' Look for the last appropriate child
      ndxList = ndxThis.SelectNodes("./child::eTree[" & strHdNounTypesTb & "]", conTb)
      ' Did we get something?
      If (ndxList Is Nothing) Then Return ndxThis
      If (ndxList.Count = 0) Then
        ' Perhaps we have a sequence of NP ... CONJP?
        ndxNoun = GetConjpHeadNoun(ndxThis)
        ' Any result?
        If (ndxNoun Is Nothing) Then
          ' Find out what the list of <eTree> children is
          ndxList = ndxThis.SelectNodes("./child::eTree")
          If (ndxList.Count > 0) Then
            ' Return failure
            Return ndxThis
          Else
            ' Really don't know...
            Return ndxThis
          End If
        End If
      Else
        ' We have something, so get to the last element
        ndxNoun = ndxList.Item(ndxList.Count - 1)
      End If
      ' Return the result
      Return ndxNoun
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetHeadNounNode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ndxThis
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetHeadNode
  ' Goal:   Return the constituent that is the head of the phrase in [ndxThis]
  '         - This phrase may be <eTree> or <forest>
  '         - Definition of head depends on the language, specified in [strOp]
  ' History:
  ' 24-04-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetHeadNode(ByRef ndxThis As XmlNode, ByVal strOp As String) As XmlNode
    Dim ndxList As XmlNodeList = Nothing  ' The children 
    Dim ndxWork As XmlNode = Nothing      ' Working node
    Dim strLabel As String                ' The node's label value
    Dim strLang As String                 ' The language to be used
    Dim strConst As String = ""           ' List of constituent children
    Dim intI As Integer                   ' Counter
    Dim arForestHd() As String = {"IP-*", "PP*", "NP*"}

    Try
      ' Validate 
      If (ndxThis Is Nothing) OrElse (Not DoLike(ndxThis.Name, "forest|eTree")) Then Return Nothing
      ' Determine the language
      Select Case strOp
        Case "DepChe"
          strLang = "che"
        Case "DepEng"
          strLabel = "en"
          ' Not yet implemented...
          Return Nothing
        Case Else
          ' Failure
          Return Nothing
      End Select
      ' Get all the <eTree> children, but no punctuation
      ndxList = ndxThis.SelectNodes("./child::eTree[count(child::eLeaf[@Type='Punct'])=0]")
      ' If there is no (valid) child, then there is no head
      If (ndxList.Count = 0) Then Return Nothing
      ' If there is only one child, then that is the head
      If (ndxList.Count = 1) Then Return ndxList(0)
      ' There are more children, so we need to evaluate them, depending on what we are
      If (ndxThis.Name = "forest") Then
        ' This is one method of determining what my head is
        ' The head of a <forest> node is very limited...
        ' However, in fact each (valid) <forest> child should be treated as "root" head...
        For intI = 0 To arForestHd.Length - 1
          ' Try get a child 
          If (SomeChild(ndxThis, arForestHd(intI), ndxWork)) Then
            ' Return this node
            Return ndxWork
          End If
        Next intI
        ' Getting here means that we do not yet know what the head is
        Stop
      End If
      ' Now we only need to deal with <eTree> nodes...
      ' Get the node's label
      strLabel = ndxThis.Attributes("Label").Value
      If (DoLike(strLabel, "CP-REL*")) Then
        ' Expect to find an IP-SUB
        If (SomeChild(ndxThis, "IP-SUB*", ndxWork)) Then Return ndxWork
      ElseIf (DoLike(strLabel, "IP-INF*")) Then
        ' Expect to find a VB+N
        If (SomeChild(ndxThis, "VB+N|VB+N-*", ndxWork)) Then Return ndxWork
        ' We can also have "TRUE" infinitival clauses, and the the infinitive verb is the head
        If (SomeChild(ndxThis, "VB|VB-*", ndxWork)) Then Return ndxWork
      ElseIf (DoLike(strLabel, "IP-MAT*")) Then
        ' Expect to find a finite verb
        If (SomeChild(ndxThis, "VBD*|VBP*|BED*|BEP*", ndxWork)) Then Return ndxWork
      ElseIf (DoLike(strLabel, "IP-SUB*")) Then
        ' Expect to find an attributive participle form???
        If (SomeChild(ndxThis, "[BV]A[GN]A*", ndxWork)) Then Return ndxWork
      ElseIf (DoLike(strLabel, "IP-PPL*")) Then
        ' Expect to find a predicative participle form???
        If (SomeChild(ndxThis, "[BV]A[GN]P*", ndxWork)) Then Return ndxWork
      ElseIf (DoLike(strLabel, "PP|PP-*")) Then
        ' Expect to find a postposition or an adverb
        If (SomeChild(ndxThis, "P|P-*", ndxWork)) Then Return ndxWork
        ' Otherwise an adverb
        If (SomeChild(ndxThis, "ADV|ADV-*", ndxWork)) Then Return ndxWork
      ElseIf (DoLike(strLabel, "NP|NP-*")) Then
        ' Get all N, NS, NPR alikes
        ndxList = ndxThis.SelectNodes("./child::eTree[tb:matches(string(@Label), 'N|NS|N-*|NS-*|NPR|NPR-*')]", conTb)
        ' Maybe we have the correct solution?
        If (ndxList.Count = 1) Then Return ndxList(0)
        ' Check for contingency
        If (ndxList.Count = 0) Then Stop
        ' There's more than one possibility: take the last one
        Return ndxList(ndxList.Count - 1)
      ElseIf (DoLike(strLabel, "CONJP*")) Then
        ' Expect to find a conjunction
        If (SomeChild(ndxThis, "CONJ|CONJ-*", ndxWork)) Then Return ndxWork
      End If
      ' We are in trouble!!! We are unable to determine the head
      Stop
      For intI = 0 To ndxList.Count - 1
        If (strConst <> "") Then strConst &= " "
        strConst &= ndxList(intI).Attributes("Label").Value
      Next intI
      Debug.Print("Unable to determine the head for [" & strConst & "]")
      ' Return failure
      Return Nothing
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetHeadNode error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function

End Module
