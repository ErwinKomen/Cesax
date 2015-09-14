Imports System.Xml
Imports System.Text.RegularExpressions
Module modAdapt
  ' ================================= LOCAL VARIABLES =========================================================
  Private pdxFile As XmlDocument                ' The XML document we are currently working on...
  Private pdxPara As XmlDocument                ' Parallel psdx file
  Private ndxList As XmlNodeList                ' List of all the NP items
  Private ndxListP As XmlNodeList               ' List of parallel NP items
  Private tblPro As DataTable = Nothing         ' Pronoun resolution table
  Private dtrPerPro() As DataRow = Nothing      ' The personal pronouns from the table
  Private dtrDemPro() As DataRow = Nothing      ' The demonstrative pronouns
  Private dtrAdv() As DataRow = Nothing         ' The adverbs from the Category table (marked "Adv-...")
  Private dtrVerb() As DataRow = Nothing        ' The verbs from the Category table (marked "Vb-...")
  Private tblPeriod As DataTable = Nothing      ' Period definitions: from period Id to period name
  Private tblPeriodFile As DataTable = Nothing  ' From filename to period ID
  Private tblCatTrack As DataTable = Nothing    ' Table with the pronoun tracking information
  Private tblAdvTrack As DataTable = Nothing    ' Table with the adverb tracking information
  Private tblVbTrack As DataTable = Nothing     ' Table with the verb tracking information
  Private colMissing As New StringColl          ' A collection of strings
  Private colHeadAnim As New StringColl         ' Heads that have been found to be animate
  Private colHeadInanim As New StringColl       ' Heads that were found to be inanimate
  Private colToronto As New StringColl          ' Correspondence between Toronto number and cameron id
  Private colTagReport As New StringColl        ' Tagging report comments
  Private strHdPossible As String = "N*|ADJ*|PRO*|WN*|Q*"
  Private strHtmlHead As String = "<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8' /></head>"
  Private loc_strLemmaSource As String = "ERK"  ' Source of lemmatization information
  Private Structure Inflected
    Dim Word As String      ' The actual word
    Dim WordType As String  ' Word type (POS equivalent)
    Dim Lemma As String     ' Lemma of the word
    Dim GrCase As String    ' Grammatical case
    Dim GrGender As String  ' Gender
    Dim GrNumber As String  ' Number
    Dim Grperson As String  ' Person
    Dim Tense As String     ' Tense
    Dim Mood As String      ' Mood
    Dim ComType As String   ' Complementation type
  End Structure
  Private loc_arOEfrom() As String = {"ea", "aw", "j", "i", "y", "y", "th", "dh", "ky", "cy", "mon", "man", "ae", "e", "cs", "x", "gc"}
  Private loc_arOEto() As String = {"ae", "au", "i", "y", "i", "ie", "dh", "th", "cy", "ky", "man", "mon", "e", "ae", "sc", "sc", "cg"}
  ' ================================ GLOBAL VARIABLES ==========================================================
  Public strLeafCondition As String = "(@Type = 'Vern' and not(parent::eTree[@Label = 'CODE' or @Label = '.']) )" & _
                                       " or (@Type = 'Punct' and parent::eTree[@Label = 'CONJ'])"
  Public strLeafCondition_NoFW As String = "(@Type = 'Vern' and not(parent::eTree[@Label = 'CODE' or @Label = '.' or @Label = 'FW']) )" & _
                                       " or (@Type = 'Punct' and parent::eTree[@Label = 'CONJ'])"
  Public str_UnaccDir As String = ""            ' Directory where unaccusative parallel psdx files are stored
  ' ============================================================================================================
  ' ------------------------------------------------------------------------------------
  ' Name:   SetLemmaSource
  ' Goal:   Set or get the lemma source abbreviation of the person that has provided the lemma's
  ' History:
  ' 20-01-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function SetLemmaSource(ByVal strIn As String) As Boolean
    Try
      loc_strLemmaSource = strIn
      Return True
    Catch ex As Exception
      Return False
    End Try
  End Function
  Public Function GetLemmaSource() As String
    Return loc_strLemmaSource
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   OnePsdFile
  ' Goal:   Make PSD output file from one PSDX input file
  ' History:
  ' 26-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function OnePsdFile(ByVal strInFile As String, ByVal strOutFile As String, _
                             Optional ByVal strType As String = "PsdSimple") As Boolean
    Try
      ' Validate
      If (Not IO.File.Exists(strInFile)) Then
        Status("Could not find file " & strInFile)
        Return False
      End If
      ' Show we are reading
      Status("Reading file " & strInFile)
      ' Try read this file into an XML structure
      If (ReadXmlDoc(strInFile, pdxFile)) Then
        ' Set the current file
        pdxCurrentFile = pdxFile
        ' Initialise the pdxList
        If (Not InitCurrentFile()) Then
          Logging("modAdapt/OnePsdFile: Could not initialise file " & strInFile)
          Return False
        End If
        ' Make PSD output from the current input
        If (Not DoExport(strType, strOutFile)) Then
          Logging("modAdapt/OnePsdFile: Could not produce PSD from " & strInFile)
          Return False
        End If
        ' Return success
        Return True
      Else
        ' Could not read the file, so return failure
        Status("Could not read the XML file: " & strInFile)
        Return False
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OnePsdFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   OneCalcFeaturesFile
  ' Goal:   Calculate the [strCat] types for all the constituents in the intput [strInFile]
  '           and store the resulting tree in the [strOutFile]
  ' History:
  ' 26-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function OneCalcFeaturesFile(ByVal strInFile As String, ByVal strOutFile As String, _
                                      ByVal strCat As String, Optional ByVal bCheck As Boolean = False) As Boolean
    Dim strParaFile As String = ""  ' Parallel file
    'Dim arFeatDef() As String   ' Feature definitions

    Try
      ' Validate
      If (Not IO.File.Exists(strInFile)) Then
        Status("Could not find file " & strInFile)
        Return False
      End If
      '' Do we have feature definitions?
      'If (strFeatDefFile = "") Then
      '  ReDim arFeatDef(0)
      'Else
      '  arFeatDef = Split(IO.File.ReadAllText(strFeatDefFile), vbLf)
      'End If
      ' Show we are reading
      Status("Reading file " & strInFile)
      ' Try read this file into an XML structure
      If (ReadXmlDoc(strInFile, pdxFile)) Then
        ' Any other processing?
        Select Case strCat
          Case CAT_VBUNACC
            ' Read parallel file
            strParaFile = str_UnaccDir & "\" & IO.Path.GetFileName(strInFile)
            If (Not ReadXmlDoc(strParaFile, pdxPara)) Then Logging("Could not open parallel file: " & strParaFile) : Return False
        End Select
        ' Continue...
        If (Not OneDoFeaturesPdx(pdxFile, strCat, bCheck)) Then
          ' Check if saving is needed
          If (bNeedSaving) Then
            ' Ask for confirmation
            Select Case MsgBox("Save changes?", MsgBoxStyle.YesNo)
              Case MsgBoxResult.Yes
                ' Save the PSDX file
                pdxFile.Save(strOutFile)
            End Select
          End If
          ' Return negatively
          Return False
        End If
        'If (Not OneNpFeaturesPdx(pdxFile)) Then Return False
        ' Write the result to the destination
        pdxFile.Save(strOutFile)
        ' Return success
        Return True
      Else
        ' Could not read the file, so return failure
        Status("Could not read the XML file: " & strInFile)
        Return False
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OneNpFeaturesFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   TagInit
  ' Goal:   Initialize the tag report
  ' History:
  ' 30-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub TagInit()
    colTagReport.Clear()
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   TagComment
  ' Goal:   Add a comment to the tag report
  ' History:
  ' 30-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub TagComment(ByVal strIn As String)
    Try
      colTagReport.Add(strIn)
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/TagComment error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   RemainReport
  ' Goal:   Return a html remain report
  ' History:
  ' 11-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function RemainReport() As String
    Dim colHtml As New StringColl
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim intI As Integer       ' Counter
    Dim intPtc As Integer     ' Percentage

    Try
      ' Start header
      colHtml.Add("<h3>Remain report</h3><table>")
      colHtml.Add("<tr><td>Vern</td><td>POS</td><td>Parent</td><td>Freq</td><td>Location</td><td>Clause</td></tr>")
      ' Question the dictionary with remainders
      dtrFound = tdlRemain.Tables("Remain").Select("", "Pos ASC, Vern ASC, PrntLabel ASC, Freq DESC")
      For intI = 0 To dtrFound.Length - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ dtrFound.Length
        Status("Remain Report " & intPtc & "%", intPtc)
        ' Access this line for the report
        With dtrFound(intI)
          ' Add the information in this line
          colHtml.Add("<tr><td>" & .Item("Vern") & "</td>" & vbCrLf & _
                      "<td>" & .Item("Pos") & "</td>" & vbCrLf & _
                      "<td>" & .Item("PrntLabel") & "</td>" & vbCrLf & _
                      "<td align='right'>" & .Item("Freq") & "</td>" & vbCrLf & _
                      "<td>[" & .Item("File") & ":" & .Item("forestId") & "." & .Item("EtreeId") & "]</td>" & vbCrLf & _
                      "<td>" & .Item("Clause") & "</td>" & vbCrLf & _
                      "</tr>")
        End With
      Next intI
      colHtml.Add("</table><p>")
      Return colHtml.Text
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/RemainReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return "error"
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   TagReport
  ' Goal:   Return a html tag report
  ' History:
  ' 30-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function TagReport() As String
    Dim intI As Integer
    Dim colHtml As New StringColl

    Try
      ' Start header
      colHtml.Add("<h3>Tag report</h3><table>")
      colHtml.Add("<tr><td>#</td><td>book</td><td>Done</td><td>Need</td></tr>")
      For intI = 0 To colTagReport.Count - 1
        colHtml.Add("<tr><td>" & intI + 1 & "</td>" & colTagReport.Item(intI) & "</tr>")
      Next intI
      colHtml.Add("</table><p>")
      Return colHtml.Text
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/TagReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return "error"
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ReadTorontoList
  ' Goal:   Read the "textlist.sgml" file of the Toronto corpus and store in conversion array
  ' History:
  ' 29-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function ReadTorontoList() As Boolean
    Dim dtrThis As DataRow = Nothing  ' Datarow
    Dim arLine() As String    ' Array of lines
    Dim arItem() As String    ' Elements within a line
    Dim arFile() As String    ' Array of files
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim strText As String     ' Text
    Dim strDelim As String    ' Delimiter
    Dim strToron As String    ' Toronto number
    Dim strCamno As String    ' Cameron ID
    Dim strTitle As String    ' Title
    Dim strPreCamno As String = "_r(camno:"
    Dim strPreTitle As String = "_r(t:"
    Dim strFile As String = "textlist.sgml" ' Where the correspondence is kept
    Dim intI As Integer       ' Counter

    Try
      ' Check if the structure already contains toronto-camno-file information
      If (tdlMorphTag.Tables("TorCam").Rows.Count > 0) Then Return True
      ' Read the list
      If (Not IO.File.Exists(strFile)) Then Return False
      strText = IO.File.ReadAllText(strFile)
      strDelim = GetDelim(strText, vbCrLf, vbCr, vbLf)
      arLine = Split(strText, strDelim)
      ' Process the file
      For intI = 0 To arLine.Length - 1
        ' Split the line into items
        arItem = Split(arLine(intI), " ")
        ' Is this a correct line?
        If (arItem(0) = "<!ENTITY") Then
          ' This is the correct line, so get the toronto number and the cameron id
          strToron = arItem(1)
          strCamno = arItem(5)
          ' Add a row to the correct datastructure
          dtrThis = AddOneDataRow(tdlMorphTag, "TorCam", "TorCamId", "TorCamList")
          If (dtrThis Is Nothing) Then TagComment("ReadTorontoList: could not add row to <TorCam>") : Return False
          With dtrThis
            .Item("Toron") = strToron
            .Item("Camno") = strCamno
            .Item("Title") = ""
            .Item("File") = ""
          End With
        End If
      Next intI
      ' Look for all files in the OE-tagged directory (or below)
      arFile = IO.Directory.GetFiles(strOEtaggedDir, "*.tagged", IO.SearchOption.AllDirectories)
      For intI = 0 To arFile.Length - 1
        ' Read this file
        strText = IO.File.ReadAllText(arFile(intI))
        ' Turn it into an array
        strDelim = GetDelim(strText, vbCrLf, vbCr, vbLf)
        arLine = Split(strText, strDelim)
        ' I expect the title and the camno to be on the first and second place
        strTitle = arLine(0) : If (InStr(strTitle, strPreTitle) = 0) Then TagComment("ReadToronto: cannot find title = " & arFile(intI)) : Return False
        strCamno = arLine(1) : If (InStr(strCamno, strPreCamno) = 0) Then TagComment("ReadToronto: cannot find camno = " & arFile(intI)) : Return False
        ' Unpack the title and the camno
        strTitle = Mid(strTitle, strPreTitle.Length + 1) : strTitle = Left(strTitle, strTitle.Length - 1)
        strCamno = Mid(strCamno, strPreCamno.Length + 1) : strCamno = Left(strCamno, strCamno.Length - 1)
        ' The cameron number may not end on an additional A, B, C, D
        If (DoLike(Right(strCamno, 1), "A|B|C|D")) AndAlso (Left(Right(strCamno, 2), 1) <> ".") Then strCamno = Left(strCamno, strCamno.Length - 1)
        ' Find the datarow with this camno
        dtrFound = tdlMorphTag.Tables("TorCam").Select("Camno='" & strCamno & "'")
        Select Case dtrFound.Length
          Case 0
            ' Did not find it
            TagComment("ReadToronto: could not find camno = " & strCamno)
            Return False
          Case 1
            ' Is it already populated?
            If (dtrFound(0).Item("Title").ToString = "") Then
              ' Add it
              With dtrFound(0)
                .Item("Title") = strTitle
                .Item("File") = arFile(intI)
              End With
            Else
              ' We need to make a new entry
              strToron = dtrFound(0).Item("Toron")
              dtrThis = AddOneDataRow(tdlMorphTag, "TorCam", "TorCamId", "TorCamList")
              If (dtrThis Is Nothing) Then TagComment("ReadTorontoList: could not add row to <TorCam>") : Return False
              With dtrThis
                .Item("Toron") = strToron
                .Item("Camno") = strCamno
                .Item("Title") = strTitle
                .Item("File") = arFile(intI)
              End With
            End If
          Case Else
            '' More than one possibility
            'TagComment("ReadToronto: ambiguous camno = " & strCamno)
            ' Return False
            ' We need to make a new entry
            strToron = dtrFound(0).Item("Toron")
            dtrThis = AddOneDataRow(tdlMorphTag, "TorCam", "TorCamId", "TorCamList")
            If (dtrThis Is Nothing) Then TagComment("ReadTorontoList: could not add row to <TorCam>") : Return False
            With dtrThis
              .Item("Toron") = strToron
              .Item("Camno") = strCamno
              .Item("Title") = strTitle
              .Item("File") = arFile(intI)
            End With
        End Select
      Next intI
      ' Make sure changes are saved in the appropriate file
      tdlMorphTag.WriteXml(strMorphTagFile)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/ReadTorontoList error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetCamno
  ' Goal:   Given the Toronto number, get the Cameron number
  ' History:
  ' 29-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetCamno(ByVal strToron As String) As String
    Dim dtrFound() As DataRow   ' Result of SELECT
    Dim strCamno As String = "" ' Cameron ID

    Try
      ' Validate
      If (colToronto.Count = 0) Then
        ' Try read it
        If (Not ReadTorontoList()) Then Return ""
      End If
      ' Possibly adapt the toronto number
      strToron = Left(strToron, 6)
      ' Try to find this toronto entry
      dtrFound = tdlMorphTag.Tables("TorCam").Select("Toron='" & strToron & "'")
      If (dtrFound.Length = 0) Then Return ""
      ' Retrieve the cameron ID
      strCamno = dtrFound(0).Item("Camno").ToString
      ' Return the Cameron ID
      Return strCamno
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetCamno error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   OneTaggedFeaturesFile
  ' Goal:   Facilitate adding features from OE-tagged to this YCOE file 
  ' Algorithm:
  '  FOR each $file in YCOE.files
  '  $first := $file.FirstForest
  '  $torF  := GetToronto($first)
  '  IF (EXISTS($torF)) THEN
  '    $tagFile := GetTaggedFile($torF, ‘textlist.sgml’)
  '    FOR each $forest in $file.forests
  '  	   $tor := GetToronto($forest)
  '      IF ($tor != $torF) THEN
  '        $tagFile := GetTaggedFile($tor, ‘textlist.sgml’)
  '      END IF
  '      FOR each $leaf in $forest.leaves
  '        $tagUnit := GetNextUnit($tagFile)
  '        $leaf.Features := $tagUnit.Features
  '      NEXT $leaf
  '    NEXT $forest
  '   END IF
  ' NEXT $file
  ' History:
  ' 29-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function OneTaggedFeaturesFile(ByVal strInFile As String, ByVal strOutFile As String, ByRef intDone As Integer, ByRef intNeed As Integer) As Boolean
    Dim strShort As String = ""     ' Short file name
    Dim bChanged As Boolean = False ' Any changes?

    Try
      ' Validate
      If (Not IO.File.Exists(strInFile)) Then Return False
      ' Determine short filename
      strShort = IO.Path.GetFileNameWithoutExtension(strOutFile)
      ' Try read this file into an XML structure
      If (ReadXmlDoc(strInFile, pdxFile)) Then
        ' Check if the OE-tagged information has already been added
        If (pdxFile.SelectSingleNode("//teiHeader/revisionDesc/change[@comment='" & OE_TAGGED_DONE & "']") Is Nothing) Then
          ' Set current file value
          strCurrentFile = strInFile
          ' Continue...
          If (Not OneTaggedFeaturesPdx(pdxFile, bChanged)) Then
            ' Check if saving is needed
            If (bNeedSaving) Then
              ' Ask for confirmation
              Select Case MsgBox("Save changes?", MsgBoxStyle.YesNo)
                Case MsgBoxResult.Yes
                  ' Save the PSDX file
                  pdxFile.Save(strOutFile)
              End Select
            End If
            ' Return negatively
            Return False
          End If
          ' Indicate that we have made the OE-tagged changes
          AddRevDesc(pdxCurrentFile, strUserName, Format(Today, "d"), OE_TAGGED_DONE)
          'If (Not OneNpFeaturesPdx(pdxFile)) Then Return False
          If (bChanged) Then
            ' Write the result to the destination
            pdxFile.Save(strOutFile)
          Else
            ' Report that there were no changes
            TagComment("Did not save [" & strShort & "]: there were no changes")
          End If
        End If
        ' Check how many have been done and how many still need doing
        If (OneTaggedCheckPdx(pdxFile, intDone, intNeed)) Then
          ' Give an appropriate comment
          TagComment("book=[" & strShort & "] done=" & intDone & " need=" & intNeed)
        End If
        ' Return success
        Return True
      End If
      ' Could not read the file, so return failure
      Status("Could not read the XML file: " & strInFile)
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OneTaggedFeaturesFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   OneTaggingReportFile
  ' Goal:   Facilitate morphology processing of OE files 
  ' History:
  ' 18-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function OneTaggingReportFile(ByVal strInFile As String, ByRef intDone As Integer, ByRef intNeed As Integer, _
                                       Optional ByVal strLabels As String = "", Optional ByVal strLngAbbr As String = "") As Boolean
    Dim strIn As String = ""          ' Short file name (input)
    Dim strVern As String = ""        ' Vernacular
    Dim strLemma As String = ""       ' Each entry should have a lemma value
    Dim strPos As String = ""         ' part of speech 
    Dim strLabel As String = ""       ' label of <eTree> parent
    Dim strFeat As String = ""        ' List of features
    ' Dim bChanged As Boolean           ' Change flag

    Try
      ' Validate
      If (Not IO.File.Exists(strInFile)) Then Return False
      ' Get the short filename
      strIn = IO.Path.GetFileNameWithoutExtension(strInFile)
      ' Try read this file into an XML structure
      If (ReadXmlDoc(strInFile, pdxFile)) Then
        '' Delete empty lemma's
        'bChanged = False
        'If (Not OneEmptyLemmaCheck(pdxFile, bChanged)) Then Return False
        '' Are we changed?
        'If (bChanged) Then
        '  ' Save the changed psdx file
        '  pdxFile.Save(strInFile)
        'End If
        strCurrentFile = strInFile
        ' Check how many have been done and how many still need doing
        If (OneTaggedCheckPdx(pdxFile, intDone, intNeed, strLabels, strLngAbbr)) Then
          ' Give an appropriate comment
          ' TagComment("book=[" & strIn & "] done=" & intDone & " need=" & intNeed)
          TagComment("<td>" & strIn & "</td><td align='right'>" & intDone & "</td><td align='right'>" & intNeed & "</td>")
        End If
        ' Return success
        Return True
      End If
      ' Could not read the file, so return failure
      Status("Could not read the XML file: " & strInFile)
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OneTaggingReportFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   OneMorphCatFile
  ' Goal:   Gather all the POS categories, asking whether they are "Main" or "Derived"
  '         And if they are derived, what their "Head" is 
  ' History:
  ' 11-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function OneMorphCatFile(ByVal strInFile As String) As Boolean
    Dim ndxFor As XmlNode = Nothing   ' Forest node
    Dim ndxWork As XmlNode = Nothing  ' Working node
    Dim ndxThis As XmlNode = Nothing  ' Current node
    Dim ndxList As XmlNodeList        ' List of <eLeaf> nodes
    Dim strIn As String = ""          ' Short file name (input)
    Dim strVern As String = ""        ' Vernacular
    Dim strLemma As String = ""       ' Each entry should have a lemma value
    Dim strPrntLabel As String = ""   ' Label of parent phrase
    Dim strPos As String = ""         ' part of speech 
    Dim strLabel As String = ""       ' label of <eTree> parent
    Dim strFeat As String = ""        ' List of features
    Dim strShort As String = ""       ' Short file name
    Dim intCount As Integer           ' Number of child nodes
    Dim intPtc As Integer             ' Percentage

    Try
      ' Validate
      If (Not IO.File.Exists(strInFile)) Then Return False
      ' Get the short filename
      strIn = IO.Path.GetFileNameWithoutExtension(strInFile)
      ' Try read this file into an XML structure
      If (ReadXmlDoc(strInFile, pdxFile)) Then
        ' Set current file
        pdxCurrentFile = pdxFile
        strCurrentFile = strInFile
        strShort = IO.Path.GetFileNameWithoutExtension(strInFile)
        ' Look for the first <forest> element
        If (Not GetFirstForest(pdxFile, ndxFor)) Then Return False
        ' Determine how many elements there are potentially
        If (ndxFor Is Nothing) Then
          intCount = 0
        ElseIf (ndxFor.ParentNode Is Nothing) Then
          intCount = 0
        Else
          intCount = ndxFor.ParentNode.ChildNodes.Count
        End If
        ' Walk through all the <forest> elements FORWARD
        While (Not ndxFor Is Nothing)
          ' Show where we are (how do we KNOW where we are?)
          intPtc = (ndxFor.Attributes("forestId").Value) * 100 \ intCount
          Status("OE-pos categories of [" & IO.Path.GetFileNameWithoutExtension(strCurrentFile) & "] " & intPtc & "%", intPtc)
          ' Get all the <eLeaf> elements of this forest, provided they are vernacular (skipping punctuation, skipping Foreign Words)
          ' NB: there may be some punctuation that slipped through; their parent label is "."
          ndxList = ndxFor.SelectNodes("./descendant::eLeaf[" & strLeafCondition_NoFW & "]")
          ' Walk them all
          For intI = 0 To ndxList.Count - 1
            ' Get my parent, which is the <eTree> we need to have the POS from
            ndxThis = ndxList(intI).ParentNode
            ' Get the POS from this one
            strPos = ndxThis.Attributes("Label").Value
            ' Set the word in the sectiontext
            frmMain.SectionText(VernToEnglish(ndxList(intI).Attributes("Text").Value))
            ' Add this category
            If (Not MorphCatAdd(strPos)) Then Return False
          Next intI
          ' Go to the next forest
          ndxFor = ndxFor.NextSibling
        End While
        ' Return success
        Return True
      End If
      ' Could not read the file, so return failure
      Status("Could not read the XML file: " & strInFile)
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OneMorphCatFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   OneRemainReportFile
  ' Goal:   Add the report information of [strInFile] to the report 
  ' History:
  ' 11-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function OneRemainReportFile(ByVal strInFile As String, ByVal strLngAbbr As String, Optional ByVal strLabels As String = "") As Boolean
    Dim ndxFor As XmlNode = Nothing   ' Forest node
    Dim ndxWork As XmlNode = Nothing  ' Working node
    Dim ndxThis As XmlNode = Nothing  ' Current node
    Dim ndxList As XmlNodeList        ' List of <eLeaf> nodes
    Dim ndxLeaf As XmlNode          ' My own leaf
    Dim intCount As Integer           ' Number of child nodes
    Dim strIn As String = ""          ' Short file name (input)
    Dim strVern As String = ""        ' Vernacular
    Dim strClause As String = ""      ' Clause
    Dim strClauseH As String = ""     ' Clause with highlighted vernacular word
    Dim strLemma As String = ""       ' Each entry should have a lemma value
    Dim strPrntLabel As String = ""   ' Label of parent phrase
    Dim strPos As String = ""         ' part of speech 
    Dim strLabel As String = ""       ' label of <eTree> parent
    Dim strFeat As String = ""        ' List of features
    Dim strShort As String = ""       ' Short file name
    Dim strVernCombi As String = "" ' Combined vernacular
    Dim bFound As Boolean             ' Process this one or not?
    Dim intCombi As Integer         ' Number of parts to a word
    Dim intPtc As Integer             ' Percentage
    Dim intPos As Integer             ' Position
    Dim intI As Integer               ' Counter
    Dim intK As Integer               ' Counter

    Try
      ' Validate
      If (Not IO.File.Exists(strInFile)) Then Return False
      ' Get the short filename
      strIn = IO.Path.GetFileNameWithoutExtension(strInFile)
      ' Try read this file into an XML structure
      If (ReadXmlDoc(strInFile, pdxFile)) Then
        ' Set current file
        pdxCurrentFile = pdxFile
        strCurrentFile = strInFile
        strShort = IO.Path.GetFileNameWithoutExtension(strInFile)
        ' Look for the first <forest> element
        If (Not GetFirstForest(pdxFile, ndxFor)) Then Return False
        ' Determine how many elements there are potentially
        If (ndxFor Is Nothing) Then
          intCount = 0
        ElseIf (ndxFor.ParentNode Is Nothing) Then
          intCount = 0
        Else
          intCount = ndxFor.ParentNode.ChildNodes.Count
        End If
        ' Walk through all the <forest> elements FORWARD
        While (Not ndxFor Is Nothing)
          ' Show where we are (how do we KNOW where we are?)
          intPtc = (ndxFor.Attributes("forestId").Value) * 100 \ intCount
          Status("Tagged remains of [" & IO.Path.GetFileNameWithoutExtension(strCurrentFile) & "] " & intPtc & "%", intPtc)
          ' Get the clause
          strClause = StringOEtoTagged(LCase(ndxFor.SelectSingleNode("./div[@lang='org']/seg").InnerText)).Replace("$", "")
          ' Get all the <eLeaf> elements of this forest, provided they are vernacular (skipping punctuation, skipping Foreign Words)
          ' NB: there may be some punctuation that slipped through; their parent label is "."
          If (strLabels = "") Then
            ndxList = ndxFor.SelectNodes("./descendant::eLeaf[" & strLeafCondition_NoFW & "]")
          Else
            ' Only include those that are in [strLabels]
            ndxList = ndxFor.SelectNodes("./descendant::eLeaf[(" & strLeafCondition_NoFW & ") and " & _
                                         " parent::eTree[tb:matches(@Label, '" & strLabels & "')] ]", conTb)
          End If
          ' Walk them all
          For intI = 0 To ndxList.Count - 1
            ' Get my parent
            ndxThis = ndxList(intI).ParentNode : bFound = False
            ' Action depends on language
            Select Case strLngAbbr
              Case "OEB"
                ' Get the @f value
                strFeat = GetFeature(ndxThis, "M", "s")
                ' Do we have this?
                bFound = (strFeat <> "")
              Case Else
                bFound = (ndxThis.SelectSingleNode("./child::fs[@type='M']") IsNot Nothing)
             End Select
            ' Process this or not?
            If (Not bFound) Then
              ' No M features, so process this one
              ' (1) Get parent node
              ndxWork = ndxThis.ParentNode
              If (ndxWork.Name = "eTree") Then strPrntLabel = LabelmainOE(ndxWork.Attributes("Label").Value) Else strPrntLabel = "$"
              ' (2) Get own label
              strPos = ndxThis.Attributes("Label").Value
              ' (3) Get vernacular text
              ' Check for parts
              If (DoLike(strPos, "*[234]1")) Then
                ' Vernacular is the combination of all the entries
                strVernCombi = ndxList(intI).Attributes("Text").Value
                intCombi = CInt(Left(Right(strPos, 2), 1))
                For intK = 1 To intCombi - 1
                  strVernCombi &= ndxThis.SelectSingleNode("./following::eLeaf[" & intK & "]").Attributes("Text").Value
                Next intK
                strVern = strVernCombi
              ElseIf (DoLike(strPos, "*[234][234]")) Then
                strVern = strVernCombi
              Else
                strVern = ndxList(intI).Attributes("Text").Value
              End If
              ' Adapt the vernacular string so that it is compatible
              strVern = StringOEtoTagged(LCase(strVern)).Replace("$", "")

              ' strVern = StringOEtoTagged(LCase(ndxList(intI).Attributes("Text").Value)).Replace("$", "")
              ' (4) Get the vernacular text of the whole sentence
              intPos = InStr(strClause, strVern)
              If (intPos > 0) Then
                strClauseH = Left(strClause, intPos - 1) & "***" & strVern & "***" & Mid(strClause, intPos + strVern.Length + 1)
              Else
                strClauseH = strClause
              End If
              ' (5) Add the information to the appropriate datatable
              If (Not MorphRemainAdd(strVern, strClauseH, strPos, strPrntLabel, strShort, ndxFor.Attributes("forestId").Value, _
                                     ndxThis.Attributes("Id").Value)) _
                Then Logging("modAdapt/OneRemainReportFile: could not add MorphRemain") : Return False
            End If
          Next intI
          ' Go to the next forest
          ndxFor = ndxFor.NextSibling
        End While
        ' Return success
        Return True
      End If
      ' Could not read the file, so return failure
      Status("Could not read the XML file: " & strInFile)
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OneRemainReportFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   OneMorphGatherFile
  ' Goal:   Facilitate morphology processing of OE files 
  ' History:
  ' 18-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function OneMorphGatherFile(ByVal strInFile As String, ByRef intDone As Integer, ByRef intNeed As Integer) As Boolean
    Dim ndxList As XmlNodeList        ' Result of select
    Dim ndxFeat As XmlNodeList        ' List of all morphological features
    Dim ndxFor As XmlNode = Nothing   ' Forest node
    Dim ndxWork As XmlNode = Nothing  ' Working node
    Dim ndxThis As XmlNode = Nothing  ' Current node
    Dim ndxLeaf As XmlNode = Nothing  ' Leaf endnode
    Dim strIn As String = ""          ' Short file name (input)
    Dim strVern As String = ""        ' Vernacular
    Dim strLemma As String = ""       ' Each entry should have a lemma value
    Dim strPos As String = ""         ' part of speech 
    Dim strLabel As String = ""       ' label of <eTree> parent
    Dim strFeat As String = ""        ' List of features
    Dim bChanged As Boolean = False   ' Any changes?
    Dim dtrFound() As DataRow         ' Result of SELECT
    Dim dtrNew As DataRow = Nothing   ' One datarow
    Dim intCount As Integer = 0       ' Counter
    Dim intPtc As Integer             ' Percentage
    Dim intI As Integer               ' Counter
    Dim intJ As Integer               ' Counter

    Try
      ' Validate
      If (Not IO.File.Exists(strInFile)) Then Return False
      ' Try read this file into an XML structure
      If (ReadXmlDoc(strInFile, pdxFile)) Then
        ' Set current file value
        strCurrentFile = strInFile : pdxCurrentFile = pdxFile : strIn = IO.Path.GetFileNameWithoutExtension(strCurrentFile)
        ' Look for the first <forest> element
        If (Not GetFirstForest(pdxFile, ndxFor)) Then Return False
        ' Determine how many elements there are potentially
        If (ndxFor Is Nothing) Then
          intCount = 0
        ElseIf (ndxFor.ParentNode Is Nothing) Then
          intCount = 0
        Else
          intCount = ndxFor.ParentNode.ChildNodes.Count
        End If
        ' Walk through all the <forest> elements FORWARD
        While (Not ndxFor Is Nothing)
          ' Show where we are (how do we KNOW where we are?)
          intPtc = (ndxFor.Attributes("forestId").Value) * 100 \ intCount
          Status("OE-morphology of [" & strIn & "] " & intPtc & "%", intPtc)
          ' Look for all locations of feature-set "M" in this forest
          ndxList = ndxFor.SelectNodes("./descendant::fs[@type='M']")
          ' Walk them
          For intI = 0 To ndxList.Count - 1
            ' Check if we are not being interrupted
            If (bInterrupt) Then Return False
            Application.DoEvents()
            ' Get this one, as well as the <eTree> node and the <eLeaf> child
            ndxThis = ndxList(intI) : ndxWork = ndxThis.ParentNode : ndxLeaf = ndxWork.SelectSingleNode("./child::eLeaf")
            strPos = ndxWork.Attributes("Label").Value
            ' ========== DEBUG ===========
            ' If (InStr(ndxLeaf.Attributes("Text").Value, "$") > 0) Then Stop
            ' ============================
            ' Adapt the vernacular string so that it is compatible
            strVern = StringOEtoTagged(LCase(ndxLeaf.Attributes("Text").Value)).Replace("$", "")
            ' Only get the MAIN part of the parent <eTree> node, e.g: NP, IP, PP etc
            strLabel = LabelmainOE(ndxWork.ParentNode.Attributes("Label").Value)
            ' Get a list of all relevant features
            ndxFeat = ndxThis.SelectNodes("./child::f") : strFeat = "" : strLemma = ""
            For intJ = 0 To ndxFeat.Count - 1
              ' Is this the lemma?
              If (ndxFeat(intJ).Attributes("name").Value = "l") Then
                ' Get it
                strLemma = ndxFeat(intJ).Attributes("value").Value
              Else
                ' Add this feature's name and value to the stack-list
                AddSemiStack(strFeat, ndxFeat(intJ).Attributes("name").Value & "=" & ndxFeat(intJ).Attributes("value").Value)
              End If
            Next intJ
            ' Create an entry for the morphological dictionary
            dtrFound = tdlMorphDict.Tables("Morph").Select("Vern='" & strVern.Replace("'", "''") & "' AND Pos='" & strPos.Replace("'", "''") & _
                       "' AND Label='" & strLabel.Replace("'", "''") & "' AND f='" & strFeat.Replace("'", "''") & "'")
            If (dtrFound.Length = 0) Then
              ' Create a new entry
              dtrNew = AddOneDataRow(tdlMorphDict, "Morph", "MorphId", "MorphList")
              With dtrNew
                .Item("Vern") = strVern : .Item("Pos") = strPos : .Item("Label") = strLabel
                .Item("l") = strLemma : .Item("f") = strFeat : .Item("File") = strIn
                .Item("forestId") = ndxFor.Attributes("forestId").Value
                .Item("EtreeId") = ndxWork.Attributes("Id").Value
                .Item("Freq") = 1
              End With
            Else
              ' Get this entry
              dtrNew = dtrFound(0)
              ' Adapt frequency
              dtrNew.Item("Freq") += 1
            End If
          Next intI
          ' Go to next forest
          ndxFor = ndxFor.NextSibling
        End While
        ' Return success
        Return True
      End If
      ' Could not read the file, so return failure
      Status("Could not read the XML file: " & strInFile)
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OneMorphGatherFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   OneMorphPropaDict
  ' Goal:   Propagate the information in [tdlMorphDict.VernPos] into this [strInFile]
  '           resulting in an updated [strOutFile]
  ' History:
  ' 08-03-2013  ERK Created
  ' 04-10-2013  ERK Added [strLabels]
  ' 20-05-2014  ERK Added language dependence
  ' ------------------------------------------------------------------------------------
  Public Function OneMorphPropaDict(ByVal strInFile As String, ByVal strOutFile As String, ByVal strLngAbbr As String, _
                                    Optional ByVal strDoTags As String = "") As Boolean
    Dim ndxList As XmlNodeList        ' Result of select
    Dim ndxFor As XmlNode = Nothing   ' Forest node
    Dim ndxWork As XmlNode = Nothing  ' Working node
    Dim ndxThis As XmlNode = Nothing  ' Current node
    Dim ndxParent As XmlNode          ' Help node
    Dim ndxLeaf As XmlNode = Nothing  ' Leaf endnode
    Dim strShort As String = ""       ' Short file name (output)
    Dim strIn As String = ""          ' Short file name (input)
    Dim strVern As String = ""        ' Vernacular
    Dim strVernCombi As String = ""   ' Vernacular of a combination
    Dim strLemma As String = ""       ' Each entry should have a lemma value
    Dim strPos As String = ""         ' part of speech 
    Dim strLabel As String = ""       ' label of <eTree> parent
    Dim strMed As String = ""         ' MED number
    Dim strFeat As String = ""        ' List of features
    Dim strH As String = ""           ' Kind of history addition
    Dim strChoice As String = ""      ' Choice we make
    Dim strGener As String = ""       ' Generalization type
    Dim strPeriod As String = ""      ' Period of this file
    'Dim arFeat() As String            ' Array of features
    'Dim arLine() As String            ' Feature name and value
    Dim bChanged As Boolean = False   ' Any changes?
    Dim bConsider As Boolean = False  ' Need to consider this one or not?
    Dim dtrFound() As DataRow         ' Result of SELECT
    Dim dtrNew As DataRow = Nothing   ' One datarow
    Dim colLemma As New StringColl    ' Alternative lemma's
    Dim intCombi As Integer           ' Number of parts
    Dim intCount As Integer = 0       ' Counter
    Dim intChanges As Integer = 0     ' Number of changes
    Dim intPtc As Integer             ' Percentage
    Dim intPos As Integer             ' Position in string
    Dim intI As Integer               ' Counter
    Dim intJ As Integer               ' Counter
    Dim intK As Integer               ' counter
    ' Dim intChoice As Integer          ' Choice made by user

    Try
      ' Validate
      If (Not IO.File.Exists(strInFile)) Then Return False
      ' Determine short filename
      strShort = IO.Path.GetFileNameWithoutExtension(strInFile)
      ' Try read this file into an XML structure
      If (ReadXmlDoc(strInFile, pdxFile)) Then
        ' Set current file value
        strCurrentFile = strInFile : pdxCurrentFile = pdxFile : strIn = IO.Path.GetFileNameWithoutExtension(strCurrentFile)
        ' Try get epriod
        strPeriod = GetPeriod(pdxCurrentFile)
        ' Look for the first <forest> element
        If (Not GetFirstForest(pdxFile, ndxFor)) Then Return False
        ' Determine how many elements there are potentially
        If (ndxFor Is Nothing) Then
          intCount = 0
        ElseIf (ndxFor.ParentNode Is Nothing) Then
          intCount = 0
        Else
          intCount = ndxFor.ParentNode.ChildNodes.Count
        End If
        ' Other initialisation: a history note with the short date attached to it
        strH = "propVP_" & Format(Today, "d")
        ' Walk through all the <forest> elements FORWARD
        While (Not ndxFor Is Nothing)
          ' Show where we are (how do we KNOW where we are?)
          intPtc = (ndxFor.Attributes("forestId").Value) * 100 \ intCount
          Status("VernPos-morphdict-propagation [" & strIn & "] " & intPtc & "%", intPtc)
          ' Repair comma <eLeafs>...
          ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf)>0 and child::eLeaf[@Type='Vern'] and @Label=',']")
          ' Repair the type to 'Punct'
          For intI = 0 To ndxList.Count - 1
            ndxThis = ndxList(intI)
            ' Repair this child's type
            ndxLeaf = ndxThis.SelectSingleNode("./child::eLeaf")
            ndxLeaf.Attributes("Type").Value = "Punct"
            bChanged = True : intChanges += 1
          Next intI
          ' Repair conjunction <eLeafs>...
          ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf)>0 and child::eLeaf[@Type='Punct'] and @Label='CONJ']")
          ' Repair the type to 'Punct'
          For intI = 0 To ndxList.Count - 1
            ndxThis = ndxList(intI)
            ' Repair this child's type
            ndxLeaf = ndxThis.SelectSingleNode("./child::eLeaf")
            ' Check the text of this type
            If (ndxLeaf.Attributes("Text").Value = "&") Then
              ' This must be of type vernacular
              ndxLeaf.Attributes("Type").Value = "Vern"
              bChanged = True : intChanges += 1
            End If
          Next intI
          ' Remove entries with an empty lemma
          ndxList = ndxFor.SelectNodes("./descendant::fs[child::f[@name='l' and @value=' '] ]")
          ' Remove the bad guys
          For intI = 0 To ndxList.Count - 1
            ' Look at this one
            ndxThis = ndxList(intI) : ndxParent = ndxThis.ParentNode
            ' Remove all my children below
            ndxThis.RemoveAll()
            ' Remove myself
            ndxParent.RemoveChild(ndxThis)
            bChanged = True : intChanges += 1
          Next intI
          ' Look for all <eTree> nodes that have an <eLeaf> child and that do not yet have a feature-set of type "M"
          If (strDoTags = "") Then
            ' If we have language "OEB", we need to do much more
            If (strLngAbbr = "OEB") Then
              ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf)>0 and " & _
                                          "not(tb:matches(@Label,'CODE|.|,|FW')) " & _
                                          " and child::eLeaf[1][@Type='Vern']]", conTb)
            Else
              ' Default behaviour
              ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf)>0 and " & _
                                          "count(child::fs[@type='M'])=0 and not(tb:matches(@Label,'CODE|.|,|FW')) " & _
                                          " and child::eLeaf[1][@Type='Vern']]", conTb)
            End If
          Else
            ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf)>0 and " & _
                                         "count(child::fs[@type='M'])=0 and tb:matches(@Label,'" & strDoTags & "') " & _
                                         " and child::eLeaf[1][@Type='Vern']]", conTb)
          End If
          ' Walk all these <eTree> nodes and see if we can add features to them
          For intI = 0 To ndxList.Count - 1
            ' Check if we are not being interrupted
            If (bInterrupt) Then Return False
            Application.DoEvents()
            ' Get the <eTree> node and the <eLeaf> child
            ndxThis = ndxList(intI) : ndxLeaf = ndxThis.SelectSingleNode("./child::eLeaf") : ndxWork = ndxThis.ParentNode
            ' Get the <eTree> node's label
            strPos = ndxThis.Attributes("Label").Value
            ' Check for parts
            If (DoLike(strPos, "*[234]1")) Then
              ' This is part ONE of a larger whole
              ' Stop
              ' Vernacular is the combination of all the entries
              strVernCombi = ndxLeaf.Attributes("Text").Value
              'intCombi = CInt(Right(strPos, 1))
              'For intK = 1 To intCombi - 1
              '  strVernCombi &= ndxThis.SelectSingleNode("./following::eLeaf[" & intK & "]").Attributes("Text").Value
              'Next intK
              intCombi = CInt(Left(Right(strPos, 2), 1))
              For intK = 1 To intCombi - 1
                strVernCombi &= ndxThis.SelectSingleNode("./following::eLeaf[" & intK & "]").Attributes("Text").Value
              Next intK
              strVern = strVernCombi
            ElseIf (DoLike(strPos, "*[234][234]")) Then
              strVern = strVernCombi
            Else
              strVern = ndxLeaf.Attributes("Text").Value
            End If
            ' Now adapt [strPos] itself...
            If (strPos Like "*[234][1234]") Then
              strPos = Left(strPos, strPos.Length - 2)
            End If
            ' ========== DEBUG ===========
            '  If (strVern = "wur+ta+d") Then Stop
            ' ============================
            ' Adapt the vernacular string so that it is compatible
            If (Left(strCurrentFile, 2).ToLower = "cm") Then
              strVern = StringOEtoTagged(LCase(strVern), True).Replace("$", "")
            Else
              strVern = StringOEtoTagged(LCase(strVern)).Replace("$", "")
            End If
            ' Only get the MAIN part of the parent <eTree> node, e.g: NP, IP, PP etc
            If (ndxWork IsNot Nothing) AndAlso (ndxWork.Name = "eTree") Then
              strLabel = LabelmainOE(ndxWork.Attributes("Label").Value)
            Else
              strLabel = ""
            End If
            ' Check if we can get a positive ID for the Vern/Pos combination
            dtrFound = tdlMorphDict.Tables("VernPos").Select("Vern='" & strVern.Replace("'", "''") & "' AND Pos='" & strPos & "'")
            ' If [VernPos] doesn't work, try [Morph]
            If (dtrFound.Length = 0) Then
              dtrFound = tdlMorphDict.Tables("Morph").Select("Vern='" & strVern.Replace("'", "''") & "' AND Pos='" & strPos & "'")
            End If
            ' Try out language dependant variants
            ' If (Not MorphTryLngVariants(dtrFound, strVern, strPos, strFeat, strLngAbbr, strLemma, bChanged, intChanges)) Then Return False
            ' Don't immediately take "NO" for an answer - try variants for "Middle English"
            If (dtrFound.Length = 0) AndAlso (strLngAbbr = "ME") Then
              If (Not MakeMEDlemmaList(strVern, strVern, colLemma)) Then Return False
              ' Check if one of these has a hit
              For intJ = 0 To colLemma.Count - 1
                dtrFound = tdlMorphDict.Tables("VernPos").Select("Vern='" & colLemma.Item(intJ).Replace("'", "''") & "' AND Pos='" & strPos & "'")
                If (dtrFound.Length > 0) Then Exit For
              Next intJ
            End If
            If (dtrFound.Length > 0) Then
              ' Walk all possibilities!
              For intJ = 0 To dtrFound.Length - 1
                ' Access current element
                With dtrFound(intJ)
                  ' Action depends partly on the selected language
                  Select Case strLngAbbr
                    Case "OEB"
                      ' Get lemma and features
                      strLemma = .Item("l").ToString : strFeat = .Item("f").ToString
                      ' Check the lemma
                      If (strLemma <> GetFeature(ndxThis, "M", "l")) Then
                        ' Propagate lemma
                        If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "l", strLemma)) Then Return False
                        bChanged = True : intChanges += 1
                        ' Add the comment as feature "h"
                        If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "h", strH)) Then Return False
                      End If
                      ' Check if we need to propagate the features
                      If (InStr(strFeat, "BT") > 0) Then
                        '' Add the BT feature
                        'strMed = Regex.Match(strFeat, "BT\d+").Value
                        ' Propagate features, which includes the "s" feature
                        If (Not AddFeatureSet(ndxThis, "M", strFeat)) Then Return False
                        '' Propagate BT number
                        'If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "s", strMed)) Then Return False
                        bChanged = True : intChanges += 1
                        ' Add the comment as feature "h"
                        If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "h", strH)) Then Return False
                      ElseIf (DoLike(strPos, strAnyVerb)) AndAlso (Not DoLike(strPos, "NR*")) Then
                        ' This one needs to be present in the [VernPos] tabel
                        If (Not MorphAddOneVernPos(strVern, strLemma, strPos, strFeat, "autoEK")) Then Return False
                        bMorphVernPosChanged = True
                        Application.DoEvents()
                      End If
                    Case Else
                      ' Do we need to check for the parent label?
                      ' NB: alternatively the parent label equals the POS
                      If (.Item("PrntLabel").ToString = "") OrElse (.Item("PrntLabel").ToString = strLabel) OrElse (.Item("PrntLabel").ToString = strPos) Then
                        ' Get lemma and features
                        strLemma = .Item("l").ToString : strFeat = .Item("f").ToString
                        ' Get the Type and act on it
                        Select Case .Item("Type")
                          Case "LemmaFeat"
                            ' Propagate lemma
                            If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "l", strLemma)) Then Return False
                            ' Propagate features
                            If (Not AddFeatureSet(ndxThis, "M", strFeat)) Then Return False
                          Case "LemmaOnly"
                            ' Propagate lemma alone
                            If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "l", strLemma)) Then Return False
                            ' Check if there are features anyway
                            If (strFeat <> "") Then
                              ' Propagate features
                              If (Not AddFeatureSet(ndxThis, "M", strFeat)) Then Return False
                            End If
                          Case Else
                            ' I don't really know this...
                            ' Stop
                            ' Propagate lemma
                            If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "l", strLemma)) Then Return False
                            ' Are there any features?
                            If (strFeat <> "") Then
                              ' Propagate features
                              If (Not AddFeatureSet(ndxThis, "M", strFeat)) Then Return False
                            End If
                        End Select
                        ' Add the comment as feature "h"
                        If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "h", strH)) Then Return False
                        bChanged = True : intChanges += 1
                        ' Since we made a change: exit the For Intj loop
                        Exit For
                      End If
                  End Select
                End With
              Next intJ
            Else
              ' Check for verbal labels, which may depend on the particular language variant (OE, ME, eModE)
              If (MorphPosToLemma(strPeriod, strVern, strPos, strLemma, strFeat)) Then
                ' Propagate lemma
                If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "l", strLemma)) Then Return False
                If (strFeat <> "") Then
                  ' Propagate features
                  If (Not AddFeatureSet(ndxThis, "M", strFeat)) Then Return False
                End If
                bChanged = True : intChanges += 1
                ' Add this to the VernPos table
                If (Not MorphAddOneVernPos(strVern, strLemma, strPos, strFeat, "Derived from Pos+Lemma [" & Format(Now, "G") & "]")) Then Return False
              ElseIf (strLngAbbr = "OEB") AndAlso (DoLike(strPos, strAnyVerb)) Then
                 ' Get current lemma and feat 
                strLemma = GetFeature(ndxThis, "M", "l")
                strFeat = GetFeature(ndxThis, "M", "f")
                ' This one needs to be present in the [VernPos] tabel
                If (Not MorphAddOneVernPos(strVern, strLemma, strPos, strFeat, "autoEK")) Then Return False
                bMorphVernPosChanged = True
                Application.DoEvents()
              End If
            End If
          Next intI
          ' Go to next forest
          ndxFor = ndxFor.NextSibling
        End While
        ' Any changes?
        If (bChanged) Then
          ' Adapt the revision
          If (AddRevDesc(pdxCurrentFile, strUserName, Format(Now, "g"), "MorphPropagate VernPos")) Then
            ' Write the result
            Status("Saving adapted file...")
            pdxCurrentFile.Save(strInFile)
          End If
        End If
        ' Return success
        Return True
      End If
      ' Could not read the file, so return failure
      Status("Could not read the XML file: " & strInFile)
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OneMorphPropaDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphTryLngVariants
  ' Goal:   Try out language dependant variants
  ' History:
  ' 23-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function MorphTryLngVariants(ByRef dtrFound() As DataRow, ByVal strVern As String, ByVal strPos As String, ByVal strFeat As String, _
                                      ByVal strLngAbbr As String, ByRef strLemma As String, ByRef bChanged As Boolean, ByRef intChanges As Integer) As Boolean
    Dim intPos As Integer   ' Position in string

    Try
      ' OE adaptation
      If (strLngAbbr = "OE") Then
        ' Check for ^ mark
        intPos = InStr(strPos, "^")
        If (intPos > 0) Then
          dtrFound = tdlMorphDict.Tables("VernPos").Select("Vern='" & strVern.Replace("'", "''") & "' AND Pos LIKE '" & Left(strPos, intPos) & "*'")
          If (dtrFound.Length > 0) Then
            ' Get the lemma
            strLemma = dtrFound(0).Item("l").ToString
            ' Add this to the VernPos table
            If (Not MorphAddOneVernPos(strVern, strLemma, strPos, strFeat, "Derived: Pos with ^ [" & Format(Now, "G") & "]")) Then Return False
            bChanged = True : intChanges += 1
            Logging("Added " & strVern & " = " & strLemma & "_" & strPos)
          Else
            ' Try variant without the ending
            dtrFound = tdlMorphDict.Tables("VernPos").Select("Vern='" & strVern.Replace("'", "''") & "' AND Pos = '" & Left(strPos, intPos - 1) & "'")
            If (dtrFound.Length > 0) Then
              ' Get the lemma
              strLemma = dtrFound(0).Item("l").ToString
              ' Add this to the VernPos table
              If (Not MorphAddOneVernPos(strVern, strLemma, strPos, strFeat, "Derived: Pos without ^ [" & Format(Now, "G") & "]")) Then Return False
              bChanged = True : intChanges += 1
              Logging("Added " & strVern & " = " & strLemma & "_" & strPos)
            End If
          End If
        End If
        If (dtrFound.Length = 0) Then
          ' Check for RP+
          intPos = InStr(strPos, "RP+")
          If (intPos = 1) Then
            ' Stop
          End If
          If (dtrFound.Length = 0) Then
            ' Check if there is a VAG-VBN correlation
            If (strPos Like "VBN*") Then
              ' Try with VAG
              dtrFound = tdlMorphDict.Tables("VernPos").Select("Vern='" & strVern.Replace("'", "''") & "' AND Pos LIKE 'VAG*'")
              If (dtrFound.Length > 0) Then
                ' Get the lemma
                strLemma = dtrFound(0).Item("l").ToString
                ' Add this to the VernPos table
                If (Not MorphAddOneVernPos(strVern, strLemma, strPos, strFeat, "Derived: Pos VBN>VAG [" & Format(Now, "G") & "]")) Then Return False
                bChanged = True : intChanges += 1
                Logging("Added " & strVern & " = " & strLemma & "_" & strPos)
              End If
            ElseIf (strPos Like "VAG*") Then
              dtrFound = tdlMorphDict.Tables("VernPos").Select("Vern='" & strVern.Replace("'", "''") & "' AND Pos LIKE 'VBN*'")
              ' Try with VBN
              If (dtrFound.Length > 0) Then
                ' Get the lemma
                strLemma = dtrFound(0).Item("l").ToString
                ' Add this to the VernPos table
                If (Not MorphAddOneVernPos(strVern, strLemma, strPos, strFeat, "Derived: Pos VAG>VBN [" & Format(Now, "G") & "]")) Then Return False
                bChanged = True : intChanges += 1
                Logging("Added " & strVern & " = " & strLemma & "_" & strPos)
              End If
            End If
          End If
        End If
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/MorphTryLngVariants error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   OneMorphPropaVerb
  ' Goal:   Look for unlemmatized verbs in [strInFile] and try to lemmatize them,
  '           resulting in an updated [strInFile]
  ' History:
  ' 14-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function OneMorphPropaVerb(ByVal strInFile As String) As Boolean
    Dim ndxList As XmlNodeList        ' Result of select
    Dim ndxFor As XmlNode = Nothing   ' Forest node
    Dim ndxWork As XmlNode = Nothing  ' Working node
    Dim ndxThis As XmlNode = Nothing  ' Current node
    Dim ndxLeaf As XmlNode = Nothing  ' Leaf endnode
    Dim strShort As String = ""       ' Short file name (output)
    Dim strIn As String = ""          ' Short file name (input)
    Dim strVern As String = ""        ' Vernacular
    Dim strLemma As String = ""       ' Each entry should have a lemma value
    Dim strPos As String = ""         ' part of speech 
    Dim strLabel As String = ""       ' label of <eTree> parent
    Dim strFeat As String = ""        ' List of features
    Dim strH As String = ""           ' Kind of history addition
    Dim strChoice As String = ""      ' Choice we make
    Dim strGener As String = ""       ' Generalization type
    'Dim arFeat() As String            ' Array of features
    'Dim arLine() As String            ' Feature name and value
    Dim bChanged As Boolean = False   ' Any changes?
    Dim bConsider As Boolean = False  ' Need to consider this one or not?
    'Dim dtrFound() As DataRow         ' Result of SELECT
    Dim dtrNew As DataRow = Nothing   ' One datarow
    Dim intCount As Integer = 0       ' Counter
    Dim intChanges As Integer = 0     ' Number of changes
    Dim intPtc As Integer             ' Percentage
    Dim intI As Integer               ' Counter
    'Dim intJ As Integer               ' Counter
    'Dim intChoice As Integer          ' Choice made by user

    Try
      ' Validate
      If (Not IO.File.Exists(strInFile)) Then Return False
      ' Determine short filename
      strShort = IO.Path.GetFileNameWithoutExtension(strInFile)
      ' Try read this file into an XML structure
      If (ReadXmlDoc(strInFile, pdxFile)) Then
        ' Set current file value
        strCurrentFile = strInFile : pdxCurrentFile = pdxFile : strIn = IO.Path.GetFileNameWithoutExtension(strCurrentFile)
        ' Look for the first <forest> element
        If (Not GetFirstForest(pdxFile, ndxFor)) Then Return False
        ' Determine how many elements there are potentially
        If (ndxFor Is Nothing) Then
          intCount = 0
        ElseIf (ndxFor.ParentNode Is Nothing) Then
          intCount = 0
        Else
          intCount = ndxFor.ParentNode.ChildNodes.Count
        End If
        ' Other initialisation: a history note with the short date attached to it
        strH = "remVb_" & Format(Today, "d")
        ' Walk through all the <forest> elements FORWARD
        While (Not ndxFor Is Nothing)
          ' Show where we are (how do we KNOW where we are?)
          intPtc = (ndxFor.Attributes("forestId").Value) * 100 \ intCount
          Status("OE-verb-resolution [" & strIn & "] " & intPtc & "%", intPtc)
          ' Look for all <eTree> (affirmative) verbal nodes that have an <eLeaf> child and that do not yet have a feature-set of type "M"
          ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf)>0 and " & _
                                       "count(child::fs[@type='M'])=0 and tb:matches(@Label,'" & strVerbAffirm & "') " & _
                                       " and child::eLeaf[1][@Type='Vern']]", conTb)
          ' Walk all these <eTree> nodes and see if we can add features to them
          For intI = 0 To ndxList.Count - 1
            ' Check if we are not being interrupted
            If (bInterrupt) Then Return False
            Application.DoEvents()
            ' Get the <eTree> node and the <eLeaf> child
            ndxThis = ndxList(intI) : ndxLeaf = ndxThis.SelectSingleNode("./child::eLeaf") : ndxWork = ndxThis.ParentNode
            ' Get the <eTree> node's label
            strPos = ndxThis.Attributes("Label").Value
            ' Adapt the vernacular string so that it is compatible
            strVern = StringOEtoTagged(LCase(ndxLeaf.Attributes("Text").Value)).Replace("$", "")
            ' Only get the MAIN part of the parent <eTree> node, e.g: NP, IP, PP etc
            If (ndxWork IsNot Nothing) AndAlso (ndxWork.Name = "eTree") Then
              strLabel = LabelmainOE(ndxWork.Attributes("Label").Value)
            Else
              strLabel = ""
            End If
            ' Try to get a lemma for this particular verb
            If (MorphLemmaGet(ndxThis, strVern, strPos, strLemma, strFeat)) Then
              ' We have a matching lemma and perhaps some features: propagate them
              ' Propagate lemma
              If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "l", strLemma)) Then Return False
              ' Are there any features?
              If (strFeat <> "") Then
                ' Propagate features
                If (Not AddFeatureSet(ndxThis, "M", strFeat)) Then Return False
              End If
              ' Add the comment as feature "h"
              If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "h", strH)) Then Return False
              bChanged = True : intChanges += 1
            End If
          Next intI
          ' Go to next forest
          ndxFor = ndxFor.NextSibling
        End While
        ' Any changes?
        If (bChanged) Then
          ' Adapt the revision
          If (AddRevDesc(pdxCurrentFile, strUserName, Format(Now, "g"), "MorphPropagate Verb")) Then
            ' Write the result
            Status("Saving adapted file...")
            pdxCurrentFile.Save(strInFile)
          End If
        End If
        ' Return success
        Return True
      End If
      ' Could not read the file, so return failure
      Status("Could not read the XML file: " & strInFile)
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OneMorphPropaVerb error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   OneMorphPropaText
  ' Goal:   Look for unlemmatized items in [strInFile] and try to lemmatize them,
  '           resulting in an updated [strInFile]
  '         Ask the user for input where necessary
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function OneMorphPropaText(ByVal strInFile As String, ByRef bDataset As Boolean) As Boolean
    Dim ndxList As XmlNodeList        ' Result of select
    Dim ndxFor As XmlNode = Nothing   ' Forest node
    Dim ndxWork As XmlNode = Nothing  ' Working node
    Dim ndxThis As XmlNode = Nothing  ' Current node
    Dim ndxLeaf As XmlNode = Nothing  ' Leaf endnode
    Dim strShort As String = ""       ' Short file name (output)
    Dim strIn As String = ""          ' Short file name (input)
    Dim strVern As String = ""        ' Vernacular
    Dim strLemma As String = ""       ' Each entry should have a lemma value
    Dim strLev As String = "1"         ' Level
    Dim strPos As String = ""         ' part of speech 
    Dim strLabel As String = ""       ' label of <eTree> parent
    Dim strFeat As String = ""        ' List of features
    Dim strH As String = ""           ' Kind of history addition
    Dim strChoice As String = ""      ' Choice we make
    Dim strGener As String = ""       ' Generalization type
    Dim strType As String = ""        ' Type of addition to table [VernPos]
    'Dim arFeat() As String            ' Array of features
    'Dim arLine() As String            ' Feature name and value
    Dim bChanged As Boolean = False   ' Any changes?
    Dim bConsider As Boolean = False  ' Need to consider this one or not?
    'Dim dtrFound() As DataRow         ' Result of SELECT
    Dim dtrNew As DataRow = Nothing   ' One datarow
    Dim intCount As Integer = 0       ' Counter
    Dim intChanges As Integer = 0     ' Number of changes
    Dim intPtc As Integer             ' Percentage
    Dim intI As Integer               ' Counter
    'Dim intJ As Integer               ' Counter
    'Dim intChoice As Integer          ' Choice made by user

    Try
      ' Validate
      If (Not IO.File.Exists(strInFile)) Then Return False
      bDataset = False
      ' Determine short filename
      strShort = IO.Path.GetFileNameWithoutExtension(strInFile)
      ' Try read this file into an XML structure
      If (ReadXmlDoc(strInFile, pdxFile)) Then
        ' Set current file value
        strCurrentFile = strInFile : pdxCurrentFile = pdxFile : strIn = IO.Path.GetFileNameWithoutExtension(strCurrentFile)
        ' Look for the first <forest> element
        If (Not GetFirstForest(pdxFile, ndxFor)) Then Return False
        ' Determine how many elements there are potentially
        If (ndxFor Is Nothing) Then
          intCount = 0
        ElseIf (ndxFor.ParentNode Is Nothing) Then
          intCount = 0
        Else
          intCount = ndxFor.ParentNode.ChildNodes.Count
        End If
        ' Other initialisation: a history note with the short date attached to it
        strH = "remText_" & Format(Today, "d")
        ' Walk through all the <forest> elements FORWARD
        While (Not ndxFor Is Nothing)
          If (bInterrupt) Then Return False
          ' Show where we are (how do we KNOW where we are?)
          intPtc = (ndxFor.Attributes("forestId").Value) * 100 \ intCount
          Status("OE-lemmatization [" & strIn & "] " & intPtc & "%", intPtc)
          ' Look for all <eTree> (affirmative) verbal nodes that have an <eLeaf> child and that do not yet have a feature-set of type "M"
          ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf)>0 and " & _
                                       "count(child::fs[@type='M'])=0 and not(tb:matches(@Label,'CODE|.|,|FW')) " & _
                                       " and child::eLeaf[1][@Type='Vern']]", conTb)
          ' Walk all these <eTree> nodes and see if we can add features to them
          For intI = 0 To ndxList.Count - 1
            ' Check if we are not being interrupted
            If (bInterrupt) Then Return False
            Application.DoEvents()
            ' Get the <eTree> node and the <eLeaf> child
            ndxThis = ndxList(intI) : ndxLeaf = ndxThis.SelectSingleNode("./child::eLeaf") : ndxWork = ndxThis.ParentNode
            ' Get the <eTree> node's label
            strPos = ndxThis.Attributes("Label").Value
            ' Adapt the vernacular string so that it is compatible
            strVern = StringOEtoTagged(LCase(ndxLeaf.Attributes("Text").Value)).Replace("$", "")
            ' Only get the MAIN part of the parent <eTree> node, e.g: NP, IP, PP etc
            If (ndxWork IsNot Nothing) AndAlso (ndxWork.Name = "eTree") Then
              strLabel = LabelmainOE(ndxWork.Attributes("Label").Value)
            Else
              strLabel = ""
            End If
            ' Try to get a lemma for this particular item
            If (MorphLemmaGet(ndxThis, strVern, strPos, strLemma, strFeat)) Then
              ' We have a matching lemma and perhaps some features: propagate them
              ' Propagate lemma
              If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "l", strLemma)) Then Return False
              ' Are there any features?
              If (strFeat <> "") Then
                ' Propagate features
                If (Not AddFeatureSet(ndxThis, "M", strFeat)) Then Return False
                strType = "LemmaFeat"
              Else
                strType = "LemmaOnly"
              End If
              ' Add the comment as feature "h"
              If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "h", strH)) Then Return False
              ' Keep track of the changes we have
              bChanged = True : intChanges += 1
              ' Add the Vern/Pos/Lemma/Feat into the table [VernPos]
              If (Not VernPosAdd(strPos, strVern, strType, strLemma, strLev, strFeat)) Then Return False
              bDataset = True
            End If
          Next intI
          ' Go to next forest
          ndxFor = ndxFor.NextSibling
        End While
        ' Any changes?
        If (bChanged) Then
          ' Adapt the revision
          If (AddRevDesc(pdxCurrentFile, strUserName, Format(Now, "g"), "MorphPropagate Verb")) Then
            ' Write the result
            Status("Saving adapted file...")
            pdxCurrentFile.Save(strInFile)
          End If
        End If
        ' Return success
        Return True
      End If
      ' Could not read the file, so return failure
      Status("Could not read the XML file: " & strInFile)
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OneMorphPropaText error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   OneMorphPropaFile
  ' Goal:   Propagate morphology information from [tdlMorphDict] into this [strInFile]
  '           resulting in an updated [strOutFile]
  ' History:
  ' 19-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function OneMorphPropaFile(ByVal strInFile As String, ByVal strOutFile As String, ByVal bDoAsk As Boolean) As Boolean
    Dim ndxList As XmlNodeList        ' Result of select
    Dim ndxFor As XmlNode = Nothing   ' Forest node
    Dim ndxWork As XmlNode = Nothing  ' Working node
    Dim ndxThis As XmlNode = Nothing  ' Current node
    Dim ndxLeaf As XmlNode = Nothing  ' Leaf endnode
    Dim strShort As String = ""       ' Short file name (output)
    Dim strIn As String = ""          ' Short file name (input)
    Dim strVern As String = ""        ' Vernacular
    Dim strLemma As String = ""       ' Each entry should have a lemma value
    Dim strPos As String = ""         ' part of speech 
    Dim strLabel As String = ""       ' label of <eTree> parent
    Dim strFeat As String = ""        ' List of features
    Dim strH As String = ""           ' Kind of history addition
    Dim strChoice As String = ""      ' Choice we make
    Dim strGener As String = ""       ' Generalization type
    Dim arFeat() As String            ' Array of features
    Dim arLine() As String            ' Feature name and value
    Dim bChanged As Boolean = False   ' Any changes?
    Dim bConsider As Boolean = False  ' Need to consider this one or not?
    Dim dtrFound() As DataRow         ' Result of SELECT
    Dim dtrNew As DataRow = Nothing   ' One datarow
    Dim intCount As Integer = 0       ' Counter
    Dim intChanges As Integer = 0     ' Number of changes
    Dim intPtc As Integer             ' Percentage
    Dim intI As Integer               ' Counter
    Dim intJ As Integer               ' Counter
    Dim intChoice As Integer          ' Choice made by user

    Try
      ' Validate
      If (Not IO.File.Exists(strInFile)) Then Return False
      ' Determine short filename
      strShort = IO.Path.GetFileNameWithoutExtension(strInFile)
      ' Try read this file into an XML structure
      If (ReadXmlDoc(strInFile, pdxFile)) Then
        ' Set current file value
        strCurrentFile = strInFile : pdxCurrentFile = pdxFile : strIn = IO.Path.GetFileNameWithoutExtension(strCurrentFile)
        ' Look for the first <forest> element
        If (Not GetFirstForest(pdxFile, ndxFor)) Then Return False
        ' Determine how many elements there are potentially
        If (ndxFor Is Nothing) Then
          intCount = 0
        ElseIf (ndxFor.ParentNode Is Nothing) Then
          intCount = 0
        Else
          intCount = ndxFor.ParentNode.ChildNodes.Count
        End If
        ' Other initialisation: a history note with the short date attached to it
        strH = "propa_" & Format(Today, "d")
        ' Walk through all the <forest> elements FORWARD
        While (Not ndxFor Is Nothing)
          ' Show where we are (how do we KNOW where we are?)
          intPtc = (ndxFor.Attributes("forestId").Value) * 100 \ intCount
          Status("OE-morph-propagation [" & strIn & "] " & intPtc & "%", intPtc)
          ' Look for all <eTree> nodes that have an <eLeaf> child and that do not yet have a feature-set of type "M"
          ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf)>0 and " & _
                                       "count(child::fs[@type='M'])=0 and not(tb:matches(@Label,'CODE|.|,|FW')) " & _
                                       " and child::eLeaf[1][@Type='Vern']]", conTb)
          ' Walk all these <eTree> nodes and see if we can add features to them
          For intI = 0 To ndxList.Count - 1
            ' Check if we are not being interrupted
            If (bInterrupt) Then Return False
            Application.DoEvents()
            ' Get the <eTree> node, the <eLeaf> child, and the parent of the <eTree> node
            ndxThis = ndxList(intI) : ndxWork = ndxThis.ParentNode : ndxLeaf = ndxThis.SelectSingleNode("./child::eLeaf")
            ' Get the <eTree> node's label
            strPos = ndxThis.Attributes("Label").Value
            ' Adapt the vernacular string so that it is compatible
            strVern = StringOEtoTagged(LCase(ndxLeaf.Attributes("Text").Value)).Replace("$", "")
            ' Only get the MAIN part of the parent <eTree> node, e.g: NP, IP, PP etc
            If (ndxWork IsNot Nothing) AndAlso (ndxWork.Name = "eTree") Then
              strLabel = LabelmainOE(ndxWork.Attributes("Label").Value)
            Else
              strLabel = "*"
            End If
            strFeat = ""
            ' Check if this is a forbidden combination
            dtrFound = tdlMorphDict.Tables("Check").Select("Vern='" & strVern.Replace("'", "''") & "' AND Pos='" & strPos.Replace("'", "''") & _
                       "' AND PrntLabel='" & strLabel.Replace("'", "''") & "'")
            bConsider = True
            If (dtrFound.Length > 0) Then
              ' It is in there, so have a look at it
              Select Case dtrFound(0).Item("Status")
                Case "Bad"
                  bConsider = False
              End Select
            End If
            ' Are we going to consider it?
            If (bConsider) Then
              ' See how many matches we find for this combination
              '  (Get them in decreasing frequency)
              dtrFound = tdlMorphDict.Tables("Morph").Select("Vern='" & strVern.Replace("'", "''") & "' AND Pos='" & strPos.Replace("'", "''") & _
                         "' AND Label LIKE '" & strLabel.Replace("'", "''") & "'", "Freq DESC")
              Select Case dtrFound.Length
                Case 0
                  ' There is no match, so we widen it to the combination Vern/POS
                  dtrFound = tdlMorphDict.Tables("Morph").Select("Vern='" & strVern.Replace("'", "''") & "' AND Pos='" & strPos.Replace("'", "''") & _
                         "'", "Freq DESC")
                  Select Case dtrFound.Length
                    Case 0
                      ' There is no match, so skip it for now!
                    Case Else
                      ' Unique match --> Check if this needs automatic propagation
                      Select Case MorphDictGenType(strPos, strLabel)
                        Case "full"   ' Propagate lemma and features
                          ' Full generalization: take lemma + features
                          AddSemiStack(strFeat, dtrFound(0).Item("l"))
                          AddSemiStack(strFeat, dtrFound(0).Item("f"))
                          ' Add the combination in the <Morph> table
                          If (Not MorphDictAdd(strVern, strPos, strLabel, dtrFound(0).Item("l"), _
                                  dtrFound(0).Item("f") & ";h=full")) Then Logging("OneMorphPropaFile problem #1") : Return False
                        Case "restr"  ' Propagate lemma alone
                          ' Full generalization: take lemma + features
                          AddSemiStack(strFeat, dtrFound(0).Item("l"))
                          ' Add the combination in the <Morph> table
                          If (Not MorphDictAdd(strVern, strPos, strLabel, dtrFound(0).Item("l"), _
                                  dtrFound(0).Item("f") & "h=restr")) Then Logging("OneMorphPropaFile problem #2") : Return False
                        Case Else
                          ' Unique match, so process it here: ask user permission!!
                          intChoice = MorphVernPosParentAsk(strVern, strPos, strLabel, dtrFound, ndxThis, strChoice, strGener)
                          If (intChoice >= 0) Then
                            ' Depends on generalization 
                            Select Case strGener
                              Case "full"
                                ' Full generalization: take lemma + features
                                AddSemiStack(strFeat, dtrFound(intChoice).Item("l"))
                                AddSemiStack(strFeat, dtrFound(intChoice).Item("f"))
                                ' Add full generalization in the <Gen> table entries
                                If (Not MorphDictGenAdd(strPos, strLabel, "full")) Then Return False
                                ' Add the combination in the <Morph> table
                                If (Not MorphDictAdd(strVern, strPos, strLabel, dtrFound(intChoice).Item("l"), _
                                        dtrFound(intChoice).Item("f") & ";h=full")) Then Logging("OneMorphPropaFile problem #3") : Return False
                              Case "restr"
                                ' Restricted generalization: take lemma only
                                AddSemiStack(strFeat, dtrFound(intChoice).Item("l"))
                                ' Add restricted generalization in the <Gen> table entries
                                If (Not MorphDictGenAdd(strPos, strLabel, "restr")) Then Return False
                                ' Add the combination in the <Morph> table
                                If (Not MorphDictAdd(strVern, strPos, strLabel, dtrFound(intChoice).Item("l"), _
                                        "h=restr")) Then Logging("OneMorphPropaFile problem #4") : Return False
                              Case Else
                                ' User agreed, so continue
                                AddSemiStack(strFeat, dtrFound(intChoice).Item("l"))
                                AddSemiStack(strFeat, strChoice)
                                ' Add the combination in the <Morph> table
                                If (Not MorphDictAdd(strVern, strPos, strLabel, dtrFound(intChoice).Item("l"), _
                                        strChoice)) Then Logging("OneMorphPropaFile problem #5") : Return False
                            End Select
                            bChanged = True : intChanges += 1
                          End If
                      End Select
                  End Select
                Case 1
                  ' There is a unique match, so process the features...
                  AddSemiStack(strFeat, dtrFound(0).Item("l"))
                  AddSemiStack(strFeat, dtrFound(0).Item("f"))
                  bChanged = True : intChanges += 1
                Case Else
                  ' There are more matches, and we will have to consider the best one with respect to the remainder
                  If (bDoAsk) Then
                    ' Ask the user for a decision
                    With frmMorph
                      ' Pass on sentence
                      .Sentence = ClauseToHtml(ndxThis)
                      ' Pass on array of options
                      .SetChoices(dtrFound)
                      ' Ask user
                      Select Case .ShowDialog
                        Case DialogResult.OK
                          ' Select this one
                          intChoice = .Choice
                          'strFeat = dtrFound(intChoice).Item("f")
                          AddSemiStack(strFeat, dtrFound(intChoice).Item("l"))
                          AddSemiStack(strFeat, dtrFound(intChoice).Item("f"))
                          bChanged = True : intChanges += 1
                        Case DialogResult.Cancel
                          ' Quit
                          bInterrupt = True
                          Return False
                      End Select
                    End With
                  End If
              End Select
              ' Do we have any features?
              If (strFeat <> "") Then
                ' Convert the features into an array
                arFeat = Split(strFeat, ";")
                ' Expect the lemma on place 0
                If (arFeat.Length > 0) Then
                  ' Process lemma
                  ' Add feature name and value
                  If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "l", arFeat(0))) Then Status("OneMorphPropaFile: Problem adding feature") : Return False
                  bChanged = True : intChanges += 1
                  ' Do the other features
                  For intJ = 1 To arFeat.Length - 1
                    ' Validate
                    If (arFeat(intJ) <> "") Then
                      ' Get feature name and value
                      arLine = Split(arFeat(intJ), "=")
                      If (arLine.Length = 2) Then
                        ' Add feature name and value
                        If (Not AddFeature(pdxCurrentFile, ndxThis, "M", arLine(0), arLine(1))) Then Status("OneMorphPropaFile: Problem adding feature") : Return False
                        intChanges += 1
                      End If
                    End If
                  Next intJ
                  ' Add one history feature, so that we can recognize "propagated" M features
                  If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "h", strH)) Then Status("OneMorphPropaFile: problem adding feature @h") : Return False
                  ' ======== DEBUGGING =============
                  ' Debug.Print("MorphPropa [" & strIn & "]:" & ndxThis.Attributes("Id").Value)
                  ' ================================
                End If  ' arFeat.Length > 0
              End If    ' strFeat <> ""
            End If      ' bConsider = True
          Next intI
          ' Go to next forest
          ndxFor = ndxFor.NextSibling
        End While
        ' Any changes?
        If (bChanged) Then
          ' Adapt the revision
          If (AddRevDesc(pdxCurrentFile, strUserName, Format(Now, "g"), "MorphPropagation")) Then
            ' Write the result
            Status("Saving adapted file...")
            pdxCurrentFile.Save(strInFile)
          End If
        End If
        ' Return success
        Return True
      End If
      ' Could not read the file, so return failure
      Status("Could not read the XML file: " & strInFile)
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OneMorphPropaFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   OneMorphCorrectFile
  ' Goal:   Correct morphology information from in the [strInFile]
  ' History:
  ' 21-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function OneMorphCorrectFile(ByVal strInFile As String, ByRef intChanges As Integer) As Boolean
    Dim ndxList As XmlNodeList        ' Result of select
    Dim ndxFor As XmlNode = Nothing   ' Forest node
    Dim ndxWork As XmlNode = Nothing  ' Working node
    Dim ndxThis As XmlNode = Nothing  ' Current node
    Dim strShort As String = ""       ' Short file name (output)
    Dim strIn As String = ""          ' Short file name (input)
    Dim strVern As String = ""        ' Vernacular
    Dim strLemma As String = ""       ' Each entry should have a lemma value
    Dim strPos As String = ""         ' part of speech 
    Dim strPOSmain As String = ""     ' Main part of POS label
    Dim strCheck As String = ""       ' Resulting pOS of checking
    Dim strLabel As String = ""       ' label of <eTree> parent
    Dim strFeat As String = ""        ' List of features
    Dim strCombi As String = ""       ' Combination
    Dim strH As String = ""           ' Kind of history addition
    Dim strMWcat As String = ""       ' MWcategory (POS in OE_tagged)
    Dim strChoice As String = ""      ' The choice we make
    Dim bChanged As Boolean = False   ' Any changes?
    Dim dtrNew As DataRow = Nothing   ' One datarow
    Dim dtrFound() As DataRow         ' Result of select
    Dim intCount As Integer = 0       ' Counter
    Dim intPtc As Integer             ' Percentage
    Dim intI As Integer               ' Counter
    Dim intJ As Integer               ' Counter
    Dim intId As Integer              ' Id of current <eTree> element
    Dim intPos As Integer             ' Position within string
    Dim intForId As Integer = 0       ' Id of current <forest>
    Dim bProcess As Boolean           ' Need to process this one or not?

    Try
      ' Validate
      If (Not IO.File.Exists(strInFile)) Then Return False
      ' Determine short filename
      strShort = IO.Path.GetFileNameWithoutExtension(strInFile)
      intChanges = 0
      ' Try read this file into an XML structure
      If (ReadXmlDoc(strInFile, pdxFile)) Then
        ' Set current file value
        strCurrentFile = strInFile : pdxCurrentFile = pdxFile : strIn = IO.Path.GetFileNameWithoutExtension(strCurrentFile)
        ' Look for the first <forest> element
        If (Not GetFirstForest(pdxFile, ndxFor)) Then Return False
        ' Determine how many elements there are potentially
        If (ndxFor Is Nothing) Then
          intCount = 0
        ElseIf (ndxFor.ParentNode Is Nothing) Then
          intCount = 0
        Else
          intCount = ndxFor.ParentNode.ChildNodes.Count
        End If
        ' Other initialisation: a history note with the short date attached to it
        strH = "mrp-corr_" & Format(Today, "d")
        ' Walk through all the <forest> elements FORWARD
        While (Not ndxFor Is Nothing)
          ' Show where we are (how do we KNOW where we are?)
          intPtc = (ndxFor.Attributes("forestId").Value) * 100 \ intCount
          Status("OE-morph-correction [" & strIn & "] " & intPtc & "%", intPtc)
          ' Look for all <eTree> nodes that have an <eLeaf> child and that have any M features
          ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf)>0 and " & _
                                       "count(child::fs[@type='M'])>0 and not(tb:matches(@Label,'CODE|.|,|FW'))" & _
                                       " and child::eLeaf[1][@Type='Vern']]", conTb)
          ' Walk all these <eTree> nodes and see if we need to correct features
          For intI = 0 To ndxList.Count - 1
            ' Check if we are not being interrupted
            If (bInterrupt) Then Return False
            Application.DoEvents()
            ' Where are we?
            intForId = ndxFor.Attributes("forestId").Value
            ' Get the <eTree> node and the parent of the <eTree> node
            ndxThis = ndxList(intI) : ndxWork = ndxThis.ParentNode : intId = ndxThis.Attributes("Id").Value
            ' Get the <eTree> node's label (the part-of-speech)
            strPos = ndxThis.Attributes("Label").Value
            strPOSmain = LabelmainOE(strPos, "-^")
            intPos = InStr(strPos, "^")
            ' ============== TEMPORARILY NEEDED =====================
            ' Possibly correct "lemma" into "l"
            strFeat = GetFeature(ndxThis, "M", "lemma")
            If (strFeat <> "") Then
              ' Change the feature
              With ndxThis.SelectSingleNode("./child::fs/child::f[@name='lemma']")
                .Attributes("name").Value = "l" : bChanged = True : intChanges += 1
              End With
            End If
            ' =======================================================
            ' Get the lemma - that is needed anyway
            strVern = ndxThis.SelectSingleNode("./child::eLeaf[1]").Attributes("Text").Value
            strVern = StringOEtoTagged(LCase(strVern).Replace("$", ""))
            strLemma = GetFeature(ndxThis, "M", "l") : strMWcat = ""
            ' Check if we already noted that some action had to be taken here
            dtrFound = tdlMorphDict.Tables("Check").Select("Vern='" & strVern.Replace("'", "''") & "' AND " & _
                                                           "Lemma='" & strLemma.Replace("'", "''") & "' AND " & _
                                                           "POS LIKE '" & strPOSmain & "'")
            If (dtrFound.Length > 0) Then
              ' Act on the action stated here
              Select Case dtrFound(0).Item("Status")
                Case "Bad"
                  ' Still need to ask the user
                  bProcess = True
                Case "Okay"
                  ' No need to do anything here!
                  bProcess = False
                Case "ChangeToLemmaDictPOS"
                  ' Change the POS to that of the lemma dictionary
                  ' (But this makes no sense!! we are not going to change the YCOE...)
                  Stop
                  bProcess = False
                Case "DeleteLemmaVernPOS", "DeleteLemmaVern"
                  ' Delete the lemma and the features
                  If (Not DeleteTaggedFeatures(ndxThis)) Then Status("Could not delete tagged features") : Return False
                  bChanged = True : intChanges += 1
                  bProcess = False
                Case "ChangeLemmaFeat"
                  ' Delete old features
                  If (Not DeleteTaggedFeatures(ndxThis)) Then Status("Could not delete tagged features") : Return False
                  ' Change the lemma and the features as in the [Correct] value
                  strLemma = dtrFound(0).Item("Correct")
                  If (InStr(strLemma, ";") > 0) Then
                    strFeat = Mid(strLemma, InStr(strLemma, ";") + 1)
                    strLemma = Left(strLemma, InStr(strLemma, ";") - 1)
                  Else
                    strFeat = ""
                  End If
                  If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "l", strLemma)) Then Status("OneMorphCorrectFile: lemma") : Return False
                  If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "h", strH)) Then Status("OneMorphCorrectFile: help info") : Return False
                  If (strFeat <> "") Then
                    If (Not AddFeatureSet(ndxThis, "M", strFeat)) Then Status("OneMorphCorrectFile: feat") : Return False
                  End If
                  bChanged = True : intChanges += 1
                  bProcess = False
                Case Else
                  Stop
              End Select
            Else
              bProcess = False
            End If
            ' Do we need to process?
            If (bProcess) Then
              ' Get the features
              strFeat = GetMfeatures(ndxThis)
              ' Compare the lemma + POS 
              If (strLemma <> "") Then
                ' Check whether the POS is correct
                strCheck = CheckMorphPOS(strLemma, strPOSmain, strMWcat)
                If (strCheck <> "") AndAlso (strCheck <> strPOSmain) Then
                  ' We need to ask
                  strChoice = MorphCheck(ndxThis, strLemma, strPOSmain, strCheck, strMWcat, strFeat)
                  Select Case strChoice
                    Case "Cancel"
                      ' Ask if user wants to quit entirely
                      If (MsgBox("Would you like to quit entirely?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes) Then
                        bInterrupt = True
                      End If
                      Return False
                    Case "AllowPOS"
                      ' YCOE is correct, so add this combination into the <Lemma> section of the morphological dictionary
                      MorphLemmaAdd(strLemma, strPOSmain, strMWcat)
                      bChanged = True : intChanges += 1
                    Case "ChangeToLemmaDictPOS"
                      ' All occurrences of Vern/Lemma should have POS as in [strCheck]
                      dtrNew = AddOneDataRow(tdlMorphDict, "Check", "CheckId", "CheckList")
                      dtrNew.Item("Vern") = strVern : dtrNew.Item("Lemma") = strLemma : dtrNew.Item("POS") = strPOSmain
                      dtrNew.Item("Status") = "ChangeToLemmaDictPOS" : dtrNew.Item("Correct") = strCheck
                      bChanged = True : intChanges += 1
                    Case "DeleteLemmaVernPOS"
                      ' Always delete this @lemma (plus @features) for combination Vern/POS
                      dtrNew = AddOneDataRow(tdlMorphDict, "Check", "CheckId", "CheckList")
                      dtrNew.Item("Vern") = strVern : dtrNew.Item("Lemma") = strLemma : dtrNew.Item("POS") = strPOSmain
                      dtrNew.Item("Status") = "DeleteLemmaVernPOS"
                      ' Delete this combination from the <Morph> table
                      dtrFound = tdlMorphDict.Tables("Morph").Select("l='" & strLemma.Replace("'", "''") & "' AND Vern='" & strVern.Replace("'", "''") & "' AND Pos='" & strPos & "'")
                      If (dtrFound.Length = 0) Then Stop Else dtrFound(0).Delete()
                      ' dtrFound(0).Delete()
                      bChanged = True : intChanges += 1
                    Case "DeleteLemmaVern"
                      ' Always delete this @lemma (plus @features) for occurrences of Vern, no matter what POS
                      dtrNew = AddOneDataRow(tdlMorphDict, "Check", "CheckId", "CheckList")
                      dtrNew.Item("Vern") = strVern : dtrNew.Item("Lemma") = strLemma : dtrNew.Item("POS") = "*"
                      dtrNew.Item("Status") = "DeleteLemmaVern"
                      ' Delete all these combinations from the <Morph> table
                      dtrFound = tdlMorphDict.Tables("Morph").Select("l='" & strLemma.Replace("'", "''") & "' AND Vern='" & strVern.Replace("'", "''") & "'")
                      If (dtrFound.Length = 0) Then Stop
                      For intJ = dtrFound.Length - 1 To 0 Step -1
                        dtrFound(intJ).Delete()
                      Next intJ
                      bChanged = True : intChanges += 1
                    Case Else
                      ' Check if this is a number
                      If (IsNumeric(strChoice)) Then
                        ' The user has selected an entry from the table <Morph>
                        dtrFound = tdlMorphDict.Tables("Morph").Select("MorphId=" & strChoice)
                        If (dtrFound.Length = 0) Then Stop
                        strCombi = ""
                        AddSemiStack(strCombi, "l=" & dtrFound(0).Item("l"))
                        AddSemiStack(strCombi, dtrFound(0).Item("f"))
                        ' De we need to change something to or add something to the <check>?
                        If (Not CheckMorphAdd(strVern, strPOSmain, strLemma, "ChangeLemmaFeat", strCombi)) Then Stop
                        bChanged = True : intChanges += 1
                        ' Change lemma and feat values
                        strLemma = dtrFound(0).Item("l")
                        strFeat = dtrFound(0).Item("f")
                        ' Check if an additional entry in <Morph> is needed
                        dtrFound = tdlMorphDict.Tables("Morph").Select("Vern='" & strVern.Replace("'", "''") & "' AND Pos='" & strPos & _
                                                                       "' AND Label='" & LabelmainOE(ndxThis.ParentNode.Attributes("Label").Value) & "'")
                        If (dtrFound.Length = 0) Then
                          ' Add this entry
                          dtrNew = AddOneDataRow(tdlMorphDict, "Morph", "MorphId", "MorphList")
                          With dtrNew
                            .Item("Vern") = strVern : .Item("Pos") = strPos : .Item("Label") = LabelmainOE(ndxThis.ParentNode.Attributes("Label").Value)
                            .Item("l") = strLemma : .Item("f") = strFeat : .Item("File") = strIn
                            .Item("forestId") = ndxThis.Attributes("forestId").Value
                            .Item("EtreeId") = ndxThis.Attributes("Id").Value
                            .Item("Freq") = 1
                          End With
                        End If
                      Else
                        ' This option is not known
                        Stop
                      End If
                  End Select
                End If

                'Select Case MsgBox("forest[" & intForId & "/" & intId & "] has:" & vbCrLf & _
                '                   "vern=" & vbTab & "[" & VernToEnglish(strVern) & "] " & vbCrLf & _
                '                   "lemma=" & vbTab & "[" & strLemma & "] " & vbCrLf & _
                '                   "Ycoe POS=" & vbTab & strPos & vbCrLf & _
                '                   "Ycoe f=" & vbTab & GetMfeatures(ndxThis) & vbCrLf & _
                '                   "Tagged POS=" & vbTab & strCheck & "/" & strMWcat & vbCrLf & _
                '                   "Is YCOE correct?", MsgBoxStyle.YesNoCancel)
                '  Case MsgBoxResult.Cancel
                '    bInterrupt = True : Return False
                '  Case MsgBoxResult.Yes
                '    ' There is a mismatch!
                '    'Stop
                '    ' YCOE is correct, so add this combination into the morphological dictionary
                '    MorphLemmaAdd(strLemma, strPOSmain, strMWcat)
                '  Case MsgBoxResult.No
                '    ' There is a mismatch!
                '    'Stop
                '    ' Delete the lemma and the other features
                '    If (Not DeleteTaggedFeatures(ndxThis)) Then Status("Could not delete tagged features") : Return False
                '    bChanged = True : intChanges += 1
                'End Select
              End If
            End If
            ' get the case feature value
            strFeat = GetFeature(ndxThis, "M", "case")
            ' Do we have a case feature?
            If (strFeat <> "") Then
              ' is this a verb?
              If (DoLike(strPos, "V*|B*|H*|M*|AX*")) AndAlso (intPos = 0) Then
                ' It should not have a case feature -- delete all features except the lemma
                If (Not DeleteTaggedFeatures(ndxThis)) Then Status("Could not delete tagged features") : Return False
                bChanged = True : intChanges += 1
                ' Check if lemma is correct
                Select Case CheckMorphStatus(strVern, strPos, strLemma)
                  Case "Okay" ' Add the lemma
                    ' Add the lemma again 
                    If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "l", strLemma)) Then Status("OneMorphCorrectFile: lemma") : Return False
                    If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "h", strH)) Then Status("OneMorphCorrectFile: help info") : Return False
                  Case "Bad"  ' No need to do anything!
                  Case Else
                    ' We need to ask
                    Select Case MsgBox("forest[" & intForId & "/" & intId & "] has:" & vbCrLf & _
                                       "vern=" & vbTab & "[" & strVern & "] " & vbCrLf & _
                                       "lemma=" & vbTab & "[" & strLemma & "] " & vbCrLf & _
                                       "POS=" & vbTab & strPos & vbCrLf & _
                                       "Is that correct?", MsgBoxStyle.YesNoCancel)
                      Case MsgBoxResult.Cancel
                        bInterrupt = True : Return False
                      Case MsgBoxResult.Yes
                        ' Add the lemma again 
                        If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "l", strLemma)) Then Status("OneMorphCorrectFile: lemma") : Return False
                        If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "h", strH)) Then Status("OneMorphCorrectFile: help info") : Return False
                        ' Indicate this one is good
                        If (Not CheckMorphAdd(strVern, strPos, strLemma, "Okay")) Then Return False
                      Case MsgBoxResult.No
                        ' Indicate this one is bad
                        If (Not CheckMorphAdd(strVern, strPos, strLemma, "Bad")) Then Return False
                    End Select
                End Select
              ElseIf DoLike(strPos, "CONJ|P") Then
                ' These labels may not have any features
                If (ndxThis.SelectSingleNode("./child::fs[@type='M']/child::f[not(name='l')]") IsNot Nothing) Then
                  ' It should not have a case feature -- delete all features except the lemma
                  strLemma = GetFeature(ndxThis, "M", "l")
                  If (Not DeleteTaggedFeatures(ndxThis)) Then Status("Could not delete tagged features") : Return False
                  strVern = ndxThis.SelectSingleNode("./child::eLeaf[1]").Attributes("Text").Value
                  bChanged = True : intChanges += 1
                  ' Check if lemma is correct
                  Select Case CheckMorphStatus(strVern, strPos, strLemma)
                    Case "Okay" ' Add the lemma
                      ' Add the lemma again 
                      If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "l", strLemma)) Then Status("OneMorphCorrectFile: lemma") : Return False
                      If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "h", strH)) Then Status("OneMorphCorrectFile: help info") : Return False
                    Case "Bad"  ' No need to do anything!
                    Case Else
                      ' We need to ask
                      Select Case MsgBox("forest[" & intForId & "/" & intId & "] has:" & vbCrLf & _
                                         "vern=" & vbTab & "[" & strVern & "] " & vbCrLf & _
                                         "lemma=" & vbTab & "[" & strLemma & "] " & vbCrLf & _
                                         "POS=" & vbTab & strPos & vbCrLf & _
                                         "Is that correct?", MsgBoxStyle.YesNoCancel)
                        Case MsgBoxResult.Cancel
                          bInterrupt = True : Return False
                        Case MsgBoxResult.Yes
                          ' Add the lemma again 
                          If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "l", strLemma)) Then Status("OneMorphCorrectFile: lemma") : Return False
                          If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "h", strH)) Then Status("OneMorphCorrectFile: help info") : Return False
                          ' Indicate this one is good
                          If (Not CheckMorphAdd(strVern, strPos, strLemma, "Okay")) Then Return False
                        Case MsgBoxResult.No
                          ' Indicate this one is bad
                          If (Not CheckMorphAdd(strVern, strPos, strLemma, "Bad")) Then Return False
                      End Select
                  End Select
                End If
              Else
                ' It is probably allowed to have a case feature -- check the case
                If (intPos > 0) Then
                  ' Get the case part
                  strPos = Mid(strPos, intPos + 1)
                  ' Look at this (ought to be correct) features
                  Select Case strPos
                    Case "N"  ' Nominative
                      If (strFeat <> "nm") Then
                        If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "case", "nm")) Then Status("OneMorphCorrectFile: nominative case") : Return False
                        If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "h", strH)) Then Status("OneMorphCorrectFile: help info") : Return False
                        bChanged = True : intChanges += 1
                      End If
                    Case "D"  ' Dative
                      If (strFeat <> "dt") Then
                        If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "case", "dt")) Then Status("OneMorphCorrectFile: dative case") : Return False
                        If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "h", strH)) Then Status("OneMorphCorrectFile: help info") : Return False
                        bChanged = True : intChanges += 1
                      End If
                    Case "A"  ' Accusative
                      If (strFeat <> "ac") Then
                        If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "case", "ac")) Then Status("OneMorphCorrectFile: accusative case") : Return False
                        If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "h", strH)) Then Status("OneMorphCorrectFile: help info") : Return False
                        bChanged = True : intChanges += 1
                      End If
                    Case "I"  ' Instrumental
                      If (strFeat <> "is") Then
                        If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "case", "is")) Then Status("OneMorphCorrectFile: instrumental case") : Return False
                        If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "h", strH)) Then Status("OneMorphCorrectFile: help info") : Return False
                        bChanged = True : intChanges += 1
                      End If
                    Case "G"  ' Genitive
                      If (strFeat <> "gn") Then
                        If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "case", "gn")) Then Status("OneMorphCorrectFile: genitive case") : Return False
                        If (Not AddFeature(pdxCurrentFile, ndxThis, "M", "h", strH)) Then Status("OneMorphCorrectFile: help info") : Return False
                        bChanged = True : intChanges += 1
                      End If
                  End Select
                End If
              End If
            End If
          Next intI
          ' Go to next forest
          ndxFor = ndxFor.NextSibling
        End While
        ' Any changes?
        If (bChanged) Then
          ' Adapt the revision
          If (AddRevDesc(pdxCurrentFile, strUserName, Format(Now, "g"), "MorphCorrection")) Then
            ' Write the result
            Status("Saving adapted file (" & intChanges & " changes)...")
            pdxCurrentFile.Save(strInFile)
          End If
        End If
        ' Return success
        Return True
      End If
      ' Could not read the file, so return failure
      Status("Could not read the XML file: " & strInFile)
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OneMorphCorrectFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ClauseToHtml
  ' Goal:   Make a report of what we are to return
  ' History:
  ' 19-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function ClauseToHtml(ByRef ndxThis As XmlNode) As String
    Dim colHtml As New StringColl ' What we return
    Dim ndxFor As XmlNode         ' Forest
    Dim ndxList As XmlNodeList    ' Whole clause
    Dim strWord As String         ' One word
    Dim intI As Integer           ' Counter

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return ""
      If (ndxThis.Name <> "eTree") Then Return ""
      ' Start clause
      colHtml.Add("<html><body>")
      ' Get forest
      ndxFor = ndxThis.SelectSingleNode("./ancestor::forest")
      ' Start output
      colHtml.Add("[" & ndxFor.Attributes("TextId").Value & ":" & ndxFor.Attributes("forestId").Value & "] ")
      ' Get the nodes of this clause
      ndxList = ndxFor.SelectNodes("./descendant::eLeaf")
      For intI = 0 To ndxList.Count - 1
        ' What kind of node is this?
        If (ndxList(intI).Attributes("Type").Value <> "Star") Then
          ' ======== DEBUG =-=-=-=-=-=-=-=-=-
          ' Debug.Print(ndxList(intI).Attributes("Text").Value)
          ' ==================================
          ' Get the word
          strWord = VernToEnglish(ndxList(intI).Attributes("Text").Value)
          ' Are we there?
          If (ndxList(intI).SelectSingleNode("./ancestor-or-self::eTree[@Id=" & ndxThis.Attributes("Id").Value & "]") Is Nothing) Then
            ' This is NOT the selected node
            colHtml.Add(strWord)
          Else
            ' This IS the selected node
            colHtml.Add("<font color='red'>" & strWord & "</font>")
          End If
        End If
      Next intI
      ' Finish html
      colHtml.Add("</body></html>")
      ' Return the result
      Return colHtml.Text
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/ClauseToHtml error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphologyReport
  ' Goal:   Make a report of what we are to return
  ' History:
  ' 18-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphologyReport(ByRef strMorphR As String) As Boolean
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim strHtml As String = ""    ' Output
    Dim colHtml As New StringColl ' Gather what we return
    Dim intI As Integer           ' Counter

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Start report
      colHtml.Add("<html><body><table><tr><td>Vern</td><td>POS</td><td>Parent</td><td>Lemma</td><td>Features</td><td>Freq</td><td>Loc</td></tr>")
      ' Combine and select data
      dtrFound = tdlMorphDict.Tables("Morph").Select("", "Vern ASC, l ASC, f ASC, Label ASC")
      For intI = 0 To dtrFound.Length - 1
        With dtrFound(intI)
          ' Add this line
          colHtml.Add("<tr><td>" & .Item("Vern") & "</td><td>" & .Item("Pos") & "</td><td>" & .Item("Label") & "</td><td>" & .Item("l") & _
                      "</td><td>" & .Item("f") & "</td><td>" & .Item("Freq") & "</td><td>" & .Item("File") & ":" & .Item("forestId") & "_" & .Item("EtreeId") & "</td></tr>")
        End With
      Next intI
      ' Finish report
      colHtml.Add("</table></body></html>")
      strHtml = colHtml.Text
      ' Save the morpholoy report somewhere
      strMorphR = GetDocDir() & "\MorphReport.htm"
      IO.File.WriteAllText(strMorphR, strHtml)
      ' Return success
      Status("Made morphology report")
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/MorphologyReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphCategories
  ' Goal:   Give a report of what is already tagged in the texts
  ' History:
  ' 11-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphCategories(ByVal strDirIn As String, ByRef strCatRepFile As String) As Boolean
    Dim arInFile() As String      ' List of files
    Dim strInFile As String       ' Oone file
    Dim strHtml As String = ""    ' Output
    Dim colHtml As New StringColl ' Gather what we return
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim intI As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage

    Try
      ' Consider all *.psdx input files
      arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories)
      ' Prepare tag-report anyway
      Status("Preparing categories report...")
      For intI = 0 To arInFile.Count - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ arInFile.Count
        Status("Processing " & intPtc & "%", intPtc)
        ' Get the input file name
        strInFile = arInFile(intI)
        ' Also show it on the main form
        Logging("English POS-categories [" & IO.Path.GetFileNameWithoutExtension(strInFile) & "] (" & intI + 1 & "/" & arInFile.Count & ")")
        ' Process this one file
        If (Not OneMorphCatFile(strInFile)) Then
          ' Inform user
          Status("Error in processing: " & strInFile)
          Return False
        End If
      Next intI
      ' Get the report
      colHtml.Add("<html><body><h1>Categories report</h1><p><table>" & _
          "<tr><td>POS</td><td>Head</td><td>Freq</td></tr>")
      dtrFound = tdlMorphDict.Tables("Cat").Select("", "Pos ASC, Freq DESC")
      For intI = 0 To dtrFound.Length - 1
        With dtrFound(intI)
          colHtml.Add("<tr><td>" & .Item("Pos") & "</td><td>" & .Item("Head") & "</td><td>" & .Item("Freq") & "</td></tr>")
        End With
      Next intI
      colHtml.Add("</table>")
      colHtml.Add("</body></html>")
      strHtml = colHtml.Text
      ' Save the morpholoy report somewhere
      strCatRepFile = GetDocDir() & "\CatReport.htm"
      IO.File.WriteAllText(strCatRepFile, strHtml)
      ' Write the morphological dictionary
      Status("Writing morphological dictionary...")
      tdlMorphDict.WriteXml(strMorphDictFile)
      ' Return success
      Status("Made categories report")
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/MorphCategories error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MakeTaggingReport
  ' Goal:   Make a report of what we are to return
  ' History:
  ' 18-02-2013  ERK Created
  ' 04-10-2013  ERK Added [strLabels]
  ' ------------------------------------------------------------------------------------
  Public Function MakeTaggingReport(ByVal strDirIn As String, ByRef strTagRepFile As String, Optional ByVal strLabels As String = "", _
                                    Optional ByVal strLngAbbr As String = "") As Boolean
    Dim arInFile() As String      ' List of files
    Dim strInFile As String       ' Oone file
    Dim strHtml As String = ""    ' Output
    Dim colHtml As New StringColl ' Gather what we return
    Dim intI As Integer           ' Counter
    Dim intD As Integer = 0       ' Number done
    Dim intN As Integer = 0       ' Number needed
    Dim intDone As Integer = 0    ' Total done
    Dim intNeed As Integer = 0    ' Total needed
    Dim intPtc As Integer         ' Percentage

    Try
      ' Consider all *.psdx input files
      arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories)
      ' Prepare tag-report anyway
      Status("Preparing English tagging report...")
      For intI = 0 To arInFile.Count - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ arInFile.Count
        Status("Processing " & intPtc & "%", intPtc)
        ' Get the input file name
        strInFile = arInFile(intI)
        ' Also show it on the main form
        Logging("English tagging [" & IO.Path.GetFileNameWithoutExtension(strInFile) & "] (" & intI + 1 & "/" & arInFile.Count & ")")
        ' Process this one file
        If (Not OneTaggingReportFile(strInFile, intD, intN, strLabels, strLngAbbr)) Then
          ' Inform user
          Status("Error in processing: " & strInFile)
          Return False
        End If
        ' Add totals
        intDone += intD : intNeed += intN
      Next intI
      ' Get the report
      colHtml.Add("<html><body><h1>Tag report</h1><p><table>" & _
          "<tr><td>Restriction:</td><td align='right'>" & IIf(strLabels = "", "all", "verbs") & "</td></tr>" & _
          "<tr><td>Done:</td><td align='right'>" & intDone & "</td></tr>" & _
          "<tr><td>Need:</td><td align='right'>" & intNeed & "</td></tr>" & _
          "<tr><td>Percentage:</td><td align='right'>" & Format(100 * intDone / (intDone + intNeed), "0.00") & "</td></tr>" & _
          "</table></p>")
      colHtml.Add(TagReport)
      colHtml.Add("</body></html>")
      strHtml = colHtml.Text
      ' Save the morpholoy report somewhere
      strTagRepFile = GetDocDir() & "\TaggingReport.htm"
      IO.File.WriteAllText(strTagRepFile, strHtml)
      ' Return success
      Status("Made English tagging report")
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/MakeTaggingReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MakeRemainReport
  ' Goal:   Make a report of what remains to be tagged
  ' History:
  ' 11-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MakeRemainReport(ByVal strDirIn As String, ByRef strTagRepFile As String, ByVal strLngAbbr As String, _
                                   Optional ByVal strLabels As String = "") As Boolean
    Dim arInFile() As String      ' List of files
    Dim strInFile As String       ' Oone file
    Dim strHtml As String = ""    ' Output
    Dim strPeriod As String = ""  ' Period of first file
    Dim colHtml As New StringColl ' Gather what we return
    Dim intI As Integer           ' Counter
    Dim intD As Integer = 0       ' Number done
    Dim intN As Integer = 0       ' Number needed
    Dim intDone As Integer = 0    ' Total done
    Dim intNeed As Integer = 0    ' Total needed
    Dim intPtc As Integer         ' Percentage

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Consider all *.psdx input files
      arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories)
      ' Prepare tag-report anyway
      Status("Preparing remain report...")
      For intI = 0 To arInFile.Count - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ arInFile.Count
        Status("Processing " & intPtc & "%", intPtc)
        ' Get the input file name
        strInFile = arInFile(intI)
        ' Also show it on the main form
        Logging("Remain [" & IO.Path.GetFileNameWithoutExtension(strInFile) & "] (" & intI + 1 & "/" & arInFile.Count & ")")
        ' Process this one file
        If (Not OneRemainReportFile(strInFile, strLngAbbr, strLabels)) Then
          ' Inform user
          Status("Error in processing: " & strInFile)
          Return False
        End If
        ' Get period of first file
        If (strPeriod = "") AndAlso (pdxCurrentFile IsNot Nothing) Then
          strPeriod = GetPeriod(pdxCurrentFile)
        End If
      Next intI
      ' Write the morphological dictionary
      Status("Writing MorphDict with [Remain] section...")
      tdlRemain.WriteXml(strMorphRemainFile)
      ' Get the report
      colHtml.Add("<html><body>")
      colHtml.Add(RemainReport)
      colHtml.Add("</body></html>")
      strHtml = colHtml.Text
      ' Save the remain report somewhere
      strTagRepFile = GetDocDir() & "\RemainReport_" & strPeriod & ".htm"
      IO.File.WriteAllText(strTagRepFile, strHtml)
      ' Return success
      Status("Made Remain report")
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/MakeRemainReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   OEtaggedWordsLike
  ' Goal:   Check if [strIn] is like [strWord] 
  ' History:
  ' 04-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function OEtaggedWordsLike(ByVal strIn As String, ByVal strWord As String) As Boolean
    'Dim arFrom() As String = {"ea", "aw", "j", "i", "y", "th", "dh", "ky", "cy", "mon", "man", "ae", "e"}
    'Dim arTo() As String = {"ae", "au", "i", "y", "i", "dh", "th", "cy", "ky", "man", "mon", "e", "ae"}
    Dim intI As Integer   ' Counter
    Dim bOkay As Boolean  ' Result of comparison
    Try
      ' Initial comparison
      bOkay = (strIn = strWord)
      ' What if there is no initial likeness?
      If (Not bOkay) Then
        ' Try alternative swappings
        For intI = 0 To loc_arOEfrom.Length - 1
          If (strIn = strWord.Replace(loc_arOEfrom(intI), loc_arOEto(intI))) Then bOkay = True : Exit For
          If (strIn.Replace(loc_arOEfrom(intI), loc_arOEto(intI)) = strWord) Then bOkay = True : Exit For
        Next intI
      End If
      'bOkay = (strIn = strWord) OrElse _
      '       (strIn = strWord.Replace("ea", "ae")) OrElse _
      '       (strIn = strWord.Replace("aw", "au")) OrElse _
      '       (strIn = strWord.Replace("j", "i")) OrElse _
      '       (strIn.Replace("i", "y") = strWord) OrElse _
      '       (strIn = strWord.Replace("i", "y"))
      ' Additional possibilities
      If (Not bOkay) Then
        If (strIn = "&" AndAlso DoLike(strWord, "and|ond")) Then
          bOkay = True
        ElseIf (strWord = "&" AndAlso DoLike(strIn, "and|ond")) Then
          bOkay = True
        End If
      End If
      ' Return the result
      Return bOkay
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OEtaggedWordsLike error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   OneEmptyLemmaCheck
  ' Goal:   Check how many features have been added and how many still need doing 
  ' History:
  ' 07-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function OneEmptyLemmaCheck(ByRef pdxFile As XmlDocument, ByRef bChanged As Boolean) As Boolean
    Dim ndxFor As XmlNode = Nothing   ' Forest node
    Dim ndxWork As XmlNode = Nothing  ' Working node
    Dim ndxThis As XmlNode = Nothing  ' Current node
    Dim ndxList As XmlNodeList        ' List of <eLeaf> nodes
    Dim intCount As Integer           ' Number of child nodes
    Dim intPtc As Integer             ' Percentage
    Dim intI As Integer               ' Counter

    Try
      ' Validate
      If (pdxFile Is Nothing) Then Return False
      ' Set the current file
      pdxCurrentFile = pdxFile
      ' Look for the first <forest> element
      If (Not GetFirstForest(pdxFile, ndxFor)) Then Return False
      ' Determine how many elements there are potentially
      If (ndxFor Is Nothing) Then
        intCount = 0
      ElseIf (ndxFor.ParentNode Is Nothing) Then
        intCount = 0
      Else
        intCount = ndxFor.ParentNode.ChildNodes.Count
      End If
      ' Walk through all the <forest> elements FORWARD
      While (Not ndxFor Is Nothing)
        ' Show where we are (how do we KNOW where we are?)
        intPtc = (ndxFor.Attributes("forestId").Value) * 100 \ intCount
        Status("OE-empty checking of [" & IO.Path.GetFileNameWithoutExtension(strCurrentFile) & "] " & intPtc & "%", intPtc)
        ' Get all the <eTree> elements of this forest that have been coded with an empty lemma
        ndxList = ndxFor.SelectNodes("./descendant::eTree[child::fs/child::f[@name='h']/@value = 'propVP_25-4-2013']")
        ' Walk them all in reverse
        For intI = ndxList.Count - 1 To 0 Step -1
          ' get to the <fs> child
          ndxThis = ndxList(intI).SelectSingleNode("./child::fs[@type='M']")
          ' Delete the fs/f children
          If (Not DelXmlNodeChild(ndxThis, "f", "name;h")) Then Return False
          If (Not DelXmlNodeChild(ndxThis, "f", "name;l")) Then Return False
          ' Delete the fs child itself
          If (Not DelXmlNodeChild(ndxList(intI), "fs", "type;M")) Then Return False
          ' Signal we have changed
          bChanged = True
        Next intI
        ' Go to the next forest
        ndxFor = ndxFor.NextSibling
      End While
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OneEmptyLemmaCheck error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   OneTaggedCheckPdx
  ' Goal:   Check how many features have been added and how many still need doing 
  ' History:
  ' 07-02-2013  ERK Created
  ' 19-05-2013  ERK Added the [strLngAbbr] subcategorization option for [MED] processing
  ' ------------------------------------------------------------------------------------
  Public Function OneTaggedCheckPdx(ByRef pdxFile As XmlDocument, ByRef intDone As Integer, ByRef intNeed As Integer, _
                                    Optional ByVal strLabels As String = "", Optional ByVal strLngAbbr As String = "") As Boolean
    Dim ndxFor As XmlNode = Nothing   ' Forest node
    Dim ndxWork As XmlNode = Nothing  ' Working node
    Dim ndxThis As XmlNode = Nothing  ' Current node
    Dim ndxList As XmlNodeList        ' List of <eLeaf> nodes
    Dim dtrFound() As DataRow         ' Result of SELECT
    Dim strSrc As String = ""         ' Source feature
    Dim strF As String = ""           ' Working variable for [MED] processing
    Dim strLemma As String = ""       ' Lemma
    Dim strLabel As String = ""       ' POS label
    Dim strVern As String = ""        ' Vernacular form
    Dim strMed As String = ""         ' MED number
    Dim intCount As Integer           ' Number of child nodes
    Dim intPtc As Integer             ' Percentage
    Dim intI As Integer               ' Counter
    Dim intJ As Integer               ' Counter
    Dim bIsME As Boolean = False      ' Is this middle English?
    Dim bIsOE As Boolean = False      ' Is this the OE that needs special treatment?
    Dim bChanged As Boolean = False   ' Needs saving
    Dim bFound As Boolean             ' Found or not?
    Dim colForm As New StringColl     ' Collection of forms

    Try
      ' Validate
      If (pdxFile Is Nothing) Then Return False
      ' Set the current file
      pdxCurrentFile = pdxFile : SetXmlDocument(pdxCurrentFile)
      ' Look for the first <forest> element
      If (Not GetFirstForest(pdxFile, ndxFor)) Then Return False
      ' Determine how many elements there are potentially
      If (ndxFor Is Nothing) Then
        intCount = 0
      ElseIf (ndxFor.ParentNode Is Nothing) Then
        intCount = 0
      Else
        intCount = ndxFor.ParentNode.ChildNodes.Count
      End If
      ' Check for middle English/MED processing
      bIsME = (strLngAbbr = "MED" AndAlso Left(ndxFor.Attributes("TextId").Value.ToLower, 2) = "cm")
      If (Not bIsME) Then
        bIsOE = (strLngAbbr = "OED")
      End If
      ' Other initialisations
      intDone = 0 : intNeed = 0
      ' Walk through all the <forest> elements FORWARD
      While (Not ndxFor Is Nothing)
        ' Show where we are (how do we KNOW where we are?)
        intPtc = (ndxFor.Attributes("forestId").Value) * 100 \ intCount
        ' Status("OE-tagged checking of [" & IO.Path.GetFileNameWithoutExtension(strCurrentFile) & "] " & intPtc & "%", intPtc)
        ' Get all the <eLeaf> elements of this forest, provided they are vernacular (skipping punctuation, skipping Foreign Words)
        ' NB: there may be some punctuation that slipped through; their parent label is "."
        If (strLabels = "") Then
          ndxList = ndxFor.SelectNodes("./descendant::eLeaf[" & strLeafCondition_NoFW & "]")
        Else
          ' Only include those that are in [strLabels]
          ndxList = ndxFor.SelectNodes("./descendant::eLeaf[(" & strLeafCondition_NoFW & ") and " & _
                                       " parent::eTree[tb:matches(@Label, '" & strLabels & "')] ]", conTb)
        End If
        ' Walk them all
        For intI = 0 To ndxList.Count - 1
          ' Get my parent, which is the <eTree> node
          ndxThis = ndxList(intI).ParentNode
          Select Case strLngAbbr
            Case "OEB"
              strF = GetFeature(ndxThis, "M", "s")
              If (Regex.IsMatch(strF, "BT\d+")) Then
                intDone += 1
              Else
                intNeed += 1
              End If
            Case "OED"
              ' Check if we have "M" features (morphology) or not
              If (ndxThis.SelectSingleNode("./child::fs[@type='M']") Is Nothing) Then
                ' No M features
                intNeed += 1
              Else
                ' There are M features
                intDone += 1
              End If
              ' =================== CHECKING the OE numbers ===================
              ' Get the vernacular form
              strVern = ndxThis.SelectSingleNode("./descendant::eLeaf[@Type='Vern'][1]").Attributes("Text").Value
              strLabel = ndxThis.Attributes("Label").Value
              strLemma = GetFeature(ndxThis, "M", "l")
              ' Does this already contain a "M::src" feature?
              strF = GetFeature(ndxThis, "M", "src")
              If (strF = "") OrElse (GetMEDorOED(strF) = "") Then
                ' Make all the forms needed
                colForm.Clear()
                colForm.Add(StringOEtoTagged(strVern, True).Replace("$", "").Replace("'", "''"))
                colForm.AddUnique(StringOEtoTagged(strVern, False).Replace("$", "").Replace("'", "''"))
                colForm.AddUnique(StringOEtoTagged(strLemma, True).Replace("$", "").Replace("'", "''"))
                colForm.AddUnique(StringOEtoTagged(strLemma, False).Replace("$", "").Replace("'", "''"))
                ' Find the corresponding lemma
                bFound = False : ReDim dtrFound(0)
                For intJ = 0 To colForm.Count - 1
                  dtrFound = tdlMorphDict.Tables("VernPos").Select("Vern='" & colForm.Item(intJ) & "' AND Pos='" & strLabel & "'")
                  ' If (dtrFound.Length = 0) Then Stop
                  If (dtrFound.Length > 0) Then bFound = True : Exit For
                Next intJ
                If (Not bFound) Then
                  For intJ = 0 To colForm.Count - 1
                    dtrFound = tdlMorphDict.Tables("Morph").Select("Vern='" & colForm.Item(intJ) & "' AND Label='" & strLabel & "'")
                    If (dtrFound.Length > 0) Then bFound = True : Exit For
                  Next intJ
                End If
                ' Action depends on the number of results
                Select Case dtrFound.Length
                  Case 0
                    ' Check if it needs adaptation because of its label
                    If (DoLike(strLabel, "B*|*+B*")) Then
                      ' Set the M/l feature
                      AddFeature(pdxCurrentFile, ndxThis, "M", "l", "beon")
                      'strSrc = GetFeature(ndxThis, "M", "src") : AddSemiStack(strSrc, "s=MED4048")
                      'AddFeature(pdxCurrentFile, ndxThis, "M", "src", strSrc)
                      AddFeature(pdxCurrentFile, ndxThis, "M", "s", "BT3713")
                    ElseIf (DoLike(strLabel, "HV*|*+HV*")) Then
                      ' Set the M/l feature
                      AddFeature(pdxCurrentFile, ndxThis, "M", "l", "habban")
                      'strSrc = GetFeature(ndxThis, "M", "src") : AddSemiStack(strSrc, "s=MED20173")
                      'AddFeature(pdxCurrentFile, ndxThis, "M", "src", strSrc)
                      AddFeature(pdxCurrentFile, ndxThis, "M", "s", "BT17802")
                    End If
                  Case 1
                    ' Yes, there is a result
                    With dtrFound(0)
                      ' Add the correct lemma
                      AddFeature(pdxCurrentFile, ndxThis, "M", "l", .Item("l").ToString)
                      'strSrc = GetFeature(ndxThis, "M", "src") : AddSemiStack(strSrc, .Item("f").ToString)
                      'AddFeature(pdxCurrentFile, ndxThis, "M", "src", strSrc)
                      ' Propagate features
                      If (Not AddFeatureSet(ndxThis, "M", .Item("f").ToString)) Then Return False
                      ' Indicate that we are changed
                      bChanged = True
                    End With
                  Case Else
                    ' There are multiple results, so I should combine them into one item
                    strF = "" : strLemma = ""
                    For intJ = 0 To dtrFound.Length - 1
                      AddSemiStack(strF, GetMEDorOED(dtrFound(intJ).Item("f").ToString), True, "|")
                      AddSemiStack(strLemma, dtrFound(intJ).Item("l").ToString, True)
                    Next intJ
                    ' Add this to the node
                    AddFeature(pdxCurrentFile, ndxThis, "M", "l", strLemma)
                    strSrc = GetFeature(ndxThis, "M", "src") : AddSemiStack(strSrc, "s=" & strF)
                    'AddFeature(pdxCurrentFile, ndxThis, "M", "src", strSrc)
                    ' Propagate features
                    If (Not AddFeatureSet(ndxThis, "M", strSrc)) Then Return False
                    ' Indicate that we are changed
                    bChanged = True
                End Select

              Else
                ' Check if correction within tables [VernPos] and/or [Morph] is called for
                If (Not CheckMedNumber(strVern, strLabel, GetFeature(ndxThis, "M", "l"), strF)) Then Return False
              End If
            Case "MED"
              ' Check if we have "M" features (morphology) or not
              If (ndxThis.SelectSingleNode("./child::fs[@type='M']") Is Nothing) Then
                ' No M features
                intNeed += 1
              Else
                ' There are M features
                intDone += 1
              End If
              ' Get the vernacular form
              strVern = ndxThis.SelectSingleNode("./descendant::eLeaf[@Type='Vern'][1]").Attributes("Text").Value
              strLabel = ndxThis.Attributes("Label").Value
              strLemma = GetFeature(ndxThis, "M", "l")
              ' Does this already contain a "M::src" feature?
              strF = GetFeature(ndxThis, "M", "src")
              If (strF = "") OrElse (Not DoLike(strF, "*s=[MO]ED*")) Then
                ' Make all the forms needed
                colForm.Clear()
                colForm.Add(StringOEtoTagged(strVern, True).Replace("$", "").Replace("'", "''"))
                colForm.AddUnique(StringOEtoTagged(strVern, False).Replace("$", "").Replace("'", "''"))
                colForm.AddUnique(StringOEtoTagged(strLemma, True).Replace("$", "").Replace("'", "''"))
                colForm.AddUnique(StringOEtoTagged(strLemma, False).Replace("$", "").Replace("'", "''"))
                ' Find the corresponding lemma
                bFound = False : ReDim dtrFound(0)
                For intJ = 0 To colForm.Count - 1
                  dtrFound = tdlMorphDict.Tables("VernPos").Select("Vern='" & colForm.Item(intJ) & "' AND Pos='" & strLabel & "'")
                  ' If (dtrFound.Length = 0) Then Stop
                  If (dtrFound.Length > 0) Then bFound = True : Exit For
                Next intJ
                If (Not bFound) Then
                  For intJ = 0 To colForm.Count - 1
                    dtrFound = tdlMorphDict.Tables("Morph").Select("Vern='" & colForm.Item(intJ) & "' AND Label='" & strLabel & "'")
                    If (dtrFound.Length > 0) Then bFound = True : Exit For
                  Next intJ
                End If
                ' Action depends on the number of results
                Select Case dtrFound.Length
                  Case 0
                    ' Check if it needs adaptation because of its label
                    If (DoLike(strLabel, "B*|*+B*")) Then
                      ' Set the M/l feature
                      AddFeature(pdxCurrentFile, ndxThis, "M", "l", "ben")
                      'strSrc = GetFeature(ndxThis, "M", "src") : AddSemiStack(strSrc, "s=MED4048")
                      'AddFeature(pdxCurrentFile, ndxThis, "M", "src", strSrc)
                      AddFeature(pdxCurrentFile, ndxThis, "M", "s", "MED4048")
                    ElseIf (DoLike(strLabel, "HV*|*+HV*")) Then
                      ' Set the M/l feature
                      AddFeature(pdxCurrentFile, ndxThis, "M", "l", "haven")
                      'strSrc = GetFeature(ndxThis, "M", "src") : AddSemiStack(strSrc, "s=MED20173")
                      'AddFeature(pdxCurrentFile, ndxThis, "M", "src", strSrc)
                      AddFeature(pdxCurrentFile, ndxThis, "M", "s", "MED20173")
                    ElseIf (DoLike(strLabel, "D[OA]*|*+D[OA]*")) Then
                      ' Set the M/l feature
                      AddFeature(pdxCurrentFile, ndxThis, "M", "l", "don")
                      'strSrc = GetFeature(ndxThis, "M", "src") : AddSemiStack(strSrc, "s=MED12367")
                      'AddFeature(pdxCurrentFile, ndxThis, "M", "src", strSrc)
                      AddFeature(pdxCurrentFile, ndxThis, "M", "s", "MED12367")
                    Else
                      ' Check if this is a Modal
                      If (strLabel = "MD") Then
                        ' Initialise: 
                        strMed = ""
                        ' Depends on the particular modal we are looking at
                        If (DoLike(strVern, "cu*|ca*|co*|k*")) Then
                          strLemma = "connen" : strMed = "MED9326"
                        ElseIf (DoLike(strVern, "au*|ac*|ag*|ah*|aw*|ow*")) Then
                          strLemma = "ouen" : strMed = "MED31051"
                        ElseIf (DoLike(strVern, "d*")) Then
                          strLemma = "durren" : strMed = "MED12843"
                        ElseIf (DoLike(strVern, "mo*t*")) Then
                          strLemma = "moten" : strMed = "MED28752"
                        ElseIf (DoLike(strVern, "m*")) Then
                          strLemma = "mouen" : strMed = "MED28782"
                        ElseIf (DoLike(strVern, "o*t|ou*|og*")) Then
                          strLemma = "ouen" : strMed = "MED31051"
                        ElseIf (DoLike(strVern, "s*|x*")) Then
                          strLemma = "shulen" : strMed = "MED40162"
                        ElseIf (DoLike(strVern, "w*")) Then
                          strLemma = "willen" : strMed = "MED52797"
                        End If
                        If (strMed = "") Then
                          ' Log this, so I can add/correct it later in [VernPos] and in [Morph]
                          Logging(StringOEtoTagged(strVern, True) & vbTab & strLabel & vbTab & strLemma & vbTab & "-")
                        Else
                          ' Process it
                          AddFeature(pdxCurrentFile, ndxThis, "M", "l", strLemma)
                          'strSrc = GetFeature(ndxThis, "M", "src") : AddSemiStack(strSrc, "s=" & strMed)
                          'AddFeature(pdxCurrentFile, ndxThis, "M", "src", strSrc)
                          AddFeature(pdxCurrentFile, ndxThis, "M", "s", strMed)
                        End If
                      End If
                    End If
                  Case 1
                    ' Yes, there is a result
                    With dtrFound(0)
                      ' Add the correct lemma
                      AddFeature(pdxCurrentFile, ndxThis, "M", "l", .Item("l").ToString)
                      'strSrc = GetFeature(ndxThis, "M", "src") : AddSemiStack(strSrc, .Item("f").ToString)
                      'AddFeature(pdxCurrentFile, ndxThis, "M", "src", strSrc)
                      ' Propagate features
                      If (Not AddFeatureSet(ndxThis, "M", .Item("f").ToString)) Then Return False
                      ' Indicate that we are changed
                      bChanged = True
                    End With
                  Case Else
                    ' There are multiple results, so I should combine them into one item
                    strF = "" : strLemma = ""
                    For intJ = 0 To dtrFound.Length - 1
                      AddSemiStack(strF, GetMEDorOED(dtrFound(intJ).Item("f").ToString), True, "|")
                      AddSemiStack(strLemma, dtrFound(intJ).Item("l").ToString, True)
                    Next intJ
                    ' Add this to the node
                    AddFeature(pdxCurrentFile, ndxThis, "M", "l", strLemma)
                    strSrc = GetFeature(ndxThis, "M", "src") : AddSemiStack(strSrc, "s=" & strF)
                    'AddFeature(pdxCurrentFile, ndxThis, "M", "src", strSrc)
                    ' Propagate features
                    If (Not AddFeatureSet(ndxThis, "M", strSrc)) Then Return False
                    ' Indicate that we are changed
                    bChanged = True
                End Select
              Else
                ' Check if correction within tables [VernPos] and/or [Morph] is called for
                If (Not CheckMedNumber(strVern, strLabel, GetFeature(ndxThis, "M", "l"), strF)) Then Return False
              End If
            Case Else
              ' Check if we have "M" features (morphology) or not
              If (ndxThis.SelectSingleNode("./child::fs[@type='M']") Is Nothing) Then
                ' No M features
                intNeed += 1
              Else
                ' There are M features
                intDone += 1
              End If
          End Select
        Next intI
        ' Go to the next forest
        ndxFor = ndxFor.NextSibling
      End While
      ' DO changes need to be saved?
      If (bChanged) Then
        ' Need to save changes
        pdxCurrentFile.Save(strCurrentFile)
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OneTaggedCheckPdx error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   OneTaggedFeaturesPdx
  ' Goal:   Actually try to add features from OE-tagged to this YCOE file 
  ' Algorithm:
  '    =======================================================
  '    FOR each $forest in $file.forests
  '  	   $tor := GetToronto($forest)
  '      IF ($tor != $torF) THEN
  '        $tagFile := GetTaggedFile($tor, ‘textlist.sgml’)
  '      END IF
  '      FOR each $leaf in $forest.leaves
  '        $tagUnit := GetNextUnit($tagFile)
  '        $leaf.Features := $tagUnit.Features
  '      NEXT $leaf
  '    NEXT $forest
  '    =======================================================
  ' History:
  ' 29-01-2013  ERK Created
  ' 04-02-2013  ERK function OEtaggedWordsLike() for the proper comparison
  ' ------------------------------------------------------------------------------------
  Public Function OneTaggedFeaturesPdx(ByRef pdxFile As XmlDocument, ByRef bChanged As Boolean, Optional ByVal bRecalculate As Boolean = False) As Boolean
    Dim ndxFor As XmlNode = Nothing   ' Forest node
    Dim ndxWork As XmlNode = Nothing  ' Working node
    Dim ndxThis As XmlNode = Nothing  ' Current node
    Dim ndxList As XmlNodeList        ' List of <eLeaf> nodes
    Dim strToron As String = ""       ' Toronto ID
    Dim strCamno As String = ""       ' Cameron ID
    Dim strOEfile As String = ""      ' Filename used for OE-tagging
    Dim arTagged() As String = Nothing  ' Array with lines from OE-tagged file
    Dim strLine As String = ""        ' One line from the tagged array
    Dim strWord As String = ""        ' One word
    Dim strCombi As String = ""       ' combined word
    Dim objInfl As New Inflected      ' One inflected type
    Dim colMis As New StringColl      ' Mismatch-showing collection
    Dim intTagged As Integer = -1     ' Counter in the tagged array
    Dim intSentStart As Integer = -1  ' Sentence start in the tagged array
    Dim intSentEnd As Integer = -1    ' Sentence end in the tagged array
    Dim intI As Integer               ' Counter
    Dim intJ As Integer               ' Counter
    Dim intThis As Integer            ' Current number
    Dim intTagThis As Integer         ' Current number under [arTagged]
    Dim intPtc As Integer             ' Percentage
    Dim intCount As Integer           ' Number of child nodes
    Dim bHaveWord As Boolean = False  ' Do we have the word or not?
    Dim bHaveFile As Boolean = False  ' Do we have a relevant file loaded?

    Try
      ' Validate
      If (pdxFile Is Nothing) Then Return False
      ' Set the current file
      pdxCurrentFile = pdxFile
      ' Look for the first <forest> element
      If (Not GetFirstForest(pdxFile, ndxFor)) Then Return False
      ' Determine how many elements there are potentially
      If (ndxFor Is Nothing) Then
        intCount = 0
      ElseIf (ndxFor.ParentNode Is Nothing) Then
        intCount = 0
      Else
        intCount = ndxFor.ParentNode.ChildNodes.Count
      End If
      ' Other Initialisations
      bChanged = False
      ' Walk through all the <forest> elements FORWARD
      While (Not ndxFor Is Nothing)
        ' Look for and clear any previous "M" type annotations
        If (Not DeleteTaggedFeatures(ndxFor)) Then Status("Could not delete previous annotation") : Return False
        ' Check if we are not being interrupted
        If (bInterrupt) Then Return False
        ' Get toronto ID -- if existing
        ndxThis = ndxFor.SelectSingleNode("./descendant::eLeaf[tb:Like(@Text, '<T0*|<T1*|<T2*|<T3*') and parent::eTree[@Label = 'CODE']]", conTb)
        If (ndxThis IsNot Nothing) Then
          ' Do we need to change it?
          If (Mid(ndxThis.Attributes("Text").Value, 2, 6) <> strToron) Then
            ' Get and change the toronto Id
            strToron = Mid(ndxThis.Attributes("Text").Value, 2, 6)
            ' Get the corresponding Cameron Id
            strCamno = GetCamno(strToron)
            ' Load the OE-tagged information with this cameron number (if any)
            Status("Loading OE-tagged " & strCamno)
            If (GetTaggedFile(strCamno, arTagged, strOEfile)) Then
              ' Reset the tagged array index
              intTagged = 0
              ' Indicate that we have a relevant file
              bHaveFile = True
              ' Show we are ready loading
              Status("Ready loading..." & strCamno)
            Else
              ' If have file is set to false, we don't have to repeat messages
              If (bHaveFile) Then
                ' We had a file, but not anymore
                TagComment("Could not retrieve tagged file for [" & strCamno & "]")
                bHaveFile = False
              End If
              'Return True
            End If
          End If
        End If
        ' Show where we are (how do we KNOW where we are?)
        intPtc = (ndxFor.Attributes("forestId").Value) * 100 \ intCount
        Status("OE-tagged features of [" & IO.Path.GetFileNameWithoutExtension(strCurrentFile) & _
               "/" & strCamno & "] " & intPtc & "%", intPtc)
        ' Do we actually (still) have a Toronto number?
        If (bHaveFile) AndAlso (strCamno <> "") AndAlso (intTagged < arTagged.Length) Then
          ' Get all the <eLeaf> elements of this forest, provided they are vernacular (skipping punctuation)
          ' NB: there may be some punctuation that slipped through; their parent label is "."
          ndxList = ndxFor.SelectNodes("./descendant::eLeaf[" & strLeafCondition & "]")
          ' Note the start of this sentence in [arTagged]
          intSentStart = GotoNextTagged(arTagged, intTagged)
          ' Double check to see if [arTagged] still has elements
          If (intTagged < arTagged.Length - 1) Then
            ' Visit all these <eLeaf> elements
            For intI = 0 To ndxList.Count - 1
              ' Fix this number
              intThis = intI : intTagThis = intTagged
              ' Get the word, but in "tagged" format for comparison, and without $ symbol
              strWord = StringOEtoTagged(LCase(ndxList(intI).Attributes("Text").Value)).Replace("$", "")
              ' Get the <eTree> node corresponding
              ndxThis = ndxList(intI).ParentNode
              ' Get the next element situated in [arTagged]
              If (Not GetNextTagged(arTagged, intTagged, objInfl, strWord, ndxFor.Attributes("Location").Value)) Then
                TagComment("OneTaggedFeaturesPdx (1): could not get next tag line forest = " & _
                           ndxFor.Attributes("forestId").Value)
                Return False
              End If
              ' Check if this is the correct word
              bHaveWord = OEtaggedWordsLike(strWord, objInfl.Word.ToString)
              If (Not bHaveWord) Then
                ' Check if the word is PART of a larger whole
                If (InStr(strWord, objInfl.Word) = 1) AndAlso (intTagged < arTagged.Length - 1) Then
                  ' Start a combi
                  strCombi = objInfl.Word
                  ' Get the next element situated in [arTagged]
                  If (Not GetNextTagged(arTagged, intTagged, objInfl, strWord & "+", ndxFor.Attributes("Location").Value)) Then
                    ' TagComment("OneTaggedFeaturesPdx: could not get next tag line forest = " & ndxFor.Attributes("forestId").Value)
                    If (ndxFor.NextSibling Is Nothing) AndAlso (intI >= ndxList.Count - 1) Then
                      ' This is probably the end anyway, so let's okay it
                      'bChanged = True (changing should be done where we REALLY change something, further down )
                      Return True
                    Else
                      ' make a severe comment
                      TagComment("OneTaggedFeaturesPdx (2): could not get next tag line forest = " & ndxFor.Attributes("forestId").Value)
                      ' Return failure
                      Return False
                    End If
                  End If
                  ' Combine
                  strCombi &= objInfl.Word
                  ' Do we have it?
                  bHaveWord = OEtaggedWordsLike(strWord, strCombi)
                  ' Remedy the value of [intTagged] if this is not a match
                  If (Not bHaveWord) Then intTagged -= 1
                ElseIf (InStr(objInfl.Word, strWord) = 1) AndAlso (intI < ndxList.Count - 2) Then
                  ' Start a combi
                  strCombi = strWord
                  ' Read the next word in the PDX
                  strWord = StringOEtoTagged(LCase(ndxList(intI + 1).Attributes("Text").Value)).Replace("$", "")
                  ' Get the <eTree> node corresponding
                  ndxThis = ndxList(intI + 1).ParentNode
                  ' Combine the two
                  strCombi &= strWord
                  ' Do we have it?
                  bHaveWord = OEtaggedWordsLike(strCombi, objInfl.Word.ToString)
                  ' Only increment [intI] if we have it!
                  If (bHaveWord) Then intI += 1
                End If
                ' Still not?
                If (Not bHaveWord) Then
                  ' ============ Try synchronization!! ==========================================
                  ' Check if the current word in [ndxList] is equal to the NEXT word in [arTagged]
                  If (intTagged < arTagged.Length - 2) AndAlso (OEtaggedWordsLike(strWord, TagToWord(arTagged(intTagged + 1)))) Then
                    ' So we need to read the next word in [arTagged]
                    If (Not GetNextTagged(arTagged, intTagged, objInfl, strWord & "+", ndxFor.Attributes("Location").Value)) Then
                      TagComment("OneTaggedFeaturesPdx (3): could not get next tag line forest = " & ndxFor.Attributes("forestId").Value) : Return False
                    End If
                    ' And we adapt the flag
                    bHaveWord = True
                  ElseIf (ndxList.Count > intI + 2) Then
                    ' Check if the current word in [arTagged] is equal to the NEXT word in [ndxList]
                    strWord = StringOEtoTagged(LCase(ndxList(intI + 1).Attributes("Text").Value)).Replace("$", "")
                    If (OEtaggedWordsLike(strWord, objInfl.Word.ToString)) Then
                      ' Okay, we have it, but we have to move the pointer one further
                      intI += 1
                      ' make sure we point to the correct place
                      ndxThis = ndxList(intI).ParentNode
                      ' And we adapt the flag
                      bHaveWord = True
                    ElseIf (ndxList.Count > intI + 3) Then
                      ' We could try to move one MORE place ahead?
                      strWord = StringOEtoTagged(LCase(ndxList(intI + 2).Attributes("Text").Value)).Replace("$", "")
                      If (OEtaggedWordsLike(strWord, objInfl.Word.ToString)) Then
                        ' Okay, we have it, but we have to move the pointer one further
                        intI += 2
                        ' make sure we point to the correct place
                        ndxThis = ndxList(intI).ParentNode
                        ' And we adapt the flag
                        bHaveWord = True
                      End If
                    End If
                  End If
                  ' ============ End synchronization trial =======================================
                  ' Are we still not sure?
                  If (Not bHaveWord) Then
                    ' Ask user
                    With frmTagSync
                      ' Put the OE-tagged clause in place
                      .OEtagged(arTagged, intTagThis - intThis, ndxList.Count, intTagThis)
                      ' Put the YCOE clause in place
                      .YCOE(ndxList, intThis)
                      ' Set the location
                      .LocShow = IO.Path.GetFileNameWithoutExtension(strCurrentFile) & ":" & ndxFor.Attributes("forestId").Value
                      ' Set the filename to be shown
                      .OEfile = strOEfile
                      ' Check what the user wants
                      Select Case .ShowDialog
                        Case DialogResult.Abort
                          ' End of file has been reached -- start looking for a new Tagged file, while retaining current ndx stuff
                          ndxFor = ndxFor.PreviousSibling : bHaveFile = False
                          Exit For
                        Case MsgBoxResult.Yes, MsgBoxResult.Ok
                          ' Check the offset parameters the user may have changed
                          If (.OEtagOffset <> 0) Then
                            ' We should reset the OEtagged corpus point, but subtract one, since it had already gone ahead earlier
                            intTagged = intTagThis + .OEtagOffset
                            ' Get the word from YCOE again, but in "tagged" format for comparison, and without $ symbol
                            strWord = StringOEtoTagged(LCase(ndxList(intI).Attributes("Text").Value)).Replace("$", "")
                            ' Get the next element situated in [arTagged]
                            ' Also get the correct [objInfl]
                            If (Not GetNextTagged(arTagged, intTagged, objInfl, strWord, ndxFor.Attributes("Location").Value)) Then
                              TagComment("OneTaggedFeaturesPdx (4): could not get next tag line forest = " & ndxFor.Attributes("forestId").Value) : Return False
                            End If
                          ElseIf (.YcoeOffset <> 0) Then
                            ' Reset the Ycoe corpus point
                            intI += .YcoeOffset
                            ' make sure we point to the correct place
                            ndxThis = ndxList(intI).ParentNode
                          End If
                          ' Signal that we have it
                          bHaveWord = True
                        Case MsgBoxResult.No
                          ' Are we at the end of the current tagged file?
                          If (intTagged >= arTagged.Length - 1) Then
                            ' End of file has been reached -- start looking for a new Tagged file, while retaining current ndx stuff
                            ndxFor = ndxFor.PreviousSibling : bHaveFile = False
                            Exit For
                          End If
                          ' We should skip this one
                          ' Stop
                        Case MsgBoxResult.Cancel
                          Return False
                        Case DialogResult.Retry
                          ' This is the signal that we have to go to the next line of YCOE, and try it from there
                          intTagged -= 1
                          Exit For
                        Case DialogResult.Ignore
                          ' This signals that we need to skip through to @forestId identified
                          ndxWork = ndxFor.SelectSingleNode("./following-sibling::forest[@forestId=" & .ForestId & "]")
                          ' Or should we have preceding sibling?
                          If (ndxWork Is Nothing) Then
                            ndxWork = ndxFor.SelectSingleNode("./preceding-sibling::forest[@forestId=" & .ForestId & "]")
                          End If
                          ' But we need to go one place back before we exit this FOR loop
                          If (ndxWork IsNot Nothing) Then ndxFor = ndxWork.PreviousSibling
                          intTagged -= 1
                          Exit For
                      End Select
                    End With
                  End If
                End If
              End If
              ' Did we actually find it?
              If bHaveWord Then
                ' Show what we are doing for debugging
                ' Debug.Print("Copying features from [" & objInfl.Word & "] --> [" & ndxThis.SelectSingleNode("./descendant::eLeaf").Attributes("Text").Value & "]")
                ' Copy all the features we need
                If (objInfl.ComType <> "") Then If (Not AddFeature(pdxFile, ndxThis, "M", "comType", objInfl.ComType)) Then Return False
                If (objInfl.GrCase <> "") Then If (Not AddFeature(pdxFile, ndxThis, "M", "case", objInfl.GrCase)) Then Return False
                If (objInfl.GrGender <> "") Then If (Not AddFeature(pdxFile, ndxThis, "M", "gender", objInfl.GrGender)) Then Return False
                If (objInfl.GrNumber <> "") Then If (Not AddFeature(pdxFile, ndxThis, "M", "number", objInfl.GrNumber)) Then Return False
                If (objInfl.Grperson <> "") Then If (Not AddFeature(pdxFile, ndxThis, "M", "person", objInfl.Grperson)) Then Return False
                If (objInfl.Lemma <> "") Then If (Not AddFeature(pdxFile, ndxThis, "M", "lemma", objInfl.Lemma)) Then Return False
                If (objInfl.Mood <> "") Then If (Not AddFeature(pdxFile, ndxThis, "M", "mood", objInfl.Mood)) Then Return False
                If (objInfl.Tense <> "") Then If (Not AddFeature(pdxFile, ndxThis, "M", "tense", objInfl.Tense)) Then Return False
                ' Signal changes
                bChanged = True
              Else
                ' ================== DEBUG ===============================
                ' Show the two sentences and their correspondence
                colMis.Clear() : strWord = "" : strCombi = ""
                For intJ = 0 To ndxList.Count - 1
                  ' Note the critical point
                  If (intJ = intI) Then
                    strWord &= ">> "
                    strCombi &= ">> "
                  End If
                  ' Let [strWord] contain the contents of [arTagged]
                  If (intSentStart + intJ <= arTagged.Length - 1) Then
                    strWord &= TagToWord(arTagged(intSentStart + intJ)) & " "
                  Else
                    strWord &= "[END] "
                  End If
                  ' Let [strCombi] contain the contents of [ndxList]
                  strCombi &= StringOEtoTagged(LCase(ndxList(intJ).Attributes("Text").Value)).Replace("$", "") & " "
                Next intJ
                TagComment("OneTaggedFeaturesPdx mismatch (skip):" & vbCrLf & "arTagged = " & strWord & vbCrLf & _
                            "ndxList = " & strCombi & vbCrLf)
                ' Stop

                ' ========================================================
              End If
              ' Double check where we are
              If (intTagged >= arTagged.Length) Then Exit For
            Next intI
          End If
        End If
        ' Go to the next forest
        ndxFor = ndxFor.NextSibling
      End While
      ' Return success
      'bChanged = True (Change flag must be set where we actually change something)
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OneTaggedFeaturesPdx error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DeleteTaggedFeatures
  ' Goal:   Delete all features within [ndxFor] of type <fs type='M'>
  ' History:
  ' 06-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function DeleteTaggedFeatures(ByRef ndxFor As XmlNode) As Boolean
    Dim ndxList As XmlNodeList  ' List of all <fs> nodes in this forest
    Dim ndxThis As XmlNode      ' Working node
    Dim ndxParent As XmlNode    ' my parent node
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (ndxFor Is Nothing) Then Return False
      ' Make a list of all relevant nodes
      ndxList = ndxFor.SelectNodes("./descendant::fs[@type='M']")
      ' Walk them all in reverse order
      For intI = ndxList.Count - 1 To 0 Step -1
        ' Look at this one
        ndxThis = ndxList(intI) : ndxParent = ndxThis.ParentNode
        ' Remove all my children below
        ndxThis.RemoveAll()
        ' Remove myself
        ndxParent.RemoveChild(ndxThis)
      Next intI
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/DeleteTaggedFeatures error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   TagToWord
  ' Goal:   Get the initial word (if any) from the "tagged" string
  ' History:
  ' 30-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function TagToWord(ByVal strIn As String) As String
    Dim intPos As Integer   ' Position

    Try
      ' Find position in string
      intPos = InStr(strIn, "_")
      If (intPos = 0) Then Return ""
      Return Left(strIn, intPos - 1)
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/TagToWord error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   StringTaggedToOE
  ' Goal:   Convert the "tagged" string into OE format
  ' History:
  ' 29-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function StringTaggedToOE(ByVal strIn As String) As String
    Dim arCharIn() As String = {"th", "ae", "dh", "gh", "Th", "Ae", "Dh", "Gh", "TH", "AE", "DH", "GH"}
    Dim arCharOut() As String = {"+t", "+a", "+d", "+g", "+T", "+A", "+D", "+G", "+T", "+A", "+D", "+G"}
    Dim intI As Integer    ' Counter

    Try
      ' Validate
      If (strIn = "") Then Return ""
      ' Replace matters
      For intI = 0 To arCharIn.Length - 1
        strIn = strIn.Replace(arCharIn(intI), arCharOut(intI))
      Next intI
      ' Return the result
      Return strIn
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/StringTaggedToOE error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return strIn
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   StringOEtoTagged
  ' Goal:   Convert the OE format into the "tagged" one
  ' History:
  ' 29-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function StringOEtoTagged(ByVal strIn As String, Optional ByVal bYoghAs3 As Boolean = False) As String
    Dim arCharIn() As String = {"+t", "+a", "+d", "+g", "+e", "+T", "+A", "+D", "+G", "+E", _
                                "ǽ", "Æ", "æ", "ð", "đ", "þ", "Þ", "Đ", "ȝ", "Ȝ", "á", "à", "é", "è", "ë", "í", "ì", "ó", "ò", "ú", "ù", "ý", "Á", "É"}
    Dim arCharOut() As String = {"th", "ae", "dh", "gh", "e", "Th", "Ae", "Dh", "Gh", "E", _
                                 "ae", "Ae", "ae", "dh", "dh", "th", "Th", "Dh", "y", "Y", "a", "a", "e", "e", "e", "i", "i", "o", "o", "u", "o", "y", "A", "E"}
    Dim arCharOutY3() As String = {"th", "ae", "dh", "3", "e", "Th", "Ae", "Dh", "3", "E", _
                                 "ae", "Ae", "ae", "dh", "dh", "th", "Th", "Dh", "3", "3", "a", "a", "e", "e", "e", "i", "i", "o", "o", "u", "o", "y", "A", "E"}
    Dim intI As Integer    ' Counter

    Try
      ' Validate
      If (strIn = "") Then Return ""
      ' Replace matters
      For intI = 0 To arCharIn.Length - 1
        If (bYoghAs3) Then
          strIn = strIn.Replace(arCharIn(intI), arCharOutY3(intI))
        Else
          strIn = strIn.Replace(arCharIn(intI), arCharOut(intI))
        End If
      Next intI
      ' Return the result
      Return strIn
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/StringOEtoTagged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return strIn
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetNextTagSentence
  ' Goal:   Get the start and end element of the next sentence within the [arTagged] array
  ' History:
  ' 30-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetNextTagSentence(ByRef arTagged() As String, ByRef intTagged As Integer, _
                                       ByRef intSentStart As Integer, ByRef intSentEnd As Integer) As Boolean
    Try
      ' Validate
      If (arTagged.Length = 0) Then Return False
      If (intTagged >= arTagged.Length) Then Return False
      ' Read until we encounter _bos
      While (arTagged(intTagged) <> "_bos")
        intTagged += 1
      End While
      ' Set the start-of-sentence
      intTagged += 1 : intSentStart = intTagged
      ' Read until we encounter _eos
      While (arTagged(intTagged) <> "_eos")
        intTagged += 1
      End While
      ' Set the end-of-sentence
      intSentEnd = intTagged - 1
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetNextTagSentence error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GotoNextTagged
  ' Goal:   Goto the first element in [arTagged] that contains an actual word
  ' History:
  ' 29-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GotoNextTagged(ByRef arTagged() As String, ByRef intTagged As Integer) As Integer
    Try
      ' Validate
      If (arTagged.Length = 0) Then Return False
      If (intTagged >= arTagged.Length) Then Return False
      ' We should be able to continue reading in the [arTagged] array for elements
      'While (intTagged < arTagged.Length) AndAlso ((arTagged(intTagged) Is Nothing) OrElse (Left(arTagged(intTagged), 1) = "_") _
      '                                             OrElse (arTagged(intTagged) = "") OrElse (InStr(arTagged(intTagged), "_t(f") > 0) _
      '                                             OrElse (InStr(arTagged(intTagged), "_f") > 0))
      While (intTagged < arTagged.Length) AndAlso ((arTagged(intTagged) Is Nothing) OrElse (Left(arTagged(intTagged), 1) = "_") _
                                                   OrElse (arTagged(intTagged) = ""))
        ' Skip unnecessary lines containing non-lexical data
        intTagged += 1
      End While
      ' Return this position
      Return intTagged
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GotoNextTagged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return intTagged
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetNextTagged
  ' Goal:   Get the next element on the [arTagged] array, starting from [intTagged]
  '         Return the element in [strLine]
  ' History:
  ' 29-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetNextTagged(ByRef arTagged() As String, ByRef intTagged As Integer, ByRef objInfl As Inflected, _
                                 ByVal strTarget As String, ByVal strLoc As String) As Boolean
    Dim arPart() As String      ' part of line
    Dim arMorph() As String     ' Morphology part
    Dim strWord As String = ""  ' Current word
    Dim strLine As String = ""  ' One line
    Dim strMorph As String = "" ' Morphology part
    Dim strCom As String = ""   ' Complementation type
    Dim intPos As Integer       ' Position

    Try
      ' Validate
      If (arTagged.Length = 0) Then Return False
      If (intTagged >= arTagged.Length) Then Return False
      ' Initialise: 
      With objInfl
        .ComType = "" : .Lemma = "" : .Word = "" : .WordType = "" : .GrCase = "" : .GrGender = "" : .GrNumber = "" : .Grperson = ""
        .Tense = "" : .Mood = ""
      End With
      ' We should be able to continue reading in the [arTagged] array for elements
      ' Also skip Foreign words
      GotoNextTagged(arTagged, intTagged)
      'While (intTagged < arTagged.Length) AndAlso ((arTagged(intTagged) Is Nothing) _
      '        OrElse (Left(arTagged(intTagged), 1) = "_") OrElse (InStr(arTagged(intTagged), "_t(f") > 0))
      '  ' Skip unnecessary lines containing non-lexical data
      '  intTagged += 1
      'End While
      ' Is anything left?
      If (intTagged < arTagged.Length) Then
        ' Get the lexical item
        strLine = arTagged(intTagged)
        ' Set the pointer to the next line
        intTagged += 1
        ' Find the lemma break
        intPos = InStr(strLine, "_l:")
        If (intPos = 0) Then
          ' Get the word by looking for an underscore
          intPos = InStr(strLine, "_")
          If (intPos = 0) Then
            strWord = strLine
          Else
            strWord = Left(strLine, intPos - 1)
          End If
          ' The only thing we can do is get the word back
          objInfl.Word = strWord
          ' Show we have a problem 
          TagComment("GetnextTagged: no lemma [" & strLine & "] (word=[" & strTarget & "] loc=[" & strLoc & "])")
          ' Return okay, even if there is no lemma
          Return True
        End If
        strWord = Left(strLine, intPos - 1) : objInfl.Word = strWord
        strLine = Mid(strLine, intPos + 3)
        ' Expecting "morphology" part
        intPos = InStr(strLine, "_m(")
        If (intPos = 0) Then TagComment("GetnextTagged: morphology missing (word=[" & strTarget & "] loc=[" & strLoc & "])") : Return False
        objInfl.Lemma = Left(strLine, intPos - 1)
        strLine = Mid(strLine, intPos + 3)
        ' Find the end of morphology
        intPos = InStr(strLine, ")")
        If (intPos = 0) Then TagComment("GetnextTagged: morphology end missing (word=[" & strTarget & "] loc=[" & strLoc & "])") : Return False
        ' Read morphology
        strMorph = Left(strLine, intPos - 1)
        strLine = Mid(strLine, intPos + 1)
        ' Is anything left?
        If (strLine <> "") Then
          ' Check if complementation type follows
          intPos = InStr(strLine, "_com:")
          If (intPos > 0) Then
            ' Get the complementation type start
            strLine = Mid(strLine, intPos + 5)
            objInfl.ComType = Left(strLine, 1)
          End If
        End If
        ' Split the morphology part up
        arPart = Split(strMorph, "_")
        ' Parse the morphology
        For inti = 0 To arPart.Length - 1
          arMorph = Split(arPart(inti), ":")
          Select Case arMorph(0)
            Case "c"  ' Case
              objInfl.GrCase = arMorph(1)
            Case "w"  ' Word type (links to POS)
              objInfl.WordType = arMorph(1)
            Case "n"  ' Number
              objInfl.GrNumber = arMorph(1)
            Case "p"  ' Person
              objInfl.Grperson = arMorph(1)
            Case "g"  ' Gender
              objInfl.GrGender = arMorph(1)
            Case "p1" ' First person
              objInfl.Grperson = "1"
            Case "p2" ' Second person
              objInfl.Grperson = "2"
            Case "p3" ' Third person
              objInfl.Grperson = "3"
            Case "t"  ' Tense
              objInfl.Tense = arMorph(1)
            Case "m"  ' Mood
              objInfl.Mood = arMorph(1)
            Case Else
              TagComment("GetNextTagged: unknown morphology part [" & arMorph(0) & "] (word=[" & strTarget & "] loc=[" & strLoc & "])")
              Return False
          End Select
        Next inti
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetNextTagged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetTaggedFile
  ' Goal:   Read the OE-tagged file(s) with the indicated [strCamno] into an array
  ' History:
  ' 29-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetTaggedFile(ByVal strCamno As String, ByRef arTagged() As String, ByRef strOEfile As String) As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim strFile As String     ' Filename
    Dim strText As String     ' Text
    Dim strDelim As String    ' Delimiter
    Dim arTemp() As String    ' Temporary copy
    Dim intI As Integer       ' Counter
    Dim intPos As Integer     ' Position to add the new array

    Try
      ' validate
      If (strCamno = "") Then Return False
      ' Try get the result
      dtrFound = tdlMorphTag.Tables("TorCam").Select("Camno='" & strCamno & "'", "File ASC")
      If (dtrFound.Length = 0) Then Return False
      ' Initialise the [arTagged] array
      ReDim arTagged(0) : strOEfile = ""
      ' Walk through all the files
      For intI = 0 To dtrFound.Length - 1
        ' Get this file name
        strFile = dtrFound(intI).Item("File")
        ' Do we actually have a file?
        If (strFile = "") Then Return False
        ' Check existence of file
        If (Not IO.File.Exists(strFile)) Then TagComment("GetTaggedFile: could not find " & strFile) : Return False
        ' Add filename to stack
        AddSemiStack(strOEfile, IO.Path.GetFileNameWithoutExtension(strFile))
        ' Read the file
        strText = IO.File.ReadAllText(strFile)
        ' Transform into array
        strDelim = GetDelim(strText, vbCrLf, vbCr, vbLf)
        arTemp = Split(strText, strDelim)
        ' Find the position
        intPos = IIf(arTagged.Length = 1, 0, arTagged.Length)
        ' Make sure there is enough room
        ReDim Preserve arTagged(0 To arTagged.Length + arTemp.Length - 1)
        ' Add to existing array
        Array.Copy(arTemp, 0, arTagged, intPos, arTemp.Length)
      Next intI
      ' Debug.Print(arTagged(10980), arTagged(1231), arTagged(1232))
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OneTaggedFeaturesPdx error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   OneDoFeaturesPdx
  ' Goal:   Add the types for this category at the necessary constituents in the XML [pdxFile] document
  ' History:
  ' 26-05-2010  ERK Created
  ' 01-09-2011  ERK Added CAT_NPTYPE
  ' 07-11-2012  ERK Added CAT_COGN
  ' 15-01-2013  ERK Added [arFeatDef] in verband met CAT_VB voor unaccusative marking
  ' 17-01-2013  ERK Added processing of CAT_PNORM
  ' 23-01-2013  ERK Changed CAT_VB processing to use tblFeat
  ' 10-02-2014  ERK Addec CAT_VBUNACC processing
  ' ------------------------------------------------------------------------------------
  Public Function OneDoFeaturesPdx(ByRef pdxFile As XmlDocument, ByVal strCategory As String, _
          Optional ByVal bRecalculate As Boolean = False) As Boolean
    Dim intI As Integer               ' Counter
    Dim intPtc As Integer             ' Percentage
    Dim intIPdist As Integer          ' Distance from source to target in IP boundaries
    Dim intCount As Integer           ' Number of child nodes
    Dim intJ As Integer               ' Counter
    Dim ndxFor As XmlNode = Nothing   ' Forest node
    Dim ndxThis As XmlNode = Nothing  ' Current node
    Dim ndxForP As XmlNode = Nothing  ' Parallel forest node
    Dim strNPtype As String           ' The actual NP type
    Dim strPGN As String              ' The Person/Gender/Number for this NP
    Dim strGrRole As String           ' The grammatical role
    Dim strVbType As String = ""      ' Verb type
    Dim strPlemma As String = ""      ' Lemma of preposition
    Dim strCogState As String         ' Cognitive status
    Dim strNewNPtype As String = ""   ' Newly determined NP type (NOT USED YET)
    Dim strNewPGN As String = ""      ' Newly determined PGN     (NOT USED YET)
    Dim strNewGrRole As String = ""   ' Newly determined GrRole  (NOT USED YET)
    Dim strAdvType As String = ""     ' The type of adverb
    Dim strPeriod As String = ""      ' The period determined for this file
    Dim strLabel As String            ' The label of the current node we are looking at
    Dim strSearch As String = ""      ' Search criteria for this category
    Dim strSearchP As String = ""     ' Search criteria for this category
    Dim strValue As String            ' Value of a feature that is already present
    Dim strLog As String = ""         ' What we log
    Dim strLoc As String = ""         ' Location
    Dim strForHtml As String = ""     ' Text of the forest in html
    Dim colBack As New StringColl     ' Report
    Dim arPgn() As String             ' Array of PGN values
    Dim colDone As New NodeColl       ' Collection of nodes we have treated already
    Dim bSbjOnly As Boolean = False    ' Only do subjects for animacy?
    Try
      ' Determine the period for this file
      strPeriod = GetPeriod(pdxFile)
      strCurrentPeriod = strPeriod
      ' Set the current file
      pdxCurrentFile = pdxFile
      ' Show hwere we are
      Select Case strCategory
        Case "Animacy"
          Status("Adding features for animacy...")
        Case Else
          Status("Adding features for the " & strCategory & "s...")
      End Select
      ' Determine the search criteria for this category
      Select Case strCategory
        Case CAT_ADV
          ' We look for adverbs that are part of an NP or of a PP
          strSearch = "NP|NP-*|PP|PP-*"
        Case CAT_VB, CAT_VBUNACC, CAT_VBTYPE
          ' We look for all kinds of verbs, and they may be auxiliary or non-auciliary
          ' (Perl file [change_verb_tags_ME.pl] states tag may start with V or B)
          ' strSearch = "V*|B*" : strSearchP = "V*|B*|U-V*|U-B*"
          strSearch = "VA*|VB*|BA*|BE*|HV*|AX*|MD*|DO*|DA*|*+VB*|*+BE*|*+HV*|*+AX*|*+MD*|*+DO*|*+DA*"
          strSearchP = "VA*|VB*|BA*|BE*|HV*|AX*|MD*|DO*|DA*|U-VA*|U-VB*|U-BA*|U-BE*|U-HV*|U-AX*|U-MD*|U-DO*|U-DA*|*+VB*|*+BE*|*+HV*|*+AX*|*+MD*|*+DO*|*+DA*"
        Case CAT_PNORM
          ' Look for all kinds of prepositions
          strSearch = "P"
        Case CAT_NP, CAT_NPTYPE, CAT_ANIM
          ' We look to label all necessary NPs
          strSearch = strNPsourceTypes
        Case CAT_WH
          ' We look for all kinds of WH NPs
          strSearch = "WNP*"
        Case CAT_GR
          ' We look to label all necessary NPs
          strSearch = strNPsourceTypes
          ' Start report
          colBack.Add("<html><body><h3>Grammatical role changes</h3>")
        Case CAT_ANIM
          colHeadAnim.Clear() : colHeadInanim.Clear()
        Case CAT_COGN
          ' We look to label all necessary NPs
          strSearch = strNPsourceTypes
          ' Start report
          colBack.Add("<html><body><h3>Cognitive states</h3>")
          colBack.Add("<table>")
        Case Else
          ' This category is not implemented!
          Return False
      End Select
      ' Look for the first <forest> element
      If (Not GetFirstForest(pdxFile, ndxFor)) Then Return False
      Select Case strCategory
        Case CAT_VBUNACC
          If (Not GetFirstForest(pdxPara, ndxForP)) Then Return False
      End Select
      ' Determine how many elements there are potentially
      If (ndxFor Is Nothing) Then
        intCount = 0
      ElseIf (ndxFor.ParentNode Is Nothing) Then
        intCount = 0
      Else
        intCount = ndxFor.ParentNode.ChildNodes.Count
      End If
      Select Case strCategory
        Case CAT_ANIM
          ' We need to walk the forests BACKWARDS
          ndxFor = ndxFor.ParentNode.LastChild
          While (Not ndxFor Is Nothing)
            ' Check where we are
            intPtc = (intCount - ndxFor.Attributes("forestId").Value) * 100 \ intCount
            Status("Checking " & strCategory & " features " & intPtc & "%", intPtc)
            ' Process the <eTree> elements in this forest recursively backwards
            If (Not DoEtreeBack(ndxFor, colDone, CAT_ANIM, intPtc, bSbjOnly)) Then Return False
            ' Go to previous <forest> element
            ndxFor = ndxFor.PreviousSibling
          End While
        Case Else
          ' Walk through all the <forest> elements FORWARD
          While (Not ndxFor Is Nothing)
            ' Show where we are (how do we KNOW where we are?)
            intPtc = (ndxFor.Attributes("forestId").Value) * 100 \ intCount
            Status("Adding " & strCategory & " features " & intPtc & "%", intPtc)
            ' Check if we are not being interrupted
            If (bInterrupt) Then Return False
            ' Find all the necessary elements in this forest
            ndxList = ndxFor.SelectNodes(".//eTree[tb:Like(string(@Label), '" & strSearch & "')]", conTb)
            ' Some categories may still need different treatment here...
            Select Case strCategory
              Case CAT_VBUNACC
                ndxListP = ndxForP.SelectNodes(".//eTree[tb:Like(string(@Label), '" & strSearchP & "')]", conTb)
                ' Double check
                If (ndxListP.Count <> ndxList.Count) Then Stop
                ' ndxListP = ndxFor.SelectNodes(".//eTree[tb:Like(string(@Label), '" & strSearch & "')]", conTb)
            End Select
            ' Visit all these elements
            For intI = 0 To ndxList.Count - 1
              ' Get this element
              ndxThis = ndxList(intI)
              ' Action depends on the feature category we are processing
              Select Case strCategory
                Case CAT_ADV
                  ' For adverbs we need to have ALL the children of type ADV and FP
                  ndxThis = ndxThis.FirstChild
                  While (Not ndxThis Is Nothing)
                    ' Is this an <eTree>?
                    If (ndxThis.Name = "eTree") Then
                      ' Check this child
                      If (DoLike(ndxThis.Attributes("Label").Value, "ADV|ADV-*|FP")) Then
                        ' We found an Adverb part of an NP or a PP
                        ' (1) Get the Adverb Type
                        If (GetAdvType(ndxThis, strPeriod, strAdvType)) Then
                          ' (2) Check the result
                          If (strAdvType <> "") Then
                            ' (3) Check if a feature is already present
                            strValue = GetFeature(ndxThis, "Adv", "AdvType")
                            If (strValue = "") OrElse (strValue = "(empty)") Then
                              ' (4) Add the adverb type in the file
                              If (Not AddFeature(pdxFile, ndxThis, "Adv", "AdvType", strAdvType)) Then Return False
                            End If
                          End If
                        Else
                          ' Show there is a problem
                          Status("Unable to get the Adverb type from:" & vbCrLf & NodeInfo(ndxThis))
                          ' Return failure
                          Return False
                        End If
                      End If
                    End If
                    ' Go to the next child
                    ndxThis = ndxThis.NextSibling
                  End While
                Case CAT_PNORM  ' Add normalized (lemmatized) form of the preposition
                  ' Get the current preposition lemma
                  strPlemma = GetFeature(ndxThis, "P", "lemma")
                  If (bRecalculate) OrElse (strPlemma = "") OrElse (strPlemma = "unknown") Then
                    ' Determine the preposition lemma
                    If (GetPlemma(ndxThis, strPlemma)) Then
                      ' If result is positive, then process it
                      If (strPlemma <> "") Then
                        ' If we have [bRecalculate] and the feature differs from what it is, then check
                        If (bRecalculate) AndAlso (strPlemma <> GetFeature(ndxThis, "P", "lemma")) Then
                          ' Okay, this is a difference - do we need to ask permission, or just signal the difference?
                          If (Not GetForestLoc(ndxFor, strLoc)) Then Return False
                          strLog = "[" & strLoc & "] - " & GetFeature(ndxThis, "P", "lemma") & _
                                   " --> " & strPlemma & " [" & NodeText(ndxThis, True) & "]<br>"
                          colBack.Add(strLog)
                        End If
                        ' See if we can process the P-lemma information
                        If (Not AddFeature(pdxFile, ndxThis, "P", "lemma", strPlemma)) Then Return False
                      End If
                    Else
                      ' Show there is a problem
                      Status("Unable to get the preposition lemma from EtreeId=" & ndxThis.Attributes("Id").Value)
                      ' Return failure
                      Return False
                    End If
                  End If
                Case CAT_VB, CAT_VBTYPE     ' Add marking to indicate whether this verb is unaccusative or not
                  ' Get the current verb type
                  strVbType = GetFeature(ndxThis, "Vb", "VbType")
                  ' ========== DEBUG =========
                  ' If (strVbType = "unacc") Then Stop
                  ' ==========================
                  If (bRecalculate) OrElse (strVbType = "") OrElse (strVbType = "unknown") OrElse (strVbType = "other") Then
                    ' Determine the verb type
                    If (GetVbType(ndxThis, strVbType, strCategory)) Then
                      ' If result is positive, then process it
                      If (strVbType <> "") Then      ' USED TO HAVE: AndAlso (Not bRecalculate OrElse strVbType <> "other") Then
                        ' ============== Debugging ============
                        ' If (strVbType <> "empty") Then Stop
                        ' =====================================
                        ' If we have [bRecalculate] and the feature differs from what it is, then check
                        If (bRecalculate) AndAlso (strVbType <> GetFeature(ndxThis, "Vb", "VbType")) Then
                          ' Okay, this is a difference - do we need to ask permission, or just signal the difference?
                          If (Not GetForestLoc(ndxFor, strLoc)) Then Logging("OneDoFeaturesPdx: Cat_Vb problem") : Return False
                          strLog = "[" & strLoc & "] - " & GetFeature(ndxThis, "Vb", "VbType") & _
                                   " --> " & strVbType & " [" & NodeText(ndxThis, True) & "]<br>"
                          colBack.Add(strLog)
                        End If
                        ' ========= DEBUG
                        ' If (NodeText(ndxThis) <> "") Then Stop
                        ' ===============
                        ' See if we can process the NPtype information
                        If (Not AddFeature(pdxFile, ndxThis, "Vb", "VbType", strVbType)) Then Logging("OneDoFeaturesPdx: Cat_Vb problem adding feature") : Return False
                      End If
                    Else
                      ' Show there is a problem
                      Status("Unable to get the verb type from EtreeId=" & ndxThis.Attributes("Id").Value)
                      ' Return failure
                      Return False
                    End If
                  End If
                  ' Also do away with the feature M/VBtype
                  If (Not DelFeature(ndxThis, "M", "VbType")) Then Return False
                Case CAT_VBUNACC     ' Add marking to indicate whether this verb is unaccusative or not
                  ' Get the verb type from the parallel one
                  strVbType = GetFeature(ndxListP(intI), "Vb", "VbType")
                  If (strVbType = "") Then
                    ' Give it a default value
                    strVbType = "other"
                    ' Check if it should be unaccusative
                    If (ndxListP(intI).Attributes("Label").Value Like "U-*") Then strVbType = "unacc"
                  Else
                    ' We already seem to have a verb type, and that is okay
                    If (strVbType = "unacc") Then
                      ' Debugging
                      ' Stop
                    End If
                  End If
                  ' Store it
                  If (Not AddFeature(pdxFile, ndxThis, "Vb", "VbType", strVbType)) Then Return False
                Case CAT_GR   ' Only check and update the Grammatical Role feature
                  ' See if the Grammatical Role feature is there already
                  strGrRole = GetFeature(ndxThis, "NP", "GrRole")
                  If (bRecalculate) OrElse (strGrRole = "") OrElse (strGrRole = "unknown") Then
                    ' Get the Grammatical role
                    If (GetGrRole(ndxThis, strGrRole)) Then
                      ' If result is positive, then process it
                      If (strGrRole <> "") Then
                        ' If we have [bRecalculate] and the feature differs from what it is, then check
                        If (bRecalculate) AndAlso (strGrRole <> GetFeature(ndxThis, "NP", "GrRole")) Then
                          ' Okay, this is a difference - do we need to ask permission, or just signal the difference?
                          If (Not GetForestLoc(ndxFor, strLoc)) Then Return False
                          strLog = "[" & strLoc & "] - " & GetFeature(ndxThis, "NP", "GrRole") & _
                                   " --> " & strGrRole & " [" & NodeText(ndxThis, True) & "]<br>"
                          colBack.Add(strLog)
                        End If
                        ' See if we can process the NPtype information
                        If (Not AddFeature(pdxFile, ndxThis, "NP", "GrRole", strGrRole)) Then Return False
                      End If
                    Else
                      ' Show there is a problem
                      Status("Unable to get the Grammatical Role from Id=" & ndxThis.Attributes("Id").Value)
                      ' Return failure
                      Return False
                    End If
                  End If
                Case CAT_COGN   ' Check and update the Cognitive Status (GHZ 1993)
                  ' See if the feature is already there
                  strCogState = GetFeature(ndxThis, "NP", "CgState")
                  If (bRecalculate) OrElse (strCogState = "") OrElse (strCogState = "unknown") Then
                    ' Determine the cognitive state
                    If (GetCogState(ndxThis, strCogState)) Then
                      ' If result is positive, then process it
                      If (strCogState <> "") Then
                        ' If we have [bRecalculate] and the feature differs from what it is, then check
                        If (bRecalculate) AndAlso (strCogState <> GetFeature(ndxThis, "NP", "CgState")) Then
                          '' Okay, this is a difference - do we need to ask permission, or just signal the difference?
                          'If (Not GetForestLoc(ndxFor, strLoc)) Then Return False
                          'strLog = "[" & strLoc & "] - " & GetFeature(ndxThis, "NP", "CgState") & _
                          '         " --> " & strCogState & " [" & NodeText(ndxThis, True) & "]<br>"
                          'colBack.Add(strLog)
                        End If
                        ' See if we can process the NPtype information
                        If (Not AddFeature(pdxFile, ndxThis, "NP", "CgState", strCogState)) Then Return False
                      End If
                    Else
                      ' Show there is a problem
                      Status("Unable to get the cognitive status from Id=" & ndxThis.Attributes("Id").Value)
                      ' Return failure
                      Return False
                    End If
                  End If
                Case CAT_NPTYPE
                  ' See if the NP type feature is there already
                  strNPtype = GetFeature(ndxThis, "NP", "NPtype")
                  ' Should we recalculate?
                  If (bRecalculate) OrElse (strNPtype = "") Then
                    ' Set the NPtype and the PGN to "unknown"
                    strNPtype = "unknown" : strPGN = "unknown"
                    ' Get the NP type
                    If (GetNpType(ndxThis, strPeriod, strNPtype, strPGN)) Then
                      ' If result is positive, then process it
                      If (strNPtype <> "") Then
                        ' See if we can process the NPtype information
                        If (Not AddFeature(pdxFile, ndxThis, "NP", "NPtype", strNPtype)) Then Return False
                      End If
                    Else
                      ' Show there is a problem
                      Logging("Unable to get the NP type from Id=" & ndxThis.Attributes("Id").Value)
                      ' Return failure
                      Return False
                    End If
                  End If
                Case CAT_NP, CAT_WH
                  ' See if the NP type feature is there already
                  strNPtype = GetFeature(ndxThis, "NP", "NPtype")
                  strPGN = GetFeature(ndxThis, "NP", "PGN")
                  If (bRecalculate) OrElse (strNPtype = "") OrElse (strPGN = "") OrElse (strNPtype = "unknown") _
                     OrElse (strPGN = "unknown") Then
                    ' Set the NPtype and the PGN to "unknown"
                    strNPtype = "unknown" : strPGN = "unknown" : strGrRole = "unknown" : intIPdist = -1
                    ' Get the NP type
                    If (GetNpType(ndxThis, strPeriod, strNPtype, strPGN)) Then
                      ' If result is positive, then process it
                      If (strNPtype <> "") Then
                        ' See if we can process the NPtype information
                        If (Not AddFeature(pdxFile, ndxThis, "NP", "NPtype", strNPtype)) Then Return False
                      End If
                      ' Do we have a PGN result?
                      If (strPGN <> "") Then
                        ' Just get the label
                        strLabel = ndxThis.Attributes("Label").Value
                        ' Check if the PGN is not "unknown", and if we should ask the user or not...
                        If (DoLike(strLabel, "PRO|PRO$|PRO^*|PRO$^*")) AndAlso (strPGN = "unknown") AndAlso _
                            (bUserProPGN) Then
                          ' Ask user to adapt the PGN
                          If (UserAdaptPGN(ndxThis, strPGN, False)) Then
                            ' =========== DEBUG
                            'If (InStr(strPGN, ";") > 0) Then Stop
                            ' ====================
                            ' Transform into array of values
                            arPgn = Split(strPGN, ";")
                            ' Add all values from the array
                            For intJ = 0 To UBound(arPgn)
                              ' The adapted value is in [strPGN], but what about the tdlSettings?
                              If (Not AddPronounPGN(strNPtype, arPgn(intJ), NodeText(ndxThis, True))) Then
                                ' Don't do anything in case of failure -- nothing goes wrong basically...
                              End If
                            Next intJ
                          End If
                        End If
                        ' ================= DEBUG ===========
                        If (strPGN = "2empty") Then Stop
                        ' ===================================
                        ' See if we can process the PGN information
                        If (Not AddFeature(pdxFile, ndxThis, "NP", "PGN", strPGN)) Then Return False
                      End If
                    Else
                      ' Show there is a problem
                      Status("Unable to get the NP type from Id=" & ndxThis.Attributes("Id").Value)
                      ' Return failure
                      Return False
                    End If
                  End If
                  ' See if the Grammatical Role feature is there already
                  strGrRole = GetFeature(ndxThis, "NP", "GrRole")
                  If (bRecalculate) OrElse (strGrRole = "") OrElse (strGrRole = "unknown") Then
                    ' Get the Grammatical role
                    If (GetGrRole(ndxThis, strGrRole)) Then
                      ' If result is positive, then process it
                      If (strGrRole <> "") Then
                        ' See if we can process the NPtype information
                        If (Not AddFeature(pdxFile, ndxThis, "NP", "GrRole", strGrRole)) Then Return False
                      End If
                    Else
                      ' Show there is a problem
                      Status("Unable to get the Grammatical Role from Id=" & ndxThis.Attributes("Id").Value)
                      ' Return failure
                      Return False
                    End If
                  End If
                  ' Get the IP distance
                  If (GetIpDist(ndxThis, intIPdist)) Then
                    ' Only process this result if it is not too negative
                    If (intIPdist >= -1) Then
                      ' See if we can process the IP distance information
                      If (Not AddFeature(pdxFile, ndxThis, "coref", "IPdist", intIPdist)) Then Return False
                    End If
                  Else
                    ' Show there is a problem
                    Status("Unable to get the IP distance at Id=" & ndxThis.Attributes("Id").Value)
                    ' Return failure
                    Return False
                  End If
              End Select
            Next intI
            ' Get all <eTree> descendants, and remove their IP/Number feature
            ndxList = ndxFor.SelectNodes(".//eTree")
            For intI = 0 To ndxList.Count - 1
              ' Check and remove feature
              If (HasFeature(ndxList(intI), "IP", "Number")) Then
                DelXmlNodeChild(ndxList(intI), "fs", "type;IP")
              End If
            Next intI
            ' Forest processing for some categories
            Select Case strCategory
              Case CAT_COGN
                ' I would like to have one line for the forest as output, where each of the constituents 
                '   with a particular cognitive state is shown with the status ID and a particular color
                strForHtml = ""
                If (Not TravNode(ndxFor, "CgnState", strForHtml)) Then Return False
                If (strForHtml <> "") Then
                  colBack.Add("<tr><td>" & ndxFor.Attributes("forestId").Value & "</td><td>" & _
                              strForHtml & "</td></tr>")
                  ' Do we have a back translation?
                  If (Not ndxFor.SelectSingleNode("./div[@lang='eng']") Is Nothing) Then
                    ' Add the backtranslation
                    colBack.Add("<tr><td></td><td>" & XmlString(ndxFor.SelectSingleNode("./div[@lang='eng']").InnerText) & "</td></tr>")
                  End If
                End If
            End Select
            ' Go to the next forest
            ndxFor = ndxFor.NextSibling
            Select Case strCategory
              Case CAT_VBUNACC
                ndxForP = ndxForP.NextSibling
            End Select
          End While
      End Select

      ' Make finishing stuff for this category
      Select Case strCategory
        Case CAT_GR, CAT_NPTYPE, CAT_ADV, CAT_NP, CAT_WH, CAT_ANIM, CAT_COGN, CAT_VB, CAT_VBTYPE, CAT_VBUNACC, CAT_PNORM
          ' Finish report
          colBack.Add("</body></html>")
          ' Put this on the report page
          frmMain.wbReport.DocumentText = colBack.Text
        Case Else
          ' This category is not implemented!
          Return False
      End Select
      ' REturn success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OneDoFeaturesPdx error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DoEtreeBack
  ' Goal:   Recursive function to go backwards, finding <eTree> elements and acting upon them
  ' History:
  ' 16-04-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function DoEtreeBack(ByRef ndxThis As XmlNode, ByRef colDone As NodeColl, ByVal strType As String, _
                               ByVal intPtc As Integer, ByVal bSbjOnly As Boolean) As Boolean
    Dim ndxChild As XmlNode ' Child node

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Check for interrupt
      If (bInterrupt) Then Return False
      ' First check all its children - last to first
      ndxChild = ndxThis.LastChild
      While (Not ndxChild Is Nothing) And (Not bInterrupt)
        ' Process this child
        If (Not DoEtreeBack(ndxChild, colDone, strType, intPtc, bSbjOnly)) Then Return False
        ndxChild = ndxChild.PreviousSibling
      End While
      ' Check what we have
      Select Case ndxThis.Name
        Case "eTree"
          ' Perform the action for this item
          Select Case strType
            Case CAT_ANIM
              If (Not OneChainAnim(ndxThis, colDone, intPtc, bSbjOnly)) Then Return False
            Case Else
              ' Nothing else implemented yet
          End Select
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/DoEtreeBack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   OneChainAnim
  ' Goal:   Figure out the animacy of this one chain
  ' History:
  ' 16-04-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function OneChainAnim(ByRef ndxThis As XmlNode, ByRef colDone As NodeColl, _
                                ByVal intPtc As Integer, ByVal bSbjOnly As Boolean) As Boolean
    Dim ndxWork As XmlNode              ' Working node
    Dim ndxForest As XmlNode = Nothing  ' Forest node
    Dim strLastText As String = ""      ' Last node text
    Dim strRefType As String            ' The kind of reference we have here
    Dim strGrRole As String             ' Grammatical role
    Dim strAnimacy As String = ""       ' Value for the animacy: '-', 'a', 'i'
    Dim strPgn As String = ""           ' Person/Gender/Number
    Dim strNPtype As String = ""        ' The NP type of me
    Dim strHeadN As String = ""         ' Head noun
    Dim strLabel As String              ' My label
    Dim strHdLabel As String            ' Head label
    Dim strFtAnim As String             ' Current value of the feature "anim"
    Dim colChain As New StringColl      ' The text of this chain
    Dim dtrChain As DataRow = Nothing   ' Datarow for this chain
    Dim dtrItem As DataRow = Nothing    ' One datarow for an item
    Dim intChainId As Integer = 0       ' ID for this chain
    Dim intItemId As Integer = 0        ' ID for this item
    Dim intLen As Integer = 0           ' Length of this chain
    Dim intNoTraceLen As Integer = 0    ' Length of this chain excluding traces
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
          ' Do we need to check for subjects only?
          If (bSbjOnly) Then
            ' Check subjecthood
            strGrRole = GetFeature(ndxThis, "NP", "GrRole")
            If (strGrRole <> "Subject") Then Return True
          End If
          ' Check if this node has already been processed
          If (colDone.Exists(ndxThis)) Then Return True
          ' Check if the animacy is already known for this first item on the chain
          strFtAnim = GetFeature(ndxThis, "NP", "anim")
          If (strFtAnim = "") OrElse (strFtAnim = "-") Then
            ' Go into a loop for the FIRST time, trying to figure out the animacy...
            strAnimacy = "-" : colChain.Clear() : colChain.Add("<html><body>")
            ndxWork = ndxThis
            ' Get the first head noun
            If (GetHeadNoun(ndxWork, strHeadN)) Then
              colChain.Add("Head=<font color='orange'>" & VernToEnglish(strHeadN) & "</font><br>")
              strHdLabel = GetHeadNounNode(ndxWork).Attributes("Label").Value
            Else
              colChain.Add("(head undetermined)<br>")
              strHdLabel = ""
            End If
            While (Not ((ndxWork Is Nothing) OrElse (colDone.Exists(ndxWork)) OrElse bStop))
              ' Is the animacy already known?
              If (strAnimacy = "-") Then
                ' Get my label
                strLabel = ndxWork.Attributes("Label").Value
                ' Get the PGN feature
                strPgn = GetFeature(ndxWork, "NP", "PGN")
                If (DoLike(strPgn, "1*|2*|3m*|3f*")) Then
                  ' The chain must be animate
                  strAnimacy = "a"
                ElseIf (DoLike(strPgn, "3ns")) Then
                  ' The chain must be inanimate
                  strAnimacy = "i"
                ElseIf (DoLike(strLabel, "CP*|IP*|*MSR*|*ADV*|*TMP*")) Then
                  ' The chain must be inanimate
                  strAnimacy = "i"
                ElseIf (DoLike(strHdLabel, "ADJ*")) Then
                  ' The chain must be inanimate
                  strAnimacy = "i"
                ElseIf (DoLike(strLabel, "*VOC*")) Then
                  ' The chain must be inanimate
                  strAnimacy = "a"
                Else
                  ' Get the NPtype
                  strNPtype = GetFeature(ndxWork, "NP", "NPtype")
                  ' Determine animacy based on the NP type??
                  If (DoLike(strNPtype, "NumNP")) Then
                    ' This must be inanimate
                    strAnimacy = "i"
                  Else
                    ' Look for special cases
                    If (Not ndxWork.SelectSingleNode("./child::eTree[tb:Like(@Label, 'NUM*|CP-FRL*|IP-PPL*')]", conTb) Is Nothing) Then
                      ' This must be inanimate
                      strAnimacy = "i"
                    ElseIf (Not ndxWork.SelectSingleNode("./child::eTree[tb:Like(@Label, 'ONE|OTHERS')]", conTb) Is Nothing) Then
                      ' This must be animate
                      strAnimacy = "a"
                    Else
                      ' Get the head noun (if possible)
                      If (GetHeadNoun(ndxWork, strHeadN)) Then
                        ' Check if it is on the list
                        If (colHeadAnim.Exists(strHeadN)) Then
                          strAnimacy = "a"
                        ElseIf (colHeadInanim.Exists(strHeadN)) Then
                          strAnimacy = "i"
                        ElseIf (DoLike(strHeadN, "*ion|*ing")) Then
                          strAnimacy = "i"
                        End If
                      End If
                    End If
                  End If
                End If
              Else
                ' We can just as well stop!
                bStop = True
              End If
              ' Last resort: ask
              If (strAnimacy = "-") Then
                ' Give information about this one
                ' strLastText = NodeText(ndxWork, True)
                strLastText = NodeInfoPlusChildren(ndxWork)
                ' Put information on a string
                colChain.Add(strLastText & " - <font color='red'>" & strRefType & "</font>[" & _
                             "<font color='green'>" & strNPtype & "</font>]<br>")
                '' Get the forestnode of me
                'If (Not GetForestNode(ndxWork, ndxForest)) Then Return False
              End If
              ' Note our current Reference type
              strRefType = CorefAttr(ndxWork, "RefType")
              ' Calculate whether we need to stop
              bStop = (bStop) OrElse (Not DoLike(strRefType, "Identity|CrossSpeech"))
              ' Go to the next antecedent
              ndxWork = CorefDst(ndxWork)
            End While
            ' Last resort animacy: Identity points to IP
            If (strAnimacy = "-") AndAlso (strRefType = "Identity") AndAlso (Not ndxWork Is Nothing) Then
              ' Check this label
              If (ndxWork.Attributes("Label").Value Like "IP*") Then strAnimacy = "i"
            End If
            ' Did we retrieve the animacy?
            If (strAnimacy = "-") Then
              ' See if one element needs to be added to the chain
              If (bStop AndAlso (DoLike(strRefType, "Identity|CrossSpeech|Inferred"))) Then
                ' Go to the last node
                ndxWork = CorefDst(ndxWork)
                ' Show the information of the last node
                strLastText = NodeInfoPlusChildren(ndxWork)
                ' Put information on a string
                colChain.Add(strLastText & " - <font color='red'>" & strRefType & "</font>[" & _
                             "<font color='green'>" & strNPtype & "</font>]<br>")
              End If
              ' Make a larger line of the first one
              colChain.Add(NodeInfoPlusChildren(ndxThis.SelectSingleNode("./ancestor-or-self::forest[1]"), _
                                                VernToEnglish(strHeadN)) & "<br>")
              ' See if there is any PDE translation of this line
              ndxWork = ndxThis.SelectSingleNode("./ancestor-or-self::forest[1]/div[@lang='eng']/seg")
              If (Not ndxWork Is Nothing) Then
                colChain.Add("<font color='green'>" & ndxWork.InnerText & "</font><br>")
              End If
              ' Finish the chain
              colChain.Add("</body></html>")
              ' Ask user for the animacy
              With frmAnimacy
                ' Set the chain
                .Chain = colChain.Text
                .Percentage = intPtc
                ' Ask user
                Select Case .ShowDialog
                  Case DialogResult.OK
                    ' Get the animacy value
                    strAnimacy = .Animacy
                  Case DialogResult.Cancel
                    ' Set interrupt
                    bInterrupt = True
                    Return False
                End Select
              End With
            End If
          End If
          ' Go into a loop for the SECOND time, filling in the animacy for all the members
          ndxWork = ndxThis : bStop = False
          While (Not ((ndxWork Is Nothing) OrElse (colDone.Exists(ndxWork)) OrElse bStop))
            ' Is animacy already determined?
            strFtAnim = GetFeature(ndxWork, "NP", "anim")
            If (strFtAnim = "") OrElse (strFtAnim = "-") Then
              ' Add the animacy feature
              If (Not AddFeature(pdxCurrentFile, ndxWork, "NP", "anim", strAnimacy)) Then Return False
              ' Indicate we have changed
              bNeedSaving = True
              ' Check if we need to put the head noun on the list
              If (strAnimacy <> "-") Then
                ' Get the head noun
                If (GetHeadNoun(ndxWork, strHeadN)) Then
                  ' Put it the list
                  If (strAnimacy = "a") Then
                    colHeadAnim.AddUnique(strHeadN)
                  Else
                    colHeadInanim.AddUnique(strHeadN)
                  End If
                End If
              End If
            Else
              If (GetHeadNoun(ndxWork, strHeadN)) Then
                Select Case GetFeature(ndxWork, "NP", "anim")
                  Case "a"
                    colHeadAnim.AddUnique(strHeadN)
                  Case "i"
                    colHeadInanim.AddUnique(strHeadN)
                End Select
              End If
            End If
            ' File this node
            colDone.AddUnique(ndxWork)
            ' Note our current Reference type
            strRefType = CorefAttr(ndxWork, "RefType")
            ' Calculate whether we need to stop
            bStop = (Not DoLike(strRefType, "Identity|CrossSpeech"))
            ' Go to the next antecedent
            ndxWork = CorefDst(ndxWork)
          End While
        Case ""
          ' No need to take action!!
        Case Else
          ' This should not occur, but may happen with other files -- just signal this
          Logging("modAdapt/OneChainAnim unknown reference type = " & strRefType & " in: " & _
                  ndxThis.SelectSingleNode("./ancestor-or-self::forest[1]").Attributes("File").Value)
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OneChainAnim error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitProTrack
  ' Goal:   Initialise the Pronoun Tracking collection
  ' History:
  ' 28-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub InitCatTrack()
    Try
      ' Does a table already exist?
      If (tblCatTrack Is Nothing) Then
        ' Does not exist, so create one
        tblCatTrack = New DataTable("CatTrack")
        ' Create the columns
        tblCatTrack.Columns.Add("Category", Type.GetType("System.String"))
        tblCatTrack.Columns.Add("Type", Type.GetType("System.String"))
        tblCatTrack.Columns.Add("Period", Type.GetType("System.String"))
        tblCatTrack.Columns.Add("Location", Type.GetType("System.String"))
        tblCatTrack.Columns.Add("Example", Type.GetType("System.String"))
      Else
        ' Already exists, so only reset the rows
        tblCatTrack.Rows.Clear()
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/InitProTrack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   HasProTrack
  ' Goal:   Whether there are any ProTrack entries
  ' History:
  ' 28-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function HasCatTrack() As Boolean
    Try
      Return (tblCatTrack.Rows.Count > 0)
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/HasCatTrack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddProTrack
  ' Goal:   Add one item to the Pronount Tracking set
  '         Each item consists of "Pronoun;Type;Period"
  '         Each example consists of "Location;Example"
  ' History:
  ' 31-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function AddCatTrack(ByVal strPro As String, ByVal strType As String, ByVal strPeriod As String, _
                         ByRef ndxThis As XmlNode) As Boolean
    Dim rowThis As DataRow          ' one new row
    Dim dtrFound() As DataRow       ' Result of searching
    Dim strLocation As String = ""  ' Location of example
    Dim strExample As String        ' Actual example

    Try
      ' Validate
      If (tblCatTrack Is Nothing) Then
        ' Warn usr
        MsgBox("modAdapt/AddCatTrack: pronoun tracking has not been properly initialised...")
        Return False
      End If
      ' Check whether this combination is already in the table
      dtrFound = tblCatTrack.Select("Category='" & strPro.Replace("'", "''") & "' AND Type='" & strType & _
                                    "' AND Period='" & strPeriod & "'")
      ' Any results?
      If (dtrFound.Length = 0) Then
        ' Get the text of the line we're in
        strExample = GetOrgText(ndxThis, strLocation)
        ' Make this new row
        rowThis = tblCatTrack.NewRow
        With rowThis
          ' Assign values to the members of this new row
          .Item("Category") = strPro
          .Item("Type") = strType
          .Item("Period") = strPeriod
          .Item("Location") = strLocation
          .Item("Example") = strExample
        End With
        ' Add this new row to the table
        tblCatTrack.Rows.Add(rowThis)
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/AddCatTrack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CatTrackHtml
  ' Goal:   Get a report of all the CatTrack period items in HTML form
  ' History:
  ' 28-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function CatTrackHtml(ByVal strCategory As String) As String
    Dim dtrFound() As DataRow         ' Result of SELECT function
    Dim dtrCat() As DataRow = Nothing ' Result of SELECT on [tblCat]
    Dim tblCat As DataTable           ' The category datatable
    Dim strText As String             ' The text to be added
    Dim strType As String = ""        ' Class of pronoun description
    Dim strPeriod As String = ""      ' Period 
    Dim intI As Integer               ' Counter
    Dim intNum As Integer             ' Number of elements
    Dim intSection As Integer = 0     ' The section number <h2>
    Dim bInTable As Boolean = False   ' Whether we are inside a table

    Try
      ' Validate
      If (tblCatTrack Is Nothing) Then Return ""
      If (tblCatTrack.Rows.Count = 0) Then Return ""
      ' Set the category datatable
      Select Case strCategory
        Case CAT_NP
          tblCat = tdlSettings.Tables("Pronoun")
        Case Else
          tblCat = tdlSettings.Tables("Category")
      End Select
      ' Select the rows in the appropriate order
      dtrFound = tblCatTrack.Select("", "Type ASC, Period ASC, Category ASC")
      ' Start HTML
      strText = "<html>" & strHtmlHead & "<body><h1>" & vbCrLf & "Category report</h1><p>" & vbCrLf
      strText &= "(Last) file: " & IO.Path.GetFileNameWithoutExtension(strCurrentFile) & "<p>" & vbCrLf
      strText &= "Date: " & Format(Now, "f") & "<p>" & vbCrLf
      intNum = dtrFound.Length
      ' Walk all the elements
      For intI = 0 To intNum - 1
        ' Access this element
        With dtrFound(intI)
          ' Are we entering a new class?
          If (strType <> .Item("Type")) Then
            ' See if a previous period needs to be finished
            If (bInTable) Then
              ' Finish a previous period
              strText &= "</table></div><p>" & vbCrLf
              ' Show we are no longer in a table
              bInTable = False
            End If
            ' Note this new type
            strType = .Item("Type")
            ' Increment section number
            intSection += 1
            ' Start new class/type of pronouns using a <H2> level heading
            strText &= "<p><h2>" & intSection & ". " & strType & "</h2>" & "<br>" & vbCrLf
            ' TODO: add description of this class from [tblPro]
            If (Not tblCat Is Nothing) Then
              ' Look for a category definition with this name
              dtrCat = tblCat.Select("Name='" & strType & "'")
              ' Start a description in HTML if possible
              If (dtrCat.Length > 0) Then
                ' Get the description from the first element
                strText &= dtrCat(0).Item("Descr") & "<br>" & vbCrLf
                ' Are there any notes?
                If (dtrCat(0).Item("Notes") & "" <> "") Then
                  ' Include the notes
                  strText &= "<b>Note:</b><br>&nbsp'&nbsp;" & _
                    CStr(dtrCat(0).Item("Notes")).Replace(vbCrLf, "<br>") & "<br>" & vbCrLf
                End If
              End If
            End If
            ' Has a table been made?
            If (dtrCat.Length > 0) AndAlso (Not bInTable) Then
              ' Show the variants for the different periods
              strText &= "<table align='center'><tr><td>Period</td><td>Variants</td></tr>" & vbCrLf & _
                "<tr><td>Old English</td><td>" & VernToEnglish(dtrCat(0).Item("OE").ToString) & "</td></tr>" & vbCrLf & _
                "<tr><td>Middle English</td><td>" & VernToEnglish(dtrCat(0).Item("ME").ToString) & "</td></tr>" & vbCrLf & _
                "<tr><td>early Modern English</td><td>" & VernToEnglish(dtrCat(0).Item("eModE").ToString) & "</td></tr>" & vbCrLf & _
                "<tr><td>Modern British English</td><td>" & VernToEnglish(dtrCat(0).Item("MBE").ToString) & "</td></tr>" & vbCrLf & _
                "</table>" & vbCrLf
              ' Note new period, which belongs to a new type anyway
              strPeriod = .Item("Period")
              ' Start new period
              strText &= "<p>Period=" & strPeriod & "<br>" & vbCrLf
              ' Do introduce the variant example table
              strText &= "Examples of this type and period:<br>" & _
                "<div><table><tr><td><b>Category</b></td><td><b>Example</b></td></tr>" & vbCrLf
              ' Show we are inside a table
              bInTable = True
            End If
          End If
          If (strPeriod <> .Item("Period")) Then
            ' See if a previous period needs to be finished
            If (bInTable) Then
              ' Finish a previous period
              strText &= "</table></div>" & vbCrLf
              ' Show we are no longer in a table
              bInTable = False
            End If
            ' Note this new period
            strPeriod = .Item("Period")
            ' Start new period
            strText &= "<p>Period=" & strPeriod & "<br>" & vbCrLf
            ' Has a table been made?
            If (Not bInTable) Then
              ' Do introduce the variant example table
              strText &= "Examples of this type and period:<br>" & _
                "<div><table><tr><td><b>Category</b></td><td><b>Example</b></td></tr>" & vbCrLf
              ' Show we are inside a table
              bInTable = True
            End If
          End If
          ' Show this element in the variant example table
          strText &= "<tr><td>" & .Item("Category").ToString & "</td>" & _
            "<td>" & .Item("Example").ToString & " [" & .Item("Location").ToString & "]</td>" & _
            "</tr>" & vbCrLf
        End With
      Next intI
      ' See if a previous period needs to be finished
      If (bInTable) Then
        ' Finish a previous period
        strText &= "</table></div>" & vbCrLf
        ' Show we are no longer in a table
        bInTable = False
      End If
      ' Finish the HTML stuff
      strText &= "</body></html>" & vbCrLf
      ' Return the result
      Return strText
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/CatTrackHtml error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return nothing
      Return ""
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   VariantsHtml
  ' Goal:   Give the variants for the indicated period
  ' History:
  ' 31-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function VariantsHtml(ByRef dtrPro() As DataRow, ByVal strPeriod As String) As String
    Dim strText As String = ""  ' Variants description in HTML
    Dim bDoOE As Boolean = False    ' Include OE?
    Dim bDoME As Boolean = False    ' Include ME?
    Dim bDoModE As Boolean = False  ' Include eModE?
    Dim bAdded As Boolean = False   ' Did we add anything?

    Try
      ' Add the variants for this period
      If (dtrPro Is Nothing) Then
        ' There is no variant definition
        strText &= "(No variant definitions for this type/period.)<br>" & vbCrLf
      Else
        ' See if there are variant definitions for this period
        If (dtrPro(0).Item(strPeriod).ToString = "") Then
          ' There is no variant definition
          strText &= "(No variant definitions for this type/period.)<br>" & vbCrLf
        Else
          ' Include the variant definition
          strText &= "Variant definition:<br>&nbsp;&nbsp;" & dtrPro(0).Item(strPeriod) & _
            "<br>" & vbCrLf
          ' Should any previous variant definitions be included?
          Select Case strPeriod
            Case "OE"
              ' Don't include previous variants
            Case "ME"
              ' Include OE
              bDoOE = True
            Case "eModE"
              ' Include OE, ME
              bDoOE = True : bDoME = True
            Case "MBE"
              ' Include OE, ME, eModE
              bDoOE = True : bDoME = True : bDoModE = True
          End Select
          ' Include OE?
          If (bDoOE) AndAlso (dtrPro(0).Item("OE").ToString <> "") Then
            strText &= "&nbsp;&nbsp;OE: " & dtrPro(0).Item("OE") & "&nbsp;"
            bAdded = True
          End If
          ' Include ME?
          If (bDoME) AndAlso (dtrPro(0).Item("ME").ToString <> "") Then
            strText &= "&nbsp;&nbsp;ME: " & dtrPro(0).Item("ME") & "&nbsp;"
            bAdded = True
          End If
          ' Include eModE?
          If (bDoModE) AndAlso (dtrPro(0).Item("eModE").ToString <> "") Then
            strText &= "&nbsp;&nbsp;eModE: " & dtrPro(0).Item("eModE") & "&nbsp;"
            bAdded = True
          End If
          ' Need finishing touch?
          If (bAdded) Then
            strText &= "<br>" & vbCrLf
          End If
        End If
      End If
      ' Return the result
      Return strText
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/VariantsHtml error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return nothing
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitMissing
  ' Goal:   Initialise the missing periods collection
  ' History:
  ' 28-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub InitMissing()
    Try
      ' Initialise it
      colMissing.Clear()
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/InitMissing error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   HasMissing
  ' Goal:   Whether there are any missing pronouns
  ' History:
  ' 28-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function HasMissing() As Boolean
    Try
      Return (colMissing.Count > 0)
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/HasMissing error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddMissing
  ' Goal:   Add one item to the missing pronoun set
  ' History:
  ' 28-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub AddMissing(ByVal strPro As String, ByVal strType As String, ByVal strPeriod As String, _
                         ByRef ndxThis As XmlNode)
    Dim strText As String   ' The text to be added
    Dim strExample As String  ' An example
    Dim strLoc As String = "" ' Location of example

    Try
      ' Make this item
      strText = strPro & ";" & strType & ";" & strPeriod
      ' Get the text of the line we're in
      strExample = GetOrgText(ndxThis, strLoc)
      ' Try add this item
      colMissing.AddUnique(strText, strExample & " [" & strLoc & "]")
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/AddMissing error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   MissingHtml
  ' Goal:   Get a report of all the missing period items in HTML form
  ' History:
  ' 28-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MissingHtml(ByVal strCategory As String) As String
    Dim strText As String   ' The text to be added
    Dim arLine() As String  ' One line
    Dim intI As Integer     ' Counter
    Dim intJ As Integer     ' Second counter

    Try
      ' Start HTML
      strText = "<html>" & strHtmlHead & "<body>" & vbCrLf
      strText &= "Report of missing " & strCategory & "s<p>" & vbCrLf
      strText &= "(Last) file: " & IO.Path.GetFileNameWithoutExtension(strCurrentFile) & "<p>" & vbCrLf
      strText &= "Date: " & Format(Now, "f") & "<p>" & vbCrLf
      strText &= _
        "<div><table><tr><td>#</td><td>" & strCategory & _
        "</td><td>Type</td><td>Period</td><td>Example</td></tr>" & vbCrLf
      ' Visit all the elements
      For intI = 0 To colMissing.Count - 1
        ' Start this line in the table
        strText &= "<tr><td>" & intI + 1 & "</td>"
        ' Add this element
        arLine = Split(colMissing.Item(intI), ";")
        For intJ = 0 To UBound(arLine)
          ' Add this cell
          strText &= "<td>" & arLine(intJ) & "</td>"
        Next intJ
        ' Add the example
        strText &= "<td>" & colMissing.Exmp(intI) & "</td>"
        ' Finish this row
        strText &= "</tr>" & vbCrLf
      Next intI
      ' Finish html
      strText &= "</table></div></body></html>" & vbCrLf
      ' Return the result
      Return strText
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/AddMissing error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return nothing
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddPeriod
  ' Goal:   Add the combination of this period with the indicated file
  ' History:
  ' 14-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AddPeriod(ByVal strFile As String, ByVal strPeriod As String) As Boolean
    Dim strName As String     ' The name of this file
    Dim intPerId As Integer   ' PeriodId value
    Dim dtrFound() As DataRow ' Result from SELECT
    Dim dtrParent As DataRow  ' The parent 
    Dim dtrThis As DataRow    ' A new datarow

    Try
      ' Validate
      If (strFile = "") OrElse (strPeriod = "") Then Return False
      ' We really do need to get the FILEname, not just the name...
      If (IO.File.Exists(strFile)) Then
        ' Get the name of this file
        strName = LCase(IO.Path.GetFileNameWithoutExtension(strFile))
      Else
        ' Perhaps this is a name instead of a full file?
        If (InStr(strFile, "\") = 0) Then
          ' Yes, this is a name
          strName = strFile
          strFile = strCurrentFile
          ' Double check
          If (strFile = "") OrElse (Not IO.File.Exists(strFile)) Then Return False
        Else
          ' This should have been a file, so return failure
          Return False
        End If
      End If
      ' Read this file into a pdx structure
      Status("Add period: reading xml file " & strName & " to derive information...")
      ' Do we need to initialize period + periodfile tables?
      If (tblPeriod Is Nothing) OrElse (tblPeriodFile Is Nothing) Then
        ' Initialise them
        tblPeriod = tdlPeriods.Tables("Period")
        tblPeriodFile = tdlPeriods.Tables("PeriodFile")
      End If
      ' Now find the PeriodId for this File
      dtrFound = tblPeriod.Select("Name='" & strPeriod & "'")
      ' Validate result
      If (dtrFound.Length = 0) Then Return False
      dtrParent = dtrFound(0)
      ' Get the period Id
      intPerId = dtrParent.Item("PeriodId")
      ' Add the period file information
      dtrFound = tblPeriodFile.Select("Name='" & strName.Replace("'", "''") & "'")
      If (dtrFound.Length = 0) Then
        ' Now we know for sure we need to read the file...
        If (Not ReadXmlDoc(strFile, pdxFile)) Then Return False
        ' Make a new entry
        dtrThis = tblPeriodFile.NewRow
        With dtrThis
          .Item("Name") = strName
          .Item("Search") = pdxFile.SelectSingleNode("//forest[string-length(@TextId)>0]").Attributes("TextId").Value
          .Item("IPmat") = pdxFile.SelectNodes("//eTree[starts-with(@Label,'IP-MAT')]").Count
          .Item("IPsub") = pdxFile.SelectNodes("//eTree[starts-with(@Label,'IP-SUB')]").Count
          .Item("PeriodId") = intPerId
          ' Make sure we get a UNIQUE FileId number
          .Item("FileId") = GetUniqueId(tblPeriodFile, "FileId")
          '.Item("FileId") = tblPeriodFile.Rows.Count + 1
          ' Add the correct parent
          .SetParentRow(dtrParent)
        End With
        ' Add this row
        tblPeriodFile.Rows.Add(dtrThis)
        ' Make sure the resulting period file is saved
        tdlPeriods.WriteXml(strPeriodFile)
        Status("Saved new period information to " & strPeriodFile)
      End If
      ' REturn success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/AddPeriod error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AskUserPeriod
  ' Goal:   Ask the user to supply the period for file [strFile]
  ' History:
  ' 16-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AskUserPeriod(ByVal strFile As String, ByRef strPeriod As String) As Boolean
    Try
      ' Request the period for this file
      With frmPeriod
        ' Set the filename
        .FileName = strFile
        ' Ask for the period
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK, Windows.Forms.DialogResult.Yes
            ' We have the correct period name
            strPeriod = .Period
            ' Now process this File/Period combination
            If (Not AddPeriod(strFile, strPeriod)) Then Exit Function
            ' REturn success
            Return True
        End Select
      End With
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/AskUserPeriod error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   HasPeriod
  ' Goal:   Check if the indicated file has a period entry or not
  ' History:
  ' 28-05-2010  ERK Created
  ' 08-07-2013  ERK Deal with files having an apostrophe in the name
  ' ------------------------------------------------------------------------------------
  Public Function HasPeriod(ByVal strFile As String, ByRef strPeriod As String, _
                            Optional ByVal bGeneral As Boolean = True) As Boolean
    Dim strName As String     ' The name of this file
    Dim intPerId As Integer   ' PeriodId value
    Dim dtrFound() As DataRow ' Result from SELECT

    Try
      ' Validate
      If (strFile = "") Then Return False
      ' Get the name of this file
      If (InStr(strFile, "\") > 0) Then
        strName = LCase(IO.Path.GetFileNameWithoutExtension(strFile))
      Else
        strName = LCase(strFile)
      End If
      ' Do we need to initialize period + periodfile tables?
      If (tblPeriod Is Nothing) OrElse (tblPeriodFile Is Nothing) Then
        ' Validate
        If (tdlPeriods Is Nothing) Then
          Logging("HasPeriod: No PeriodFile has been loaded")
          Return False
        End If
        ' Initialise them
        tblPeriod = tdlPeriods.Tables("Period")
        tblPeriodFile = tdlPeriods.Tables("PeriodFile")
      End If
      ' Now find the PeriodId for this File
      dtrFound = tblPeriodFile.Select("Name='" & strName.Replace("'", "''") & "'")
      ' Found anything?
      If (dtrFound.Length > 0) Then
        ' Get the period id
        intPerId = dtrFound(0).Item("PeriodId")
        ' Now find the name of this period
        dtrFound = tblPeriod.Select("PeriodId=" & intPerId)
        ' Found anything?
        If (dtrFound.Length > 0) Then
          ' Evaluate the name of this period
          strPeriod = Trim(dtrFound(0).Item("Name"))
          ' Is it valid?
          If (strPeriod <> "") Then
            ' Need to make a general period?
            If (bGeneral) Then strPeriod = GetGeneralPeriod(strPeriod)
            ' Return success
            Return True
          End If
        End If
      End If
      ' Return Failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/HasPeriod error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetGeneralPeriod
  ' Goal:   Get the GENERAL period (OE, ME, eModE, MBE) from the specific period (M1, M3 etc)
  ' History:
  ' 16-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetGeneralPeriod(ByVal strPeriod As String) As String
    Try
      ' Is it valid?
      If (strPeriod <> "") Then
        ' Evaluate the period name
        Select Case Left(strPeriod, 1)
          Case "O"
            strPeriod = "OE"
          Case "M"
            strPeriod = "ME"
          Case "E"
            strPeriod = "eModE"
          Case "B"
            strPeriod = "MBE"
        End Select
      End If
      ' Return the adapted result
      Return strPeriod
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetGeneralPeriod error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetPeriod
  ' Goal:   Get the period for this file: OE, ME, eModE, MBE
  ' History:
  ' 28-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetPeriod(ByRef pdxThis As XmlDocument, Optional ByVal bNeeded As Boolean = True) As String
    Dim ndxThis As XmlNode        ' <forestGrp> node with the file name
    Dim strFile As String         ' Name of the file we're in
    Dim strPeriod As String = ""  ' Name of this period

    Try
      ' Validate
      If (pdxThis Is Nothing) OrElse (strPeriodFile = "") Then Return ""
      ' See if the period information (subtypeing) is in the header
      ndxThis = pdxThis.SelectSingleNode("./descendant::teiHeader/descendant::creation")
      If (ndxThis IsNot Nothing) Then
        ' Check the appropriate attribute
        If (ndxThis.Attributes("subtype") IsNot Nothing) Then
          strPeriod = ndxThis.Attributes("subtype").Value
          If (strPeriod <> "") Then Return strPeriod
        End If
      End If
      ' Get to the file name
      ndxThis = pdxThis.SelectSingleNode("./descendant::forestGrp")
      ' Validate
      If (Not ndxThis Is Nothing) Then
        ' Get the file name
        strFile = ndxThis.Attributes("File").Value
        ' Try get the period
        If (HasPeriod(strFile, strPeriod)) Then
          ' Return this period
          Return strPeriod
        Else
          ' Do we really need to know it at this point?
          If (bNeeded) Then
            ' We really do need to get the period, so ask the user...
            If (AskUserPeriod(strFile, strPeriod)) Then
              ' Return the generalisation for this period
              Return GetGeneralPeriod(strPeriod)
            End If
          Else
            ' We don't really need to know it, so make it "unknown"
            Return "unknown"
          End If
        End If
      End If
      ' Return failure
      Return ""
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetPeriod error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddAttribute
  ' Goal:   Add the attribute to the node, if the attribute does not exist there
  ' History:
  ' 26-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AddAttribute(ByRef ndxThis As XmlNode, ByVal strAttrName As String, _
                               ByVal strAttrValue As String) As Boolean
    Return AddAttribute(pdxCurrentFile, ndxThis, strAttrName, strAttrValue)
  End Function
  Public Function AddAttribute(ByRef pdxThis As XmlDocument, ByRef ndxThis As XmlNode, _
                               ByVal strAttrName As String, ByVal strAttrValue As String) As Boolean
    Dim atrThis As XmlAttribute ' New attribute

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Check if the attribute is there
      If (ndxThis.Attributes(strAttrName) Is Nothing) Then
        ' Create an attribute
        atrThis = pdxThis.CreateAttribute(strAttrName)
        ' Append this attribute
        ndxThis.Attributes.Append(atrThis)
      End If
      ' Set the value of this attribute
      ndxThis.Attributes(strAttrName).Value = strAttrValue
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/AddAttribute error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddFeature
  ' Goal:   Add the feature [strFname] with value [strFvalue] to the node
  '         Result:
  '           <fs type='[strFStype]'>
  '             <f name='[strFname] value='[strFvalue]' />
  '           </fs>
  ' History:
  ' 26-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AddFeature(ByRef pdxThis As XmlDocument, ByRef ndxThis As XmlNode, ByVal strFStype As String, _
                             ByVal strFname As String, ByVal strFvalue As String) As Boolean
    Dim ndxChild As XmlNode     ' Child node of type <fs>

    Try
      ' ============== DEBUG
      'If (ndxThis.Name = "eTree") AndAlso (ndxThis.Attributes("Id").Value = 85) AndAlso (strFStype = "dep") _
      '   AndAlso (strFname = "rel") Then
      '  Stop
      'End If
      ' =======================
      ' Validate
      If (pdxThis Is Nothing) Then Return False
      ' Add this inside <fs type='NP'> to <f name='NPtype' value='...'/>
      ndxChild = SetXmlNodeChild(pdxThis, ndxThis, "fs", "type;" & strFStype, "type", strFStype)
      ' Validate
      If (ndxChild Is Nothing) Then
        ' Show error
        Status("Could not add an appropriate <fs type='" & strFStype & "'> child")
        ' Return failure
        Return False
      End If
      ' Add or replace the <f> child
      If (SetXmlNodeChild(pdxThis, ndxChild, "f", "name;" & strFname, "value", strFvalue) Is Nothing) Then
        ' Show error
        Status("Could not add an appropriate <f name='" & strFname & "'> child")
        ' Return failure
        Return False
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/AddFeature error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddFeatureSet
  ' Goal:   Add a set of features
  ' History:
  ' 11-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AddFeatureSet(ByRef ndxThis As XmlNode, ByVal strSet As String, ByVal strFeat As String) As Boolean
    Dim arFeat() As String  ' Feature set broken up into array
    Dim arLine() As String  ' One feature name + value

    Try
      If (strFeat <> "") Then
        ' Convert the features into an array
        arFeat = Split(strFeat, ";")
        ' Expect the lemma on place 0
        If (arFeat.Length > 0) Then
          ' Do the features
          For intJ = 0 To arFeat.Length - 1
            ' Validate
            If (arFeat(intJ) <> "") Then
              ' Get feature name and value
              arLine = Split(arFeat(intJ), "=")
              If (arLine.Length = 2) Then
                ' Add feature name and value
                If (Not AddFeature(pdxCurrentFile, ndxThis, "M", arLine(0), arLine(1))) Then Status("AddFeatureSet: Problem adding feature") : Return False
              End If
            End If
          Next intJ
        End If  ' arFeat.Length > 0
      End If    ' strFeat <> ""
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/AddFeatureSet error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetCatValues
  ' Goal:   Get a list of all possible values
  ' History:
  ' 17-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetCatValues(ByRef dtrAll() As DataRow) As String
    Dim intI As Integer         ' Counter
    Dim intPos As Integer       ' Position within string
    Dim strValue As String = "" ' Where we store values
    Dim strOne As String        ' One result

    Try
      ' Validate
      If (dtrAll Is Nothing) Then Return ""
      ' Add all values
      For intI = 0 To dtrAll.Length - 1
        ' Get one result
        strOne = dtrAll(intI).Item("Name")
        ' Find the hyphen
        intPos = InStr(strOne, "-")
        If (intPos = 0) Then
          ' This value is empty
          strOne = ""
        Else
          ' Get this value
          strOne = Mid(strOne, intPos + 1)
        End If
        ' Add this value to the array
        AddSemiStack(strValue, strOne)
      Next intI
      ' Return the result
      Return strValue
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetCatValues error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetOneAdv
  ' Goal:   Determine the adverb type of the adverb given
  ' History:
  ' 17-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetOneAdv(ByVal strAdvWord As String, ByVal strPeriod As String, _
    ByRef ndxThis As XmlNode) As String
    Dim strAdvType As String          ' The result
    Dim strAdvValues As String        ' The adverb values currently recognized
    Dim strAdvWords As String         ' Adverbs
    Dim strCategory As String = "Adv" ' The category: adverb
    Dim dtrThis As DataRow = Nothing  ' New datarow
    Dim dtrFound() As DataRow         ' Result of SELECT

    Try
      ' Validate
      If (strAdvWord = "") Then Return "unknown"
      ' Set the Adverb table
      If (dtrAdv Is Nothing) Then
        dtrAdv = tdlSettings.Tables("Category").Select("Name LIKE 'Adv-*'")
      End If
      ' Try to determine the type of adverb
      strAdvType = GetOneCatType(dtrAdv, strAdvWord, strPeriod)
      ' Is result okay?
      If (strAdvType = "") Then
        ' Ask the user to supply the adverb type
        With frmCatType
          ' Set the category
          .Category = strCategory
          ' Get a list of all adverb types
          strAdvValues = GetCatValues(dtrAdv)
          .CatValue = strAdvValues
          ' Derive the clause from this constituent
          .Clause = GetIPfromConstituent(ndxThis)
          ' Derive the Syntax from this constituent
          .Syntax(ndxThis)
          ' Get all the ones that are like me
          .Alike = GetFeatLike(strCategory, strAdvWord)
          ' Try to get the adverb type
          Select Case .ShowDialog
            Case DialogResult.OK, DialogResult.Yes
              ' Process the result
              strAdvType = .CatValue
            Case DialogResult.Cancel, DialogResult.Abort, DialogResult.No
              ' Abort this operation
              bInterrupt = True
              Return ""
          End Select
        End With
        ' Did we get a value?
        If (strAdvType = "") Then
          ' Keep list of unknown matches
          AddMissing(strAdvWord, strCategory, strPeriod, ndxThis)
          ' Didn't get a match...
          Return "unknown"
        Else
          ' We received a value - check if this adverb type already exists
          dtrFound = tdlSettings.Tables("Category").Select("Name='" & strCategory & "-" & strAdvType & "'" & _
                                                           " AND Type='" & strAdvType & "'")
          ' Got any results?
          If (dtrFound.Length = 0) Then
            ' Add a row to the correct table
            dtrThis = AddOneDataRow(tdlSettings, "Category", "CatId", "CategoryList")
            ' Add values for this row
            With dtrThis
              .Item("Name") = strCategory & "-" & strAdvType
              .Item("Descr") = "<Add your description here>"
              .Item("Type") = strAdvType
              .Item("OE") = ""
              .Item("ME") = ""
              .Item("eModE") = ""
              .Item("MBE") = ""
              .Item("Notes") = "This category was added automatically on " & Format(Today, "g")
            End With
            ' Also adapt the datarow containing the adverbs
            dtrAdv = tdlSettings.Tables("Category").Select("Name LIKE 'Adv-*'")
          Else
            dtrThis = dtrFound(0)
          End If
          ' It should now be in [dtrThis] - Check if it is in the correct period
          strAdvWords = dtrThis.Item(strPeriod).ToString
          If (Not IsInSemiStack(strAdvWords, strAdvWord, "|")) Then
            ' Add it here!
            AddSemiStack(strAdvWords, strAdvWord, False, "|")
            dtrThis.Item(strPeriod) = strAdvWords
            ' Make sure changes are saved
            XmlSaveSettings(strSetFile)
          End If
          ' Also add the adverb type to the feature dictionary
          If (Not AddFeatDict(strCategory, strAdvType, strPeriod, strAdvWord)) Then Return False
        End If
      End If
      ' Is there any result?
      If (strAdvType <> "") Then
        ' Note this result
        If (Not AddCatTrack(strAdvWord, strCategory & "-" & strAdvType, strPeriod, ndxThis)) Then
          Return "unknown"
        End If
      End If
      ' Return the result - whatever it is
      Return strAdvType
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetOneAdv error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return "unknown"
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetProPGN
  ' Goal:   Determine the Person/Gender/Number of the input [strPro]
  '         First the possibilities for the indicated period are tried,
  '           then those for earlier periods
  ' History:
  ' 26-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetProPGN(ByVal strPro As String, ByVal bDemPro As Boolean, ByVal strPeriod As String, _
    ByRef ndxThis As XmlNode) As String
    Dim strPGN As String        ' The result
    Dim strType As String = ""  ' The type of pronoun

    Try
      ' Validate
      If (strPro = "") Then Return "unknown"
      ' Set the correct table
      If (tblPro Is Nothing) OrElse (dtrPerPro Is Nothing) OrElse (dtrDemPro Is Nothing) Then
        tblPro = tdlSettings.Tables("Pronoun")
        dtrPerPro = tblPro.Select("Name LIKE 'Per*'")
        dtrDemPro = tblPro.Select("Name LIKE 'Dem*'")
      End If
      ' Determine which to choose: Personal or Demonstrative pronoun list
      If (bDemPro) Then
        ' Okay, we can try it
        strPGN = GetOnePGN(dtrDemPro, strPro, strPeriod, strType)
      Else
        ' First try it as personal pronoun
        strPGN = GetOnePGN(dtrPerPro, strPro, strPeriod, strType)
        ' Any result?
        If (strPGN = "") Then
          ' Now try it as demonstrative
          strPGN = GetOnePGN(dtrDemPro, strPro, strPeriod, strType)
        End If
      End If
      ' Is result okay?
      If (strPGN = "") Then
        ' Keep list of unknown matches
        AddMissing(strPro, IIf(bDemPro, "Dem", "Per"), strPeriod, ndxThis)
        ' Didn't get a match...
        Return "unknown"
      Else
        ' Note this result
        If (Not AddCatTrack(strPro, strType, strPeriod, ndxThis)) Then
          Return "unknown"
        End If
        ' Return the result
        Return strPGN
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetProPGN error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return "unknown"
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetOnePGN
  ' Goal:   Determine the Person/Gender/Number of the input [strPro]
  '         First the possibilities for the indicated period are tried,
  '           then those for earlier periods
  ' History:
  ' 28-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetOnePGN(ByRef dtrThis() As DataRow, ByVal strPro As String, _
                             ByVal strPeriod As String, ByRef strType As String) As String
    Dim strVariants As String ' Variants for this entry
    Dim arVariant() As String ' Array of matches
    Dim intI As Integer       ' Counter
    Dim intJ As Integer       ' Counter
    Dim intK As Integer       ' Variant counter (in order of decreasing periods)
    Dim bPeriodOk As Boolean  ' Whether this period should be checked or not
    Dim strResult As String = ""  ' The resulting period

    Try
      ' ============= DEBUG ===============
      'If (LCase(strPro) = "+dam") Then Stop
      '===================================
      ' Initialise period OK flag
      bPeriodOk = False
      ' Visit all periods in increasing order (that is, starting from MBE > ME > OE)
      For intK = 0 To UBound(arPeriod)
        ' See if this period is ok
        If (arPeriod(intK) = strPeriod) Then bPeriodOk = True
        ' Only start processing if this period is okay
        If (bPeriodOk) Then
          ' Try find the correct entry
          For intI = 0 To dtrThis.Length - 1
            ' Get the variants for this entry and for this period
            strVariants = dtrThis(intI).Item(arPeriod(intK)) & ""
            If (strVariants <> "") Then
              ' Get an array of variants
              arVariant = Split(Trim(LCase(strVariants)), "|")
              ' Any results?
              If (arVariant.Length > 0) Then
                ' Walk through the variants
                For intJ = 0 To arVariant.Length - 1
                  ' There are results, so see if there is a match with the pattern stored in [arVariant]
                  If (strPro Like arVariant(intJ)) Then
                    ' Also return the TYPE of this pronoun
                    strType = dtrThis(intI).Item("Name").ToString
                    ' ============= DEBUG ===============
                    'If (LCase(strPro) = "+dam") Then Stop
                    '===================================
                    ' Okay, there is a match, so return the result
                    ' Add the result to the semistack (only unique elements)
                    AddSemiStack(strResult, dtrThis(intI).Item("PGN"), True)
                  End If
                Next intJ
              End If
            End If
          Next intI
        End If
        ' If results have been found in the "highest" period, they should suffice
        If (strResult <> "") Then
          ' Return these results
          Return strResult
        End If
      Next intK
      ' =============== DEBUG ===============
      'If (InStr(strResult, ";") > 0) Then Stop
      ' =====================================
      ' Return the semistack with results
      Return strResult
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetOnePGN error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetOneCatType
  ' Goal:   Determine the adverb type of the input [strPro]
  '         First the possibilities for the indicated period are tried,
  '           then those for earlier periods
  ' History:
  ' 17-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  'Private Function GetOneCatType(ByRef dtrThis() As DataRow, ByVal strAdvWord As String, _
  '                           ByVal strPeriod As String, ByRef strCatType As String) As String
  Private Function GetOneCatType(ByRef dtrThis() As DataRow, ByVal strAdvWord As String, _
                             ByVal strPeriod As String) As String
    Dim strVariants As String ' Variants for this entry
    Dim arVariant() As String ' Array of matches
    Dim intI As Integer       ' Counter
    Dim intJ As Integer       ' Counter
    Dim intK As Integer       ' Variant counter (in order of decreasing periods)
    Dim bPeriodOk As Boolean  ' Whether this period should be checked or not
    Dim strResult As String = ""  ' The resulting period

    Try
      ' Initialise period OK flag
      bPeriodOk = False
      ' Visit all periods in increasing order (that is, starting from MBE > ME > OE)
      For intK = 0 To UBound(arPeriod)
        ' See if this period is ok
        If (arPeriod(intK) = strPeriod) Then bPeriodOk = True
        ' Only start processing if this period is okay
        If (bPeriodOk) Then
          ' Try find the correct entry
          For intI = 0 To dtrThis.Length - 1
            ' Get the variants for this entry and for this period
            strVariants = dtrThis(intI).Item(arPeriod(intK)) & ""
            If (strVariants <> "") Then
              ' Get an array of variants
              arVariant = Split(Trim(LCase(strVariants)), "|")
              ' Any results?
              If (arVariant.Length > 0) Then
                ' Walk through the variants
                For intJ = 0 To arVariant.Length - 1
                  ' There are results, so see if there is a match with the pattern stored in [arVariant]
                  If (strAdvWord Like arVariant(intJ)) Then
                    '' Also return the TYPE of this category
                    'strCatType = dtrThis(intI).Item("Name").ToString
                    ' Okay, there is a match, so return the result
                    ' Add the result to the semistack (only unique elements)
                    AddSemiStack(strResult, dtrThis(intI).Item("Type"), True)
                  End If
                Next intJ
              End If
            End If
          Next intI
        End If
        ' If results have been found in the "highest" period, they should suffice
        If (strResult <> "") Then
          ' Return these results
          Return strResult
        End If
      Next intK
      ' Return the semistack with results
      Return strResult
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetOneCatType error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetAdvType
  ' Goal:   Determine the Adverb type of this node
  ' History:
  ' 17-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetAdvType(ByRef ndxThis As XmlNode, ByVal strPeriod As String, _
                              ByRef strAdvType As String) As Boolean
    Dim strLeaf As String       ' The text of the <eLeaf> for a category

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Validate type
      If (ndxThis.Attributes("Label") Is Nothing) Then Return False
      ' Set default values for [strAdvType]
      strAdvType = "unknown"
      ' Get the leaf text from myself - the actual adverb or focus particle
      strLeaf = LCase(LeafText(ndxThis))
      ' Check if this adverb/particle already occurs for this period in the database
      strAdvType = GetOneAdv(strLeaf, strPeriod, ndxThis)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetAdvType error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetNpType
  ' Goal:   Determine the NP type of this node
  ' History:
  ' 26-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetNpType(ByRef ndxThis As XmlNode, ByVal strPeriod As String, ByRef strNPtype As String, _
                             ByRef strPGN As String) As Boolean
    Dim ndxChild As XmlNode     ' First <eTree> child
    Dim intChildren As Integer  ' Number of children of type <eTree>
    Dim strChLabel As String    ' The label of the child
    Dim strHdLabel As String    ' The label of the last child
    Dim strLeaf As String       ' The text of the <eLeaf> for a Dem/Pro
    Dim bHasConj As Boolean     ' Whether the NP contains a conjunction
    Dim bHasName As Boolean     ' Whether the NP contains a name (NPR)

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Validate type
      If (ndxThis.Attributes("Label") Is Nothing) Then Return False
      ' Set default values for NPtype and PGN
      strNPtype = "unknown" : strPGN = "unknown"
      '' =============== DEBUGGING ===================
      'If (ndxThis.Attributes("Id").Value = 426) Then Stop
      '' =============================================
      ' Get the number of children
      intChildren = ndxThis.SelectNodes("./child::eTree").Count
      ' Validate
      If (intChildren = 0) Then
        ' There is no <eTree> child, but that is okay
        strChLabel = ndxThis.Attributes("Label").Value
        ' Get the leaf text from myself
        strLeaf = LCase(LeafText(ndxThis))
        bHasConj = False : bHasName = False
        strHdLabel = ""
      Else
        ' Get the first child, which should determine the NP type??
        ndxChild = ndxThis.SelectSingleNode("./child::eTree[1]")
        If (ndxChild.Attributes("Label") Is Nothing) Then Return False
        ' Action depends on the label of the first child
        strChLabel = ndxChild.Attributes("Label").Value
        ' Get the actual <eLeaf> (in lowercase) in order to determine the Person/Gender/Number
        strLeaf = LCase(LeafText(ndxChild))
        ' Find out if it has a name as child
        bHasName = (ndxThis.SelectNodes("./eTree[tb:Like(string(@Label),'NR*|NPR*')]", conTb).Count > 0)
        ' Find out what the head nominal element is (there is always at most 1 nominal element)
        ' In order of importance the head can be: nominal, adjectival, pronominal, WH-nominal
        strHdLabel = GetHdLabel(ndxThis)
      End If
      ' Delete a possible initial "$" from the leaf
      If (Left(strLeaf, 1) = "$") Then strLeaf = Trim(Mid(strLeaf, 2))
      ' Check for Question Constituent types
      If (ndxThis.Attributes("Label").Value Like "WNP*") Then
        ' Provide default values
        strNPtype = "Wh" : strPGN = "3"
        ' Need to process a question NP
        If (intChildren = 0) Then
          ' This is a zero question NP
          strNPtype = "ZeroWh"
        ElseIf (intChildren = 1) Then
          ' This could be a Wh demonstrative or Wh pronoun -- look a the child labels
          If (DoLike(strChLabel, "D^*|D")) Then
            ' THis is a Wh demonstrative
            strNPtype = "WhDem"
          ElseIf (strChLabel Like "WPRO*") Then
            ' THis is a WH pronoun
            strNPtype = "WhPro"
          End If
        End If
        ' Return success
        Return True
      End If
      ' Start checking for different NP types
      If (strChLabel Like "PRO*") AndAlso ((intChildren <= 1) OrElse ((intChildren > 1) AndAlso _
              (Not ndxThis.SelectSingleNode("./descendant::eLeaf[starts-with(@Text,'*ICH*-')]") Is Nothing))) Then
        '' ======= DEBUGGING =========
        'If (intChildren = 2) Then Stop
        '' ===========================
        ' This is a pronoun
        If (InStr(strChLabel, "$") > 0) Then
          strNPtype = "PossPro"
        Else
          strNPtype = "Pro"
        End If
        ' Get the PGN of the pronoun
        strPGN = GetProPGN(strLeaf, False, strPeriod, ndxThis)
      ElseIf (strChLabel Like "PRO*") Then
        ' This is another kind of Pronoun constituent, with more attached to it
        strNPtype = "ProNP"
        ' Get the PGN of the pronoun
        strPGN = GetProPGN(strLeaf, False, strPeriod, ndxThis)
      ElseIf (DoLike(strChLabel, "EX|EX-*")) Then
        ' This is an expletive
        strNPtype = "Expl"
        ' Expletives are always 3ns, I guess
        strPGN = "3ns"
      ElseIf (intChildren = 1) AndAlso DoLike(strChLabel, "N|N^*") Then
        ' This is a bare noun
        strNPtype = "Bare"
        strPGN = "3s"
      ElseIf (intChildren = 1) AndAlso DoLike(strChLabel, "NS|NS^*|*+NS|*+NS^*") Then
        ' This is a bare noun
        strNPtype = "Bare"
        strPGN = "3p"
      ElseIf (intChildren = 0) AndAlso (strLeaf = "*con*") Then
        ' This is a subject elided under conjunction
        strNPtype = "ZeroSbj"
        strPGN = "empty"
      ElseIf (intChildren = 0) AndAlso (Left(strLeaf, 1) = "*") Then
        ' This is an NP marked as a trace
        strNPtype = "Trace"
        ' If this is a WH trace, then the NP type should be "empty"
        ' NOTE: the [strleaf] contains lower case...
        If (strLeaf Like "[*]t[*]*") Then
          strPGN = "empty"
        Else
          strPGN = "unknown"
        End If
      ElseIf (intChildren > 1) AndAlso (strChLabel Like "PRO$*") Then
        ' This is an Anchored NP: a PRO$ determiner with a nominal head somewhere
        strNPtype = "AnchoredNP"
        ' Determine the p/g/n
        bHasConj = (ndxThis.SelectNodes("./eTree[starts-with(@Label,'CONJ')]").Count > 0)
        strPGN = IIf(bHasConj, "3p", "3" & NumberType(strHdLabel))
      ElseIf (DoLike(strChLabel, "PRO$*|NP-POS*")) Then
        ' If there are no children, this is a pronoun
        If (intChildren = 0) Then
          ' A possessive pronoun...
          strNPtype = "PossPro"
        ElseIf (strHdLabel Like "N*") Then
          ' This is a definite NP, which has a genitive/possessive determiner
          strNPtype = "DefNP"
        End If
        ' Look at number of children
        If (intChildren > 1) Then
          ' Determine the p/g/n
          bHasConj = (ndxThis.SelectNodes("./eTree[starts-with(@Label,'CONJ')]").Count > 0)
          strPGN = IIf(bHasConj, "3p", "3" & NumberType(strHdLabel))
        Else
          ' Get the PGN of the HEAD NOUN
          strPGN = GetProPGN(strLeaf, False, strPeriod, ndxThis)
        End If
      ElseIf (DoLike(strChLabel, "NPR$|NR$")) AndAlso (intChildren = 0) Then
        ' This is a possessive determiner name
        strNPtype = "Proper"
        ' Singular name
        strPGN = "3s"
      ElseIf (DoLike(strChLabel, "NPRS$|NRS$")) AndAlso (intChildren = 0) Then
        ' This is a possessive determiner name
        strNPtype = "Proper"
        ' Plural name
        strPGN = "3p"
      ElseIf (DoLike(strChLabel, "NUM*|SUCH*|ONE*|OTHER*")) Then
        ' This is a numeral like one, two, three etc.
        strNPtype = "IndefNP"
        ' If this is a number, then we have to look at the number itself
        If (strChLabel Like "NUM*") Then
          strPGN = "3" & IIf(strLeaf Like "on*", "s", "p")
        Else
          strPGN = "3" & NumberType(strHdLabel)
        End If
      ElseIf (strChLabel Like "N^*") Then
        ' This is an OE NP with a noun, and without demonstrative
        strNPtype = "IndefNP"
        strPGN = "3s"
      ElseIf (strChLabel Like "NS^*") Then
        ' This is an OE NP with a noun, and without demonstrative
        strNPtype = "IndefNP"
        strPGN = "3p"
      ElseIf (strChLabel Like "Q*") OrElse (strHdLabel Like "Q*") Then
        ' This NP starts with a quantifier
        strNPtype = "QuantNP"
        ' PGN is just third person something (it would be nice to detect plurality)
        strPGN = "3" & NumberType(strHdLabel)
      ElseIf (DoLike(strChLabel, "FW*")) Then
        ' THis is a full NP
        strNPtype = "FullNP" : strPGN = "3ns"
      ElseIf (DoLike(strChLabel, "CP-FRL*")) Then
        ' THis is a full NP /what I want to say/, /who you have met yesterday/ etc
        strNPtype = "FullNP"
        ' The PGN is 3s
        strPGN = "3"
      ElseIf (DoLike(strChLabel, "NP*|CONJ*") AndAlso bHasConj) Then
        ' This is a full NP
        strNPtype = "FullNP"
        ' It must be 3p
        strPGN = "3p"
      ElseIf (strChLabel Like "D*") Then
        ' This is a demonstrative or a determiner - see how many children there arae
        If (intChildren = 1) OrElse ((intChildren = 2) AndAlso _
              (Not ndxThis.SelectSingleNode("./descendant::eLeaf[starts-with(@Text,'*ICH*-')]") Is Nothing)) Then
          '' ======= DEBUGGING =========
          'If (intChildren = 2) Then Stop
          '' ===========================
          ' This is an indep. dem. pn
          strNPtype = "Dem"
          ' Get the PGN of the pronoun
          strPGN = GetProPGN(strLeaf, True, strPeriod, ndxThis)
        ElseIf (strChLabel Like "D^*") OrElse (strChLabel Like "DPRO^*") Then
          ' This is a demonstrative heading an NP in Old English
          strNPtype = "DemNP"
          ' Get the PGN of the pronoun
          strPGN = GetProPGN(strLeaf, True, strPeriod, ndxThis)
          ' DOuble check the number
          If (strPGN = "3") Then
            ' The demonstrative is only specified for 3rd person --> get the number from the head noun
            strPGN &= NumberType(strHdLabel)
          ElseIf (DoLike(strPGN, "3m|3f|3n")) Then
            ' The demonstrative is specified for person and gender --> do get the number from the head noun!
            strPGN &= NumberType(strHdLabel)
          End If
        Else
          ' Also find out if one of the children has a label starting with "CONJ"
          bHasConj = (ndxThis.SelectNodes("./eTree[starts-with(@Label,'CONJ')]").Count > 0)
          ' Action depends on the text value
          Select Case strLeaf
            Case "a", "A", "an", "An"
              ' Indefinite NP
              strNPtype = "IndefNP"
              ' Determine PGN
              strPGN = IIf(bHasConj, "3p", "3" & NumberType(strHdLabel))
            Case "the", "The", "+te", "+de"
              ' Definite NP
              strNPtype = "DefNP"
              ' Determine PGN
              strPGN = IIf(bHasConj, "3p", "3" & NumberType(strHdLabel))
            Case Else
              ' ========== DEBUG ==================
              ' If (Not ndxThis.SelectSingleNode("./child::eTree/child::eLeaf[@Text='+te']") Is Nothing) Then Stop
              ' ===================================
              ' This is a definite NP starting with a demonstrative
              strNPtype = "DemNP"
              ' PGN is determined on the basis of the demonstrative's PGN
              strPGN = GetProPGN(strLeaf, True, strPeriod, ndxThis)
              '' Double check plural/singular
              'If (bHasConj) AndAlso (strPGN = "3s") Then strPGN = "3p"
              ' But if the head noun determines otherwise, we should follow it, rather than the demonstrative
              If (strPGN = "3s") AndAlso (NumberType(strHdLabel) = "p") Then
                ' Take plural rather than singular
                strPGN = "3p"
              ElseIf (DoLike(strPGN, "3|3m|3f|3n")) Then
                ' Take number possibly from head noun
                strPGN &= NumberType(strHdLabel)
              End If
          End Select
        End If
      ElseIf (bHasName) Then
        ' THis is some kind of combination with a proper name
        strNPtype = "Proper"
        ' Also find out if one of the children has a label starting with "CONJ"
        bHasConj = (ndxThis.SelectNodes("./eTree[starts-with(@Label,'CONJ')]").Count > 0)
        ' Ascertain the plurality
        strPGN = IIf(bHasConj, "3p", "3" & NumberType(strHdLabel))
      ElseIf DoLike(strHdLabel, "N|N^*|NS|NS^*|*+NS|*+NS^*") Then
        ' This probably is a full NP 
        strNPtype = "FullNP"
        ' Check number of children
        If (intChildren = 2) Then
          ' Check the second child
          strChLabel = ndxThis.SelectNodes("./eTree")(1).Attributes("Label").Value
          If (strChLabel Like "PP*") Then
            ' We will still call this a bare NP
            strNPtype = "BareWithPP"
          End If
        End If
        strPGN = "3" & NumberType(strHdLabel)
      Else
        ' Return unknown
        strNPtype = "unknown"
        ' Ascertain the plurality
        strPGN = IIf(bHasConj, "3p", "3" & NumberType(strHdLabel))
      End If
      ' ========== DEBUGGING
      'If (strChLabel Like "[*]T[*]*") AndAlso (strPGN = "unknown") Then Stop
      If (strPGN = "2empty") Then Stop
      ' =============================
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetNpType error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetHdLabel
  ' Goal:   Get the label of the head element
  ' History:
  ' 18-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetHdLabel(ByRef ndxThis As XmlNode) As String
    Dim ndxNext As XmlNode      ' Next node
    Dim strLabel As String      ' The label
    Dim strHead As String = ""  ' The head we are going to return

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return ""
      ' ===== DEBUGGING ======
      'If (ndxThis.Attributes("Id").Value = 21194) Then Stop
      ' ======================
      ' Get the last child and work backwards
      ndxNext = ndxThis.LastChild
      ' Go into a loop
      While (Not ndxNext Is Nothing)
        ' Only consider <eTree> items
        If (ndxNext.Name = "eTree") Then
          ' See if this has the correct label
          strLabel = ndxNext.Attributes("Label").Value
          ' Consider this label
          If (DoLike(strLabel, strHdPossible)) Then
            ' This one fulfills the criteria
            strHead = strLabel
            ' Return it
            Return strHead
          End If
        End If
        ' Go to the next sibling
        ndxNext = ndxNext.PreviousSibling
      End While
      ' Return the resulting head
      Return strHead
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetHdLabel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   NumberType
  ' Goal:   Check whether this word type can be plural or not
  ' History:
  ' 18-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function NumberType(ByVal strLabel As String) As String
    Try
      ' There is a limited number of categories that can be plural...
      If (DoLike(strLabel, "NS*|NRS*|NPRS*|*+NS*")) Then
        ' This is plural
        Return "p"
      ElseIf (DoLike(strLabel, "NR*|NPR*|N*|*+N*")) Then
        ' This is singular
        Return "s"
      Else
        ' Impossible to determine the number
        Return ""
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/NumberType error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   IPnumberForest
  ' Goal:   Recursive function to walk parts of a <forest> and add IP numbers to
  '         all the <eTree> elements in here
  ' History:
  ' 17-12-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function IPnumberForest(ByRef ndxThis As XmlNode, ByRef intIPnumber As Integer, _
                                  ByRef bDoIncr As Boolean) As Boolean
    Dim ndxNext As XmlNode  ' The next <eTree> element with respect to me...
    Dim strLabel As String  ' Label of this <eTree>

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Action depends on the kind of node
      Select Case ndxThis.Name
        Case "forest"
          ' Process all the children of this <forest>
          ndxNext = ndxThis.FirstChild
          While (Not ndxNext Is Nothing)
            ' Process this child
            If (Not IPnumberForest(ndxNext, intIPnumber, bDoIncr)) Then Return False
            ' Go to next child
            ndxNext = ndxNext.NextSibling
          End While
        Case "eTree"
          ' If this still has the older IP/Number feature, then delete this
          If (Not ndxThis.SelectSingleNode("./child::fs[@name='IP']") Is Nothing) Then

          End If
          ' Check if this is the start of a new IP
          strLabel = ndxThis.Attributes("Label").Value
          If (DoLike(strLabel, strIPstackType)) Then
            ' Reset the incrementation necessity
            bDoIncr = False
            ' Increment the IP number
            intIPnumber += 1
            ' Store this number to the current node - as attribute
            If (Not AddAttribute(ndxThis, "IPnum", intIPnumber)) Then Return False
            ' AddFeature(pdxCurrentFile, ndxThis, "IP", "Number", intIPnumber)
            ' Process all the children of this IP
            ndxNext = ndxThis.FirstChild
            While (Not ndxNext Is Nothing)
              ' Process this child
              If (Not IPnumberForest(ndxNext, intIPnumber, bDoIncr)) Then Return False
              ' Go to next child
              ndxNext = ndxNext.NextSibling
            End While
            ' Signal that the IP number should be incremented at the next <eTree> element
            bDoIncr = True
          ElseIf (Not DoLike(strLabel, strSkipTypes)) Then
            ' Should we increment anyway? (But no need to do that for conjunctions!)
            If (bDoIncr) AndAlso (Not DoLike(strLabel, "CONJ*")) Then
              ' Yes, increment
              intIPnumber += 1
              ' Reset
              bDoIncr = False
            End If
            ' Store this number to the current node - as attribute
            ' AddFeature(pdxCurrentFile, ndxThis, "IP", "Number", intIPnumber)
            If (Not AddAttribute(ndxThis, "IPnum", intIPnumber)) Then Return False
            ' Process all the children of this <eTree> element
            ndxNext = ndxThis.FirstChild
            While (Not ndxNext Is Nothing)
              ' Process this child
              If (Not IPnumberForest(ndxNext, intIPnumber, bDoIncr)) Then Return False
              ' Go to next child
              ndxNext = ndxNext.NextSibling
            End While
          End If
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/IPnumberForest error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetIpNumber
  ' Goal:   Get the IPnum attribute, which is the number of the IP to which [ndxThis] belongs
  ' History:
  ' 17-12-2010  ERK Created
  ' 20-12-2010  ERK Changed from feature IP/Number to attribute IPnum
  ' ------------------------------------------------------------------------------------
  Public Function GetIpNumber(ByRef ndxThis As XmlNode, ByRef intIPnumber As Integer) As Boolean
    Dim ndxFor As XmlNode = Nothing ' Forest node
    Dim bDoIncr As Boolean          ' Flag
    Dim intPtc As Integer           ' Percentage
    Dim intNum As Integer = 1       ' IP number
    Dim intForest As Integer = 0    ' Running number for the forest we are in
    Dim intCount As Integer         ' Total number of forests

    Try
      ' Initialise IP number
      intIPnumber = 0
      ' Validate
      If (ndxThis Is Nothing) Then Logging("GetIpNumber: ndxThis is Nothing") : Return False
      If (ndxThis.Name <> "eTree") Then Logging("GetIpNumber: ndxThis is no <eTree>") : Return False
      ' Try to get the number of this IP
      If (ndxThis.Attributes("IPnum") Is Nothing) Then
        ' No such attribute yet!
        intIPnumber = -1
      Else
        ' Additional testing
        If (ndxThis.Attributes("IPnum") Is Nothing) Then Logging("GetIpNumber: node without [IPnum] attribute (1)") : Return False
        ' Get the number
        intIPnumber = ndxThis.Attributes("IPnum").Value
      End If
      ' Do we have a result?
      If (intIPnumber < 0) Then
        ' There is no result: IP numbers must be recalculated!
        Status("Recalculating IP-numbers...")
        Logging("GetIpNumber: Recalculating IP-numbers...")
        ' Go to the first forest
        If (Not GetFirstForest(pdxCurrentFile, ndxFor)) Then Logging("GetIpNumber: could not get first forest") : Return False
        ' Determine how many there are
        intCount = ndxFor.ParentNode.ChildNodes.Count
        If (intCount <= 0) Then
          ' Something is wrong...
          Logging("GetIpNumber: could not count the number of <forest> nodes (warning only)")
          ' Take an imaginary number
          intCount = 10000
        End If
        ' Initialise the number counting
        intNum = 1
        ' Walk all forests
        While (Not ndxFor Is Nothing)
          ' Show where we are
          intPtc = (intForest * 100) \ intCount
          Status("Adding IP numbers " & intPtc & "%", intPtc)
          ' Reset flag
          bDoIncr = False
          ' Process the IP-Number features for this part
          If (Not IPnumberForest(ndxFor, intNum, bDoIncr)) Then Logging("GetIpNumber: IPnumberForest did not succeed") : Return False
          ' Go to the next forest
          ndxFor = ndxFor.NextSibling
          intForest += 1
        End While
        ' Make sure Dirty is set, so that changes get saved
        frmMain.SetDirty(True)
        ' Show we added IP numbers
        Status("Adding IP numbers done")
        ' Still need to get the number of myself
        If (ndxThis.Attributes("IPnum") Is Nothing) Then Logging("GetIpNumber: node without [IPnum] attribute (2)") : Return False
        intIPnumber = ndxThis.Attributes("IPnum").Value
        ' One more error check
        If (intIPnumber < 0) Then Logging("GetIpNumber: intIPnumber is negative") : Return False
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetIpNumber error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetIpDist
  ' Goal:   Determine the distance in IP boundaries between the source [ndxThis] and
  '           its target node, if there is such a node
  ' History:
  ' 14-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetIpDist(ByRef ndxThis As XmlNode, ByRef intIpDist As Integer) As Boolean
    Dim ndxAnt As XmlNode     ' My antecedent
    Dim intIPsrcN As Integer  ' The number of the source IP
    Dim intIPdstN As Integer  ' The number of the antecedent IP

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Validate type
      If (ndxThis.Attributes("Label") Is Nothing) Then Return False
      ' Check whether this node is referential
      If (Not HasCoref(ndxThis)) Then
        ' This indicates that there is in fact no coreference relation
        intIpDist = -2
        Return True
      End If
      ' Get my own number
      If (Not GetIpNumber(ndxThis, intIPsrcN)) Then Return False
      ' Get the antecedent node
      ndxAnt = CorefDst(ndxThis)
      ' Do we have an antecedent?
      If (ndxAnt Is Nothing) Then
        ' No antecedent, so distance zero
        intIpDist = 0
      Else
        ' Get the antecedent's number
        If (Not GetIpNumber(ndxAnt, intIPdstN)) Then Return False
        ' Double check the result
        If (intIPsrcN < 0) OrElse (intIPdstN < 0) Then
          ' This means that the node does not belong to an IP
          intIpDist = 0
          Return True
        End If
        ' subtract the numbers
        intIpDist = intIPsrcN - intIPdstN
        ' Check the result and adapt to -1 for cataphores
        If (intIpDist < 0) Then intIpDist = -1
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetIpDist error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetCogState
  ' Goal:   Determine the Cognitive State of this node:
  '         inf  In Focus     - it
  '         act  Activated    - that, this, this N
  '         fam  Familiar     - that N
  '         unq  Uniquely Idt - the N
  '         ref  Referential  - indefinite this N
  '         typ  Type Idt     - a N
  '         non  (none)       - bare N that is "Inert"
  ' History:
  ' 07-11-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetCogState(ByRef ndxThis As XmlNode, ByRef strCgState As String) As Boolean
    Dim ndxParent As XmlNode    ' The parent <eTree> of me
    Dim ndxAnt As XmlNode       ' My antecedent
    Dim ndxWork As XmlNode      ' Working node
    Dim intIpDist As Integer    ' IP distance to antecedent
    Dim strLabel As String      ' My own label
    Dim strNPtype As String     ' Type of the current NP
    Dim strRefType As String    ' The referential state of myself

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Validate type
      If (ndxThis.Attributes("Label") Is Nothing) Then Return False
      ' Initialisation for when the current NP should not have a cognitive status at all
      strCgState = ""
      ' The NP type and RefType must be available, otherwise we cannot do anything
      strNPtype = GetFeature(ndxThis, "NP", "NPtype") : If (strNPtype = "") Then Return True
      strRefType = GetFeature(ndxThis, "coref", "RefType") : If (strRefType = "") Then Return True
      ' Get my IP parent, because we are only determine the status of IP children
      ndxParent = ndxThis.ParentNode
      If (ndxParent Is Nothing) Then Return True
      If (ndxParent.Name = "eTree") Then
        If (Not ndxParent.Attributes("Label").Value Like "IP*") Then Return True
      End If
      ' Set default Cognitive Status values for NPs that are IP children
      strCgState = "unknown"
      ' Get my label
      strLabel = ndxThis.Attributes("Label").Value
      ' Get my antecedent (if any)
      ndxAnt = CorefDst(ndxThis)
      ' Action partly depends on the NP type
      Select Case strNPtype
        Case "Bare"
          ' If this is a bare NP, then it may have no status if it is "Inert"
          strCgState = IIf(strRefType = "Inert", "non", "typ")
        Case "Trace"
          ' We really are not going to assign cognitive status to a trace
          strCgState = "non"
        Case "IndefNP", "QuantNP"
          ' The cognitive status is "type identifiable"
          strCgState = "typ"
        Case "Proper", "FullNP"
          ' Do we have an antecedent?
          If (strRefType = "Assumed") Then
            ' There is a non-textual antecedent
          ElseIf (ndxAnt Is Nothing) Then
            ' There is no antecedent
            strCgState = IIf(strNPtype = "Proper", "ref", "typ")
          Else
            ' There is a textual antecedent
          End If
        Case "Wh", "ZeroWh", "WhDem", "WhPro"
          strCgState = strCgState
        Case "DefNP", "AnchoredNP"
          strCgState = "unq"
        Case "PossPro"
          ' This should not be possible
          strCgState = "posspro"
        Case "DemNP"
          ' Do we have an antecedent?
          If (ndxAnt Is Nothing) Then
            ' There is no antecedent
            strCgState = "fam"    ' At least it must be familiar
          Else
            ' There is an antecedent: continue...
            Select Case strRefType
              Case "Assumed"
                ' There is no *textual* antecedent
                strCgState = "fam"  ' It is familiar
              Case "Inferred"
                ' It just infers, so starts something new, but is identifiable
                strCgState = "unq"
              Case Else
                ' How far away is the antecedent?
                intIpDist = GetFeature(ndxThis, "coref", "IPdist")
                If (intIpDist = 0) Then
                  ' Same line pronoun --> activated
                  strCgState = "act"
                ElseIf (intIpDist < 2) Then
                  ' Nearby -- check the grammatical role of the antecedent
                  Select Case GetFeature(ndxAnt, "NP", "GrRole")
                    Case "Subject"
                      ' Must be in focus
                      strCgState = "inf"
                    Case Else
                      ' A pro/dem with non-subject antecedent may only be "in focus" if there
                      '   is no other pro/dem in the current IP
                      ndxWork = ndxParent.SelectSingleNode("./child::eTree[not(@Id='" & ndxThis.Attributes("Id").Value & "') and child::fs/child::f[@name='NPtype' and (@value='Dem' or @value='Pro')]]")
                      If (ndxWork Is Nothing) Then
                        ' There is no other pro/dem in the current IP
                        strCgState = "inf"
                      Else
                        ' There is another pronoun or demonstrative in the current IP, and that may link back to the previous subject
                        strCgState = "act"
                      End If
                  End Select
                Else
                  ' Too far away! It is only activated
                  strCgState = "act"
                End If
            End Select
          End If
        Case "Pro", "Dem", "ProNP"
          ' Do we have an antecedent?
          If (ndxAnt Is Nothing) Then
            ' There is no antecedent
            strCgState = "fam"    ' At least it must be familiar
          Else
            ' There is an antecedent: continue...
            Select Case strRefType
              Case "Assumed"
                ' There is no *textual* antecedent
                strCgState = "fam"  ' It is familiar
              Case "Inferred"
                ' It just infers (how is that possible??)
                strCgState = "pro-inf"
              Case Else
                ' How far away is the antecedent?
                intIpDist = GetFeature(ndxThis, "coref", "IPdist")
                If (intIpDist = 0) Then
                  ' Same line pronoun --> activated
                  strCgState = "act"
                ElseIf (intIpDist < 2) Then
                  ' Nearby -- check the grammatical role of the antecedent
                  Select Case GetFeature(ndxAnt, "NP", "GrRole")
                    Case "Subject"
                      ' Must be in focus
                      strCgState = "inf"
                    Case Else
                      ' A pro/dem with non-subject antecedent may only be "in focus" if there
                      '   is no other pro/dem in the current IP
                      ndxWork = ndxParent.SelectSingleNode("./child::eTree[not(@Id='" & ndxThis.Attributes("Id").Value & "') and child::fs/child::f[@name='NPtype' and (@value='Dem' or @value='Pro')]]")
                      If (ndxWork Is Nothing) Then
                        ' There is no other pro/dem in the current IP
                        strCgState = "inf"
                      Else
                        ' There is another pronoun or demonstrative in the current IP, and that may link back to the previous subject
                        strCgState = "act"
                      End If
                  End Select
                Else
                  ' Too far away! It is only activated
                  strCgState = "act"
                End If
            End Select
          End If
        Case "Expl"
        Case "ZeroSbj"
          ' This must be in focus
          strCgState = "inf"
        Case "unknown"
          ' If the NP type is not known, then the CGtype can also not be known
          strCgState = "unknown"
        Case Else
          ' We are not able to handle this type of NP, which is weird
          strCgState = "NPtype=" & strNPtype
      End Select

      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetCogState error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetPlemma
  ' Goal:   Determine the Preposition Lemma of this node
  ' History:
  ' 17-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetPlemma(ByRef ndxThis As XmlNode, ByRef strPlemma As String) As Boolean
    Dim strVern As String       ' My own text

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Validate type
      If (ndxThis.Attributes("Label") Is Nothing) Then Return False
      ' Get my own vernacular text
      strVern = NodeText(ndxThis)
      ' See if the dictionary contains it
      strPlemma = GetFeatEntry(CAT_PNORM, strVern)
      ' Check the result
      If (strPlemma = "") Then
        ' Set default values for Preposition Lemma
        strPlemma = "unknown"
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetPlemma error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   GetVbType
  ' Goal:   Determine the Verb Type of this node
  '         Use the array in [arFeatDef], which contains all verbs that are to be considered as [Unaccusative]
  ' History:
  ' 15-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetVbType(ByRef ndxThis As XmlNode, ByRef strVbType As String, ByVal strCategory As String) As Boolean
    Dim strVern As String       ' My own text
    Dim strLemma As String      ' Lemma

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Validate type
      If (ndxThis.Attributes("Label") Is Nothing) Then Return False
      ' Set default values for Verb type
      strVbType = "unknown"
      ' Get my own vernacular text
      strVern = NodeText(ndxThis)
      ' Check if this one is in [arFeatDef]
      If (strVern = "") Then
        strVbType = "empty"
      Else
        Select Case strCategory
          Case CAT_VB
            ' See if the dictionary contains it
            strVbType = GetFeatEntry(CAT_VB, strVern)
          Case CAT_VBTYPE
            ' Default type: empty
            strVbType = ""
            ' Get the lemma of this node
            strLemma = GetFeature(ndxThis, "M", "l")
            If (strLemma <> "") Then
              ' ============= DEBUG ==========
              ' If (strLemma = "wenen" OrElse strLemma = "witen") Then Stop
              ' ==============================
              strVbType = GetFeatEntry(CAT_VBTYPE, strLemma, GetFeature(ndxThis, "M", "s"))
            End If
        End Select
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetVbType error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetGrRole
  ' Goal:   Determine the Grammatical Role of this node
  ' 
  '         Currently the following grammatical roles are being assigned:
  '           Subject   - Subject of a clause
  '           Argument  - Direct or indirect object, or other NP that is child of an IP
  '           NonArg    - Vocatives and others that should not be regarded as arguments
  '           PossDet   - Possessive determiner of another NP
  '           LeftDis   - Left dislocated NP
  '           PPobject  - Object governed by a preposition
  '           Oblique   - 
  '           None      - This NP must not be regarded as having any role
  '           unknown   - We are unable to determine the grammatical role of this NP
  ' History:
  ' 26-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetGrRole(ByRef ndxThis As XmlNode, ByRef strGrRole As String) As Boolean
    Dim ndxParent As XmlNode    ' The parent <eTree> of me
    Dim strParLabel As String   ' Label of my parent
    Dim strLabel As String      ' My own label

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Validate type
      If (ndxThis.Attributes("Label") Is Nothing) Then Return False
      ' Set default values for NPtype and PGN
      strGrRole = "unknown"
      ' Get my label
      strLabel = ndxThis.Attributes("Label").Value
      ' We might be able to act fast...
      If DoLike(strLabel, "NP*SBJ*|NP*NOM|NP-NOM-RSP|IP*SBJ*") Then
        ' Try to get the FIRST parent
        ndxParent = ndxThis.ParentNode
        If (ndxParent Is Nothing) OrElse (ndxParent.Name = "forest") Then
          ' We cannot be a subject
          strGrRole = "unknown"
        Else
          ' This only is a subject, if the parent is of an IP type
          If (DoLike(ndxParent.Attributes("Label").Value, "IP*")) Then
            ' This is a subject, so no problem!
            strGrRole = "Subject"
          Else
            ' This is NOT a subject, but it is a subject-like NP
            strGrRole = "unknown"
          End If
        End If
      ElseIf (DoLike(strLabel, "*-OB1|*-OB2")) Then
        ' Several NP types are already labelled as arguments
        strGrRole = "Argument"
      ElseIf (DoLike(strLabel, "*-VOC|*-TMP|*-ADV|*-MSR|*-PRN")) Then
        ' Several NP types should not be regarded as arguments: vocative, temporals, adverbials, measure nouns
        strGrRole = "NonArg"
      ElseIf DoLike(strLabel, "PRO$*|NPR$*|NP*GEN*|NP*POS") Then
        ' This is a possessor
        strGrRole = "PossDet"
      ElseIf (strLabel Like "NP*LFD*") Then
        ' This is a left dislocated NP, which could be anything, but is a separate category
        strGrRole = "LeftDis"
      ElseIf (strLabel Like "WNP*") Then
        ' This is a question NP, which could be subject, object -- depending on the NP in the cluase
        strGrRole = "Question"
      ElseIf (Not strLabel Like "NP*") Then
        ' This is not an NP, so we won't be able to deal with it
        strGrRole = "unknown"
      Else
        ' Go into a loop to determine my "proper" parent
        ndxParent = ndxThis
        Do
          ' Determine my parent
          ndxParent = ndxParent.ParentNode
          ' Validate
          If (ndxParent Is Nothing) Then Return False
          If (ndxParent.Name <> "eTree") Then
            ' If the parent is unknown, we will not be able to determine the grammatical role
            strGrRole = "unknown"
            Return True
          End If
          strParLabel = ndxParent.Attributes("Label").Value
        Loop While (DoLike(strParLabel, "CONJP*|FRAG*|NP*|QTP*")) AndAlso _
                   (Not DoLike(strParLabel, "*-GEN*|*-POS|*-OB1|*-OB2|*-SBJ|*-NOM*"))
        ' Action depends on the label of my parent
        If (strParLabel Like "PP*") Then
          ' We are the object of a PP
          strGrRole = "PPobject"
        ElseIf (strParLabel Like "REF*") Then
          ' This is a reference, which cannot count as something we can refer to
          strGrRole = "None"
        ElseIf (DoLike(strParLabel, "*-SBJ|*-NOM*")) Then
          ' We are a non-subject NP directly under an IP, or we are part of a larger OB1/OB2 NP
          strGrRole = "Subject"
        ElseIf (DoLike(strParLabel, "IP*|*-OB1|*-OB2")) Then
          ' We are a non-subject NP directly under an IP, or we are part of a larger OB1/OB2 NP
          strGrRole = "Argument"
        ElseIf (strParLabel Like "ADJP*") Then
          ' We are a modifier buried somewhere within an ADJP
          strGrRole = "Mod"
        ElseIf (DoLike(strParLabel, "NP*-GEN*|NP*-POS")) Then
          ' We are (inside) the determiner of another NP
          strGrRole = "PossDet"
        Else
          ' Right now I'm unable to determine what we are, so return oblique
          strGrRole = "Oblique"
        End If
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetGrRole error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   OneGrRoleFile
  ' Goal:   Calculate the grammatical roles for all the NPs in the intput [strInFile]
  '           and store the resulting tree in the [strOutFile]
  ' History:
  ' 26-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function OneGrRoleFile(ByVal strInFile As String, ByVal strOutFile As String) As Boolean
    Dim intI As Integer     ' Counter
    Dim intPtc As Integer   ' Percentage
    Dim strGrRole As String ' Grammatical role result
    Dim ndxThis As XmlNode  ' Current node

    Try
      ' Validate
      If (Not IO.File.Exists(strInFile)) Then
        Status("Could not find file " & strInFile)
        Return False
      End If
      ' Try read this file into an XML structure
      If (ReadXmlDoc(strInFile, pdxFile)) Then
        ' Select all the NP items of the XML structure
        ndxList = pdxFile.SelectNodes("//eTree[starts-with(@Label, 'NP') or @Label='PRO$']")
        ' Visit all these NP's
        For intI = 0 To ndxList.Count - 1
          ' Show where we are
          intPtc = (intI + 1) * 100 \ ndxList.Count
          Status("Processing file " & strInFile & " " & intPtc & "%", intPtc)
          ' Get the current node
          ndxThis = ndxList(intI)
          ' ============ DEBUGGING ==============
          ' If (ndxThis.Attributes("Label").Value = "NPR$") Then Stop
          ' =====================================
          ' Set the NPtype and the PGN to "unknown"
          strGrRole = "unknown"
          ' Get the NP type
          If (Not GetGrRole(ndxThis, strGrRole)) Then
            ' Show there is a problem
            Status("Unable to get the Grammatical Role from Id=" & ndxThis.Attributes("Id").Value)
            ' Return failure
            Return False
          End If
          ' If result is positive, then process it
          If (strGrRole <> "") Then
            ' See if we can process the NPtype information
            If (Not AddFeature(pdxFile, ndxThis, "NP", "GrRole", strGrRole)) Then Return False
          End If
        Next intI
        ' Write the result to the destination
        pdxFile.Save(strOutFile)
        ' Return success
        Return True
      Else
        ' Could not read the file, so return failure
        Status("Could not read the XML file: " & strInFile)
        Return False
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/OneGrRoleFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddPronounPGN
  ' Goal:   Add the pronoun/demonstrative [strVern] to the list of items having
  '           the indicated [strPGN]
  '         The [strNPtype] is either "Dem" or "Pro" for demonstrative or personal pronouns
  ' History:
  ' 16-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function AddPronounPGN(ByVal strNPtype As String, ByVal strPGN As String, ByVal strVern As String) As Boolean
    Dim strName As String   ' The name we will be looking for (partly)
    Dim dtrPro() As DataRow ' The correct element to treat

    Try
      ' Validate
      If (strNPtype = "") OrElse (strPGN = "") OrElse (strVern = "") Then Return False
      ' Sort out what we have
      Select Case strNPtype
        Case "Dem"
          strName = "Dem*"
        Case "Pro", "PossDet", "PossPro"
          strName = "Per*"
        Case Else
          ' This is an unknown NP type, so we won't be able to treat it
          Return False
      End Select
      ' Determine the correct row in the pronoun table
      dtrPro = tdlSettings.Tables("Pronoun").Select("Name LIKE '" & strName & "' AND PGN='" & strPGN & "'")
      If (dtrPro.Length > 0) Then
        ' Did we get more than one row?
        If (dtrPro.Length > 1) Then
          ' There is ambiguity in where to put it...
          Return False
        Else
          ' Validate
          If (strCurrentPeriod = "") OrElse (Not DoLike(strCurrentPeriod, "OE|ME|eModE|MBE")) Then
            ' The correct period indicator is missing...
            Return False
          End If
          ' We have exactly one row back --> process the result
          If (dtrPro(0).Item(strCurrentPeriod) & "" = "") Then
            dtrPro(0).Item(strCurrentPeriod) = strVern
          Else
            dtrPro(0).Item(strCurrentPeriod) &= "|" & strVern
          End If
          ' Make sure changes are saved
          XmlSaveSettings(strSetFile)
        End If
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/AddPronounPGN error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  Public Function MosActShow() As String
    Dim strBack As String = ""  ' What we return
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (tdlSettings.Tables("MosAct").Rows.Count = 0) Then Return ""
      ' Access it
      With tdlSettings.Tables("MosAct")
        ' Walk all values
        For intI = 0 To .Rows.Count - 1
          If (intI > 0) Then strBack &= ","
          strBack &= .Rows(intI).Item("Name").ToString
        Next intI
      End With
      ' Return the result
      Return strBack
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/MosActShow error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
End Module
