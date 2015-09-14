Imports System.Xml
Imports System.Text.RegularExpressions
Module modMorph
  ' ================================= LOCAL VARIABLES =========================================================
  Dim strMorphOElexicon As String = "D:\Data files\Corpora\OE-tagged\lexicon\lemmas"
  Private colVerbsBE As New StringColl      ' Collection of BE verb lemma's
  Private colVerbsAX As New StringColl      ' Collection of AX verb lemma's
  Private colVerbsVB As New StringColl      ' Collection of VB verb lemma's
  Private arPreVerb() As String = {"aefter", "nidher", "nither", "nydher", "nyther", _
                                   "geond", "dhurh", "thurh", "fordh", "forth", "under", _
                                   "eond", "with", "widh", "wyth", "wydh", "fore", "ofer", _
                                   "aet", "ymb", "upa", "upp", "oth", _
                                   "ac", "an", "em", "in", "of", "on", "to", "up", "ut", _
                                   "a"}
  ' i)	Vertaling: th=2, dh=3, ae=4, hw=5, sc=6, ea=7, eo=8, ie=9
  Private arSeqIn() As String = {"th", "dh", "ae", "hw", "sc", "ea", "eo", "ie"}
  Private arSeqOut() As String = {"2", "3", "4", "5", "6", "7", "8", "9"}
  Private lev_strSubstOrg As String = "abcdefghijklmnopqrstuvwxyz23456789"
  Private lev_strSubstCs1 As String = "e ktaw  y c nmu    do uci 327gx4ee"
  Private lev_strSubstCs2 As String = "op  ou ce q   ebk   w  se         "
  Private lev_strDelCosts As String = "4444441144444444444444444444444444"
  Private lev_strInsCosts As String = "4444441114444444444444441444444444"
  Private smp_intInf As Integer = 100 ' Infinite
  Private smp_strVowel As String = "aeio4789"
  Private smp_strAmbi As String = "yu"
  Private smp_strCons As String = "bcdfghjklmnpqrstvwxz2356"
  Private smp_colPath As New StringColl     ' Collection of EditOp path elements
  Private smp_intPath As Integer            ' Number of items in [smp_colPath]
  Private objSim As New Similarity          ' We need this in order to perform similarity measurements
  Private mrp_colSimi As New StringColl     ' Collection with similarity possibilities
  Private mrp_colHand As New StringColl     ' Those that need to be done by hand
  Private mrp_colMult As New StringColl     ' Those that have multiple values
  Private mrp_colFull As New StringColl     ' Those that have 100% reliability
  ' ================================ GLOBAL VARIABLES ==========================================================
  Public objM_Lem As DgvHandle              ' Morphology: lemma handler
  Public objM_VP As DgvHandle               ' Morphology: vernpos handler
  Public objM_Morph As DgvHandle            ' Morphology: morph handler
  Public objM_Rem As DgvHandle              ' Morphology: remain handler
  Public tdlOEdict As DataSet               ' Old English dictionary
  Public tdlLemmaVern As DataSet            ' Additions to the [Morph] table that need to be discussed
  Public tdlLemmaAsk As DataSet             ' Additions to the [Morph] table that are under question
  Public tdlBT As DataSet = Nothing         ' Hold B&T dataset
  ' ============================================================================================================
  ' ------------------------------------------------------------------------------------
  ' Name:   CheckMorphAdd
  ' Goal:   Add the combination of [strVern] and [strPos] with status [strStatus]
  ' History:
  ' 22-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function CheckMorphAdd(ByVal strVern As String, ByVal strPos As String, ByVal strLemma As String, ByVal strStatus As String, _
                                Optional ByVal strCorrect As String = "") As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim dtrNew As DataRow     ' New datarow

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Look for the entry
      dtrFound = tdlMorphDict.Tables("Check").Select("Vern='" & strVern.Replace("'", "''") & "' AND Pos='" & strPos.Replace("'", "''") & _
                                                     "' AND Lemma='" & strLemma.Replace("'", "''") & "'")
      If (dtrFound.Length > 0) Then
        ' It is already there -- get it
        dtrNew = dtrFound(0)
      Else
        ' It is not yet there, so create it
        dtrNew = AddOneDataRow(tdlMorphDict, "Check", "CheckId", "CheckList")
        With dtrNew
          .Item("Vern") = strVern : .Item("Pos") = strPos : .Item("Lemma") = strLemma
        End With
      End If
      ' Adapt the status
      dtrNew.Item("Status") = strStatus
      ' Need to change the correct status?
      dtrNew.Item("Correct") = strCorrect
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/CheckMorphAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CheckMorphStatus
  ' Goal:   If the combination of [strVern] and [strPos] and [strLemma] is available, give its status
  '         If no status is known, return empty ""
  ' History:
  ' 22-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function CheckMorphStatus(ByVal strVern As String, ByVal strPos As String, ByVal strLemma As String) As String
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim strStatus As String = ""  ' Status

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Look for the entry
      dtrFound = tdlMorphDict.Tables("Check").Select("Vern='" & strVern.Replace("'", "''") & "' AND Pos='" & strPos.Replace("'", "''") & _
                                                     "' AND Lemma='" & strLemma.Replace("'", "''") & "'")
      If (dtrFound.Length > 0) Then
        ' It is there -- get its status
        strStatus = dtrFound(0).Item("Status")
      End If
      ' Return status
      Return strStatus
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/CheckMorphAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphReadTaggedLexicon
  ' Goal:   Read the "lemmas" file and incorporate it into the MorphDict, if not present yet
  ' History:
  ' 22-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphReadTaggedLexicon() As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim dtrLemma() As DataRow ' Another one
    Dim dtrNew As DataRow     ' New entry
    Dim strText As String     ' Text of file
    Dim strDelim As String    ' Delimiter
    Dim strEntry As String    ' One entry
    Dim strLemma As String    ' One lemma
    Dim strWMpos As String    ' One w: entry
    Dim strYCOEpos As String  ' The POS within YCOE
    Dim arLine() As String    ' File in lines
    Dim intPos As Integer     ' Position in string
    Dim intI As Integer       ' Counter
    Dim intJ As Integer       ' Counter

    Try
      ' Validate
      If (tdlMorphTag Is Nothing) Then Logging("MorphTag dictionary is not loaded") : Return False
      If (tdlMorphDict Is Nothing) Then Logging("MorphDict is not loaded") : Return False
      If (tdlMorphDict.Tables("Lemma").Rows.Count > 1) Then
        ' Warn user
        Logging("MorphDict already contains entries")
        ' This is okay, but don't proceed
        Return True
      End If
      ' Try read file
      If (IO.File.Exists(strMorphOElexicon)) Then
        ' Read it
        strText = IO.File.ReadAllText(strMorphOElexicon)
        ' Get delimiter
        strDelim = GetDelim(strText, vbCrLf, vbCr, vbLf)
        arLine = Split(strText, strDelim)
        ' Walk through them
        For intI = 0 To arLine.Length - 1
          ' Get this entry
          strEntry = arLine(intI)
          ' Retrieve lemma
          intPos = InStr(strEntry, "_l:")
          If (intPos > 0) Then
            ' Get this entry
            strEntry = Mid(strEntry, intPos + 3)
            ' Get lemma
            intPos = InStr(strEntry, "_")
            If (intPos > 0) Then
              strLemma = Left(strEntry, intPos - 1)
              strEntry = Mid(strEntry, intPos)
              ' Repetitively get grammatical categories
              intPos = InStr(strEntry, "w:")
              While (intPos > 0)
                ' Eat to start
                strEntry = Mid(strEntry, intPos + 2)
                ' Find out where it ends
                intPos = InStr(strEntry, ")")
                strWMpos = Left(strEntry, intPos - 1)
                ' Find the POS that belongs to this [strWmPos]
                dtrFound = tdlMorphTag.Tables("Tag").Select("OEtag = '_w:" & strWMpos & "'")
                If (dtrFound.Length > 0) Then
                  For intJ = 0 To dtrFound.Length - 1
                    ' Get the YCOE tag
                    strYCOEpos = dtrFound(intJ).Item("POS")
                    ' Store this combination of @lemma, @wmpos
                    dtrLemma = tdlMorphDict.Tables("Lemma").Select("Vern = '" & strLemma.Replace("'", "''") & "'" & _
                                  " AND MWcat = '" & strWMpos.Replace("'", "''") & "'" & _
                                  " AND Pos = '" & strYCOEpos.Replace("'", "''") & "'")
                    If (dtrLemma.Length = 0) Then
                      ' Only create one if this is necessary
                      dtrNew = AddOneDataRow(tdlMorphDict, "Lemma", "LemmaId", "LemmaList")
                      With dtrNew
                        .Item("Vern") = strLemma : .Item("MWcat") = strWMpos : .Item("Pos") = strYCOEpos
                      End With
                    End If
                  Next intJ
                Else
                  ' This situation should not occur -- ask user
                  MsgBox("MorphReadTaggedLexicon: cannot find meaning of tag [" & strWMpos & "]")
                  Stop
                End If
                ' Try get next one
                intPos = InStr(intPos + 1, strEntry, "w:")
              End While
            End If
          End If
        Next intI
        ' Save the morphological dictionary
        Status("Saving morphological dictionary: " & strMorphDictFile)
        tdlMorphDict.WriteXml(strMorphDictFile)
        ' Return success
        Status("Okay")
        Return True
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphReadTaggedLexicon error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CheckMorphPOS
  ' Goal:   Look for the combination [strVern] and [strPos]:
  '         - if available              --> return the same [strPos]
  '         - if [strVern] is available --> return the [Pos] that goes with it
  '         - otherwise return empty ""
  ' History:
  ' 22-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function CheckMorphPOS(ByVal strVern As String, ByVal strPos As String, ByRef strMWcat As String) As String
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim strStatus As String = ""  ' Status
    Dim intI As Integer           ' Counter

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      '' Make sure we only get the first part of the POS
      'strPos = LabelmainOE(strPos, "-^")
      ' Get the correct vernacular form
      strVern = StringOEtoTagged(LCase(strVern)).Replace("$", "")
      ' Look for the entry
      dtrFound = tdlMorphDict.Tables("Lemma").Select("Vern='" & strVern.Replace("'", "''") & "'")
      ' Try to get the correct status
      For intI = 0 To dtrFound.Length - 1
        ' Keep track
        strMWcat = dtrFound(intI).Item("MWcat")
        ' Is this a hit?
        If (strPos Like dtrFound(intI).Item("Pos")) Then
          ' There is a hit
          strStatus = strPos
          Exit For
        Else
          ' There's a different pOS
          strStatus = dtrFound(intI).Item("Pos")
        End If
      Next intI
      ' Return status
      Return strStatus
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/CheckMorphPOS error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphLemmaAdd
  ' Goal:   Add the combination [strLemma], [strPos] to the dictionary
  ' History:
  ' 23-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphLemmaAdd(ByVal strLemma As String, ByVal strPos As String, ByVal strMWcat As String) As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim dtrNew As DataRow     ' New datarow

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Look for the entry
      dtrFound = tdlMorphDict.Tables("Lemma").Select("Vern='" & strLemma.Replace("'", "''") & "' AND Pos='" & strPos.Replace("'", "''") & _
                                                     "' AND MWcat='" & strMWcat.Replace("'", "''") & "'")
      If (dtrFound.Length > 0) Then
        ' It is already there -- get it
        dtrNew = dtrFound(0)
      Else
        ' It is not yet there, so create it
        dtrNew = AddOneDataRow(tdlMorphDict, "Lemma", "LemmaId", "LemmaList")
        With dtrNew
          .Item("Vern") = strLemma : .Item("Pos") = strPos : .Item("MWcat") = strMWcat
        End With
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphLemmaAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetMfeatures
  ' Goal:   Get all the "M" features for this node
  ' History:
  ' 23-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetMfeatures(ByRef ndxThis As XmlNode) As String
    Dim strFeat As String = ""  ' Where we return all the features
    Dim ndxWork As XmlNode      ' Working node

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return ""
      ' Find match
      ndxWork = ndxThis.SelectSingleNode("./child::fs[@type='M']")
      If (ndxWork Is Nothing) Then Return ""
      ' Walk all children
      ndxWork = ndxWork.FirstChild
      While (Not ndxWork Is Nothing)
        ' Skip lemma
        If (ndxWork.Attributes("name").Value <> "l") Then
          ' Add this feature
          AddSemiStack(strFeat, ndxWork.Attributes("name").Value & "=" & ndxWork.Attributes("value").Value)
        End If
        ' Go to next sibling
        ndxWork = ndxWork.NextSibling
      End While
      ' Return the result
      Return strFeat
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/GetMfeatures error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   VernInMorphDict
  ' Goal:   Look for all <Vern> entries in [tdlMorphDict]
  ' History:
  ' 04-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function VernInMorphDict(ByVal strVern As String, ByRef arMorphDict() As String) As Boolean
    Dim dtrFound() As DataRow ' Result of selection
    Dim intI As Integer       ' Counter

    Try
      ' Validate
      If (strVern = "") Then Return False
      ' Try find the vernacular in the morphological dictionary
      dtrFound = tdlMorphDict.Tables("Morph").Select("Vern='" & strVern.Replace("'", "''") & "'")
      ' Any results?
      If (dtrFound.Length = 0) Then Return False
      ' Make the array
      ReDim arMorphDict(0 To dtrFound.Length - 1)
      For intI = 0 To dtrFound.Length - 1
        With dtrFound(intI)
          ' Add the relevant features to the array
          arMorphDict(intI) = .Item("MorphId") & ";" & .Item("Pos") & ";" & .Item("Label") & ";" & .Item("l") & ";" & .Item("f")
        End With
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/VernInMorphDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphCheck
  ' Goal:   Check constituent [ndxThis] with the user
  '         This constituent has [strPOSmain] as main part of its @Label attribute
  '         We have previously decided that:
  '         - it has lemma [stLemma]
  '         But this lemma correlates with:
  '         - OE-tagged category [strMWcat]
  '         - YCOE POS label [strCheck]
  '         Let the user decide how to handle this discrepancy:
  '         1) 
  '         2) 
  '         3) 
  '         4) 
  ' History:
  ' 23-02-2013  ERK Created
  ' 04-03-2013  ERK Worked out implementation and switched to other form
  ' ------------------------------------------------------------------------------------
  Public Function MorphCheck(ByRef ndxThis As XmlNode, ByVal strLemma As String, ByVal strPOSmain As String, ByVal strCheck As String, _
                              ByVal strMWcat As String, ByVal strFeat As String) As String
    Dim arMorphDict(0) As String  ' Places in the morphological dictionary with <Lemma>
    Dim strVern As String         ' Vernacular

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return ""
      ' Use the form
      With frmMorphCheck
        ' Show the clause
        .Sentence = ClauseToHtml(ndxThis)
        .POS = strPOSmain
        .Lemma = strLemma
        .LemmaDictCat = strMWcat
        .LemmaDictPOS = strCheck
        .Features = strFeat.Replace(";", vbCrLf)
        ' Get and check vernacular
        strVern = ndxThis.SelectSingleNode("./child::eLeaf[1]").Attributes("Text").Value
        strVern = StringOEtoTagged(LCase(strVern).Replace("$", ""))
        If (Not VernInMorphDict(strVern, arMorphDict)) Then Stop
        .MorphDict = arMorphDict
        ' See what user decides
        Select Case .ShowDialog
          Case DialogResult.Cancel
            ' We need to cancel
            bInterrupt = True
            Return "Cancel"
        End Select
        ' Retrieve the user decision
        Return .Choice
      End With
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphCheck error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphDictAdd
  ' Goal:   Add the combination [strLemma], [strPos], [strParent] to the dictionary
  '         under the entry of [strVern]
  '         NB: only do this if it does not exist yet
  ' History:
  ' 23-02-2013  ERK Created
  ' 28-02-2014  ERK Additional validation
  ' ------------------------------------------------------------------------------------
  Public Function MorphDictAdd(ByVal strVern As String, ByVal strPos As String, ByVal strParent As String, _
                                ByVal strLemma As String, ByVal strFeat As String, _
                                Optional ByVal bUseFeat As Boolean = False, Optional ByVal intEntryId As Integer = -1, _
                                Optional ByVal bIsLemma As Boolean = False) As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim dtrNew As DataRow     ' New datarow
    Dim strSelect As String   ' Select statement

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Look for the entry
      strSelect = "Vern='" & strLemma.Replace("'", "''") & "' AND Pos='" & strPos.Replace("'", "''") & "'"
      ' Optionally include @Label (which should point to the parent label)
      If (strParent <> "") AndAlso (strParent <> strPos) Then
        strSelect &= " AND Label='" & strParent.Replace("'", "''") & "'"
      End If
      ' Optionally include @f (if it is not empty)
      If (bUseFeat) AndAlso (strFeat <> "") Then
        strSelect &= " AND f='" & strFeat & "'"
      End If
      ' Do the finding
      dtrFound = tdlMorphDict.Tables("Morph").Select(strSelect)
      ' Act on the results
      If (dtrFound.Length > 0) Then
        ' It is already there -- get it
        dtrNew = dtrFound(0)
        ' Possibly adapt lemma and features
        If (strLemma <> "") Then dtrNew.Item("l") = strLemma
        If (strFeat <> "") Then dtrNew.Item("f") = strFeat
        If (intEntryId >= 0) Then dtrNew.Item("EntryId") = intEntryId
        dtrNew.Item("t") = IIf(bIsLemma, "l", "d")
      Else
        ' It is not yet there, so create it
        dtrNew = AddOneDataRow(tdlMorphDict, "Morph", "MorphId", "MorphList")
        With dtrNew
          .Item("Vern") = strVern : .Item("l") = strLemma : .Item("Pos") = strPos : .Item("Label") = strParent : .Item("f") = strFeat
          If (intEntryId >= 0) Then
            .Item("EntryId") = intEntryId
          End If
        End With
        Logging("MorphDictAdd [" & strVern & "]: l=[" & strLemma & "], POS=" & strPos, False)
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphDictAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphVernPosParentAsk
  ' Goal:   Ask user if the combination Vern/Pos/Parent + lemma/feat okay is
  ' History:
  ' 05-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphVernPosParentAsk(ByVal strVern As String, ByVal strPos As String, ByVal strParent As String, _
                                        ByRef dtrFound() As DataRow, ByRef ndxThis As XmlNode, ByRef strFeat As String, _
                                        ByRef strGener As String) As Integer
    Dim intChoice As Integer = -1 ' User made choice

    Try
      ' Validate
      If (strVern = "") Or (strPos = "") Then Return False
      If (ndxThis Is Nothing) Then Return ""
      ' Ask the question
      With frmMorph
        '.POS = strPos
        '.Features = strFeat.Replace(";", vbCrLf)
        .OwnLabel = strPos
        .ParentLabel = strParent
        ' Show the clause
        .Sentence = ClauseToHtml(ndxThis)
        ' Pass on array of options
        .SetChoices(dtrFound)
        .Information = "Parent label:" & vbCrLf & strParent & vbCrLf & vbCrLf & _
          "Choose list item with correct lemma + features!"
        strGener = ""
        ' Ask user
        Select Case .ShowDialog
          Case DialogResult.OK
            ' Select this one
            intChoice = .Choice
            strFeat = .Features
          Case DialogResult.No
            ' Give it a default value
            intChoice = 0
            strFeat = .Features
          Case DialogResult.None
            ' None of them is valid
            intChoice = -1
          Case DialogResult.Cancel
            ' Quit
            intChoice = -1
            bInterrupt = True
          Case DialogResult.Ignore  ' Full generalization
            ' Signal that all is well, but we want full generalization
            intChoice = .Choice
            strGener = "full"
          Case DialogResult.Retry   ' Restricted generalization
            ' Signal that all is well, but we want restricted generalization
            intChoice = .Choice
            strGener = "restr"
        End Select
      End With
      ' Return user decision
      Return intChoice
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphVernPosParentAsk error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphDictGenAdd
  ' Goal:   -
  ' History:
  ' 05-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphDictGenAdd(ByVal strPos As String, ByVal strPrntLabel As String, ByVal strType As String) As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim dtrNew As DataRow     ' New datarow

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Look for the entry
      dtrFound = tdlMorphDict.Tables("Gen").Select("Pos='" & strPos & "' AND PrntLabel='" & strPrntLabel & _
                                                     "' AND Type='" & strType & "'")
      If (dtrFound.Length > 0) Then
        ' It is already there -- get it
        dtrNew = dtrFound(0)
      Else
        ' It is not yet there, so create it
        dtrNew = AddOneDataRow(tdlMorphDict, "Gen", "GenId", "GenList")
        With dtrNew
          .Item("Pos") = strPos : .Item("PrntLabel") = strPrntLabel : .Item("Type") = strType
        End With
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphDictGenAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphDictGenType
  ' Goal:   -
  ' History:
  ' 05-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphDictGenType(ByVal strPos As String, ByVal strPrntLabel As String) As String
    Dim dtrFound() As DataRow ' Result of SELECT

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Look for the entry
      dtrFound = tdlMorphDict.Tables("Gen").Select("Pos='" & strPos & "' AND PrntLabel='" & strPrntLabel & "'")
      If (dtrFound.Length = 0) Then
        ' It is not there, so return empty
        Return ""
      Else
        ' It is there, so return the type
        Return dtrFound(0).Item("Type")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphDictGenType error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   VernPosAdd
  ' Goal:   Add an item to the <VernPos> table
  ' History:
  ' 07-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function VernPosAdd(ByVal strPos As String, ByVal strVern As String, ByVal strType As String, _
                             ByVal strLemma As String, ByVal strLev As String, ByVal strFeat As String, _
                             Optional ByVal strSrc As String = "") As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim dtrNew As DataRow     ' New datarow

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Look for the entry
      'dtrFound = tdlMorphDict.Tables("VernPos").Select("Pos='" & strPos & "' AND Vern='" & strVern.Replace("'", "''") & _
      '                                               "' AND Type='" & strType & "'")
      dtrFound = tdlMorphDict.Tables("VernPos").Select("Pos='" & strPos & "' AND Vern='" & strVern.Replace("'", "''") & "'")
      If (dtrFound.Length > 0) Then
        ' It is already there -- get it
        dtrNew = dtrFound(0)
        ' Double check on [Type] 
        If (dtrNew.Item("Type").ToString <> strType) Then
          dtrNew.Item("Type") = strType
        End If
        ' Double check on [Features]
        If (dtrNew.Item("f").ToString = "") AndAlso (strType <> "") Then
          dtrNew.Item("f") = strFeat
        End If
      Else
        ' It is not yet there, so create it
        dtrNew = AddOneDataRow(tdlMorphDict, "VernPos", "VernPosId", "VernPosList")
        With dtrNew
          .Item("Pos") = strPos : .Item("Vern") = strVern : .Item("Type") = strType
          .Item("l") = strLemma : .Item("f") = strFeat : .Item("Lev") = strLev : .Item("Src") = strSrc
        End With
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/VernPosAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   VernPosType
  ' Goal:   Get a type from the <VernPos> table
  ' History:
  ' 07-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function VernPosType(ByVal strPos As String, ByVal strVern As String, ByRef strLemma As String, _
                              ByRef strFeat As String) As String
    Dim dtrFound() As DataRow ' Result of SELECT

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Look for the entry
      dtrFound = tdlMorphDict.Tables("VernPos").Select("Pos='" & strPos & "' AND Vern='" & strVern.Replace("'", "''") & "'")
      If (dtrFound.Length > 0) Then
        ' It is there - return the type, the lemma nd the feature
        strLemma = dtrFound(0).Item("l")
        strFeat = dtrFound(0).Item("f")
        Return dtrFound(0).Item("Type")
      End If
      ' Return failure
      Return "none"
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/VernPosType error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphDictAdapt
  ' Goal:   Ask user if the combination Vern/Pos/Parent + lemma/feat okay is
  ' History:
  ' 07-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphDictAdapt() As Boolean
    Dim dtrFound() As DataRow   ' Ordered view of <Morph>
    Dim dtrItem() As DataRow    ' Part of <Morph>
    Dim strFeat As String = ""  ' Value of @f
    Dim strH As String = ""     ' Value of @h
    Dim strVern As String = ""  ' Vernacular
    Dim strPos As String = ""   ' POS
    Dim strLem As String = ""   ' Lemma
    Dim strLev As String = "1"  ' Level
    Dim bFeat As Boolean        ' Are features same?
    Dim bLem As Boolean         ' Are lemma's same?
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter
    Dim intLen As Integer       ' Length of table
    Dim intPtc As Integer       ' Position
    Dim lstThis As List(Of String)

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      tdlMorphDict.AcceptChanges()
      ' Clear previous table
      ClearTable(tdlMorphDict.Tables("VernPos"))
      tdlMorphDict.AcceptChanges()
      Application.DoEvents()
      ' Access <Morph>
      With tdlMorphDict.Tables("Morph")
        ' Walk the whole table
        intLen = .Rows.Count
        For intI = 0 To intLen - 1
          ' Show where we are 
          intPtc = (intI + 1) * 100 \ intLen
          Status("MorphDictAdapt 1 " & intPtc & "%", intPtc)
          ' Access this element
          With .Rows(intI)
            ' Get the @F value
            strFeat = .Item("f").ToString : strH = "" : strLem = .Item("l").ToString
            ' Turn it into a list
            lstThis = Split(strFeat, ";").OrderBy(Function(x) x).ToList()
            '' Make sure the features are ORDERED alphabetically
            'lstThis = lstThis.OrderBy(Function(x) x).ToList
            ' find h= member
            For intJ = lstThis.Count - 1 To 0 Step -1
              ' Is this the one
              If (lstThis.Item(intJ) Like "h=*") Then
                ' Get it
                strH = lstThis.Item(intJ)
                ' Delete it
                lstThis.Remove(strH)
                ' Exit For
              ElseIf (lstThis.Item(intJ) Like "lemma=*") Then
                ' Is the actual lemme empty?
                If (strLem = "") Then
                  strLem = lstThis.Item(intJ)
                  lstThis.Remove(strLem)
                  .Item("l") = Mid(strLem, InStr(strLem, "=") + 1)
                End If
              End If
            Next intJ
            ' Join them together
            ' Get rebuilt @f
            strFeat = Join(lstThis.ToArray, ";")
            ' Adapt the features in the tdl
            .Item("f") = strFeat
            ' Possibly add H feature
            If (strH <> "") Then .Item("h") = strH
          End With
        Next intI
      End With
      ' Order the table <Morph>
      dtrFound = tdlMorphDict.Tables("Morph").Select("", "Vern ASC, Pos ASC")
      For intI = 0 To dtrFound.Length - 1
        ' Show where we are 
        intPtc = (intI + 1) * 100 \ dtrFound.Length
        Status("MorphDictAdapt 2 " & intPtc & "%", intPtc)
        ' Consider this combination
        strVern = dtrFound(intI).Item("Vern").ToString : strPos = dtrFound(intI).Item("Pos").ToString
        dtrItem = tdlMorphDict.Tables("Morph").Select("Vern='" & strVern & "' AND Pos='" & strPos & "'", _
                                                      "l ASC, f ASC")
        ' The first feature and lemma are always needed
        strFeat = dtrItem(0).Item("f").ToString : strLem = dtrItem(0).Item("l").ToString
        intLen = dtrItem.Length
        If (intLen = 1) Then
          ' Check frequency
          If (dtrFound(intI).Item("Freq") > 1) Then
            ' It occurs more, so list it
            If (Not VernPosAdd(strPos, strVern, "LemmaFeat", strLem, strLev, strFeat)) Then Return False
          End If
        Else
          ' Length must be larger than 1 --> are they all the same?
          bFeat = True : bLem = True
          For intJ = 1 To dtrItem.Length - 1
            ' Check
            If (strFeat <> dtrItem(intJ).Item("f").ToString) Then bFeat = False
            If (strLem <> dtrItem(intJ).Item("l").ToString) Then bLem = False : Exit For
          Next intJ
          ' Any good?
          If (bLem) Then
            ' Yes, lemma's are same!
            If (bFeat) Then
              ' The combinationof lemma + feature may be propagated
              If (Not VernPosAdd(strPos, strVern, "LemmaFeat", strLem, strLev, strFeat)) Then Return False
            Else
              ' Only lemma may be propagated
              If (Not VernPosAdd(strPos, strVern, "LemmaOnly", strLem, strLev, strFeat)) Then Return False
            End If
          End If
        End If
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphDictAdapt error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphRemainAdd
  ' Goal:   Add the combination Vern/Pos/ParentLabel to the dictionary
  ' History:
  ' 11-03-2013  ERK Created
  ' 28-02-2014  ERK Do NOT take the parent label into account anymore...
  ' ------------------------------------------------------------------------------------
  Public Function MorphRemainAdd(ByVal strVern As String, ByVal strClauseH As String, ByVal strPos As String, ByVal strPrntLabel As String, _
                                 ByVal strFile As String, ByVal intForestId As Integer, ByVal intEtreeId As Integer) As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim dtrNew As DataRow     ' New datarow
    Dim strSelect As String = ""
    Dim bUseParent As Boolean = False

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Look for the entry
      strSelect = "Vern='" & strVern.Replace("'", "''") & "' AND Pos='" & strPos.Replace("'", "''") & "'"
      If (bUseParent) Then
        strSelect &= " AND PrntLabel='" & strPrntLabel.Replace("'", "''") & "'"
      End If
      dtrFound = tdlRemain.Tables("Remain").Select(strSelect)
      If (dtrFound.Length > 0) Then
        ' It is already there -- get it
        dtrNew = dtrFound(0)
        ' Adapt the frequency
        dtrNew.Item("Freq") += 1
      Else
        ' It is not yet there, so create it
        dtrNew = AddOneDataRow(tdlRemain, "Remain", "RemainId", "RemainList")
        With dtrNew
          .Item("Vern") = strVern : .Item("Pos") = strPos : .Item("PrntLabel") = strPrntLabel : .Item("Clause") = strClauseH
          .Item("Freq") = 1 : .Item("File") = strFile : .Item("forestId") = intForestId : .Item("EtreeId") = intEtreeId
        End With
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphRemainAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphSuffixRules
  ' Goal:   Derive suffix-rewrite rules in table [Morph]
  '         Look for candidates in table [Remain], against lemma's in [Morph] and [Lemma]
  '         Add candidate Vern/Pos/Lemma combinations in [VernPos]
  ' History:
  ' 16-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphSuffixRules() As Boolean
    Dim dtrFound() As DataRow ' Table of remainder
    Dim strVern As String     ' Current vern
    Dim strPos As String      ' Current POS
    Dim strLemma As String    ' Lemma
    Dim strSfx As String      ' Suffix
    Dim strRewrite As String  ' How the suffix must be re-written to get the lemma
    Dim intI As Integer       ' Counter
    Dim intPtc As Integer     ' Percentage
    Dim intNum As Integer = 0 ' Number of changes

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Get to the [Morph] table, and look for suffix-rewrite rules
      dtrFound = tdlMorphDict.Tables("Morph").Select("", "Pos ASC, Vern ASC, Freq DESC")
      ' Walk them all
      For intI = 0 To dtrFound.Length - 1
        If (bInterrupt) Then Return False
        ' Where are we?
        intPtc = (100 * (intI + 1) \ dtrFound.Length)
        Status("Morph suffix detection " & intPtc & "%", intPtc)
        ' Get the vern and pos and the lemma
        With dtrFound(intI)
          strVern = .Item("Vern") : strPos = .Item("Pos") : strLemma = .Item("l")
        End With
        strSfx = "" : strRewrite = ""
        ' Get a suffix-rewrite rule
        If (GetSuffixRewrite(strVern, strLemma, strSfx, strRewrite)) Then
          ' Add this rule to the Rule table
          If (Not MorphAddSfxRule(strSfx, strRewrite, strPos, strVern)) Then Logging("modMorph/MorphSuffixRules error #1") : Return False
          ' We found one!
          intNum += 1
        End If
      Next intI
      ' Show what we have done
      Logging("MorphSuffixRules found [" & intNum & "] matches")
      ' Save file if there are more matches
      If (intNum > 0) Then
        Status("Saving rules...")
        tdlMorphDict.WriteXml(strMorphDictFile)
        Status("Creating suffix report...")
        If (Not MorphSuffixReport()) Then Return False
        Status("Ready")
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphSuffixRules error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphAddSfxRule
  ' Goal:   Add a rule from [strSfx] > [strRewrite] for POS = strPos
  ' History:
  ' 16-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function MorphAddSfxRule(ByVal strSfx As String, ByVal strRewrite As String, ByVal strPos As String, ByVal strVern As String) As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim dtrNew As DataRow     ' New datarow

    Try
      ' Validate
      If (strSfx = "") Then Return False
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Look for the entry
      dtrFound = tdlMorphDict.Tables("Suffix").Select("Sfx='" & strSfx.Replace("'", "''") & "' AND Rew='" & strRewrite.Replace("'", "''") & _
                                                     "' AND Pos='" & strPos.Replace("'", "''") & "'")
      If (dtrFound.Length > 0) Then
        ' It is already there -- get it
        dtrNew = dtrFound(0)
        dtrNew.Item("Freq") += 1
      Else
        ' It is not yet there, so create it
        dtrNew = AddOneDataRow(tdlMorphDict, "Suffix", "SuffixId", "SuffixList")
        With dtrNew
          .Item("Sfx") = strSfx : .Item("Pos") = strPos : .Item("Rew") = strRewrite : .Item("Freq") = 1
          .Item("Vern") = strVern
        End With
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphAddSfxRule error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphSuffixReport
  ' Goal:   Make an HTML report of the suffixes and their frequencies
  ' History:
  ' 19-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphSuffixReport() As Boolean
    Dim colHtml As New StringColl ' Here we store each HTML line
    Dim strHtml As String         ' Complete report
    Dim strSfxRepFile As String   ' File where we store the report
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim intI As Integer           ' Counter

    Try
      ' Get the report
      colHtml.Add("<html><body><h1>Suffix report</h1><p><table>" & _
          "<tr><td>POS</td><td align='right'>Freq</td><td>Suffix</td><td>Rewrite</td><td>Vern</td></tr>")
      ' Get correct table
      dtrFound = tdlMorphDict.Tables("Suffix").Select("", "Pos ASC, Freq DESC, Sfx ASC")
      ' Walk all the lines
      For intI = 0 To dtrFound.Length - 1
        With dtrFound(intI)
          colHtml.Add("<tr><td>" & .Item("Pos") & "</td><td align='right'>" & .Item("Freq") & "</td><td>" & .Item("Sfx") & "</td><td>" & _
                      .Item("Rew") & "</td><td>" & .Item("Vern") & "</td></tr>")
        End With
      Next intI
      ' Finish the HTML file
      colHtml.Add("</table></body></html>")
      strHtml = colHtml.Text
      ' Save the morpholoy report somewhere
      strSfxRepFile = GetDocDir() & "\SuffixReport.htm"
      IO.File.WriteAllText(strSfxRepFile, strHtml)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphSuffixReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphDistReport
  ' Goal:   Make an HTML report of the distances, lengths and their frequencies
  ' History:
  ' 30-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphDistReport() As Boolean
    Dim colHtml As New StringColl ' Here we store each HTML line
    Dim strHtml As String         ' Complete report
    Dim strSfxRepFile As String   ' File where we store the report
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim intI As Integer           ' Counter

    Try
      ' Get the report
      colHtml.Add("<html><body><h1>Distance report</h1><p><table>" & _
          "<tr><td>Length</td><td>Cost</td><td align='right'>Freq</td><td>Type</td><td>Example</td></tr>")
      ' Get correct table
      dtrFound = tdlMorphDict.Tables("Dist").Select("", "Len DESC, Cost ASC, Freq DESC")
      ' Walk all the lines
      For intI = 0 To dtrFound.Length - 1
        With dtrFound(intI)
          colHtml.Add("<tr><td>" & .Item("Len") & "</td><td>" & .Item("Cost") & "</td><td align='right'>" & .Item("Freq") & "</td>" & _
                    "<td>" & .Item("Type") & "</td><td>" & .Item("Vern") & " <i>" & .Item("Lemma") & "</i></td></tr>")
        End With
      Next intI
      ' Finish the HTML file
      colHtml.Add("</table></body></html>")
      strHtml = colHtml.Text
      ' Save the distance report somewhere
      strSfxRepFile = GetDocDir() & "\DistReport.htm"
      IO.File.WriteAllText(strSfxRepFile, strHtml)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphDistReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphSimiReport
  ' Goal:   Make an HTML report of the similarity values, lengths and their frequencies
  ' History:
  ' 30-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphSimiReport() As Boolean
    Dim colHtml As New StringColl ' Here we store each HTML line
    Dim strHtml As String         ' Complete report
    Dim strSfxRepFile As String   ' File where we store the report
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim intI As Integer           ' Counter

    Try
      ' Get the report
      colHtml.Add("<html><body><h1>Similarity report</h1><p><table>" & _
          "<tr><td>Length</td><td>SimiValue</td><td align='right'>Freq</td><td>Type</td><td>Example</td></tr>")
      ' Get correct table
      dtrFound = tdlMorphDict.Tables("Simi").Select("", "Len DESC, Value ASC, Freq DESC")
      ' Walk all the lines
      For intI = 0 To dtrFound.Length - 1
        With dtrFound(intI)
          colHtml.Add("<tr><td>" & .Item("Len") & "</td><td>" & .Item("Value") & "</td><td align='right'>" & .Item("Freq") & "</td>" & _
                    "<td>" & .Item("Type") & "</td><td>" & .Item("Vern") & " <i>" & .Item("Lemma") & "</i></td></tr>")
        End With
      Next intI
      ' Finish the HTML file
      colHtml.Add("</table></body></html>")
      strHtml = colHtml.Text
      ' Save the distance report somewhere
      strSfxRepFile = GetDocDir() & "\SimiReport.htm"
      IO.File.WriteAllText(strSfxRepFile, strHtml)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphSimiReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetSuffixRewrite
  ' Goal:   Compare [strVern] with the lemma [strLemma] and see if a rewrite rule
  '           can be formed with [strSfx] as suffix and [strRewrite] as rewrite
  ' History:
  ' 16-03-2013  ERK Created
  ' 14-02-2014  ERK No more 0 suffix adding
  ' ------------------------------------------------------------------------------------
  Private Function GetSuffixRewrite(ByVal strVern As String, ByVal strLemma As String, _
                                      ByRef strSfx As String, ByRef strRewrite As String) As Boolean
    Dim intPos As Integer   ' Position
    Dim intLen As Integer   ' length of the minimum
    Dim intPtc As Integer = 35  ' Minimum length percentage
    Dim bBreak As Boolean   ' Reached break-point

    Try
      ' Validate
      If (strVern = "") OrElse (strLemma = "") Then Return False
      ' ================ DEBUG =========
      ' If (InStr(strVern, "abhorr") > 0) Then Stop
      ' ================================
      ' Compare vern and lemma --> if they are equal, we have no rewrite rule
      If (strVern = strLemma) Then Return False
      intLen = Math.Min(strVern.Length, strLemma.Length) : bBreak = False
      ' Look for commonalities
      For intPos = 1 To intLen
        ' Check if they are still equal
        If (Mid(strVern, intPos, 1) <> Mid(strLemma, intPos, 1)) Then bBreak = True : intPos -= 1 : Exit For
      Next intPos
      ' If we have reached the length of the smallest, then go one back
      If (intPos - 1 >= intLen) Then intPos -= 1
      ' Where is [intPos] with respect to the total size of [strVern]?
      If (intPos * 100 \ strVern.Length <= intPtc) Then Return False
      ' We have a rewrite-possibility: set up the rule
      strSfx = Mid(strVern, intPos)
      strRewrite = Mid(strLemma, intPos)
      ' Double check
      If (strSfx.Length * 100 \ strVern.Length > (100 - intPtc)) Then Return False
      If (strRewrite.Length * 100 \ strLemma.Length > (100 - intPtc)) Then Return False
      ' Do not allow empty suffixes
      If (strSfx = "") Then Return False
      ' ========= DEBUG =========
      'If (strSfx = strRewrite) Then Stop
      'If (strRewrite <> "") Then Stop
      ' =========================
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/GetSuffixRewrite error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphDeriveRules
  ' Goal:   Extract rewrite rules in table [VernPos]:
  '           they turn a derived POS category into a main category (a "head" category)
  '         Make use of the information in table [Cat]
  ' History:
  ' 13-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphDeriveRules() As Boolean
    Dim dtrFound() As DataRow ' Table of remainder
    Dim dtrCat() As DataRow   ' Category check 
    Dim dtrNear() As DataRow  ' Result of looking for @Near ones
    Dim strVern As String     ' Current vern
    Dim strVernRe As String   ' Rewrite of current vernacular
    Dim strPos As String      ' Current POS
    Dim strLemma As String    ' Lemma
    Dim strNear As String     ' Near one
    Dim strSfx As String      ' Suffix
    Dim strRewrite As String  ' How the suffix must be re-written to get the lemma
    Dim strHeadPos As String  ' POS of the head category
    Dim strPath As String     ' Path of rules
    Dim strType As String     ' Type of VernPos entry
    'Dim arRule() As String    ' Array of rules
    Dim intI As Integer       ' Counter
    Dim intJ As Integer       ' Counter
    Dim intPtc As Integer     ' Percentage
    Dim intNum As Integer = 0 ' Number of changes
    Dim intCum As Integer = 0 ' Cumulative
    Dim intOps As Integer = 0 ' Number of operations
    Dim intCost As Integer    ' Costs
    Dim intValue As Integer   ' Value
    Dim intValueRe As Integer   ' Rewrite similarity value
    Dim bDeriveEmpty As Boolean ' Whether there are derive rules already or not
    Dim bEditOpEmpty As Boolean ' Whether edit operations are empty or not
    Dim bDistEmpty As Boolean   ' Whether distance measurements are empty or not
    Dim bSimiEmpty As Boolean   ' Whether similarity measurements are empty or not
    Dim bEnvEmtpy As Boolean    ' Wheterh environment frequency calculation has been done yet or not
    Dim colRewr As New StringColl

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Get the categories that need checking
      dtrCat = tdlMorphDict.Tables("Cat").Select("Type = 'Derived'")
      ' Check on derive rules
      bDeriveEmpty = (tdlMorphDict.Tables("Derive").Rows.Count < 10)
      bEditOpEmpty = (tdlMorphDict.Tables("EditOp").Rows.Count < 10)
      bDistEmpty = (tdlMorphDict.Tables("Dist").Rows.Count < 10)
      bSimiEmpty = (tdlMorphDict.Tables("Simi").Rows.Count < 10)
      bEnvEmtpy = (tdlMorphDict.Tables("Env").Rows.Count < 10)
      ' Initialisations for some varieties of English: do not use Edit, Dist, Simi and Env
      If (DoLike(strMorphDictFile, "*DictME*|*DictModE*|*DictEmodE*")) Then
        bEditOpEmpty = False : bDistEmpty = False : bSimiEmpty = False : bEnvEmtpy = False
      End If
      ' Get to the [Morph] table, and look for suffix-rewrite rules
      dtrFound = tdlMorphDict.Tables("VernPos").Select("", "Pos ASC, Vern ASC")
      ' Walk them all
      For intI = 0 To dtrFound.Length - 1
        If (bInterrupt) Then Return False
        ' Where are we?
        intPtc = (100 * (intI + 1) \ dtrFound.Length)
        Status("Morph derivation rules detection " & intPtc & "%", intPtc)
        ' Get the vern and pos and the lemma
        With dtrFound(intI)
          strVern = .Item("Vern") : strPos = .Item("Pos") : strLemma = .Item("l") : strType = .Item("Type")
        End With
        ' Middle-Englihs: adapt lemma
        strLemma = strLemma.Replace("-", "")
        ' Other initializations
        strSfx = "" : strRewrite = "" : strPath = "" : intCum = 400
        ' Check whether the category is one of those where we need derive-rules
        dtrCat = tdlMorphDict.Tables("Cat").Select("Type = 'Derived' AND Pos = '" & strPos & "'")
        If (dtrCat.Length > 0) AndAlso (InStr(strVern, "'") = 0) AndAlso (InStr(strLemma, "'") = 0) Then
          ' Do we need to do derive rules?
          If (bDeriveEmpty) Then
            ' Get a derive-rewrite rule
            If (GetSuffixRewrite(strVern, strLemma, strSfx, strRewrite)) Then
              ' There is a derive-rewrite rule, so now get the HEAD
              strHeadPos = dtrCat(0).Item("Head")
              ' Add this rule to the Rule table
              If (Not MorphAddDeriveRule(strSfx, strRewrite, strPos, strHeadPos, strVern, "Main")) Then Logging("modMorph/MorphDeriveRules error #1") : Return False
              ' We found one!
              intNum += 1
            End If
            ' Check if there is a @Near feature in this entry of table [Cat]
            strHeadPos = dtrCat(0).Item("Near").ToString
            If (strHeadPos <> "") Then
              ' Yes, there is a @Near feature --> Are there entries in [VernPos] with the @Near as POS and the [strLemma] as lemma?
              dtrNear = tdlMorphDict.Tables("VernPos").Select("Pos = '" & strHeadPos & "' AND l='" & strLemma & "'")
              ' Walk the results
              For intJ = 0 To dtrNear.Length - 1
                With dtrNear(intJ)
                  ' Regard this item as the "lemma" for the moment
                  strNear = .Item("Vern")
                End With
                strSfx = "" : strRewrite = ""
                ' Try get a derive-rewrite rule
                If (GetSuffixRewrite(strVern, strNear, strSfx, strRewrite)) Then
                  ' There is a derive-rewrite rule
                  ' Add this rule to the Rule table
                  If (Not MorphAddDeriveRule(strSfx, strRewrite, strPos, strHeadPos, strVern, "Deriv")) Then Logging("modMorph/MorphDeriveRules error #2") : Return False
                  ' We found one!
                  intNum += 1
                End If
              Next intJ
            End If
          End If
          ' De we need to make EditOp elements?
          If (bEditOpEmpty) AndAlso (strType <> "LemmaAmbi") AndAlso (InStr(strLemma, ";") = 0) Then
            ' Derive edit op rules for the transformation from [strVern] to [strLemma]
            intCost = SimpleDistance(strVern, strLemma, strPath, intOps, intCum)
            ' Add the rule(s)
            If (Not MorphAddEditRule(strVern, strLemma, strPos, strPath)) Then Logging("modMorph/MorphDeriveRules error #3") : Return False
            ' Keep track of matters
            intNum += 1
          End If
          ' De we need to make Dist elements?
          If (bDistEmpty) AndAlso (strType <> "LemmaAmbi") AndAlso (InStr(strLemma, ";") = 0) AndAlso (InStr(strVern, "&") = 0) Then
            ' Derive edit op rules for the transformation from [strVern] to [strLemma]
            intCost = SimpleDistance(strVern, strLemma, strPath, intOps, intCum)
            ' Add the rule(s)
            If (Not MorphAddDist(strVern, strLemma, strPos, intCost)) Then Logging("modMorph/MorphDeriveRules error #4") : Return False
            ' Keep track of matters
            intNum += 1
          End If
          ' De we need to make Dist elements?
          If (bSimiEmpty) AndAlso (strType <> "LemmaAmbi") AndAlso (InStr(strLemma, ";") = 0) AndAlso (InStr(strVern, "&") = 0) Then
            ' OLD: intValue = 1000 * objSim.CompareStrings(strVern, strLemma)
            ' OLD: intValue = SimpleSimilarity(strVern, strLemma)
            ' First get sinmilarity of strVern
            intValue = objSim.LevenSimi(strVern, strLemma)
            ' See if we can get a suffix-rewrite
            If (Not MorphSuffixRewrite(strVern, strPos, colRewr)) Then Return False
            strVernRe = IIf(colRewr.Count = 0, strVern, colRewr.Item(0))
            If (strVern <> strVernRe) Then
              ' Then find the similarity between the re-write of [strVern] and the lemma
              intValueRe = objSim.LevenSimi(strVernRe, strLemma)
              ' See which one is better
              If (intValueRe > intValue) Then intValue = intValueRe
            End If
            ' Add the rule(s)
            If (Not MorphAddSimi(strVern, strLemma, strPos, intValue)) Then Logging("modMorph/MorphDeriveRules error #5") : Return False
            ' ========== DEBUGGING ===========
            If (intValue < 50) Then
              Debug.Print(strPos & ";" & strVern & " L=" & strLemma & " S=" & intValue & "%")
            End If
            ' ================================
            ' Keep track of matters
            intNum += 1
          End If
          If (bEnvEmtpy) AndAlso (strType <> "LemmaAmbi") AndAlso (InStr(strLemma, ";") = 0) AndAlso (InStr(strVern, "&") = 0) Then
            ' Derive environment frequences from [strLemma]
            If (Not MorphAddEnv(strLemma)) Then Logging("modMorph/MorphDeriveRules error #6") : Return False
            ' Keep track
            intNum += 1
          End If
        End If
      Next intI
      ' Show what we have done
      Logging("MorphDeriveRules found [" & intNum & "] matches")
      ' Save file if there are more matches
      If (intNum > 0) Then
        Status("Saving adapted MorphDict...")
        tdlMorphDict.WriteXml(strMorphDictFile)
        If (bDistEmpty) Then
          Status("Creating distance report...")
          If (Not MorphDistReport()) Then Return False
        End If
        If (bSimiEmpty) Then
          Status("Creating similarity report...")
          If (Not MorphSimiReport()) Then Return False
        End If
        Status("Creating derive report...")
        If (Not MorphSuffixReport()) Then Return False
        Status("Ready")
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphDeriveRules error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphSuffixRewrite
  ' Goal:   See if any rewrite rule matches [strVern]+[strPos], and if so, apply it
  ' History:
  ' 03-06-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function MorphSuffixRewrite(ByVal strVern As String, ByVal strPos As String, ByRef colRewr As StringColl) As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim strSfx As String      ' The suffix
    Dim intI As Integer       ' Counter

    Try
      ' Validate
      If (strVern = "") Then Return ""
      ' If no POS is supplied, then do not try to rewrite
      If (strPos = "") Then Return strVern
      ' Initialise
      colRewr.Clear()
      ' Get all the rewrite rules that could apply, and get them in the right order
      dtrFound = tdlMorphDict.Tables("Suffix").Select("Pos='" & strPos & "' AND Freq > 2", "Freq DESC")
      ' If there is nothing, get anything
      If (dtrFound.Length = 0) Then
        dtrFound = tdlMorphDict.Tables("Suffix").Select("Pos='" & strPos & "'", "Freq DESC")
      End If
      For intI = 0 To dtrFound.Length - 1
        ' See if this rule applies
        With dtrFound(intI)
          strSfx = .Item("Sfx")
          If (Right(strVern, strSfx.Length) = strSfx) Then
            ' Suffix applies, so make the rewrite
            colRewr.AddUnique(Left(strVern, strVern.Length - strSfx.Length) & .Item("Rew").ToString)
          End If
        End With
      Next intI
      ' No rewrite applies, so return the vernacular
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphSuffixRewrite error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphAddDeriveRule
  ' Goal:   Add a rule from [strSfx] > [strRewrite] for  POS = strPos
  '                                             and for Head = strHead
  ' History:
  ' 13-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function MorphAddDeriveRule(ByVal strSfx As String, ByVal strRewrite As String, ByVal strPos As String, ByVal strHead As String, _
                                      ByVal strVern As String, ByVal strType As String) As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim dtrNew As DataRow     ' New datarow

    Try
      ' Validate
      If (strSfx = "") Then Return False
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Look for the entry
      dtrFound = tdlMorphDict.Tables("Derive").Select("Sfx='" & strSfx.Replace("'", "''") & "' AND Rew='" & strRewrite.Replace("'", "''") & _
            "' AND Head='" & strHead & "' AND Pos='" & strPos.Replace("'", "''") & "'")
      If (dtrFound.Length > 0) Then
        ' It is already there -- get it
        dtrNew = dtrFound(0)
        dtrNew.Item("Freq") += 1
      Else
        ' It is not yet there, so create it
        dtrNew = AddOneDataRow(tdlMorphDict, "Derive", "DeriveId", "DeriveList")
        With dtrNew
          .Item("Sfx") = strSfx : .Item("Pos") = strPos : .Item("Rew") = strRewrite : .Item("Freq") = 1
          .Item("Vern") = strVern : .Item("Head") = strHead : .Item("Type") = strType
        End With
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphAddDeriveRule error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphAddEditRule
  ' Goal:   Add a rule from [strSrc] > [strRewrite] for  POS = strPos
  ' History:
  ' 21-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function MorphAddEditRule(ByVal strVern As String, ByVal strLemma As String, ByVal strPos As String, _
                                      ByVal strPath As String) As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim dtrNew As DataRow     ' New datarow
    Dim arOp() As String      ' Array of edit operations
    Dim arLine() As String    ' One line
    Dim strLine As String     ' One line
    Dim strOp As String       ' One operation
    Dim strS As String        ' Source letter
    Dim strT As String        ' Rewrite (target) letter
    Dim intI As Integer       ' Counter

    Try
      ' Validate
      If (strVern = "") OrElse (strLemma = "") Then Return False
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Get the array of edit operations and perform other initialisations
      arOp = Split(strPath, "|") : strOp = "" : strS = "" : strT = ""
      ' Walk all operations
      For intI = 0 To arOp.Length - 1
        ' Get this line
        strLine = Trim(arOp(intI))
        If (strLine <> "") Then
          ' Split the line
          arLine = Split(strLine, ";")
          ' Action depends on the first character
          strOp = arLine(0)
          Select Case strOp
            Case "i", "a"
              ' Get source and target character
              strS = "" : strT = arLine(1)
            Case "d"
              ' Get source and target character
              strS = arLine(1) : strT = ""
            Case "s"
              ' Get source and target letters
              strS = arLine(1) : strT = arLine(2)
            Case "e"
              ' There is no need to keep track of the 'equality' operations
            Case Else
              Logging("MorphAddEditRule: unknown edit operation [" & strOp & "]")
              bInterrupt = True
              Return False
          End Select
        End If
        If (strOp <> "e") AndAlso (InStr(strS, "'") = 0) AndAlso (InStr(strT, "'") = 0) Then
          ' Look for the entry
          dtrFound = tdlMorphDict.Tables("EditOp").Select("Src='" & strS & "' AND Rew='" & strT & _
                "' AND Type='" & strOp & "' AND Pos='" & strPos & "'")
          If (dtrFound.Length > 0) Then
            ' It is already there -- get it
            dtrNew = dtrFound(0)
            dtrNew.Item("Freq") += 1
          Else
            ' It is not yet there, so create it
            dtrNew = AddOneDataRow(tdlMorphDict, "EditOp", "EditOpId", "EditOpList")
            With dtrNew
              .Item("Src") = strS : .Item("Pos") = strPos : .Item("Rew") = strT : .Item("Freq") = 1
              .Item("Exmp") = strVern & " > " & strLemma : .Item("Type") = strOp
            End With
          End If
        End If
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphAddEditRule error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphAddDist
  ' Goal:   Add a distance measure from [strVern] > [strLemma] (for  POS = strPos)
  ' History:
  ' 30-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function MorphAddDist(ByVal strVern As String, ByVal strLemma As String, ByVal strPos As String, _
                                      ByVal intCost As Integer) As Boolean
    Dim dtrFound() As DataRow   ' Result of SELECT
    Dim dtrNew As DataRow       ' New datarow
    Dim strAlt As String = ""   ' Alternative lemma
    Dim strType As String = ""  ' Type: unique or ambiguous
    Dim intLen As Integer       ' Length of vernacular

    Try
      ' Validate
      If (strVern = "") OrElse (strLemma = "") Then Return False
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Skip those that are equal
      If (strVern = strLemma) Then Return True
      ' Determine the length
      intLen = strVern.Length
      ' Check if there is a non-lemma with the same distance (costs) or even closer by
      If (GetOEdictEntryAlt(strVern, strLemma, strPos, intCost, strAlt)) Then
        ' Set the type
        strType = "ambi"
      Else
        ' Set the type
        strType = "uniq"
      End If
      ' Look for the entry
      dtrFound = tdlMorphDict.Tables("Dist").Select("Len=" & intLen & " AND Cost=" & intCost)
      If (dtrFound.Length > 0) Then
        ' It is already there -- get it
        dtrNew = dtrFound(0)
        dtrNew.Item("Freq") += 1
        If (strType = "ambi") AndAlso (dtrNew.Item("Type") = "uniq") Then dtrNew.Item("Type") = "ambi"
      Else
        ' It is not yet there, so create it
        dtrNew = AddOneDataRow(tdlMorphDict, "Dist", "DistId", "DistList")
        With dtrNew
          .Item("Len") = intLen : .Item("Cost") = intCost : .Item("Freq") = 1
          .Item("Vern") = strVern : .Item("Lemma") = strLemma : .Item("Type") = strType
        End With
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphAddDist error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphAddSimi
  ' Goal:   Add a similarity measure from [strVern] > [strLemma] (for  POS = strPos)
  ' History:
  ' 30-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function MorphAddSimi(ByVal strVern As String, ByVal strLemma As String, ByVal strPos As String, _
                                      ByVal intValue As Integer) As Boolean
    Dim dtrFound() As DataRow   ' Result of SELECT
    Dim dtrNew As DataRow       ' New datarow
    Dim strAlt As String = ""   ' Alternative lemma
    Dim strType As String = ""  ' Type: unique or ambiguous
    Dim intLen As Integer       ' Length of vernacular

    Try
      ' Validate
      If (strVern = "") OrElse (strLemma = "") Then Return False
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' We are not really interested in exact matches
      If (strVern = strLemma) Then Return True
      ' Determine the length
      intLen = strVern.Length
      ' Check if there is a non-lemma with the same similarity (value) or even larger
      If (GetOEdictEntrySimi(strVern, strLemma, strPos, intValue, strAlt)) Then
        ' Set the type
        strType = "ambi"
      Else
        ' Set the type
        strType = "uniq"
      End If
      ' Look for the entry
      dtrFound = tdlMorphDict.Tables("Simi").Select("Len=" & intLen & " AND Value=" & intValue)
      If (dtrFound.Length > 0) Then
        ' It is already there -- get it
        dtrNew = dtrFound(0)
        dtrNew.Item("Freq") += 1
        If (strType = "ambi") AndAlso (dtrNew.Item("Type") = "uniq") Then dtrNew.Item("Type") = "ambi"
      Else
        ' It is not yet there, so create it
        dtrNew = AddOneDataRow(tdlMorphDict, "Simi", "SimiId", "SimiList")
        With dtrNew
          .Item("Len") = intLen : .Item("Value") = intValue : .Item("Freq") = 1
          .Item("Vern") = strVern : .Item("Lemma") = strLemma : .Item("Type") = strType
        End With
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphAddSimi error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphAddEnv
  ' Goal:   Add all the environments available in [strLemma]
  ' History:
  ' 31-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function MorphAddEnv(ByVal strLemma As String) As Boolean
    Dim dtrFound() As DataRow   ' Result of SELECT
    Dim dtrNew As DataRow       ' New datarow
    Dim strAlt As String = ""   ' Alternative lemma
    Dim strType As String = ""  ' Type: unique or ambiguous
    Dim strPre As String = ""   ' Preceding context
    Dim strFocus As String = "" ' Focus character
    Dim strFol As String = ""   ' Following character
    Dim intLen As Integer       ' Length of vernacular
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (strLemma = "") Then Return False
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Determine the length
      intLen = strLemma.Length
      ' Walk the whole string
      For intI = 1 To intLen
        ' Get the values for strPre, strFocus and strFol
        If (intI = 1) Then strPre = "-"
        ' Get focus character
        strFocus = Mid(strLemma, intI, 1)
        ' Get following character
        If (intI < intLen) Then
          strFol = Mid(strLemma, intI + 1, 1)
        Else
          strFol = "-"
        End If
        ' Process entry
        dtrFound = tdlMorphDict.Tables("Env").Select("Pre='" & strPre & "' AND Focus='" & strFocus & "' AND Fol='" & strFol & "'")
        If (dtrFound.Length > 0) Then
          ' It is already there -- get it
          dtrNew = dtrFound(0)
          dtrNew.Item("Freq") += 1
        Else
          ' It is not yet there, so create it
          dtrNew = AddOneDataRow(tdlMorphDict, "Env", "EnvId", "EnvList")
          With dtrNew
            .Item("Pre") = strPre : .Item("Focus") = strFocus : .Item("Freq") = 1
            .Item("Fol") = strFol
          End With
        End If
        ' Shift environment
        strPre = strFocus
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphAddEnv error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphDeriveReport
  ' Goal:   Make an HTML report of the suffixes and their frequencies
  ' History:
  ' 19-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphDeriveReport() As Boolean
    Dim colHtml As New StringColl ' Here we store each HTML line
    Dim strHtml As String         ' Complete report
    Dim strSfxRepFile As String   ' File where we store the report
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim intI As Integer           ' Counter

    Try
      ' Get the report
      colHtml.Add("<html><body><h1>Derive report</h1><p><table>" & _
          "<tr><td>POS</td><td>Head</td><td align='right'>Freq</td><td>Suffix</td><td>Rewrite</td><td>Vern</td></tr>")
      ' Get correct table
      dtrFound = tdlMorphDict.Tables("Derive").Select("", "Pos ASC, Freq DESC, Sfx ASC")
      ' Walk all the lines
      For intI = 0 To dtrFound.Length - 1
        With dtrFound(intI)
          colHtml.Add("<tr><td>" & .Item("Pos") & "</td><td>" & .Item("Head") & "</td><td align='right'>" & .Item("Freq") & _
                      "</td><td>" & .Item("Sfx") & "</td><td>" & _
                      .Item("Rew") & "</td><td>" & .Item("Vern") & "</td></tr>")
        End With
      Next intI
      ' Finish the HTML file
      colHtml.Add("</table></body></html>")
      strHtml = colHtml.Text
      ' Save the morpholoy report somewhere
      strSfxRepFile = GetDocDir() & "\DeriveReport.htm"
      IO.File.WriteAllText(strSfxRepFile, strHtml)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphDeriveReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphLemmaAmbi
  ' Goal:   Check all items in table [VernPos] of type "LemmaAmbi"
  '         Ask user if one lemma should be taken, and if so, which one
  ' History:
  ' 26-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphLemmaAmbi() As Boolean
    Dim dtrFound() As DataRow       ' Table with ambiguous entries
    Dim strVern As String           ' Vernacular word
    Dim strPos As String            ' POS tag
    Dim strLemma As String          ' Current lemma
    Dim strLemOne As String         ' One lemma
    Dim strFeat As String           ' Associated feature(s)
    Dim intI As Integer             ' Counter
    Dim intJ As Integer             ' Counter
    Dim intPtc As Integer           ' Percentage
    Dim bChanged As Boolean = False ' Are we changed?
    Dim bAmbi As Boolean            ' Lemma is ambiguous
    Dim arLemma() As String         ' Array of lemma's
    Dim strLogFile As String = ""   ' File to which I write a log-report

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Read dictionary
      If (Not MorphReadOEdict()) Then Return False
      ' Make the name for the log-file
      strLogFile = GetDocDir() & "\lemma_ambi.log"
      ' Get all ambiguous entries in [VernPos]
      dtrFound = tdlMorphDict.Tables("VernPos").Select("Type='LemmaAmbi'", "Pos ASC, Vern ASC")
      ' Walk all these entries
      For intI = 0 To dtrFound.Length - 1
        ' Access this entry
        With dtrFound(intI)
          strVern = .Item("Vern") : strPos = .Item("Pos") : strLemma = .Item("l") : strFeat = .Item("f").ToString
        End With
        ' Show where we are
        intPtc = (intI + 1) * 100 \ dtrFound.Length
        Status("LemmaAmbi " & intPtc & "% (i=" & intI + 1 & "/" & dtrFound.Length & ")", intPtc)
        ' First check if the lemma for this entry is ambiguous after all
        arLemma = Split(strLemma, ";") : bAmbi = False
        strLemOne = Split(arLemma(1), "|")(0)
        For intJ = 2 To arLemma.Length - 1
          If (Split(arLemma(intJ), "|")(0) <> strLemOne) Then
            bAmbi = True : Exit For
          End If
        Next intJ
        If (bAmbi) Then
          ' Ambiguous --> Ask user
          With frmLemmaAmbi
            ' Supply form with necessary information
            .Vern = strVern
            .Pos = strPos
            .Lemma = strLemma
            .Feat = strFeat
            ' Show and ask user
            Select Case .ShowDialog
              Case DialogResult.Cancel
                ' User wants to stop
                Exit For
              Case DialogResult.Ignore
                ' Do not take this one into account
              Case DialogResult.OK
                ' Okay, process the user's choice
                dtrFound(intI).Item("l") = .LemmaChosen
                dtrFound(intI).Item("f") = .Feat
                dtrFound(intI).Item("Type") = "LemmaOnly"
                ' Show what we have done
                IO.File.AppendAllText(strLogFile, "Vern=[" & strVern & "] Pos=[" & strPos & "] lemma=[" & strLemma & "] Chosen=[" & .LemmaChosen & "]" & vbCrLf)
            End Select
          End With
        Else
          ' Not ambiguous --> Process the lemma
          dtrFound(intI).Item("l") = strLemOne
          dtrFound(intI).Item("f") = ""
          dtrFound(intI).Item("Type") = "LemmaOnly"
          ' Show what we have done
          IO.File.AppendAllText(strLogFile, "Vern=[" & strVern & "] Pos=[" & strPos & "] lemma=[" & strLemma & "] Chosen=[" & strLemOne & "]" & vbCrLf)
        End If
      Next intI
      ' Make sure all changes are processed
      tdlMorphDict.AcceptChanges()
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphLemmaAmbi error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetPrefixRewrite
  ' Goal:   See if there is a prefix rewrite possible for [strIn], and if so return it
  '           inside [strVernRew]
  ' History:
  ' 23-04-2013  ERK Created
  ' 17-05-2013  ERK May return more than one possible rewrite
  ' ------------------------------------------------------------------------------------
  Private Function GetPrefixRewrite(ByVal strIn As String, ByRef colVernRew As StringColl) As Boolean
    Dim dtrFound() As DataRow     ' Table of prefixes
    Dim strVernRew As String = "" ' Rewrite
    Dim intI As Integer           ' Counter

    Try
      ' Validate
      If (strIn = "") Then Return False
      ' Initialise
      colVernRew.Clear()
      ' Get all prefixes that may match
      dtrFound = tdlMorphDict.Tables("Prefix").Select("", "Size DESC")
      For intI = 0 To dtrFound.Length - 1
        With dtrFound(intI)
          If (strIn Like .Item("Pre").ToString & "*") Then
            ' We've got the best matching prefix
            strVernRew = .Item("Rew").ToString & Mid(strIn, .Item("Pre").ToString.Length + 1)
            ' Add to collection
            colVernRew.Add(strVernRew)
            '            strVernRew = Left(strIn, strIn.Length - .Item("Pre").ToString.Length) & .Item("Rew").ToString
            Return True
          End If
        End With
      Next intI
      ' Are there any results?
      If (colVernRew.Count > 0) Then Return True
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/GetPrefixRewrite error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphCatList
  ' Goal:   Given a main POS category, get a list of all the categories belonging to it
  ' History:
  ' 06-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphCatList(ByVal strPosHead As String, Optional ByVal bDoLike As Boolean = False) As String
    Dim dtrFound() As DataRow     ' Result of select
    Dim strPosList As String = "" ' All categories belonging to this one
    Dim intI As Integer           ' Counter

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) OrElse (tdlMorphDict.Tables("Cat") Is Nothing) Then Return False
      ' Initialise the list
      If (bDoLike) Then
        strPosList = strPosHead
      Else
        strPosList = "(Pos='" & strPosHead & "'"
      End If
      ' Find the category
      dtrFound = tdlMorphDict.Tables("Cat").Select("Head='" & strPosHead & "'")
      If (dtrFound.Length > 0) Then
        ' Add all these categories
        For intI = 0 To dtrFound.Length - 1
          If (bDoLike) Then
            strPosList &= "|" & dtrFound(intI).Item("Pos").ToString
          Else
            strPosList &= " OR Pos='" & dtrFound(intI).Item("Pos").ToString & "'"
          End If
        Next intI
      End If
      ' Finish the list
      If (Not bDoLike) Then strPosList &= ")"
      ' Return positively
      Return strPosList
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphCatList error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphCatMain
  ' Goal:   Given a POS, provide the main POS category from table [Cat]
  ' History:
  ' 06-05-2014  ERK Created
  ' 11-06-2014  ERK Added [bRestricted] option
  ' ------------------------------------------------------------------------------------
  Public Function MorphCatMain(ByVal strPos As String, Optional ByVal bRestricted As Boolean = False) As String
    Dim dtrFound() As DataRow   ' Result of select
    Dim strPosMain As String    ' Main category
    Dim intI As Integer         ' Counter
    Dim arMain() As String = {"*ADJ*", "*ADV*", "BE|*+BE|HV|*+HV|MD|*+MD|AX|*+AX", "NR", "D|*PRO*|*Q*"}
    Dim arRest() As String = {"ADJ", "ADV", "VB", "N", "pronoun"}

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) OrElse (tdlMorphDict.Tables("Cat") Is Nothing) Then Return False
      ' Find the category
      dtrFound = tdlMorphDict.Tables("Cat").Select("Pos='" & strPos & "'")
      If (dtrFound.Length = 0) Then
        ' Ask the user what it is
        If (Not MorphCatAdd(strPos)) Then Return False
        ' Try again
        dtrFound = tdlMorphDict.Tables("Cat").Select("Pos='" & strPos & "'")
      End If
      If (dtrFound.Length = 0) Then
        ' Notify the user
        Logging("MorphCatMain: cannot find main category for [" & strPos & "]")
        strPosMain = ""
      ElseIf (dtrFound(0).Item("Type").ToString = "Main") Then
        ' This is the main category!
        strPosMain = strPos
      Else
        strPosMain = dtrFound(0).Item("Head").ToString
      End If
      ' Should we look restricted?
      If (bRestricted) Then
        For intI = 0 To arMain.Length - 1
          If (DoLike(strPosMain, arMain(intI))) Then
            strPosMain = arRest(intI) : Exit For
          End If
        Next intI
      End If
      ' Return positively
      Return strPosMain
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphCatMain error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  Public Function MorphGetPosDict(ByVal strPos As String, ByRef strPosNear As String) As String
    Dim dtrCat() As DataRow       ' Category
    Dim strPosDict As String = ""    ' POS according to dictionary

    Try
      ' Action depends on the POS we have
      Select Case strPos
        Case "VB", "NEG+VB", "RP+VB"
          strPosDict = "Vb"
        Case "N", "N^N"
          strPosDict = "N"
        Case "NR", "NR^N"
          strPosDict = "NR"
        Case "ADJ"
          strPosDict = "Adj"
        Case "ADV"
          strPosDict = "Adv"
        Case Else
          ' Get category by looking at the HEAD category in table [Cat]
          dtrCat = tdlMorphDict.Tables("Cat").Select("Pos='" & strPos & "'")
          If (dtrCat.Length > 0) Then
            ' Get the actual NEAR pos
            strPosNear = dtrCat(0).Item("Near").ToString
            ' Derive the [PosDict], if possible...
            If (DoLike(dtrCat(0).Item("Head").ToString, "*BE|*HV|*VB|*MD|*AX")) Then
              strPosDict = "Vb"
            ElseIf (DoLike(dtrCat(0).Item("Head").ToString, "ADJ*")) Then
              strPosDict = "Adj"
            ElseIf (DoLike(dtrCat(0).Item("Head").ToString, "ADV*")) Then
              strPosDict = "Adv"
            ElseIf (dtrCat(0).Item("Head").ToString = "N") Then
              strPosDict = "N"
            ElseIf (dtrCat(0).Item("Head").ToString = "NR") Then
              strPosDict = "NR"
            Else
              strPosDict = ""
            End If
          ElseIf (strPos Like "ADJ*") Then
            strPosDict = "Adj"
          ElseIf (strPos Like "ADV*") Then
            strPosDict = "Adv"
          Else
            ' Don't know the POS
            strPosDict = ""
          End If
      End Select
      ' Return the result
      Return strPosDict
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphGetPosDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMEprojLemmatizeRewrite_Click
  ' Goal:   Find the best matching derived form or rewritten lemma
  ' History:
  ' 09-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MorphLemmatizeBestMatch(ByRef intNum As Integer) As Boolean
    Dim strVern As String         ' Current vern
    Dim strVernRew As String = "" ' Vernacular rewritten with other prefix (if at all possible)
    Dim strPos As String          ' Current POS
    Dim strLemma As String        ' Lemma
    Dim strLemForm As String      ' Lemma based on form-match
    Dim strLemRewr As String      ' lemma based on rewrite rules
    Dim strClause As String       ' Clause
    Dim strLoc As String          ' Location
    Dim strDeriv As String = ""   ' Derivation that has been made
    Dim strDerivForm As String    ' Derivation from form
    Dim strDerivRewr As String    ' Derivation from rewrite
    Dim strFeatForm As String = ""   ' Features, including MED/OED or other number
    Dim strFeatRewr As String = ""   ' Features, including MED/OED or other number
    Dim strFeat As String = ""    ' Features
    Dim strFile As String         ' Filename
    Dim strSimFile As String      ' File name for similarities
    Dim dtrFound() As DataRow     ' Table of remainder
    Dim intForestId As Integer    ' Forest Id
    Dim intEtreeId As Integer     ' Etree Id
    Dim intI As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage
    Dim intScoreForm As Integer   ' Match score of the best match with the same POS
    Dim intScoreRewr As Integer   ' Match score of the best match after applying form-to-lemma rewrite rules
    Dim intScore As Integer       ' Match score (best overall)
    Dim intFreq As Integer        ' Frequency of best match - overall

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      If (tdlRemain Is Nothing) Then Return False
      ' Initialise
      mrp_colHand.Clear() : mrp_colSimi.Clear() : mrp_colFull.Clear() : intNum = 0 : intFreq = 0 : strDerivForm = "" : strDerivRewr = ""
      ' Get to the [Remain] table 
      ' dtrFound = tdlMorphDict.Tables("Remain").Select("Pos='VB'", "Pos ASC, Vern ASC, Freq DESC")
      dtrFound = tdlRemain.Tables("Remain").Select("", "Pos ASC, Vern ASC, Freq DESC")
      ' Walk them all
      For intI = 0 To dtrFound.Length - 1
        ' Check for interrupt
        If (bInterrupt) Then Return False
        ' Access the current [Remain] entry
        With dtrFound(intI)
          strVern = .Item("Vern") : strPos = .Item("Pos") : strLemma = ""
          strFile = .Item("File") : intForestId = .Item("forestId") : intEtreeId = .Item("eTreeId")
          strClause = .Item("Clause").ToString
          strLoc = strFile & ":" & intForestId
          intFreq = .Item("Freq")
        End With
        ' Other initialisations
        strLemForm = "" : strLemRewr = ""
        ' Where are we?
        intPtc = (100 * (intI + 1) \ dtrFound.Length)
        Status("Morph lemmatize by dictionary " & intPtc & "% found=[" & intNum & "] POS=" & strPos & " w=" & strVern, intPtc)
        ' First trial: look for *exact* match

        ' Find the best matching form with the same POS
        If (Not GetMorphDictMatch(strVern, strPos, "Form", strLemForm, intScoreForm, strDerivForm, strFeatForm)) Then Return False
        ' Find the best matching form after applying rewrite rules
        If (Not GetMorphDictMatch(strVern, strPos, "Rewrite", strLemRewr, intScoreRewr, strDerivRewr, strFeatRewr)) Then
          Select Case MsgBox("There is a hickup in Lemmatization. Continue anyway?", MsgBoxStyle.YesNoCancel)
            Case MsgBoxResult.No, MsgBoxResult.Cancel
              Return False
          End Select
        End If
        ' Check which is the winner
        If (intScoreForm > intScoreRewr) Then
          intScore = intScoreForm : strLemma = strLemForm : strDeriv = strDerivForm : strFeat = strFeatForm
        ElseIf (intScoreForm < intScoreRewr) Then
          intScore = intScoreRewr : strLemma = strLemRewr : strDeriv = strDerivRewr : strFeat = strFeatRewr
        ElseIf (intScoreForm = 0) Then
          ' Equal and zero
          intScore = intScoreForm : strLemma = strLemForm : strDeriv = strDerivForm : strFeat = strFeatForm
        Else
          ' They are equal, but not zero
          intScore = intScoreForm : strLemma = strLemForm : strDeriv = strDerivForm : strFeat = strFeatRewr
        End If
        ' Process the result
        Select Case intScore
          Case Is >= 97
            '' Automatically process in the [VernPos] table
            'If (Not VernPosAdd(strPos, strVern, "LemmaOnly", strLemma, "1", "", strDeriv & " score=" & intScore & " [" & Format(Now, "g") & "]")) Then
            '  Logging("VernPosAdd problem for [" & strVern & "]") : Return False
            'End If
            frmMain.tbGenSource.AppendText(strVern & " (" & strPos & ") l=" & strLemma & vbCrLf)
            ' Add to [Full] csv list
            If (Not AddSimiLine(strVern, strPos, "Vb", intScore, strDeriv, strLemma, strFeat, intFreq, strLoc, strClause, mrp_colFull)) Then Return False
            ' Keep track of number of instances
            intNum += 1
          Case Is >= 75
            ' Add to [Simi] csv list
            If (Not AddSimiLine(strVern, strPos, "Vb", intScore, strDeriv, strLemma, strFeat, intFreq, strLoc, strClause, mrp_colSimi)) Then Return False
          Case Else
            ' Add to [Hand] csv list
            If (Not AddSimiLine(strVern, strPos, "Vb", intScore, strDeriv, strLemma, strFeat, intFreq, strLoc, strClause, mrp_colHand)) Then Return False
        End Select
      Next intI
      ' Save the results of Full
      strSimFile = GetDocDir() & "\MorphDictFull_" & Today.Day & MonthName(Month(Today)) & "-auto.csv"
      IO.File.WriteAllText(strSimFile, mrp_colFull.Text)
      ' Inform user of these results
      Logging("Look for good matches in: " & strSimFile)
      ' Save the results of simi
      strSimFile = GetDocDir() & "\MorphDictSimi_" & Today.Day & MonthName(Month(Today)) & "-auto.csv"
      IO.File.WriteAllText(strSimFile, mrp_colSimi.Text)
      ' Inform user of these results
      Logging("Look for suggestions in: " & strSimFile)
      ' Save the results of hand
      strSimFile = GetDocDir() & "\MorphDictHand_" & Today.Day & MonthName(Month(Today)) & "-auto.csv"
      IO.File.WriteAllText(strSimFile, mrp_colHand.Text)
      ' Show what we have done
      Logging("MorphLemmatizeByDict found [" & intNum & "] <full> matches (accuracy >= 97%)")
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphLemmatizeBestMatch error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddSimiLine
  ' Goal:   Add one line to collection [colThis] with a suggestion of vernacular/lemma combination
  ' History:
  ' 09-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function AddSimiLine(ByVal strIn As String, ByVal strPos As String, ByVal strPosDict As String, ByVal intValMax As Integer, _
                                 ByVal strSfxType As String, ByVal strWordMax As String, ByVal strFeat As String, ByVal intFreq As Integer, _
                                 ByVal strLoc As String, ByVal strClause As String, ByRef colThis As StringColl) As Boolean
    Dim strPre As String = ""
    Dim strSimiLine As String = ""  ' The line we are adding

    Try
      ' Record what we have found in tabular form, so that the user could still make use of it
      strSimiLine = strPre & """" & strIn & """" & ";" & _
                      """" & strPos & """" & ";" & _
                      """" & strPosDict & """" & ";" & _
                      intValMax & ";" & _
                      """" & strSfxType & """" & ";" & _
                      """" & strWordMax & """" & ";" & _
                      """" & strFeat & """" & ";" & _
                      intFreq & ";" & _
                      """" & "[" & strLoc & "]" & """" & ";" & _
                      """" & strClause.Replace("""", "'").Replace(";", " ") & """"
      colThis.AddUnique(strSimiLine)
      ' Add the line to the revision output
      If (intValMax > 0) Then
        frmMain.tbGenRevision.AppendText(strSimiLine & vbCrLf)
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/AddSimiLine error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphLemmatizeByDict
  ' Goal:   Using the XML dictionary (Bosworth & Toller derived, or the ME Gutenberg)
  '           attempt to determine the lemma for main entries, such as ADJ, VB, N, N^N, NR, NR^N and so on
  ' History:
  ' 08-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphLemmatizeByDict() As Boolean
    Dim dtrFound() As DataRow     ' Table of remainder
    'Dim dtrCat() As DataRow       ' Category
    Dim strVern As String         ' Current vern
    Dim strVernRew As String = "" ' Vernacular rewritten with other prefix (if at all possible)
    Dim strPos As String          ' Current POS
    Dim strPosDict As String      ' POS according to dictionary
    Dim strPosNear As String      ' Near POS
    Dim strLemma As String        ' Lemma
    'Dim strFeat As String         ' Features
    Dim strClause As String       ' Clause
    Dim strLoc As String          ' Location
    Dim strLev As String          ' Level within table [VernPos]
    Dim strFile As String         ' Filename
    Dim strSimFile As String      ' File name for similarities
    Dim intForestId As Integer    ' Forest Id
    Dim intEtreeId As Integer     ' Etree Id
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage
    Dim intNum As Integer = 0     ' Number of changes
    Dim colPfxRew As New StringColl ' Possible prefix rewrites

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      If (tdlRemain Is Nothing) Then Return False
      ' Initialise
      mrp_colHand.Clear() : mrp_colSimi.Clear() : mrp_colMult.Clear()
      ' Get to the [Remain] table 
      dtrFound = tdlRemain.Tables("Remain").Select("", "Pos ASC, Vern ASC, Freq DESC")
      ' =========== DEBUG: alleen infinitives ======================
      ' dtrFound = tdlMorphDict.Tables("Remain").Select("Pos = 'VB'", "Pos ASC, Vern ASC, Freq DESC")
      ' ===========================================================
      ' Walk them all
      For intI = 0 To dtrFound.Length - 1
        If (bInterrupt) Then Return False
        ' Access the current [Remain] entry
        With dtrFound(intI)
          strVern = .Item("Vern") : strPos = .Item("Pos") : strLemma = ""
          strFile = .Item("File") : intForestId = .Item("forestId") : intEtreeId = .Item("eTreeId")
          strClause = .Item("Clause").ToString
          strLoc = strFile & ":" & intForestId
        End With
        ' Where are we?
        intPtc = (100 * (intI + 1) \ dtrFound.Length)
        Status("Morph lemmatize by dictionary " & intPtc & "% found=[" & intNum & "] POS=" & strPos & " w=" & strVern, intPtc)
        '' ========= DEBUGGING: only verbs =========
        'If (DoLike(strPos, "*VB*|*VA*")) Then
        '  ' =========================================
        ' =========== DEBUGGING =============
        'If (InStr(strVern, "forseowonlicne") > 0) Then
        '  Stop
        ' ===================================
        ' Initialisations
        strPosNear = ""
        ' Action depends on the POS we have
        strPosDict = MorphGetPosDict(strPos, strPosNear)

        'Select Case strPos
        '  Case "VB", "NEG+VB", "RP+VB"
        '    strPosDict = "Vb"
        '  Case "N", "N^N"
        '    strPosDict = "N"
        '  Case "NR", "NR^N"
        '    strPosDict = "NR"
        '  Case "ADJ"
        '    strPosDict = "Adj"
        '  Case "ADV"
        '    strPosDict = "Adv"
        '  Case Else
        '    ' Get category by looking at the HEAD category in table [Cat]
        '    dtrCat = tdlMorphDict.Tables("Cat").Select("Pos='" & strPos & "'")
        '    If (dtrCat.Length > 0) Then
        '      ' Get the actual NEAR pos
        '      strPosNear = dtrCat(0).Item("Near").ToString
        '      ' Derive the [PosDict], if possible...
        '      If (DoLike(dtrCat(0).Item("Head").ToString, "*BE|*HV|*VB|*MD|*AX")) Then
        '        strPosDict = "Vb"
        '      ElseIf (DoLike(dtrCat(0).Item("Head").ToString, "ADJ*")) Then
        '        strPosDict = "Adj"
        '      ElseIf (DoLike(dtrCat(0).Item("Head").ToString, "ADV*")) Then
        '        strPosDict = "Adv"
        '      ElseIf (dtrCat(0).Item("Head").ToString = "N") Then
        '        strPosDict = "N"
        '      ElseIf (dtrCat(0).Item("Head").ToString = "NR") Then
        '        strPosDict = "NR"
        '      Else
        '        strPosDict = ""
        '      End If
        '    ElseIf (strPos Like "ADJ*") Then
        '      strPosDict = "Adj"
        '    ElseIf (strPos Like "ADV*") Then
        '      strPosDict = "Adv"
        '    Else
        '      ' Don't know the POS
        '      strPosDict = ""
        '    End If
        'End Select
        ' Do we have a dictionary POS?
        If (strPosDict <> "") Then
          ' Initialise
          strLev = "1"
          ' =========== DEBUGGING =============
          'If (InStr(strVern, "hihtes") > 0) Then
          '  Stop
          'End If
          ' ===================================
          ' Try to find a dictionary entry
          If (GetOEdictEntry(strVern, strPosDict, strPos, strPosNear, strLemma, strLev, strLoc, strClause)) Then
            ' =========== DEBUG ==============
            ' If (DoLike(strLemma, "acsigan")) Then Stop
            ' ================================
            ' Found a lemma for me: put it in the [VernPos] table
            If (Not VernPosAdd(strPos, strVern, "LemmaOnly", strLemma, strLev, "")) Then Logging("VernPosAdd problem for [" & strVern & "]") : Return False
            ' Keep track of number of instances
            intNum += 1
            frmMain.tbGenSource.AppendText(strVern & " (" & strPos & ") l=" & strLemma & vbCrLf)
          ElseIf (strPos = "ADV") Then
            ' See if there is an entry under "Adj"
            If (GetOEdictEntry(strVern, "Adj", "ADJ", "", strLemma, strLev, strLoc, strClause)) Then
              ' =========== DEBUG ==============
              ' If (DoLike(strLemma, "acsigan")) Then Stop
              ' ================================
              ' Found a lemma for me: put it in the [VernPos] table
              If (Not VernPosAdd(strPos, strVern, "LemmaOnly", strLemma, strLev, "")) Then Logging("VernPosAdd problem for [" & strVern & "]") : Return False
              ' Keep track of number of instances
              intNum += 1
              frmMain.tbGenSource.AppendText(strVern & " (" & strPos & ") l=" & strLemma & vbCrLf)
            End If
          ElseIf (strPosDict = "Vb") AndAlso (Left(strVern, 2) = "ge") AndAlso (strVern.Length > 6) Then
            ' Try to find a lemma for the part following on [ge]
            If (GetOEdictEntry(Mid(strVern, 3), strPosDict, strPos, strPosNear, strLemma, strLev, strLoc, strClause)) Then
              ' =========== DEBUG ==============
              ' If (DoLike(strLemma, "acsigan")) Then Stop
              ' ================================
              ' Check if there is an entry for [ge] + lemma after all
              If (tdlOEdict.Tables("Entry").Select("l='ge" & strLemma & "'").Length = 0) Then
                ' There is NO such entry: put it in the [VernPos] table
                If (Not VernPosAdd(strPos, strVern, "LemmaOnly", strLemma, strLev, "")) Then Logging("VernPosAdd problem for [" & strVern & "]") : Return False
                frmMain.tbGenSource.AppendText(strVern & " (" & strPos & ") l=" & strLemma & vbCrLf)
              Else
                ' There IS such an entry: put it in the [VernPos] table
                If (Not VernPosAdd(strPos, strVern, "LemmaOnly", "ge" & strLemma, strLev, "")) Then Logging("VernPosAdd problem for [" & strVern & "]") : Return False
                frmMain.tbGenSource.AppendText(strVern & " (" & strPos & ") l=ge" & strLemma & vbCrLf)
              End If
              ' Keep track of number of instances
              intNum += 1
            End If
          End If
        Else
          ' the POS does not belong to the major categories, but perhaps it is a verb form we may be able to process?
          If (DoLike(strPos, "*VBD*|*BED*|*HVD*|*AXD*|*MDD*")) Then
            strPosDict = "past|pastSg|pastPl"
          ElseIf (DoLike(strPos, "*VBP*|*BEP*|*HVP*|*AXP*|*MDP*|MD")) Then
            strPosDict = "3rd pres"
          ElseIf (DoLike(strPos, "*VAG*|*VBN*|*BAG|*BEN|*HAG|*HVN|*AXG|*AXN")) Then
            strPosDict = "ptp|ptpIs"
          Else
            If (DoLike(strPos, "V*|B*|H*|MD*|AX*")) AndAlso Not (DoLike(strPos, "VB^*|VBI")) Then
              Stop
            End If
            strPosDict = ""
          End If
          ' Initialise level
          strLev = "1"
          ' Do we have a dictionary POS?
          If (strPosDict <> "") Then
            ' Try to find a dictionary entry
            If (GetOEdictLexfun(strVern, strPosDict, strPos, strLemma)) Then
              ' =========== DEBUG ==============
              ' If (DoLike(strLemma, "acsigan")) Then Stop
              ' ================================
              ' Found a lemma for me: put it in the [VernPos] table
              If (Not VernPosAdd(strPos, strVern, "LemmaOnly", strLemma, strLev, "")) Then Logging("VernPosAdd problem for [" & strVern & "]") : Return False
              ' Keep track of number of instances
              intNum += 1
              frmMain.tbGenSource.AppendText(strVern & " (" & strPos & ") l=" & strLemma & vbCrLf)
            Else
              ' Try the best prefix rewrite
              ' Try to find a prefix rewrite
              If (GetPrefixRewrite(strVern, colPfxRew)) Then
                ' Check them all
                For intJ = 0 To colPfxRew.Count - 1
                  strVernRew = colPfxRew.Item(intJ)
                  If (GetOEdictLexfun(strVernRew, strPosDict, strPos, strLemma)) Then
                    ' =========== DEBUG ==============
                    ' If (DoLike(strLemma, "acsigan")) Then Stop
                    ' ================================
                    ' Found a lemma for me: put it in the [VernPos] table
                    If (Not VernPosAdd(strPos, strVern, "LemmaOnly", strLemma, strLev, "")) Then Logging("VernPosAdd problem for [" & strVern & "]") : Return False
                    ' Keep track of number of instances
                    intNum += 1
                    frmMain.tbGenSource.AppendText(strVern & " (" & strPos & ") l=" & strLemma & vbCrLf)
                  End If
                Next intJ
              End If
            End If
          End If
        End If
        '' ========= DEBUGGING: only verbs =========
        '' =========================================
        'End If
        'Else
        '  ' SKIP this one
        'End If
      Next intI
      ' Save the results of simi
      strSimFile = GetDocDir() & "\MorphDictSimi_" & Today.Day & MonthName(Month(Today)) & "-auto.csv"
      IO.File.WriteAllText(strSimFile, mrp_colSimi.Text)
      ' Inform user of these results
      Logging("Look for suggestions in: " & strSimFile)
      ' Save the results of hand
      strSimFile = GetDocDir() & "\MorphDictHand_" & Today.Day & MonthName(Month(Today)) & "-auto.csv"
      IO.File.WriteAllText(strSimFile, mrp_colHand.Text)
      ' Save the results of multi
      strSimFile = GetDocDir() & "\MorphDictMult_" & Today.Day & MonthName(Month(Today)) & "-auto.csv"
      IO.File.WriteAllText(strSimFile, mrp_colMult.Text)
      ' Show what we have done
      Logging("MorphLemmatizeByDict found [" & intNum & "] matches")
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphLemmatizeByDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphLemmatizeRewrite
  ' Goal:   Using the suffix-rewrite rules in table [Suffix],
  '           attempt to determine the lemma for words that have not yet been lemmatized
  '         Action depends on the number of different lemma's that are possible
  '         zero - leave as is
  '         one  - accept this as lemma
  '         more - add a list of lemma's, possibly extended with confidence?
  ' History:
  ' 22-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphLemmatizeRewrite() As Boolean
    Dim dtrFound() As DataRow     ' Table of remainder
    Dim dtrRewri() As DataRow     ' Suffix rewrite rules for one particular POS
    Dim strVern As String         ' Current vern
    Dim strVernAlt As String      ' Alternative form of vernacular
    Dim strPos As String          ' Current POS
    Dim strLastPos As String = "" ' Previous POS
    Dim strLemma As String        ' Lemma
    Dim strLev As String = "1"    ' Level
    ' Dim strFeat As String         ' Features
    Dim strType As String         ' What's our addition type?
    Dim strSuffix As String       ' Suffix we are looking for
    Dim strRewr As String         ' Rewrite belonging to it
    Dim colLem As New StringColl  ' Collection of possible lemma's
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage
    Dim intSfx As Integer         ' Length of suffix
    Dim intNum As Integer = 0     ' Number of changes

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Get to the [Remain] table
      dtrFound = tdlMorphDict.Tables("Remain").Select("", "Pos ASC, Vern ASC, Freq DESC")
      dtrRewri = Nothing
      ' Walk them all
      For intI = 0 To dtrFound.Length - 1
        If (bInterrupt) Then Return False
        ' Where are we?
        intPtc = (100 * (intI + 1) \ dtrFound.Length)
        Status("Morph lemmatize rewrite " & intPtc & "%", intPtc)
        ' Get the POS
        strPos = dtrFound(intI).Item("Pos")
        ' Check if this differs from the previous POS
        If (strPos <> strLastPos) Then
          ' Adapt the last POS flag
          strLastPos = strPos
          ' rewrite rules may be broader for ADJ* and ADV*
          If (strPos Like "AD[JV]*") Then
            ' Use adapted POS search
            dtrRewri = tdlMorphDict.Tables("Suffix").Select("Pos LIKE '" & Left(strPos, 3) & "*'", "Freq DESC")
          ElseIf (strPos Like "WAD[JV]*") Then
            ' Use adapted POS search
            dtrRewri = tdlMorphDict.Tables("Suffix").Select("Pos LIKE '" & Left(strPos, 4) & "*'", "Freq DESC")
          Else
            ' Find all possible suffix-rewrite rules for this POS
            dtrRewri = tdlMorphDict.Tables("Suffix").Select("Pos='" & strPos & "'", "Freq DESC")
          End If
        End If
        ' Get the word itself
        strVern = dtrFound(intI).Item("Vern") : colLem.Clear()
        ' Try to apply different rewrite rules
        For intJ = 0 To dtrRewri.Length - 1
          ' Get this suffix
          strSuffix = dtrRewri(intJ).Item("Sfx")
          intSfx = strSuffix.Length
          ' Try to match the suffix
          If (Right(strVern, intSfx) = strSuffix) Then
            ' Apply the rewrite
            strRewr = Left(strVern, strVern.Length - intSfx) & dtrRewri(intJ).Item("Rew")
            ' See if this re-written word figures as lemma in the [Lemma] table
            If (LemmaMatch(strRewr, strPos)) Then
              ' Add to the lemma collection
              colLem.Add(strRewr, "|sfx=" & dtrRewri(intJ).Item("Sfx") & "|frq=" & dtrRewri(intJ).Item("Freq"))
            End If
          End If
        Next intJ
        ' Any results?
        If (colLem.Count = 0) Then
          ' No match: does the word contain a "y"?
          If (InStr(strVern, "y") > 0) Then
            ' No match: try to replace "y" for "i"
            strVernAlt = strVern.Replace("y", "i")
            For intJ = 0 To dtrRewri.Length - 1
              ' Get this suffix
              strSuffix = dtrRewri(intJ).Item("Sfx")
              intSfx = strSuffix.Length
              ' Try to match the suffix
              If (Right(strVernAlt, intSfx) = strSuffix) Then
                ' Apply the rewrite
                strRewr = Left(strVernAlt, strVernAlt.Length - intSfx) & dtrRewri(intJ).Item("Rew")
                ' See if this re-written word figures as lemma in the [Lemma] table
                If (LemmaMatch(strRewr, strPos)) Then
                  ' Add to the lemma collection
                  colLem.Add(strRewr, "|sfx=" & dtrRewri(intJ).Item("Sfx") & "|frq=" & dtrRewri(intJ).Item("Freq"))
                End If
              End If
            Next intJ
            ' Still no solution?
            If (colLem.Count = 0) Then
              strVernAlt = strVern.Replace("y", "ie")
              For intJ = 0 To dtrRewri.Length - 1
                ' Get this suffix
                strSuffix = dtrRewri(intJ).Item("Sfx")
                intSfx = strSuffix.Length
                ' Try to match the suffix
                If (Right(strVernAlt, intSfx) = strSuffix) Then
                  ' Apply the rewrite
                  strRewr = Left(strVernAlt, strVernAlt.Length - intSfx) & dtrRewri(intJ).Item("Rew")
                  ' See if this re-written word figures as lemma in the [Lemma] table
                  If (LemmaMatch(strRewr, strPos)) Then
                    ' Add to the lemma collection
                    colLem.Add(strRewr, "|sfx=" & dtrRewri(intJ).Item("Sfx") & "|frq=" & dtrRewri(intJ).Item("Freq"))
                  End If
                End If
              Next intJ
            End If
          End If
        End If
        ' Action depends on the amount of matches
        Select Case colLem.Count
          Case 0  ' Don't do anything at this point
          Case 1
            ' There is one unique match: add it
            ' Add this information into table [VernPos] to be processed later on
            strType = "LemmaOnly"
            If (Not VernPosAdd(strPos, strVern, strType, colLem.Item(0), strLev, "")) Then Logging("MorphLemmatizeRewrite: problem in VernPosAdd") : Return False
            ' Keep track of numbers
            intNum += 1
          Case Else
            ' There's more than one match
            'Stop
            ' We need to add all possible lemma's that were found
            strLemma = colLem.Count
            For intJ = 0 To colLem.Count - 1
              strLemma &= ";" & colLem.Item(intJ) & colLem.Exmp(intJ)
            Next intJ
            ' Add this information into table [VernPos] to be processed later on
            strType = "LemmaAmbi"
            If (Not VernPosAdd(strPos, strVern, strType, strLemma, strLev, "")) Then Logging("MorphLemmatizeRewrite: problem in VernPosAdd") : Return False
            ' Keep track of numbers
            intNum += 1
        End Select
      Next intI
      ' Show what we have done
      Logging("MorphLemmatizeRewrite found [" & intNum & "] matches")
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphLemmatizeRewrite error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphRewrites
  ' Goal:   Get all possible prefix + suffix rewrites for me
  ' History:
  ' 09-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphRewrites(ByVal strVern As String, ByVal strPos As String, ByRef colRew As StringColl) As Boolean
    Static strLastPos As String   ' last used POS
    Static dtrRewri() As DataRow  ' Suffix rewrite rules for one particular POS
    Static dtrPref() As DataRow   ' Prefix replacements (irrespective of POS)
    Dim strSuffix As String       ' Suffix we address
    Dim strPrefix As String       ' Prefix we address
    Dim strRewr As String         ' Rewrite we use
    Dim strRewr2 As String        ' Second rewrite
    Dim intSfx As Integer         ' Length of suffix
    Dim intPfx As Integer         ' Length of prefix
    Dim intJ As Integer           ' Counter
    Dim intI As Integer           ' Counter

    Try
      ' Validate
      If (strVern = "") OrElse (strPos = "") Then Return ""
      ' Initialize
      If (strLastPos = "") Then dtrRewri = Nothing : dtrPref = Nothing
      colRew.Clear()
      ' Possibly initialise prefixes
      If (dtrPref Is Nothing) Then
        ' Validat
        If (tdlMorphDict.Tables("Prefix") Is Nothing) Then
          dtrPref = Nothing
        Else
          ' Get prefixes - in decreasing size order
          dtrPref = tdlMorphDict.Tables("Prefix").Select("", "Size DESC")
          ' dtrPref = tdlOEdict.Tables("Entry").Select("Cf <> ''")
        End If
      End If
      ' Where are we?
      If (strPos <> strLastPos) Then
        ' Assign lastpos
        strLastPos = strPos
        ' rewrite rules may be broader for ADJ* and ADV*
        If (strPos Like "AD[JV]*") Then
          ' Use adapted POS search
          dtrRewri = tdlMorphDict.Tables("Suffix").Select("Pos LIKE '" & Left(strPos, 3) & "*'", "Freq DESC")
        ElseIf (strPos Like "WAD[JV]*") Then
          ' Use adapted POS search
          dtrRewri = tdlMorphDict.Tables("Suffix").Select("Pos LIKE '" & Left(strPos, 4) & "*'", "Freq DESC")
        Else
          ' Find all possible suffix-rewrite rules for this POS
          dtrRewri = tdlMorphDict.Tables("Suffix").Select("Pos='" & strPos & "'", "Freq DESC")
        End If
      End If
      ' Try to apply different rewrite rules
      For intJ = 0 To dtrRewri.Length - 1
        ' Get this suffix
        strSuffix = dtrRewri(intJ).Item("Sfx")
        intSfx = strSuffix.Length
        ' Try to match the suffix
        If (Right(strVern, intSfx) = strSuffix) Then
          ' Apply the rewrite
          strRewr = Left(strVern, strVern.Length - intSfx) & dtrRewri(intJ).Item("Rew")
          ' Add this rewrite
          colRew.AddUnique(strRewr)
          ' Possibly add dh <> th rewrites
          If (InStr(strRewr, "dh") > 0) Then
            colRew.AddUnique(strRewr.Replace("dh", "th"))
          ElseIf (InStr(strRewr, "th") > 0) Then
            colRew.AddUnique(strRewr.Replace("th", "dh"))
          ElseIf (InStr(strRewr, "ui") > 0) Then
            colRew.AddUnique(strRewr.Replace("ui", "wi"))
            colRew.AddUnique(strRewr.Replace("ui", "fi"))
          ElseIf (InStr(strRewr, "gg") > 0) Then
            colRew.AddUnique(strRewr.Replace("gg", "cg"))
          ElseIf (InStr(strRewr, "k") > 0) Then
            colRew.AddUnique(strRewr.Replace("k", "c"))
          End If
          ' Try possible prefix rewrites
          For intI = 0 To dtrPref.Length - 1
            ' Check if this prefix matches
            ' strPrefix = dtrPref(intI).Item("l")
            strPrefix = dtrPref(intI).Item("Pre")
            intPfx = strPrefix.Length
            If (Left(strRewr, intPfx) = strPrefix) Then
              ' Replace
              ' strRewr2 = dtrPref(intI).Item("Cf") & Mid(strRewr, intPfx + 1)
              strRewr2 = dtrPref(intI).Item("Rew") & Mid(strRewr, intPfx + 1)
              ' Add to rewrite collection
              colRew.AddUnique(strRewr2)
              ' Possibly add dh <> th rewrites
              If (InStr(strRewr2, "dh") > 0) Then
                colRew.AddUnique(strRewr2.Replace("dh", "th"))
              ElseIf (InStr(strRewr2, "th") > 0) Then
                colRew.AddUnique(strRewr2.Replace("th", "dh"))
              ElseIf (InStr(strRewr2, "ui") > 0) Then
                colRew.AddUnique(strRewr2.Replace("ui", "wi"))
                colRew.AddUnique(strRewr2.Replace("ui", "fi"))
              ElseIf (InStr(strRewr2, "gg") > 0) Then
                colRew.AddUnique(strRewr2.Replace("gg", "cg"))
              ElseIf (InStr(strRewr2, "k") > 0) Then
                colRew.AddUnique(strRewr2.Replace("k", "c"))
              End If
            End If
          Next intI
        End If
      Next intJ
      ' Add possible prefix rewrites for the unchanged form
      strRewr = strVern
      ' Possibly add dh <> th rewrites
      If (InStr(strRewr, "dh") > 0) Then
        colRew.AddUnique(strRewr.Replace("dh", "th"))
      ElseIf (InStr(strRewr, "th") > 0) Then
        colRew.AddUnique(strRewr.Replace("th", "dh"))
      ElseIf (InStr(strRewr, "ui") > 0) Then
        colRew.AddUnique(strRewr.Replace("ui", "wi"))
        colRew.AddUnique(strRewr.Replace("ui", "fi"))
      ElseIf (InStr(strRewr, "gg") > 0) Then
        colRew.AddUnique(strRewr.Replace("gg", "cg"))
      ElseIf (InStr(strRewr, "k") > 0) Then
        colRew.AddUnique(strRewr.Replace("k", "c"))
      End If
      ' Try possible prefix rewrites
      For intI = 0 To dtrPref.Length - 1
        ' Check if this prefix matches
        ' strPrefix = dtrPref(intI).Item("l")
        strPrefix = dtrPref(intI).Item("Pre")
        intPfx = strPrefix.Length
        If (Left(strRewr, intPfx) = strPrefix) Then
          ' Replace
          ' strRewr2 = dtrPref(intI).Item("Cf") & Mid(strRewr, intPfx + 1)
          strRewr2 = dtrPref(intI).Item("Rew") & Mid(strRewr, intPfx + 1)
          ' Add to rewrite collection
          colRew.AddUnique(strRewr2)
          ' Possibly add dh <> th rewrites
          If (InStr(strRewr2, "dh") > 0) Then
            colRew.AddUnique(strRewr2.Replace("dh", "th"))
          ElseIf (InStr(strRewr2, "th") > 0) Then
            colRew.AddUnique(strRewr2.Replace("th", "dh"))
          ElseIf (InStr(strRewr2, "ui") > 0) Then
            colRew.AddUnique(strRewr2.Replace("ui", "wi"))
            colRew.AddUnique(strRewr2.Replace("ui", "fi"))
          ElseIf (InStr(strRewr2, "gg") > 0) Then
            colRew.AddUnique(strRewr2.Replace("gg", "cg"))
          ElseIf (InStr(strRewr2, "k") > 0) Then
            colRew.AddUnique(strRewr2.Replace("k", "c"))
          End If
        End If
      Next intI
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphRewrites error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SfxRewMatch
  ' Goal:   Check if vernacular word [strVern] with part-of-speech [strPos] matches
  '           with [strWord] using any of the suffix-rewrite rules in table [Suffix]
  ' History:
  ' 22-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function SfxRewMatch(ByVal strVern As String, ByVal strWord As String, ByVal strPos As String) As Boolean
    Static strLastPos As String   ' last used POS
    Static dtrRewri() As DataRow  ' Suffix rewrite rules for one particular POS
    Dim strSuffix As String       ' Suffix we address
    Dim strRewr As String         ' Rewrite we use
    Dim intSfx As Integer         ' Length of suffix
    Dim intJ As Integer           ' Counter

    Try
      ' Initialize
      If (strLastPos = "") Then dtrRewri = Nothing
      ' Do not proceed if the first TWO letters do not match
      ' NEE!!! If (Left(strVern, 2) <> Left(strWord, 2)) Then Return False
      ' Stel dat het woord met th begint en er een dh variant is...
      ' Where are we?
      If (strPos <> strLastPos) Then
        ' Assign lastpos
        strLastPos = strPos
        ' rewrite rules may be broader for ADJ* and ADV*
        If (strPos Like "AD[JV]*") Then
          ' Use adapted POS search
          dtrRewri = tdlMorphDict.Tables("Suffix").Select("Pos LIKE '" & Left(strPos, 3) & "*'", "Freq DESC")
        ElseIf (strPos Like "WAD[JV]*") Then
          ' Use adapted POS search
          dtrRewri = tdlMorphDict.Tables("Suffix").Select("Pos LIKE '" & Left(strPos, 4) & "*'", "Freq DESC")
        Else
          ' Find all possible suffix-rewrite rules for this POS
          dtrRewri = tdlMorphDict.Tables("Suffix").Select("Pos='" & strPos & "'", "Freq DESC")
        End If
      End If
      ' Try to apply different rewrite rules
      For intJ = 0 To dtrRewri.Length - 1
        ' Get this suffix
        strSuffix = dtrRewri(intJ).Item("Sfx")
        intSfx = strSuffix.Length
        ' Try to match the suffix
        If (Right(strVern, intSfx) = strSuffix) Then
          ' Apply the rewrite
          strRewr = Left(strVern, strVern.Length - intSfx) & dtrRewri(intJ).Item("Rew")
          ' See if the words now match
          If (strRewr = strWord) Then
            Return True
          End If
          ' Compare in several different spellingvariants
          If (OEtaggedWordsLike(strRewr, strWord)) Then
            ' This should be the correct one
            Return True
          End If
        End If
      Next intJ
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/SfxRewMatch error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DeriveRewMatch
  ' Goal:   Check if vernacular word [strVern] with part-of-speech [strPos] matches
  '           with [strWord] using any of the suffix-rewrite rules in table [Suffix]
  '         The word we look at comes with category [strPos], and we first check if
  '           this category is a derived category or not
  ' History:
  ' 13-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function DeriveRewMatch(ByVal strVern As String, ByVal strWord As String, ByVal strPos As String, _
                                  Optional ByVal bUseNear As Boolean = False) As Boolean
    Static strLastPos As String   ' last used POS
    Static bLastUseNear As Boolean ' Last used [bUseNear]
    Static dtrRewri() As DataRow  ' Suffix rewrite rules for one particular POS
    Dim strSuffix As String       ' Suffix we address
    Dim strRewr As String         ' Rewrite we use
    Dim intSfx As Integer         ' Length of suffix
    Dim intJ As Integer           ' Counter

    Try
      ' Initialize
      If (strLastPos = "") Then dtrRewri = Nothing
      ' Is this a derive POS?
      If (tdlMorphDict.Tables("Cat").Select("Type='Derived' AND Pos='" & strPos & "'").Length = 0) Then Return False
      ' Do not proceed if the first TWO letters do not match
      ' NEE!!! If (Left(strVern, 2) <> Left(strWord, 2)) Then Return False
      ' Stel dat het woord met th begint en er een dh variant is...
      ' Where are we?
      If (strPos <> strLastPos) OrElse (bUseNear <> bLastUseNear) Then
        ' Assign lastpos
        strLastPos = strPos : bLastUseNear = bUseNear
        ' Use near or not?
        If (bUseNear) Then
          ' Find all possible suffix-rewrite rules for this particular POS
          dtrRewri = tdlMorphDict.Tables("Derive").Select("Type='Deriv' AND Pos='" & strPos & "'", "Freq DESC")
        Else
          ' Find all possible suffix-rewrite rules for this particular POS
          dtrRewri = tdlMorphDict.Tables("Derive").Select("Type='Main' AND Pos='" & strPos & "'", "Freq DESC")
        End If
      End If
      ' Try to apply different rewrite rules
      For intJ = 0 To dtrRewri.Length - 1
        ' Get this suffix
        strSuffix = dtrRewri(intJ).Item("Sfx")
        intSfx = strSuffix.Length
        ' Try to match the suffix
        If (Right(strVern, intSfx) = strSuffix) Then
          ' Apply the rewrite
          strRewr = Left(strVern, strVern.Length - intSfx) & dtrRewri(intJ).Item("Rew")
          ' ========= DEBUG ===========
          If (bUseNear) AndAlso (Left(strRewr, 5) = Left(strWord, 5)) Then Stop
          ' ===========================
          ' See if the words now match
          If (strRewr = strWord) Then
            Return True
          End If
          ' Compare in several different spellingvariants
          If (OEtaggedWordsLike(strRewr, strWord)) Then
            ' This should be the correct one
            Return True
          End If
        End If
      Next intJ
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/DeriveRewMatch error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphSweep
  ' Goal:   Address the entries from table [Remain] using information from [Morph]
  '           and then adding entries into table [VernPos]
  ' History:
  ' 14-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphSweep() As Boolean
    Dim dtrFound() As DataRow ' Table of remainder
    Dim strVern As String     ' Current vern
    Dim strPos As String      ' Current POS
    Dim strLemma As String    ' Lemma
    Dim strLev As String = "1"  ' Level
    Dim strFeat As String     ' Features
    Dim strType As String     ' What's our addition type?
    Dim intI As Integer       ' Counter
    Dim intJ As Integer       ' Counter
    Dim intPtc As Integer     ' Percentage
    Dim intNum As Integer = 0 ' Number of changes

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Get to the [Remain] table
      dtrFound = tdlMorphDict.Tables("Remain").Select("", "Pos ASC, Vern ASC, Freq DESC")
      ' Walk them all
      For intI = 0 To dtrFound.Length - 1
        If (bInterrupt) Then Return False
        ' Where are we?
        intPtc = (100 * (intI + 1) \ dtrFound.Length)
        Status("Morph sweep " & intPtc & "%", intPtc)
        ' Get the vern and pos
        strVern = dtrFound(intI).Item("Vern") : strPos = dtrFound(intI).Item("Pos")
        ' Check how many have the same Vern+Pos combination (they only differ in @PrntLabel)
        intJ = intI
        While (intJ < dtrFound.Length - 1) AndAlso (dtrFound(intJ + 1).Item("Vern") = strVern) AndAlso _
              (dtrFound(intJ + 1).Item("Pos") = strPos)
          ' We are allowed to proceed
          intJ += 1
        End While
        ' Try to find a lemma + features in table [Morph]
        strLemma = "" : strFeat = ""
        If (MorphLemmaGet(Nothing, strVern, strPos, strLemma, strFeat)) Then
          ' Add this information into table [VernPos], to be processed later on
          strType = IIf(strFeat = "", "LemmaOnly", "LemmaFeat")
          If (Not VernPosAdd(strPos, strVern, "", strLemma, strLev, strFeat)) Then Logging("MorphSweep: problem in VernPosAdd") : Return False
          ' Keep track of numbers
          intNum += 1
        End If
      Next intI
      ' Show what we have done
      Logging("MorphSweep found [" & intNum & "] matches")
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphSweep error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   LemmaMatch
  ' Goal:   Check if [strVern] is a lemma in the table [Lemma]
  '           taking into account the POS in [strPos]
  ' History:
  ' 22-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function LemmaMatch(ByVal strVern As String, ByVal strPos As String) As Boolean
    Dim dtrFound() As DataRow ' Result of select
    Dim intI As Integer       ' Counter

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) OrElse (tdlMorphDict.Tables("Lemma") Is Nothing) Then Return False
      ' Try to match
      dtrFound = tdlMorphDict.Tables("Lemma").Select("Vern='" & strVern & "'")
      '' What if we have some kind of a match?
      'If (dtrFound.Length > 0) AndAlso (strPos Like "N*") Then
      '  Debug.Print("Vern=[" & strVern & "], POS = " & strPos)
      '  Stop
      'End If
      ' Possibly adapt the POS
      If (strPos.Length > 3) AndAlso (strPos Like "AD[JV]*") Then
        strPos = Left(strPos, 3)
      ElseIf (strPos.Length > 4) AndAlso (strPos Like "WAD[JV]*") Then
        strPos = Left(strPos, 4)
      End If
      ' Check the POS
      For intI = 0 To dtrFound.Length - 1
        With dtrFound(intI)
          ' Some POS need special treatment
          If (LabelLemmaOE(strPos) Like .Item("Pos") & "*") Then Return True
        End With
      Next intI
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/LemmaMatch error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphLemmaCheck
  ' Goal:   Check the status of the lemma [strLemma] in table [Morph]
  '         But it should be there *not* under [strPos], but under the part-of-speech that is
  '           the major lemma category of [strPos]
  ' History:
  ' 24-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphLemmaCheck(ByVal strLemma As String, ByVal strVern As String, ByVal strPos As String, ByRef strHead As String, _
                                  ByRef strBack As String, Optional ByVal bUseVern As Boolean = False) As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim dtrCat() As DataRow   ' Result of SELECT
    Dim strSelect As String   ' Select statement

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) OrElse (tdlMorphDict.Tables("Morph") Is Nothing) Then Return False
      If (tdlMorphDict.Tables("Cat") Is Nothing) Then Return False
      ' Find out the real POS
      dtrCat = tdlMorphDict.Tables("Cat").Select("Type = 'Derived' AND Pos = '" & strPos & "'")
      If (dtrCat.Length > 0) Then
        ' Adapt the POS we are looking for
        strPos = dtrCat(0).Item("Head").ToString
      End If
      strHead = strPos
      ' Construct SELECT statement
      strSelect = "l='" & strLemma.Replace("'", "''") & "' AND Pos='" & strPos & "'"
      If (bUseVern) AndAlso (strVern <> "") AndAlso (strVern <> strLemma) Then
        strSelect &= " AND Vern='" & strVern.Replace("'", "''") & "'"
      End If
      ' Try to find the lemma
      dtrFound = tdlMorphDict.Tables("Morph").Select(strSelect)
      If (dtrFound.Length = 0) Then
        ' Double check: does this have a hyphen?
        If (InStr(strLemma, "-") > 0) Then
          If (InStr(strPos, "RP+") = 1) Then
            ' Look for the part *after* the hyphen, and assume that the prefix is okay
            strSelect = "l='" & Mid(strLemma, InStrRev(strLemma, "-") + 1).Replace("'", "''") & "' AND Pos='" & Mid(strPos, 4) & "'"
            If (bUseVern) AndAlso (strVern <> "") AndAlso (strVern <> strLemma) Then
              strSelect &= " AND Vern='" & strVern.Replace("'", "''") & "'"
            End If
          Else
            ' Look for the part *after* the hyphen, and assume that the prefix is okay
            strSelect = "l='" & Mid(strLemma, InStr(strLemma, "-") + 1).Replace("'", "''") & "' AND Pos='" & strPos & "'"
            If (bUseVern) AndAlso (strVern <> "") AndAlso (strVern <> strLemma) Then
              strSelect &= " AND Vern='" & strVern.Replace("'", "''") & "'"
            End If
          End If
          ' Try to find the lemma
          dtrFound = tdlMorphDict.Tables("Morph").Select(strSelect)
          Select Case dtrFound.Length
            Case 0
              ' This lemma does not exist
              strBack = "none"
            Case 1
              ' The lemma is unique
              strBack = "unique"
            Case Else
              ' There is ambiguity
              strBack = "ambi"
          End Select
        Else
          ' This lemma does not exist
          strBack = "none"
        End If
      ElseIf (dtrFound.Length = 1) Then
        ' The lemma is unique
        strBack = "unique"
      Else
        ' There is ambiguity
        strBack = "ambi"
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphLemmaCheck error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphLemmaGet
  ' Goal:   Given a Vern+POS combination, get a common lemma and a common set of features
  ' History:
  ' 14-03-2013  ERK Created
  ' 22-03-2013  ERK Only ask the user if [ndxThis] holds an item
  ' ------------------------------------------------------------------------------------
  Public Function MorphLemmaGet(ByRef ndxThis As XmlNode, ByVal strVern As String, ByVal strPos As String, ByRef strLemma As String, _
                                ByRef strFeat As String) As Boolean
    'Dim dtrFound() As DataRow     ' Result of SELECT
    'Dim arFeat() As String        ' Array of features
    Dim colFeat As New StringColl ' Collection of features
    Dim strPosAlt As String       ' Alternative part-of-speech
    Dim strLemmaAlt As String     ' Alternative lemma
    Dim strVernAlt As String      ' Alternative vernacular word
    Dim strVernPre As String      ' Vernacular with different preverb
    'Dim intPos As Integer         ' Position within string
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    'Dim intFreqT As Integer       ' Sum frequency
    'Dim intFreqI As Integer       ' Incremental Frequency
    Dim intPtc As Integer = 95    ' Accuracy percentage

    Try
      ' Validate
      If (strVern = "") Then strLemma = "" : Return True
      If (strPos = "") Then Return False
      ' =========== DEBUG ============
      ' If (strVern = "middaneard") Then Stop
      ' If (strVern = "east") Then Stop
      ' ==============================
      ' Initialise
      strLemma = ""
      If (Not MorphBuildLists()) Then Return False
      ' Is this a modal?
      If (Left(strPos, 2) = "MD") Then
        ' Determine the lemma by matching
        If (strVern Like "c*") Then
          strLemma = "cunnan"
        ElseIf (strVern Like "d*") Then
          strLemma = "durran"
        ElseIf (strVern Like "m*[ghc]*") Then
          strLemma = "magan"
        ElseIf (strVern Like "m*") Then
          strLemma = "motan"
        ElseIf (strVern Like "sc*") Then
          strLemma = "sculan"
        ElseIf (strVern Like "th*rf*") Then
          strLemma = "thurfan"
        ElseIf (strVern Like "w*l*") Then
          strLemma = "willan"
        Else
          ' We will need to ask for a lemma anyway!
        End If
      ElseIf (strPos = "NR^N") Then
        ' Accept this word as the lemma
        strLemma = strVern : strFeat = "case=nm"
      Else
        ' See if Vern/Pos can be found in the Morph table
        If (Not MorphLemmaFeat(strVern, strPos, strLemma, strFeat, 95)) Then Return False
      End If
      ' Have we gotten the lemma yet?
      If (strLemma = "") Then
        ' Try some more options, depending on the POS we have
        If (strPos Like "ADJ*") Then
          ' Look for this form in accusative, dative or other inflected forms
          If (Not MorphLemmaFeat(strVern, "ADJ^*", strLemma, strFeat, 95)) Then Return False
          If (strLemma <> "") Then MorphRemoveFeat(strFeat, "case")
          ' Did we get it?
          If (strLemma = "") Then
            ' We haven't gotten it, so try stripping off different suffixes
            If (DoLike(strVern, "*ne|*an|*re|*or|*ra")) Then
              ' Look for this form in accusative, dative or other inflected forms
              If (Not MorphLemmaFeat(Left(strVern, strVern.Length - 2), "ADJ*", strLemma, strFeat, 95)) Then Return False
              If (strLemma <> "") Then MorphRemoveFeat(strFeat, "case")
            ElseIf (DoLike(strVern, "*e|*a")) Then
              ' Look for this form in accusative, dative or other inflected forms
              If (Not MorphLemmaFeat(Left(strVern, strVern.Length - 1), "ADJ*", strLemma, strFeat, 95)) Then Return False
              If (strLemma <> "") Then MorphRemoveFeat(strFeat, "case")
            End If
          End If
        ElseIf (strPos = "FP") Then
          If (strVern Like "b*") Then
            strLemma = "butan"
          Else
            strLemma = "furthum"
          End If
        ElseIf (strPos = "C") Then
          If (DoLike(strVern, "[dt][eiuy]|[dt]hae")) Then
            strLemma = "the"
          ElseIf (DoLike(strVern, "[dt]*[ae]*[dt]*h")) Then
            strLemma = "thaet"
          ElseIf (DoLike(strVern, "[dt]*[ae]*[dt]*e")) Then
            strLemma = "thaette"
          End If
        ElseIf (strPos Like "ADV^*") Then
          ' Look for any adverbs
          If (Not MorphLemmaFeat(strVern, "ADV*", strLemma, strFeat, 95)) Then Return False
          ' Did we get it?
          If (strLemma = "") Then
            ' We haven't gotten it, so try stripping off different suffixes
            If (DoLike(strVern, "*[oeu]st")) Then
              ' Look for this form in accusative, dative or other inflected forms
              If (Not MorphLemmaFeat(Left(strVern, strVern.Length - 3), "ADV*", strLemma, strFeat, 95)) Then Return False
            ElseIf (DoLike(strVern, "*[oeu]s[dt]h")) Then
              ' Look for this form in accusative, dative or other inflected forms
              If (Not MorphLemmaFeat(Left(strVern, strVern.Length - 4), "ADV*", strLemma, strFeat, 95)) Then Return False
            ElseIf (DoLike(strVern, "*s[dt]")) Then
              ' Look for this form in accusative, dative or other inflected forms
              If (Not MorphLemmaFeat(Left(strVern, strVern.Length - 2), "ADV*", strLemma, strFeat, 95)) Then Return False
            End If
          End If
        ElseIf (strPos = "N") Then
          ' Look for any nouns
          If (Not MorphLemmaFeat(strVern, "N^*", strLemma, strFeat, 95)) Then Return False
        ElseIf (strPos Like "N^*") Then
          ' Look for any nouns
          If (Not MorphLemmaFeat(strVern, "N", strLemma, strFeat, 95)) Then Return False
        ElseIf (strPos = "NR") Then
          ' Look for any nouns
          If (Not MorphLemmaFeat(strVern, "NR^*", strLemma, strFeat, 95)) Then Return False
        ElseIf (strPos Like "NR^*") Then
          ' Look for any nouns
          If (Not MorphLemmaFeat(strVern, "NR", strLemma, strFeat, 95)) Then Return False
        ElseIf (ndxThis IsNot Nothing) AndAlso (strPos Like "B*") Then
          ' This must be a form of "be", so ask for possible values
          If (Not MorphLemmaAsk(ndxThis, strVern, strPos, "BE", strLemma)) Then Return False
        ElseIf (ndxThis IsNot Nothing) AndAlso (strPos Like "AX*") Then
          ' This is a verb, so offer the lexemes that are available for auxilary verbs
          If (Not MorphLemmaAsk(ndxThis, strVern, strPos, "AX", strLemma)) Then Return False
        ElseIf (ndxThis IsNot Nothing) AndAlso (strPos Like "V*") Then
          ' This is a verb, so offer the lexemes that are available for verbs in general
          If (Not MorphLemmaAsk(ndxThis, strVern, strPos, "VB", strLemma)) Then Return False
        ElseIf (strPos Like "HV*") OrElse (strPos Like "HA*") Then
          ' Must be "to have"
          strLemma = "habban"
        ElseIf (strPos Like "RP+*") Then
          ' Temporarily strip off the RP+
          strPosAlt = Mid(strPos, 4)
          ' It could be that this form is known with the revised POS
          If (Not MorphLemmaFeat(strVern, strPosAlt, strLemma, strFeat, 95)) Then Return False
          ' Any results?
          If (strLemma = "") Then
            ' Check what the preverb is
            For intI = 0 To arPreVerb.Length - 1
              ' Do we have this preverb?
              If (Left(strVern, arPreVerb(intI).Length) = arPreVerb(intI)) Then
                ' Make an alternative lemma
                strVernAlt = Mid(strVern, arPreVerb(intI).Length + 1) : strLemmaAlt = ""
                ' This is a verb, so offer the lexemes that are available for verbs in general
                If (Not MorphLemmaFeat(strVernAlt, strPosAlt, strLemmaAlt, strFeat, 95)) Then Return False
                ' Do we have a result?
                If (strLemmaAlt = "") Then
                  ' Alternative strategy: look for a form with the same POS, but with a preverb
                  For intJ = 0 To arPreVerb.Length - 1
                    ' Skip equal preverb
                    If (intJ <> intI) Then
                      ' Get an adapted vernacular
                      strVernPre = arPreVerb(intJ) & strVernAlt : strLemmaAlt = ""
                      ' Offer the combination [VernPre] / [PosAlt]
                      If (Not MorphLemmaFeat(strVernPre, strPosAlt, strLemmaAlt, strFeat, 95)) Then Return False
                      ' Alternative POS
                      If (strLemmaAlt = "") Then
                        If (Not MorphLemmaFeat(strVernPre, strPos, strLemmaAlt, strFeat, 95)) Then Return False
                      End If
                      ' Any results?
                      If (strLemmaAlt <> "") Then
                        strLemmaAlt = Mid(strLemmaAlt, arPreVerb(intJ).Length + 1)
                        Exit For
                      End If
                    End If
                  Next intJ
                End If
                ' Do we now have a result?
                If (strLemmaAlt <> "") Then
                  ' Process the result
                  strLemma = arPreVerb(intI) & strLemmaAlt
                  Exit For
                End If
              End If
            Next intI
          End If
        End If
      End If
      If (strLemma = "") AndAlso (ndxThis IsNot Nothing) Then
        ' We really do have to ask the user now...
        If (Not MorphLemmaAsk(ndxThis, strVern, strPos, "", strLemma)) Then Return False
      End If
      ' Some features may be added just because of the POS 
      Select Case strPos
        Case "BE", "HV"
          colFeat.AddUnique("mood=ib")
        Case "BE^D", "HV^D"
          colFeat.AddUnique("mood=ii")
        Case "BAG", "VAG", "BAG^N", "VAG^N", "HAG", "HAG^N"
          colFeat.AddUnique("mood=pt")
          colFeat.AddUnique("tense=pa")
        Case "BAG^A", "VAG^A", "HAG^A"
          colFeat.AddUnique("mood=pt")
          colFeat.AddUnique("tense=pa")
          colFeat.AddUnique("case=ac")
        Case "BAG^D", "VAG^D", "HAG^D"
          colFeat.AddUnique("mood=pt")
          colFeat.AddUnique("tense=pa")
          colFeat.AddUnique("case=dt")
        Case "BEN", "VBN", "BEN^N", "VBN^N", "HAN", "HAN^N"
          colFeat.AddUnique("mood=pt")
          colFeat.AddUnique("tense=pa")
        Case "BEN^A", "VBN^A", "HAN^A"
          colFeat.AddUnique("mood=pt")
          colFeat.AddUnique("tense=pa")
          colFeat.AddUnique("case=ac")
        Case "BEN^D", "VBN^D", "HAN^D"
          colFeat.AddUnique("mood=pt")
          colFeat.AddUnique("tense=pa")
          colFeat.AddUnique("case=dt")
        Case "BED", "MDD", "VBD", "HVD", "AXD"
          colFeat.AddUnique("tense=pa")
        Case "BEI", "MDI", "VBI", "HVI", "AXI"
          colFeat.AddUnique("mood=im")
        Case "BEP", "MDP", "VBP", "HVP", "AXP"
          colFeat.AddUnique("tense=pr")
        Case "BEDI", "MDDI", "VBDI", "HVDI", "AXDI"
          colFeat.AddUnique("tense=pa")
          colFeat.AddUnique("mood=in")
        Case "BEDS", "MDDS", "VBDS", "HVDS", "AXDS"
          colFeat.AddUnique("tense=pa")
          colFeat.AddUnique("mood=su")
        Case "BEPI", "MDPI", "VBPI", "HVPI", "AXPI"
          colFeat.AddUnique("tense=pr")
          colFeat.AddUnique("mood=in")
        Case "BEPS", "MDPS", "VBPS", "HVPS", "AXPS"
          colFeat.AddUnique("tense=pr")
          colFeat.AddUnique("mood=su")
        Case (strPos Like "*^G")
          colFeat.AddUnique("case=gn")
        Case (strPos Like "*^D")
          colFeat.AddUnique("case=dt")
        Case (strPos Like "*^A")
          colFeat.AddUnique("case=ac")
      End Select
      ' Re-combine the features that have matched
      strFeat = colFeat.Semi
      ' Double check
      If (strFeat = ";") Then strFeat = ""
      ' Return success
      Return (strLemma <> "")
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphLemmaGet error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphLemmaAsk
  ' Goal:   Ask for a lemma and return it
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphLemmaAsk(ByRef ndxThis As XmlNode, ByVal strVern As String, ByVal strPos As String, ByVal strType As String, _
                                ByRef strLemma As String) As Boolean
    Try
      ' Validate
      If (strVern = "") OrElse (strPos = "") Then Return False
      ' Initialise
      strLemma = ""
      ' Start asking
      With frmLemmaAsk
        .Word = strVern
        .Pos = strPos
        .Context = ClauseToHtml(ndxThis)
        Select Case strType
          Case "BE"
            .MorphLemma = Split(colVerbsBE.Text, vbCrLf)
          Case "AX"
            .MorphLemma = Split(colVerbsAX.Text, vbCrLf)
          Case "VB"
            .MorphLemma = Split(colVerbsVB.Text, vbCrLf)
        End Select
        Select Case .ShowDialog
          Case DialogResult.Cancel
            bInterrupt = True
            Return False
          Case DialogResult.Yes
            ' We need to save the file and stop
            If (pdxCurrentFile IsNot Nothing) AndAlso (strCurrentFile <> "") Then
              pdxCurrentFile.Save(strCurrentFile)
              Status("Changes have been saved to: " & strCurrentFile)
            End If
            bInterrupt = True
            Return False
          Case DialogResult.Ignore
            ' Ignore this one, but do not switch on interrupt
            Return False
        End Select
        ' Get the lemma
        strLemma = .Lemma
      End With
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphLemmaAsk error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphBuildLists
  ' Goal:   Build collections of possible BE, AX and VB lemma's
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphBuildLists() As Boolean
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim intI As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage

    Try
      ' Validate
      If (colVerbsAX.Count > 0) OrElse (colVerbsBE.Count > 0) OrElse (colVerbsVB.Count > 0) Then Return True
      ' Find all AX verbs
      dtrFound = tdlMorphDict.Tables("Morph").Select("Pos LIKE 'AX*'")
      For intI = 0 To dtrFound.Length - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ dtrFound.Length
        Status("MorphBuild AX " & intPtc & "%", intPtc)
        ' Add the lemma
        colVerbsAX.AddUnique(dtrFound(intI).Item("l"))
      Next intI
      ' Find all BE verbs
      dtrFound = tdlMorphDict.Tables("Morph").Select("Pos LIKE 'BE*'")
      For intI = 0 To dtrFound.Length - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ dtrFound.Length
        Status("MorphBuild BE " & intPtc & "%", intPtc)
        ' Add the lemma
        colVerbsBE.AddUnique(dtrFound(intI).Item("l"))
      Next intI
      ' Find all VB verbs
      dtrFound = tdlMorphDict.Tables("Morph").Select("Pos LIKE 'V*'")
      For intI = 0 To dtrFound.Length - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ dtrFound.Length
        Status("MorphBuild VB " & intPtc & "%", intPtc)
        ' Add the lemma
        colVerbsVB.AddUnique(dtrFound(intI).Item("l"))
      Next intI
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphBuildLists error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphLemmaFeat
  ' Goal:   See if we can find Vern/Pos in the [Morph] table
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function MorphLemmaFeat(ByVal strVern As String, ByVal strPos As String, ByRef strLemma As String, _
                                   ByRef strFeat As String, ByVal intPtc As Integer) As Boolean
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim arFeat() As String        ' Array of features
    Dim colFeat As New StringColl ' Collection of features
    Dim strVernAlt As String      ' Spelling variation
    Dim strPosAlt As String       ' Adapted POS
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intFreqT As Integer       ' Sum frequency
    Dim intFreqI As Integer       ' Incremental Frequency
    Dim arSpelIn() As String = {"dh", "th", "y", "ie"}
    Dim arSpelOut() As String = {"th", "dh", "ie", "y"}

    Try
      ' Check if POS adaptation is needed
      If (strPos Like "ADV*") Then
        strPosAlt = "ADV"
      ElseIf (strPos Like "ADJ*") Then
        strPosAlt = "ADJ"
      ElseIf (strPos Like "WADJ*") Then
        strPosAlt = "WADJ"
      ElseIf (strPos Like "WADV*") Then
        strPosAlt = "WADV"
      ElseIf (strPos Like "N^*") Then
        strPosAlt = "N^*"
      Else
        strPosAlt = strPos
      End If
      ' See if Vern/Pos can be found in the Morph table
      dtrFound = tdlMorphDict.Tables("Morph").Select("Vern='" & strVern & "' AND Pos LIKE '" & strPos & "'", "Freq DESC")
      ' Any results?
      If (dtrFound.Length = 0) Then
        ' Try alternative POS
        dtrFound = tdlMorphDict.Tables("Morph").Select("Vern='" & strVern & "' AND Pos LIKE '" & strPosAlt & "'", "Freq DESC")
        If (dtrFound.Length = 0) Then
          ' Try spelling variations
          For intI = 0 To arSpelIn.Length - 1
            strVernAlt = strVern.Replace(arSpelIn(intI), arSpelOut(intI))
            dtrFound = tdlMorphDict.Tables("Morph").Select("Vern='" & strVernAlt & "' AND Pos LIKE '" & strPos & "'", "Freq DESC")
            ' Any results?
            If (dtrFound.Length > 0) Then Exit For
          Next intI
          ' Try spelling variations with different POS
          For intI = 0 To arSpelIn.Length - 1
            strVernAlt = strVern.Replace(arSpelIn(intI), arSpelOut(intI))
            dtrFound = tdlMorphDict.Tables("Morph").Select("Vern='" & strVernAlt & "' AND Pos LIKE '" & strPosAlt & "'", "Freq DESC")
            ' Any results?
            If (dtrFound.Length > 0) Then Exit For
          Next intI
        End If
      End If
      ' Process them
      If (dtrFound.Length > 0) Then
        ' Get the total frequency
        intFreqT = 0 : For intI = 0 To dtrFound.Length - 1
          intFreqT += dtrFound(intI).Item("Freq")
        Next intI
        ' Take the top-frequency lemma and the top-frequency features
        strFeat = dtrFound(0).Item("f") : colFeat.AddSemi(strFeat)
        strLemma = dtrFound(0).Item("l") : intFreqI = dtrFound(0).Item("Freq")
        ' Process the remainder
        For intI = 1 To dtrFound.Length - 1
          ' Access this element
          With dtrFound(intI)
            ' Adapt frequency
            intFreqI += dtrFound(0).Item("Freq")
            ' Are there differences in lemma within the first 90%?
            If (strLemma <> .Item("l")) AndAlso ((100 * intFreqI \ intFreqT) < intPtc) Then
              ' We cannot find the correct lemma
              strLemma = ""
              Exit For
            End If
            ' Are we within the first 90%?
            If ((100 * intFreqI \ intFreqT) < intPtc) Then
              ' What are the features of this item?
              arFeat = Split(.Item("f"), ";")
              ' Check if all of the features in [colFeat] are present within [arFeat]
              For intJ = colFeat.Count - 1 To 0 Step -1
                If (Not arFeat.Contains(colFeat.Item(intJ))) Then
                  ' We have to remove that feature from the collection
                  colFeat.DelItem(intJ)
                End If
              Next intJ
            End If
          End With
        Next intI
      Else
        ' Try to look in the [Lemma] table
        dtrFound = tdlMorphDict.Tables("Lemma").Select("Vern='" & strVern & "' AND Pos='" & LabelLemmaOE(strPos) & "'")
        If (dtrFound.Length > 0) Then
          strLemma = strVern
        End If
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphLemmaFeat error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphRemoveFeat
  ' Goal:   Remove feature of type [strType] from semistack [strFeat]
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function MorphRemoveFeat(ByRef strFeat As String, ByVal strType As String) As Boolean
    Dim colFeat As New StringColl ' Collection of features
    Dim arFeat() As String        ' Array of features
    Dim intI As Integer           ' Counter

    Try
      arFeat = Split(strFeat, ";")
      For intI = 0 To arFeat.Length - 1
        If (InStr(arFeat(intI), strType & "=") = 0) Then
          ' Copy it 
          colFeat.Add(arFeat(intI))
        End If
      Next intI
      ' Recombine
      strFeat = colFeat.Semi
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphRemoveFeat error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphMakeDict
  ' Goal:   Make an OE dictionary (OBSOLETE)
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphMakeDict() As Boolean
    Dim intI As Integer       ' Counter
    Dim intJ As Integer       ' Counter
    Dim intPtc As Integer     ' Percentage
    Dim intPos As Integer     ' Position
    Dim intSpaces As Integer  ' Number of spaces
    Dim arLine() As String    ' Dictionary divided into lines
    Dim arSfm() As String     ' Line broken up into sfm's
    Dim colSfm As New List(Of String)  ' Collection of sfm's
    Dim strLine As String     ' One line
    Dim strSfm As String      ' One sfm
    Dim strText As String     ' Remainder
    Dim strItem As String     ' One item
    Dim strFileIn As String = "d:\data files\corpora\dictionaries\2013_OEdict_db.txt"
    Dim strFileOut As String = "d:\data files\corpora\dictionaries\2013_OEdict_db.sfm"

    Try
      ' Validate
      If (Not IO.File.Exists(strFileIn)) Then Return False
      ' Load dictionary into array
      arLine = Split(IO.File.ReadAllText(strFileIn), "\lex ")
      ' Process each line
      For intI = 0 To arLine.Length - 1
        ' Where are we?
        intPtc = (intI + 1) * 100 \ arLine.Length
        Status("Dictionary " & intPtc & "%", intPtc)
        ' Get this line
        strLine = MyTrim(arLine(intI))
        strLine = strLine.Replace(vbCr, "")
        strLine = strLine.Replace(vbLf, " ")
        ' Break line up into sfms
        arSfm = Split(strLine, "\") : colSfm.Clear()
        ' Check the first item
        strText = Trim(arSfm(0))
        intPos = InStr(strText, " see ")
        If (intPos > 0) Then
          colSfm.Add(Left(strText, intPos - 1))
          colSfm.Add("cf " & Mid(strText, intPos + 5))
        Else
          colSfm.Add(strText)
        End If
        ' Process the sfm's
        For intJ = 1 To arSfm.Length - 1
          ' Get this item
          strItem = Trim(arSfm(intJ))
          ' Split into sfm text and body
          intPos = InStr(strItem, " ")
          If (intPos > 0) Then
            strSfm = Left(strItem, intPos - 1)
            If (Left(strSfm, 2) = "mp") AndAlso (strSfm.Length > 2) Then
              '  Stop
              strItem = "mp " & Mid(strItem, 3)
              intPos = 3
              strSfm = "mp"
            End If
            strText = Trim(Mid(strItem, intPos + 1))
            ' Get the number of spaces in this item
            intSpaces = Split(strText, " ").Count
            If (intSpaces > 1) Then
              ' We have to split this one after the first space
              intSpaces = InStr(strText, " ")
              colSfm.Add(strSfm & " " & Left(strText, intSpaces))
              colSfm.Add("txt " & Mid(strText, intSpaces + 1))
            ElseIf (intJ > 1) AndAlso (strSfm = "prn") AndAlso (intJ = arSfm.Length - 1) Then
              ' This needs to be "deleted" --> do not add it to the SFM collection!
              ' Stop
              ' Just add the [strText] to the last part of the previous one
              colSfm.Item(colSfm.Count - 1) &= " " & strText
            Else
              ' No split is needed
              colSfm.Add(strItem)
            End If
          Else
            ' If it does not break up...
            colSfm.Add(strItem)
          End If
        Next intJ
        ' Recombine 
        arLine(intI) = Join(colSfm.ToArray, vbCrLf & "\")

      Next intI
      ' Recombine text
      strText = Join(arLine, vbCrLf & "\lex ")
      ' Write text back
      IO.File.WriteAllText(strFileOut, strText)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphMakeDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitBT
  ' Goal:   Initialise the B&T reader
  ' History:
  ' 22-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function InitBT() As Boolean
    Dim strFile As String = "d:\data files\dbase\BTlemma.xml"

    Try
      ' Check
      If (tdlBT IsNot Nothing) Then Return True
      ' Validate
      If (Not IO.File.Exists(strFile)) Then
        Return False
      End If
      ' Load the B&T
      If (Not ReadDataset("BTlemList.xsd", strFile, tdlBT)) Then Return False
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/InitBT error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetBtLemma
  ' Goal:   Given the BT number, get the lemma
  ' History:
  ' 24-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetBtLemma(ByVal intBt As Integer, ByRef strLemma As String) As Boolean
    Dim dtrFound() As DataRow ' Result of select

    Try
      ' Validate
      If (tdlBT Is Nothing) Then Return False
      ' Look for the number
      dtrFound = tdlBT.Tables("BT").Select("bt=" & intBt)
      If (dtrFound.Length = 0) Then
        Return False
      Else
        strLemma = dtrFound(0).Item("l").ToString
        Return True
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/GetBtLemma error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   PosBTandVP
  ' Goal:   Compare BT and VernPos POS tags
  ' History:
  ' 13-06-2014  ERK Derived from [MorphResolveBT]
  ' ------------------------------------------------------------------------------------
  Private Function PosBTandVP(ByVal strPosBT As String, ByVal strPosVP As String) As Boolean
    Dim strPosVPred As String ' Reduced VernPos pos

    Try
      ' Determine the reduced VP pos
      If (DoLike(strPosVP, "V*|*+V*|B[EA]*|*+B[EA]*|MD*|HV*|*+HV*")) Then
        strPosVPred = "VB"
      ElseIf (DoLike(strPosVP, "ADJ*")) Then
        strPosVPred = "ADJ"
      ElseIf (DoLike(strPosVP, "ADV*")) Then
        strPosVPred = "ADV"
      ElseIf (DoLike(strPosVP, "NUM*")) Then
        strPosVPred = "NUM"
      ElseIf (DoLike(strPosVP, "N|N^*|NR|NR^*")) Then
        strPosVPred = "N"
      ElseIf (DoLike(strPosVP, "P")) Then
        strPosVPred = "P"
      Else
        strPosVPred = strPosVP
      End If
      ' Make a comparison
      Return (strPosBT = strPosVPred)
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/PosBTandVP error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphResolveVernPosBT
  ' Goal:   Review [VernPos], where no BT numbers are supplied, try to get these from the BT dictionary
  ' History:
  ' 13-06-2014  ERK Derived from [MorphResolveBT]
  ' ------------------------------------------------------------------------------------
  Public Function MorphResolveVernPosBT(ByVal strLngAbbr As String, ByRef bChanged As Boolean) As Boolean
    Dim dtrVP() As DataRow    ' VernPos table in order of ascending @Vern
    Dim dtrMorph() As DataRow ' Morph table 
    Dim dtrNew As DataRow     ' New datarow
    Dim strVern As String     ' Vernacular
    Dim strLemma As String    ' Lemma
    Dim strLemmaOld As String ' Old lemma that is to be replaced
    Dim strPos As String      ' Part of speech
    Dim strV As String        ' @v attribute
    Dim strF As String        ' @f attribute
    Dim strBtFile As String = "d:\data files\dbase\OEbtDict.xml"
    Dim rdThis As XmlReader = Nothing
    Dim setXmlRd As New XmlReaderSettings()
    Dim pdxThis As New XmlDocument
    Dim ndxThis As XmlNode    ' One BT node
    Dim bFound As Boolean     ' Found hit orn ot
    Dim intBT As Integer      ' BT number
    Dim intI As Integer       ' Counter
    Dim intJ As Integer       ' Counter
    Dim intPtc As Integer     ' Percentage

    Try
      ' Validate
      If (strLngAbbr <> "OE") Then Return False
      ' Read the BT lemma *list*
      Status("Reading BT...")
      If (Not InitBT()) Then Logging("Could not initialize BT reading") : Return False
      ' Start reading the BT xml file
      If (Not IO.File.Exists(strBtFile)) Then Return False
      ' Initialise XML reader
      setXmlRd.ProhibitDtd = False : setXmlRd.IgnoreComments = True : setXmlRd.IgnoreWhitespace = True
      rdThis = XmlReader.Create(strBtFile, setXmlRd) : ndxThis = Nothing
      ' Read the first <BT> node
      rdThis.ReadToFollowing("BT")
      If (Not XmlReaderGetNextNode(rdThis, pdxThis, "BT", ndxThis)) Then Return False
      ' Get the @v attribute
      strV = ndxThis.Attributes("v").Value
      ' Get all the entries from [Remain] in order
      Status("Reading VernPos...")
      dtrVP = tdlMorphDict.Tables("VernPos").Select("", "Vern ASC")
      For intI = 0 To dtrVP.Length - 1
        ' Get this Vern item and its POS
        strVern = dtrVP(intI).Item("Vern").ToString
        strPos = dtrVP(intI).Item("Pos").ToString
        strLemmaOld = dtrVP(intI).Item("l").ToString
        ' Show where we are
        intPtc = (intI + 1) * 100 \ dtrVP.Length
        Status("MorphResolveVernPosBT " & intPtc & "% [" & Left(strVern, 3) & "]", intPtc)
        ' Get the @f attribute to see whether a BT number is known already
        strF = dtrVP(intI).Item("f").ToString
        If (Not Regex.IsMatch(strF, "BT\d+")) Then
          ' No BT number in @f yet ==> check if we can find the form in OEbtDict.xml...
          ' Is the vern okay already?
          bFound = False
          If (strV <> strVern) Then
            ' Check if the @Vern we are looking for can be found here
            While (StrComp(strV, strVern) <= 0) AndAlso (Not rdThis.EOF)
              ' Read further
              If (XmlReaderGetNextNode(rdThis, pdxThis, "BT", ndxThis)) Then
                ' Get the @v
                strV = ndxThis.Attributes("v").Value
                ' Check if we are there
                If (StrComp(strV, strVern) = 0) AndAlso PosBTandVP(ndxThis.Attributes("ps").Value, strPos) Then bFound = True : Exit While
              End If
            End While
          End If
          ' What happened?
          If (bFound) Then
            ' We have a hit! - get the BT number
            intBT = ndxThis.Attributes("bt").Value
            ' Look up this value in the list
            If (GetBtLemma(intBT, strLemma)) Then
              ' We have a lemma and a BT number!! -- Process in [VernPos]
              AddSemiStack(strF, "s=BT" & Format(intBT, "000000"))
              dtrVP(intI).Item("f") = strF : dtrVP(intI).Item("l") = strLemma
              ' Show our findings
              Logging(strPos & vbTab & strVern & vbTab & strLemmaOld & vbTab & strLemma & vbTab & "BT" & intBT, False)
              ' Make sure the entries in [Morph] for this lemma is adapted too
              dtrMorph = tdlMorphDict.Tables("Morph").Select("l='" & strLemmaOld & "'")
              For intJ = 0 To dtrMorph.Length - 1
                With dtrMorph(intJ)
                  ' Check if this has the correct POS
                  If (MorphCatMain(.Item("Pos").ToString) = MorphCatMain(strPos)) Then
                    ' Adapt the lemma and the @f attribute
                    strF = .Item("f").ToString
                    If (Not Regex.IsMatch(strF, "BT\d+")) Then
                      AddSemiStack(strF, "s=BT" & Format(intBT, "000000"))
                      .Item("f") = strF : .Item("l") = strLemma
                    End If
                  End If
                End With
              Next intJ
              bChanged = True
            End If
          End If
        End If
        ' Check if we are out of options...
        If (rdThis.EOF) Then Exit For
      Next intI
      ' Close reader
      rdThis.Close()
      ' Return positively
      Return True
    Catch ex As Exception
      ' Close reader
      If (Not rdThis Is Nothing) Then rdThis.Close()
      ' Show error
      HandleErr("modMorph/MorphResolveVernPosBT error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphResolveBT
  ' Goal:   Review [Remain], try to put entries in [VernPos], taking them from the BT dictionary
  ' History:
  ' 24-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphResolveBT(ByVal strLngAbbr As String, ByRef bChanged As Boolean) As Boolean
    Dim dtrRem() As DataRow   ' Remaining stuff in order
    Dim dtrNew As DataRow     ' New datarow
    Dim strVern As String     ' Vernacular
    Dim strLemma As String    ' Lemma
    Dim strPos As String      ' Part of speech
    Dim strV As String        ' @v attribute
    Dim strBtFile As String = "d:\data files\dbase\OEbtDict.xml"
    Dim rdThis As XmlReader = Nothing
    Dim setXmlRd As New XmlReaderSettings()
    Dim pdxThis As New XmlDocument
    Dim ndxThis As XmlNode    ' One BT node
    Dim intBT As Integer      ' BT number
    Dim intI As Integer       ' Counter
    Dim intPtc As Integer     ' Percentage

    Try
      ' Validate
      If (strLngAbbr <> "OE") Then Return False
      ' Read the BT lemma *list*
      Status("Reading BT...")
      If (Not InitBT()) Then Logging("Could not initialize BT reading") : Return False
      ' Start reading the BT xml file
      If (Not IO.File.Exists(strBtFile)) Then Return False
      ' Initialise XML reader
      setXmlRd.ProhibitDtd = False : setXmlRd.IgnoreComments = True : setXmlRd.IgnoreWhitespace = True
      rdThis = XmlReader.Create(strBtFile, setXmlRd) : ndxThis = Nothing
      ' Read the first <BT> node
      rdThis.ReadToFollowing("BT")
      If (Not XmlReaderGetNextNode(rdThis, pdxThis, "BT", ndxThis)) Then Return False
      ' Get the @v attribute
      strV = ndxThis.Attributes("v").Value
      ' Get all the entries from [Remain] in order
      Status("Reading remain...")
      dtrRem = tdlRemain.Tables("Remain").Select("", "Vern ASC")
      For intI = 0 To dtrRem.Length - 1
        ' Get this Vern item and its POS
        strVern = dtrRem(intI).Item("Vern").ToString
        strPos = dtrRem(intI).Item("Pos").ToString
        ' Is the vern okay already?
        If (strV <> strVern) Then
          ' Check if the @Vern we are looking for can be found here
          While (StrComp(strV, strVern) <= 0) AndAlso (Not rdThis.EOF)
            ' Read further
            If (XmlReaderGetNextNode(rdThis, pdxThis, "BT", ndxThis)) Then
              ' Get the @v
              strV = ndxThis.Attributes("v").Value
              ' Check if we are there
              If (StrComp(strV, strVern) = 0) AndAlso (ndxThis.Attributes("ps").Value Like "V*") Then Exit While
            End If
          End While
        End If
        ' What happened?
        If (strV = strVern) Then
          ' We have a hit! - get the BT number
          intBT = ndxThis.Attributes("bt").Value
          ' Look up this value in the list
          If (GetBtLemma(intBT, strLemma)) Then
            ' We have a lemma and a BT number!! -- Process in [VernPos]
            If (Not MorphDictAdd(strVern, strPos, "", strLemma, "s=BT" & intBT)) Then Return False
            ' Add it to VernPos
            If (Not VernPosAdd(strPos, strVern, "LemmaFeat", strLemma, "1", "s=BT" & intBT)) Then Return False
            bChanged = True
          End If
        End If
        ' Check if we are out of options...
        If (rdThis.EOF) Then Exit For
      Next intI
      ' Close reader
      rdThis.Close()
      ' Return positively
      Return True
    Catch ex As Exception
      ' Close reader
      If (Not rdThis Is Nothing) Then rdThis.Close()
      ' Show error
      HandleErr("modMorph/MorphResolveBT error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphResolveTable
  ' Goal:   Process the resolved entries from table [Remain] into table [VernPos]
  '         The resolved entries are in a tab-separated text file
  ' History:
  ' 27-04-2013  ERK Created
  ' 20-01-2014  ERK Added information source value
  ' ------------------------------------------------------------------------------------
  Public Function MorphResolveTable(ByVal strFile As String, Optional ByVal strLngAbbr As String = "") As Boolean
    Dim strText As String     ' Whole text
    Dim strDelim As String    ' Delimiter
    Dim arText() As String    ' Whole text
    Dim arLine() As String    ' One line
    Dim arLemm() As String    ' One lemma complex
    Dim strPos As String      ' Part of speech
    Dim strLemma As String    ' Lemma
    Dim strTryLem As String   ' Trial lemma
    Dim strTryPos As String   ' Trial POS
    Dim strLev As String = "1"  ' Level
    Dim strVern As String     ' Vernacular
    Dim strFeats As String    ' Features
    Dim strSrc As String
    Dim strType As String     ' Type of addition
    Dim strHead As String     ' POS of the lemma-category
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim dtrNew As DataRow     ' One added row
    Dim bSkip As Boolean      ' Skip this element or not?
    Dim bSpecial As Boolean   ' Special treatment of tables (e.g. for OE; one column contains okay markers)
    Dim bMaySkip As Boolean = False ' we may skip those items that are there already
    Dim bAskSkip As Boolean = False ' Skipping has been asked
    Dim intNum As Integer = 0 ' Counter
    Dim intI As Integer       ' Counter
    Dim intPtc As Integer     ' Percentage
    Dim intResolveColumns As Integer = 8    ' Number of columns used in RESOLVE excel table
    Dim intColVern As Integer = 0
    Dim intColPos As Integer = 1
    Dim intColLemma As Integer = 5
    Dim intColFeat As Integer = 6
    Dim colLemmaVern As New StringColl  ' Alle lemma-vern combinations
    Dim colAsk As New StringColl        ' Lemma's to be checked with the user

    Try
      ' Validate
      If (Not IO.File.Exists(strFile)) Then Return False
      ' Check if BT needs to be initialised
      If (strLngAbbr = "OE") Then
        Status("Reading BT...")
        If (Not InitBT()) Then Logging("Could not initialize BT reading") : Return False
      End If
      ' OE special treatment
      If (strLngAbbr = "OE") Then
        Select Case MsgBox("Does column [6] contain feature(s) (=Yes) or Check marks (=No)?", MsgBoxStyle.YesNoCancel)
          Case MsgBoxResult.Cancel
            Return False
          Case MsgBoxResult.Yes
            bSpecial = False
          Case MsgBoxResult.No
            bSpecial = True
        End Select
      End If
      ' Get lemma source
      strSrc = GetLemmaSource()
      ' Read the file
      strText = IO.File.ReadAllText(strFile)
      ' Get delimiter
      strDelim = GetDelim(strText, vbCrLf, vbCr, vbLf)
      arText = Split(strText, strDelim)
      ' Initialise the dataset for additions
      If (tdlLemmaAsk IsNot Nothing) Then tdlLemmaAsk.Dispose()
      tdlLemmaAsk = Nothing
      If (Not CreateDataSet("MorphDict.xsd", tdlLemmaAsk)) Then Return False
      ' Check line by line for the lemma's that have been supplied
      For intI = 0 To arText.Length - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ arText.Length
        Status("Reading table " & intPtc & "%", intPtc)
        ' Get this line
        ' arLine = Split(arText(intI), ";")
        If (InStr(arText(intI), vbTab) > 0) Then
          arLine = Split(arText(intI), vbTab)
        ElseIf (InStr(arText(intI), ";") > 0) Then
          arLine = Split(arText(intI), ";")
        Else
          ' Cannot determine how it is split up...
          arLine = Split(arText(intI), ";")
        End If
        ' How's this? Allow more columns if needed
        If (arLine.Length >= intResolveColumns) Then
          ' Does it contain a resolution in column X?\
          ' (NB: also make sure that the lemma is not just a space...)
          strLemma = Trim(arLine(intColLemma))
          If (strLemma <> "") Then
            ' Get elements
            strVern = arLine(intColVern)
            strPos = arLine(intColPos)
            If (strPos Like "*[234][1234]") Then
              strPos = Left(strPos, strPos.Length - 2)
            End If
            ' ======================= DEBUG
            ' If (strPos = "?") Then Stop
            ' =====================================
            ' Include or skip?
            bSkip = False
            ' See if there are any features
            If (strLngAbbr <> "OE") OrElse (Not bSpecial) Then strFeats = arLine(intColFeat)
            ' For OE we need to check if this lemma has been 'okayed'
            If (strLngAbbr = "OE") AndAlso (bSpecial) Then
              If (arLine(intColFeat) <> "✓") OrElse (strPos = "?") Then
                ' Indicate that this one needs to be skipped
                bSkip = True
              End If
            End If
            ' Check if the entry already occurs in [VernPos]
            dtrFound = tdlMorphDict.Tables("VernPos").Select("Vern='" & arLine(intColVern).Replace("'", "''") & "'" & _
                                                             " AND Pos='" & arLine(intColPos) & "'")
            If (dtrFound.Length > 0) Then
              If (Not bAskSkip) Then
                Select Case MsgBox("Some Vern/Pos combinations have already been labelled." & vbCrLf & _
                                   "Shall I re-do these? (Y)", vbYesNoCancel)
                  Case MsgBoxResult.Cancel
                    Return False
                  Case MsgBoxResult.No
                    bMaySkip = True
                  Case MsgBoxResult.Yes
                    bMaySkip = False
                End Select
                ' Indicate we have been asked
                bAskSkip = True
              End If
              ' Can we skip?
              If (bMaySkip) Then bSkip = True
            End If
            If (Not bSkip) Then
              ' Add lemma + vernacular to the collection
              colLemmaVern.AddUnique(strLemma & ";" & strVern & ";" & strPos, "-")
              Application.DoEvents()
              ' Add entry to [tdlLemmaAsk]
              dtrNew = AddOneDataRow(tdlLemmaAsk, "Morph", "MorphId", "MorphList")
              With dtrNew
                ' Use @t alternatively: to indicate the POS of the lemma
                .Item("Vern") = arLine(intColVern) : .Item("l") = arLine(5) : .Item("Pos") = arLine(intColPos) : .Item("t") = ""
                ' Use the @Label attribute to give the forest location
                .Item("Label") = arLine(7)
                ' Use the @File attribute to contain the example
                If (arLine.Length >= 9) Then
                  .Item("File") = arLine(8)
                Else
                  .Item("File") = ""
                End If
                ' Use @h to signal status -- default is "ok"
                .Item("h") = "ok"
              End With
            End If
          End If
        End If
      Next intI
      ' Initialise the dataset for additions
      If (tdlLemmaVern IsNot Nothing) Then tdlLemmaVern.Dispose()
      tdlLemmaVern = Nothing
      If (Not CreateDataSet("MorphDict.xsd", tdlLemmaVern)) Then Return False
      ' Check all lemma's on their existence in [Morph]
      For intI = 0 To colLemmaVern.Count - 1
        ' Validate
        If (colLemmaVern.Item(intI) <> "") Then
          ' Access lemma, vern and POS
          arLemm = Split(colLemmaVern.Item(intI), ";") : strType = "" : strHead = "" : bSkip = False
          If (Not bSkip) Then
            ' Check existence of this lemma
            If (Not MorphLemmaCheck(arLemm(0), arLemm(1), arLemm(2), strHead, strType)) Then Return False
            ' Special case for non-existing OE lemma's: check BT lemma list
            If (strType = "none") AndAlso (strLngAbbr = "OE") Then
              ' Define ps
              If (DoLike(arLemm(2), "N|N^*")) Then
                strTryPos = "N"
              ElseIf (DoLike(arLemm(2), "V*|RP+*")) Then
                strTryPos = "VB"
              ElseIf (DoLike(arLemm(2), "ADJ*")) Then
                strTryPos = "ADJ"
              ElseIf (DoLike(arLemm(2), "ADV*")) Then
                strTryPos = "ADJ"
              Else
                strTryPos = arLemm(2).Replace("'", "''")
              End If
              ' Define lemma
              strTryLem = arLemm(0).Replace("'", "''")
              ' Check if the lemma occurs in the list of lemma's for this kind of word in the BT list
              dtrFound = tdlBT.Tables("BT").Select("b='" & strTryLem.Replace("-", "") & "' AND ps='" & strTryPos & "'")
              If (dtrFound.Length > 0) Then
                strType = "unique"
              ElseIf (InStr(strTryLem, "-") > 0) Then
                ' Try the part after the last hyphen
                strTryLem = Mid(strTryLem, InStrRev(strTryLem, "-") + 1)
                dtrFound = tdlBT.Tables("BT").Select("b='" & strTryLem & "' AND ps='" & strTryPos & "'")
                If (dtrFound.Length > 0) Then strType = "unique"
              End If
            End If
            Select Case strType
              Case "ambi"
                ' Lemma is ambiguous --> this is possible: there may be multiple sences and transitivity types
              Case "unique"
                ' Lemma is unique --> okay!!
              Case "none"
                ' Lemma does not exist --> signal this problem by adding it to the [tdlLemmaVern] dataset
                dtrNew = AddOneDataRow(tdlLemmaVern, "Morph", "MorphId", "MorphList")
                With dtrNew
                  ' Use @t alternatively: to indicate the POS of the lemma
                  .Item("Vern") = arLemm(1) : .Item("l") = arLemm(0) : .Item("Pos") = arLemm(2) : .Item("t") = strHead
                  ' Use @h to signal what action needs to be taken -- default is "add"
                  .Item("h") = "add"
                End With
                ' Find the entry in [tdlLemmaAsk]
                dtrFound = tdlLemmaAsk.Tables("Morph").Select("Vern='" & arLemm(1).Replace("'", "''") & "' AND l='" & arLemm(0).Replace("'", "''") & "'" & _
                                                              " AND Pos='" & arLemm(2).Replace("'", "''") & "'")
                If (dtrFound.Length > 0) Then
                  dtrFound(0).Item("h") = "check"
                  dtrFound(0).Item("t") = strHead
                End If
            End Select
          End If
        End If
      Next intI
      ' Prepare a table with all the elements that need to be checked
      dtrFound = tdlLemmaAsk.Tables("Morph").Select("h='check'", "Vern ASC, l ASC")
      If (dtrFound.Length > 0) Then
        ' Make a report
        colAsk.Add("<html><body><table><tr><td>Vernacular</td><td>POS</td><td>Short</td><td>Ptc</td>" & _
                   "<td>Derivation</td><td>Suggestion</td><td>Done</td><td>Text:forest</td><td>Sentence</td></tr>")
        For intI = 0 To dtrFound.Length - 1
          With dtrFound(intI)
            ' Add one line
            colAsk.Add("<tr><td>" & .Item("Vern").ToString & "</td><td>" & .Item("Pos").ToString & "</td><td>" & .Item("t").ToString & "</td><td>0</td><td>(-)</td>" & _
                       "<td>" & .Item("l") & "</td><td>?</td><td>" & .Item("Label").ToString & "</td><td>" & .Item("File").ToString & "</td></tr>")
          End With
        Next intI
        ' Finish
        colAsk.Add("</table></body></html>")
        frmMain.wbReport.DocumentText = colAsk.Text
        ' Switch to report
        frmMain.TabControl1.SelectedTab = frmMain.tpReport
        ' Show what we mean
        Status("There are some problematic cases to be reviewed")
      Else
        ' Show we are okay
        Status("There are no problematic cases")
      End If
      ' Need to ask anything?
      If (tdlLemmaVern.Tables("Morph").Rows.Count > 0) Then
        ' Present the problems in a form...
        With frmLemmaVern
          Select Case .ShowDialog
            Case DialogResult.No, DialogResult.Cancel
              ' Don't continue!
              Return False
          End Select
          ' The results should be available in [tdlLemmaVern]
        End With
      End If
      ' Process line by line
      For intI = 0 To arText.Length - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ arText.Length
        Status("Resolving table " & intPtc & "%", intPtc)
        ' Get this line
        ' arLine = Split(arText(intI), ";")
        If (InStr(arText(intI), vbTab) > 0) Then
          arLine = Split(arText(intI), vbTab)
        ElseIf (InStr(arText(intI), ";") > 0) Then
          arLine = Split(arText(intI), ";")
        Else
          ' Cannot determine how it is split up...
          arLine = Split(arText(intI), ";")
        End If
        ' How's this? Allow more columns if needed
        If (arLine.Length >= intResolveColumns) Then
          ' Does it contain a resolution in column X?\
          ' (NB: also make sure that the lemma is not just a space...)
          strLemma = Trim(arLine(intColLemma))
          If (strLemma <> "") Then
            ' Get elements
            strVern = arLine(intColVern)
            strPos = arLine(intColPos)
            If (strPos Like "*[234][1234]") Then
              strPos = Left(strPos, strPos.Length - 2)
            End If
            ' ======================= DEBUG
            ' If (strPos = "?") Then Stop
            ' =====================================
            bSkip = False
            ' For OE we need to check if this lemma has been 'okayed'
            If (strLngAbbr = "OE") AndAlso (bSpecial) Then
              If (arLine(intColFeat) <> "✓") OrElse (strPos = "?") Then bSkip = True
            End If
            ' Check if an adaptation is needed within table [Morph]
            dtrFound = tdlLemmaVern.Tables("Morph").Select("Vern='" & strVern.Replace("'", "''") & "' AND Pos='" & strPos & "'")
            If (dtrFound.Length > 0) Then
              ' Check what changes there are
              Select Case dtrFound(0).Item("h").ToString
                Case "change"
                  ' Make sure the lemma is changed before it is added to table [Morph]
                  strLemma = dtrFound(0).Item("l").ToString
                  ' Check if this new lemma exists in [Morph]
                  strType = "" : strHead = ""
                  If (Not MorphLemmaCheck(strLemma, strVern, strPos, strHead, strType)) Then Return False
                  If (strType = "none") AndAlso (InStr(strPos, "+") = 0) Then
                    ' lemma does not yet exist here, so add try add it
                    If (Not MorphDictAdd(strLemma, strHead, "", strLemma, "")) Then Return False
                    ' Also try add the vernacular
                    If (Not MorphDictAdd(strVern, strPos, strPos, strLemma, "")) Then Return False
                  End If
                Case "add"
                  ' Add the lemma to table [Morph] if needed
                  strType = "" : strHead = ""
                  If (Not MorphLemmaCheck(strLemma, strVern, strPos, strHead, strType)) Then Return False
                  If (strType = "none") AndAlso (InStr(strPos, "+") = 0) Then
                    ' lemma does not yet exist here, so add it
                    If (Not MorphDictAdd(strLemma, strHead, "", strLemma, "")) Then Return False
                    ' Also try add the vernacular form
                    If (Not MorphDictAdd(strVern, strPos, strPos, strLemma, "")) Then Return False
                  End If
                Case "delete"
                  bSkip = True
                Case Else
                  Stop
              End Select
            End If
            ' Can we continue?
            If (Not bSkip) Then
              ' Check if the *lemma* is already in the [Morph] table
              strType = "" : strHead = ""
              If (Not MorphLemmaCheck(strLemma, strLemma, strPos, strHead, strType)) Then Return False
              If (strType = "none") AndAlso (InStr(strPos, "+") = 0 OrElse DoLike(strPos, "RP+*|NEG+*")) Then
                ' lemma does not yet exist here, so add it
                If (Not MorphDictAdd(strLemma, strHead, "", strLemma, "")) Then Return False
                ' Also try add the vernacular form
                If (Not MorphDictAdd(strVern, strPos, strPos, strLemma, "")) Then Return False
              End If
              ' See if there are any features
              If (strLngAbbr <> "OE") OrElse (Not bSpecial) Then strFeats = arLine(intColFeat)
              ' Possibly add "source" feature
              If (Not DoLike(strSrc, "EK|ERK")) Then AddSemiStack(strFeats, "src=" & strSrc)
              ' Check if the lemma that has been determined actually is available in the table [Morph] (24/feb/2014)

              ' ============= debug ==============
              ' If (strFeats <> "") Then Stop
              ' ==================================
              ' Determine type
              strType = IIf(strFeats = "", "LemmaOnly", "LemmaFeat")
              ' Found a lemma for me: put it in the [VernPos] table
              If (Not VernPosAdd(strPos, strVern, strType, strLemma, strLev, strFeats)) Then Logging("VernPosAdd problem for [" & strVern & "]") : Return False
              ' Keep track of number of instances
              intNum += 1
              frmMain.tbGenSource.AppendText(strVern & " (" & strPos & ") l=" & strLemma & vbCrLf)
              Application.DoEvents()
            End If
          End If
        End If
      Next intI
      ' Show how much we've done
      Logging("Resolved " & intNum & " entries")
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphResolveTable error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   VernToMorph
  ' Goal:   Check all entries in [VernPos] for presence in [Morph] and add them if needed
  ' History:
  ' 25-02-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function VernToMorph(ByRef intNum As Integer) As Boolean
    Dim dtrMorph() As DataRow ' Result of querying [Morph]
    Dim dtrVpos() As DataRow  ' Walking [VernPos]
    Dim strVern As String     ' Current vernacular
    Dim strLemma As String    ' Current lemma
    Dim strPos As String      ' Current POS
    Dim strType As String     ' Type of presence
    Dim strHead As String     ' Lemma POS category
    Dim strFeats As String    ' Current features
    Dim intI As Integer       ' Counter
    Dim intPtc As Integer     ' Percentage

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Initialize
      intNum = 0
      ' Walk vernpos
      dtrVpos = tdlMorphDict.Tables("VernPos").Select("", "Vern ASC, l ASC, Pos ASC")
      For intI = 0 To dtrVpos.Count - 1
        ' Where are we?
        intPtc = (intI + 1) * 100 \ dtrVpos.Length
        Status("VernToMorph " & intPtc & "%", intPtc)
        ' Access this element
        With dtrVpos(intI)
          strVern = .Item("Vern").ToString : strLemma = .Item("l").ToString : strPos = .Item("Pos").ToString
          strFeats = .Item("f").ToString
        End With
        ' Check if it exists inside [Morph]
        strType = "" : strHead = ""
        If (Not MorphLemmaCheck(strLemma, strVern, strPos, strHead, strType)) Then Return False
        If (strType = "none") Then
          ' lemma does not yet exist here, so add it
          If (Not MorphDictAdd(strVern, strPos, strPos, strLemma, strFeats)) Then Return False
          intNum += 1
          Logging("v=[" & strVern & "] l=[" & strLemma & "] POS=[" & strPos & "]")
        End If
        ' Also check if the lemma is not there yet
        If (Not MorphLemmaCheck(strLemma, strLemma, strHead, strHead, strType)) Then Return False
        If (strType = "none") Then
          ' lemma does not yet exist here, so add it
          If (Not MorphDictAdd(strLemma, strHead, strHead, strLemma, "")) Then Return False
          intNum += 1
        End If
      Next intI
      ' Any changes?
      If (intNum > 0) Then
        Status("Saving: " & strMorphDictFile)
        tdlMorphDict.WriteXml(strMorphDictFile)
      End If
      ' Show how much we've done
      Logging("Added " & intNum & " entries")
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/VernToMorph error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphPfxOEdict
  ' Goal:   Read the OE dictionary from the expected place; if not available: download from internet
  ' History:
  ' 23-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MorphPfxOEdict(Optional ByVal bForce As Boolean = False) As Boolean
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim dtrMorph() As DataRow     ' Result of SELECT on tdlMorphDict
    Dim dtrNew As DataRow         ' New row
    Dim strPfx As String          ' Prefix
    Dim strRew As String          ' Rewrite
    Dim strTmp As String          ' Temporary string
    Dim arRew() As String         ' Array with rewrite rules
    Dim colThis As New StringColl ' Collection
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intPos As Integer         ' Counter
    Dim bNeed As Boolean = False  ' Need saving or not?

    Try
      ' Validate
      If (tdlOEdict Is Nothing) Then Logging("English dictionary not loaded...") : Return False
      If (tdlMorphDict Is Nothing) Then Logging("MorphDict not loaded...") : Return False
      ' Check if table [Prefix] already has members --> do not duplicate work!
      If (Not bForce) AndAlso (tdlMorphDict.Tables("Prefix").Rows.Count > 0) Then Return True
      ' Access the correct elements from [tdlOEdict], table [Entry]
      ' Old: dtrFound = tdlOEdict.Tables("Entry").Select("l LIKE '*-' AND Cf <> ''")
      dtrFound = tdlOEdict.Tables("Sense").Select("l LIKE '*-' AND Cf <> ''")
      ' Walk through all of them
      For intI = 0 To dtrFound.Length - 1
        ' Check if this one is okay
        If (dtrFound(intI).Item("Def").ToString = "") Then
          ' Get the prefix and the rewrite
          strPfx = dtrFound(intI).Item("l").ToString
          arRew = Split(dtrFound(intI).Item("Cf").ToString, ",") : colThis.Clear()
          ' Pre-process the rewrite strings
          For intJ = 0 To arRew.Length - 1
            ' Evaluate this one
            intPos = InStr(arRew(intJ), "(")
            If (intPos = 0) AndAlso (intJ = 0) Then
              ' Add to collection
              colThis.Add(Trim(arRew(intJ)))
            ElseIf (intPos > 0) Then
              ' This holds two items: extract them...
              strRew = Trim(Left(arRew(intJ), intPos - 1))
              colThis.Add(strRew)
              ' Find out where ) ends
              strTmp = Mid(arRew(intJ), intPos + 1)
              intPos = InStr(strTmp, ")")
              If (intPos > 0) Then
                strRew &= Trim(Left(strTmp, intPos - 1))
                colThis.Add(strRew)
              End If
            End If
          Next intJ
          ' Process each of the possible rewrites
          For intJ = 0 To colThis.Count - 1
            ' Get the rewrite string
            strRew = colThis.Item(intJ)
            ' Double check
            If (Right(strPfx, 1) = "-") AndAlso (Right(strRew, 1) = "-") Then
              ' Strip off the hyphens
              strPfx = Left(strPfx, strPfx.Length - 1) : strRew = Left(strRew, strRew.Length - 1)
              ' Does [tdlMorphDict] already have it?
              dtrMorph = tdlMorphDict.Tables("Prefix").Select("Pre = '" & strPfx & "' AND Rew = '" & strRew & "'")
              If (dtrMorph.Length = 0) Then
                ' Make a new one
                dtrNew = AddOneDataRow(tdlMorphDict, "Prefix", "PrefixId", "PrefixList")
                With dtrNew
                  .Item("Pre") = strPfx : .Item("Rew") = strRew : .Item("Size") = strPfx.Length
                End With
                ' Flag we need saving
                bNeed = True
              End If
            End If
          Next intJ
        End If
      Next intI
      ' Need saving?
      If (bNeed) Then
        ' Save the [tdlMorphdict]
        Status("Saving added prefix entries...")
        tdlMorphDict.WriteXml(strMorphDictFile)
        ' Show how many we added
        Logging("Prefix entries added: " & tdlMorphDict.Tables("Prefix").Rows.Count)
        ' Show we are ready
        Status("Ready")
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphPfxOEdict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphReadOEdict
  ' Goal:   Read the OE dictionary from the expected place; if not available: download from internet
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MorphReadOEdict(Optional ByVal strDictName As String = "OEdict_out.xml") As Boolean
    'Dim strDictName As String = "OEdict_out.xml"
    Dim strDictFile As String   'Location of OE dictionary
    Dim strDictUrl As String    ' URL of dictionary file

    Try
      ' Validate: Do not read the dictionary again!!!
      If (tdlOEdict IsNot Nothing) Then Status("English dictionary loaded...") : Return True
      ' Determine the location of the dictionary file
      strDictFile = GetSetDir() & "\" & strDictName
      If (Not IO.File.Exists(strDictFile)) Then
        ' Try to download it from internet
        strDictUrl = "http://erwinkomen.ruhosting.nl/software/Cesax/" & strDictName
        ' Try to download it
        Status("Trying to download " & strDictName & "...")
        My.Computer.Network.DownloadFile(strDictUrl, strDictFile, "", "", True, 5000, True)
        If (Not IO.File.Exists(strDictFile)) Then Return False
      End If
      ' Read the dictionary into a dataset
      If (Not ReadDataset("OEdict.xsd", strDictFile, tdlOEdict)) Then Return False
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphReadOEdict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphDictSfmToXml
  ' Goal:   Convert .sfm dictionary into .xml one
  ' History:
  ' 15-03-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MorphDictSfmToXml() As Boolean
    Dim strFileIn As String = "d:\data files\corpora\dictionaries\OEdict_input.sfm"
    Dim strFileAdp As String = "d:\data files\corpora\dictionaries\OEdict_adapt.sfm"
    Dim strFileOut As String = "d:\data files\corpora\dictionaries\OEdict_out.xml"
    Dim strLine As String     ' One line
    Dim strSfm As String      ' One sfm
    Dim strPos As String      ' Part-of-speech category
    Dim strLexeme As String   ' The lexeme we are treating
    Dim strText As String     ' Remainder
    Dim strTextBkup As String ' Backup of [strText]
    Dim strLv As String       ' lexical function vernacular value
    Dim strLv2 As String      ' Second one
    Dim arCf() As String      ' Array of cross-references
    Dim arLine() As String    ' Dictionary divided into lines
    Dim intI As Integer       ' Counter
    Dim intJ As Integer       ' Counter
    Dim intPtc As Integer     ' Percentage
    Dim intPos As Integer     ' Position
    Dim intPosBkup As Integer ' Backup of [intPos]
    Dim intStart As Integer   ' Start within word
    Dim intSense As Integer   ' Current sense number
    Dim intEntryId As Integer ' Current entry
    Dim bHasSense As Boolean  ' Does this entry contain \sn markers?
    Dim bHasCat As Boolean    ' Does this entry have a category (POS)?
    Dim bHasCapital As Boolean ' First letter is a capital letter
    Dim dtrNew As DataRow     ' Current datarow we are working on
    Dim dtrFound() As DataRow ' What we find
    Dim dtrSub() As DataRow   ' Belonging to current entry
    Dim dtrSense As DataRow   ' Sense
    Dim dtrLexF As DataRow    ' Lexical function/value pair
    Dim colLine As New StringColl ' Output collection
    Dim arVbInfo() As String = {"3rd pres", "past 3rd sing", "3rd sing.", "3rd sing", "past", "pret", "ptp is", "ptp"}
    Dim arVbMean() As String = {"3rd pres", "pastSg", "pastSg", "3rd sing", "3rd sing", "past", "past", "ptpIs", "ptp"}
    Dim arVbMeanPos() As String = {"VBP", "VBD", "VBD", "VBP", "VBP", "VBD", "VBD", "VAN", "VAN"}
    Dim arCatInfo() As String = {"comp adj", "cmp adj ", "spl adj ", "adv ", "n (-", "m (-", "f (-", "? (-", "pron ", "prep ", "adj "}
    Dim arCatMean() As String = {"Adj", "Adj", "Adj", "Adv", "N", "N", "N", "N", "Pro", "P", "Adj"}
    Dim arVbCfIn() As String = {"past 3rd sing of ", "past participle of ", "pret 3rd sing of ", "past 3rd sing of ", "pres 3rd sing of ", _
                                "ptp of ", "past of "}
    Dim arVbCfOut() As String = {"pastSg", "pastPtp", "past", "past", "3rd sing", _
                                 "ptp", "past"}
    Dim encThis As System.Text.Encoding

    Try
      ' Validate
      If (Not IO.File.Exists(strFileIn)) Then Return False
      ' Create a new dataset
      If (Not CreateDataSet("OEdict.xsd", tdlOEdict)) Then Return False
      ' Read the SFM dictionary
      Status("Reading SFM dictionary...")
      ' encThis = System.Text.Encoding.GetEncoding(1252)
      encThis = System.Text.Encoding.UTF8
      arLine = Split(IO.File.ReadAllText(strFileIn, encThis), vbCrLf) : dtrNew = Nothing : dtrSense = Nothing : strPos = "" : strLexeme = ""
      ' Parse the dictionary line by line
      For intI = 0 To arLine.Length - 1
        ' Where are we?
        intPtc = (intI + 1) * 100 \ arLine.Length
        Status("SfmToXml " & intPtc & "%", intPtc)
        ' Get this line
        strLine = MyTrim(arLine(intI))
        ' Skip empty lines
        If (strLine <> "") Then
          ' ========== DEBUG ===============
          ' If (InStr(strLine, "PastSg of") > 0) Then Stop
          ' ================================
          ' There should be an sfm part and something else
          If (Left(strLine, 1) <> "\") Then
            ' Add this to the current definition
            dtrSense.Item("Def") &= " " & strLine
          Else
            intPos = InStr(strLine, " ")
            If (intPos = 0) Then Stop
            strSfm = Left(strLine, intPos - 1) : intPosBkup = intPos
            strText = StringOEtoTagged(Mid(strLine, intPos + 1)) : strTextBkup = strText
            Select Case strSfm
              Case "\lex"
                ' Finish up the previous one
                If (dtrNew IsNot Nothing) Then
                  ' Does the previous entry have a POS?
                  If (dtrNew.Item("Pos").ToString = "") AndAlso (strPos <> "") Then
                    ' No it doesn't but there is a POS available --> assign it
                    dtrNew.Item("Pos") = strPos
                  End If
                End If
                ' Possibly strip off a numeral at the end
                If (IsNumeric(Right(strText, 1))) Then strText = Left(strText, strText.Length - 1)
                ' Check first letter capitalisation
                bHasCapital = System.Char.IsUpper(strText, 0)
                ' Debug.Print(System.Char.IsUpper(strText, 1))
                ' Get the lexeme
                strLexeme = LCase(strText)
                strLexeme = strLexeme.Replace("1", "")
                ' Start a new lexeme
                dtrNew = AddOneDataRow(tdlOEdict, "Entry", "EntryId", "EntryList")
                dtrNew.Item("l") = strLexeme
                intEntryId = dtrNew.Item("EntryId")
                ' ========== DEBUG ===============
                ' If (strText = "aet") Then Stop
                ' ================================
                ' Start afresh with sense number, category and POS
                intSense = 1 : strPos = ""
                bHasCat = False : bHasSense = False
              Case "\mp"  ' Morphemic Property (may be more than one)
                ' Determine what kind of type this is
                If (Left(strText, 1) = "s") Then
                  strText = "vtype=" & strText
                Else
                  strText = "comType=" & strText
                End If
                If (dtrNew.Item("f").ToString = "") Then
                  dtrNew.Item("f") = strText
                Else
                  dtrNew.Item("f") &= ";" & strText
                End If
              Case "\c"   ' Part of speech
                dtrNew.Item("Pos") = strText : strPos = strText
                ' Check for names
                If (strPos = "N" AndAlso bHasCapital) Then strPos = "NR"
                ' Set [HasCat]
                bHasCat = True
              Case "\prn"   ' Way to pronounce the utterance
                ' Remove left and right bracket
                strText = strText.Replace("[", "")
                strText = Trim(strText.Replace("]", ""))
                dtrNew.Item("Phn") = strText
              Case "\sn"    ' Sense number
                ' We have sense!
                bHasSense = True
                ' Now get the sense...
                intPos = InStr(strText, ".")
                If (intPos = 0) Then
                  Stop
                Else
                  intSense = CInt(Left(strText, intPos - 1))
                  strText = Trim(Mid(strText, intPos + 1))
                  ' Already make this sense
                  dtrSense = AddOneDataRowWithParent(tdlOEdict, "Sense", "SenseId", dtrNew)
                  dtrSense.Item("N") = intSense
                  dtrSense.Item("EntryId") = intEntryId
                  dtrSense.Item("l") = strLexeme
                  ' Set the text, if there is some
                  If (strText <> "") Then
                    dtrSense.Item("Def") = strText
                  End If
                  ' Accept the default category we have gathered so far
                  If (bHasCat) Then
                    ' Check for names
                    If (strPos = "N" AndAlso bHasCapital) Then strPos = "NR"
                    ' Adapt
                    dtrSense.Item("Pos") = strPos
                    dtrNew.Item("Pos") = strPos
                  End If
                End If
              Case "\lf"  ' Lexical function
                dtrLexF = AddOneDataRowWithParent(tdlOEdict, "LexFun", "LexFunId", dtrNew)
                With dtrLexF
                  ' Give it the correct entry id, lexical function and lexical value
                  .Item("EntryId") = intEntryId : .Item("lf") = strText
                  ' Check if the next line contains \lv
                  intPos = InStr(arLine(intI + 1), "\lv ")
                  If (intPos > 0) Then
                    ' Get the lv
                    intI += 1
                    strText = Mid(arLine(intI), intPos + 4)
                    .Item("lv") = strText
                  End If
                End With
              Case "\lv"  ' Lexical vernacular
                ' We cannot have a \lv without preceding \lf
                Stop
              Case "\gn"    ' Gender
                ' Add the gender as a feature
                strText = "gn=" & strText
                If (dtrNew.Item("f").ToString = "") Then
                  dtrNew.Item("f") = strText
                Else
                  dtrNew.Item("f") &= ";" & strText
                End If
              Case "\txt"   ' Definition
                ' Check if this is a variant needing special treatment...
                If (StringContains(strText, arVbCfIn, True)) Then
                  ' This line needs special treatment: take off the first part
                  For intJ = 0 To arVbCfIn.Length - 1
                    ' Is this the one?
                    If (InStr(strText, arVbCfIn(intJ)) = 1) Then
                      ' Okay, we have it: process it
                      dtrSense.Item("Pos") = "Vb/" & arVbCfOut(intJ)
                      strText = Trim(Mid(strText, arVbCfIn(intJ).Length + 1))
                      strText = Trim(strText.Replace(";", ""))
                      dtrSense.Item("Cf") = strText
                      ' We can exit the loop
                      Exit For
                    End If
                  Next intJ
                ElseIf (InStr(strText, "see ") = 1) Then
                  ' This is a cross-reference\
                  strText = Mid(strText, 5)
                  ' If (InStr(strText, ";") > 0) Then Stop
                  strText = Trim(strText.Replace(";", ""))
                  dtrSense.Item("Cf") = strText
                  ' Make sure we signal that all is "eaten"
                  strText = ""
                ElseIf (strPos = "Vb") OrElse (StringContains(strText, arVbInfo)) Then
                  ' Check if this is a verb, in which case the \txt field needs further scrutiny
                  ' Check if POS is defined
                  If (strPos = "") Then
                    ' POS is not defined, but this has the bearmarks of a verb, so make it a verb
                    strPos = "Vb" : bHasCat = True
                    ' Adjust the main entry's settings
                    dtrNew.Item("Pos") = strPos
                    ' Do we have a sense already?
                    If (bHasSense) Then
                      dtrSense.Item("Pos") = strPos
                    End If
                  End If
                  ' Field may need further scrutiny
                  For intJ = 0 To arVbInfo.Length - 1
                    intPos = InStr(strText, arVbInfo(intJ))
                    If (intPos > 0) Then
                      ' ========== DEBUGING ====
                      ' If (intJ = 3) OrElse (InStr(strTextBkup, "past 3rd") > 0) Then Stop
                      ' =========================
                      ' Get this lexical function out of it
                      strText = Left(strText, intPos - 1) & Mid(strText, intPos + 1 + arVbInfo(intJ).Length)
                      ' Eat spaces
                      'intPos += arVbInfo(intJ).Length + 1
                      While (InStr(" " & ";" & vbCrLf, Mid(strText, intPos, 1)) > 0) AndAlso (intPos < strText.Length)
                        intPos += 1
                      End While
                      ' Read a word
                      intStart = intPos
                      While (InStr(" " & ";" & vbCrLf, Mid(strText, intPos, 1)) = 0) AndAlso (intPos < strText.Length)
                        intPos += 1
                      End While
                      ' Get the word
                      strLv = Mid(strText, intStart, intPos - intStart)
                      ' Check
                      If (strLv = "") Then
                        ' ========== DEBUGING ====
                        ' If (arVbInfo(intJ) <> "pret") Then Stop
                        ' =========================
                        ' Reset the SFM and text parts and [intPos]
                        intPos = intPosBkup
                        strText = strTextBkup
                        ' and do not proceed!!
                      Else
                        ' Proceed with the LV
                        If (Right(strLv, 1) = ",") Then strLv = Left(strLv, strLv.Length - 1)
                        ' Adapt...
                        strText = Trim(Left(strText, intStart - 1) & Mid(strText, intPos))
                        ' Process the lex-vernacular
                        intPos = InStr(strLv, "/")
                        If (intPos > 0) Then
                          ' We need to process two versions
                          strLv2 = Left(strLv, intPos - 1)
                          dtrLexF = AddOneDataRowWithParent(tdlOEdict, "LexFun", "LexFunId", dtrNew)
                          With dtrLexF
                            ' Give it the correct entry id, lexical function and lexical value
                            .Item("EntryId") = intEntryId
                            .Item("lf") = IIf(arVbMean(intJ) = "past", "pastSg", arVbMean(intJ))
                            .Item("lv") = strLv2
                          End With
                          ' Derive the second one:
                          ' (1) strip off the /
                          strLv = Mid(strLv, intPos + 1)
                          ' Check if the first fits into the second
                          If (InStr(strLv2, strLv) = 0) AndAlso (Left(strLv2, 2) <> Left(strLv, 2)) Then
                            ' No fit --> probably suffix
                            If (strLv <> "on") Then
                              ' Stop
                            Else
                              strLv = strLv2 & strLv
                            End If
                          Else
                            ' First fits in second --> keep separate
                            'strLv = strLv2 & Mid(strLv, intPos + 1)
                          End If

                          dtrLexF = AddOneDataRowWithParent(tdlOEdict, "LexFun", "LexFunId", dtrNew)
                          With dtrLexF
                            ' Give it the correct entry id, lexical function and lexical value
                            .Item("EntryId") = intEntryId
                            .Item("lf") = IIf(arVbMean(intJ) = "past", "pastPl", arVbMean(intJ))
                            .Item("lv") = strLv
                          End With
                        Else
                          ' Only one version is needed
                          dtrLexF = AddOneDataRowWithParent(tdlOEdict, "LexFun", "LexFunId", dtrNew)
                          With dtrLexF
                            ' Give it the correct entry id, lexical function and lexical value
                            .Item("EntryId") = intEntryId : .Item("lf") = arVbMean(intJ) : .Item("lv") = strLv
                          End With
                        End If
                      End If
                    End If
                  Next intJ
                ElseIf (Not bHasCat) Then
                  ' Try to determine the category by looking for markers
                  For intJ = 0 To arCatInfo.Length - 1
                    ' Do we have a match?
                    If (Left(strText, arCatInfo(intJ).Length) = arCatInfo(intJ)) Then
                      ' DO NOT Adapt the text -- the info should stay in there
                      ' strText = Mid(strText, arCatInfo(intJ).Length + 1)
                      ' Does this go to the category or to the sense?
                      If (bHasSense) Then
                        ' Only give this sense a POS category
                        dtrSense.Item("Pos") = arCatMean(intJ)
                      Else
                        ' Determine the category
                        strPos = arCatMean(intJ)
                        ' Check for names
                        If (strPos = "N" AndAlso bHasCapital) Then strPos = "NR"
                        ' Adjust the main entry's settings
                        dtrNew.Item("Pos") = strPos
                        ' Exit gracefully
                        bHasCat = True
                        Exit For
                      End If
                      ' DO NOT Adapt the text -- the info should stay in there
                      ' strText = Trim(Mid(strText, arCatInfo(intJ).Length + 1))
                      ' We have the category, so now leave...
                      Exit For
                    End If
                  Next intJ
                  '' Do we now have a category?
                  'If (Not bHasCat) AndAlso (Left(strText, 4) = "see ") Then
                  '  ' Adapt the text
                  '  strText = Mid(strText, 5)
                  '  ' Make this a "conform" one
                  'End If
                End If
                ' Have we got any text left?
                If (strText <> "") Then
                  ' Check if this sense already exists
                  dtrFound = tdlOEdict.Tables("Sense").Select("EntryId=" & intEntryId & " AND N=" & intSense)
                  If (dtrFound.Length = 0) Then
                    ' Add a new sense level with this definition
                    dtrSense = AddOneDataRowWithParent(tdlOEdict, "Sense", "SenseId", dtrNew)
                    dtrSense.Item("N") = intSense
                    dtrSense.Item("EntryId") = intEntryId
                    dtrSense.Item("Def") = strText
                    dtrSense.Item("l") = strLexeme
                    ' The POS must have been supplied earlier
                    If (bHasCat) Then
                      ' Check for names
                      If (strPos = "N" AndAlso bHasCapital) Then strPos = "NR"
                      ' Adapt POS
                      dtrSense.Item("Pos") = strPos
                    End If
                  Else
                    ' Keep this sense
                    dtrSense = dtrFound(0)
                    ' Do we already have a definition?
                    If (dtrSense.Item("Def").ToString = "") Then
                      dtrSense.Item("Def") = strText
                    Else
                      dtrSense.Item("Def") &= " " & strText
                    End If
                  End If
                End If
              Case "\cf"    ' Cross reference
                ' Divide the cross-references
                If (InStr(strText, ";") > 0) Then
                  arCf = Split(strText, ";")
                Else
                  arCf = Split(strText, ",")
                End If
                ' Is this only one?
                If (arCf.Length = 1) Then
                  ' Also add it to the main entry
                  dtrNew.Item("Cf") = strText
                End If
                ' Add each one to a sense
                For intJ = 0 To arCf.Length - 1
                  ' Adapt this CF
                  strText = Trim(arCf(intJ))
                  If (strText <> "") Then
                    ' Add a sense with this CF
                    dtrSense = AddOneDataRowWithParent(tdlOEdict, "Sense", "SenseId", dtrNew)
                    dtrSense.Item("N") = intSense : intSense += 1
                    dtrSense.Item("EntryId") = intEntryId
                    dtrSense.Item("Cf") = strText
                    dtrSense.Item("l") = strLexeme
                  End If
                Next intJ
              Case "\_sh"
                ' Just skip
              Case Else
                Stop
            End Select
          End If
        End If
        ' Go to the next line
      Next intI
      ' Make sure everything is up to date
      tdlOEdict.AcceptChanges()
      ' Save the result
      tdlOEdict.WriteXml(strFileOut)
      ' Make adapted output dictionery in SFM
      dtrFound = tdlOEdict.Tables("Entry").Select("", "EntryId ASC")
      colLine.Add("\_sh v3.0  869  OEdict")
      For intI = 0 To dtrFound.Length - 1
        ' Access this element
        With dtrFound(intI)
          ' Start output
          colLine.Add(vbCrLf & "\lex " & .Item("l"))
          ' Do we have pronunciation?
          If (.Item("Phn").ToString <> "") Then colLine.Add("\prn " & .Item("Phn"))
          ' Hopefully POS? --> add category
          If (.Item("Pos").ToString <> "") Then colLine.Add("\c " & .Item("Pos"))
          ' Morphemic properties?
          If (.Item("f").ToString <> "") Then colLine.Add("\mp " & .Item("f"))
          ' Cross reference?
          If (.Item("Cf").ToString <> "") Then colLine.Add("\cf " & .Item("Cf"))
        End With
        ' Get any possible lexical functions
        dtrSub = tdlOEdict.Tables("LexFun").Select("EntryId=" & dtrFound(intI).Item("EntryId"))
        For intJ = 0 To dtrSub.Length - 1
          ' Add this lexical function
          With dtrSub(intJ)
            colLine.Add("\lf " & .Item("lf"))
            colLine.Add("\lv " & .Item("lv"))
          End With
        Next intJ
        ' Get any possible senses
        dtrSub = tdlOEdict.Tables("Sense").Select("EntryId=" & dtrFound(intI).Item("EntryId"))
        If (dtrSub.Length = 1) Then
          ' Just add this without sense number
          colLine.Add("\txt " & dtrSub(0).Item("Def"))
        ElseIf (dtrSub.Length > 0) Then
          For intJ = 0 To dtrSub.Length - 1
            ' Add this lexical function
            With dtrSub(intJ)
              colLine.Add("\sn " & .Item("N"))
              ' Do we have a POS?
              If (.Item("Pos").ToString <> "") Then
                colLine.Add("\ps " & .Item("Pos"))
              End If
              If (.Item("Def").ToString <> "") Then
                colLine.Add("\txt " & .Item("Def"))
              End If
              ' Do we have a cross reference?
              If (.Item("Cf").ToString <> "") Then
                ' Add the CF
                colLine.Add("\cf " & .Item("Cf"))
              End If
            End With
          Next intJ
        End If
      Next intI
      ' Save the adapted dictionary
      IO.File.WriteAllText(strFileAdp, colLine.Text, encThis)
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphDictSfmToXml error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   StringContains
  ' Goal:   Check if the string [strText] contains an item from [arThis]
  ' History:
  ' 08-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function StringContains(ByRef strText As String, ByRef arThis() As String, Optional ByVal bInitial As Boolean = False) As Boolean
    Dim intI As Integer ' Counter
    Dim intPos As Integer ' Position

    Try
      ' Validate
      If (strText = "") Then Return False
      ' Check all possibilities
      For intI = 0 To arThis.Length - 1
        ' Check this item
        intPos = InStr(strText, arThis(intI))
        If (bInitial) Then
          ' Must be initial
          If (intPos = 1) Then Return True
        Else
          ' May be anywhere
          If (intPos > 0) Then Return True
        End If
      Next intI
      ' Default: return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/StringContains error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   AskOEdictEntry
  ' Goal:   Try to retrieve an entry [strIn] from the OE dictionary, given the POS
  '         If it is there, then return positively, and put an adapted lemma in [strLemma] if needed
  ' History:
  ' 08-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function AskOEdictEntry(ByVal strIn As String, ByVal strPos As String, ByVal strFile As String, ByVal intForestId As Integer, _
                                 ByVal intEtreeId As Integer, ByRef strLemma As String) As Boolean
    Dim ndxThis As XmlNode  ' Node of me
    Dim arFiles() As String ' Result of looking for file
    Dim strSrchDir As String = "d:\data files\corpora\english\xml\adapted"

    Try
      ' See if we already have the file
      If (InStr(strCurrentFile, strFile) = 0) Then
        ' Open the correct file
        arFiles = IO.Directory.GetFiles(strSrchDir, strFile & ".psdx", IO.SearchOption.AllDirectories)
        If (arFiles.Length > 0) Then
          ' Set current file
          strCurrentFile = arFiles(0)
          ' Load the file properly
          pdxCurrentFile = Nothing
          If (Not ReadXmlDoc(strCurrentFile, pdxCurrentFile)) Then Return False
          ' Go to the correct node
          ndxThis = pdxCurrentFile.SelectSingleNode("//eTree[@Id=" & intEtreeId & "]")
          ' We really want to know the verbal lemma's, so ask the user
          If (MorphLemmaAsk(ndxThis, strIn, strPos, strPos, strLemma)) Then
            ' Return the lemma
            Return True
          End If
        End If
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/AskOEdictEntry error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetOEdictEntryAlt
  ' Goal:   Try to retrieve an entry [strIn] from the OE dictionary, given the POS
  '         The entry may not be equal to [strLemma], and its SimpleDistance should be equal to or larger than [intCost]
  ' Return: True  - if there IS an alternative
  '         False - if there is NO alternative
  ' History:
  ' 30-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function GetOEdictEntryAlt(ByVal strIn As String, ByVal strLemma As String, ByVal strPos As String, ByVal intCost As Integer, _
                                    ByRef strAlt As String) As Boolean
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim strPosDict As String      ' POS as defined in dictionary
    Dim strPosNear As String = "" ' Near POS
    Dim strWord As String         ' Current possibility
    Dim strPath As String = ""    ' Detailed substitution path
    Dim intSimD As Integer        ' Simple distance
    Dim intSimMin As Integer      ' Minimum distance
    Dim intNum As Integer = 0     ' Total number of insertions, deletions and substitutions
    Dim intI As Integer           ' Counter

    Try
      ' Determine the dictionary's POS
      strPosDict = MorphGetPosDict(strPos, strPosNear)
      ' Need to load dictionary?
      If (tdlOEdict Is Nothing) Then
        ' Read dictionary...
        If (Not MorphReadOEdict()) Then Status("Could not read the OE dictionary database") : Return False
      End If
      ' Try find entry/entries
      dtrFound = tdlOEdict.Tables("Sense").Select("l='" & strIn & "' AND Pos='" & strPosDict & "'")
      ' Any results?
      If (dtrFound.Length > 0) Then
        ' Evaluate results
        If (dtrFound.Length = 1) Then
          ' We have the one and only result
          Return False
        Else
          ' There's more than one possibility, so there is an alternative lemma
          strAlt = dtrFound(1).Item("l") : Return True
        End If
      End If
      ' Go through POS-possible results in dictionary stepwise...
      dtrFound = tdlOEdict.Tables("Sense").Select("Pos='" & strPosDict & "'")
      intSimMin = intCost
      ' Walk the results
      For intI = 0 To dtrFound.Length - 1
        ' Get current possibility
        strWord = dtrFound(intI).Item("l")
        ' Get simple distance
        intSimD = SimpleDistance(strIn, strWord, strPath, intNum, intSimMin)
        ' Check the result
        If (strWord <> strLemma) AndAlso (intSimD <= intCost) Then
          ' There is at least one lemma with the same distance or even nearer by
          strAlt = strWord : Return True
        End If
      Next intI
      ' No lemma found that is nearer by
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/GetOEdictEntryAlt error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetOEdictEntrySimi
  ' Goal:   Try to retrieve an entry [strIn] from the OE dictionary, given the POS
  '         The entry may not be equal to [strLemma], and its Similarity should be equal to or larger than [intValue]
  ' Return: True  - if there IS an alternative
  '         False - if there is NO alternative
  ' History:
  ' 30-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function GetOEdictEntrySimi(ByVal strIn As String, ByVal strLemma As String, ByVal strPos As String, _
                                     ByVal intSval As Integer, ByRef strAlt As String) As Boolean
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim strPosDict As String      ' POS as defined in dictionary
    Dim strPosNear As String = "" ' Near POS
    Dim strWord As String         ' Current possibility
    Dim strPath As String = ""    ' Detailed substitution path
    Dim intSimD As Integer        ' Simple distance
    Dim intNum As Integer = 0     ' Total number of insertions, deletions and substitutions
    Dim intI As Integer           ' Counter

    Try
      ' Determine the dictionary's POS
      strPosDict = MorphGetPosDict(strPos, strPosNear)
      ' Need to load dictionary?
      If (tdlOEdict Is Nothing) Then
        ' Read dictionary...
        If (Not MorphReadOEdict()) Then Status("Could not read the OE dictionary database") : Return False
      End If
      ' Try find entry/entries
      dtrFound = tdlOEdict.Tables("Sense").Select("l='" & strIn & "' AND Pos='" & strPosDict & "'")
      ' Any results?
      If (dtrFound.Length > 0) Then
        ' Evaluate results
        If (dtrFound.Length = 1) Then
          ' We have the one and only result
          Return False
        Else
          ' There's more than one possibility, so there is an alternative lemma
          strAlt = dtrFound(1).Item("l") : Return True
        End If
      End If
      ' Go through POS-possible results in dictionary stepwise...
      dtrFound = tdlOEdict.Tables("Sense").Select("Pos='" & strPosDict & "'")
      ' Walk the results
      For intI = 0 To dtrFound.Length - 1
        ' Get current possibility
        strWord = dtrFound(intI).Item("l")
        ' Get similarity
        intSimD = objSim.CompareStrings(strIn, strWord)
        ' Check the result
        If (strWord <> strLemma) AndAlso (intSimD >= intSval) Then
          ' There is at least one lemma with the same distance or even nearer by
          strAlt = strWord : Return True
        End If
      Next intI
      ' No lemma found that is nearer by
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/GetOEdictEntrySimi error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetOEdictEntry
  ' Goal:   Try to retrieve an entry [strIn] from the OE dictionary, given the POS
  '         If it is there, then return positively, and put an adapted lemma in [strLemma] if needed
  ' History:
  ' 08-04-2013  ERK Created
  ' 16-05-2013  ERK Added [PosNear]
  ' ------------------------------------------------------------------------------------------------------------
  Public Function GetOEdictEntry(ByVal strIn As String, ByVal strPosDict As String, ByVal strPos As String, ByVal strPosNear As String, _
                                 ByRef strLemma As String, ByRef strLev As String, Optional ByVal strLoc As String = "", _
                                 Optional ByVal strClause As String = "") As Boolean
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim dtrCat() As DataRow       ' Result of SELECT
    Dim strWord As String         ' Current possibility
    Dim strInRew As String        ' Rewrite of word
    Dim strSimiLine As String     ' File with prospective similarities in tab format
    Dim strPath As String = ""    ' Detailed substitution path
    Dim strPosPrev As String = "" ' Previous POS
    Dim strPre As String = ""     ' Possible prefix before [strIn]
    Dim colRewr As New StringColl ' Possible rewrites for me
    Dim colLev As New StringColl  ' Levenshtein collection
    Dim colSim As New StringColl  ' Simple distance collection
    Dim colBest As New StringColl ' Best results after measuring distance
    ' Dim intLevD As Integer       ' Levenshtein distance (weighted)
    ' Dim intLevMin As Integer     ' Minimum distance
    ' Dim intLevIdx As Integer     ' Index of the best one
    Dim intValue As Integer       ' Similarity measure
    Dim intValMax As Integer      ' Highest similarity value
    Dim intEquals As Integer      ' Number of equals
    Dim strWordMax As String      ' Best matching lemma
    Dim strSfxType As String      ' Whether closest word uses a suffix rewrite rule
    ' Dim intSimD As Integer       ' Simple distance
    ' Dim intSimMin As Integer      ' Minimum distance
    ' Dim intSimIdx As Integer     ' Index of the best
    Dim intNum As Integer = 0     ' Total number of insertions, deletions and substitutions
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim bFound As Boolean         ' Found flag

    Try
      '' Trial: just process those that have a [PosNear]
      'If (strPosNear = "") Then Return False
      ' Initialize level
      strLev = "1" : If (strLoc = "") Then strPre = "**"
      ' Try find entry/entries
      dtrFound = tdlOEdict.Tables("Sense").Select("l='" & strIn & "' AND Pos='" & strPosDict & "'")
      ' Any results?
      If (dtrFound.Length > 0) Then
        ' Evaluate results
        If (dtrFound.Length = 1) Then
          ' We have the one and only result
          strLemma = strIn : Return True
        Else
          ' There's more than one possibility -- but the lemma's are apparently the same, and that is what counts
          strLemma = strIn : Return True
        End If
      End If
      ' Try looking for entries that have no POS marked, but do have "conform"
      dtrFound = tdlOEdict.Tables("Sense").Select("l='" & strIn & "' AND Cf<>''")
      If (dtrFound.Length > 0) Then
        ' Find the "CF" entry
        strWord = dtrFound(0).Item("Cf").ToString
        dtrFound = tdlOEdict.Tables("Sense").Select("l='" & strWord & "'")
        ' Check the result
        If (dtrFound.Length > 0) Then
          ' Does this have the correct POS?
          If (dtrFound(0).Item("Pos").ToString = strPosDict) Then
            ' Accept the lemma
            strLemma = dtrFound(0).Item("l").ToString
            Return True
          ElseIf (dtrFound(0).Item("Cf").ToString <> "''") Then
            ' Try going one more level deep
            strWord = dtrFound(0).Item("Cf").ToString
            dtrFound = tdlOEdict.Tables("Sense").Select("l='" & strWord & "'")
            If (dtrFound.Length > 0) Then
              ' Does this have the correct POS?
              If (dtrFound(0).Item("Pos").ToString = strPosDict) Then
                ' Accept the lemma
                strLemma = dtrFound(0).Item("l").ToString
                Return True
              End If
            End If
          End If
        End If
      End If
      ' Third attempt: Get all possible prefix and suffix-rewrites for [strIn]
      If (Not MorphRewrites(strIn, strPos, colRewr)) Then Return False
      ' Try them all
      For intI = 0 To colRewr.Count - 1
        ' Check this rewrite
        dtrFound = tdlOEdict.Tables("Sense").Select("l='" & colRewr.Item(intI) & "' AND Pos='" & strPosDict & "'")
        If (dtrFound.Length > 0) Then
          ' Found it!
          strLemma = colRewr.Item(intI)
          Return True
        End If
      Next intI
      ' Initialize
      strSfxType = "" : colBest.Clear() : intEquals = 0
      ' Fourth attempt: go through POS-possible results in dictionary stepwise...
      If (strPosNear = "") Then
        dtrFound = tdlOEdict.Tables("Sense").Select("Pos='" & strPosDict & "'")
        intValMax = 0 : strWordMax = ""
        ' Get a suffix rewrite
        If (Not MorphSuffixRewrite(strIn, strPos, colRewr)) Then Return False
        ' Walk the results
        For intI = 0 To dtrFound.Length - 1
          ' Get current possible lemma
          strWord = dtrFound(intI).Item("l")
          ' ============ Double check =============
          ' If (strWord = "forbeodendlic" OrElse strWord = "forsewenlic") Then Stop
          ' =======================================
          ' Get the dictionary value with the initially highest similarity
          intValue = objSim.LevenDist(strIn, strWord, intEquals)
          ' Check if this one has a higher similarity so far
          If (intValue > intValMax) Then
            intValMax = intValue : strWordMax = strWord : strSfxType = ""
            colBest.Clear() : colBest.AddUnique(strWord, strSfxType & ";" & intEquals)
          End If
          ' Check the other variants
          For intJ = 0 To colRewr.Count - 1
            ' Get this variant for the word
            strInRew = colRewr.Item(intJ)
            ' Check this one's similarity measure
            intValue = objSim.LevenDist(strInRew, strWord, intEquals)
            ' Check if this surpasses the similarity we have right now
            If (intValue > intValMax) AndAlso (intValue > 0) Then
              ' Take this value + word as the best guess
              intValMax = intValue : strWordMax = strWord : strSfxType = strInRew
              ' Reset the collectino of best ones
              colBest.Clear() : colBest.AddUnique(strWord, strSfxType & ";" & intEquals)
              ' ============ Double check =============
              'If (intValMax = 100) Then Stop
              ' =======================================
            ElseIf (intValue = intValMax) AndAlso (intValue > 0) Then
              ' Add to the collection of best ones
              colBest.AddUnique(strWord, strSfxType)
            End If
          Next intJ
        Next intI
        ' Similarity acceptance depends on length
        ' If (intValMax >= GetSimiThreshold(strIn)) Then
        ' ============ DEBUG =========
        ' If (InStr(strIn, "brother") > 0) Then Stop
        ' ============================
        If (intValMax >= 97) Then
          ' Check the number of matches
          Select Case colBest.Count
            Case 1
              Debug.Print("Best match to [" & strPre & strIn & "] is: " & strWordMax & " (simi=" & intValMax & "%)")
              strLemma = strWordMax
              Return True
            Case Else
              ' Ask user
              Stop
          End Select
        ElseIf (intValMax >= 75) Then
          ' Look out for weird cases
          ' If (strLoc = "") Then Stop
          ' Action depends on the number of matches
          Select Case colBest.Count
            Case 1
              ' Record what we have found in tabular form, so that the user could still make use of it
              strSimiLine = strPre & strIn & ";" & strPos & ";" & strPosDict & ";" & intValMax & ";" & strSfxType & ";" & _
                                strWordMax & ";" & "" & ";" & "[" & strLoc & "]" & ";" & strClause
              mrp_colSimi.AddUnique(strSimiLine)
              ' Add the line to the revision output
              frmMain.tbGenRevision.AppendText(strSimiLine & vbCrLf)
            Case Else
              ' Now what? Add all the items to the output collection??
              For intI = 0 To colBest.Count - 1
                ' Add this one
                strSimiLine = strPre & strIn & ";" & strPos & ";" & strPosDict & ";" & intValMax & ";" & intI + 1 & ": " & colBest.Exmp(intI) & ";" & _
                           colBest.Item(intI) & ";" & "" & ";" & "[" & strLoc & "]" & ";" & strClause
                mrp_colMult.AddUnique(strSimiLine)
              Next intI
              ' Determine the number of equals
              ' Stop
          End Select
        Else
          ' We did not find a good match. Try something else: find a match in the set of those with my derived POS
          dtrCat = tdlMorphDict.Tables("Cat").Select("Pos='" & strPos & "'") : bFound = False
          If (dtrCat.Length > 0) Then
            ' Check if this is a derived type
            'If (dtrCat(0).Item("Type").ToString = "Derived") Then
            If (True) Then
              If (strPos <> strPosPrev) Then
                strPosPrev = strPos
                ' Select all entries with the exact same POS in table [VernPos]
                dtrFound = tdlMorphDict.Tables("VernPos").Select("Pos='" & strPos & "' AND Lev <> '2'", "Vern ASC")
              End If
              ' Check all the entries
              strWord = ""
              For intI = 0 To dtrFound.Length - 1
                ' Treat this entry?
                If (strWord <> dtrFound(intI).Item("Vern").ToString) Then
                  ' Okay treat it
                  strWord = dtrFound(intI).Item("Vern")
                  ' Get the dictionary value with the initially highest similarity
                  intValue = objSim.LevenDist(strIn, strWord, intEquals)
                  ' Check if this one has a higher similarity so far
                  If (intValue > intValMax) Then
                    intValMax = intValue : strWordMax = dtrFound(intI).Item("l")
                    strSfxType = "VernPos[" & strWord & "]"
                    colBest.Clear() : colBest.AddUnique(strWordMax, strSfxType & ";" & intEquals)
                    bFound = True
                  End If
                End If
              Next intI
            End If
          End If
          ' ========= Debugging ====
          ' If (Left(strLoc, 2) <> "co") Then Stop
          ' =========================
          ' Found anything?
          If (bFound) AndAlso (intValMax >= 75) Then
            ' What we now supply is second level
            strLev = "2"
            ' Action depends on valmax
            If (intValMax >= 97) Then
              ' Check the number of matches
              Select Case colBest.Count
                Case 1
                  Debug.Print("Best match to [" & strPre & strIn & "] is: " & strWordMax & " (simi=" & intValMax & "%)")
                  strLemma = strWordMax
                  Return True
                Case Else
                  ' Ask user
                  Stop
              End Select
            ElseIf (intValMax >= 75) Then
              ' Look out for weird cases
              ' If (strLoc = "") Then Stop
              ' Action depends on the number of matches
              Select Case colBest.Count
                Case 1
                  ' Record what we have found in tabular form, so that the user could still make use of it
                  strSimiLine = strPre & strIn & ";" & strPos & ";" & strPosDict & ";" & intValMax & ";" & strSfxType & ";" & _
                                    strWordMax & ";" & "" & ";" & "[" & strLoc & "]" & ";" & strClause
                  mrp_colSimi.AddUnique(strSimiLine)
                  ' Add the line to the revision output
                  frmMain.tbGenRevision.AppendText(strSimiLine & vbCrLf)
                Case Else
                  ' Now what? Add all the items to the output collection??
                  For intI = 0 To colBest.Count - 1
                    ' Add this one
                    strSimiLine = strPre & strIn & ";" & strPos & ";" & strPosDict & ";" & intValMax & ";" & intI + 1 & ": " & colBest.Exmp(intI) & ";" & _
                               colBest.Item(intI) & ";" & "" & ";" & "[" & strLoc & "]" & ";" & strClause
                    mrp_colMult.AddUnique(strSimiLine)
                  Next intI
                  ' Determine the number of equals
                  ' Stop
              End Select
            End If
          Else
            ' Put this one in a separate file
            strSimiLine = strPre & strIn & ";" & strPos & ";" & strPosDict & ";" & intValMax & ";" & strSfxType & ";" & _
                              strWordMax & ";" & "" & ";" & "[" & strLoc & "]" & ";" & strClause
            mrp_colHand.AddUnique(strSimiLine)
          End If
        End If
      End If

      '' fifth attempt: go through POS-possible results in dictionary stepwise...
      ''   18/may: calculate the Levenshtein distance (weighted)
      'dtrFound = tdlOEdict.Tables("Sense").Select("Pos='" & strPosDict & "'")
      'colLev.Clear() : intSimMin = 500
      '' Walk the results
      'For intI = 0 To dtrFound.Length - 1
      '  ' Get current possibility
      '  strWord = dtrFound(intI).Item("l")
      '  ' Compare in several different spellingvariants
      '  If (OEtaggedWordsLike(strIn, strWord)) Then
      '    ' This should be the correct one
      '    strLemma = strWord
      '    Return True
      '  End If
      '  ' See if there is a suffix-rewrite rule that helps for this item
      '  ' Note: use the 'real' POS
      '  If (SfxRewMatch(strIn, strWord, strPos)) Then
      '    ' This should be the correct one
      '    strLemma = strWord
      '    Return True
      '  End If
      '  ' See if there is a derive-rewrite rule that helps for this item
      '  ' Note: use the 'real' POS
      '  If (DeriveRewMatch(strIn, strWord, strPos, False)) Then
      '    ' This should be the correct one
      '    strLemma = strWord
      '    Return True
      '  End If
      '  ' See if there is a derive-near rule that helps for this item
      '  ' This is where we see if we can find a match for [PosNear]
      '  If (strPosNear <> "") AndAlso (DeriveRewMatch(strIn, strWord, strPosNear, True)) Then
      '    ' The vernacular word can be re-written as [strWord] with POS = [strPosNear]
      '    ' Can we find a lemma entry for this?
      '    If (GetOEdictEntry(strWord, strPosDict, strPosNear, "", strLemma)) Then
      '      ' Okay, this should be the lemma, then!
      '      Return True
      '    End If
      '  End If
      'Next intI
      '' Check what the best Levenshtein match was
      ''Debug.Print("Best match to [" & strIn & "] is: " & colLev.Exmp(intLevIdx) & " (dist=" & intLevMin & ")" & _
      ''             ";" & "simple: " & colSim.Exmp(intSimIdx) & " (dist=" & intSimMin & ")")
      'Debug.Print("Best match to [" & strIn & "] is: " &  colSim.Exmp(intSimIdx) & " (dist=" & intSimMin & ")")
      ' Return without result
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/GetOEdictEntry error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetSimiThreshold
  ' Goal:   Get a similarity threshold, depending on the length of the input string
  ' History:
  ' 03-06-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function GetSimiThreshold(ByVal strIn As String) As Integer
    Try
      Select Case strIn.Length
        Case Is > 10
          Return 83
        Case 10
          Return 86
        Case 9
          Return 87
        Case 8
          Return 88
        Case 7
          Return 89
        Case Else
          ' Default threshold is 90
          Return 90
      End Select
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/GetSimiThreshold error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return 100
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetOEdictLexfun
  ' Goal:   Try to retrieve an entry [strIn] from the OE dictionary, given the POS
  '         If it is there, then return positively, and put an adapted lemma in [strLemma] if needed
  ' History:
  ' 08-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function GetOEdictLexfun(ByVal strIn As String, ByVal strPosDict As String, ByVal strPos As String, ByRef strLemma As String) As Boolean
    Dim dtrFound() As DataRow   ' Result of SELECT
    'Dim dtrEnt() As DataRow     ' Looking for the entry
    Dim strWord As String       ' One word
    Dim arLf() As String        ' Array of LF entries
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter
    'Dim intEntryId As Integer   ' ID of entry

    Try
      ' Get entries
      arLf = Split(strPosDict, "|") : dtrFound = Nothing
      ' Try find entry/entries
      For intI = 0 To arLf.Length - 1
        ' Get anything from this one?
        dtrFound = tdlOEdict.Tables("Lexfun").Select("lv='" & strIn & "' AND lf='" & arLf(intI) & "'")
        If (dtrFound.Length > 0) Then Exit For
      Next intI
      ' Any results?
      If (dtrFound.Length > 0) Then
        If (DoOEdictLexFun(dtrFound, strLemma)) Then Return True
        '' Evaluate results
        'If (dtrFound.Length = 1) Then
        '  ' We have the one and only result, but now we still need to get the lemma...
        '  intEntryId = dtrFound(0).Item("EntryId")
        '  dtrEnt = tdlOEdict.Tables("Entry").Select("EntryId=" & intEntryId)
        '  If (dtrEnt.Length = 1) Then
        '    strLemma = dtrEnt(0).Item("l")
        '    Return True
        '  End If
        'Else
        '  ' If we come here, it means that there is ambiguity...
        '  ' Stop
        '  ' Check if this would lead to a different lemma
        '  ' (1) first get the first lemma
        '  intEntryId = dtrFound(0).Item("EntryId")
        '  dtrEnt = tdlOEdict.Tables("Entry").Select("EntryId=" & intEntryId)
        '  If (dtrEnt.Length = 1) Then
        '    strLemma = dtrEnt(0).Item("l")
        '  Else
        '    ' Wasting time
        '    Return False
        '  End If
        '  ' (2) compare this with the others
        '  For intI = 1 To dtrFound.Length - 1
        '    intEntryId = dtrFound(intI).Item("EntryId")
        '    dtrEnt = tdlOEdict.Tables("Entry").Select("EntryId=" & intEntryId)
        '    If (dtrEnt.Length = 1) Then
        '      If (strLemma <> dtrEnt(0).Item("l")) Then
        '        ' Lemma's are not the same
        '        Return False
        '      End If
        '    Else
        '      ' Wasting time
        '      Return False
        '    End If
        '  Next intI
        '  ' The lemma's are the same, so we're fine
        '  Return True
        'End If
      End If
      ' Third attempt: go through POS-possible results stepwise...
      For intI = 0 To arLf.Length - 1
        ' Get anything from this one?
        dtrFound = tdlOEdict.Tables("Lexfun").Select("lf='" & arLf(intI) & "'")
        ' Walk the results
        For intJ = 0 To dtrFound.Length - 1
          ' Get current possibility
          strWord = dtrFound(intJ).Item("lv")
          ' Compare in several different spellingvariants
          If (OEtaggedWordsLike(strIn, strWord)) Then
            ' This should be the correct one: try to get the lemma
            If (DoOEdictLexFun(dtrFound, strLemma, intJ)) Then Return True
          End If
          ' See if there is a suffix-rewrite rule that helps for this item
          ' Note: use the REAL p-o-s
          If (SfxRewMatch(strIn, strWord, strPos)) Then
            ' This should be the correct one: try to get the lemma
            If (DoOEdictLexFun(dtrFound, strLemma, intJ)) Then Return True
          End If
        Next intJ
      Next intI
      ' Return without result
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/GetOEdictLexfun error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoOEdictLexFun
  ' Goal:   Process the results found in GetOEdictLexFun
  ' History:
  ' 22-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function DoOEdictLexFun(ByRef dtrFound() As DataRow, ByRef strLemma As String, Optional ByVal intIdx As Integer = -1) As Boolean
    Dim dtrEnt() As DataRow     ' Result of SELECT
    Dim intEntryId As Integer   ' ID of the entry

    Try
      ' Evaluate results
      If (dtrFound.Length = 1) OrElse (intIdx >= 0) Then
        ' We have the one and only result, but now we still need to get the lemma...
        If (intIdx < 0) Then intIdx = 0
        intEntryId = dtrFound(intIdx).Item("EntryId")
        dtrEnt = tdlOEdict.Tables("Entry").Select("EntryId=" & intEntryId)
        If (dtrEnt.Length = 1) Then
          strLemma = dtrEnt(0).Item("l")
          Return True
        End If
      Else
        ' If we come here, it means that there is ambiguity...
        ' Stop
        ' Check if this would lead to a different lemma
        ' (1) first get the first lemma
        intEntryId = dtrFound(0).Item("EntryId")
        dtrEnt = tdlOEdict.Tables("Entry").Select("EntryId=" & intEntryId)
        If (dtrEnt.Length = 1) Then
          strLemma = dtrEnt(0).Item("l")
        Else
          ' Wasting time
          Return False
        End If
        ' (2) compare this with the others
        For intI = 1 To dtrFound.Length - 1
          intEntryId = dtrFound(intI).Item("EntryId")
          dtrEnt = tdlOEdict.Tables("Entry").Select("EntryId=" & intEntryId)
          If (dtrEnt.Length = 1) Then
            If (strLemma <> dtrEnt(0).Item("l")) Then
              ' Lemma's are not the same
              Return False
            End If
          Else
            ' Wasting time
            Return False
          End If
        Next intI
        ' The lemma's are the same, so we're fine
        Return True
      End If
      ' Return without result
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/DoOEdictLexFun error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphCatAdd
  ' Goal:   Add the [strPos] into the [Cat] table
  ' History:
  ' 13-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphCatAdd(ByVal strPos As String) As Boolean
    Dim dtrFound() As DataRow   ' Result of SELECT
    Dim dtrNew As DataRow       ' New datarow
    Dim strHead As String = ""  ' Head of this category
    Dim strDef As String = ""   ' Default category
    Dim strNear As String = ""  ' Near POS
    Dim intPos As Integer       ' Position in string

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Look for the entry
      dtrFound = tdlMorphDict.Tables("Cat").Select("Pos='" & strPos.Replace("'", "''") & "'")
      If (dtrFound.Length > 0) Then
        ' It is already there -- get it
        dtrNew = dtrFound(0)
        ' Increment the frequency
        dtrNew.Item("Freq") += 1
      Else
        ' It is not yet there, so create it
        dtrNew = AddOneDataRow(tdlMorphDict, "Cat", "CatId", "CatList")
        With dtrNew
          .Item("Pos") = strPos : .Item("Freq") = 1
          ' Determine the default
          If (DoLike(strPos, "*2[12]|*3[123]|*4[1234]")) Then
            ' Do not change the definition
            strDef = "-"
          ElseIf (Right(strPos, 1) = "$") Then
            strDef = Left(strPos, strPos.Length - 1)
          ElseIf (Right(strPos, 1) = "S") Then
            strDef = Left(strPos, strPos.Length - 1)
          ElseIf (DoLike(strPos, "DO*|DA*")) And (strPos.Length > 2) Then
            strDef = "DO"
          ElseIf (Left(strPos, 1) = "V") And (strPos.Length > 2) Then
            strDef = "VB"
          ElseIf (Left(strPos, 4) = "RP+V") And (strPos.Length > 5) Then
            strDef = "RP+VB"
          ElseIf (Left(strPos, 5) = "NEG+V") And (strPos.Length > 6) Then
            strDef = "NEG+VB"
          ElseIf (Left(strPos, 1) = "B") And (strPos.Length > 2) Then
            strDef = "BE"
          ElseIf (Left(strPos, 4) = "RP+B") And (strPos.Length > 5) Then
            strDef = "RP+BE"
          ElseIf (Left(strPos, 5) = "NEG+B") And (strPos.Length > 6) Then
            strDef = "NEG+BE"
          ElseIf (Left(strPos, 2) = "MD") And (strPos.Length > 2) Then
            strDef = "MD"
          ElseIf (Left(strPos, 4) = "RP+M") And (strPos.Length > 5) Then
            strDef = "RP+MD"
          ElseIf (Left(strPos, 5) = "NEG+M") And (strPos.Length > 6) Then
            strDef = "NEG+MD"
          ElseIf (Left(strPos, 2) = "AX") And (strPos.Length > 2) Then
            strDef = "AX"
          ElseIf (Left(strPos, 5) = "RP+AX") And (strPos.Length > 5) Then
            strDef = "RP+AX"
          ElseIf (Left(strPos, 6) = "NEG+AX") And (strPos.Length > 6) Then
            strDef = "NEG+AX"
          ElseIf (Left(strPos, 1) = "H") And (strPos.Length > 2) Then
            strDef = "HV"
          ElseIf (Left(strPos, 4) = "RP+H") And (strPos.Length > 5) Then
            strDef = "RP+HV"
          ElseIf (Left(strPos, 5) = "NEG+H") And (strPos.Length > 6) Then
            strDef = "NEG+HV"
          ElseIf (Left(strPos, 3) = "ADJ") And (strPos.Length > 3) Then
            strDef = "ADJ"
          ElseIf (Left(strPos, 3) = "ADV") And (strPos.Length > 3) Then
            strDef = "ADV"
          ElseIf (InStr(strPos, "^") > 0) Then
            strDef = Left(strPos, InStr(strPos, "^") - 1)
          ElseIf (InStr(strPos, "-") > 0) Then
            intPos = InStrRev(strPos, "-")
            If (IsNumeric(Mid(strPos, intPos + 1))) Then
              strDef = Left(strPos, intPos - 1)
            End If
          Else
            strDef = "-"
          End If
          ' Ask what the head category is of this item
          strHead = InputBox("What is the head of [" & strPos & "]?", "Category processing", strDef)
          .Item("Head") = strHead
          .Item("Type") = IIf(strHead = "-", "Main", "Derived")
          ' Check existence of ^ in the POS
          intPos = InStr(strPos, "^")
          If (intPos > 0) Then
            ' Break up the POS in two parts
            strNear = Left(strPos, intPos - 1)
            ' Check if [Near] is the same as [Head]
            If (strNear = strHead) Then
              strNear = ""
            End If
            ' Add a [Near]
            .Item("Near") = strNear
          End If
        End With
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MorphCatAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SandeepAlign
  ' Goal:   Align word [strS] and word [strT]
  ' Source: web > arxiv.org/ftp/arxiv/papers/1210/1210.8398.pdf
  ' Return: Similarity measure between 0 and 100
  ' Extensension:
  '         - If [intMax] is larger than 0, than don't calculate distances larger than [intMax]
  ' History:
  ' 01-06-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function SandeepAlign(ByRef strS_in As String, ByRef strT_in As String) As Integer
    Dim strS As String = MorphSeqToNum(strS_in, True)
    Dim strT As String = MorphSeqToNum(strT_in, True)
    Dim intM As Integer = strS.Length
    Dim intN As Integer = strT.Length
    Dim intMatch As Integer = 0   ' Number of matches
    Dim intK As Integer = 0       ' Counter
    Dim intCount1 As Integer = 0  ' Counter
    Dim intCount2 As Integer = 0  ' Counter

    Try
      ' Check for the obvious
      If (strS_in = strT_in) Then Return 100
      If (strS_in = "" OrElse strT_in = "") Then Return 0
      ' Loop through
      While (intN - intK >= 1)
        While (intCount2 + (intN - intK) <= intN)
          If (intCount2 + (intN - intK) <= intN) Then
            intCount1 = 0
          Else
            While (intCount1 < intM - (intN - intK))
              If (Mid(strT, intCount2, intN - intK) = Mid(strS, intCount1, intN - intK)) Then
                ' MARK[count1,count1+(n-k))
                ' // Record in S, where match has occurred.
                ' MARK[count2,count2+(n-k))
                ' // Record in V, where match has occurred.
              End If
              intCount1 += 1
            End While
          End If
          intCount2 += 1
        End While
        intCount1 = 0 : intCount2 = 0
        intK += 1
      End While

    Catch ex As Exception

    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   LevenshteinDistance
  ' Goal:   Calculate the distance between word [strS] and word [strT]
  ' Source: web > en.wikipedia.org/wiki/Levenshtein_distance
  '         Hjelmqvist, Sten (26 Mar 2012), "Fast, memory efficient Levenshtein algorithm"
  '         web > www.codeproject.com/Articles/13525/Fast-memory-efficient-Levenshtein-algorithm
  ' Extensension:
  '         - If [intMax] is larger than 0, then don't calculate distances larger than [intMax]
  ' History:
  ' 18-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function LevenshteinDistance(ByRef strS As String, ByRef strT As String, _
                                      Optional ByVal intMax As Integer = 0) As Integer
    Dim arV0(0 To strT.Length + 1) As Integer ' Work vector for distances
    Dim arV1(0 To strT.Length + 1) As Integer ' Work vector for distances
    Dim intCostS As Integer               ' Costs for substitution
    Dim intCostD As Integer               ' Costs for deletion
    Dim intCostI As Integer               ' Costs for insertion
    Dim intI As Integer                   ' Counter
    Dim intJ As Integer                   ' Counter

    Try
      ' Simple comparison of strings
      If (strS = strT) Then Return 0
      If (strS = "") Then Return strT.Length
      If (strT = "") Then Return strS.Length
      ' Initialize v0 (the previous row of distances)
      ' This row is A[0][i]: edit distance for an empty s
      ' The distance is just the number of characters to delete from t
      For intI = 0 To arV0.Length - 1
        arV0(intI) = intI
      Next intI
      ' Walk through string [strS]
      For intI = 1 To strS.Length
        ' Calculate v1 (current row distances) from the previous row v0
        '  First element of v1 is A[i+1][0]
        '  Edit distance is delete (i+1) chars from s to match empty t
        arV1(0) = intI * 4

        ' Use formula to fill in the rest of the row
        'for (int j = 0; j < t.Length; j++)
        For intJ = 1 To strT.Length
          ' Compare s[i] with t[j]
          intCostS = LevSubstCost(Mid(strS, intI, 1), Mid(strT, intJ, 1))
          ' ============== DEBUG ===========
          If (intCostS = 0) AndAlso (Mid(strS, intI, 1) <> Mid(strT, intJ, 1)) Then Stop
          ' ================================
          ' Only continue if this is not zero
          If (intCostS = 0) Then
            ' Note the minimum: zero
            arV1(intJ) = arV0(intJ - 1)
          Else
            ' Check out what deletion costs
            intCostD = LevDelCost(Mid(strS, intI, 1))
            ' Check out what insertion costs
            intCostI = LevInsCost(Mid(strT, intJ, 1))
            ' Check the minimum
            arV1(intJ) = Math.Min(arV1(intJ - 1) + intCostI, Math.Min(arV0(intJ) + intCostD, arV0(intJ - 1) + intCostS))
          End If

          ' Calculate the costs between s[i] and t[j]
          '    var cost = (s[i] == t[j]) ? 0 : 1;
          ' OLD: intCost = IIf(Mid(strS, intI, 1) = Mid(strT, intJ, 1), 0, 1)
          ' Calculate minimum
          '    v1[j + 1] = Minimum(v1[j] + 1, v0[j + 1] + 1, v0[j] + cost);
          ' OLD: arV1(intJ + 1) = Math.Min(arV1(intJ) + 1, Math.Min(arV0(intJ + 1) + 1, arV0(intJ) + intCost))
        Next intJ

        ' Copy v1 (current row) to v0 (previous row) for next interation
        For intJ = 0 To arV0.Length - 1
          ' COpy this element
          arV0(intJ) = arV1(intJ)
        Next intJ
      Next intI
      ' ========== DEBUG ============
      If (arV1(strT.Length) = 0) Then Stop
      ' =============================
      ' The distance is the last element in v1
      Return arV1(strT.Length)
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/LevenshteinDistance error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   LevSubstCost
  ' Goal:   What are the costs of substituting S for T?
  '         I use a measure between 0 and 5
  ' History:
  ' 18-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function LevSubstCost(ByVal chS As Char, ByVal chT As Char) As Integer
    Dim intIdx As Integer ' Index

    Try
      ' Are they equal? --> no cost
      If (chS = chT) Then Return 0
      ' Get the index within lev_strSubstOrg
      intIdx = Asc(LCase(chS)) - Asc("a") + 1
      ' Check range --> return extra large distance if needed
      If (intIdx <= 0) OrElse (intIdx > lev_strSubstOrg.Length) Then Return 4
      ' Check if the costs are 1, 2 or 3
      If (Mid(lev_strSubstCs1, intIdx, 1) = chT) Then
        Return 1
      ElseIf (Mid(lev_strSubstCs2, intIdx, 1) = chT) Then
        Return 2
      End If
      ' No special case: cost is predefined constant
      Return 9
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/LevSubstCost error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   LevDelCost
  ' Goal:   What are the costs of deleting S?
  ' History:
  ' 18-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function LevDelCost(ByVal chS As Char) As Integer
    Dim intIdx As Integer   ' Index
    Dim intCost As Integer  ' Costs

    Try
      ' Get the index within lev_strSubstOrg
      intIdx = Asc(LCase(chS)) - Asc("a") + 1
      ' Check range --> return extra large distance if needed
      If (intIdx <= 0) OrElse (intIdx > lev_strSubstOrg.Length) Then
        intCost = 4
      Else
        intCost = CInt(Mid(lev_strDelCosts, intIdx, 1))
      End If
      ' Adaptation of costs...
      If (intCost > 3) Then intCost = 9
      ' Return the deletion costs (which are 1 or 2)
      Return intCost
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/LevDelCost error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   LevInsCost
  ' Goal:   What are the costs of inserting S?
  ' History:
  ' 18-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function LevInsCost(ByVal chS As Char) As Integer
    Dim intIdx As Integer ' Index
    Dim intCost As Integer  ' Costs

    Try
      ' Get the index within lev_strSubstOrg
      intIdx = Asc(LCase(chS)) - Asc("a") + 1
      ' Check range --> return extra large distance if needed
      If (intIdx <= 0) OrElse (intIdx > lev_strSubstOrg.Length) Then
        intCost = 4
      Else
        intCost = CInt(Mid(lev_strInsCosts, intIdx, 1))
      End If
      ' Return the insertion costs (which are 1 or 2)
      Return intCost
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/LevInsCost error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DamerauLevenshteinDistance
  ' Goal:   Calculate the distance between word [strS] and word [strT]
  ' Source: en.wikipedia.org/wiki/Damerau%E2%80%93Levenshtein_distance
  ' History:
  ' 18-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function DamerauLevenshteinDistance(ByVal strS As String, ByVal strT As String) As Integer
    Dim arScore = New Integer(strS.Length + 1, strT.Length + 1) {}
    Dim INF = strS.Length + strT.Length
    Dim intI As Integer     ' Counter
    Dim intJ As Integer     ' Counter

    Try
      If (strS.Length = 0) Then
        If (strT.Length = 0) Then
          Return 0
        Else
          Return strT.Length
        End If
      ElseIf (strT.Length = 0) Then
        Return strS.Length
      End If

      arScore(0, 0) = INF
      For intI = 0 To strS.Length
        arScore(intI + 1, 1) = intI
        arScore(intI + 1, 0) = INF
      Next
      For intJ = 0 To strT.Length
        arScore(1, intJ + 1) = intJ
        arScore(0, intJ + 1) = INF
      Next

      Dim sdThis = New SortedDictionary(Of Char, Integer)()
      For Each letter In (strS & strT)
        If Not sdThis.ContainsKey(letter) Then
          sdThis.Add(letter, 0)
        End If
      Next

      For i = 1 To strS.Length
        Dim intDB = 0
        For j = 1 To strT.Length
          Dim intI1 = sdThis(strT(j - 1))
          Dim intJ1 = intDB

          If strS(i - 1) = strT(j - 1) Then
            arScore(i + 1, j + 1) = arScore(i, j)
            intDB = j
          Else
            arScore(i + 1, j + 1) = Math.Min(arScore(i, j), Math.Min(arScore(i + 1, j), arScore(i, j + 1))) + 1
          End If

          arScore(i + 1, j + 1) = Math.Min(arScore(i + 1, j + 1), arScore(intI1, intJ1) + (i - intI1 - 1) + 1 + (j - intJ1 - 1))
        Next j

        sdThis(strS(i - 1)) = i
      Next i
      ' Return the total
      Return arScore(strS.Length + 1, strT.Length + 1)
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/DamerauLevenshteinDistance error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try

  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MorphSeqToNum
  ' Goal:   Convert certain char sequences into numbers
  ' History:
  ' 31-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MorphSeqToNum(ByVal strIn As String, ByVal bForward As Boolean) As String
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
      HandleErr("modMorph/MorphSeqToNum error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SimpleDistance
  ' Goal:   Calculate the distance between [strS] and [strT] using my own method
  '         Return: the total costs
  '         Also return:
  '         [strPath]:  details of substitutions, insertions and deletions
  '         [intNum]:   total number of substitutions, insertions and deletions
  '         Make use of the minimum so far [intCumMin], so that calculations do not have to
  '           continue when the costs are above [intCumMin]
  ' History:
  ' 20-05-2013  ERK Created
  ' 21-05-2013  ERK Use look-ahead window to help determine whether a substitution is justified
  ' ------------------------------------------------------------------------------------
  Public Function SimpleDistance(ByRef strS_in As String, ByRef strT_in As String, ByRef strPath As String, _
                                 ByRef intNum As Integer, ByVal intCumMin As Integer) As Integer
    Dim intIdxS As Integer = 1    ' Source index
    Dim intIdxT As Integer = 1    ' Target index
    Dim intCost As Integer = 0    ' Total costs
    Dim intThis As Integer = 0    ' Current costs
    Dim intNumSub As Integer = 0  ' Number of substitutions
    Dim intNumIns As Integer = 0  ' Number of insertions
    Dim intNumDel As Integer = 0  ' Number of deletions
    Dim intThr_CC As Integer = 5  ' One threshold     --> Depends on lev_strDelCosts
    Dim intThr_AA As Integer = 2  ' Another threshold --> Depends on lev_strDelCosts
    Dim intLkAhd As Integer = 2   ' Size of the look-ahead window
    Dim chS As Char               ' Current source character
    Dim chT As Char               ' Current target character
    Dim strLkAhd As String = ""   ' Look ahead window
    Dim strS As String            ' Corrected input string
    Dim strT As String            ' Corrected input string

    Try
      ' Simple comparison of strings
      If (strS_in = strT_in) Then Return 0
      If (strS_in = "") OrElse (strT_in = "") Then Return smp_intInf
      ' ========= DEBUG ===========
      ' If (strS = "aece" AndAlso strT = "ece") Then Stop
      ' ===========================
      ' Initialize
      strPath = "" : smp_colPath.Clear() : smp_intPath = 0
      strS = MorphSeqToNum(strS_in, True) : strT = MorphSeqToNum(strT_in, True)
      ' Walk through the source letter by letter
      While (intIdxS <= strS.Length) AndAlso (intIdxT <= strT.Length) AndAlso (intCost < intCumMin)
        ' Check source and target vowel
        chS = Mid(strS, intIdxS, 1) : chT = Mid(strT, intIdxT, 1)
        If (chS = chT) Then
          ' Source and target are equal, so there are no costs. But keep track of what happens
          SimpleOneEq(intIdxS, intIdxT)
        Else
          ' Since they are not equal, we need to have a look-ahead window in the target string
          strLkAhd = Mid(strT, intIdxT + 1, intLkAhd)
          If (InStr(strLkAhd, chS) > 0) Then
            ' The source is coming up in the look-ahead window of target, 
            '   so we skip the current target character by inserting it
            intCost += SimpleOneInsOrDel("i", chT, intIdxT, intNumIns)
          ElseIf (InStr(Mid(strS, intIdxS + 1, intLkAhd), chT) > 0) Then
            ' The target is in the upcoming window of the source, so delete the current source character
            intCost += SimpleOneInsOrDel("d", chS, intIdxS, intNumDel)
          Else
            ' Source character is not found in the upcoming letters of target, so determine s,d,i in the usual mannner
            If (InStr(smp_strCons, chS) > 0) Then
              ' If the source is a consonant, we need to have a consonant or ambi in the target at some point
              If (InStr(smp_strCons, chT) > 0 OrElse InStr(smp_strAmbi, chT) > 0) Then
                ' Do substitution or else delection
                intCost += SimpleSubOrDel(chS, chT, intThr_CC, intIdxS, intIdxT, intNumSub, intNumDel)
              Else
                ' The source is consonant, the target a vowel --> Insert the target vowel
                intCost += SimpleOneInsOrDel("i", chT, intIdxT, intNumIns)
              End If
            ElseIf (InStr(smp_strAmbi, chS) > 0) Then
              ' The source is an ambivalent one --> pass certain ones, but hold others
              If (InStr(smp_strAmbi, chT) > 0) OrElse (InStr(smp_strVowel, chT) > 0) Then
                ' Ambivalent source characters can be substituted by another ambivalent one or by a vowel
                intCost += SimpleSubOrDel(chS, chT, intThr_CC, intIdxS, intIdxT, intNumSub, intNumDel)
              Else
                ' Ambivalent source with consonent target --> get the cost
                intThis = LevSubstCost(chS, chT)
                ' Take 2 as threshold
                intCost += SimpleSubOrDel(chS, chT, intThr_AA, intIdxS, intIdxT, intNumSub, intNumDel)
              End If
            ElseIf (chS = "&") Then
              ' We should advance the source at any rate
              intIdxS += 1
              ' If the target has "and", then advance the target
              If (DoLike(Mid(strT, intIdxT, 3), "and|ond|ant|ont")) Then
                ' Skip target without costs
                intIdxT += 3
                ' strPath &= "e|e|e|"
                smp_colPath.Add("e") : smp_colPath.Add("e") : smp_colPath.Add("e")
                smp_intPath = smp_colPath.Count
              Else
                ' Note that we kind of deleted the source at some costs
                ' strPath &= "d;&;1|"
                smp_colPath.Add("d;&;1")
                smp_intPath = smp_colPath.Count
              End If
            Else
              ' Source is a vowel ...
              If (InStr(smp_strCons, chT) > 0) Then
                ' We cannot have a vowel being replaced by a consonant --> delete it
                intCost += SimpleOneInsOrDel("d", chS, intIdxS, intNumDel)
              ElseIf (InStr(smp_strAmbi, chT) > 0) Then
                ' Some vowels can be replace by an ambivalent character in the target
                intThis = LevSubstCost(chS, chT)
                ' Take 2 as threshold
                intCost += SimpleSubOrDel(chS, chT, intThr_AA, intIdxS, intIdxT, intNumSub, intNumDel)
              Else
                ' Vowels can be replaced by other vowels...
                intCost += SimpleSubOrDel(chS, chT, intThr_CC, intIdxS, intIdxT, intNumSub, intNumDel)
              End If
            End If
          End If
        End If
      End While
      ' Delete any remaining source characters
      While (intIdxS <= strS.Length) AndAlso (intCost < intCumMin)
        ' Get this source character
        chS = Mid(strS, intIdxS, 1)
        ' Delete it
        intCost += SimpleOneInsOrDel("d", chS, intIdxS, intNumDel)
      End While
      ' Insert any remaining destination characters
      While (intIdxT <= strT.Length) AndAlso (intCost < intCumMin)
        ' Get target character
        chT = Mid(strT, intIdxT, 1)
        ' Insert it
        intCost += SimpleOneInsOrDel("a", chT, intIdxT, intNumIns)
      End While
      ' Make [strPath] anew
      strPath = SimplePath()
      ' Calculate the total number of operations
      intNum = intNumDel + intNumSub + intNumIns
      ' Check distance
      If (intCost = 0) AndAlso (intCumMin > 0) Then Stop
      ' Return the total costs
      Return intCost
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/SimpleDistance error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return smp_intInf
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SimpleSimilarity
  ' Goal:   Calculate the similarity between [strS] and [strT] using my own method
  '         Return: a similarity measure between 0 and 100
  ' Notes:  The similarity measure is not very good, because words are not fully
  '           economically aligned towards one another. Unfortunately...
  ' History:
  ' 01-06-2013  ERK Derived from SimpleDistance
  ' ------------------------------------------------------------------------------------
  Public Function SimpleSimilarity(ByRef strS_in As String, ByRef strT_in As String) As Integer
    Dim intIdxS As Integer = 1    ' Source index
    Dim intIdxT As Integer = 1    ' Target index
    Dim intSimi As Integer = 0    ' Total similarity
    Dim intThis As Integer = 0    ' Current costs
    Dim intNumSub As Integer = 0  ' Number of substitutions
    Dim intNumIns As Integer = 0  ' Number of insertions
    Dim intNumDel As Integer = 0  ' Number of deletions
    Dim intThr_CC As Integer = 5  ' One threshold     --> Depends on lev_strDelCosts
    Dim intThr_AA As Integer = 2  ' Another threshold --> Depends on lev_strDelCosts
    Dim intEq As Integer = 5      ' Standard benefit for equality
    Dim intLkAhd As Integer = 2   ' Size of the look-ahead window
    Dim chS As Char               ' Current source character
    Dim chT As Char               ' Current target character
    Dim strLkAhd As String = ""   ' Look ahead window
    Dim strS As String            ' Corrected input string
    Dim strT As String            ' Corrected input string
    Dim strS_out As String = ""   ' Output string S
    Dim strT_out As String = ""   ' Output string T

    Try
      ' Simple comparison of strings
      If (strS_in = strT_in) Then Return 100
      If (strS_in = "") OrElse (strT_in = "") Then Return 0
      ' ========= DEBUG ===========
      ' If (strS = "aece" AndAlso strT = "ece") Then Stop
      ' ===========================
      ' Initialize
      strS = MorphSeqToNum(strS_in, True) : strT = MorphSeqToNum(strT_in, True)
      ' Walk through the source letter by letter
      While (intIdxS <= strS.Length) AndAlso (intIdxT <= strT.Length)
        ' Check source and target vowel
        chS = Mid(strS, intIdxS, 1) : chT = Mid(strT, intIdxT, 1)
        If (chS = chT) Then
          ' Source and target are equal, so we add the standard benefit for similarity
          intSimi += intEq : strS_out &= chS : strT_out &= chT : intIdxS += 1 : intIdxT += 1
        Else
          ' Since they are not equal, we need to have a look-ahead window in the target string
          strLkAhd = Mid(strT, intIdxT + 1, intLkAhd)
          ' Check if source character is in look-ahead window of target
          If (InStr(strLkAhd, chS) > 0) Then
            ' The source is coming up in the look-ahead window of target, 
            '   so we skip the current target character == insertion rule
            intSimi += 5 - LevInsCost(chT) : strS_out &= "-" : strT_out &= chT : intIdxT += 1
          ElseIf (InStr(Mid(strS, intIdxS + 1, intLkAhd), chT) > 0) Then
            ' The target is in the upcoming window of the source, so delete the current source character
            intSimi += 5 - LevDelCost(chS) : strS_out &= chS : strT_out &= "-" : intIdxS += 1
          Else
            ' Source character is not found in the upcoming letters of target, so determine s,d,i in the usual mannner
            If (InStr(smp_strCons, chS) > 0) Then
              ' If the source is a consonant, we need to have a consonant or ambi in the target at some point
              If (InStr(smp_strCons, chT) > 0 OrElse InStr(smp_strAmbi, chT) > 0) Then
                ' Do substitution or else delection
                intSimi += 5 - SimpleSubDel(chS, chT, intThr_CC, intIdxS, intIdxT, strS_out, strT_out)
              Else
                ' The source is consonant, the target a vowel --> Insert the target vowel
                intSimi += 5 - LevInsCost(chT) : strS_out &= "-" : strT_out &= chT : intIdxT += 1
              End If
            ElseIf (InStr(smp_strAmbi, chS) > 0) Then
              ' The source is an ambivalent one --> pass certain ones, but hold others
              If (InStr(smp_strAmbi, chT) > 0) OrElse (InStr(smp_strVowel, chT) > 0) Then
                ' Ambivalent source characters can be substituted by another ambivalent one or by a vowel
                intSimi += 5 - SimpleSubDel(chS, chT, intThr_CC, intIdxS, intIdxT, strS_out, strT_out)
              Else
                ' Ambivalent source with consonent target --> get the cost
                intSimi += 5 - SimpleSubDel(chS, chT, intThr_AA, intIdxS, intIdxT, strS_out, strT_out)
              End If
            ElseIf (chS = "&") Then
              ' We should advance the source at any rate
              intIdxS += 1 : strS_out &= chS
              ' If the target has "and", then advance the target
              If (DoLike(Mid(strT, intIdxT, 3), "and|ond|ant|ont")) Then
                ' Skip target without costs
                intIdxT += 3 : intSimi += intEq * 3 : strT_out &= Mid(strT, intIdxT, 3)
              Else
                ' Note that we kind of deleted the source at some costs
                intSimi += 5 - LevDelCost(chS) : strT_out &= "-"
              End If
            Else
              ' Source is a vowel ...
              If (InStr(smp_strCons, chT) > 0) Then
                ' We cannot have a vowel being replaced by a consonant --> delete it
                intSimi += 5 - LevDelCost(chS) : strS_out &= chS : strT_out &= "-" : intIdxS += 1
              ElseIf (InStr(smp_strAmbi, chT) > 0) Then
                ' Some vowels can be replace by an ambivalent character in the target
                intSimi += 5 - SimpleSubDel(chS, chT, intThr_AA, intIdxS, intIdxT, strS_out, strT_out)
              Else
                ' Vowels can be replaced by other vowels...
                intSimi += 5 - SimpleSubDel(chS, chT, intThr_CC, intIdxS, intIdxT, strS_out, strT_out)
              End If
            End If
          End If
        End If
      End While
      ' Delete any remaining source characters
      While (intIdxS <= strS.Length)
        ' Get this source character
        chS = Mid(strS, intIdxS, 1)
        ' Delete it
        ' intSimi += SimpleOneInsOrDel("d", chS, intIdxS, intNumDel)
        intSimi += 5 - LevDelCost(chS) : strS_out &= chS : strT_out &= "-" : intIdxS += 1
      End While
      ' Insert any remaining destination characters
      While (intIdxT <= strT.Length)
        ' Get target character
        chT = Mid(strT, intIdxT, 1)
        ' Insert it
        ' intSimi += SimpleOneInsOrDel("a", chT, intIdxT, intNumIns)
        intSimi += 5 - LevInsCost(chT) : strS_out &= "-" : strT_out &= chT : intIdxT += 1
      End While
      ' Normalize the similarity measure to a positive number between 0 and 100
      intSimi = IIf(intSimi < 0, 0, 100 * intSimi / (intEq * strS_out.Length))
      '' Check distance
      'If (intSimi = 0) Then Stop
      ' ========== DEBUG =========
      Debug.Print(intSimi & "% - " & strS_in & "/" & strT_in & " S=[" & strS_out & "] >> T=[" & strT_out & "]")
      ' ==========================
      ' Return the total similarity measure
      Return intSimi
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/SimpleSimilarity error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return smp_intInf
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   SimplePath
  ' Goal:   Combine subsequent s,d,i into one 
  '         Then come back with the path string
  ' History:
  ' 21-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function SimplePath() As String
    Dim strPath As String = ""    ' Path to be returned
    Dim strThis As String = ""    ' Current one
    Dim strPrev As String = ""    ' Previous one
    Dim strThisOp As String = ""  ' Current operation
    Dim strPrevOp As String = ""  ' Previous operation
    Dim strThisS As String = ""   ' Current source
    Dim strThisT As String = ""   ' Current target
    Dim strPrevS As String = ""   ' Previous source
    Dim strPrevT As String = ""   ' Previous target
    Dim colThis As New StringColl ' Where we gather the elements
    Dim arThis() As String        ' Split one
    Dim arPrev() As String        ' Split previous one
    Dim intI As Integer           ' Counter
    Dim intThisCost As Integer    ' Current cost
    Dim intPrevCost As Integer    ' Current cost
    Dim bReDo As Boolean          ' Redo previous one or not?

    Try
      ' Validate
      If (smp_colPath.Count = 0) Then Return ""
      ' Walk the array
      For intI = 0 To smp_colPath.Count - 1
        ' Get current one
        strThis = smp_colPath.Item(intI) : strThisOp = Left(strThis, 1)
        ' Split the previous and the current in any case
        arThis = Split(strThis, ";")
        arPrev = Split(strPrev, ";")
        ' Get current S, T and cost
        Select Case arThis(0)
          Case "s"
            strThisS = arThis(1) : strThisT = arThis(2) : intThisCost = arThis(3)
          Case "i", "a"
            strThisS = "" : strThisT = arThis(1) : intThisCost = arThis(2)
          Case "d"
            strThisS = arThis(1) : strThisT = "" : intThisCost = arThis(2)
        End Select
        ' Get previous S, T and cost
        Select Case arPrev(0)
          Case "s"
            strPrevS = arPrev(1) : strPrevT = arPrev(2) : intPrevCost = arPrev(3)
          Case "i", "a"
            strPrevS = "" : strPrevT = arPrev(1) : intPrevCost = arPrev(2)
          Case "d"
            strPrevS = arPrev(1) : strPrevT = "" : intPrevCost = arPrev(2)
        End Select
        ' Be positive!
        bReDo = True
        ' Compare this operation with previous one
        Select Case strThisOp & strPrevOp
          Case "ss", "sd", "si", "sa", "ds", "di", "da", "is", "id", "as", "ad"
            ' This becomes type "s"
            strThis = "s;" & strPrevS & strThisS & ";" & strPrevT & strThisT & ";" & intPrevCost + intThisCost
          Case "dd"
            ' This stays "d"
            strThis = "d;" & strPrevS & strThisS & ";" & intPrevCost + intThisCost
          Case "ii", "ia", "ai", "ai"
            ' This becomes type "i"
            strThis = "i;" & strPrevT & strThisT & ";" & intPrevCost = intThisCost
          Case "aa"
            ' This stays type "a"
            strThis = "a;" & strPrevT & strThisT & ";" & intPrevCost = intThisCost
          Case Else
            ' No need to redo
            bReDo = False
        End Select
        ' Need to redo previous one?
        If (bReDo) Then
          colThis.Item(colThis.Count - 1) = strThis
        Else
          ' Just add it
          colThis.Add(strThis)
        End If
        ' Note last operation
        strPrev = strThis : strPrevOp = Left(strThis, 1)
      Next intI
      ' Return the path
      Return colThis.Text.Replace(vbCrLf, "|")
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/SimplePath error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SimpleOneEq
  ' Goal:   Add one equality operation
  ' History:
  ' 21-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub SimpleOneEq(ByRef intIdxS As Integer, ByRef intIdxT As Integer)
    Try
      smp_colPath.Add("e") : smp_intPath = smp_colPath.Count
      'strPath &= "e|"
      intIdxS += 1 : intIdxT += 1
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/SimpleOneEq error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Function SimpleSubDel(ByVal chS As Char, ByVal chT As Char, ByVal intThreshold As Integer, _
                                ByRef intIdxS As Integer, ByRef intIdxT As Integer, _
                                ByRef strS_out As String, ByRef strT_out As String) As Integer
    Dim intThis As Integer  ' Costs
    Dim intCost As Integer  ' Costs

    Try
      intThis = LevSubstCost(chS, chT)
      If (intThis < intThreshold) Then
        ' Do substitution
        intCost = intThis : strS_out &= chS : strT_out &= chT : intIdxS += 1 : intIdxT += 1
      Else
        ' Do deletion
        intCost = LevDelCost(chS) : strS_out &= chS : strT_out &= "-" : intIdxS += 1
      End If
      ' Return the results
      Return intCost
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/SimpleSubDel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return smp_intInf
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SimpleSubOrDel
  ' Goal:   Substitute S for T if not too costy, otherwise delete S
  ' History:
  ' 20-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function SimpleSubOrDel(ByVal chS As Char, ByVal chT As Char, ByVal intThreshold As Integer, ByRef intIdxS As Integer, ByRef intIdxT As Integer, _
                                   ByRef intNumSub As Integer, ByRef intNumDel As Integer) As Integer
    Dim intThis As Integer  ' Costs
    Dim strCur As String    ' Current operation

    Try
      ' Check the cost of a substitution
      intThis = LevSubstCost(chS, chT)
      If (intThis < intThreshold) Then
        ' Okay, perform substitution
        intIdxS += 1 : intIdxT += 1
        ' What is the operation?
        strCur = "s;" & chS & ";" & chT & ";" & intThis
        ' Keep track of what we do
        smp_colPath.Add(strCur)
        smp_intPath = smp_colPath.Count
        'strPath &= strCur & "|"
        ' Bookkeeping: keep track of the number of substitutions
        intNumSub += 1
      Else
        ' Don't do a substitution: delete this input character
        intThis = SimpleOneInsOrDel("d", chS, intIdxS, intNumDel)
      End If
      ' Return the costs for this operation
      Return intThis
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/SimpleSubOrDel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return smp_intInf
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   SimpleOneInsOrDel
  ' Goal:   Insert, delete or append the character
  ' History:
  ' 21-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function SimpleOneInsOrDel(ByVal strOp As String, ByVal chX As Char, ByRef intIdxX As Integer, _
                                   ByRef intNumOp As Integer) As Integer
    Dim intThis As Integer  ' Costs
    Dim strCur As String    ' Current operation
    Dim arCur() As String   ' Array

    Try
      intThis = LevInsCost(chX) : intIdxX += 1
      ' Check what the last operation was
      If (smp_intPath > 0) AndAlso (Left(smp_colPath.Item(smp_intPath - 1), 1) = strOp) Then
        ' We need to combine the operations
        arCur = Split(smp_colPath.Item(smp_intPath - 1), ";")
        ' Adapt deletion or insertion + counting
        arCur(1) &= chX : arCur(2) += intThis
        ' Re-combine result
        smp_colPath.Item(smp_intPath - 1) = Join(arCur, ";")
      Else
        ' This is a separate operation
        strCur = strOp & ";" & chX & ";" & intThis
        smp_colPath.Add(strCur)
        smp_intPath = smp_colPath.Count
      End If
      '' Keep track of what we do
      'strPath &= smp_colPath.Item(smp_intPath - 1) & "|"
      ' Keep track of the number of operations
      intNumOp += 1
      ' Return the costs for this operation
      Return intThis
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/SimpleOneInsOrDel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return smp_intInf
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMEprojReportVernpos_Click
  ' Goal:   Provide a report that contains all entries in [Vernpos] and corresponding lemma's
  ' History:
  ' 23-01-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function MakeVernposReport(ByVal strLang As String, ByRef strDictRepFile As String) As Boolean
    Dim dtrFound() As DataRow       ' Result of selecting
    Dim colBack As New StringColl   ' What we return
    Dim strBack As String = ""      ' Content of [colBack]
    Dim strLine As String = ""      ' One line for the report
    Dim strSrc As String = ""       ' Source of this addition
    Dim strFeats As String = ""     ' Features
    Dim arFeat() As String          ' Array of features
    Dim intPtc As Integer           ' Percentage
    Dim intI As Integer             ' Counter
    Dim intJ As Integer             ' Counter

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Logging("Could not open the morphological dictionary dataset") : Return False
      If (strLang = "") Then Return False
      ' Select the vernpos table
      dtrFound = tdlMorphDict.Tables("VernPos").Select("", "Vern ASC, Pos ASC, l ASC")
      ' Write a header line
      strLine = "Vern" & ";" & "Pos" & ";" & "Lemma" & ";" & "Source" & ";" & "Features"
      colBack.Add(strLine)
      ' Walk through the results
      For intI = 0 To dtrFound.Length - 1
        ' Where are we?
        intPtc = (intI + 1) * 100 \ dtrFound.Length
        Status("Vernpos table: " & intPtc & "%", intPtc)
        ' Access this entry
        With dtrFound(intI)
          ' Preparations...
          arFeat = Split(.Item("f").ToString, ";")
          Select Case strLang
            Case "OE"
              strSrc = "B&T"
            Case "ME"
              strSrc = "Gutenberg"
              ' Double check in features
              intJ = arFeat.Length - 1
              If (intJ >= 0) Then
                If (arFeat(intJ) Like "src=*") Then
                  ' Get the source...
                  strSrc = Mid(arFeat(intJ), 5)
                  Array.Resize(arFeat, intJ)
                End If
              End If
            Case "eModE", "LmodE", "lModE"
              strSrc = "GCIDE"
            Case Else
              strSrc = "unknown"
          End Select
          ' Recombine features
          strFeats = Join(arFeat, "|")
          ' Prepare one line for the report
          strLine = .Item("Vern").ToString & ";" & _
                    .Item("Pos").ToString & ";" & _
                    .Item("l").ToString & ";" & _
                    strSrc & ";" & strFeats & ";"
          ' Store the line
          colBack.Add(strLine)
        End With
      Next intI
      ' Finish the report
      Status("finishing report...")
      strBack = colBack.Text
      ' Save the remain report somewhere
      strDictRepFile = GetDocDir() & "\VernposReport_" & strLang & ".csv"
      IO.File.WriteAllText(strDictRepFile, strBack, System.Text.Encoding.UTF8)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MakeDictReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MakeDictReport
  ' Goal:   Provide an HTML file and report that contains all entries in [Morph] and corresponding lemma's
  ' History:
  ' 18-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MakeDictReport(ByVal strLang As String, ByRef strDictRepFile As String) As Boolean
    Dim dtrDict() As DataRow        ' Result of selecting in the correct dictionary
    Dim dtrSense() As DataRow       ' Sense child
    Dim dtrFound() As DataRow       ' Result of selecting
    Dim dtrLexF() As DataRow        ' LexFun entries
    Dim dtrMorph As DataRow         ' One record
    Dim colHtml As New StringColl   ' Gather what we return
    Dim strHtml As String = ""      ' Report text
    Dim strLemma As String = ""     ' (New) lemma we are dealing with
    Dim strVern As String = ""      ' Vernacular form
    Dim strDef As String = ""       ' Definition
    Dim strLangDict As String = ""  ' Language dictionary
    Dim strPosField As String = ""  ' Which field in MorphDict must be taken as POS?
    Dim strPos As String = ""       ' Part-of-speech
    Dim strFeat As String = ""      ' Features
    Dim colForm As New StringColl   ' Collection of forms for this entry (form + POS)
    Dim intPtc As Integer           ' Percentage
    Dim intIdx As Integer           ' Index
    Dim intI As Integer             ' Counter
    Dim intJ As Integer             ' Counter
    ' Conversion from OEdict.xml @lf to MorphDict @POS and @f fields
    Dim arVbMean() As String = {"3rd pres", "pastSg", "pastPl", "3rd sing", "past", "ptpIs", "ptp"}
    Dim arVbMeanPos() As String = {"VBP", "VBD", "VBD", "VBP", "VBD", "VAN", "VAN"}
    Dim arVbMeanF() As String = {"number=sg;person=3", "number=sg", "number=pl", "number=sg;person=3", "", "", ""}

    Try
      ' Validate

      ' Double check
      Select Case strLang
        Case "OE"
          strPosField = "Pos"
        Case "ME"
          strPosField = "Label"
        Case "eModE", "LmodE"
        Case Else
          Return False
      End Select
      ' Name of language dictionary is standard, cross the board
      strLangDict = strLang & "dict_out.xml"
      ' Initialise morphological dictionary
      MorphDictIni(False, "MorphDict" & strLang & ".xml")
      If (tdlMorphDict Is Nothing) Then Logging("Could not open the morphological dictionary dataset") : Return False
      ' Try read the OE dictionary
      If (Not MorphReadOEdict(strLangDict)) Then Status("Could not read the OE dictionary database") : Return False
      ' Start the report
      colHtml.Add("<html><body><table>")
      ' Check if the @t attribute has been added already
      If (tdlMorphDict.Tables("Morph").Select("t='l'").Length = 0) Then
        ' Copy all that is needed from OEdict.Entry --> Morphdict.Morph
        dtrDict = tdlOEdict.Tables("Entry").Select("", "l ASC, Pos ASC")
        For intI = 0 To dtrDict.Length - 1
          ' Show where we are
          intPtc = (intI + 1) * 100 \ dtrDict.Length
          Status("Typing pass 1: " & intPtc & "%", intPtc)
          ' Get this entry
          strLemma = dtrDict(intI).Item("l").ToString : strPos = UCase(dtrDict(intI).Item("Pos").ToString)
          ' Is this a real lemma?
          If (Left(strLemma, 1) Like "[a-zA-Z]") AndAlso (strPos <> "") Then
            ' Check if this entry is in the MorphDict.Morph
            dtrFound = tdlMorphDict.Tables("Morph").Select("Vern='" & strLemma & "' AND l='" & strLemma & _
                                                           "' AND Pos='" & strPos & "'")
            If (dtrFound.Length = 0) Then
              ' Check if there is an adapted POS in the <sense> entry
              dtrSense = tdlOEdict.Tables("Sense").Select("EntryId=" & dtrDict(intI).Item("EntryId").ToString, "SenseId ASC")
              If (dtrSense.Length > 0) Then
                ' Evaluate the first sense's POS value
                If (dtrSense(0).Item("Pos").ToString <> "") AndAlso (dtrSense(0).Item("Pos").ToString <> strPos) Then
                  ' Take this as POS
                  strPos = UCase(dtrSense(0).Item("Pos").ToString)
                End If
              End If
              ' Add the entry to the table
              dtrMorph = AddOneDataRow(tdlMorphDict, "Morph", "MorphId", "MorphList")
              With dtrMorph
                .Item("Vern") = strLemma : .Item("l") = strLemma : .Item("EntryId") = dtrDict(intI).Item("EntryId")
                .Item("f") = dtrDict(intI).Item("f").ToString : .Item("t") = "l" : .Item("Pos") = strPos
              End With
            End If
            ' Copy all <LexFun> entries --> MorphDict.morph
            dtrLexF = tdlOEdict.Tables("LexFun").Select("EntryId=" & dtrDict(intI).Item("EntryId").ToString)
            For intJ = 0 To dtrLexF.Length - 1
              ' Get this LV entry
              strVern = dtrLexF(intJ).Item("lv").ToString
              ' Evaluate the LF entry
              intIdx = Array.FindIndex(arVbMean, Function(strIn As String) strIn = dtrLexF(intJ).Item("lf").ToString)
              If (intIdx >= 0) Then
                strPos = arVbMeanPos(intIdx) : strFeat = arVbMeanF(intIdx)
              Else
                strFeat = ""
              End If
              ' Check if this entry is in the MorphDict.Morph
              dtrFound = tdlMorphDict.Tables("Morph").Select("Vern='" & strVern & "' AND l='" & strLemma & _
                                                             "' AND Pos='" & strPos & "'")
              If (dtrFound.Length = 0) Then
                ' Add the entry to the table
                dtrMorph = AddOneDataRow(tdlMorphDict, "Morph", "MorphId", "MorphList")
                With dtrMorph
                  .Item("Vern") = strVern : .Item("l") = strLemma : .Item("EntryId") = dtrLexF(intJ).Item("EntryId")
                  .Item("f") = strFeat : .Item("t") = "d" : .Item("Pos") = strPos
                End With
              End If
            Next intJ
          End If
        Next intI
        ' Mark everthing as "derived" first
        dtrFound = tdlMorphDict.Tables("Morph").Select("")
        For intI = 0 To dtrFound.Length - 1
          ' Show where we are
          intPtc = (intI + 1) * 100 \ dtrFound.Length
          Status("Typing pass 3: " & intPtc & "%", intPtc)
          ' Mark as derived
          If (dtrFound(intI).Item("t").ToString = "") Then dtrFound(intI).Item("t") = "d"
        Next intI
        ' Add the type information from LangDict to Morph Table
        dtrDict = tdlOEdict.Tables("Entry").Select("", "l ASC, Pos ASC")
        For intI = 0 To dtrDict.Length - 1
          ' Show where we are
          intPtc = (intI + 1) * 100 \ dtrDict.Length
          Status("Typing pass 4: " & intPtc & "%", intPtc)
          'If (dtrDict(intI).Item("t").ToString = "d") Then
          ' Get this lemma
          strLemma = dtrDict(intI).Item("l").ToString
          ' Validate
          ' If (strLemma = "blowan") Then Stop
          ' Get the corresponding lemma in MorphDict
          dtrFound = tdlMorphDict.Tables("Morph").Select("Vern='" & strLemma & "' AND l='" & strLemma & _
                                                         "' AND " & strPosField & "='" & dtrDict(intI).Item("Pos").ToString & "'")
          ' mark this form as the lemma
          If (dtrFound.Length > 0) Then
            ' Does it have a "d" mark?
            If (dtrFound(0).Item("t").ToString = "d") Then
              ' Change it
              dtrFound(0).Item("t") = "l"
            End If
          End If
          'End If
        Next intI
        ' Save the resulting Morphdict
        Status("Saving updated morphdict...")
        tdlMorphDict.WriteXml(strMorphDictFile)
      End If
      ' Prepare selection in order
      dtrFound = tdlMorphDict.Tables("Morph").Select("", "l ASC, t DESC, " & strPosField & " ASC, MorphId ASC, Vern ASC")
      ' Walk report 
      For intI = 0 To dtrFound.Length - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ dtrFound.Length
        Status("Processing " & intPtc & "%", intPtc)
        ' Access this entry
        With dtrFound(intI)
          ' Is this a new entry?
          If (strLemma <> .Item("l").ToString) Then
            ' Possibly add alternative forms for the previous lemma
            If (Not AddDictLemmaForms(strLemma, colForm, colHtml)) Then Return False
            ' Yes, this is a new lemma
            strLemma = .Item("l").ToString
            ' Default sense value
            strDef = "(no definition given)"
            ' Make new collection of forms
            colForm.Clear() : colForm.AddUnique(strLemma.Replace("-", "") & "_" & .Item(strPosField).ToString)
            ' Get current vern form 
            strVern = .Item("Vern").ToString.Replace("-", "")
            colForm.AddUnique(strVern & "_" & .Item(strPosField).ToString)
            ' Lookup and add the definition for this lemma
            dtrDict = tdlOEdict.Tables("Entry").Select("l='" & strLemma & "'")
            If (dtrDict.Length > 0) Then
              ' Get the sense
              dtrSense = tdlOEdict.Tables("Sense").Select("EntryId=" & dtrDict(0).Item("EntryId").ToString)
              If (dtrSense.Length > 0) Then
                ' Add the definition here in the table
                strDef = dtrSense(0).Item("Def").ToString
              End If
            End If
            ' Add lemma entry
            colHtml.Add("<tr><td><b><font color='darkblue'>" & strLemma & "</font></b></td><td>" & strLemma & "</td><td>" & .Item(strPosField).ToString & _
                        "</td><td><font color='red'>" & strDef & "</font></td></tr>")
          Else
            ' Get vern form 
            strVern = .Item("Vern").ToString.Replace("-", "")
            If (Not colForm.Exists(strVern & "_" & .Item(strPosField).ToString)) Then
              colForm.AddUnique(strVern & "_" & .Item(strPosField).ToString)
              ' This is a sub-entry
              colHtml.Add("<tr><td><font color='blue'>" & strLemma & "</font></td><td>" & .Item("Vern").ToString & "</td><td>" & .Item(strPosField).ToString & _
                          "</td><td><font color='green'>" & .Item("f").ToString & "</font></td></tr>")
            End If
          End If
        End With
      Next intI
      ' Possibly add alternative forms for the last lemma
      If (Not AddDictLemmaForms(strLemma, colForm, colHtml)) Then Return False
      ' Finish the report
      colHtml.Add("</table></body></html>")
      strHtml = colHtml.Text
      ' Save the remain report somewhere
      strDictRepFile = GetDocDir() & "\DictReport_" & strLang & ".htm"
      IO.File.WriteAllText(strDictRepFile, strHtml, System.Text.Encoding.UTF8)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/MakeDictReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   AddDictLemmaForms
  ' Goal:   Look in Morphdict.VernPos table for alternative forms not yet in [colForm]
  ' History:
  ' 18-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function AddDictLemmaForms(ByVal strLemma As String, ByRef colForm As StringColl, ByRef colHtml As StringColl) As Boolean
    Dim strCombi As String    ' Combinatino form
    Dim strPosL As String     ' POS of the lemma
    Dim strPattern As String   ' Pattern in terms of POS to look for
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim intI As Integer       ' Counter

    Try
      ' Validate
      If (colForm.Count = 0) Then Return True
      ' Get the lemma-part-of-speech
      strPosL = colForm.Item(0)
      If (InStr(strPosL, "_") > 0) Then
        strPosL = Mid(strPosL, InStr(strPosL, "_") + 1)
      End If
      ' Initialise
      strPattern = strPosL & "*"
      ' Determine the POS string to look for
      Select Case strPosL
        Case "VB"
          strPattern = "V*"
        Case "BE"
          strPattern = "B*"
        Case "MD"
          strPattern = "MD*"
        Case "AX"
          strPattern = "AX*"
        Case "N"
        Case "ADJ"
          strPattern = "ADJ*"
        Case "ADV"
          strPattern = "ADV*"
      End Select
      ' Look for results
      dtrFound = tdlMorphDict.Tables("VernPos").Select("l='" & strLemma & "' AND Pos LIKE '" & strPattern & "'", "Pos ASC")
      For intI = 0 To dtrFound.Length - 1
        ' Access entry
        With dtrFound(intI)
          ' Check if the entry is already present in the [colForm]
          strCombi = .Item("Vern") & "_" & .Item("Pos")
          If (Not colForm.Exists(strCombi)) Then
            ' Add this to [colForm]
            colForm.AddUnique(strCombi)
            ' Add this form to the output table, since it did not exist in [colForm] yet...
            colHtml.Add("<tr><td><font color='blue'>" & strLemma & "</font></td><td>" & _
                        .Item("Vern").ToString & "</td><td>" & .Item("Pos").ToString & _
                       "</td><td>(Derived)<font color='green'>" & .Item("f").ToString & "</font></td></tr>")
          End If
        End With
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/AddDictLemmaForms error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ReadMorphOEanalyzer
  ' Goal:   Try to read the "analyzer.sql" file from Ondrej Tichy and transform it into
  '           something I can use
  ' History:
  ' 08-02-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function ReadMorphOEanalyzer() As Boolean
    Dim tdlLemma As DataSet = Nothing ' lemma dictionary
    Dim tdlVern As DataSet = Nothing  ' vernacular dictionary
    Dim dtrThis As DataRow            ' One new datarow
    Dim dtrFound() As DataRow         ' Result of SELECT
    Dim arLine() As String            ' One line broken up
    Dim arVal() As String             ' Array of values
    Dim tdlAna As New DataSet         ' This will hold the results
    Dim strLine As String             ' One line
    Dim strOut As String              ' One output line
    Dim strCsv As String              ' One CSV line
    Dim strForm As String             ' One form
    Dim strOutFile As String          ' Actual output file
    Dim strSearch As String = ""      ' Search line
    Dim intI As Integer               ' Counter
    Dim intJ As Integer               ' COunter
    Dim intPtc As Integer             ' Percentage
    Dim intId As Integer              ' Unique ID per record
    Dim colLemmaBT As New StringColl        ' Collection with lemma-BT numbers
    Dim colForm As New StringColl           ' Collection of forms
    Dim rdThis As IO.StreamReader = Nothing ' Reader of the file
    Dim wrThis As IO.StreamWriter = Nothing ' Output
    Dim wrLem As IO.StreamWriter = Nothing  ' Lemma writer
    Dim wrCsv As IO.StreamWriter = Nothing  ' CSV writer
    Dim bIsCorpus As Boolean = False        ' type of output
    Dim bUseDataset As Boolean = False      ' Use dataset
    Dim strSqlFile As String = "d:\xampp\htdocs\analyzer.sql"
    Dim strFrmFile As String = "d:\xampp\htdocs\OEforms.xml"
    Dim strCrpFile As String = "d:\xampp\htdocs\OEcorpus.xml"
    Dim strLemFile As String = "d:\xampp\htdocs\OElemma.xml"
    Dim strCsvFile As String = "d:\xampp\htdocs\OElemma.csv"
    Dim strFormCsvFile As String = "d:\xampp\htdocs\OEforms.csv"

    Try
      ' Validate
      If (Not IO.File.Exists(strSqlFile)) Then Logging("SQL input file not found: " & strSqlFile) : Return False
      ' Ask for instructions
      Select Case MsgBox("Read analyzer.sql?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Cancel
          Return False
        Case MsgBoxResult.Yes
          ' Start reading
          rdThis = New IO.StreamReader(strSqlFile) : intI = 0 : intId = 0
          ' Find out what to do
          Select Case MsgBox("Would you like to extract info from [forms_u1](Y) or [corpus](N)?", MsgBoxStyle.YesNoCancel)
            Case MsgBoxResult.Cancel
              Return False
            Case MsgBoxResult.Yes
              bIsCorpus = False
            Case MsgBoxResult.No
              bIsCorpus = True
          End Select
          If (bIsCorpus) Then
            strSearch = "INSERT INTO `corpus` VALUES"
            strOutFile = strCrpFile
          Else
            strSearch = "INSERT INTO `forms_u1` VALUES"
            strOutFile = strFrmFile
          End If
          ' Start output
          If (Not CreateDataSet("BTlemma.xsd", tdlVern)) Then Return False
          dtrThis = Nothing
          wrThis = New IO.StreamWriter(strOutFile, False, System.Text.Encoding.UTF8)
          wrCsv = New IO.StreamWriter(strFormCsvFile, False, System.Text.Encoding.UTF8)
          strOut = "<BT>" : wrThis.WriteLine(strOut)
          While (Not rdThis.EndOfStream)
            ' Read one line
            strLine = rdThis.ReadLine
            ' Is this the line we want?
            If (InStr(strLine, strSearch) > 0) Then
              ' Read the values
              arLine = Split(strLine, "(")
              ' Show where we are
              intI += 1
              ' Status("Reading sql line #" & intI & " with " & arLine.Length & " entries")
              ' Process all values
              For intJ = 0 To arLine.Length - 1
                arVal = Split(arLine(intJ), ",")
                Application.DoEvents()
                ' where are we?
                intPtc = (intJ + 1) * 100 \ arLine.Length
                Status("Reading sql line #" & intI & " with " & arLine.Length & " entries " & intPtc & "%", intPtc)
                ' Is this a good one?
                If (Not bIsCorpus) AndAlso (arVal.Length > 18) Then
                  ' Process this entry
                  ' 1  = vernacular
                  ' 3  = lemma
                  ' 2  = BT identifier
                  ' 9  = feature(s)
                  ' 13 = POS
                  ' strOut = arVal(1) & ";" & arVal(2) & ";" & arVal(9)
                  intId += 1
                  If (bUseDataset) Then
                    arVal(1) = arVal(1).Replace("’", "").Replace("'", "")
                    arVal(9) = arVal(9).Replace("’", "").Replace("'", "")
                    ' Does it already exist?
                    If (tdlVern.Tables("V").Select("v='" & arVal(1).Replace("'", "''") & _
                                                   "' AND bt=" & arVal(2) & " AND f='" & arVal(9) & "'").Length = 0) Then
                      ' More normalisation
                      arVal(3) = arVal(3).Replace("’", "").Replace("'", "")
                      arVal(13) = arVal(13).Replace("’", "").Replace("'", "")
                      If (Not CreateNewRow(tdlVern, "V", "id", intId, dtrThis, _
                          intId, arVal(1), arVal(3), arVal(9), arVal(2), arVal(13))) Then Return False
                    End If
                  Else
                    strOut = "<E id=" & """" & intId & """" & _
                              " v=" & arVal(1).Replace("’", "'") & _
                              " bt=" & """" & arVal(2) & """" & _
                              " f=" & arVal(9).Replace("’", "'") & " />"
                    wrThis.WriteLine(strOut)
                    ' Process the BT-identifier and the lemma
                    colLemmaBT.AddUnique(arVal(2), arVal(3) & ";" & arVal(13))
                    ' Also output to csv output writer 
                    strCsv = arVal(1).Replace("’", "'").Replace("'", """") & "," & _
                              arVal(3).Replace("’", "'").Replace("'", """") & "," & _
                              """" & "http://www.bosworthtoller.com/" & Format(CInt(arVal(2)), "000000") & """" & "," & _
                              arVal(9).Replace("’", "'").Replace("'", """") & "," & _
                              arVal(13).Replace("’", "'").Replace("'", """")
                    wrCsv.WriteLine(strCsv)
                  End If
                ElseIf (bIsCorpus) AndAlso (arVal.Length > 3) Then
                  ' Process this entry: n, word, freq, texts
                  intId += 1
                  arVal(1) = arVal(1).Replace("’", """")
                  arVal(1) = arVal(1).Replace("'", """")
                  strOut = "<C id=" & """" & intId & """" & _
                              " n=" & """" & arVal(0) & """" & _
                              " w=" & arVal(1).Replace("’", """") & _
                              " f=" & """" & arVal(2) & """" & _
                              " t=" & """" & arVal(3) & """" & " />"
                  wrThis.WriteLine(strOut)
                  ' Process the BT-identifier and the lemma
                  colLemmaBT.AddUnique(arVal(1) & "_" & arVal(0))
                End If
              Next intJ
            End If
          End While
          ' Finish table
          strOut = "</BT>" : wrThis.WriteLine(strOut)
          wrThis.Close() : wrCsv.Close()
          ' Have we used a dataset?
          If (bUseDataset) Then
            ' Write the results
            tdlVern.WriteXml(strFrmFile)
          End If
          If (Not bIsCorpus) AndAlso (Not bUseDataset) Then
            ' Start output
            wrThis = New IO.StreamWriter(strLemFile, False, System.Text.Encoding.UTF8)
            strOut = "<BT>" : wrThis.WriteLine(strOut)
            ' Write the BT to lemma identification
            For intI = 0 To colLemmaBT.Count - 1
              ' Where are we?
              intPtc = (intI + 1) * 100 \ colLemmaBT.Count
              Status("Lemma to BT " & intPtc & "%", intPtc)
              ' Process this lemma-BT entry
              arVal = Split(colLemmaBT.Exmp(intI), ";")
              strOut = "<L id=" & """" & intI + 1 & """" & _
                        " bt=" & """" & colLemmaBT.Item(intI) & """" & _
                        " l=" & arVal(0).Replace("’", "'").Replace("'", """") & _
                        " ps=" & arVal(1).Replace("’", "'").Replace("'", """") & " />"
              wrThis.WriteLine(strOut)
            Next intI
            ' Finish table
            strOut = "</BT>" : wrThis.WriteLine(strOut)
            ' Close 
            rdThis.Close()
            wrThis.Close()
          End If
      End Select
      ' Ask for follow-up
      Select Case MsgBox("Would you like to sort the lemma dictionary into a csv file?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Cancel
          Return False
        Case MsgBoxResult.Yes
          ' Load the lemma dictionary
          If (Not ReadDataset("BTlemma.xsd", strLemFile, tdlLemma)) Then Status("Could not read lemma dictionary") : Return False
          ' Sort the output
          dtrFound = tdlLemma.Tables("L").Select("", "l ASC, bt ASC")
          wrThis = New IO.StreamWriter(strCsvFile, False, System.Text.Encoding.UTF8)
          ' Write the otuput
          For intI = 0 To dtrFound.Length - 1
            ' Access this member
            With dtrFound(intI)
              ' Prepare output string
              strOut = """" & .Item("l").ToString & """" & "," & """" & .Item("ps").ToString & """" & "," & .Item("bt").ToString & "," & _
                """" & "http://www.bosworthtoller.com/" & Format(CInt(.Item("bt").ToString), "000000")
              wrThis.WriteLine(strOut)
            End With
          Next intI
          wrThis.Close()
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/ReadMorphOEanalyzer error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Close 
      rdThis.Close()
      wrThis.Close()
      ' Return failure
      Return False
    End Try
  End Function
End Module

