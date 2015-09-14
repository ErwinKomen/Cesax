Imports System.Xml
Module modConLLX
  ' ===================================== GLOBAL VARIABLES ==========================================================
  Public dep_strCanBeHeadNode As String = "./child::eTree[" & _
        "count(child::eLeaf[@Type='Punct' or @Type='Zero'])=0 and @Label!='CODE'" & _
        " and (tb:matches(@Label, '*LFD*') " & _
        "      or (count(child::eLeaf[@Type='Vern'])>0) " & _
        "      or not(tb:matches(@Label, '" & loc_strLabelNumber & "')))]"
  Public tdlLiftPosDic As DataSet = Nothing   ' Possible <liftpos> dictionary for language
  Public objPosTag As PosTag = Nothing        ' POS-tagging bassed on aggluspell for this language
  ' ===================================== LOCAL CONSTANTS ===========================================================
  Private Const MALT_PARSER As String = "maltparser-1.7.2.jar"
  Private Const MALT_CHECHEN As String = "chechen.mco"
  Private Const MALT_SPANISH As String = "spanish.mco"
  Private Const MALT_OLDDUTCH As String = "nederlands.mco"
  Private Const MALT_LIB_1 As String = "liblinear-1.8.jar"
  Private Const MALT_LIB_2 As String = "libsvm.jar"
  Private Const MALT_LIB_3 As String = "log4j.jar"
  Private Const MALT_INTRO As String = "-jar " & """"
  Private Const MALT_LOC As String = " -c $mco -m parse -i $inp -o $out"
  Private Const MALT_MCO As String = " -a nivreeager -c $mco -i $inp -m learn"
  ' ===================================== LOCAL VARIABLES ===========================================================
  Private Structure CONLLX
    Dim Id As Integer
    Dim Form As String
    Dim Lemma As String
    Dim CposTag As String
    Dim PosTag As String
    Dim Feats As String
    Dim Head As String    ' Use STRING here to allow _
    Dim DepRel As String
    Dim Phead As String   ' Use STRING here to allow _
    Dim PdepRel As String
  End Structure
  Private Structure HdRuleSpec
    Dim phrase As String  ' Phrase label
    Dim dir As String     ' Direction: l or r
    Dim hdlist As String  ' semi-colon separated list of possible heads
  End Structure
  Private Structure DepRelSpec
    Dim rel As String     ' Relation string
    Dim par As String     ' Parent node 
    Dim lablist As String ' vertical bar separated list of possible (functional) labels
  End Structure
  Private Structure CompactPos
    Dim Cpos As String    ' Compact POS
    Dim poslist As String ' vertical bar separate list of possible POS tags
  End Structure
  Private Structure MorphPos
    Dim Pos As String       ' Regular POS
    Dim morphlist As String ' vertical bar separate list to recognize morphemes
    Dim poslist As String   ' Vertical bar separated list to recognize POS used in the original
  End Structure
  Private Structure PosRule
    Dim Pos As String     ' Articulate POS
    Dim Cpos As String    ' Compact POS (converted to upper case)
    Dim ftlist As String  ' List of features
    Dim vern As String    ' Vernacular specifications (optional; converted to lower case)
  End Structure
  Private Structure PhraseRule
    Dim phrase As String  ' Phrase label
    Dim poslist As String ' List of vertical-bar separated possible POS tags
    Dim dep As String     ' Dependency relation that is required (one)
  End Structure
  Private strHttpDir As String = "http://erwinkomen.ruhosting.nl/software/stp/"
  Private strMaltParser As String = ""
  Private strMaltShellIntro As String = ""
  Private strMaltShellCreateMco As String = ""
  Private pdxConv As XmlDocument        ' One XML document new 
  Private tdlConv As DataSet = Nothing  ' One dataset
  Private arRootNode() As String = {"N", "V", "S", "A", "F", "Z", "W", "C", "P", "D", "R", """"}
  Private arRootLink() As String = {"NP", "IP-MAT", "PP", "ADJP", ".", "NUMP", "WP", "CONJP", "NP", "DP", "RP", "."}
  Private arVerbPos() As String = { _
    "VMIP*", "VBP", "VSIP*", "AXP", "VAIP*", "BEP", _
    "VMII*|VMIS*", "VBD", "VSII*|VSIS*", "AXD", "VAII*|VAIS*", "BED", _
    "VMIF*", "VBF", "VSIF*", "AXF", "VAIF*", "BEF", _
    "VMIC*", "VBC", "VSIC*", "AXC", "VAIC*", "BEC", _
    "VMSP*", "VBPS", "VSSP*", "AXPS", "VASP*", "BEPS", _
    "VMSI*", "VBDS", "VSSI*", "AXDS", "VASI*", "BEDS", _
    "VMSF*", "VBFS", "VSSF*", "AXFS", "VASF*", "BEFS", _
    "VMG*", "VAG", "VSG*", "AXG", "VAG*", "BAG", _
    "VMM*", "VBI", "VSM*", "AXI", "VAM*", "BEI", _
    "VMN*", "VB", "VSN*", "AX", "VAN*", "BE", _
    "VMP*", "VAN", "VSP*", "AXN", "VAP*", "BEN"}
  Private arHeadRule_lng() As String = {"nl", "HeadRule_du.txt", _
                                        "en", "HeadRule_en.txt", _
                                        "che", "HeadRule_ch.txt", _
                                        "es", "HeadRule_es.txt", _
                                        "lez", "HeadRule_lez.txt", _
                                        "lbe-Latn", "HeadRule_lbe.txt", _
                                        "lbe-Cyr", "HeadRule_lbe.txt"}
  Private loc_strLabelNumber As String = "*-[123456789]|*-[123456789][0123456789]"
  Private arHeadRule() As HdRuleSpec            ' Head dependency rules
  Private arDepRel() As DepRelSpec              ' Dependency relations list
  Private arCposRule() As CompactPos            ' Translation from POS to CPOS
  Private arPosRule() As PosRule                ' Derive POS from CPOS and features
  Private arPhraseRule() As PhraseRule          ' List of phrase rules
  Private arMorphPos() As MorphPos              ' List of translation from morpheme to POS
  Private bLangInit As Boolean = False          ' Language initialised or not?
  Private objXpFun As New XPathFunctions        ' Link to Xpath functions
  Private colDoHdStack As New StringColl        ' Collection of <eTree> Id values 
  Private colDoHdDupli As New StringColl        ' Collection of allowed duplicates
  Private bDuplicating As Boolean = False       ' Duplication mode
  Private loc_intCurrentForestId As Integer     ' forestId of currently selected item
  Private loc_ndxCurrentForest As XmlNode = Nothing ' Current forest
  ' ==========================================================================================================
  '-----------------------------------------------------------------------------------------------------------
  ' Name:       ExpandOneConllToPsdx()
  ' Goal:       Expand the syntactic analysis of one file in CONLLX format to the currently loaded psdx file
  ' History:
  ' 17-09-2013  ERK Created
  ' 19=06-2014  ERK Added bDoDep flag
  '-----------------------------------------------------------------------------------------------------------
  Public Function ExpandOneConllToPsdx(ByVal strConFile As String, ByVal strLang As String, _
                                       Optional ByVal bDepToConst As Boolean = True) As Boolean
    Dim strText As String = ""    ' String of file content
    Dim strBrk As String = ""     ' Break symbol used
    Dim arText() As String        ' File broken up into lines
    Dim ndxFor As XmlNode         ' Current forest
    Dim intI As Integer           ' Counter

    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Logging("ExpandOneConllToPsdx: no pdxCurrentFile") : Return False
      If (Not IO.File.Exists(strConFile)) Then Logging("ExpandOneConllToPsdx: file not found - " & strConFile) : Return False
      ' Make sure language initialisation has taken place
      If (Not DoLangInit(strLang)) Then Logging("ExpandOneConllToPsdx: could not initialise language " & strLang) : Return False
      ' Check there are phrase rules
      If (arPhraseRule Is Nothing) Then Logging("No phrase rules defined for [" & strLang & "]") : Return False
      ' Make sure current document is treated
      SetXmlDocument(pdxCurrentFile)
      ' Read the file into a string
      strText = IO.File.ReadAllText(strConFile)
      strBrk = GetDelim(strText, vbCrLf, vbCr, vbLf)
      arText = Split(strText, strBrk)
      ' Find the first forest
      ndxFor = pdxCurrentFile.SelectSingleNode("//forest")
      If (ndxFor Is Nothing) Then Logging("ExpandOneConllToPsdx: no first forest found") : Return False
      ' Start reading the array
      For intI = 0 To arText.Length - 1
        ' Validate
        If (ndxFor Is Nothing) AndAlso (arText(intI) <> "") Then
          ' There is a problem: we are expecting a section, including a forest, but it is no there
          Stop
          Logging("modConLLX/ExpandOneConllToPsdx: there is a mismatch between <forest> nodes and sections")
          bInterrupt = True
          Return False
        End If
        ' Check if this starts a section
        If (InStr(arText(intI), "1" & vbTab) = 1) Then
          ' Process this forest
          If (Not OneConllxToPsdxForest(ndxFor, arText, intI, strLang, True, bDepToConst)) Then Return False
          ' Go to the next forest
          ndxFor = ndxFor.NextSibling
        End If
      Next intI
      ' Re-calculate the <eTree>@Id values
      AdaptEtreeId(1)
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/ExpandOneConllToPsdx error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '-----------------------------------------------------------------------------------------------------------
  ' Name:       ExpandOneConllDepToPsdx()
  ' Goal:       Expand the syntactic analysis of one file in (already in psdx) to the currently loaded psdx file
  ' History:
  ' 19-06-2014  ERK created
  '-----------------------------------------------------------------------------------------------------------
  Public Function ExpandOneConllDepToPsdx(ByVal strLang As String, Optional ByVal bDoDep As Boolean = True) As Boolean
    Dim strText As String = ""    ' String of file content
    Dim strBrk As String = ""     ' Break symbol used
    Dim ndxFor As XmlNode         ' Current forest

    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Logging("ExpandOneConllDepToPsdx: no pdxCurrentFile") : Return False
      ' Make sure language initialisation has taken place
      If (Not DoLangInit(strLang)) Then Logging("ExpandOneConllDepToPsdx: could not initialise language " & strLang) : Return False
      ' Check there are phrase rules
      If (arPhraseRule Is Nothing) Then Logging("No phrase rules defined for [" & strLang & "]") : Return False
      ' Make sure current document is treated
      SetXmlDocument(pdxCurrentFile)
      ' Find the first forest
      ndxFor = pdxCurrentFile.SelectSingleNode("//forest")
      If (ndxFor Is Nothing) Then Logging("ExpandOneConllToPsdx: no first forest found") : Return False
      ' Start reading the array
      While (ndxFor IsNot Nothing)
        ' Process this forest
        If (Not OneConllDepToPsdxForest(ndxFor, strLang)) Then Return False
        ' Go to the next forest
        ndxFor = ndxFor.NextSibling
       End While
      ' Re-calculate the <eTree>@Id values
      AdaptEtreeId(1)
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/ExpandOneConllDepToPsdx error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
 
  '----------------------------------------------------------------------------------------
  ' Name:       ConvertOneConllxToPsdx()
  ' Goal:       Convert one file in CONLLX format to destination (psdx)
  ' History:
  ' 16-08-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function ConvertOneConllxToPsdx(ByVal strSrcFile As String, ByVal strDstFile As String, _
      Optional ByVal strLang As String = "none") As Boolean
    Dim strText As String = ""    ' String of file content
    Dim strBrk As String = ""     ' Break symbol used
    Dim strShort As String = ""   ' Short file name
    Dim strTextId As String = ""  ' Text ID
    Dim strType As String = ""    ' Type of eLeaf
    Dim strValue As String = ""   ' Value of <eLeaf>
    Dim strLabel As String = ""   ' Label value
    Dim strFunct As String = ""   ' Function
    Dim arText() As String        ' File broken up into lines
    Dim arSect() As CONLLX        ' Array of CONLLX elements
    Dim ndxWork As XmlNode        ' Working node
    Dim ndxFor As XmlNode         ' Current forest
    Dim ndxForGrp As XmlNode      ' Pointer to <forestGrp>
    Dim intForestId As Integer = 1  ' ID of the forest
    Dim intEtreeId As Integer = 1 ' ID of <eTree>
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intId As Integer = 0      ' DUmmy

    Try
      ' Read the file into a string
      strText = IO.File.ReadAllText(strSrcFile)
      ' strBrk = GetDelim(strText, vbCrLf, vbCr, vbLf)
      ' Determining linebreak
      strBrk = GetDelimDeep(strText)
      arText = Split(strText, strBrk)
      ' Create a new xml document and set the header to initial values
      If (Not CreatePsdxHeader(pdxConv, strSrcFile, strShort)) Then Return False
      ' Do we need to ask for a language?
      If (strLang = "ask") Then
        ' Yes, ask for a language
         With frmConvert
          Select Case .ShowDialog
            Case Windows.Forms.DialogResult.OK
              strLang = LangToEthno(.Language)
            Case Else
              Return False
          End Select
        End With
      End If
      ' Initialise
      ' strShort = IO.Path.GetFileNameWithoutExtension(strSrcFile)
      strTextId = strShort.Replace(" ", "")
      If (Not DoLangInit(strLang)) Then Return False

      ndxForGrp = pdxConv.SelectSingleNode(".//forestGrp")
      ' Walk through the text
      For intI = 0 To arText.Length - 1
        ' Check if this starts a section
        If (InStr(arText(intI), "1" & vbTab) = 1) Then
          ' Start new CONLL section
          ReDim arSect(0) : intJ = 0
          ' Create a new <forest> element and add its properties
          ndxFor = AddXmlChild(ndxForGrp, "forest", "forestId", intForestId, "attribute", _
                               "TextId", strTextId, "attribute", _
                               "File", strShort & ".psdx", "attribute", _
                               "Location", strTextId & "." & Format(intForestId, "0000"), "attribute")
          intForestId += 1
          ' Add the <divs>: one for Spanish (org) and one for English (to be added)
          ndxWork = AddXmlChild(ndxFor, "div", "lang", "org", "attribute", "seg", "", "child")
          ndxWork = AddXmlChild(ndxFor, "div", "lang", "eng", "attribute", "seg", "", "child")
          ' Process this <forest>
          If (Not OneConllxToPsdxForest(ndxFor, arText, intI, strLang, False)) Then Return False
        End If
      Next intI
      ' Set the current one
      InitCurrentFile()
      ' Re-calculate the <eTree>@Id values
      AdaptEtreeId(1)
      ' Save dataset as file
      pdxConv.Save(strDstFile)
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/ConvertOneConllxToPsdx error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       OneConllxToPsdxForest()
  ' Goal:       Convert one forest section in [arText] and process it in [ndxFor]
  ' History:
  ' 17-09-2013  ERK Created
  ' 19-06-2014  ERK Added [bDepToConst] variable
  '----------------------------------------------------------------------------------------
  Private Function OneConllxToPsdxForest(ByRef ndxFor As XmlNode, ByRef arText() As String, ByRef intI As Integer, _
      ByVal strLang As String, ByVal bSynchronize As Boolean, Optional ByVal bDepToConst As Boolean = True) As Boolean
    Dim ndxWork As XmlNode        ' Working node
    Dim ndxLeaf As XmlNode        ' Leaf child
    Dim ndxHead As XmlNode        ' My head
    '   Dim ndxList As XmlNodeList    ' List of nodes
    Dim ndxChild As XmlNodeList   ' <eTree> children of <ndxFor> when bSynchronize is true
    Dim strType As String = ""    ' Type of eLeaf
    Dim strValue As String = ""   ' Value of <eLeaf>
    Dim arSect() As CONLLX        ' Array of CONLLX elements
    Dim arLine() As String        ' One line in pieces
    Dim intEtreeId As Integer = 1 ' ID of <eTree>
    Dim intJ As Integer           ' Counter
    Dim intK As Integer           ' Counter
    Dim intChild As Integer       ' Which child are we?
    Dim strHead As String         ' Number of my head

    Try
      ' Validate
      If (ndxFor Is Nothing) OrElse (ndxFor.Name <> "forest") Then Return False
      ' Are we synchronizing?
      If (bSynchronize) Then
        ' Read the list of children
        ndxChild = ndxFor.SelectNodes("./child::eTree")
      Else
        ndxChild = Nothing
      End If
      ' Note where we are
      intChild = 0
      ' This is the start of a new sentence (=section) --> read it 
      While (arText(intI) <> "")
        ' Read this line
        arLine = Split(arText(intI), vbTab)
        ' Check how many columns we have
        If (arLine.Length <> 10 AndAlso arLine.Length <> 8) Then
          ' This is not CONLLX
          MsgBox("This file is not CONLL-X format, since it has [" & arLine.Length & "] columns instead of 8 or 10")
          bInterrupt = True
          Return False
        End If
        ' ============ DEBUG ==================
        ' If (arLine(0) = "115") Then Stop
        ' =====================================
        ' Disperse into CONLL format
        ReDim Preserve arSect(0 To intJ)
        With arSect(intJ)
          .Id = arLine(0) : .Form = arLine(1) : .Lemma = arLine(2)
          .CposTag = arLine(3) : .PosTag = arLine(4) : .Feats = arLine(5)
          .Head = arLine(6) : .DepRel = arLine(7)
          If (arLine.Length = 8) Then
            .Phead = "_" : .PdepRel = "_"
          Else
            .Phead = arLine(8) : .PdepRel = arLine(9)
          End If
          '' ===== DEBUG =======
          'If (intForestId = 29) And (intJ = 2) Then
          '  Stop
          '  Debug.Print(AscW(arLine(1)))
          'End If
          '' ===================
          ' Make sure we do not process empty forms
          If (.Form = "") OrElse (.Form = " ") OrElse (.Form = vbTab) Then .Form = "*e*"
          ' Determine type of <eLeaf>
          Select Case .Form
            Case "(", "["
              strType = "Punct" : strValue = "("
            Case ")", "]"
              strType = "Punct" : strValue = ")"
            Case ".", "!", "?"
              strType = "Punct" : strValue = "."
            Case ",", ":", ";", "-"
              strType = "Punct" : strValue = ","
            Case """", "'", "«", "»"
              strType = "Punct" : strValue = """"
            Case Else
              If (Left(.Form, 1) = "*") Then
                ' EMpty node
                strType = "Star" : strValue = ""
              Else
                strType = "Vern" : strValue = ""
              End If
          End Select
          ' Are we synchronizing?
          If (bSynchronize) Then
            ' Get the correct child
            ndxWork = ndxChild(intChild) : ndxLeaf = ndxWork.SelectSingleNode("./child::eLeaf")
            ' Check if this is the correct POS and text
            If (ndxWork.Attributes("Label").Value <> .PosTag) OrElse _
               (ndxLeaf Is Nothing) OrElse (ndxLeaf.Attributes("Text").Value <> .Form) Then
              ' Bad synchronization
              Logging("OneConllxToPsdxForest: synchronization problem at child=" & intChild)
              bInterrupt = True
              Return False
            End If
          Else
            ' We are CREATING, not synchronizing: add one XML child under [ndxFor]
            If (strType = "Punct") AndAlso (strValue <> "") Then
              ndxWork = AddEtreeChild(ndxFor, intEtreeId, strValue, 0, 0)
            Else
              ndxWork = AddEtreeChild(ndxFor, intEtreeId, UCase(.CposTag), 0, 0)
            End If
            intEtreeId += 1
            ' Add <eLeaf> to the <eTree> node
            AddEleafChild(ndxWork, strType, .Form, 0, 0)
          End If
          ' Add features to the <eTree> node
          If (Not AddFeatureXml(ndxWork, "con", "drel", .DepRel)) Then Return False
          If (Not AddFeatureXml(ndxWork, "con", "l", .Lemma)) Then Return False
          If (Not AddFeatureXml(ndxWork, "con", "pos", .PosTag)) Then Return False
          If (Not AddFeatureXml(ndxWork, "con", "f", .Feats)) Then Return False
          If (Not AddFeatureXml(ndxWork, "con", "hd", .Head)) Then Return False
          If (Not AddFeatureXml(ndxWork, "con", "id", .Id)) Then Return False
        End With
        ' Increment intJ
        intJ += 1
        ' Continue with intI
        intI += 1
        ' Go to next child
        intChild += 1
      End While
      ' We have read one sentence into CONLL, now analyze and add this to the psdx

      '' Dependency to tree SIMPLE:
      '' Visit all nodes in order and position them under the indicate head
      'ndxChild = ndxFor.SelectNodes("./child::eTree")
      'For intK = 0 To ndxChild.Count - 1
      '  ' Find out what my head is
      '  strHead = GetFeature(ndxChild(intK), "con", "hd")
      '  If (strHead = "") Then Stop
      '  ndxHead = ndxFor.SelectSingleNode("./descendant::eTree[child::fs[@type='con']/child::f[@name='id']/@value = '" & strHead & "']")
      '  If (ndxHead IsNot Nothing) Then
      '    ndxHead.AppendChild(ndxChild(intK))
      '  End If
      'Next intK

      ' Continue with Dependency to Constituency??
      If (Not bDepToConst) Then Return True
      If (Not OneConllDepToPsdxForest(ndxFor, strLang)) Then Return False
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/OneConllxToPsdxForest error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       OneConllDepToPsdxForest()
  ' Goal:       Convert one forest section (within psdx) from dependency to constituency
  ' History:
  ' 19-06-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function OneConllDepToPsdxForest(ByRef ndxFor As XmlNode, ByVal strLang As String) As Boolean
    Dim ndxList As XmlNodeList    ' List of nodes
    Dim ndxWork As XmlNode        ' Working node
    Dim strType As String = ""    ' Type of eLeaf
    Dim strValue As String = ""   ' Value of <eLeaf>
    Dim intEtreeId As Integer = 1 ' ID of <eTree>
    Dim intJ As Integer           ' Counter

    Try
      ' (1) Check all ROOT <eTree> constituents
      ndxList = ndxFor.SelectNodes("./descendant::eTree[child::fs/child::f[@name='hd' and @value='0']]")
      For intJ = 0 To ndxList.Count - 1
        ' Process this rootnode recursively
        If (Not OneConllxNode(ndxList(intJ), ndxFor, strLang)) Then Return False
      Next intJ
      '' Process mono phrases
      'If (Not OneConllxMonoPhr(ndxFor, strLang)) Then Return False
      ' Check if V adaptation is needed
      If (strLang <> "none") Then
        ' Visit all the children with V-label
        ndxList = ndxFor.SelectNodes("./descendant::eTree[@Label = 'V']")
        For intJ = 0 To ndxList.Count - 1
          ndxWork = ndxList(intJ)
          ' Adapt the label for the V-node
          ndxWork.Attributes("Label").Value = ConLLxAdaptVerbLabel(GetFeature(ndxWork, "con", "pos"), strLang)
        Next intJ
        Select Case strLang
          Case "che", "ch"
            ' Look for all CP-RELs
            ndxList = ndxFor.SelectNodes("./descendant::eTree[@Label = 'CP-REL']")
            For intJ = 0 To ndxList.Count - 1
              ' (1) If the following sibling is an NP
              ndxWork = ndxList(intJ).NextSibling
              If (ndxWork IsNot Nothing) Then
                If (ndxWork.Attributes("Label").Value Like "NP*") Then
                  ' I must go under it
                  ndxWork.PrependChild(ndxList(intJ))
                End If
              End If
              ' (2) Check it is under an NP
              If (Not DoLike(ndxList(intJ).ParentNode.Attributes("Label").Value, "NP*")) Then
                ' Insert an NP level above the CP
                ndxWork = Nothing
                If (Not eTreeInsertLevel(ndxList(intJ), ndxWork)) Then Return False
                ' ============ DEBUG: Test =====================
                If (Not TestEndNodeIdOrder(ndxWork)) Then
                  Debug.Print("OneConllxNode: bad order")
                  Logging("modConLLX/OneConllxToPsdxForest warning: bad constituent order at CP-REL forestId=" & _
                          ndxFor.Attributes("forestId").Value)
                End If
                ' ==============================================
                ' Set the phrase type
                ndxWork.Attributes("Label").Value = "NP-NOM"
              End If
              ' (3) Check out the IP-SUB under the CP-REL
              ndxWork = ndxList(intJ).SelectSingleNode("./child::eTree[@Label = 'IP-SUB']")
              If (ndxWork IsNot Nothing) Then
                ' Move any preceding siblings under the IP
                While (ndxWork.PreviousSibling IsNot Nothing)
                  ndxWork.PrependChild(ndxWork.PreviousSibling)
                End While
              End If
            Next intJ
        End Select
      End If
      ' Make sure the sentence is re-analyzed
      ndxWork = Nothing
      eTreeSentence(ndxFor, ndxWork)
      ' Return positively
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/OneConllDepToPsdxForest error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       GetPhraseRule()
  ' Goal:       Derive the phrase label from [strPos] and [strDep]
  ' History:
  ' 05-09-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function GetPhraseRule(ByVal strPos As String, ByVal strDep As String) As String
    Dim intI As Integer ' Counter

    Try
      ' Validate
      If (strPos = "") Then Return ""
      If (arPhraseRule Is Nothing) Then Return ""
      strDep = LCase(strDep)
      ' Walk the list
      For intI = 0 To arPhraseRule.Length - 1
        With arPhraseRule(intI)
          ' See if the pOS fits
          If (DoLike(strPos, .poslist)) AndAlso (DoLike(strDep, .dep)) Then
            ' Return the phrase
            Return .phrase
          End If
        End With
      Next intI
      ' Getting here means failure...
      Return ""
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/GetPhraseRule error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return ""
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       GetPhraseRuleRev()
  ' Goal:       Derive the phrase label from [strPos] and [strDep], but only if it is needed
  ' History:
  ' 17-09-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function GetPhraseRuleRev(ByVal strPos As String, ByVal strDep As String) As String
    Dim intI As Integer ' Counter

    Try
      ' Validate
      If (strPos = "") Then Return ""
      If (arPhraseRule Is Nothing) Then Return ""
      strDep = LCase(strDep)
      ' Walk the list
      For intI = 0 To arPhraseRule.Length - 1
        With arPhraseRule(intI)
          ' See if the pOS fits
          If (.dep <> "*") AndAlso (DoLike(strPos, .poslist)) AndAlso (DoLike(strDep, .dep)) Then
            ' Return the phrase
            Return .phrase
          End If
        End With
      Next intI
      ' Walk the list again, but now with ".dep" equal to "*"
      For intI = 0 To arPhraseRule.Length - 1
        With arPhraseRule(intI)
          ' See if the pOS fits
          If (.dep = "*") AndAlso (DoLike(strPos, .poslist)) AndAlso (DoLike(strDep, .dep)) Then
            ' Return the phrase
            Return .phrase
          End If
        End With
      Next intI
      ' Getting here means failure...
      Return ""
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/GetPhraseRuleRev error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return ""
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       OneConllxMonoPhr()
  ' Goal:       Process phrases that only have one child
  ' History:
  ' 17-09-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function OneConllxMonoPhr(ByRef ndxFor As XmlNode, ByVal strLang As String) As Boolean
    Dim ndxThis As XmlNode = Nothing  ' Me
    Dim ndxWork As XmlNode = Nothing  ' Node
    Dim ndxChild As XmlNode = Nothing ' One child
    Dim ndxList As XmlNodeList        ' List of nodes
    Dim strNodeDrel As String         ' Value of feature "drel"
    Dim strPhrase As String           ' Phrase label
    Dim intI As Integer               ' Counter
    Dim intJ As Integer               ' Counter

    Try
      ' Validate
      If (ndxFor Is Nothing) Then Return False
      If (ndxFor.Name <> "forest") Then Return False
      ' Walk all the phrase rules
      For intI = 0 To arPhraseRule.Length - 1
        ' Check out dependency relation (take the first one)
        strNodeDrel = Split(arPhraseRule(intI).dep, "|")(0)
        ' Is this something?
        If (strNodeDrel <> "*") Then
          ' Get the phrase
          strPhrase = arPhraseRule(intI).phrase
          ' Okay, this is substantial -- find all nodes with this relation, and no <eTree> children
          ndxList = ndxFor.SelectNodes("./descendant::eTree[child::fs/child::f[@name='drel' and @value='" & strNodeDrel & "']" & _
                                       " and count(child::eTree) = 0" & _
                                       " and tb:matches(@Label, '" & arPhraseRule(intI).poslist & "')]", conTb)
          ' Walk them in reverse order
          For intJ = ndxList.Count - 1 To 0 Step -1
            ' Get me
            ndxThis = ndxList(intJ)
            ' Insert a level above me
            If (Not eTreeInsertLevel(ndxThis, ndxWork)) Then Return False
            ' Make [ndxWork] the new [ndxThis]
            ndxThis = ndxWork
            ' Set the phrase type
            ndxThis.Attributes("Label").Value = strPhrase
          Next intJ
        End If
      Next intI
      ' Return positively 
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/OneConllxMonoPhr error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       OneConllxNode()
  ' Goal:       Process this head
  ' History:
  ' 19-08-2013  ERK Created
  ' 04-09-2013  ERK Added [strLang] extension
  '----------------------------------------------------------------------------------------
  Private Function OneConllxNode(ByRef ndxThis As XmlNode, ByRef ndxFor As XmlNode, _
                                 ByVal strLang As String) As Boolean
    Dim strLabel As String            ' Label of me
    Dim ndxWork As XmlNode = Nothing  ' Node
    Dim ndxChild As XmlNode = Nothing ' One child
    Dim ndxSave As XmlNode = Nothing  ' Additional working node
    Dim ndxList As XmlNodeList        ' List of nodes
    Dim ndxDepList As XmlNodeList     ' List of nodes that are dependent upon me
    Dim ndxFoll As XmlNode            ' Element before which I should be inserted
    Dim intL As Integer               ' Counter
    Dim intI As Integer               ' Counter
    Dim intFeatId As Integer          ' Value of feature "id"
    Dim intHeadId As Integer          ' Value of feature "hd"
    Dim intChildId As Integer         ' Value of feature "id" for ndxChild
    Dim strNodePos As String          ' Value of feature "pos"
    Dim strNodeDrel As String         ' Value of feature "drel"
    Dim strPhrase As String           ' Phrase label

    Try
      ' Validate
      If (ndxThis Is Nothing) OrElse (ndxFor Is Nothing) Then Return False
      ' Capture my ID
      intFeatId = GetFeature(ndxThis, "con", "id")
      intHeadId = GetFeature(ndxThis, "con", "hd")
      ' Select all nodes that have my id as their head
      ndxList = ndxFor.SelectNodes("./descendant::eTree[child::fs/child::f[@name='hd' and @value='" & intFeatId & "']]")
      ' Get my own label
      strLabel = ndxThis.Attributes("Label").Value
      ' ============= DEBUG
      'If (ndxFor.Attributes("forestId").Value = 32) AndAlso (strLabel Like "VAGA-*") Then
      '  Stop
      'End If
      ' =======================
      Select Case strLang
        Case "nl", "du"
          ' Adapt my own label at any rate
          ndxThis.Attributes("Label").Value = GetFeature(ndxThis, "con", "pos")
      End Select
      ' Additional levels depend on the number of nodes having me as head
      If (ndxList.Count = 0) Then
        ' I am not the head of anyone else!
        ' Depending on who I am, I may not need an additional head
        ' I need to have my POS type
        strNodePos = GetFeature(ndxThis, "con", "pos")
        ' Safeguard the @con/drel feature
        strNodeDrel = GetFeature(ndxThis, "con", "drel")
        ' Determine the phrase label 
        strPhrase = GetPhraseRuleRev(strNodePos, strNodeDrel)
        ' Got anything?
        If (strPhrase <> "") Then
          ' Check for additional levels above [strPhrase]
          Select Case strLang
            Case "ch", "che"
              If (strPhrase = "CP-REL") Then
                ' We need to insert an IP-SUB between the RC and us
                If (Not eTreeInsertLevel(ndxThis, ndxWork)) Then Return False
                ' ============ DEBUG: Test =====================
                If (Not TestEndNodeIdOrder(ndxWork)) Then
                  Debug.Print("OneConllxNode: bad order")
                  Logging("modConLLX/OneConllxNode warning: bad constituent order at IP-SUB forestId=" & _
                          ndxFor.Attributes("forestId").Value)
                End If
                ' ==============================================
                ' Make [ndxWork] the new [ndxThis]
                ndxThis = ndxWork
                ' Set the phrase type
                ndxThis.Attributes("Label").Value = "IP-SUB"
              End If
          End Select
          ' Insert a level above me
          If (Not eTreeInsertLevel(ndxThis, ndxWork)) Then Return False
          ' Make [ndxWork] the new [ndxThis]
          ndxThis = ndxWork
          ' Set the phrase type
          ndxThis.Attributes("Label").Value = strPhrase
        End If
      Else
        ' If my own @con/id never figures as head, then I don't need an additional level
        ' Check if an additional level needs to be made
        Select Case strLang
          Case "nl", "du"   ' Dutch
            ' Action depends on what the label is
            Select Case UCase(strLabel)
              Case ".", """", ",", "!", "?", ":", ";"    ' Punctuation: don't insert an additional level
              Case Else
                ' I need to have my POS type
                strNodePos = GetFeature(ndxThis, "con", "pos")
                ' Safeguard the @con/drel feature
                strNodeDrel = GetFeature(ndxThis, "con", "drel")
                ' Determine the phrase label 
                strPhrase = GetPhraseRule(strNodePos, strNodeDrel)
                ' Did we get any result?
                If (strPhrase = "") AndAlso (Not DoLike(strNodePos, "C|Q|INT|ADV_NEG|D|RP|NEG|D$")) Then
                  ' Double check if this is what I want
                  ' Stop
                  Logging("There is no phrase rule for pos=[" & strNodePos & "] and Drel=[" & strNodeDrel & "]")
                Else
                  ' Insert a level above me
                  If (Not eTreeInsertLevel(ndxThis, ndxWork)) Then Return False
                  ' Make [ndxWork] the new [ndxThis]
                  ndxThis = ndxWork
                  ' Set the phrase type
                  ndxThis.Attributes("Label").Value = strPhrase
                End If
            End Select
          Case "ch", "che", "lez"  ' Chechen and Lezgi
            ' Action depends on what the label is
            Select Case UCase(strLabel)
              Case ".", """", ",", "!", "?", ":", ";", "«", "»"   ' Punctuation: don't insert an additional level
              Case Else
                ' I need to have my POS type
                strNodePos = GetFeature(ndxThis, "con", "pos")
                ' Safeguard the @con/drel feature
                strNodeDrel = GetFeature(ndxThis, "con", "drel")
                ' Determine the phrase label 
                strPhrase = GetPhraseRule(strNodePos, strNodeDrel)
                ' Did we get any result?
                If (strPhrase = "") AndAlso (Not DoLike(strNodePos, "C")) Then
                  ' Double check if this is what I want
                  ' Stop
                  Logging("Warning. No phrase rule found for combination:" & vbCrLf & _
                          "POS = " & strNodePos & " Drel=" & strNodeDrel)
                  ' Take a default phrase above me
                  strPhrase = "PhraseOver_" & strNodePos
                End If
                If (strPhrase = "CP-REL") Then
                  ' Save ourselves...
                  ndxSave = ndxThis
                  '' All preceding siblings that depend upon me...
                  'ndxDepList = ndxSave.SelectNodes("./preceding-sibling::eTree[" & _
                  '    "child::fs[@type='con']/child::f[@name='hd' and @value='" & _
                  '    GetFeature(ndxSave, "con", "id") & "']]")
                  ' We need to insert an IP-SUB between the RC and us
                  If (Not eTreeInsertLevel(ndxThis, ndxWork)) Then Return False
                  ' ============ DEBUG: Test =====================
                  If (Not TestEndNodeIdOrder(ndxWork)) Then
                    Debug.Print("OneConllxNode: bad order")
                    Logging("modConLLX/OneConllxNode warning: bad constituent order at IP-SUB forestId=" & _
                            ndxFor.Attributes("forestId").Value)
                  End If
                  ' ==============================================
                  ' Make [ndxWork] the new [ndxThis]
                  ndxThis = ndxWork
                  ' Set the phrase type
                  ndxThis.Attributes("Label").Value = "IP-SUB"
                  '' Prepend all members of [ndxDepList] under [IP-SUB]
                  'For intI = 0 To ndxDepList.Count - 1
                  '  ' Prepend this one
                  '  eTreeJoinUnder(ndxDepList(intI), ndxThis, True)
                  'Next intI
                End If
                ' Insert a level above me
                If (Not eTreeInsertLevel(ndxThis, ndxWork)) Then Return False
                ' ============ DEBUG: Test =====================
                If (Not TestEndNodeIdOrder(ndxWork)) Then
                  Debug.Print("OneConllxNode: bad order")
                  Logging("modConLLX/OneConllxNode warning: bad constituent order at IP-MAT(?) forestId=" & _
                          ndxFor.Attributes("forestId").Value)
                End If
                ' ==============================================
                ' Make [ndxWork] the new [ndxThis]
                ndxThis = ndxWork
                ' Set the phrase type
                ndxThis.Attributes("Label").Value = strPhrase
                ' Check if this is a top node
                If (intHeadId = 0) AndAlso (strPhrase Like "IP*") Then
                  ndxThis.Attributes("Label").Value = "IP-MAT"
                End If
            End Select
          Case "en"         ' English
            ' This is not yetimplemented for English
            Logging("OneConllxNode: not yet implemented for English") : Return False
          Case "es"         ' Spanish
            ' Action depends on what the label is
            Select Case UCase(strLabel)
              Case "F", """", ",", "."    ' Punctuation: don't insert an additional level
              Case "Z", "N", "S", "D", "A", "W", "C", "P", "R"
                ' Safeguard the @con/drel feature
                strNodeDrel = GetFeature(ndxThis, "con", "drel")
                ' The label "C" needs to be changed, depending on the DREL
                If (strLabel = "P") Then
                  ndxThis.Attributes("Label").Value = "PRO"
                ElseIf (strLabel = "C") AndAlso (DoLike(strNodeDrel, "COMP|DO|IO|MOD")) Then
                  strLabel = "S" : ndxThis.Attributes("Label").Value = strLabel
                End If
                ' Insert a level above me
                If (Not eTreeInsertLevel(ndxThis, ndxWork)) Then Return False
                ' Make [ndxWork] the new [ndxThis]
                ndxThis = ndxWork
                ' Get the index
                intL = Array.FindIndex(arRootNode, Function(strIn As String) strIn = UCase(strLabel))
                ' Determine the phrase type
                ndxThis.Attributes("Label").Value = arRootLink(intL)
                ' Possibly adapt the label of the child
                ndxThis.Attributes("Label").Value = DrelToLabel(ndxThis.Attributes("Label").Value, strNodeDrel)
              Case "V"    ' The kind of IP depends on the kind of verb we have
                ' I need to have my POS type
                strNodePos = GetFeature(ndxThis, "con", "pos")
                ' Insert a level above me
                If (Not eTreeInsertLevel(ndxThis, ndxWork)) Then Return False
                ' Make [ndxWork] the new [ndxThis]
                ndxThis = ndxWork
                ' Am I the root?
                If (intHeadId = "0") Then
                  ndxThis.Attributes("Label").Value = "IP-MAT"
                Else
                  ndxThis.Attributes("Label").Value = VerbTypeToNodeType(strNodePos)
                End If
              Case Else
                Logging("modConLLX/ConvertOneConLLx error: don't know what to do with root node label [" & strLabel & "]")
                Status("See LOG for error report")
                bInterrupt = True
                Return False
            End Select
          Case "none"
            ' No additional level needs to be made
        End Select
      End If

      ' Walk all these nodes that have me (intFeatId) as their head
      For intI = 0 To ndxList.Count - 1
        ' Access this node
        ndxChild = ndxList(intI)
        ' Get its @con/id feature value before we start meddling with it
        intChildId = GetFeature(ndxChild, "con", "id")
        ' Process this child
        If (Not OneConllxNode(ndxChild, ndxFor, strLang)) Then Return False
        ' Does [ndxThis] have children already?
        If (ndxThis.SelectSingleNode("./child::eTree") Is Nothing) Then
          ' Let this become the FIRST child under [ndxThis]
          ndxThis.AppendChild(ndxChild)
          ' ============ DEBUG: Test =====================
          If (Not TestEndNodeIdOrder(ndxThis)) Then Debug.Print("OneConllxNode: bad order")
          ' ==============================================
        Else
          ' Find the first child under [ndxThis] with an @con/id feature value that is 
          '   higher than the one of ndxChild
          'ndxFoll = ndxThis.SelectSingleNode("./child::eTree[" & _
          '          "child::fs[@type='con']/child::f[@name='id' and (@value > " & intChildId & ")]]")
          ndxFoll = ndxThis.SelectSingleNode("./child::eTree[" & _
                    "descendant::fs[@type='con']/child::f[@name='id' and (@value > " & intChildId & ")]]")
          ' validate
          If (ndxFoll Is Nothing) Then
            ' Let this become LAST child under [ndxThis]
            ndxThis.AppendChild(ndxChild)
            ' ============ DEBUG: Test =====================
            If (Not TestEndNodeIdOrder(ndxThis)) Then Debug.Print("OneConllxNode: bad order")
            ' ==============================================
          Else
            ' Insert [ndxChild] before [ndxFoll]
            ndxThis.InsertBefore(ndxChild, ndxFoll)
            ' ============ DEBUG: Test =====================
            If (Not TestEndNodeIdOrder(ndxThis)) Then Debug.Print("OneConllxNode: bad order")
            ' ==============================================
          End If
        End If
        '' Possibly adapt the label of the child
        'strNodeDrel = GetFeature(ndxChild, "con", "drel")
        'ndxChild.Attributes("Label").Value = DrelToLabel(ndxChild.Attributes("Label").Value, strNodeDrel)
      Next intI
      ' Return positively 
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/OneConllxNode error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       TestEndNodeIdOrder()
  ' Goal:       Text if the end-nodes are in the correct order
  ' History:
  ' 19-08-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function TestEndNodeIdOrder(ByRef ndxFor As XmlNode) As Boolean
    Dim ndxList As XmlNodeList  ' List of nodes
    Dim intConId As Integer     ' Value of current @con/id
    Dim intLastId As Integer    ' Value of previous id
    Dim intI As Integer         ' Counter

    Try
      ' Get a list of end-nodes under <forest>
      ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf)>0]")
      ' Get the first item
      If (ndxList.Count > 0) Then
        ' Get the @con/id value
        intLastId = GetFeature(ndxList(0), "con", "id")
        For intI = 1 To ndxList.Count - 1
          ' Get the ID of this one
          intConId = GetFeature(ndxList(intI), "con", "id")
          ' Check for mismatches
          If (intConId < intLastId) Then Return False
          ' Adapt last id
          intLastId = intConId
        Next intI
      End If
      ' We are okay
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/TestEndNodeIdOrder error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       VerbTypeToNodeType()
  ' Goal:       Given the detailed POS, give the node label that should govern it
  ' History:
  ' 19-08-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function VerbTypeToNodeType(ByVal strVerbType As String) As String
    Try
      ' Action depends on the kind of verb I am
      If (strVerbType Like "VM*") Then
        ' Main verb
        Return "IP-SUB"
      ElseIf (strVerbType Like "VA*") Then
        ' Auxiliary
        Return "IP-SUB"
      ElseIf (DoLike(strVerbType, "VMG*|VSG*")) Then
        ' Gerund
        Return "IP-PPL"
      ElseIf (DoLike(strVerbType, "VMM*|VSM*")) Then
        ' Imperative
        Return "IP-IMV"
      ElseIf (DoLike(strVerbType, "VMN*|VSN*")) Then
        ' Infinitive
        Return "IP-INF"
      ElseIf (DoLike(strVerbType, "VMP*|VSP*")) Then
        ' Participial
        Return "IP-PPL"
      Else
        Return "IP-SUB"
      End If
    Catch ex As Exception
      ' Give error
      HandleErr("modTxtToPsd/VerbTypeToNodeType error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return ""
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       DrelToLabel()
  ' Goal:       Given the current @Label and the @drel value, provide an adapted @Label
  ' History:
  ' 19-08-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function DrelToLabel(ByVal strLabel As String, ByVal strFunct As String) As String
    Try
      Select Case strFunct
        Case "SUBJ"     ' Subject
          strLabel &= "-SBJ"
        Case "DO"       ' Direct object
          strLabel &= "-OB1"
        Case "IO"       ' Indirect object
          strLabel &= "-OB2"
        Case "OBLC"     ' Oblique object
          strLabel &= "-" & strFunct
        Case "ADV"      ' Adverbial object
          strLabel &= "-" & strFunct
        Case "AUX"      ' Auxilary
          strLabel &= "-" & strFunct
        Case "BYAG"     ' By agent complement
          strLabel &= "-" & strFunct
        Case "ATR"      ' Attribute
          strLabel &= "-" & strFunct
        Case "PRD"      ' Predicative complement
          strLabel &= "-" & strFunct
        Case "OPRD"     ' Object predicative complement
          strLabel &= "-" & strFunct
        Case "PP-LOC"   ' Locative PP
          strLabel &= "-LOC"
        Case "PP-DIR"   ' Directional PP
          strLabel &= "-DIR"
        Case "SUBJ-GAP" ' Subject in a gapping construction
          strLabel &= "-SBJ-GAP"
        Case "COMP-GAP" ' Complement in a gapping construction
          strLabel &= "-" & strFunct
        Case "MOD-GAP"  ' Modifier in a gapping construction
          strLabel &= "-" & strFunct
        Case "VOC"      ' Vocative
          strLabel &= "-" & strFunct
        Case "MIMPERS"  ' Impersonal marker
          strLabel &= "-" & strFunct
        Case "MPAS"     ' Passive marker ("se")
          strLabel &= "-" & strFunct
        Case "MPRON"    ' Pronominal marker
          strLabel &= "-" & strFunct
        Case "COMP"     ' Complement (of N, ADJ, ADV, PREP)
          If (strLabel <> "C") Then
            strLabel &= "-" & strFunct
          End If
        Case "MOD"      ' Modifier
          strLabel &= "-" & strFunct
        Case "NEG"      ' Negation
          strLabel &= "-" & strFunct
        Case "SPEC"     ' Specifier
          strLabel &= "-" & strFunct
        Case "COORD"    ' Coordination
          If (strLabel = "C") Then
            strLabel = "CONJ"
          Else
            strLabel &= "-" & strFunct
          End If
        Case "CONJ"     ' Conjunction
          strLabel &= "-" & strFunct
        Case "punct"    ' Punctuation
      End Select
      ' Return the resulting label
      Return strLabel
    Catch ex As Exception
      ' Give error
      HandleErr("modTxtToPsd/DrelToLabel error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return ""
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       ConLLxAdaptVerbLabel()
  ' Goal:       Make an appropriate label V* or AX* or BE*, depending on part-of-speech strPos
  ' History:
  ' 19-08-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function ConLLxAdaptVerbLabel(ByVal strPos As String, ByVal strLang As String) As String
    Dim intI As Integer ' Counter

    Try
      ' Check all possibilities
      For intI = 0 To arVerbPos.Length - 1 Step 2
        If (DoLike(strPos, arVerbPos(intI))) Then
          ' Got you!
          Return arVerbPos(intI + 1)
        End If
      Next intI
      ' Return the resulting label
      Return "V"
    Catch ex As Exception
      ' Give error
      HandleErr("modTxtToPsd/ConLLxAdaptVerbLabel error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetOneCoNLL
  ' Goal:   Convert relevant information into one line for CoNLL-X
  ' History:
  ' 24-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetOneCoNLL(ByVal intDepId As Integer, ByVal strOrg As String, ByVal strCat As String, _
                      ByVal strFeats As String, ByVal intDepHd As String, ByVal strDepRel As String) As String
    Try
      ' Adaptations...
      If (strFeats = "") Then strFeats = "_"
      ' Combine
      Return intDepId & vbTab & strOrg & vbTab & strOrg & vbTab & PosToCpos(strCat) & vbTab & strCat & vbTab & _
                           strFeats & vbTab & intDepHd & vbTab & strDepRel & vbTab & "_" & vbTab & "_"
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/GetOneCoNLL error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetConLLpos
  ' Goal:   Get POS-tags into the CoNLL-X format
  ' History:
  ' 17-09-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetConLLpos(ByRef ndxThis As XmlNode) As String
    Dim strOrg As String = ""           ' Original text
    Dim strCat As String = ""           ' category
    Dim strDepRel As String = "_"       ' Dependency relation type (empty)
    Dim strFeats As String = ""         ' Features belonging to this node
    Dim intDepId As Integer             ' Dependency ID
    Dim intDepHd As String = "_"        ' Dependency head ID (empty)
    Dim ndxChild As XmlNode             ' Child node
    Dim ndxList As XmlNodeList          ' List of nodes
    Dim intI As Integer                 ' Counter
    Dim objThis As New StringColl

    Try
      ' Now walk all the <eTree> end nodes
      ndxList = ndxThis.SelectNodes("./descendant::eTree[count(child::eLeaf)>0]")
      For intI = 0 To ndxList.Count - 1
        ' Set the correct ID
        intDepId = intI + 1
        ' Access the child
        ndxChild = ndxList(intI)
        ' Get my own POS and my vernacular value
        strCat = ndxChild.Attributes("Label").Value
        strOrg = ndxChild.SelectSingleNode("./child::eLeaf").Attributes("Text").Value
        ' Get my features, but exclude features "Cat" and "Cyr", as well as "mbt" and "c"
        strFeats = GetFeatVect(ndxChild, "Cat", "Cyr", "mbt", "c")
        ' Process the information of this child in a line of CoNLL-X format:
        ' Id, Vern, Lemma, Cpos, Pos, Feats, Head, DepRel, X, X
        objThis.Add(GetOneCoNLL(intDepId, strOrg, strCat, strFeats, intDepHd, strDepRel))
      Next intI
      ' Return the result
      Return objThis.Text
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/GetConLLpos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetCoNLL
  ' Goal:   Get information for CoNLL-X
  ' History:
  ' 24-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetCoNLL(ByRef ndxThis As XmlNode) As String
    Dim strOrg As String = ""           ' Original text
    Dim strCat As String = ""           ' category
    Dim strDepRel As String = ""        ' Dependency relation type
    Dim strFeats As String = ""         ' Features belonging to this node
    Dim intDepId As Integer             ' Dependency ID
    Dim intDepHd As Integer             ' Dependency head ID
    Dim ndxChild As XmlNode             ' Child node
    Dim objThis As New StringColl

    Try
      ' Now walk all the <eTree> end nodes
      For Each ndxChild In ndxThis.SelectNodes("./descendant::eTree[count(child::eTree)>0]")
        ' Get my own POS and my vernacular value
        strCat = ndxChild.Attributes("Label").Value
        strOrg = ndxChild.SelectSingleNode("./child::eLeaf").Attributes("Text").Value
        ' Get my features
        strFeats = GetFeatVect(ndxChild)
        ' Get the dependency information: Id, Relation, Hd
        Do
          intDepId = GetFeature(ndxChild, "dep", "id")
          intDepHd = GetFeature(ndxChild, "dep", "hd")
          strDepRel = GetFeature(ndxChild, "dep", "rel")
          If (strDepRel = "Hd") Then
            ndxChild = ndxChild.ParentNode
          End If
        Loop While (strDepRel = "Hd")
        ' Process the information of this child in a line of CoNLL-X format:
        ' Id, Vern, Lemma, Cpos, Pos, Feats, Head, DepRel, X, X
        objThis.Add(GetOneCoNLL(intDepId, strOrg, strCat, strFeats, intDepHd, strDepRel))
      Next ndxChild
      ' Return the result
      Return objThis.Text
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/GetCoNLL error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitDepLang
  ' Goal:   Read the head rules
  ' History:
  ' 27-08-2013  ERK Created
  ' 05-09-2013  ERK Added Phraserule
  ' ------------------------------------------------------------------------------------
  Private Function InitDepLang(ByRef arRule() As HdRuleSpec, ByRef arDep() As DepRelSpec, ByRef arCpos() As CompactPos, ByRef arPos() As PosRule, _
                     ByRef arPhrase() As PhraseRule, ByRef arMorph() As MorphPos, ByVal strShort As String) As Boolean
    Dim strFile As String         ' Head rule file
    Dim strSrcFile As String      ' Default file location
    Dim strLine As String         ' One line
    Dim arLine() As String        ' One line
    Dim arText() As String        ' Text of file
    Dim intI As Integer           ' Counter
    ' Dim intPos As Integer         ' Position in string
    Dim intSize As Integer = 0    ' Size
    Dim intDepSize As Integer = 0 ' Size of dependency list
    Dim intCpSize As Integer = 0  ' Size of POS to CPOS translation table
    Dim intPosSize As Integer = 0 ' Size of POS rule table
    Dim intPhrSize As Integer = 0 ' Size of phrase rule table
    Dim intMrpSize As Integer = 0 ' Size of morph to pos rule table
    Dim strSection As String = "" ' Name of this section

    Try
      ' Validate
      If (strShort = "") Then Return False
      ' Define the file location we are expecting
      strFile = GetDocDir() & "\" & strShort
      ' Is the file there?
      If (Not IO.File.Exists(strFile)) Then
        ' Attempt to find the default file
        strSrcFile = IO.Path.GetDirectoryName(Application.ExecutablePath) & "\" & strShort
        If (Not IO.File.Exists(strSrcFile)) Then Return False
        ' Copy the default to the new location
        IO.File.Copy(strSrcFile, strFile, True)
      End If
      ' Read the rules as lines
      arText = IO.File.ReadAllLines(strFile)
      ' Clear what we have
      ReDim arRule(0)
      ' Walk the lines
      For intI = 0 To arText.Length - 1
        strLine = Trim(arText(intI))
        If (strLine <> "") AndAlso (Left(strLine, 2) <> "# ") AndAlso (strLine <> "#") Then
          ' Check if this is a section indicator
          If (InStr(strLine, "#section") = 1) Then
            ' Start of a section: read the name of the section
            strSection = Trim(Mid(strLine, 9))
          End If
          ' Action depends on section
          Select Case strSection
            Case "HeadRules", "HdRules"
              ' Get the line
              arLine = Split(strLine, vbTab)
              If (arLine.Length = 3) Then
                ' Make room
                ReDim Preserve arRule(0 To intSize)
                With arRule(intSize)
                  .phrase = arLine(0) : .dir = arLine(1) : .hdlist = arLine(2)
                End With
                ' Move the array size on
                intSize += 1
              End If
            Case "DependencyRelations", "DepRel", "DepRels"
              ' Get the line
              arLine = Split(strLine, vbTab)
              If (arLine.Length = 3) Then
                ' Make room
                ReDim Preserve arDep(0 To intDepSize)
                With arDep(intDepSize)
                  .rel = arLine(0) : .par = arLine(1) : .lablist = arLine(2)
                End With
                ' Move the array size on
                intDepSize += 1
              End If
            Case "CompactPos", "Cpos", "CPOS", "CompactPOS"
              ' Get the line
              arLine = Split(strLine, vbTab)
              If (arLine.Length = 2) Then
                ' Make room
                ReDim Preserve arCpos(0 To intCpSize)
                With arCpos(intCpSize)
                  .Cpos = arLine(0) : .poslist = arLine(1)
                End With
                ' Move the array size on
                intCpSize += 1
              End If
            Case "PosRule", "PosRules", "POS", "Pos"
              ' Get the line
              arLine = Split(strLine, vbTab)
              If (arLine.Length >= 3) Then
                ' Make room
                ReDim Preserve arPos(0 To intPosSize)
                With arPos(intPosSize)
                  .Pos = arLine(0) : .Cpos = UCase(arLine(1)) : .ftlist = arLine(2)
                  If (arLine.Length = 4) Then
                    .vern = LCase(arLine(3))
                  Else
                    .vern = "*"
                  End If
                End With
                ' Move the array size on
                intPosSize += 1
              End If
            Case "PhraseRule", "Phrase", "phrase", "PhraseRules", "Phrases", "phrases"
              ' Get the line
              arLine = Split(strLine, vbTab)
              If (arLine.Length = 3) Then
                ' Make room
                ReDim Preserve arPhrase(0 To intPhrSize)
                With arPhrase(intPhrSize)
                  .phrase = arLine(0) : .poslist = arLine(1) : .dep = arLine(2)
                End With
                ' Move the array size on
                intPhrSize += 1
              End If
            Case "MorphPos", "morphpos"
              ' Conversion from morphology to POS
              ' Get the line
              arLine = Split(strLine, vbTab)
              If (arLine.Length = 3) Then
                ' Make room
                ReDim Preserve arMorph(0 To intMrpSize)
                With arMorph(intMrpSize)
                  .Pos = arLine(0) : .morphlist = arLine(1) : .poslist = arLine(2)
                End With
                ' Move the array size on
                intMrpSize += 1
              End If
          End Select
        End If
      Next intI
      ' Check for critical sections
      If (intSize = 0) Then
        Logging("The language definition file needs to have a [HeadRules] section")
        Return False
      End If
      If (intCpSize = 0) Then
        Logging("The language definition file needs to have a [CompactPos] section")
        Return False
      End If
      If (intDepSize = 0) Then
        Logging("The language definition file needs to have a [DependencyRelations] section")
        Return False
      End If
      If (intPhrSize = 0) Then
        Logging("The language definition file could do with a [PhraseRules] section for ConLLX to psdx conversion...")
        Return True
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/InitDepLang error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetHeadNodeDep
  ' Goal:   Return the constituent that is the head of the phrase in [ndxThis]
  '         - This phrase must be <eTree>
  '           (any <eTree> child of a <forest> automatically is a head)
  '         - Definition of head depends on the language, specified in [strOp]
  ' History:
  ' 27-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetHeadNodeDep(ByRef ndxThis As XmlNode, ByVal strLang As String) As XmlNode
    Dim ndxList As XmlNodeList = Nothing  ' The children 
    Dim ndxWork As XmlNode = Nothing      ' Working node
    Dim strLabel As String                ' The node's label value
    Dim strPhrase As String               ' Phrase label
    '  Dim strLang As String                 ' The language to be used
    Dim strConst As String = ""           ' List of constituent children
    Dim strDirection As String            ' Direction
    Dim arHd() As String                  ' Possible head labels
    Dim intRule As Integer                ' Index of the rule
    Dim intI As Integer                   ' Counter
    Dim intJ As Integer                   ' Counter

    Try
      ' Validate 
      If (ndxThis Is Nothing) OrElse (Not DoLike(ndxThis.Name, "eTree")) Then Return Nothing
      ' Check if we need to do something
      If (Not bLangInit) Then
        '' Determine the language
        'Select Case strOp
        '  Case "DepChe"
        '    strLang = "che"
        '  Case "DepEng"
        '    strLang = "en"
        '  Case "DepNl"    ' Dutch dependency
        '    strLang = "nl"
        '  Case Else
        '    ' Failure
        '    Return Nothing
        'End Select
        If (Not DoLangInit(strLang)) Then bInterrupt = True : Return Nothing
      End If
      ' Get all the <eTree> children
      ' Old: ndxList = ndxThis.SelectNodes("./child::eTree[count(child::eLeaf[@Type='Punct'])=0 and @Label!='CODE']")
      ' EK: changed @Type= 'Punct' into @Type!='Vern', to exclude all Star, Zero etc. categories as potential heads
      'ndxList = ndxThis.SelectNodes("./child::eTree[" & _
      ndxList = ndxThis.SelectNodes("./child::eTree[@Label != 'CODE' and tb:matches(tb:deptype(self::eTree),'*node')]", conTb)
      '  "count(child::eLeaf[@Type!='Vern'])=0 and @Label!='CODE']")
      ' If there is no <eTree> child, then this node does not function as the head of another node
      If (ndxList.Count = 0) Then Return Nothing
      ' If there is only one child, then that child is the head of this node
      If (ndxList.Count = 1) Then Return ndxList(0)
      ' Now we only need to deal with <eTree> nodes...
      ' Get the node's label
      strLabel = ndxThis.Attributes("Label").Value
      strPhrase = PhraseLabel(strLabel)
      ' Look for the correct head rule
      intRule = -1
      For intI = 0 To arHeadRule.Length - 1
        If (arHeadRule(intI).phrase = strPhrase) Then intRule = intI : Exit For
      Next intI
      ' Validate
      If (intRule < 0) Then Logging("modConLXX/GetHeadNodeDep error: " & _
                      "could not find the head rule for phrase [" & strPhrase & "]") : Return Nothing
      ' Get the list of potential head labels and the direction to look
      arHd = Split(arHeadRule(intRule).hdlist, ";")
      strDirection = arHeadRule(intRule).dir
      ' Walk all the potential head labels in preferential order
      For intI = 0 To arHd.Length - 1
        ' Look for potential heads in the indicated order of the nodes
        If (strDirection = "l") Then
          ' Walk left to right
          For intJ = 0 To ndxList.Count - 1
            ' Check if this matches
            If (DoLike(PhraseLabel(ndxList(intJ).Attributes("Label").Value), arHd(intI))) Then Return ndxList(intJ)
            If (DoLike(ndxList(intJ).Attributes("Label").Value, arHd(intI))) Then Return ndxList(intJ)
          Next intJ
        ElseIf (strDirection = "r") Then
          ' Walk right to left
          For intJ = ndxList.Count - 1 To 0 Step -1
            ' Check if this matches
            If (DoLike(PhraseLabel(ndxList(intJ).Attributes("Label").Value), arHd(intI))) Then Return ndxList(intJ)
            If (DoLike(ndxList(intJ).Attributes("Label").Value, arHd(intI))) Then Return ndxList(intJ)
          Next intJ
        Else
          ' Bad direction error
          Logging("modConLXX/GetheadNodeDep error: bad direction indicator [" & strDirection & "]")
          Return Nothing
        End If
      Next intI

      'If (strDirection = "l") Then
      '  ' Try visit left to right
      '  For intI = 0 To ndxList.Count - 1
      '    ' Match this label with the potential heads
      '    If (HeadDepMatches(arHd, PhraseLabel(ndxList(intI).Attributes("Label").Value))) Then Return ndxList(intI)
      '  Next intI
      'ElseIf (strDirection = "r") Then
      '  ' Try visit right to left
      '  For intI = ndxList.Count - 1 To 0 Step -1
      '    ' Match this label with the potential heads
      '    If (HeadDepMatches(arHd, PhraseLabel(ndxList(intI).Attributes("Label").Value))) Then Return ndxList(intI)
      '  Next intI
      'Else
      '  ' Bad direction error
      '  Logging("modConLXX/GetheadNodeDep error: bad direction indicator [" & strDirection & "]")
      '  Return Nothing
      'End If
      ' If we haven't gotten a match yet, see if one of the children has the same label I have
      ndxList = ndxThis.SelectNodes("./child::eTree[@Label = '" & ndxThis.Attributes("Label").Value & "']")
      If (ndxList.Count > 0) Then
        ' Take the first one that matches
        Return ndxList(0)
      End If
      ' ================ DEBUG ===================
      strConst = "" : ndxList = ndxThis.SelectNodes("./child::eTree")
      For intI = 0 To ndxList.Count - 1
        If (strConst <> "") Then strConst &= " "
        strConst &= ndxList(intI).Attributes("Label").Value
      Next intI
      Debug.Print("Line " & ndxThis.SelectSingleNode("./ancestor::forest").Attributes("forestId").Value)
      Debug.Print("Unable to determine the head for " & strLabel & " [" & strConst & "]")
      Debug.Print(NodeText(ndxThis))
      Stop
      ' ===========================================
      ' Return failure
      Return Nothing
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/GetHeadNodeDep error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DoLangInit
  ' Goal:   Initialise language processing for [strLang]
  ' History:
  ' 27-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function DoLangInit(ByVal strLang As String) As Boolean
    Dim bRead As Boolean = False          ' Have we read the rules?
    Dim strFile As String = ""

    Try
      ' Validate
      If (bLangInit) Then Return True
      If (strLang = "nld") Then strLang = "nl"
      ' Read dependency rules for this language
      For intI = 0 To arHeadRule_lng.Length - 1 Step 2
        If (arHeadRule_lng(intI) = strLang) Then
          ' Read the rules
          If (Not InitDepLang(arHeadRule, arDepRel, arCposRule, arPosRule, arPhraseRule, arMorphPos, arHeadRule_lng(intI + 1))) Then Return Nothing
          bRead = True
          Exit For
        End If
      Next intI
      ' Validate
      If (Not bRead) Then
        Status("Could not read the dependency rules for language " & strLang)
        Return Nothing
      End If
      ' Try to initialize a language for aggluspell POS tagging
      objPosTag = New PosTag(strLang)
      If (objPosTag IsNot Nothing) Then
        tdlLiftPosDic = objPosTag.dict
      End If
      ' Set the flag
      bLangInit = True
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/DoLangInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   getPosFromLift
  ' Goal:   Assuming a <liftpos> dictionary has been read, try to get the POS value
  '           for the word back
  ' History:
  ' 17-07-2015  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function getPosFromLift(ByVal strWord As String, ByVal bCyr As Boolean, _
          ByRef strDef As String, ByRef strLat As String, ByRef strLemma As String) As String
    Dim strLngField As String   ' Field name of the target-language
    Dim strLngOther As String   ' Field for the non-target-language
    Dim strAlloField As String  ' Allophone field for affix dic
    Dim strAlt As String = ""   ' Alternative parse
    Dim intParses As Integer    ' number of parses
    Dim dtrFound() As DataRow   ' Result of SELECT
    Dim dtrPar As DataRow       ' One datarow
    Dim intI As Integer         ' Counter
    Dim prsThis As Parse = Nothing

    Try
      ' Validate
      If (tdlLiftPosDic Is Nothing) Then Return ""
      ' Determine which language field to use
      strLngField = IIf(bCyr, "lexc", "lex")
      strLngOther = IIf(bCyr, "lex", "lexc")
      strAlloField = IIf(bCyr, "ac", "a")
      ' Can we use parsing or not?
      If (objPosTag IsNot Nothing AndAlso IO.File.Exists(objPosTag.affixFile)) Then
        ' Tell the parser which main field is for the language
        objPosTag.lexField = strLngField
        objPosTag.alloField = strAlloField
        objPosTag.affixLanguage = IIf(bCyr, "cyr", "lat")
        prsThis = Nothing
        ' We can try and use full parsing
        If (objPosTag.DoParse(strWord, 0, strAlt, "", "dict|rewrite")) Then
          intParses = objPosTag.parses.Count
          If (intParses = 1) Then
            ' Return the information from this parse
            prsThis = objPosTag.parses(0)
          ElseIf (intParses > 1) Then
            ' Check if they differ in terms of POS
            Dim strPos = objPosTag.parses(0).Cat
            Dim bFound As Boolean = False
            For intI = 1 To objPosTag.parses.Count - 1
              If (objPosTag.parses(intI).Cat <> strPos) Then bFound = True : Exit For
            Next intI
            If (Not bFound) Then
              ' We can safely return the first parse, since the POS does not differ
              prsThis = objPosTag.parses(0)
            End If
          End If
        End If
        If (prsThis IsNot Nothing) Then
          ' Return the information from the selected parse
          With prsThis
            strLemma = .Lemma
            strDef = .Def
            strLat = ""
            Return .POS
          End With
        End If
      Else
        ' Look at lower case
        strWord = strWord.ToLower.Replace("'", "''")
        ' Look in the main entries
        dtrFound = tdlLiftPosDic.Tables("entry").Select(strLngField & " = '" & strWord & "'")
        If (dtrFound.Length = 0) Then
          ' Look at <pdg>
          dtrFound = tdlLiftPosDic.Tables("pdg").Select(strLngField & " = '" & strWord & "'")
          If (dtrFound.Length = 0) Then Return ""
          ' Get the parent
          dtrPar = dtrFound(0).GetParentRow("entry_pdg")
          ' Get def and lat
          strDef = dtrPar.Item("def")
          strLemma = dtrPar.Item(strLngField)
          strLat = dtrPar.Item(strLngOther)
        Else
          ' Get def and lat
          strDef = dtrFound(0).Item("def")
          strLemma = dtrFound(0).Item(strLngField)
          strLat = dtrFound(0).Item(strLngOther)
        End If
        ' Return the POS of the FIRST hit
        Return dtrFound(0).Item("pos")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/getPosFromLift error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   HeadDepMatches
  ' Goal:   See if label [strLabel] matches one in arHd
  ' History:
  ' 27-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function HeadDepMatches(ByRef arHd() As String, ByVal strLabel As String) As Boolean
    Dim intI As Integer   ' Counter

    Try
      ' Validate
      If (arHd.Length = 0) Then Return False
      If (strLabel = "") Then Return False
      ' Start matching
      For intI = 0 To arHd.Length - 1
        ' Is this a match?
        If (DoLike(strLabel, arHd(intI))) Then Return True
      Next intI
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/HeadDepMatches error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   PhraseLabel
  ' Goal:   Strip everything from [strIn], leaving only the phrase label part
  ' History:
  ' 27-08-2013  ERK Created
  ' 16-10-2013  ERK Make sure to return only the part following the last + sign
  ' ------------------------------------------------------------------------------------
  Private Function PhraseLabel(ByVal strIn As String) As String
    Dim intPos As Integer   ' Position
    Dim strTmp As String    ' Temporary string

    Try
      ' Validate
      If (strIn = "") Then Return ""
      ' Find the first non-character position except for PLUS
      intPos = 1
      While (intPos < strIn.Length) AndAlso (Mid(strIn, intPos, 1) Like "[a-zA-Z+]")
        ' Continue
        intPos += 1
      End While
      ' Return the part we have
      If (intPos < strIn.Length) Then
        strTmp = Left(strIn, intPos - 1)
      Else
        strTmp = strIn
      End If
      ' Find the rightmost + sign
      intPos = InStrRev(strTmp, "+")
      If (intPos = 0) Then
        Return strTmp
      Else
        Return Mid(strTmp, intPos + 1)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/PhraseLabel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return strIn
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   mnuToolsMorphHeadRules_Click
  ' Goal:   Get a list of the possible children each XP can have
  ' History:
  ' 27-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function TravelXPchildren(ByRef strReport As String) As Boolean
    Dim arIn() As String            ' All files
    Dim ndxList As XmlNodeList      ' List of relevant nodes in a forest
    Dim ndxChild As XmlNodeList     ' List of relevant child nodes in an <eTree>
    ' Dim ndxThis As XmlNode          ' One node
    Dim ndxFor As XmlNode = Nothing ' Forest node
    Dim arSort() As String          ' Sorted result
    Dim colResult As New StringColl ' Where we store the result lines
    Dim colBack As New StringColl   ' HTML report
    Dim strPhrase As String         ' Phrase label
    Dim strEntry As String          ' Entry
    Dim strChild As String          ' Child label
    Dim strShort As String          ' Short file name
    Dim intI As Integer             ' Counter
    Dim intJ As Integer             ' Counter
    Dim intK As Integer             ' Counter
    Dim intF As Integer             ' Forest count
    Dim intIdx As Integer           ' Index
    Dim intPtc As Integer           ' Counter
    Dim intNum As Integer           ' Number of children
    Dim strDir_en As String = "D:\Data Files\Corpora\English\xml\Adapted"
    Dim strDir_nl As String = "D:\Data Files\Corpora\Dutch\xml"
    Dim strDir_ch As String = "d:\data files\corpora\Chechen\xml\MyTask"
    Dim strDir As String = "D:\Data Files\Corpora\English\xml\Adapted"
    Dim pdxThis As XmlDocument = Nothing

    Try
      ' Ask which
      Select Case MsgBox("Would you like to process Dutch (Y) or English (N)?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Cancel
          Return False
        Case MsgBoxResult.Ok, MsgBoxResult.Yes
          strDir = strDir_nl
        Case MsgBoxResult.No
          strDir = strDir_en
      End Select
      ' Validate
      If (Not IO.Directory.Exists(strDir)) Then Return False
      arIn = IO.Directory.GetFiles(strDir, "*.psdx", IO.SearchOption.AllDirectories)
      ' Walk all files
      For intI = 0 To arIn.Length - 1
        ' Get filename
        strShort = IO.Path.GetFileNameWithoutExtension(arIn(intI))
        ' Show where we are...
        Status("Reading file " & intI + 1 & "/" & arIn.Length & ": " & strShort & "...")
        ' Read this file as psdx
        If (ReadXmlDoc(arIn(intI), pdxThis)) Then
          ' Get number of forests
          intNum = pdxThis.SelectNodes("//forest").Count : intF = 1
          ' Get the first forest
          If (Not GetFirstForest(pdxThis, ndxFor)) Then Return False
          ' Walk all forests
          While (Not ndxFor Is Nothing)
            ' Where are we?
            intPtc = intF * 100 \ intNum : intF += 1
            Status(strShort & " " & intPtc & "%", intPtc)
            ' Get all non-terminal <eTree> nodes in this forest
            ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf)=0]")
            ' Walk all the element nodes
            For intJ = 0 To ndxList.Count - 1
              ' Get my phrase label
              strPhrase = PhraseLabel(ndxList(intJ).Attributes("Label").Value)
              ' Find parent entry
              If (Not colResult.Exists(strPhrase)) Then colResult.AddUnique(strPhrase)
              intIdx = colResult.Find(strPhrase)
              ' Get parent entry
              strEntry = colResult.Exmp(intIdx)
              ' Get all relevant child-nodes in this phrase
              ndxChild = ndxList(intJ).SelectNodes("./child::eTree[" & _
                             "not(child::eLeaf[tb:matches(@Type,'Punct|Star')])" & _
                             " and not(tb:matches(@Label, 'CODE|META'))" & _
                             "]", conTb)
              If (ndxChild.Count > 0) Then
                ' Add all relevant children to my phrase
                For intK = 0 To ndxChild.Count - 1
                  ' Get child label
                  strChild = PhraseLabel(ndxChild(intK).Attributes("Label").Value)
                  ' Add child to parent
                  AddSemiStack(strEntry, strChild, True)
                Next intK
                ' Put back into [colResult]
                arSort = Split(strEntry, ";")
                Array.Sort(arSort)
                strEntry = Join(arSort, ";")
                colResult.Exmp(intIdx) = strEntry
              End If
            Next intJ
            ' Go to next forest
            ndxFor = ndxFor.NextSibling
          End While
        End If
      Next intI
      ' Start HTML result
      colBack.Add("<html><body><table><tr><td>Phrase></td><td>Children</td></tr>")
      ' Walk results
      For intI = 0 To colResult.Count - 1
        ' Show this result
        colBack.Add("<tr><td>" & colResult.Item(intI) & "</td>" & _
                      "<td>" & colResult.Exmp(intI) & "</td></tr>")
      Next intI
      ' Finish HTML
      colBack.Add("</table></body></html>")
      strReport = colBack.Text
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/TravelXPchildren error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   OneForestToConLLX
  ' Goal:   Convert one forest that has been annotated with dependency relations to a string in CONLL-X format
  ' History:
  ' 28-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function OneForestToConLLX(ByRef ndxFor As XmlNode, ByRef strSect As String) As Boolean
    Dim ndxRoot As XmlNodeList    ' List of root nodes
    Dim ndxEnd As XmlNodeList     ' List of endnodes
    Dim ndxThis As XmlNode        ' One node
    ' Dim ndxTop As XmlNode         ' Top node of one structural part
    Dim strForm As String         ' One form
    Dim strDepRel As String       ' Dependency relation
    Dim colSect As New StringColl ' List of sections
    Dim arSect() As CONLLX        ' Array for this section
    Dim intCount As Integer = 0   ' Number of root nodes
    Dim intAddId As Integer       ' ID of the first node in this section of the forest
    Dim intTopId As Integer       ' ID of root
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter

    Try
      ' Validate
      If (ndxFor Is Nothing) Then Logging("Cannot find <forest>") : Return False
      If (ndxFor.Name <> "forest") Then Logging("Cannot find <forest>") : Return False
      ' Initialise
      colDoHdDupli.Clear() : colDoHdStack.Clear()
      ' Get a list of root nodes
      ndxRoot = ndxFor.SelectNodes("./child::eTree[child::fs[@type='dep']/child::f[@name='rel' and @value='root']]")
      If (ndxRoot.Count > 1) Then
        Logging("Multiple roots at line " & ndxFor.Attributes("TextId").Value & ":" & _
                    ndxFor.Attributes("forestId").Value & " (" & ndxRoot.Count & ")")
      End If
      ' Walk all roots
      For intI = 0 To ndxRoot.Count - 1
        ' Reset counter
        intCount = 0 : intTopId = -1
        ' Calculate the ID of the first node in this section
        intAddId = GetFeature(ndxRoot(intI).SelectSingleNode( _
                      "./descendant::eTree[child::fs[@type='dep']/child::f[@name='id']]"), "dep", "id") - 1
        ' =============== DEBUG ===============
        ' If (ndxFor.Attributes("forestId").Value = 2) Then Stop
        ' =====================================
        ' Find the other head values
        If (Not DoHeadValues(ndxRoot(intI), 0)) Then Logging("OneForestToConLLX/DoHeadValues error") : Return False
        ' Adjust the heads for all the "dnode" nodes
        If (Not DoBindDnodes(ndxRoot(intI))) Then Logging("OneForestToConLLX/DoBindNodes error") : Return False
        ' Check the amount of room we need
        ndxEnd = ndxRoot(intI).SelectNodes("./descendant::eTree[child::fs[@type='dep']/child::f[@name='id']]")
        ReDim arSect(0 To ndxEnd.Count - 1)
        ' Walk the endnodes
        For intJ = 0 To ndxEnd.Count - 1
          ' Access this node
          ndxThis = ndxEnd(intJ)
          ' Validate
          If (GetFeature(ndxThis, "dep", "hd").ToString = "") Then
            ' This node does not have its head determined, so we run into problems
            Debug.Print("line " & ndxFor.Attributes("forestId").Value & ": there is no head in constituent [" & ndxThis.Attributes("Label").Value & "]")
            Stop
          End If
          ' Get the form
          strForm = ndxThis.SelectSingleNode("./child::eLeaf").Attributes("Text").Value
          ' There may be some changes to the form of the input
          strForm = strForm.Replace("’", "'")
          strForm = strForm.Replace("«", """")
          strForm = strForm.Replace("»", """")
          ' ============ DEBUG ===========
          ' If (ndxFor.Attributes("forestId").Value = 65) Then Stop
          ' ==============================
          ' Get the dependency relation
          strDepRel = GetUpperDepRel(ndxThis)
          ' Check for root
          If (strDepRel = "root") Then
            If (intTopId > 0) Then
              ' This is a very serious problem: more than one <root> relation in a forest
              Logging("The structure at [" & ndxFor.Attributes("TextId").Value & ":" & _
                    ndxFor.Attributes("forestId").Value & "] is bad." & vbCrLf & _
                    "  I found more than one node serving as ROOT." & vbCrLf & _
                    "  You should check and adapt the head rules and make sure each clause has a proper head." & vbCrLf & _
                    "  I will continue, and repair the structure for the moment." & vbCrLf & _
                    "  The structure after repair is:" & vbCrLf & strSect)
            Else
              intTopId = GetFeature(ndxThis, "dep", "id")
            End If
          End If
          ' Fill in some of the values
          With arSect(intJ)
            .Id = GetFeature(ndxThis, "dep", "id") - intAddId : .Form = strForm
            .Lemma = .Form : .Feats = "_" : .PosTag = ndxThis.Attributes("Label").Value : .CposTag = PosToCpos(.PosTag)
            .Head = GetFeature(ndxThis, "dep", "hd") : .DepRel = strDepRel : .Phead = "_" : .PdepRel = "_"
            ' Check for roots
            If (.Head = 0) Then
              intCount += 1
            Else
              .Head -= intAddId
            End If
          End With
        Next intJ
        ' Process this section
        For intJ = 0 To arSect.Length - 1
          ' Need adaptation?
          If (intCount > 1) AndAlso (arSect(intJ).Head = 0) Then
            arSect(intJ).Head = intTopId
          End If
          ' Store this one -- this takes care of CONLL-X format with columns in the correct way.
          With arSect(intJ)
            colSect.Add(.Id & vbTab & .Form & vbTab & .Lemma & vbTab & .CposTag & vbTab & .PosTag & vbTab & _
                        .Feats & vbTab & .Head & vbTab & .DepRel & vbTab & .Phead & vbTab & .PdepRel)
          End With
        Next intJ
        ' Add one empty line to separate sections
        colSect.Add("")
      Next intI
      ' Return the combined result
      strSect = colSect.Text
      ' Check
      If (intCount > 1) Then
        Logging("The structure at [" & ndxFor.Attributes("TextId").Value & ":" & _
                ndxFor.Attributes("forestId").Value & "] is bad." & vbCrLf & _
                "  You should check and adapt the head rules and make sure each clause has a proper head." & vbCrLf & _
                "  I will continue, and repair the structure for the moment." & vbCrLf & _
                "  The structure after repair is:" & vbCrLf & strSect)
        'bInterrupt = True
        'Return False
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/OneForestToConLLX error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   PosToCpos
  ' Goal:   Convert POS tag into CPOS
  ' History:
  ' 02-09-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function PosToCpos(ByVal strPos As String) As String
    Dim intI As Integer ' Counter

    Try
      ' Validate
      If (strPos = "") Then Return ""
      ' Walk the possibilities
      For intI = 0 To arCposRule.Length - 1
        If (DoLike(strPos, arCposRule(intI).poslist)) Then
          Return arCposRule(intI).Cpos
        End If
      Next intI
      ' Default: just return the POS 
      Return PhraseLabel(strPos)
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/PosToCpos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return strPos
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphToPos
  ' Goal:   Convert Lezgi morpheme into POS tag
  '         Give preference to longest recognition
  ' History:
  ' 19-03-2014  ERK Created
  ' 02-04-2014  ERK Added [strPosOrg] to the table and interpretation
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MorphToPos(ByVal strMorph As String, ByVal strPosOrg As String) As String
    Dim intI As Integer             ' Counter
    Dim intLong As Integer = 0      ' Longest recognition
    Dim intLen As Integer           ' Current length
    Dim bSpecf As Boolean           ' Give preference to more specific rules, involving Morph + POSorg
    Dim strMatch As String = "X"    ' The match
    Dim bMatch As Boolean           ' Match or not
    ' Dim colMatch As New StringColl  ' Collection of matches

    Try
      ' Validate
      If (strMorph = "") Then Return "X"
      If (arMorphPos Is Nothing) Then Return "X"
      If (arMorphPos.Length = 0) Then Return "X"
      '' DEBUG
      'If (strMorph = "COP-PST") Then Stop
      '' =====================================
      ' Walk the possibilities
      For intI = 0 To arMorphPos.Length - 1
        With arMorphPos(intI)
          If (.morphlist = "") Then
            ' Only use poslist
            bMatch = DoLike(strPosOrg, .poslist) : bSpecf = False
          ElseIf (.poslist = "") Then
            ' Only use morphlist
            bMatch = DoLike(strMorph, .morphlist) : bSpecf = False
          Else
            ' Use both of them
            bMatch = DoLike(strPosOrg, .poslist) AndAlso DoLike(strMorph, .morphlist) : bSpecf = True
          End If
          ' Do we have a match?
          If (bMatch) Then
            intLen = .morphlist.Length + .poslist.Length
            ' colMatch.Add(arMorphPos(intI).Pos, intLen)
            ' Keep track of longest
            If (intLen > intLong) OrElse (bSpecf) Then
              intLong = intLen : strMatch = .Pos
              ' Give strong preference to specific rules
              If (bSpecf) Then
                Return strMatch
              End If
            End If
          End If
        End With
      Next intI
      ' Last check
      If (strMatch = "X") AndAlso (strPosOrg <> "") Then
        strMatch = strPosOrg
      End If
      ' Return the best match
      Return strMatch
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/MorphToPos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return "X"
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetUpperDepRel
  ' Goal:   Find the nearest up-going dependency relation that is not equal to "-" or "hd"
  ' History:
  ' 29-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function GetUpperDepRel(ByRef ndxThis As XmlNode) As String
    Dim ndxWork As XmlNode  ' Working node
    Dim strRel As String    ' Relation

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return "-"
      ' Walk up
      ndxWork = ndxThis
      While (ndxWork IsNot Nothing) AndAlso (ndxWork.Name = "eTree")
        ' Check this relation
        strRel = GetFeature(ndxWork, "dep", "rel")
        ' Test it
        If (strRel <> "-") AndAlso (strRel <> "hd") Then
          ' Return this one
          Return strRel
        End If
        ' Go up
        ndxWork = ndxWork.ParentNode
      End While
      ' Return failure
      Return "-"
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/GetUpperDepRel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return "-"
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetConSectId
  ' Goal:   Look at the array [arSect] and get the indxe of the entry that has @Id value [intIdx]
  ' History:
  ' 29-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function GetConSectId(ByRef arSect() As CONLLX, ByVal intIdx As Integer) As Integer
    Dim intI As Integer   ' Counter

    Try
      For intI = 0 To arSect.Length - 1
        If (arSect(intI).Id = intIdx) Then Return intI
      Next intI
      ' Return failure
      Return -1
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/GetConSectId error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   IsDislocatedPhrase
  ' Goal:   Check if this is a dislocated phrase
  ' History:
  ' 29-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function IsDislocatedPhrase(ByRef ndxThis As XmlNode, ByRef intExtra As Integer) As Boolean
    Dim strLabel As String  ' Label
    Dim strNum As String    ' Number
    Dim intPos As Integer   ' Position in string

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      If (ndxThis.Name <> "eTree") Then Return False
      ' Get my label
      strLabel = ndxThis.Attributes("Label").Value
      ' Get position of rightmost hyphen
      intPos = InStrRev(strLabel, "-")
      ' Results?
      If (intPos = 0) Then Return False
      ' Check rightmost part for being a number
      strNum = Mid(strLabel, intPos + 1)
      If (IsNumeric(strNum)) Then
        ' Return the number
        intExtra = CInt(strNum)
        Return True
      Else
        Return False
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/IsDislocatedPhrase error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetTraceNumber
  ' Goal:   Get the numerical identifier of this trace
  ' History:
  ' 29-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function GetTraceNumber(ByRef ndxThis As XmlNode) As Integer
    Dim strLabel As String  ' Label
    Dim strNum As String    ' Number
    Dim intPos As Integer   ' Position in string

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return 0
      If (ndxThis.Name <> "eLeaf") Then Return 0
      ' Get my label
      strLabel = ndxThis.Attributes("Text").Value
      ' Get position of rightmost hyphen
      intPos = InStrRev(strLabel, "-")
      ' Results?
      If (intPos = 0) Then Return 0
      ' Check rightmost part for being a number
      strNum = Mid(strLabel, intPos + 1)
      If (IsNumeric(strNum)) Then
        ' Return the number
        Return CInt(strNum)
      Else
        Return 0
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/GetTraceNumber error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return 0
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoBindDnodes
  ' Goal:   Find all the nodes of type "dnode" in [ndxThis], and bind them to their corresponding "droot" ones
  ' History:
  ' 30-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function DoBindDnodes(ByRef ndxThis As XmlNode) As Boolean
    Dim ndxFor As XmlNode       ' Forest
    Dim ndxLeaf As XmlNode      ' Leaf
    Dim ndxStart As XmlNode     ' Working node
    Dim ndxList As XmlNodeList  ' List of nodes
    Dim ndxOld As XmlNodeList   ' List of nodes with old hd id
    Dim intHdOld As Integer
    Dim intHdNew As Integer
    Dim intExt As Integer       ' Extension number
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Get forest
      ndxFor = ndxThis.SelectSingleNode("./ancestor-or-self::forest")
      ' ========= DEBUG =========
      ' If (ndxFor.Attributes("forestId").Value = 12) Then Stop
      ' =========================
      ' Find the "dnode" nodes
      ndxList = ndxThis.SelectNodes("./descendant::eTree[tb:deptype(self::eTree) = 'dnode']", conTb)
      For intI = 0 To ndxList.Count - 1
        ' Get the extension number
        intExt = objXpFun.LabelExtNum(ndxList(intI))
        If (intExt < 0) Then
          ' Try get extension number from the <eLeaf> child
          ndxLeaf = ndxThis.SelectSingleNode("./child::eLeaf[@Type='Star']")
          intExt = GetTraceNumber(ndxLeaf)
        End If
        If (intExt > 0) Then
          ' There is a valid number to pursue...
          ' Note the old head number
          intHdOld = GetFeature(ndxList(intI), "dep", "hd")
          ' Find the corresponding node, and make sure it is not identical to myself
          ndxStart = ndxFor.SelectSingleNode("./descendant::eTree[@Id != " & ndxList(intI).Attributes("Id").Value & _
                                             " and tb:matches(@Label, '*-" & intExt & "')]", conTb)
          ' Find the corresponding [droot] for this particular [dnode]
          ndxStart = ndxFor.SelectSingleNode("./descendant::eTree[" & _
                      "@Id != " & ndxList(intI).Attributes("Id").Value & _
                      " and tb:deptype(self::eTree) = 'droot'" & _
                      " and (tb:matches(@Label, '*-" & intExt & "')" & _
                      "      or count(child::eLeaf[tb:matches(@Text, '*-" & intExt & "')])>0 )]", conTb)
          If (ndxStart Is Nothing) Then
            ' There is a problem: we have a numbered trace or dislocation, but cannot find the corresponding node
            Logging("Line " & ndxFor.Attributes("forestId").Value & vbCrLf & _
                    "  DoBindDnodes() cannot find [droot] node for n=" & intExt & vbCrLf & _
                    "  Forest = [" & NodeText(ndxFor) & "]")
            ' Stop
          Else
            ' Get the new head number
            intHdNew = GetFeature(ndxStart, "dep", "hd")
            ' Are they different?
            If (intHdOld <> intHdNew) Then
              ' Replace all occurrences of [intHdOld] under [ndxList(inti)] with [intHdNew]
              ndxOld = ndxList(intI).SelectNodes("./descendant-or-self::eTree[" & _
                        "child::fs[@type='dep']/child::f[@name='hd' and @value='" & intHdOld & "']]")
              For intJ = 0 To ndxOld.Count - 1
                AddFeature(pdxCurrentFile, ndxOld(intJ), "dep", "hd", intHdNew)
              Next intJ
            End If
          End If
        End If
        ' Find the "droot" of this one
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/DoBindDnodes error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoHeadValues
  ' Goal:   Retrieve the head values for the children of [ndxThis]
  ' History:
  ' 28-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function DoHeadValues(ByRef ndxThis As XmlNode, ByVal intDepId As Integer) As Boolean
    Dim ndxHead As XmlNode        ' Node that is the head of all the children of [ndxThis]
    Dim ndxChild As XmlNode       ' One child
    Dim ndxFor As XmlNode         ' My forest
    Dim ndxStart As XmlNode       ' Displaced node
    Dim ndxDeep As XmlNodeList    ' Deep node
    Dim ndxList As XmlNodeList    ' List of children
    Dim strDepType As String      ' Dependency type
    Dim intDepHd As Integer       ' Dependency "hd"
    Dim intI As Integer           ' Counter
    Dim intId As Integer          ' My ID
    Dim intExtra As Integer = 0   ' Extraposed Id

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      If (ndxThis.Name <> "eTree") Then Return False
      ' Get my ID
      intId = ndxThis.Attributes("Id").Value
      ' Am I running around in circles?
      If (colDoHdStack.Exists(intId)) AndAlso (Not colDoHdDupli.Exists(intId)) Then
        MsgBox("There is a problem in the line with forestId=" & ndxThis.SelectSingleNode("./ancestor::forest").Attributes("forestId").Value & vbCrLf & _
               "It is quite likely that this line contains different nodes using the same extension number." & vbCrLf & _
               "You need to repair this in the text file before I can proceed.")
        Stop
        ' make sure we are well understood
        bInterrupt = True : Return False
      Else
        If (bDuplicating) Then
          colDoHdDupli.AddUnique(intId)
        Else
          colDoHdStack.AddUnique(intId)
        End If
      End If
      ' Set the value of my own head straight away (so this includes normal endnodes, punctuation and so forth
      AddFeature(pdxCurrentFile, ndxThis, "dep", "hd", intDepId)
      ndxStart = Nothing : ndxHead = Nothing
      ' Action depends on my own dependency type
      strDepType = objXpFun.DepType(ndxThis)
      Select Case strDepType
        Case "tnode", "droot", "dnode"
          ' A trace-node needs to be processed if it is not an end-node
          ndxStart = ndxThis
          ' Determine the child with the [head] under [ndxThis]
          ndxHead = ndxThis.SelectSingleNode("./child::eTree[child::fs[@type='dep']/child::f[@name='rel' and @value='hd']]")
          ' Note: do not bother if [ndxHead] is nothing...
        Case "node"
          ' I myself may be a [node], but not if the only descendant is ZERO
          If ndxThis.SelectSingleNode("./descendant::eLeaf[@Type != 'Zero']") Is Nothing Then
            ndxHead = Nothing
          Else
            ' Normal child
            ndxStart = ndxThis
            ' Determine the child with the [head] under [ndxThis]
            ndxHead = ndxThis.SelectSingleNode("./child::eTree[child::fs[@type='dep']/child::f[@name='rel' and @value='hd']]")
            If (ndxHead Is Nothing) Then
              ndxFor = ndxThis.SelectSingleNode("./ancestor::forest")
              Logging("Line " & ndxFor.Attributes("forestId").Value & vbCrLf & _
                      "  Cannot find child with head feature under [" & ndxThis.Attributes("Label").Value & "] " & vbCrLf & _
                      "  Forest=[" & NodeText(ndxFor) & "]")
              ' Problem...
              'Stop

            End If
            ' Debug.Print(objXpFun.DepType(ndxThis))
          End If
        Case Else
          ndxHead = Nothing
      End Select
      ' Do we have a head to continue with?
      If (ndxHead IsNot Nothing) Then
        ' ========= DEBUGGING ===========
        ' Debug.Print("Head is: " & ndxHead.Attributes("Label").Value)
        ' Stop
        If (ndxStart Is Nothing) Then Stop
        ' ===============================
        ' Walk the child that IS the head
        If (Not DoHeadValues(ndxHead, intDepId)) Then Return False
        ' Is this head an endnode?
        ndxDeep = ndxHead.SelectNodes("./descendant::eTree[child::fs[@type='dep']/child::f[@name='id']]")
        If (GetFeature(ndxHead, "dep", "id") <> "") Then
          ' Get the @id value of this head
          intDepHd = GetFeature(ndxHead, "dep", "id")
        ElseIf (ndxDeep.Count = 1) Then
          intDepHd = GetFeature(ndxDeep(0), "dep", "id")
        ElseIf (GetFeature(ndxHead, "dep", "hd") <> "") Then
          intDepHd = GetFeature(ndxHead, "dep", "hd")
        Else
          ' We have a problem
          Stop
        End If

        ' Walk all the children that are NOT this head, and look at their dependency types
        ndxList = ndxStart.SelectNodes("./child::eTree[@Id != " & ndxHead.Attributes("Id").Value & "]")
        For intI = 0 To ndxList.Count - 1
          ' Access the child
          ndxChild = ndxList(intI)
          ' Get the dependency type
          strDepType = objXpFun.DepType(ndxChild)
          ' Action depends on dependency type
          Select Case strDepType
            Case "endnode", "node", "droot", "tnode", "punct", "dnode"   ' EK: new approach -- do all that need doing
              ' Now go down and walk the children, having the correct [intDepHd] in hand...
              If (Not DoHeadValues(ndxChild, intDepHd)) Then Return False
            Case Else
              ' We purposefully skip empty/nothing nodes
          End Select
        Next intI
      End If

        ' Return success
        Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/DoHeadValues error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoHeadValuesOKAY
  ' Goal:   Retrieve the head values for the children of [ndxThis]
  ' History:
  ' 28-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function DoHeadValuesOKAY(ByRef ndxThis As XmlNode, ByVal intDepId As Integer) As Boolean
    Dim ndxHead As XmlNode        ' Node that is the head of all the children of [ndxThis]
    Dim ndxChild As XmlNode       ' One child
    Dim ndxList As XmlNodeList    ' List of children
    Dim intDepHd As Integer       ' Dependency "hd"
    Dim intI As Integer           ' Counter

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      If (ndxThis.Name <> "eTree") Then Return False
      ' Set the value of my own head straight away
      AddFeature(pdxCurrentFile, ndxThis, "dep", "hd", intDepId)
      ' Check if this is an "end" node
      If (GetFeature(ndxThis, "dep", "id") <> "") Then
        ' This is an endnode: set its head to [intDepId]
        AddFeature(pdxCurrentFile, ndxThis, "dep", "hd", intDepId)
      Else
        ' Determine the child with the [head] under [ndxThis]
        ndxHead = ndxThis.SelectSingleNode("./child::eTree[child::fs[@type='dep']/child::f[@name='rel' and @value='hd']]")
        If (ndxHead Is Nothing) Then
          ' Problem...
          Stop
        End If
        ' Walk the child that IS the head
        If (Not DoHeadValuesOKAY(ndxHead, intDepId)) Then Return False
        ' Is this head an endnode?
        If (GetFeature(ndxHead, "dep", "id") <> "") Then
          ' Get the @id value of this head
          intDepHd = GetFeature(ndxHead, "dep", "id")
        ElseIf (GetFeature(ndxHead, "dep", "hd") <> "") Then
          intDepHd = GetFeature(ndxHead, "dep", "hd")
        Else
          ' We have a problem
          Stop
        End If
        ' Walk all the children that are not this head
        ndxList = ndxThis.SelectNodes("./child::eTree[@Id != " & ndxHead.Attributes("Id").Value & "]")
        For intI = 0 To ndxList.Count - 1
          ndxChild = ndxList(intI)
          If (Not DoHeadValuesOKAY(ndxChild, intDepHd)) Then Return False
        Next
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modConLLX/DoHeadValuesOKAY error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetDepRel
  ' Goal:   Give the dependency relation that constituent [ndxThis] has
  '         - This phrase may be <eTree> 
  '         - This depends on the language, specified in [strOp]
  ' History:
  ' 24-04-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetDepRel(ByRef ndxThis As XmlNode, ByVal strLang As String) As String
    Dim ndxList As XmlNodeList = Nothing  ' The children 
    Dim ndxWork As XmlNode = Nothing      ' Working node
    Dim strLabel As String                ' The node's label value
    ' Dim strLang As String                 ' The language to be used
    Dim intI As Integer                   ' Counter
    'Dim arConst() As String = {"*-SBJ|*-SBJ-*", "*-OB1|*-OB1-*", "*-OB2|*-OB2-*", _
    '                           "NP-POS*|N-GEN*|NS-GEN*|NEG", "NP-PRN*", _
    '   "*-LOC|*-TMP|*-INS|*-OBL|PP|PP-*|NP|NP-*|ADJ|ADJ-*|ADJP|ADJP-*|IP-INF|IP-INF-*|IP-PPL|ADVP*|CP-REL*"}
    'Dim arRel() As String = {"su", "obj1", "obj2", "det", "appos", "mod"}

    Try
      ' Validate 
      If (ndxThis Is Nothing) OrElse (ndxThis.Name <> "eTree") Then Return Nothing
      '' Determine the language
      'Select Case strOp
      '  Case "DepChe"
      '    strLang = "che"
      '  Case "DepEng"
      '    strLang = "en"
      '    ' Not yet implemented...
      '    Return "-"
      '  Case "DepNl"
      '    strLang = "nl"
      '  Case Else
      '    ' Failure
      '    Return "-"
      'End Select
      ' Check initialisation
      If (Not bLangInit) Then DoLangInit(strLang)
      ' Is this punctuation?
      If (ndxThis.SelectSingleNode("./child::eLeaf[@Type='Punct']") IsNot Nothing) Then
        ' This is punctuation
        Return "punct"
      End If
      ' Do some comparisons
      strLabel = ndxThis.Attributes("Label").Value
      For intI = 0 To arDepRel.Length - 1
        ' See if the label fits in here
        If (ndxThis.ParentNode.Name = "forest") Then
          ' Don't take parent into account
          If (DoLike(strLabel, arDepRel(intI).lablist)) Then
            ' Return the corresponding relation
            Return arDepRel(intI).rel
          End If
        Else
          ' Parent is an <eTree>
          If (DoLike(ndxThis.ParentNode.Attributes("Label").Value, arDepRel(intI).par)) AndAlso _
             (DoLike(strLabel, arDepRel(intI).lablist)) Then
            ' Return the corresponding relation
            Return arDepRel(intI).rel
          End If
        End If
      Next intI
      ' Different conditions
      If (ndxThis.ParentNode.Name = "eTree") AndAlso (ndxThis.ParentNode.Attributes("Label").Value.Replace("_", "-") Like "IP-PPL*") Then
        ' A Conjunction is a modifier
        strLabel = ndxThis.Attributes("Label").Value.Replace("_", "-")
        If (DoLike(strLabel, "CONJ|CONJ-*")) Then
          ' NOTE: I am not sure anymore why I decided to call this [det]...
          Return "det"
        End If
      End If
      ' Anything else goes to default "mod" assignment
      ' Debug.Print("Default [mod] role for: " & ndxThis.Attributes("Label").Value)
      ' Return failure
      Return "mod"
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/GetDepRel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   TravDep
  ' Goal:   Travel a node recursively, looking for dependency relations
  ' History:
  ' 30-08-2013  ERK Added "DepChe" format
  ' ------------------------------------------------------------------------------------
  Public Function TravDep(ByRef ndxThis As XmlNode, ByVal strLang As String) As Boolean
    Dim strDepRel As String = ""        ' Dependency relation type
    Dim ndxChild As XmlNode = Nothing   ' Children of me
    Dim ndxHead As XmlNode = Nothing    ' Head node
    Dim ndxList As XmlNodeList          ' List of nodes

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Bottom-Up processing, where the heads and dependencies are resolved
      Select Case ndxThis.Name
        Case "forest"
          ' The forest's children are all <root>s
          For Each ndxChild In ndxThis.SelectNodes("./child::eTree")
            ' Forest-children are roots
            AddFeature(pdxCurrentFile, ndxChild, "dep", "rel", "root")
            ' Try to process the child
            If (Not (TravDep(ndxChild, strLang))) Then Return False
          Next ndxChild
        Case "eTree"
          ' (1) Determine the head for this phrase
          If (ndxThis.SelectSingleNode("./child::eTree[tb:matches(tb:deptype(self::eTree),'*node')]", conTb) IsNot Nothing) Then
            ndxHead = GetHeadNodeDep(ndxThis, strLang)
            If (ndxHead Is Nothing) Then
              ' Just return if there is an interrupt already
              If (bInterrupt) Then Return False
              ' We're in trouble -- no head!!!
              bInterrupt = True
              Stop
              Return False
            End If
            ' (2) Mark it as head
            AddFeature(pdxCurrentFile, ndxHead, "dep", "rel", "hd")
            ' (3) Visit all children that are NOT the head
            For Each ndxChild In ndxThis.SelectNodes("./child::eTree[not(@Id=" & ndxHead.Attributes("Id").Value & ")]")
              ' (4) Add the dependency relation of this child
              AddFeature(pdxCurrentFile, ndxChild, "dep", "rel", GetDepRel(ndxChild, strLang))
            Next ndxChild
            ' Process the children, unless we are at the bottom!
            For Each ndxChild In ndxThis.SelectNodes("./child::eTree[count(child::eLeaf)=0]")
              ' Try to process the child
              If (Not (TravDep(ndxChild, strLang))) Then Return False
            Next ndxChild
          ElseIf DoLike(objXpFun.DepType(ndxThis), "droot") Then
            ' This is a dislocation root, so it functions as head for "dnode" type constituents
            AddFeature(pdxCurrentFile, ndxThis, "dep", "rel", "hd")
          Else
            ' Double check
            If (GetFeature(ndxThis, "dep", "rel").ToString = "hd") Then
              ' We should not modify nodes that are already assigned to be heads
              ' Stop
            Else
              ' Check if all is going well...
              ' Stop
              Select Case objXpFun.DepType(ndxThis)
                Case "node"
                  ' What if I am the only parent of my child?
                  ndxList = ndxThis.SelectNodes("./child::eTree")
                  If (ndxList.Count = 1) Then
                    ' I have one <eTree> child, which means that it is head
                    ndxChild = ndxList(0)
                    AddFeature(pdxCurrentFile, ndxChild, "dep", "rel", "hd")
                    ' Process the children, unless we are at the bottom!
                    For Each ndxChild In ndxThis.SelectNodes("./child::eTree[count(child::eLeaf)=0]")
                      ' Try to process the child
                      If (Not (TravDep(ndxChild, strLang))) Then Return False
                    Next ndxChild
                  ElseIf (ndxList.Count > 1) Then
                    ' This is a really weird situation
                    Stop
                    ' Process the children, unless we are at the bottom!
                    For Each ndxChild In ndxThis.SelectNodes("./child::eTree[count(child::eLeaf)=0]")
                      ' Try to process the child
                      If (Not (TravDep(ndxChild, strLang))) Then Return False
                    Next ndxChild
                  End If
                  ' If I have one <eTree> child, then I am the head of that node
                Case "empty", "troot"
                  ' We can just nodes that are really empty, as well as traces (marked as "troot")
                Case Else
                  ' Check if all is going well...
                  ' Stop
                  ' This is an end-node, and an end-node can never be a head, but only a modifier or something
                  AddFeature(pdxCurrentFile, ndxThis, "dep", "rel", GetDepRel(ndxThis, strLang))
              End Select
            End If
          End If
      End Select
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modEditor/TravDep error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   OneDutchPosToConllX
  ' Goal:   Learn the constituents from one Psdx file
  ' History:
  ' 02-09-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function OneDutchPosToConllX(ByVal strFile As String, ByVal strLang As String, ByRef strResult As String, ByVal bDoCode As Boolean) As Boolean
    Dim strForm As String             ' Vernacular form
    Dim strPosAll As String           ' Combined POS
    Dim strCpos As String             ' Compact pos
    Dim strPos As String              ' More specific POS tag
    Dim strFeat As String             ' Features
    Dim intId As Integer              ' ID per sentence
    Dim intI As Integer               ' Counter
    Dim intJ As Integer               ' Counter
    Dim intK As Integer               ' Counter
    Dim intPos As Integer             ' position in string
    Dim intPtc As Integer             ' Percentage
    ' Dim intCount As Integer           ' Number of child nodes
    Dim intSize As Integer = 0        ' Size of [arCon]
    Dim arText() As String            ' File as an array of lines
    Dim arItem() As String            ' One line
    Dim arLemma() As String           ' Lemma's
    Dim arPos() As String             ' POS's
    Dim arCon(0) As CONLLX            ' Array of CONLLX elements
    Dim colResult As New StringColl   ' Result
    Dim arForm() As String = {"&period;", "&colon;", "&comma;", "&semicolon;"}
    Dim arRewr() As String = {".", ":", ",", ","}

    Try
      ' Validate
      If (Not IO.File.Exists(strFile)) Then Return False
      ' Read the file as a text document
      Status("Reading [" & IO.Path.GetFileNameWithoutExtension(strFile) & "]...")
      arText = IO.File.ReadAllLines(strFile)
      ' Initializations
      intId = 1
      If (Not bLangInit) Then
        If (Not DoLangInit(strLang)) Then Return False
      End If
      ' Process all the individual lines
      For intI = 0 To arText.Length - 1
        ' Show progress
        intPtc = (intI + 1) * 100 \ arText.Length
        Status("Pos to Conll-X " & intPtc & "%", intPtc)
        ' Do this line?
        If (arText(intI) <> "") AndAlso (InStr(arText(intI), "<manuscript") = 0) Then
          ' Read this line
          arItem = Split(arText(intI), vbTab)
          ' Validate
          If (arItem.Length = 4) Then
            ' Okay, we have 4 components: Vernacular, Cleaned, Lemma, POS combi
            ' Split up this one into + divided items
            arLemma = Split(arItem(2), "+") : arPos = Split(arItem(3), "+")
            If (arLemma.Length <> arPos.Length) Then
              ' There is a problem, because lemma and POS do not have the same amount of items
              Logging("OneDutchPosToConllX problem: column 3 and 4 have a different number of '+' signs in line " & intI + 1)
              Return False
            End If
            ' Walk them
            For intJ = 0 To arLemma.Length - 1
              ' Get the Cpos
              strPosAll = arPos(intJ) : intPos = InStr(strPosAll, "(")
              If (intPos = 0) Then
                ' How can we determine Cpos?
                strCpos = "X" : strFeat = strPosAll.Replace(")", "")
              Else
                strCpos = UCase(Left(strPosAll, intPos - 1))
                strFeat = Mid(strPosAll, intPos + 1).Replace(")", "")
              End If
              If (strFeat = "") Then strFeat = "_"
              ' Determine POS
              strPos = CposToPos(strCpos, strFeat, arItem(0), strLang)
              ' Determine the real CPOS
              strCpos = PosToCpos(strPos)
              ' Some changes in form
              For intK = 0 To arForm.Length - 1
                If (arItem(0) = arForm(intK)) Then
                  arItem(0) = arRewr(intK) : arItem(1) = arItem(0) : arItem(2) = arItem(0)
                  arLemma(intJ) = arItem(0)
                  Exit For
                End If
              Next intK
              ' Get the form
              If (arLemma.Length = 1) Then
                strForm = arItem(0)
              Else
                strForm = arItem(0) & "_" & strCpos
              End If
              ' Make room for a new CONLLX element
              ReDim Preserve arCon(0 To intSize)
              ' Add information to this element
              With arCon(intSize)
                .Id = intId : .CposTag = strCpos : .DepRel = "_" : .Feats = strFeat : .Form = strForm
                .Head = "_" : .Lemma = arLemma(intJ) : .PdepRel = "_" : .Phead = "_" : .PosTag = strPos
              End With
              ' Increment the size of [arCon]
              intSize += 1 : intId += 1
            Next intJ
            ' Check if this finalizes a section: then re-start id
            If (arItem(0) = "&period;") OrElse (arItem(0) = ".") Then intId = 1
          End If
        End If
      Next intI
      ' Process the entire [arCon] section
      For intI = 0 To arCon.Length - 1
        ' Add one element
        With arCon(intI)
          colResult.Add(.Id & vbTab & .Form & vbTab & .Lemma & vbTab & .CposTag & vbTab & .PosTag & vbTab & _
                        .Feats & vbTab & .Head & vbTab & .DepRel & vbTab & .Phead & vbTab & .PdepRel)
        End With
      Next intI
      ' Pass back the result
      strResult = colResult.Text
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modConLLX/OneDutchPosToConllX error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   CposToPos
  ' Goal:   Convert a combination of CPOS, Features and Vernacular form to an extended POS
  ' History:
  ' 02-09-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function CposToPos(ByVal strCpos As String, ByVal strFeats As String, ByVal strVern As String, _
                             ByVal strLang As String) As String
    Dim intI As Integer     ' Counter

    Try
      ' Validate
      If (strLang <> "nl") Then Return strCpos
      If (strCpos = "") Then Return strCpos
      If (Not bLangInit) Then Return strCpos
      If (strCpos = "X") Then Return strCpos
      ' Walk through all the possibilities
      For intI = 0 To arPosRule.Length - 1
        ' Check if this one matches
        With arPosRule(intI)
          If (UCase(strCpos) = .Cpos) AndAlso _
             DoLike(strFeats, .ftlist) AndAlso DoLike(strVern, .vern) Then
            ' We are okay!
            Return .Pos
          End If
        End With
      Next intI
      ' Return failure
      Stop
      Return strCpos
    Catch ex As Exception
      ' Warn the user
      HandleErr("modConLLX/CposToPos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DeriveLanguage
  ' Goal:   Look at the directory name and derive the language from it
  ' History:
  ' 03-09-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function DeriveLanguage(ByVal strDirIn As String, ByRef strLang As String, Optional ByVal bDoAsk As Boolean = False) As Boolean
    Dim strReply As String = "Dutch"  ' Answer
    Dim ndxThis As XmlNode            ' List of <language> elements

    Try
      ' Do we have a current PSDX?
      If (pdxCurrentFile IsNot Nothing) Then
        ' Check the PSDX itself
        ndxThis = pdxCurrentFile.SelectSingleNode("//teiHeader/descendant::language[@name='org']")
        If (ndxThis IsNot Nothing) Then
          ' Get the "org" element
          strLang = ndxThis.Attributes("ident").Value
          Return True
        End If
      End If
      ' Determine the language from the directory
      If (InStr(LCase(strDirIn), "english") > 0) Then
        strLang = "en"
      ElseIf (InStr(LCase(strDirIn), "che") > 0) Then
        strLang = "che"
      ElseIf (InStr(LCase(strDirIn), "/lak") > 0) Then
        strLang = "lak"
      ElseIf (InStr(LCase(strDirIn), "lezgi") > 0) Then
        strLang = "lez"
      ElseIf (InStr(LCase(strDirIn), "dutch") > 0) Then
        strLang = "nl"
      ElseIf (InStr(LCase(strDirIn), "spanish") > 0) Then
        strLang = "es"
      ElseIf (bDoAsk) Then
        ' Ask for the language
        strReply = InputBox("What language are the texts written in?", DefaultResponse:=strReply)
        Select Case Trim(LCase(strReply))
          Case "dutch", "nederlands", "nl"
            strLang = "nl"
          Case "spanish", "spaans", "es"
            strLang = "es"
          Case "tsjetsjeens", "chechen", "noxchiin", "che"
            strLang = "che"
          Case "english", "engels", "en"
            strLang = "en"
          Case Else
            ' Return failure
            Return False
        End Select
      Else
        ' Return failure
        Return False
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modConLLX/DeriveLanguage error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   ParseOneMalt
  ' Goal:   Parse the .conll file [strFile] by using the MaltParser
  ' History:
  ' 17-09-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function ParseOneMalt(ByVal strFile As String, ByVal strLang As String) As Boolean
    Dim strCommand As String = ""
    Dim strOutFile As String = ""   ' Temporary output file
    Dim strErrFile As String = ""   ' Error output
    Dim strMco As String = ""       ' Location of .mco file
    Dim strDir As String
    Dim strOut As String = ""       ' Output
    Dim strErr As String = ""       ' Error

    Try
      ' Validate
      If (strFile = "") Then Return False
      If (Not IO.File.Exists(strFile)) Then Return False
      ' Determine temporary output file
      strOutFile = GetDocDir() & "\MaltTemp_" & strLang & ".conll"
      strErrFile = GetDocDir() & "\" & IO.Path.GetFileNameWithoutExtension(strFile) & "-malt.err"
      strDir = GetSetDir()
      ' Check presence of maltparser
      If (Not MaltCheck()) Then Return False
      ' Set the .mco for this language
      Select Case strLang
        Case "nl"
          strMco = MALT_OLDDUTCH
        Case "ch", "che", "lbe-Latn", "lez"
          strMco = MALT_CHECHEN
        Case Else
          Return False
      End Select
      ' Fill in the MaltShellIntro, depending on the language
      strCommand = strMaltShellIntro.Replace("$inp", """" & strFile & """")
      strCommand = strCommand.Replace("$out", """" & strOutFile & """")
      strCommand = strCommand.Replace("$mco", """" & strMco & """")
      ' Execute the JAVA command
      If (Not ExecuteOneJavaCommand(strCommand, strDir, strOut, strErr)) Then
        ' Something went wrong
        Logging("modConLLX/ParseOneMalt/Execute error")
        Return False
      End If
      ' Make name of error file
      IO.File.WriteAllText(strErrFile, strErr)
      ' Copy the output to the input file
      IO.File.Copy(strOutFile, strFile, True)
      ' Delete the temporary file
      IO.File.Delete(strOutFile)
      ' Return positively
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modConLLX/ParseOneMalt error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   CreateMaltMco
  ' Goal:   Use the .conll file [strFile] to create a .mco file with the Maltparser
  ' History:
  ' 12-02-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function CreateMaltMco(ByVal strFile As String, ByVal strLang As String) As Boolean
    Dim strCommand As String = ""
    Dim strOutFile As String = ""   ' Temporary output file
    Dim strErrFile As String = ""   ' Error output
    Dim strMco As String = ""       ' Location of .mco file
    Dim strDir As String
    Dim strOut As String = ""       ' Output
    Dim strErr As String = ""       ' Error

    Try
      ' Validate
      If (strFile = "") Then Return False
      If (Not IO.File.Exists(strFile)) Then Return False
      ' Determine temporary output file
      strErrFile = GetDocDir() & "\" & IO.Path.GetFileNameWithoutExtension(strFile) & "-malt.err"
      strDir = GetSetDir()
      ' Check presence of maltparser
      If (Not MaltCheck()) Then Return False
      ' Set the .mco for this language
      Select Case strLang
        Case "nl", "nld", "nld_hist"
          strMco = MALT_OLDDUTCH
        Case "ch", "che", "lbe-Latn"
          strMco = MALT_CHECHEN
        Case Else
          Return False
      End Select
      ' Fill in the MaltShellIntro, depending on the language
      strCommand = strMaltShellCreateMco.Replace("$inp", """" & strFile & """")
      strCommand = strCommand.Replace("$mco", """" & IO.Path.GetFileNameWithoutExtension(strMco) & """")
      ' Determine the output file's name
      strOutFile = strDir & "\" & strMco
      ' Execute the JAVA command
      If (Not ExecuteOneJavaCommand(strCommand, strDir, strOut, strErr)) Then
        ' Something went wrong
        Logging("modConLLX/CreateMaltMco/Execute error")
        Return False
      End If
      ' Make name of error file
      IO.File.WriteAllText(strErrFile, strErr)
      ' Make sure it is clear where the output is
      Logging("MCO has been created: " & strOutFile)
      Logging("Upload it to: http://erwinkomen.ruhosting.nl/software/stp/")
      ' Return positively
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modConLLX/CreateMaltMco error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       MaltCheck()
  ' Goal:       Check if the maltparser is there, and if not: download it
  ' History:
  ' 14-03-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function MaltCheck() As Boolean
    Dim strSave As String     ' Local destination
    Dim strMco As String      ' One MCO
    Dim strMlib As String     ' One malt lib file
    Dim intI As Integer       ' Counter
    Dim intLength As Integer  ' File length
    Dim arMco() As String = {MALT_CHECHEN, MALT_OLDDUTCH}
    Dim arMaltLib() As String = {MALT_LIB_1, MALT_LIB_2, MALT_LIB_3}
    Dim objFinf As IO.FileInfo

    Try
      ' Get destination directory
      strSave = GetSetDir()
      strMaltParser = strSave & "\" & MALT_PARSER
      ' Check for the parser and the definitions
      If (Not IO.File.Exists(strMaltParser)) Then
        ' Try to download it
        Status("Trying to download " & MALT_PARSER & "...")
        My.Computer.Network.DownloadFile(strHttpDir & MALT_PARSER, strMaltParser, "", "", True, 5000, True)
        If (Not IO.File.Exists(strMaltParser)) Then Return False
      End If
      ' Download the MCO files
      For intI = 0 To arMco.Count - 1
        strMco = strSave & "\" & arMco(intI)
        If (IO.File.Exists(strMco)) Then
          objFinf = New IO.FileInfo(strMco)
          intLength = objFinf.Length
        Else
          objFinf = Nothing : intLength = 0
        End If
        If (Not IO.File.Exists(strMco)) OrElse (intLength = 0) Then
          ' Try to download it
          Status("Trying to download " & arMco(intI) & "...")
          Try
            ' Try downloading
            My.Computer.Network.DownloadFile(strHttpDir & arMco(intI), strMco, "", "", True, 5000, True)
            ' Check for the answer
            If (Not IO.File.Exists(strMco)) Then Return False
          Catch ex As Exception
            ' There was a problem downloading this file
            Logging("MaltCheck warning: could not download file " & strMco)
          End Try
        ElseIf My.Computer.Network.IsAvailable Then
          Dim strTmp As String = strSave & "\temp.mco"
          Dim infTmp As IO.FileInfo

          ' Download it to a temporary file
          Try
            ' Try downloading
            Status("Comparing current mco with internet...")
            My.Computer.Network.DownloadFile(strHttpDir & arMco(intI), strTmp, "", "", False, 5000, True)
            ' Check for the answer
            If (Not IO.File.Exists(strTmp)) Then Return False
          Catch ex As Exception
            ' There was a problem downloading this file
            Logging("MaltCheck warning: could not download file " & strTmp)
          End Try
          ' Check sizes
          infTmp = New IO.FileInfo(strTmp)
          If (infTmp.Length <> objFinf.Length) Then
            ' COpy it
            IO.File.Copy(strTmp, strMco, True)
          End If

        End If
      Next intI
      ' Make sure there is a lib directory
      If (Not IO.Directory.Exists(strSave & "\lib")) Then IO.Directory.CreateDirectory(strSave & "\lib")
      ' Download Maltparser lib jar-files
      For intI = 0 To arMaltLib.Count - 1
        strMlib = strSave & "\lib\" & arMaltLib(intI)
        If (IO.File.Exists(strMlib)) Then
          objFinf = New IO.FileInfo(strMlib)
          intLength = objFinf.Length
        Else
          objFinf = Nothing : intLength = 0
        End If
        If (Not IO.File.Exists(strMlib)) OrElse (intLength = 0) Then
          ' Try to download it
          Status("Trying to download " & arMaltLib(intI) & "...")
          Try
            ' Try downloading
            My.Computer.Network.DownloadFile(strHttpDir & arMaltLib(intI), strMlib, "", "", True, 5000, True)
            ' Check for the answer
            If (Not IO.File.Exists(strMlib)) Then Return False
          Catch ex As Exception
            ' There was a problem downloading this file
            Logging("MaltCheck warning: could not download file " & strMlib)
          End Try
        End If
      Next intI
      ' Construct the parser command introduction
      strMaltShellIntro = MALT_INTRO & strMaltParser & """" & MALT_LOC
      ' Construct the maltparser mco creation
      strMaltShellCreateMco = MALT_INTRO & strMaltParser & """" & MALT_MCO
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/MaltCheck error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       ParseOneDelete()
  ' Goal:       Delete current parsing from the indicated file
  ' History:
  ' 25-09-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function ParseOneDelete(ByRef pdxThis As XmlDocument) As Boolean
    Dim ndxList As XmlNodeList  ' List of nodes
    Dim ndxFor As XmlNode       ' Forest
    Dim ndxThis As XmlNode      ' Node
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (pdxThis Is Nothing) Then Return False
      ' Get first forest
      ndxFor = pdxThis.SelectSingleNode("./descendant::forest")
      ' Start deleting from here
      While (ndxFor IsNot Nothing)
        ' Get all the end-nodes
        ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf)>0]")
        ' Attach the first one
        If (ndxList.Count > 0) Then
          ' Attach as first child
          ndxThis = ndxFor.PrependChild(ndxList(0))
          ' Walk all <eTree> children and attach them directly to the <forest>
          For intI = 1 To ndxList.Count - 1
            ' Add as child after [ndxThis]
            ndxThis = ndxFor.InsertAfter(ndxList(intI), ndxThis)
          Next intI
          ' Remove all siblings following after [ndxList(inti)]
          ndxList = ndxFor.SelectNodes("./child::eTree[count(child::eLeaf)=0]")
          For intI = ndxList.Count - 1 To 0 Step -1
            ' Remove this one
            ndxList(intI).RemoveAll()
            ndxFor.RemoveChild(ndxList(intI))
          Next intI
          ' Get a list of endnodes that still have an <eTree> child
          ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf)>0 and count(child::eTree)>0]")
          ' Walk this list
          For intI = ndxList.Count - 1 To 0 Step -1
            ' Get the <eTree> child
            ndxThis = ndxList(intI).SelectSingleNode("./child::eTree")
            ' Remove it
            ndxThis.RemoveAll()
            ndxList(intI).RemoveChild(ndxThis)
          Next intI
        End If
        ' Go  to next forest
        ndxFor = ndxFor.NextSibling
      End While
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/ParseOneDelete error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       ConllToHtml()
  ' Goal:       Convert a CONLLX file into an html table
  ' History:
  ' 02-10-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function ConllToHtml(ByVal strFile As String) As String
    Dim colThis As New StringColl ' 
    Dim arLine() As String
    Dim intI As Integer

    Try
      ' Validate
      If (strFile = "") OrElse (Not IO.File.Exists(strFile)) Then Return ""
      arLine = IO.File.ReadAllLines(strFile)
      colThis.Add("<html><body><table>")
      For intI = 0 To arLine.Length - 1
        colThis.Add("<tr><td>" & arLine(intI).Replace(vbTab, "</td><td>") & "</td></tr>")
      Next intI
      colThis.Add("</table></body></html>")
      ' Return the result
      Return colThis.Text
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/ConllToHtml error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return ""
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       CreateDepPos()
  ' Goal:       Create a POS-based dependency datastructure in a file
  ' History:
  ' 08-10-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function CreateDepPos(ByVal strFile As String, ByVal strLang As String) As Boolean
    Dim ndxFor As XmlNode = Nothing   ' Forest
    Dim ndxList As XmlNodeList        ' List of nodes
    Dim colOut As New StringColl      ' Output collection
    Dim strPos As String              ' Part of speech
    Dim strLemma As String            ' Lemma of this entry
    Dim strVern As String             ' Vernacular for this entry
    Dim strCpos As String             ' Compact POS for this entry
    Dim strOut As String              ' Output string
    Dim intI As Integer               ' Counter
    Dim intPtc As Integer             ' Percentage
    Dim intDepId As Integer           ' Dependency Id within dataset
    Dim intForestId As Integer        ' ForestId
    Dim intEtreeId As Integer

    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Return False
      ' We need a language description for the translation from Pos to Cpos
      If (Not DoLangInit(strLang)) Then bInterrupt = True : Return False
            ' Start output
      colOut.Add("<Dependency>") : intDepId = 0
      ' Get first forest
      If (Not GetFirstForest(pdxCurrentFile, ndxFor)) Then Return False
      ' Walk through all forests
      While (ndxFor IsNot Nothing)
        ' Get this id
        intForestId = ndxFor.Attributes("forestId").Value
        ' Get a list of endnodes
        ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf)>0]")
        For intI = 0 To ndxList.Count - 1
          ' Get values for this node
          strPos = ndxList(intI).Attributes("Label").Value : strCpos = PosToCpos(strPos)
          strVern = XmlEscape(NodeText(ndxList(intI))) : strLemma = strVern
          intEtreeId = ndxList(intI).Attributes("Id").Value
          intDepId += 1
          strOut = " <Dep DepId=" & """" & intDepId & """" & " forestId=" & """" & intForestId & """" & _
                     " Id=" & """" & intI + 1 & """" & " Vern=" & """" & strVern & """" & " Lemma=" & """" & strLemma & """" & _
                     " Cpos=" & """" & strCpos & """" & " Pos=" & """" & strPos & """" & " Hd = " & """" & "_" & """" & _
                     " Rel=" & """" & "_" & """" & " EtreeId=" & """" & intEtreeId & """" & " />"
          ' Adaptations
          strOut = strOut.Replace("""" & """" & """" & " ", """" & "&quot;" & """" & " ")
          colOut.Add(strOut)
        Next intI
        ' Next forest
        ndxFor = ndxFor.NextSibling
      End While
      ' Finish file
      colOut.Add("</Dependency>")
      IO.File.WriteAllText(strFile, colOut.Text, System.Text.Encoding.UTF8)
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/CreateDepPos error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       CreateDepFull()
  ' Goal:       Create a fulle dependency datastructure in a file
  ' History:
  ' 19-06-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function CreateDepFull(ByVal strFile As String, ByVal strLang As String) As Boolean
    Dim ndxFor As XmlNode = Nothing   ' Forest
    Dim ndxList As XmlNodeList        ' List of nodes
    Dim colOut As New StringColl      ' Output collection
    Dim strPos As String              ' Part of speech
    Dim strLemma As String            ' Lemma of this entry
    Dim strVern As String             ' Vernacular for this entry
    Dim strCpos As String             ' Compact POS for this entry
    Dim strOut As String              ' Output string
    Dim strDhd As String               ' My head
    Dim strDrel As String             ' Dep rel
    Dim strDid As String              ' ID
    Dim intI As Integer               ' Counter
    Dim intPtc As Integer             ' Percentage
    Dim intDepId As Integer           ' Dependency Id within dataset
    Dim intForestId As Integer        ' ForestId
    Dim intEtreeId As Integer

    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Return False
      ' We need a language description for the translation from Pos to Cpos
      If (Not DoLangInit(strLang)) Then bInterrupt = True : Return False
      ' Start output
      colOut.Add("<Dependency>") : intDepId = 0
      ' Get first forest
      If (Not GetFirstForest(pdxCurrentFile, ndxFor)) Then Return False
      ' Walk through all forests
      While (ndxFor IsNot Nothing)
        ' Get this id
        intForestId = ndxFor.Attributes("forestId").Value
        ' Get a list of endnodes
        ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf)>0]")
        For intI = 0 To ndxList.Count - 1
          ' Get values for this node
          strPos = ndxList(intI).Attributes("Label").Value : strCpos = PosToCpos(strPos)
          strVern = XmlEscape(NodeText(ndxList(intI))) : strLemma = strVern
          intEtreeId = ndxList(intI).Attributes("Id").Value
          strDid = GetFeature(ndxList(intI), "con", "id")
          strDhd = GetFeature(ndxList(intI), "con", "hd")
          strDrel = GetFeature(ndxList(intI), "con", "drel")
          intDepId += 1
          strOut = " <Dep DepId=" & """" & intDepId & """" & " forestId=" & """" & intForestId & """" & _
                     " Id=" & """" & strDid & """" & " Vern=" & """" & strVern & """" & " Lemma=" & """" & strLemma & """" & _
                     " Cpos=" & """" & strCpos & """" & " Pos=" & """" & strPos & """" & " Hd = " & """" & strDhd & """" & _
                     " Rel=" & """" & strDrel & """" & " EtreeId=" & """" & intEtreeId & """" & " />"
          ' Adaptations
          strOut = strOut.Replace("""" & """" & """" & " ", """" & "&quot;" & """" & " ")
          colOut.Add(strOut)
        Next intI
        ' Next forest
        ndxFor = ndxFor.NextSibling
      End While
      ' Finish file
      colOut.Add("</Dependency>")
      IO.File.WriteAllText(strFile, colOut.Text, System.Text.Encoding.UTF8)
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/CreateDepFull error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       DepSelChange()
  ' Goal:       Check if there is a change in forest
  ' History:
  ' 09-10-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function DepSelChange(ByVal intForId As Integer, ByRef ndxFor As XmlNode) As Boolean
    Try
      ' Validate
      If (intForId < 0) Then Return False
      ' Set forest to correct current value
      ndxFor = loc_ndxCurrentForest
      ' Was there a change?
      If (intForId <> loc_intCurrentForestId) OrElse (ndxFor Is Nothing) Then
        ' There is a change in forest
        loc_intCurrentForestId = intForId
        ndxFor = pdxCurrentFile.SelectSingleNode("./descendant::forest[@forestId = " & intForId & "]")
        loc_ndxCurrentForest = ndxFor
        Return True
      End If
      ' No need to change
      Return False
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/DepSelChange error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       GetCurrentDepRow()
  ' Goal:       Get the currently selected <eTree> node
  ' History:
  ' 16-10-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function GetCurrentDepRow(ByRef dtrThis As DataRow) As Boolean
    Dim intDepId As Integer   ' ID of currently selected entry
    Dim dtrFound() As DataRow ' Result of SELECT

    Try
      ' Get the currently selected item
      intDepId = objDepEd.SelectedId
      If (intDepId < 0) Then Return Nothing
      dtrFound = tdlDependency.Tables("Dep").Select("DepId=" & intDepId)
      If (dtrFound.Length > 0) Then
        dtrThis = dtrFound(0)
        Return True
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/DepSelChange error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       GetCurrentDepNode()
  ' Goal:       Get the currently selected <eTree> node
  ' History:
  ' 16-10-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function GetCurrentDepNode() As XmlNode
    Dim intEtreeId As Integer ' ID of <eTree>
    Dim dtrThis As DataRow = Nothing  ' Currently selected row
    Dim ndxThis As XmlNode = Nothing  ' Currently selected <eTree> node

    Try
      ' Get the currently selected row
      If (Not GetCurrentDepRow(dtrThis)) Then Return Nothing
      ' Get the associated <eTree>
      intEtreeId = CInt(dtrThis.Item("EtreeId").ToString)
      ' Find this node
      ndxThis = loc_ndxCurrentForest.SelectSingleNode("./descendant::eTree[@Id = " & intEtreeId & "]")
      ' Return the result
      Return ndxThis
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/GetCurrentDepNode error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return Nothing
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       DepChange()
  ' Goal:       Process the change of type [strType] with value [strValue]
  ' History:
  ' 16-10-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Public Sub DepChange(ByVal strType As String, ByVal strValue As String)
    Dim dtrThis As DataRow = Nothing  ' Currently selected row
    Dim ndxThis As XmlNode    ' Currently selected <eTree> node
    Dim ndxLeaf As XmlNode    ' <eLeaf> child

    Try
      ' Validate
      If (loc_ndxCurrentForest Is Nothing) Then Return
      If (objDepEd.IsSelecting) Then Exit Sub
      ' Action depends on [strType]
      Select Case strType
        Case "Org"
          ' Adapt original
          ndxThis = loc_ndxCurrentForest.SelectSingleNode("./child::div[@lang='org']/child::seg")
          If (ndxThis IsNot Nothing) Then
            ' Adapt the value
            ndxThis.InnerText = strValue
          End If
        Case "Eng"
          ' Adapt english translation
          ndxThis = loc_ndxCurrentForest.SelectSingleNode("./child::div[@lang='eng']/child::seg")
          If (ndxThis IsNot Nothing) Then
            ' Adapt the value
            ndxThis.InnerText = strValue
          End If
        Case Else
          ' Get currently selected row
          If (Not GetCurrentDepRow(dtrThis)) Then Exit Sub
          ' Get the currently selected item
          ndxThis = GetCurrentDepNode()
          ' Anything found?
          If (ndxThis IsNot Nothing) Then
            ' Action depends on [strType]
            Select Case strType
              Case "Vern"
                ' Adapt the value of the <eLeaf> child
                ndxLeaf = ndxThis.SelectSingleNode("./child::eLeaf")
                If (ndxLeaf IsNot Nothing) Then
                  ndxLeaf.Attributes("Text").Value = strValue
                End If
                ' Adapt row-value
                dtrThis.Item("Vern") = strValue
              Case "Lemma"
                ' Check if there is a lemma feature at all
                If (HasFeature(ndxThis, "dep", "l")) Then
                  ' Adapt the lemma feature
                  AddFeature(pdxCurrentFile, ndxThis, "dep", "l", strValue)
                End If
                ' Adapt row-value
                dtrThis.Item("Lemma") = strValue
              Case "Pos"
                ' Adapt the Label
                ndxThis.Attributes("Label").Value = strValue
                ' Adapt row-value
                dtrThis.Item("Pos") = strValue
                ' Also adapt the CPOS value
                frmMain.tbDepCpos.Text = PosToCpos(strValue)
              Case "Cpos"
                ' Adapt the cpos feature
                If (HasFeature(ndxThis, "dep", "cpos")) Then AddFeature(pdxCurrentFile, ndxThis, "dep", "cpos", strValue)
                ' Adapt row-value
                dtrThis.Item("Cpos") = strValue
              Case "Id"
                ' Adapt the @dep/id feature
                If (HasFeature(ndxThis, "dep", "id")) Then
                  AddFeature(pdxCurrentFile, ndxThis, "dep", "id", strValue)
                ElseIf (HasFeature(ndxThis, "con", "id")) Then
                  AddFeature(pdxCurrentFile, ndxThis, "con", "id", strValue)
                End If
                ' Adapt row-value
                dtrThis.Item("Id") = strValue
              Case "Head"
                ' Adapt the @dep/head feature
                If (HasFeature(ndxThis, "dep", "hd")) Then
                  AddFeature(pdxCurrentFile, ndxThis, "dep", "hd", strValue)
                ElseIf (HasFeature(ndxThis, "con", "hd")) Then
                  AddFeature(pdxCurrentFile, ndxThis, "con", "hd", strValue)
                End If
                ' Adapt row-value
                dtrThis.Item("Hd") = strValue
              Case "Rel"
                ' Adapt the @dep/rel feature
                If (HasFeature(ndxThis, "dep", "rel")) Then
                  AddFeature(pdxCurrentFile, ndxThis, "dep", "rel", strValue)
                ElseIf (HasFeature(ndxThis, "con", "rel")) Then
                  AddFeature(pdxCurrentFile, ndxThis, "con", "drel", strValue)
                End If
                ' Adapt row-value
                dtrThis.Item("Rel") = strValue
            End Select
          End If

      End Select

    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/DepChange error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
    End Try
  End Sub
  '----------------------------------------------------------------------------------------
  ' Name:       DepAddWordBelow()
  ' Goal:       Attach the one below me to me
  ' History:
  ' 16-10-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function DepAddWordBelow() As Boolean
    Dim ndxThis As XmlNode    ' Currently selected <eTree> node
    Dim intDepId As Integer   ' ID of currently selected entry

    Try
      ' Get the currently selected item
      ndxThis = GetCurrentDepNode() : If (ndxThis Is Nothing) Then Return False
      ' Check if there is any <eTree> following me
      If (ndxThis.SelectSingleNode("./following-sibling::eTree") Is Nothing) Then Return False
      ' Get currently selected dep
      intDepId = objDepEd.SelectedId
      ' Select the one below me
      objDepEd.SelectDgvId(intDepId + 1)
      ' Do join word above
      If (Not DepJoinWordAbove()) Then Return False
      ' Set the correct selection
      objDepEd.SelectDgvId(intDepId)
      ' Otherwise we are okay
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/DepAddWordBelow error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       DepSplitLineBelow()
  ' Goal:       Attach the current value to the one above me, which is the <eTree> preceding me
  ' History:
  ' 16-10-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function DepSplitLineBelow() As Boolean
    Dim dtrDep() As DataRow   ' Result of [Select]
    Dim ndxList As XmlNodeList  ' List of nodes
    Dim ndxThis As XmlNode    ' Currently selected <eTree> node
    Dim ndxNext As XmlNode    ' The next node, which should become the first one of the new forest
    Dim ndxNxFor As XmlNode   ' The new forest
    Dim ndxThFor As XmlNode   ' Current forest
    Dim ndxDiv As XmlNode     ' Div node
    Dim ndxSeg As XmlNode     ' Seg node
    Dim ndxAnc As XmlNode     ' Ancestor <forestGrp>
    Dim ndxWork As XmlNode    ' Working node
    Dim arEng() As String     ' English text
    Dim arOrg() As String     ' Original text
    Dim intI As Integer       ' Counter
    Dim intId As Integer      ' @Id value for current node
    Dim intForId As Integer   ' Selected forest id
    Dim intDepId As Integer   ' ID of currently selected entry

    Try
      ' Validate
      If (loc_ndxCurrentForest Is Nothing) Then Return False
      ' Get the current forest id
      intForId = loc_intCurrentForestId
      ' Get currently selected dep
      intDepId = objDepEd.SelectedId
      ' Get the currently selected item
      ndxThis = GetCurrentDepNode() : If (ndxThis Is Nothing) Then Return False
      ' Get forest
      ndxThFor = ndxThis.ParentNode : If (ndxThFor Is Nothing) OrElse (ndxThFor.Name <> "forest") Then Return False
      ndxAnc = ndxThFor.ParentNode : If (ndxAnc Is Nothing) OrElse (ndxAnc.Name <> "forestGrp") Then Return False
      ' Get the inner texts
      arEng = Split(ndxThFor.SelectSingleNode("./child::div[@lang='eng']/seg").InnerText, vbLf)
      arOrg = Split(ndxThFor.SelectSingleNode("./child::div[@lang='org']/seg").InnerText, vbLf)
      ' Validate
      If (arEng.Length <> 2) OrElse (arOrg.Length <> 2) Then Return False
      ' Get the next node
      ndxNext = ndxThis.SelectSingleNode("./following-sibling::eTree")
      If (ndxNext Is Nothing) Then Return False
      ' Create a new forest node
      SetXmlDocument(pdxCurrentFile)
      ndxNxFor = AddXmlChildAfter(ndxAnc, ndxThFor, "forest", "forestId", intForId + 1, "attribute", _
                                  "TextId", ndxThFor.Attributes("TextId").Value, "attribute", _
                                  "File", ndxThFor.Attributes("File").Value, "attribute", _
                                  "Location", ndxThFor.Attributes("TextId").Value & "." & intForId + 1, "attribute")
      ' Add <div> and <seg> stuff for [org]
      ndxDiv = AddXmlChild(ndxNxFor, "div", "lang", "org", "attribute", _
                                            "seg", arOrg(1), "child")
      ndxThFor.SelectSingleNode("./child::div[@lang='org']/seg").InnerText = arOrg(0)
      ' Add <div> and <seg> stuff for [eng]
      ndxDiv = AddXmlChild(ndxNxFor, "div", "lang", "eng", "attribute", _
                                            "seg", arEng(1), "child")
      ndxThFor.SelectSingleNode("./child::div[@lang='eng']/seg").InnerText = arEng(0)
      ' Move all the <eTree> nodes starting from [ndxNext]
      ndxWork = ndxNext
      While (ndxWork IsNot Nothing)
        ' Move the [ndxWork] node
        ndxNxFor.AppendChild(ndxWork)
        ' Get next <eTree> node
        ndxWork = ndxThis.SelectSingleNode("./following-sibling::eTree")
      End While
      ' Adapt the forest ids of the remaining forests
      ndxWork = ndxNxFor.SelectSingleNode("./following-sibling::forest")
      intForId = ndxNxFor.Attributes("forestId").Value
      While (ndxWork IsNot Nothing)
        ' Adapt the forest id of this one
        intForId += 1
        ndxWork.Attributes("forestId").Value = intForId
        ndxWork.Attributes("Location").Value = ndxWork.Attributes("TextId").Value & "." & intForId
        ' Go to next forest
        ndxWork = ndxWork.SelectSingleNode("./following-sibling::forest")
      End While
      ' Get all the rows with a forest id that is too low
      intForId = ndxNxFor.Attributes("forestId").Value
      dtrDep = tdlDependency.Tables("Dep").Select("forestId >= " & intForId)
      For intI = 0 To dtrDep.Length - 1
        ' Increment the forest Id number
        dtrDep(intI).Item("forestId") += 1
      Next intI
      ' Get all the <eTree> nodes that need a new @Id value
      ndxList = ndxNxFor.SelectNodes("./child::eTree")
      For intI = 0 To ndxList.Count - 1
        ' Get this one's etree id
        intId = ndxList(intI).Attributes("Id").Value
        ' Adapt the @Id in the datatable
        dtrDep = tdlDependency.Tables("Dep").Select("eTreeId = " & intId)
        If (dtrDep.Length > 0) Then
          dtrDep(0).Item("Id") = intI + 1
          dtrDep(0).Item("forestId") = intForId
        End If
      Next intI
      ' Accept changes
      tdlDependency.AcceptChanges()
      ' Otherwise we are okay
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/DepSplitLineBelow error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       DepJoinWordAbove()
  ' Goal:       Attach the current value to the one above me, which is the <eTree> preceding me
  ' History:
  ' 16-10-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function DepJoinWordAbove() As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim intDepId As Integer   ' ID of currently selected entry
    Dim intId As Integer      ' @dep/id value of currently selected one
    Dim ndxThis As XmlNode    ' Currently selected <eTree> node
    Dim ndxPrev As XmlNode    ' Previous <eTree>
    Dim ndxThisL As XmlNode   ' Current <eLeaf>
    Dim ndxPrevL As XmlNode   ' Previous <eLeaf>
    Dim dtrThis As DataRow = Nothing  ' Currently selected row
    Dim intForId As Integer   ' Selected forest id
    Dim intI As Integer       ' Counter

    Try
      ' Validate
      If (loc_ndxCurrentForest Is Nothing) Then Return False
      ' Get the current forest id
      intForId = loc_intCurrentForestId
      ' Get currently selected dep
      intDepId = objDepEd.SelectedId
      ' Get the currently selected item
      ndxThis = GetCurrentDepNode() : If (ndxThis Is Nothing) Then Return False
      ' Get <eLeaf>
      ndxThisL = ndxThis.SelectSingleNode("./child::eLeaf") : If (ndxThisL Is Nothing) Then Return False
      ' Get preceding node
      ndxPrev = ndxThis.SelectSingleNode("./preceding-sibling::eTree[1]")
      ' Does previous exist?
      If (ndxPrev Is Nothing) Then Return False
      ' Get <eLeaf> of ndxPrev
      ndxPrevL = ndxPrev.SelectSingleNode("./child::eLeaf") : If (ndxPrevL Is Nothing) Then Return False
      ' Attach my values to the one above
      ndxPrevL.Attributes("Text").Value &= ndxThisL.Attributes("Text").Value
      ' Delete <prev>
      ndxThis.RemoveAll()
      ndxPrev.ParentNode.RemoveChild(ndxThis)
      ' Remove the datarow
      If (Not GetCurrentDepRow(dtrThis)) Then Return False
      ' Get the @dep/id value
      intId = dtrThis.Item("Id")
      dtrThis.Delete()
      tdlDependency.AcceptChanges() : Application.DoEvents()
      ' Get the datarow containing "prev"
      dtrFound = tdlDependency.Tables("Dep").Select("forestId = " & intForId & _
                      " AND EtreeId = " & ndxPrev.Attributes("Id").Value)
      If (dtrFound.Length > 0) Then
        dtrThis = dtrFound(0)
        ' Change this one's [Vern] value
        dtrThis.Item("Vern") = ndxPrevL.Attributes("Text").Value
      End If
      ' Adapt all the necessary @dep/id values
      dtrFound = tdlDependency.Tables("Dep").Select("forestId = " & intForId & _
                      " AND Id >= " & intId)
      ' Walk all the results
      For intI = 0 To dtrFound.Length - 1
        ' Adapt the value
        dtrFound(intI).Item("Id") -= 1
      Next intI
      ' Adapt all necessary DepId values - renumber them
      dtrFound = tdlDependency.Tables("Dep").Select("", "DepId ASC")
      For intI = 0 To dtrFound.Length - 1
        dtrFound(intI).Item("DepId") = intI + 1
      Next intI
      ' Select the previous one
      objDepEd.SelectDgvId(intDepId - 1)
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/DepJoinWordAbove error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       ShowCurrentDep()
  ' Goal:       Make sure current dependency is being shown
  ' History:
  ' 28-05-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function ShowCurrentDep() As Boolean
    Dim intForestId As Integer        ' ID of current forest
    Dim ndxFor As XmlNode = Nothing   ' Forest
    Dim ndxThis As XmlNode = Nothing  ' Node

    Try
      ' Validate
      If (objDepEd Is Nothing) Then Return True
      ' Adapt the whole text filter
      If (bDepWholeText) Then
        objDepEd.Filter = ""
      Else
        ' Find out where we are
        If (Not GetCurrentForest(ndxFor, ndxThis)) Then Logging("Could not find currently selected forest") : Return False
        ' Set a filter
        If (ndxFor IsNot Nothing) Then
          intForestId = ndxFor.Attributes("forestId").Value
        Else
          intForestId = 1
        End If
        bDepChg = True
        objDepEd.Filter = "forestId = " & intForestId
        bDepChg = False
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/ShowCurrentDep error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       DepInsertWordBelow()
  ' Goal:       Insert a new word below me
  ' History:
  ' 28-12-2013  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function DepInsertWordBelow() As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim intDepId As Integer   ' ID of currently selected entry
    Dim intId As Integer      ' @dep/id value of currently selected one
    Dim ndxThis As XmlNode    ' Currently selected <eTree> node
    Dim ndxFor As XmlNode     ' Current forest
    Dim ndxNew As XmlNode     ' New node
    Dim ndxLeaf As XmlNode    ' New <eLeaf> node
    Dim ndxMax As XmlNode     ' <eTree> with the highest @Id
    Dim dtrThis As DataRow = Nothing  ' Currently selected row
    Dim intForId As Integer   ' Selected forest id
    Dim intEtreeId As Integer ' New eTree Id
    Dim intI As Integer       ' Counter

    Try
      ' Validate
      If (loc_ndxCurrentForest Is Nothing) Then Return False
      ' Get the current forest id
      intForId = loc_intCurrentForestId
      ' Get currently selected dep
      intDepId = objDepEd.SelectedId
      ' Validate: try get the currently selected item
      ndxThis = GetCurrentDepNode() : If (ndxThis Is Nothing) Then Return False
      ' Now get the node above me (probably the forest)
      ndxFor = ndxThis.ParentNode
      ' Validate: must exist
      If (ndxFor Is Nothing) OrElse (ndxThis.Name <> "eTree") Then Return False
      ' Get ID of current row
      If (Not GetCurrentDepRow(dtrThis)) Then Return False
      ' Get the @dep/id value
      intId = dtrThis.Item("Id")
      ' Create and insert a new <eTree> node
      ndxNew = CreateNewEtree(pdxCurrentFile)
      ndxFor.InsertAfter(ndxNew, ndxThis)
      ndxLeaf = CreateNewEleaf(pdxCurrentFile)
      ndxNew.AppendChild(ndxLeaf)
      ' Make sure the <eTree> has a valid Id
      'ndxMax = pdxCurrentFile.SelectSingleNode( _
      '  "( for $nd in //eTree" & _
      '  "  let $intId := $nd/@Id" & _
      '  "  order by $intId descending" & _
      '  "  return $nd)[1]")
      'intEtreeId = ndxMax.Attributes("Id").Value + 1
      intEtreeId = pdxCurrentFile.SelectNodes("./descendant::eTree").Count + 1
      ndxNew.Attributes("Id").Value = intEtreeId
      ' Adapt all the necessary @dep/id values
      dtrFound = tdlDependency.Tables("Dep").Select("forestId = " & intForId & _
                      " AND Id >= " & intId + 1)
      ' Walk all the results
      For intI = 0 To dtrFound.Length - 1
        ' Adapt the value
        dtrFound(intI).Item("Id") += 1
      Next intI
      ' Add a new row into the dependency view
      dtrThis = AddOneDataRow(tdlDependency, "Dep", "DepId", "")
      With dtrThis
        .Item("Id") = intId + 1 : .Item("EtreeId") = intEtreeId
        intDepId = .Item("DepId")
        .Item("forestId") = intForId : .Item("Hd") = "_" : .Item("Rel") = "_"
      End With
      ' Select the new item
      objDepEd.SelectDgvId(intDepId)
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/DepInsertWordBelow error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
End Module
