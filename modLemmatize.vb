Imports System.Xml
Imports System.Text.RegularExpressions
Module modLemmatize
  ' ========================================= LOCAL STRUCTURES ==================================================
  Private Structure WebsterPos
    Dim cat As String   ' Category according to webster's
    Dim examp As String ' Example
    Dim pos As String   ' POS list, divided by semicolumn
    Dim feats As String ' List of features, divided by semicolumn
  End Structure
  ' ========================================= LOCAL VARIABLES ==================================================
  Private arFormName() As String
  Private arFormPos() As String
  Private arFormFeat() As String
  Private loc_strGutenberg As String = "Gutenberg.txt"
  Private bInitGutenberg As Boolean = False
  Public bMorphVernPosChanged As Boolean = False
  Private loc_strEndIndicator As String = "*—*|*. From *|*.*"
  Private loc_strNoEnd As String = "*.,*|*;*"
  Private loc_arVerbCat() As String = {"1", "1 p. pr. pl.", "1 pr. pl.", "1 pr. s.", "2 pr. s.", "2 pr. s. subj.", _
                                       "2 pt. s.", "pr. p.", "pr. pl.", "pr. s.", "pr. p.", "pt.", "pt. pl.", _
                                       "pt. pr.", "pt. s.", "pt. s. impers.", "pt. s. reflex.", "pt. t.", "pt.-pr. s. impers."}
  Private objSim As New Similarity          ' We need this in order to perform similarity measurements
  Private arWebster() As WebsterPos         ' Array of Cat > POS entries
  Private loc_strWebsterPosFile As String = "d:\data files\corpora\dictionaries\WebsterPos.csv"
  Private bWebsterPos As Boolean = False
  Private colMEDpos As New StringColl
  Private Const REW_BEST As Integer = 10
  Private Const REW_SPCHG As Integer = 6
  Private Const REW_NOCHG As Integer = 4
  Private Const REW_WORSE As Integer = 2
  Private colMis As New StringColl  ' Missing POS stuff
  Private arPreVb() As String = {"aefter", "nither", "under", "forth", "thurh", "aweg", "fore", "from", "ofer", "uppa", _
                             "with", "for", "ful", "uta", "ymb", "am", "an", "be", _
                             "of", "on", "to", "up", "ut"}
  Private arDictAbbr() As String = {"BT", "MED", "OED"}
  Private arDictUrl() As String = {"http://www.bosworthtoller.com/", _
                                   "http://quod.lib.umich.edu/cgi/m/mec/med-idx?type=id&id=MED", _
                                   "http://www.oed.com/view/Entry/"}
  Private loc_wbDom As New WebBrowser
  Private loc_strCss As String = "body{margin:0;padding:0;background:#f5f6f7;font:12px/170% Verdana,sans-serif;color:#707070;}" & _
    "input{font:12px/100% Verdana,sans-serif;color:#707070;}" & _
    ".entry-content{border:1px solid #E9EFF3;clear:both;padding:0.5em;text-align:justify;font-size:110%;}" & _
    ".entry-content p{margin-top:0;margin-bottom:0;}" & _
    ".entry-content i{font-weight:bold;}" & _
    ".entry-content b{font-weight:bolder;color:#303030;}" & _
    ".intext-title{color:#303030;float:left;font-weight:bolder;margin-right:0.3em;}" & _
    "input#edit-3{width:90%;}" & _
    ".ext-links form{display:inline;}" & _
    ".diacritics{text-align:center;border-style:dotted;border-width:thin;border-color:#b4d7f0;margin-left:auto;margin-right:auto;width:600px;padding:5px;font-size:90%;}" & _
    "edit{color:red;font-weight:bold;padding-right:5px;padding-left:5px;}" & _
    "#entry-content grammar,.cke_show_borders grammar{display:block;margin-bottom:25px;}" & _
    "grammar{color:brown;}" & _
    "#entry-content sense,.cke_show_borders sense{display:block;padding:5px;padding-left:27px;text-indent:-22px;clear:both;border-style:dotted;border-bottom:none;border-width:thin;border-color:#D8D1D1;}" & _
    "#entry-content snum,.cke_show_borders snum{font-weight:bold;font-size:110%;color:black;}" & _
    "#entry-content def,.cke_show_borders def{text-indent:0px;font-size:110%;background-color:#F7F7F7;padding:5px 5px 5px 10px;color:#454545;}" & _
    "def{font-weight:bold;color:#black;}" & _
    "#entry-content examples,.cke_show_borders examples{display:block;text-indent:0px;}" & _
    "#entry-content ex,.cke_show_borders ex{list-style:disc inside  none;display:list-item;padding:15px 5px 0px 25px;text-indent:-12px;}#entry-content oe,.cke_show_borders oe{width:40%;text-align:left;font-size:110%;}" & _
    "#entry-content hebrew,.cke_show_borders hebrew{color:brown;unicode-bidi:bidi-override;}" & _
    "#entry-content greek,.cke_show_borders greek{color:brown;}" & _
    "#entry-content rune,.cke_show_borders rune{color:brown;font-size:150%;font-family:Junicode,Segoe UI,FreeSerif,Verdana,sans-serif;}" & _
    "#entry-content thornbar,.cke_show_borders thornbar{font-family:Segoe UI,FreeSerif,Cardo,DejaVu Sans,Junicode,Verdana,sans-serif;}" & _
    "#entry-content trans,.cke_show_borders trans{font-weight:bold;color:#4E4B74;width:40%;text-align:left;font-style:italic;}" & _
    "#entry-content trans:before{content:'Â»';}" & _
    "#entry-content trans:after{content:'Â«';}" & _
    "#entry-content references,.cke_show_borders references{font-weight:bold;color:#6A83AD;display:block;text-align:right;font-size:8pt;width:100%;border-bottom-style:dotted;border-width:thin;border-color:#D8D1D1;float:right;}" & _
    "#entry-content  a.tips{border-bottom:1px dotted;}" & _
    "#entry-content references a.tips{font-size:8pt;}" & _
    "#entry-content etym,.cke_show_borders etym{background-color:#FDF2F2;display:block;padding:10px;margin:10px;clear:both;}" & _
    "#entry-content der,.cke_show_borders der{background-color:#FFFCDA;display:block;padding:10px;margin:10px;clear:both;}" & _
    "#entry-content see,.cke_show_borders see{background-color:#F2FAFF;display:block;padding:10px;margin:10px;clear:both;}.cite{float:right;}"

  ' ==============================================================================================================
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   InitGutenberg
  ' Goal:   Load settings for Gutenberg ME dictionary and put them into arrays [arFormName] etc.
  ' History:
  ' 02-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function InitGutenberg() As Boolean
    Dim strFile As String = ""
    Dim arLine() As String
    Dim arText() As String
    Dim intI As Integer
    Dim intSize As Integer

    Try
      ' Do we need to continue?
      If (bInitGutenberg) Then Return True
      ' Read the Gutenberg definitions
      strFile = IO.Path.GetDirectoryName(Application.ExecutablePath) & "\" & loc_strGutenberg
      If (Not IO.File.Exists(strFile)) Then Return False
      arLine = IO.File.ReadAllLines(strFile)
      ' Note the length
      intSize = arLine.Length
      ' Define the sizes of the other arrays
      ReDim arFormFeat(0 To intSize - 1)
      ReDim arFormPos(0 To intSize - 1)
      ReDim arFormName(0 To intSize - 1)
      ' Fill the other arrays
      For intI = 0 To intSize - 1
        arText = Split(arLine(intI), ",")
        arFormName(intI) = arText(0)
        arFormPos(intI) = arText(1)
        arFormFeat(intI) = arText(2)
      Next intI
      ' Return positively
      bInitGutenberg = True
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/InitGutenberg error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GutenbergIndex
  ' Goal:   Get the index of the [strForm] within arFormName
  ' History:
  ' 03-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function GutenbergIndex(ByVal strForm As String) As Integer
    Dim intI As Integer ' Counter

    Try
      ' Validate
      If (strForm = "") Then Return -1
      ' Try to get it
      For intI = 0 To arFormName.Length - 1
        If (strForm = arFormName(intI)) Then
          ' ============= DEBUG ============== 
          ' Check if this contains a number
          ' If (InStr(strForm, "1") > 0 OrElse InStr(strForm, "2") > 0) Then Stop
          ' ==================================
          ' Return the index
          Return intI
        End If
      Next intI
      ' Return failure
      Return -1
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/GutenbergIndex error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetVerbLemmaPos
  ' Goal:   Given a lemma of a verb, get its POS, which is VB, HV, BE or so
  ' History:
  ' 04-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function GetVerbLemmaPos(ByVal strLemma As String, ByVal strPos As String) As String
    Try
      Select Case strPos
        Case "VB"
          If (DoLike(strLemma, "hauen|have")) Then
            Return "HV"
          ElseIf (DoLike(strLemma, "be")) Then
            Return "BE"
          Else
            Return "VB"
          End If
        Case "VBP"
          If (DoLike(strLemma, "hauen|have")) Then
            Return "HVP"
          ElseIf (DoLike(strLemma, "be")) Then
            Return "BEP"
          Else
            Return "VBP"
          End If
        Case "VBD"
          If (DoLike(strLemma, "hauen|have")) Then
            Return "HVD"
          ElseIf (DoLike(strLemma, "be")) Then
            Return "BED"
          Else
            Return "VBD"
          End If
        Case "VBI"
          If (DoLike(strLemma, "hauen|have")) Then
            Return "HVI"
          ElseIf (DoLike(strLemma, "be")) Then
            Return "BEI"
          Else
            Return "VBI"
          End If
      End Select
      ' Default
      Return strPos
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/GetVerbLemmaPos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return strPos
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphDictWebsterToTdl
  ' Goal:   Convert the Gutenberg Webster .xml file into a MorphDict
  ' History:
  ' 29-11-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MorphDictWebsterToTdl() As Boolean
    Dim strText As String = ""        ' Text
    Dim strFeats As String = ""       ' Features
    Dim strLemma As String            ' Lemma
    Dim strPos As String              ' Part of speech
    Dim strThisPos As String          ' Local POS value
    Dim arPos(0) As String            ' POS divided into array
    Dim strDef As String              ' Definition
    Dim strPrevPos As String          ' Last POS
    Dim strPrevDef As String          ' Last DEF
    Dim strForm As String             ' List of forms
    Dim strVariant As String          ' Variant form
    Dim arVar() As String             ' Array of variants
    Dim strLine As String = ""        ' One line
    Dim strNew As String = ""         ' New line
    Dim strFileIn As String = "d:\data files\corpora\dictionaries\GutenbergWebsters_v2.xml"
    Dim strFileOut As String = ""     ' Output file for current letter
    Dim strDirIn As String = "D:\Data files\Corpora\Dictionaries\GCIDE\"
    Dim strDirOut As String = "d:\data files\corpora\dictionaries\"
    Dim arFile() As String            ' Array of input files
    Dim strFileSrt As String = ""     ' Sorted file
    Dim strLetter As String           ' Current letter
    Dim rdThis As IO.StreamReader = Nothing ' A text reader
    Dim xmlDoc As New XmlDocument     ' One piece of xml we work with
    Dim ndxSense As XmlNode           ' One sense
    Dim ndxMark As XmlNode            ' Node pointing to <mark>
    Dim ndxWork As XmlNode            ' Working node
    Dim ndxPrev As XmlNode            ' Preceding node
    Dim ndxList As XmlNodeList        ' List of nodes
    Dim ndxConj As XmlNodeList        ' List of conjugation forms
    Dim colSrt As New StringColl      ' Sorted version
    Dim colPos As New StringColl      ' All POS-es with an example
    Dim colForm As New StringColl     ' temporary collection
    Dim colHtml As New StringColl     ' report
    Dim dtrFound() As DataRow         ' Sorted version
    Dim bEntry As Boolean = False     ' New entry found
    Dim bPos As Boolean = False       ' POS determined
    Dim bDef As Boolean = False       ' Defintion
    Dim bComSt As Boolean = False     ' Comment start
    Dim bComEnd As Boolean = False    ' Comment end
    Dim intEntryId As Integer         ' Index within the [OEdict.xsd] type xml source
    Dim intMorphId As Integer         ' MorphId
    Dim intLexFunId As Integer        ' Lexical function
    Dim intSenseId As Integer         ' ID of Sense
    Dim intSize As Integer = 0        ' Number of lines
    Dim intDef As Integer             ' Definition number
    Dim intPtc As Integer             ' Percentage
    Dim intPos As Integer             ' Position
    Dim intI As Integer = 0           ' Counter
    Dim intJ As Integer               ' Counter
    Dim intK As Integer = 0           ' Counter
    Dim intP As Integer               ' POS counter
    Dim intW As Integer               ' Webster entry
    Dim intLetter As Integer          ' Current letter

    Try
      '' We need to create the [tdlMorphDict]
      'MorphDictIni(True, "MorphDictModE.xml")
      'If (tdlMorphDict Is Nothing) Then Status("Could not initialize tdlMorphdict") : Return False
      ' Read webster Cat > POS conversion
      If (Not InitWebsterPos()) Then Status("Could not read WebsterPos conversion") : Return False
      ' Read the Gutenberg XML dictionary piece by piece
      Status("Converting Webster XML into MorphDict...")
      ' Initialisations
      strPos = "" : strLemma = "" : intEntryId = 0 : intMorphId = 0 : intSenseId = 0 : intLexFunId = 0 : strVariant = ""
      ' Find the correct files in the GCIDE directory
      arFile = IO.Directory.GetFiles(strDirIn, "CIDE.?.txt", IO.SearchOption.TopDirectoryOnly)
      For intLetter = 0 To arFile.Length - 1
        ' Get this file
        strFileIn = arFile(intLetter) : intI = 0 : intSize = 0
        ' Get the letter
        If (IO.Path.GetExtension(strFileIn) = ".txt") Then
          intPos = InStrRev(strFileIn, ".") : strLetter = Left(strFileIn, intPos - 1)
          strLetter = Mid(IO.Path.GetExtension(strLetter), 2)
        Else
          Stop
          strLetter = ""
        End If
        ' Check: if the output of the *NEXT* letter already exists, we don't have to proceed
        strFileOut = IO.Path.GetDirectoryName(strFileIn) & "\Webster_" & Chr(Asc(strLetter) + 1) & ".xml"
        If (Not IO.File.Exists(strFileOut)) Then
          ' Determine output
          strFileOut = IO.Path.GetDirectoryName(strFileIn) & "\Webster_" & strLetter & ".xml"
          ' Initialise textreader
          rdThis = New IO.StreamReader(strFileIn) : intSize = 0
          ' Determine the size
          While (Not rdThis.EndOfStream)
            strLine = rdThis.ReadLine
            intSize += 1
          End While
          ' Clear output
          If (IO.File.Exists(strFileOut)) Then IO.File.Delete(strFileOut)
          ' Start output
          colSrt.Clear()
          ' Is this the first letter?
          If (intLetter = 0) Then
            ' Start the XML file
            colSrt.Add("<?xml version='1.0' standalone='yes'?>")
            colSrt.Add("<OEdict>") : colSrt.Add(" <EntryList>")
          End If
          ' Initialisations
          ndxSense = Nothing : strPos = "" : strDef = "" : strForm = "" : strVariant = "" : strPrevDef = "" : strPrevPos = "" : colForm.Clear()
          intDef = 0 : strFeats = ""
          ' Reset to the start
          rdThis.Close() : rdThis = New IO.StreamReader(strFileIn)
          While (Not rdThis.EndOfStream)
            ' Show where we are
            intPtc = (intI + 1) * 100 \ intSize
            Status("WebsterToTdl " & strLetter & " " & intPtc & "%", intPtc)
            ' Get next line
            strLine = rdThis.ReadLine
            ' =========== DEBUG ===============
            ' If (InStr(strLine, "<ent>Abaist</ent>") > 0) Then Stop
            ' =================================
            ' Check for comment start/end
            bComSt = (InStr(strLine, "<--") > 0) : bComEnd = (InStr(strLine, "-->") > 0)
            ' Add other lines, where needed
            While (Not rdThis.EndOfStream) AndAlso ((Right(strLine, 4) = "<br/") OrElse (bComSt And Not bComEnd))
              strNew = rdThis.ReadLine : intI += 1
              strLine &= strNew
              bComEnd = (InStr(strLine, "-->") > 0 OrElse InStr(strNew, "-->") = 1)
            End While
            ' Check if we can pass this through
            If (InStr(strLine, "<Xpage") = 0) AndAlso (DoLike(strLine, "*<ent>*|*<pos>*|*<def>*|*<vmorph>*|*<amorph>*|*<nmorph>*|*<wordforms>*|*<altsp>*")) Then
              ' =========== DEBUG ===============
              ' If (InStr(strLine, "<ent>Abandon</ent>") > 0) Then Stop
              ' =================================
              ' Do some initial work
              strText = "<webster>" & WebsterString(strLine) & "</webster>" : intI += 1
              ' NOTE:
              ' 1) Main entry:      used to be <h1>, but in Gcide it is <ent>
              ' 2) Part of speech:  used to be <tt>, but now it is <pos>

              ' What we have should be readable as XML
              xmlDoc.LoadXml(strText)

              ' (1) Look for <ent>
              ndxList = xmlDoc.SelectNodes("./descendant::ent")
              If (ndxList.Count > 0) Then
                ' Output the ***previous*** main entry
                If (strLemma <> "") AndAlso (strPos <> "") Then
                  ' Process webster pos
                  If (Not DoWebsterPos(strPos, arPos)) Then Return False
                  ' Check for empty POS
                  If (arPos.Length = 0) Then Stop
                  ' TODO: add words to <Morph> dictionary (see below)
                  colSrt.Add("  <Entry EntryId=" & """" & intEntryId & """" & _
                    " l=" & """" & strLemma & """" & _
                    " Pos=" & """" & arPos(0) & """" & _
                    IIf(strFeats = "", "", " f=" & """" & strFeats & """") & _
                    ">")
                  '' Add this particular POS
                  'colPos.AddUnique(strPos, strLemma)
                  ' Do we need to process more POS entries straight away?
                  For intP = 1 To arPos.Length - 1
                    ' Process this <LexFun>
                    WebsterAddLexFun(colForm, intEntryId, intLexFunId, arPos(intP), strLemma)
                  Next intP
                  ' Check for more variants
                  If (strVariant <> "") Then
                    arVar = Split(strVariant, ";")
                    For intK = 0 To arVar.Length - 1
                      For intP = 0 To arPos.Length - 1
                        WebsterAddLexFun(colForm, intEntryId, intLexFunId, arPos(intP), arVar(intK))
                      Next intP
                      ' ========== DEBUG ==========
                      If (arVar(intK) = "") Then Stop
                      ' ===========================
                    Next intK
                  End If
                  ' Add what we have gathered
                  'If (strDef <> "") Then colSrt.Add(strDef)
                  'If (strForm <> "") Then colSrt.Add(strForm)
                  If (colForm.Count > 0) Then
                    For intP = 0 To colForm.Count - 1
                      colSrt.Add(colForm.Item(intP))
                    Next intP
                  End If
                  ' Finish entry
                  colSrt.Add("  </Entry>")
                End If
                ' Debug.Print(colSrt.Item(colSrt.Count - 3))
                ' Reset parameters
                strLemma = LCase(ndxList(0).InnerText) : ndxSense = Nothing : strPos = "" : strDef = "" : strForm = "" : strVariant = ""
                strPrevPos = "" : strPrevDef = "" : colForm.Clear() : intDef = 0 : strFeats = ""
                ' Lemma adaptations...
                strLemma = XmlString(strLemma)
                intEntryId += 1
                ' Process the other <ent> entries as alternatives, lexical forms
                For intK = 1 To ndxList.Count - 1
                  ' TODO: Process this as a subentry
                  Debug.Print("subentry #" & intK & "=" & ndxList(intK).InnerText)
                  AddSemiStack(strVariant, ndxList(intK).InnerText)
                  ' Stop
                Next intK

              End If

              ' (2) Only Look for <pos> if this is a main entry...
              If (xmlDoc.SelectNodes("./descendant::ent").Count > 0) Then
                ndxList = xmlDoc.SelectNodes("./descendant::pos")
                If (ndxList.Count > 0) Then
                  ' The first POS is the main entry
                  strPos = ndxList(0).InnerText
                  If (strPos <> "") Then
                    bPos = True
                    ' Process webster pos
                    If (Not DoWebsterPos(strPos, arPos)) Then Return False
                    ' Check for empty POS
                    If (arPos.Length = 0) Then Stop
                    ' Should we process it already?
                    If (ndxList.Count > 0) AndAlso (ndxList(0).SelectSingleNode("./following-sibling::pos") IsNot Nothing) Then
                      ' Look for nearest following <def>
                      ndxSense = ndxList(0).SelectSingleNode("./following-sibling::def")
                      If (ndxSense Is Nothing) Then
                        ' Did we already have this?
                        If (strPrevPos <> strPos OrElse strPrevDef <> "-") Then
                          ' Just add a Sense
                          For intP = 0 To arPos.Length - 1
                            WebsterAddSense(colForm, intEntryId, intSenseId, arPos(intP), strLemma, intDef, "-")
                          Next intP
                          ' Adapt prev values
                          strPrevPos = strPos : strPrevDef = "-"
                        End If
                      Else
                        ' Did we already have this?
                        If (strPrevPos <> strPos OrElse strPrevDef <> ndxSense.InnerText) Then
                          ' Add a sense with a definition
                          For intP = 0 To arPos.Length - 1
                            WebsterAddSense(colForm, intEntryId, intSenseId, arPos(intP), strLemma, intDef, ndxSense.InnerText)
                          Next intP
                          ' Adapt prev values
                          strPrevPos = strPos : strPrevDef = ndxSense.InnerText
                        End If
                      End If
                    End If
                    ' A <mark> definition may follow
                    ndxMark = ndxList(0).SelectSingleNode("./following-sibling::mark")
                    If (ndxMark IsNot Nothing) Then
                      ' ======= Debug.
                      ' If (ndxMark.SelectSingleNode("./preceding-sibling::sn") IsNot Nothing) Then Stop
                      ' ==================
                      ' Check if it contains the obsolete marking
                      ' Also: it may not be part of a <sn> separate sense
                      If (DoLike(ndxMark.InnerText, "[[]Obs.]")) AndAlso (ndxMark.SelectSingleNode("./preceding-sibling::sn") Is Nothing) Then
                        ' Mark feature obsolete
                        AddSemiStack(strFeats, "mark=obs", True, "|")
                      End If
                    End If
                    ' There may be other POS definitions, but these must be part of <vmorph>, <amorph> or <nmorph>
                    For intK = 1 To ndxList.Count - 1
                      Select Case ndxList(intK).ParentNode.Name
                        Case "vmorph", "amorph", "nmorph", "plu"
                          ' One or more <conjf>, <adjf> or <decf> forms should follow
                          ndxConj = ndxList(intK).SelectNodes("./following-sibling::conjf | ./following-sibling::pos")
                          For intJ = 0 To ndxConj.Count - 1
                            ' What should we do?
                            Select Case ndxConj(intJ).Name
                              Case "pos"
                                ' Leave
                                Exit For
                              Case "conjf", "adjf", "decf", "plw"
                                If (DoLike(ndxList(intK).ParentNode.Name, "[anv]morph|plu")) Then
                                  ' Get the POS
                                  ndxWork = ndxConj(intJ).SelectSingleNode("./preceding::pos[1]")
                                  If (ndxWork Is Nothing) Then
                                    Stop
                                  Else
                                    strThisPos = ndxWork.InnerText
                                    ' Process webster pos
                                    If (Not DoWebsterPos(strThisPos, arPos)) Then Return False
                                    ' Check for empty POS
                                    If (arPos.Length = 0) Then Stop
                                    ' Process POS + conjugated form...
                                    For intP = 0 To arPos.Length - 1
                                      WebsterAddLexFun(colForm, intEntryId, intLexFunId, arPos(intP), ndxConj(intJ).InnerText)
                                    Next intP
                                    ' ========== DEBUG ==========
                                    If (ndxConj(intJ).InnerText = "") Then Stop
                                    ' ===========================
                                  End If
                                End If
                            End Select

                          Next intJ
                        Case "webster"
                          ' Check if this is preceded by "see"
                          If (Not DoLike(LCase(ndxList(intK).PreviousSibling.InnerText), "*see*|*from*")) Then
                            ' No, this is not going to work
                            If (False) Then
                              ' This is a main entry --> it must be of a different sense
                              strPos = ndxList(intK).InnerText
                              ' Process webster pos
                              If (Not DoWebsterPos(strPos, arPos)) Then Return False
                              ' Check for empty POS
                              If (arPos.Length = 0) Then Stop
                              ' Look for nearest following <def>
                              ndxSense = ndxList(intK).SelectSingleNode("./following-sibling::def")
                              If (ndxSense Is Nothing) Then
                                ' Did we already have this?
                                If (strPrevPos <> strPos OrElse strPrevDef <> "-") Then
                                  ' Just add a Sense
                                  For intP = 0 To arPos.Length - 1
                                    WebsterAddSense(colForm, intEntryId, intSenseId, arPos(intP), strLemma, intDef, "-")
                                  Next intP
                                  ' Adapt prev values
                                  strPrevPos = strPos : strPrevDef = "-"
                                End If
                              Else
                                ' Did we already have this?
                                If (strPrevPos <> strPos OrElse strPrevDef <> ndxSense.InnerText) Then
                                  ' Add a sense with a definition
                                  For intP = 0 To arPos.Length - 1
                                    WebsterAddSense(colForm, intEntryId, intSenseId, arPos(intP), strLemma, intDef, ndxSense.InnerText)
                                  Next intP
                                  ' Adapt prev values
                                  strPrevPos = strPos : strPrevDef = ndxSense.InnerText
                                End If
                              End If

                            End If
                            ' Main entry POS may be followed by <plu> with a <plw> -- check
                            ndxWork = ndxList(intK).SelectSingleNode("./following-sibling::plu")
                            If (ndxWork IsNot Nothing) Then
                              ' Look for the *first* <plw> form
                              ndxWork = ndxWork.SelectSingleNode("./child::plw[1]")
                              If (ndxWork IsNot Nothing) Then
                                ' Add a lexical function entry by substantiating the main POS with feature [number=pl]
                                For intP = 0 To arPos.Length - 1
                                  WebsterAddLexFun(colForm, intEntryId, intLexFunId, arPos(intP) & ";number=pl", ndxWork.InnerText)
                                Next intP
                              End If
                            End If
                          End If
                        Case "wordforms", "def", "see", "altsp", "plu"
                          ' Okay, this is treated later on, or <def> parent doesn't matter
                        Case "mark"
                          ' Don't treat this separately here
                        Case Else
                          Stop
                      End Select
                    Next intK
                  End If  ' If (strPos <> "") Then
                End If    ' If (ndxList.Count > 0) Then

              End If      ' If (xmlDoc.SelectNodes("./descendant::ent").Count > 0) Then

              ' (3) Look for <def>
              If (ndxSense Is Nothing) Then
                ndxList = xmlDoc.SelectNodes("./descendant::def")
              Else
                ' Start looking from <ndxSense>
                ndxList = ndxSense.SelectNodes("./following::def")
              End If
              For intK = 0 To ndxList.Count - 1
                ' Process this defintion
                ndxSense = ndxList(intK)
                ' Did we already have this?
                If (strPrevPos <> strPos OrElse strPrevDef <> ndxSense.InnerText) AndAlso (strPos <> "") Then
                  ' Process webster pos
                  If (Not DoWebsterPos(strPos, arPos)) Then Return False
                  ' Check for empty POS
                  If (arPos.Length = 0) Then Stop
                  ' Add a sense with a definition
                  For intP = 0 To arPos.Length - 1
                    WebsterAddSense(colForm, intEntryId, intSenseId, arPos(intP), strLemma, intDef, ndxSense.InnerText)
                  Next intP
                  ' Adapt prev values
                  strPrevPos = strPos : strPrevDef = ndxSense.InnerText
                End If

              Next intK

              ' (4) word forms
              ' Check if this contains word forms
              ndxWork = xmlDoc.SelectSingleNode("./descendant::wordforms")
              If (ndxWork IsNot Nothing) Then
                ' Look for all instances of <er> and preceding <tt>
                ndxList = ndxWork.SelectNodes("./child::er")
                ' Walk all <er> children of <wordforms>
                For intJ = 0 To ndxList.Count - 1
                  ' Get the preceding <tt> node
                  ndxPrev = ndxList(intJ).SelectSingleNode("./preceding-sibling::pos[1]")
                  If (ndxPrev IsNot Nothing) Then
                    ' We have a matching word form + notation
                    WebsterAddLexFun(colForm, intEntryId, intLexFunId, ndxPrev.InnerText, ndxList(intJ).InnerText)
                    ' ========== DEBUG ==========
                    If (ndxList(intJ).InnerText = "") Then Stop
                    ' ===========================
                  End If
                Next intJ
              End If

              ' (5) Check if we have alternative spelling forms
              ndxWork = xmlDoc.SelectSingleNode("./descendant::altsp")
              If (ndxWork IsNot Nothing) Then
                ' Look for all instances of <asp> 
                ndxList = ndxWork.SelectNodes("./child::asp")
                ' Walk all <asp> children of <wordforms>
                For intJ = 0 To ndxList.Count - 1
                  ' strForm &= "<form id='" & intJ + 1 & "' vern='" & ndxList(intJ).InnerText & "' />" & vbCrLf
                  AddSemiStack(strVariant, ndxList(intJ).InnerText)
                Next intJ
              End If
            End If
          End While
          ' Is this the last letter?
          If (intLetter = arFile.Length - 1) Then
            ' Add final part for this section
            colSrt.Add(" </EntryList>") : colSrt.Add("</OEdict>")
          End If
          ' Save this section
          IO.File.WriteAllText(strFileOut, colSrt.Text, System.Text.Encoding.UTF8)
        End If
      Next intLetter
      ' Combine all files into one
      Status("Combining letters into one file...")
      ' Determine output file
      strFileOut = strDirOut & "ModEdict_out" & ".xml"
      ' Delete older version
      If (IO.File.Exists(strFileOut)) Then IO.File.Delete(strFileOut)
      ' Walk all input files
      For intLetter = 0 To arFile.Length - 1
        ' Show where we are
        intPtc = (intLetter + 1) * 100 \ arFile.Length
        Status("Combining " & intPtc & "%", intPtc)
        ' Get this file
        strFileIn = strDirIn & "Webster_" & Chr(Asc("A") + intLetter) & ".xml"
        ' strFileIn = arFile(intLetter)
        ' Append input file to the output
        IO.File.AppendAllText(strFileOut, IO.File.ReadAllText(strFileIn))
      Next intLetter

      'Status("Sorting...")
      ' '' Save the resulting morphdict
      ''tdlMorphDict.WriteXml(strMorphDictFile)
      '' Make a sorted version
      'dtrFound = tdlMorphDict.Tables("Morph").Select("", "Vern ASC, l ASC")
      ' '' Save this sorted version under a different name
      ''strFileSrt = GetSetDir() & "\MorphDict-tmp.xml"
      'colSrt.Clear() : colSrt.Add("<MorphDict>") : colSrt.Add(" <MorphList>")
      'For intI = 0 To dtrFound.Length - 1
      '  ' Show where we are
      '  intPtc = (intI + 1) * 100 \ dtrFound.Length
      '  Status("Sorting " & intPtc & "%", intPtc)
      '  ' Process this entry
      '  With dtrFound(intI)
      '    colSrt.Add("  <Morph MorphId=" & """" & intI + 1 & """" & " EntryId=" & """" & .Item("EntryId").ToString & """" & _
      '               " Vern=" & """" & .Item("Vern").ToString & """" & _
      '               " l=" & """" & .Item("l").ToString & """" & _
      '               " Pos=" & """" & .Item("Pos").ToString & """" & _
      '               " Label=" & """" & .Item("Label").ToString & """" & _
      '               " f=" & """" & .Item("f").ToString & """" & _
      '               " />")
      '  End With
      'Next intI
      'colSrt.Add(" </MorphList>") : colSrt.Add("</MorphDict>")
      'IO.File.WriteAllText(strMorphDictFile, colSrt.Text)
      '' Show we are ready
      'Logging("Dictionary saved at: " & strMorphDictFile)
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MorphDictWebsterToTdl error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   WebsterAddLexFun
  ' Goal:   Convert a string in the Webster xml/html dictionary to readable text
  ' History:
  ' 29-11-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function WebsterAddLexFun(ByRef colForm As StringColl, ByVal intEntryId As Integer, ByRef intLexFunId As Integer, ByVal strPos As String, _
                                     ByVal strVern As String) As Boolean
    Try
      ' make a new LexFunId
      intLexFunId += 1
      ' Adapt strVern
      strVern = XmlString(strVern)
      ' Process the entry
      colForm.Add("   <LexFun EntryId=" & """" & intEntryId & """" & " LexFunId=" & """" & intLexFunId & """" & _
           " lf=" & """" & strPos & """" & _
           " lv=" & """" & LCase(strVern) & """" & " />")
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/WebsterAddLexFun error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   WebsterAddSense
  ' Goal:   Convert a string in the Webster xml/html dictionary to readable text
  ' History:
  ' 29-11-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function WebsterAddSense(ByRef colForm As StringColl, ByVal intEntryId As Integer, ByRef intSenseId As Integer, ByVal strPos As String, _
                                     ByVal strLemma As String, ByVal intDef As Integer, ByVal strDef As String) As Boolean
    Try
      ' make a new LexFunId
      intSenseId += 1 : intDef += 1
      ' Adapt <strDef>
      strDef = XmlString(strDef)
      ' Process the entry
      colForm.Add("   <Sense EntryId=" & """" & intEntryId & """" & " SenseId=" & """" & intSenseId & """" & _
                            " l=" & """" & strLemma & """" & " N=" & """" & intDef & """" & _
                            " Pos=" & """" & strPos & """" & " Def=" & """" & strDef & """" & _
                                     " />")
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/WebsterAddSense error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   WebsterString
  ' Goal:   Convert a string in the Webster xml/html dictionary to readable text
  ' History:
  ' 29-11-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function WebsterString(ByVal strIn As String) As String
    Dim strTag As String = ""         ' Tag itself
    Dim strReg As String              ' Regular expression
    Dim colStart As MatchCollection   ' Collection of opening tags
    Dim colEnd As MatchCollection     ' Collection of closing tags
    Dim colTags As MatchCollection    ' Collection of opening and closing tags
    Dim mtThis As Match               ' One match
    Dim intPos As Integer             ' Position within string
    Dim intLen As Integer             ' Length
    Dim intI As Integer               ' Counter
    Dim intJ As Integer               ' Counter
    Dim bProblem As Boolean = False   ' There is a problem
    ' Review and possibly repair these tags:
    Dim arTag() As String = {"def", "tt", "hw", "ent", "fld", "plw", "plu", "wordforms", "i", "b", "altsp", "asp"}
    ' These tags need to be deleted and not taken into account
    Dim arDel() As String = {"mhw", "blockquote", "ety", "ets", "as", "note", "er", "wf", "p", "q", "col", _
                             "cd", "cs", "def2", "mcol", "spn", "spm", "ex", "syn", "chform", _
                             "it", "fld", "altname", "mord", "sd", "dsyn", "sing", "singw", "usage", "pre", _
                             "subtypes", "text", "song", "caption", "table", "bio"}
    ' Translation from names to symbols
    Dim arFieldIn() As String = {"&flat;", "&nbsp;", "&ndash;", "&mdash;", "&alpha;", "&hand;", _
                                  "&a;", "&keras;", "&or;", "&deg;", "& ", "&.", "&,", "<?/", "</[", _
                                  "&min;", "&acr;", "&icr;", "&omac;", "&emac;", "&osl;", "&oe;", "<--", " -- ", _
                                  "</rdquo/", "<2dot/", "<3dot/", "<4dot/", "<8star/"}
    Dim arFieldOut() As String = {"b", " ", "-", "-", "a", "h", _
                                  "a", "keras", "or", "degree", "&amp; ", "&amp;.", "&amp;,", "#", "", _
                                  "min", "acr", "icr", "omac", "emac", "osl", "", "<!--", " - ", _
                                  "<rdquo/>", "", "", "", ""}
    ' ==================== OLD CODE ================================
    Dim arMatchSt() As String = {"\<def\>", "\<tt\>", "\<hw\>"}
    Dim arMatchEn() As String = {"\<\/def\>", "\<\/tt\>", "\<\/hw\>"}
    ' ===============================================================

    Try
      ' ======== DEBUG =======
      ' If (InStr(strIn, ">Homer") > 0) Then Stop
      ' ======================
      ' Change symbols into their translation
      For intI = 0 To arFieldIn.Length - 1
        strIn = strIn.Replace(arFieldIn(intI), arFieldOut(intI))
      Next intI
      ' Remove word-final < sign
      If (Right(strIn, 1) = "<") Then
        strIn = Left(strIn, strIn.Length - 1)
      ElseIf (Right(strIn, 5) = "</def") OrElse (Right(strIn, 5) = "</syn") OrElse (Right(strIn, 5) = "</ety") Then
        strIn &= ">"
      End If
      ' Look for 'sgml-style' single stuff: <mmm/ --> <mmm/>
      colTags = Regex.Matches(strIn, "\<\w+\/[^\>]")
      While (colTags.Count > 0)
        For intI = colTags.Count - 1 To 0 Step -1
          ' Insert the > sign after the closing / sign
          mtThis = colTags(intI)
          intPos = mtThis.Index + mtThis.Length
          strIn = Left(strIn, intPos - 1) & ">" & Mid(strIn, intPos)
          'strIn = Left(strIn, colTags(intI).Index) & Left(colTags(intI).Value, colTags(intI).Value.Length - 1) & ">" & Mid(strIn, colTags(intI).Index + colTags(intI).Length)
        Next intI
        ' Try once more in the left-over
        colTags = Regex.Matches(strIn, "\<\w+\/[^\>]")
      End While

      ' Look for 'sgml-style' single stuff: <mmm/ --> <mmm/>
      colTags = Regex.Matches(strIn, "\<\w+\/$")
      ' colTags = Regex.Matches(strIn, "\<\w+\/($!\>)")
      For intI = colTags.Count - 1 To 0 Step -1
        ' Insert the > sign after the closing / sign
        strIn &= ">"
        ' strIn = Left(strIn, colTags(intI).Index) & colTags(intI).Value & ">" & Mid(strIn, colTags(intI).Index + colTags(intI).Length + 1)
      Next intI

      '' Change singular ampersands
      'strIn = strIn.Replace("& ", "&amp; ")
      '' Remove weird symbol
      'strIn = strIn.Replace("<?/", "#")
      ' Delete all tags that need to be deleted
      For intI = 0 To arDel.Length - 1
        ' Find and delete tag type one
        strTag = "<" & arDel(intI) & ">"
        If (InStr(strIn, strTag) > 0) Then
          strIn = strIn.Replace(strTag, "")
        End If
        strTag = "</" & arDel(intI) & ">"
        If (InStr(strIn, strTag) > 0) Then
          strIn = strIn.Replace(strTag, "")
        End If
      Next intI
      ' Look for any left-over non-declared entities
      colTags = Regex.Matches(strIn, "\&\w+\;")
      For intI = colTags.Count - 1 To 0 Step -1
        ' Check which one this is
        Select Case colTags(intI).Value
          Case "&amp;"
            ' Continue
          Case Else
            ' Replace & with #sign
            strIn = Left(strIn, colTags(intI).Index) & "#" & Mid(strIn, colTags(intI).Index + 2)
        End Select
      Next intI
      ' Now look specifically for ampersand replacement
      colTags = Regex.Matches(strIn, "\&")
      For intI = colTags.Count - 1 To 0 Step -1
        ' Check which one this is
        If (Mid(strIn, colTags(intI).Index + 1, 5) <> "&amp;") Then
          ' Replace it
          strIn = Left(strIn, colTags(intI).Index) & "&amp;" & Mid(strIn, colTags(intI).Index + 2)
        End If
      Next intI
      '' Look for the weird use of <name/  without an ending 
      'colTags = Regex.Matches(strIn, "\<\w+\/[^;]")
      'For intI = colTags.Count - 1 To 0 Step -1
      '  ' Replace without starting < and closing /
      '  strTag = Mid(colTags(intI).Value, 2)
      '  strTag = Left(strTag, strTag.Length - 2) & Right(colTags(intI).Value, 1)
      '  strIn = Left(strIn, colTags(intI).Index) & strTag & Mid(strIn, colTags(intI).Index + colTags(intI).Length + 1)
      'Next intI

      ' We do it differently...
      For intI = 0 To arTag.Length - 1
        ' Get the tag for this occurrance
        strTag = arTag(intI) : strReg = "\<\/?" & strTag & "\>"
        Do
          ' Make a list of matches, which includes opening and closing ones
          colTags = Regex.Matches(strIn, strReg)
          ' Initially there is no problem
          bProblem = False
          ' Walk all the matches one by one
          For intJ = 0 To colTags.Count - 1
            ' Note position *after* the xml tag
            intPos = colTags(intJ).Index + 1 + colTags(intJ).Length
            ' we are expecting an opening tag
            If (InStr(colTags(intJ).Value, "/") = 0) Then
              ' Okay, we have an opening tag -- now get the closing tag
              intJ += 1
              If (intJ >= colTags.Count) OrElse (InStr(colTags(intJ).Value, "/") = 0) Then
                ' Either:
                ' (a) There are no more matches left, or 
                ' (b) The next tag unexpectedly is an opening tag too
                If (intJ = colTags.Count - 1) Then
                  ' Change the opening tag by a closing tag
                  strIn = Left(strIn, colTags(intJ).Index) & "</" & strTag & ">" & Mid(strIn, colTags(intJ).Index + colTags(intJ).Length + 1)
                ElseIf (intJ <= colTags.Count - 2) AndAlso (InStr(colTags(intJ + 1).Value, "/") = 0) Then
                  ' Another *opening* tag follows the unexpected opening tag
                  ' Change the current opening tag into a closing tag
                  strIn = Left(strIn, colTags(intJ).Index) & "</" & strTag & ">" & Mid(strIn, colTags(intJ).Index + colTags(intJ).Length + 1)
                Else
                  ' Insert a closing tag before the first < after the xml tag
                  intPos = InStr(intPos, strIn, "<")
                  If (intPos = 0) Then intPos = strIn.Length + 1
                  strIn = Left(strIn, intPos - 1) & "</" & strTag & ">" & Mid(strIn, intPos)
                End If
                ' Indicate there has been a problem
                bProblem = True : Exit For
              End If
            Else
              ' This is a closing tag, and we are *not* expecting it
              ' Is the next one also a closing tag?
              If (intJ < colTags.Count - 1) AndAlso (InStr(colTags(intJ + 1).Value, "/") > 0) Then
                ' Change the current one into a starting tag
                strIn = Left(strIn, colTags(intJ).Index) & "<" & strTag & ">" & Mid(strIn, colTags(intJ).Index + colTags(intJ).Length + 1)
              Else
                ' Look for a position where we can insert a starting tag
                If (colTags(intJ).Index = 0) Then
                  intPos = 0
                Else
                  intPos = InStrRev(strIn, ">", colTags(intJ).Index)
                  If (intPos = 0) Then intPos = 0
                End If
                strIn = Left(strIn, intPos) & "<" & strTag & ">" & Mid(strIn, intPos + 1)
              End If
              ' Indicate there has been a problem
              bProblem = True : Exit For
            End If
          Next intJ
        Loop While (bProblem)
      Next intI


      ' ==================== OLD CODE ================================
      If (False) Then
        ' ================ GETS NEVER ACCESSED ===============
        ' Check for matching beginning and end codes
        For intI = 0 To arMatchSt.Length - 1
          colStart = Regex.Matches(strIn, arMatchSt(intI))
          colEnd = Regex.Matches(strIn, arMatchEn(intI))
          ' Do the numbers match?
          If (colStart.Count > colEnd.Count) Then
            ' There are more starters than ends
            Select Case colStart.Count - colEnd.Count
              Case 1
                ' Delete the last starter??
                Stop
              Case 2
                ' There are more starts than finishes -- change one start into a finish
                With colStart.Item(colStart.Count - 1)
                  intPos = .Index + 1
                  intLen = .Length
                End With
                strIn = Left(strIn, intPos - 1) & "</" & arTag(intI) & ">" & Mid(strIn, intPos + intLen + 1)
              Case Else
                Stop
            End Select
          ElseIf (colStart.Count < colEnd.Count) Then
            Stop
            ' There are more ends than starts -- change one start into a finish
            With colStart.Item(colStart.Count - 1)
              intPos = .Index + 1
              intLen = .Length
            End With
            strIn = Left(strIn, intPos - 1) & "</" & arTag(intI) & ">" & Mid(strIn, intPos + intLen + 1)
          End If
        Next intI
      End If
      ' ===============================================================
      ' Return the result
      Return strIn
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/WebsterString error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   CombineMEDsections
  ' Goal:   Convert the Gutenberg .xml file into a MorphDict
  ' History:
  ' 21-01-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function CombineMEDsections() As Boolean
    Dim arFile() As String      ' Array of input files...
    Dim arPos() As String       ' Array of POS tags
    Dim strLemma As String      ' Lemma
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter
    Dim intK As Integer         ' Counter
    Dim intPtc As Integer       ' Percentage
    Dim bIsName As Boolean      ' Is this a name?
    Dim tdlSection As DataSet   ' One section of the dictionary
    Dim dtrThis As DataRow      ' One row
    Dim strDirIn As String = "\Data Files\Research\2013\V2_ME\Data"
    Dim strFile As String = strDirIn & "\MED_combined.xml"

    Try
      ' Do we need to do this again?
      If (IO.File.Exists("U:" & strFile)) Then
        Select Case MsgBox("The file [" & strFile & "] already exists." & vbCrLf & _
                           "Would you like to over-write it?", MsgBoxStyle.YesNoCancel)
          Case MsgBoxResult.No
            Return True
          Case MsgBoxResult.Cancel
            Return False
        End Select
      End If
      ' Create a new dictionary that will contain all the results
      If (Not CreateDataSet("OEdict.xsd", tdlOEdict)) Then Return False
      ' Initialise the MED pos dictionary
      If (Not InitMEDpos()) Then Return False
      ' Go through all sections
      arFile = IO.Directory.GetFiles(strDirIn, "*_out.txt")
      For intI = 0 To arFile.Length - 1
        ' Where are we?
        intPtc = (intI + 1) * 100 \ arFile.Length
        Status("Combining MED " & intPtc & "%", intPtc)
        ' Read this file
        tdlSection = Nothing
        If (Not ReadDataset("OEdict.xsd", arFile(intI), tdlSection)) Then Return False
        ' Go through this section, transferring to [tdlOEdict]
        With tdlSection.Tables("Entry")
          For intJ = 0 To .Rows.Count - 1
            ' Get the lemma
            strLemma = .Rows(intJ).Item("l").ToString
            ' Remove brackets from lemma
            strLemma = strLemma.Replace("(", "") : strLemma = strLemma.Replace(")", "")
            ' Check if the first letter of the lemma is upper-case (then it is probably NPR)
            bIsName = (AscW(Left(strLemma, 1)) <> AscW(LCase(Left(strLemma, 1))))
            ' ======================================
            ' If (strLemma = "al-kinnes") Then Stop
            ' ======================================
            ' Get all POS variants
            arPos = GetMEDpos(.Rows(intJ).Item("Pos").ToString, bIsName, "-")
            ' Walk all POS tags
            For intK = 0 To arPos.Length - 1
              ' Is this something?
              If (arPos(intK) <> "") Then
                ' Make a new row for the [tdlOEdict]
                dtrThis = AddOneDataRow(tdlOEdict, "Entry", "EntryId", "EntryList")
                dtrThis.Item("l") = strLemma
                dtrThis.Item("Pos") = arPos(intK)
                ' Provide the link to the MED
                dtrThis.Item("s") = .Rows(intJ).Item("med").ToString
              End If
            Next intK
          Next intJ
        End With
      Next intI
      '' Show any remaining unresolved issues
      'For intI = 0 To colMEDpos.Count - 1
      '  Debug.Print(intI & vbTab & colMEDpos.Item(intI) & vbTab & colMEDpos.Freq(intI))
      'Next intI
      ' Save the combined dictionary
      ' strFile = strDirIn & "\MED_combined.xml"
      tdlOEdict.WriteXml(strFile)
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/CombineMEDsections error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   InitMEDpos
  ' Goal:   Initialise reading MED pos tags
  ' History:
  ' 21-01-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function InitMEDpos() As Boolean
    Dim strFile As String   ' File to be read
    Dim arText() As String  ' Contents of file
    Dim arSplit() As String ' One line
    Dim intI As Integer     ' Counter

    Try
      ' Initialise collection
      colMEDpos.Clear()
      ' Get file name
      strFile = IO.Path.GetDirectoryName(Application.ExecutablePath) & "\MED_pos.txt"
      If (Not IO.File.Exists(strFile)) Then Return False
      ' Read content
      arText = IO.File.ReadAllLines(strFile)
      ' Put into collection
      For intI = 0 To arText.Length - 1
        ' Any content?
        If (arText(intI) <> "") Then
          arSplit = Split(arText(intI), vbTab)
          If (arSplit.Length = 3) Then
            colMEDpos.Add(arSplit(1), arSplit(0) & ";" & arSplit(2))
          End If
        End If
      Next intI
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/InitMEDpos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetMEDpos
  ' Goal:   Convert MED POS into 'normal' POS
  ' History:
  ' 21-01-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function GetMEDpos(ByVal strIn As String, ByVal bIsName As Boolean, ByVal strPos As String) As String()
    Dim intPos As Integer   ' Position within string
    Dim intI As Integer     ' Counter
    Dim intJ As Integer     ' Counter
    Dim strType As String   ' Subtype
    Dim strReqPos As String ' Required POS
    Dim strOutPos As String ' Outcome POS
    Dim bFound As Boolean   ' Found flag
    Dim arBack() As String = Nothing
    Dim arItem() As String
    Dim colBack As New StringColl

    Try
      ' 
      ' Trim it
      strType = "" : strIn = strIn.Replace("?", "") : strIn = Trim(strIn)
      ' Remove FINAL comma
      If (DoLike(strIn, "*,|*;")) Then strIn = Trim(Left(strIn, strIn.Length - 1))
      ' Remove unnecessary spaces
      strIn = strIn.Replace(". ", ".")
      strIn = strIn.Replace(", ", ",")
      ' Change to lower case
      strIn = strIn.ToLower
      ' ====== DEBUG ========================
      ' If (InStr(strIn, "ppl.") > 0) AndAlso (strPos = "v.") Then Stop
      ' =====================================
      ' ================== DEBUG =====================
      If (InStr(strIn, "also") > 0 OrElse InStr(strIn, "contract") > 0) Then Stop
      ' ==============================================

      ' Delete everything starting with (
      intPos = InStr(strIn, "(")
      If (intPos > 0) Then
        strType = Trim(Mid(strIn, intPos + 1))
        strIn = Trim(Left(strIn, intPos - 1))
        ' Remove right bracket
        intPos = InStr(strType, ")")
        If (intPos > 0) Then
          strType = Trim(Left(strType, intPos - 1))
        End If
      End If
      ' Check for ampersand or "or"
      intPos = InStr(strIn, "&")
      If (intPos > 0) Then
        ' Two entries
        ReDim arBack(0 To 1)
        arBack(0) = Trim(Left(strIn, intPos - 1))
        arBack(1) = Trim(Mid(strIn, intPos + 1))
      Else
        ' Check for [or]
        intPos = InStr(strIn, " or ")
        If (intPos > 0) Then
          ' Two entries
          ReDim arBack(0 To 1)
          arBack(0) = Trim(Left(strIn, intPos - 1))
          arBack(1) = Trim(Mid(strIn, intPos + 4))
        ElseIf (DoLike(strIn, "*as subj.|*as obj.")) Then
          ' Just one entry
          ReDim arBack(0 To 0)
          arBack(0) = ""
        Else
          ' Check for [as]
          intPos = InStr(strIn, " as ")
          If (intPos > 0) Then
            ' Just one entry
            ReDim arBack(0 To 0)
            arBack(0) = Trim(Mid(strIn, intPos + 4))
          Else
            ' Just one entry
            ReDim arBack(0 To 0)
            arBack(0) = strIn
          End If
        End If
      End If
      ' Treat all entries
      For intI = 0 To arBack.Length - 1
        ' Get this entry
        strIn = Trim(arBack(intI))
        If (strIn = "") Then
          bFound = True
        Else
          ' Go and find it...
          bFound = False
          For intJ = 0 To colMEDpos.Count - 1
            ' Get required and outcome POS
            strReqPos = Split(colMEDpos.Exmp(intJ), ";")(0)
            ' =========== DEBUG ==============
            ' If (colMEDpos.Item(intJ) = strIn) Then Stop
            ' ================================
            ' Check ...
            If (colMEDpos.Item(intJ) = strIn) AndAlso (strReqPos = "" OrElse strReqPos = strPos) Then
              arBack(intI) = Split(colMEDpos.Exmp(intJ), ";")(1) : bFound = True
              ' Check for sub-types of verbs
              If (arBack(intI) = "VB") AndAlso (strType <> "") Then
                Select Case strType
                  Case "sg. 3", "sg.3", "sg. 2", "sg.3", "pr.sg.2", "pr.sg.3"
                    arBack(intI) = "VBP"
                  Case "p. t.", "p.t.", "p."
                    arBack(intI) = "VBD"
                  Case "impv.sg."
                    arBack(intI) = "VBI"
                  Case "inf."
                    ' No changes needed!
                End Select
              End If
              Exit For
            End If
          Next intJ
        End If
        ' Double check
        If (Not bFound) Then
          ' Nothing was found
          Debug.Print("Did not find POS entry: [" & strIn & "]")
          ' Add it to the missing POS bag
          colMis.AddUnique(LCase(strIn), strPos)
          ' Return our input instead
          arBack(intI) = strIn
          '' Stop
          'Return Nothing
        End If
      Next intI
      ' Enter everything into collection
      For intI = 0 To arBack.Length - 1
        If (arBack(intI) <> "") Then
          arItem = Split(arBack(intI), ";")
          For intJ = 0 To arItem.Count - 1
            If (arItem(intJ) <> "") Then
              colBack.AddUnique(Trim(arItem(intJ)))
            End If
          Next intJ
        End If
      Next intI
      ' Return the result
      ReDim arBack(0 To colBack.Count - 1)
      For intI = 0 To colBack.Count - 1
        If (bIsName) Then
          If (colBack.Item(intI) = "N") Then
            arBack(intI) = "NPR"
          ElseIf (colBack.Item(intI) = "NS") Then
            arBack(intI) = "NPRS"
          Else
            arBack(intI) = colBack.Item(intI)
          End If
        Else
          arBack(intI) = colBack.Item(intI)
        End If
      Next intI
      ' If (arBack.Length > 1) Then Stop
      Return arBack
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/GetMEDpos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return arBack
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetLastMedEntry
  ' Goal:   Find the last entry id
  ' History:
  ' 21-04-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function GetLastMedEntry(ByVal strFile As String) As Integer
    Dim strLine As String               ' One line
    Dim strText As String = ""          ' Last line that contains an id
    Dim strDiv As String = "<div id='"  ' Recognition string
    Dim intPos As Integer               ' Position in string
    Dim intId As Integer
    Dim rdThis As IO.StreamReader = Nothing

    Try
      ' Validate
      If (Not IO.File.Exists(strFile)) Then Return 0
      ' Open file
      rdThis = New IO.StreamReader(strFile)
      ' Loop through all lines
      While (Not rdThis.EndOfStream)
        strLine = rdThis.ReadLine
        If (InStr(strLine, strDiv) > 0) Then strText = strLine
      End While
      ' Close the file
      rdThis.Close()
      ' Check the last line
      If (InStr(strText, strDiv) > 0) Then
        ' Get the id
        strText = Mid(strText, strDiv.Length + 1)
        ' Find second quotation mark
        intPos = InStr(strText, "'")
        If (intPos > 0) Then
          ' Get the text
          strText = Left(strText, intPos - 1)
          If (IsNumeric(strText)) Then
            intId = CInt(strText)
            Return intId
          End If
        End If
      End If
      ' Return failure
      Return 0
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/GetLastMedEntry error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return 0
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MEDextend_Org
  ' Goal:   Extend the lemma-only MED-Lookup dictionary with:
  '         - variant forms
  '         - derived morphemes
  '         - (definitions) -- (this is perhaps not necessary?)
  ' History:
  ' 15-04-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MEDextend_Org(ByRef wbThis As WebBrowser) As Boolean
    Dim strMed As String        ' MED code
    Dim strWpage As String      ' The web-page
    Dim strPre As String        ' Pre-amble for web page
    Dim strTmpF As String       ' Temporary file
    Dim strText As String       ' Text
    Dim strLemma As String      ' The lemma we are looking at
    Dim strPos As String        ' The POS of the lemma
    Dim strEntryPos As String   ' POS of the entry as a whole
    Dim strAlt As String        ' Alternative form
    Dim strAlt2 As String       ' Alternative for alternative
    Dim strTitle As String = "" ' HTML title
    Dim arPosList() As String   ' List of POS possibilities
    Dim bIsName As Boolean      ' Is the entry a name?
    Dim bOkay As Boolean        ' Are we okay?
    Dim ndxEntry As XmlNode     ' One <Entry> record
    Dim ndxSense As XmlNode     ' One <Sense> node
    Dim ndxLexFun As XmlNode    ' One <LexFun> node
    Dim ndxHtml As XmlNode      ' One node from the <html> entry part
    Dim ndxImg As XmlNode       ' One <img> node
    Dim ndxChild As XmlNode
    Dim ndxList As XmlNodeList  ' List of nodes
    Dim ndxInfo As XmlNodeList  ' Information about an entry
    Dim colAlt As New StringColl  ' Collection of alternatives
    Dim colDef As New StringColl  ' Collection of definitions
    Dim intPos As Integer       ' Position in string
    Dim intEntryId As Integer   ' ID of this entyr
    Dim intLexFunId As Integer  ' ID of <LexFun> element
    Dim intSenseId As Integer   ' ID of <Sense> with meaning definitions
    Dim intI As Integer         ' Counter
    Dim intWait As Integer = 3000   ' Three seconds waiting
    Dim wrThis As IO.StreamWriter   ' This is where I write all the data to
    Dim pdxThis As New XmlDocument
    ' Dim strBodyStart As String = "</table>"
    Dim sgmThis As New Sgml.SgmlReader
    Dim txtThis As System.IO.TextReader
    Dim strDirIn As String = "d:\Data Files\Research\2013\V2_ME\Data"
    Dim strFile As String = strDirIn & "\MED_combined.xml"
    Dim strMedList As String = GetLocalDir() & "\MED_list.htm"
    Dim nmsDf As XmlNamespaceManager

    Try
      ' Validate
      If (Not IO.File.Exists(strFile)) Then
        ' Try alternative
        strFile = "U" & Mid(strFile, 2)
        If (Not (IO.File.Exists(strFile))) Then
          Logging("File not found: " & strFile)
          Return False
        End If
      End If
      If (pdxCurrentFile IsNot Nothing) Then Logging("Close and enter the program again") : Return False
      ' Read the MED-combined dictionary that we have right now
      If (Not ReadXmlDoc(strFile, pdxCurrentFile)) Then Logging("Could not open file") : Return False
      ' Initialise the MED pos dictionary
      If (Not InitMEDpos()) Then Return False
      strPre = "http://quod.lib.umich.edu/cgi/m/mec/med-idx?type=id&id="
      ' Determine the name of the temporary file
      strTmpF = GetLocalDir() & "\MED_tmp.htm"
      ' Check existence of MEDlist
      If (IO.File.Exists(strMedList)) Then
        ' The file already exists, so we should find the correct entry point
        intEntryId = GetLastMedEntry(strMedList)
      Else
        ' The file does not exists, so we create it
        wrThis = New IO.StreamWriter(strMedList, False, System.Text.Encoding.UTF8)
        wrThis.WriteLine("<html><body>") : wrThis.Close()
        ' Set the entry id
        intEntryId = 0
      End If
      ' Other initialisations
      intSenseId = 1 : intLexFunId = 1 : strAlt = "" : strAlt2 = "" : colMis.Clear()
      If (intEntryId = 0) Then
        ' Walk all the entries
        ndxEntry = pdxCurrentFile.SelectSingleNode("./descendant::Entry[1]")
      Else
        ' Start from some point
        ndxEntry = pdxCurrentFile.SelectSingleNode("./descendant::Entry[@EntryId = " & intEntryId & "]")
      End If
      wrThis = Nothing
      While (ndxEntry IsNot Nothing)
        ' Get attributes: MED code, entryid, POS
        strMed = ndxEntry.Attributes("s").Value
        intEntryId = ndxEntry.Attributes("EntryId").Value
        strEntryPos = ndxEntry.Attributes("Pos").Value
        ' Restrictions: for the moment only do the Verbs and only do those that have not yet been done
        ' WAS: If (ndxEntry.ChildNodes.Count = 0) AndAlso (Left(strEntryPos, 1) = "V") Then
        If (ndxEntry.ChildNodes.Count = 0) Then
          nmsDf = Nothing
          ' Determine the address (URL)
          strWpage = strPre & strMed
          ' ============= DEBUG ==============
          ' Navigate to web page and show this to me
          frmMain.TabControl1.SelectedTab = frmMain.tpReport
          bOkay = True
          ' Try to download the web page by making use of the browser
          Do
            Dim intDefault As Integer = 100
            ' Start out positively
            bOkay = True
            ' Sleep at least some time
            System.Threading.Thread.Sleep(intDefault)
            ' Only now continue
            If (Not DownloadHtml(strWpage, strTmpF, wbThis, True)) Then
              bOkay = False
            Else
              ' Transform the result
              txtThis = New IO.StreamReader(strTmpF, System.Text.Encoding.UTF8)
              With sgmThis
                .DocType = "HTML" : .WhitespaceHandling = WhitespaceHandling.All
                .CaseFolding = Sgml.CaseFolding.ToLower
                .InputStream = txtThis
              End With
              ' Read HTML as XML document
              pdxThis.PreserveWhitespace = True
              pdxThis.XmlResolver = Nothing
              pdxThis.Load(sgmThis)
              nmsDf = New XmlNamespaceManager(pdxThis.NameTable)
              nmsDf.AddNamespace("html", pdxThis.DocumentElement.NamespaceURI)
              ' Close the input file
              txtThis.Close()
              ' Check the result for a 503 error
              strTitle = pdxThis.SelectSingleNode("./descendant::html:title", nmsDf).InnerText
              If (InStr(strTitle, "503") > 0 AndAlso InStr(strTitle, "Unavailable") > 0) Then
                bOkay = False
              End If
            End If
            ' Are we okay?
            If (Not bOkay) Then
              ' sleep at least TEN seconds
              intWait = 15000
              System.Threading.Thread.Sleep(intWait)
            End If
          Loop Until bOkay

          ' Prepare for operations
          SetXmlDocument(pdxThis) : colDef.Clear() : colAlt.Clear()
          arPosList = Nothing
          strLemma = "error"
          ' Transform all the <img> nodes
          ndxImg = pdxThis.SelectSingleNode("./descendant::html:img", nmsDf)
          While (ndxImg IsNot Nothing)
            ' Get the letter that is needed
            strText = ndxImg.Attributes("alt").Value
            intPos = InStr(strText, " ")
            If (intPos > 0) Then
              strText = Left(strText, intPos - 1)
            Else
              ' If there is no space, then there is just the [macronbreve], which I interpret to be an [e]
              strText = "e"
            End If
            ' Add the text after the <img> node
            If (ndxImg.NextSibling Is Nothing) OrElse (ndxImg.NextSibling.Name <> "#text") Then
              ' There is no following node or the following node is not #text
              ' Insert a node with text after me
              ndxHtml = AddXmlChildAfter(ndxImg.ParentNode, ndxImg, "")
              ndxHtml.InnerText = strText
            Else
              ' There is already text, so add it after
              ndxImg.NextSibling.InnerText = strText & ndxImg.NextSibling.InnerText
            End If
            ' Remove the <img> node
            ndxImg.ParentNode.RemoveChild(ndxImg)
            ' Find another one
            ndxImg = pdxThis.SelectSingleNode("./descendant::html:img", nmsDf)
          End While
          ' Get all the <p><font> nodes
          ndxList = pdxThis.SelectSingleNode("./descendant::html:table", nmsDf).SelectNodes("./following::html:p[not(@align) or (@align='right' and contains(child::html:font,'['))]/child::html:font", nmsDf)
          ' The first <p><font> node must contain the head word
          If (ndxList.Count > 0) Then
            ' Get the head word node
            ndxHtml = ndxList(0)
            ' The information is in the children of <font>
            ndxChild = ndxHtml.FirstChild
            ' First child should be <strong> with the name of the entry
            strLemma = MedToEnglish(ndxChild.InnerText, strAlt)
            ' Start writing this entry to the Html writer
            wrThis = New IO.StreamWriter(strMedList, True)
            wrThis.WriteLine("<div id='" & intEntryId & "' name='" & strLemma & "' med='" & strMed & "' >")
            wrThis.WriteLine(" <p>" & ndxList(0).InnerXml & "</p>")
            ' Show where we are
            Dim intPtc As Integer = (100 * intEntryId) \ 54590
            Status("Looking at: " & strMed & " = " & strLemma & " " & intPtc & "%", intPtc)
            ' Compare with the entry lemma
            If (strLemma <> ndxEntry.Attributes("l").Value) Then
              ' The Lemma I find and the lemma in the current lexicon is not equal
              ' Stop
              If (strLemma = MedToEnglish(ndxEntry.Attributes("l").Value, strAlt)) Then
                ' Adapt the lexeme value
                ndxEntry.Attributes("l").Value = strLemma
              Else
                ' TODO: decide what to do...
                Stop
              End If
            End If
            ' Are there any alternatives
            If (strAlt <> "") Then
              ' Add the alternative as a separat <LexFun> entry
              SetXmlDocument(pdxCurrentFile)
              ndxLexFun = AddXmlChild(ndxEntry, "LexFun", "EntryId", intEntryId, "attribute", _
                                   "LexFunId", intLexFunId, "attribute", _
                                   "lf", strEntryPos, "attribute", _
                                   "lv", strAlt, "attribute")
              intLexFunId += 1
            End If
            ' Figure out of the lexeme is a name
            bIsName = (AscW(Left(strLemma, 1)) <> AscW(LCase(Left(strLemma, 1))))
            ' Get the text following the lemma
            ndxChild = ndxChild.NextSibling
            If (ndxChild IsNot Nothing) Then
              ' Get the text of this entry
              strText = Trim(ndxChild.InnerText)
              ' Get the part-of-speech
              strPos = Regex.Match(strText, "\w+\.(\w+\.)*").Value
              ' Get the text following the POS
              strText = Trim(Mid(strText, InStrRev(strText, ")") + 1))
              ' strText = Trim(Mid(strText, InStr(strText, strPos) + strPos.Length + 1))
              ' Find out what is left
              If (strText.ToLower = "also") Then
                ' Get the alternatives for the lemma
                colAlt.Clear()
                While (ndxChild IsNot Nothing) AndAlso (Left(ndxChild.InnerText, 1) <> ".")
                  ' Can we add this as alternative?
                  If (ndxChild.Name = "strong") Then
                    ' Get the alternative
                    strAlt = MedToEnglish(ndxChild.InnerText, strAlt2)
                    ' Check for alternative prefixes -- I cannot handle them...
                    If (Right(strAlt, 1) <> "-") Then
                      colAlt.Add(strAlt)
                      ' Add this alternative as a <LexFun> element
                      SetXmlDocument(pdxCurrentFile)
                      ndxLexFun = AddXmlChild(ndxEntry, "LexFun", "EntryId", intEntryId, "attribute", _
                                           "LexFunId", intLexFunId, "attribute", _
                                           "lf", strEntryPos, "attribute", _
                                           "lv", strAlt, "attribute")
                      intLexFunId += 1
                    End If
                  End If
                  ' Go to next child
                  ndxChild = ndxChild.NextSibling
                End While
              ElseIf DoLike(strText, "*)|,") Then
                ' no worries
              ElseIf (strText <> "") Then
                ' It is very likely that the text actually is one particular form of the lexeme
                ' Possibly remove "Forms"
                intPos = InStr(strText, "Forms:")
                If (intPos > 0) Then
                  strText = Trim(Mid(strText, intPos + 6))
                End If
                ' So: process this as a form
                arPosList = GetMEDpos(strText, bIsName, strPos)
                '' Keep track of missing POS stuff
                'If (arPosList Is Nothing) Then colMis.AddUnique(LCase(strText), strMed)
                'If (arPosList Is Nothing) Then
                '  ' Figure out what the text is, and what we can do with it
                '  ' Stop
                'End If
              Else
                ' We can skip empty stuff
              End If
              ' Add the variant forms of the lemma to the <Entry>

              ' There may be inflections coming up now...
              If (ndxChild IsNot Nothing) AndAlso _
                   ((InStr(ndxChild.InnerText, "Forms:") > 0) OrElse _
                     (arPosList IsNot Nothing AndAlso arPosList.Length > 0)) Then
                ' Yes, inflections are coming up -- process the first one
                ' Stop
                If (arPosList Is Nothing OrElse arPosList.Length = 0) Then
                  ' Need to retrieve the first form from the text after "forms"
                  strText = Trim(Mid(ndxChild.InnerText, InStrRev(ndxChild.InnerText, ":") + 1))
                  arPosList = GetMEDpos(strText, bIsName, strPos)
                  '' Keep track of missing POS stuff
                  'If (arPosList Is Nothing) Then colMis.AddUnique(strText, strMed)
                End If
                ' Get the first lexical value
                ndxChild = ndxChild.NextSibling
                While (ndxChild IsNot Nothing)
                  ' is this text or a node?
                  If (ndxChild.Name = "#text") Then
                    ' Get the text
                    strText = Trim(ndxChild.InnerText)
                    Select Case Left(strText, 1)
                      Case ",", "&"
                        ' Another form is following
                      Case ";"
                        ' A different POS follows: get this POS
                        strText = Trim(Mid(strText, 2))
                        arPosList = GetMEDpos(strText, bIsName, strPos)
                        '' Keep track of missing POS stuff
                        'If (arPosList Is Nothing) Then colMis.AddUnique(strText, strMed)
                      Case "."
                        ' End marker
                    End Select

                  ElseIf (ndxChild.Name = "strong") Then
                    ' This is an entry - get and process it
                    strAlt = MedToEnglish(ndxChild.InnerText, strAlt2)
                    If (arPosList IsNot Nothing) Then
                      ' Process this alternative
                      For intI = 0 To arPosList.Length - 1
                        ' Add this variant with the indicated POS
                        SetXmlDocument(pdxCurrentFile)
                        ndxLexFun = AddXmlChild(ndxEntry, "LexFun", "EntryId", intEntryId, "attribute", _
                                             "LexFunId", intLexFunId, "attribute", _
                                             "lf", arPosList(intI), "attribute", _
                                             "lv", strAlt, "attribute")
                        intLexFunId += 1
                      Next intI
                    End If
                  Else
                    ' What else could it be??
                    ' Stop
                  End If

                  ' Next child
                  ndxChild = ndxChild.NextSibling
                End While

                ' Process the following lexical values
              End If

              ' Process the other <p> elements -- or not? 
              ' ( These are the definitions, and what would I do with them??)
              For intI = 1 To ndxList.Count - 1
                ' Keep track of the HTML content
                wrThis.WriteLine(" <p>" & ndxList(intI).InnerXml & "</p>")
                ' My parent may not be a blockquote or a right-aligned <p> node
                If (ndxList(intI).SelectSingleNode("./ancestor::html:blockquote", nmsDf) Is Nothing) AndAlso _
                   (ndxList(intI).Attributes("align") Is Nothing) Then
                  ' If there is only one child, and it has a <strong> tag, then it is a number of a definition
                  If (ndxList(intI).ChildNodes.Count = 1 AndAlso ndxList(intI).FirstChild.Name = "strong") Then
                    ' There is no need to get the number of the definition: we will use our own number

                  Else
                    ' This is a definition - take it
                    colDef.Add(ndxList(intI).InnerText)
                    ' Add the definition
                    SetXmlDocument(pdxCurrentFile)
                    ndxSense = AddXmlChild(ndxEntry, "Sense", "EntryId", intEntryId, "attribute", _
                                           "SenseId", intSenseId, "attribute", _
                                           "l", strLemma, "attribute", _
                                           "N", colDef.Count, "attribute", _
                                           "Pos", strEntryPos, "attribute", _
                                           "Def", ndxList(intI).InnerText, "attribute")
                    intSenseId += 1
                  End If
                End If
              Next intI

            End If
          End If
          Debug.Print(strMed & " Lemma = [" & strLemma & "] Alt = [" & colAlt.Semi & "] Def = [" & colDef.Semi & "]")
          ' Anything to be finished?
          If (wrThis IsNot Nothing) Then
            ' Finish this entry
            wrThis.WriteLine("</div>")
            wrThis.Close()
          End If

        Else
          ' We are skipping those that have already been done as well as non-verbs
        End If
        ' Go to the next entry
        ndxEntry = ndxEntry.SelectSingleNode("./following-sibling::Entry[1]")
      End While
      ' Finish the writer of all
      wrThis = New IO.StreamWriter(strMedList, True)
      wrThis.WriteLine("</body></html>")
      wrThis.Close()
      ' Save what we have in a different file
      strFile = IO.Path.GetDirectoryName(strFile) & "\" & IO.Path.GetFileNameWithoutExtension(strFile) & _
        "_New.xml"
      pdxCurrentFile.Save(strFile)
      Logging("Saved result: " & strFile)
      ' Show which POS parts are still missing
      MsgBox("Missing POS:" & vbCrLf & colMis.Text)
      Stop

      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MEDextend_Org error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   AddLexFunItem
  ' Goal:   Add one item to <LexFun>
  ' History:
  ' 25-04-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function AddLexFunItem(ByRef ndxEntry As XmlNode, ByVal intEntryId As Integer, ByRef intLexFunId As Integer, _
                                  ByVal strLf As String, ByVal strLv As String) As Boolean
    Dim ndxLexFun As XmlNode
    Try
      SetXmlDocument(pdxCurrentFile)
      ndxLexFun = AddXmlChild(ndxEntry, "LexFun", "EntryId", intEntryId, "attribute", _
                           "LexFunId", intLexFunId, "attribute", _
                           "lf", strLf, "attribute", _
                           "lv", strLv, "attribute")
      intLexFunId += 1
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/AddLexFunItem error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MEDextend
  ' Goal:   Extend the lemma-only MED-Lookup dictionary with:
  '         - variant forms
  '         - derived morphemes
  '         - (definitions) -- (this is perhaps not necessary?)
  ' History:
  ' 15-04-2014  ERK Created
  ' 24-04-2014  ERK This versino uses the downloaded portion of the MED
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MEDextend(ByRef wbThis As WebBrowser) As Boolean
    Dim strMed As String        ' MED code
    Dim strDef As String        ' definition pre-amble
    Dim strText As String       ' Text
    Dim strLemma As String      ' The lemma we are looking at
    Dim strLemInf As String     ' Lemma of infinitive
    Dim strPos As String        ' The POS of the lemma
    Dim strEntryPos As String   ' POS of the entry as a whole
    Dim strAlt As String        ' Alternative form
    Dim strAlt2 As String       ' Alternative for alternative
    Dim strTitle As String = "" ' HTML title
    Dim arPosList() As String   ' List of POS possibilities
    Dim bIsName As Boolean      ' Is the entry a name?
    Dim ndxEntry As XmlNode     ' One <Entry> record
    Dim ndxVbInf As XmlNode     ' Verbal infinitive
    Dim ndxSense As XmlNode     ' One <Sense> node
    Dim ndxHtml As XmlNode      ' One node from the <html> entry part
    Dim ndxMed As XmlNode       ' Entry in the MED xml file
    Dim ndxChild As XmlNode
    Dim ndxList As XmlNodeList  ' List of nodes
    Dim colAlt As New StringColl  ' Collection of alternatives
    Dim colDef As New StringColl  ' Collection of definitions
    Dim intEntryId As Integer   ' ID of this entyr
    Dim intLexFunId As Integer  ' ID of <LexFun> element
    Dim intSenseId As Integer   ' ID of <Sense> with meaning definitions
    Dim intI As Integer         ' Counter
    Dim intWait As Integer = 3000   ' Three seconds waiting
    Dim wrThis As IO.StreamWriter   ' This is where I write all the data to
    Dim pdxThis As New XmlDocument
    Dim pdxMed As XmlDocument = Nothing ' MED read from internet
    Dim strMedContinue As String = "also*|?also*|* also*|cont.*|contraction*|early|earlier:|forms|later|etc.|spelled"
    Dim strMedList As String = "D:\Data files\Corpora\Dictionaries\MED\MED_list.xml"
    Dim strDirIn As String = "d:\Data Files\Research\2013\V2_ME\Data"
    Dim strFile As String = strDirIn & "\MED_combined.xml"

    Try
      ' Validate
      If (Not IO.File.Exists(strFile)) Then
        ' Try alternative
        strFile = "U" & Mid(strFile, 2)
        If (Not (IO.File.Exists(strFile))) Then
          Logging("File not found: " & strFile)
          Return False
        End If
      End If
      ' The [MedList] must also exist
      If (Not IO.File.Exists(strMedList)) Then Logging("File not found: " & strMedList) : Return False
      If (pdxCurrentFile IsNot Nothing) Then Logging("Close and enter the program again") : Return False
      ' Read the MED-combined dictionary that we have right now
      Status("Loading " & strFile & "...")
      If (Not ReadXmlDoc(strFile, pdxCurrentFile)) Then Logging("Could not open file") : Return False
      ' Read the MED data that has been taken from the internet
      Status("Loading " & strMedList & "...")
      If (Not ReadXmlDoc(strMedList, pdxMed, "html.xsd")) Then Logging("Could not open MED xml file") : Return False
      ' Other initialisations
      If (Not InitMEDpos()) Then Return False
      intSenseId = 1 : intLexFunId = 1 : strAlt = "" : strAlt2 = "" : colMis.Clear()
      If (intEntryId = 0) Then
        ' Walk all the entries
        ndxEntry = pdxCurrentFile.SelectSingleNode("./descendant::Entry[1]")
      Else
        ' Start from some point
        ndxEntry = pdxCurrentFile.SelectSingleNode("./descendant::Entry[@EntryId = " & intEntryId & "]")
      End If
      wrThis = Nothing : ndxMed = Nothing
      While (ndxEntry IsNot Nothing)
        ' Get attributes: MED code, entryid, POS
        strMed = ndxEntry.Attributes("s").Value
        intEntryId = ndxEntry.Attributes("EntryId").Value
        strEntryPos = ndxEntry.Attributes("Pos").Value
        ' Restrictions: for the moment only do the Verbs and only do those that have not yet been done
        ' WAS: If (ndxEntry.ChildNodes.Count = 0) AndAlso (Left(strEntryPos, 1) = "V") Then
        If (ndxEntry.ChildNodes.Count = 0) Then
          ' Reset the namespace definition
          ndxVbInf = Nothing
          ' Find this entry in the MED
          ' (1) First try nearby...
          If (ndxMed IsNot Nothing) Then
            ndxMed = ndxMed.SelectSingleNode("./following-sibling::div[@med = '" & strMed & "']")
          End If
          ' Found anything?
          If (ndxMed Is Nothing) Then
            ndxMed = pdxMed.SelectSingleNode("./descendant::div[@med = '" & strMed & "']")
          End If
          ' Prepare for operations
          SetXmlDocument(pdxThis) : colDef.Clear() : colAlt.Clear()
          arPosList = Nothing
          ' Get the lemma
          strLemma = ndxMed.Attributes("name").Value
          ' ================= DEBUG ================
          If (strLemma = "wersen") Then Stop
          ' ========================================
          ' Capture VAG entries
          If (strEntryPos = "VAG") Then
            ndxVbInf = Nothing
            ' Try get lemma of infinitive
            If (strLemma Like "*ing") Then
              ' Make infinitive lemma
              strLemInf = Left(strLemma, strLemma.Length - 3) & "en"
              ' See if we can get a verbal infinitive
              ndxVbInf = pdxCurrentFile.SelectSingleNode("./descendant::Entry[@Pos='VB' and @l='" & strLemInf & "']")
            End If
          End If
          ' Get all the <p> child nodes
          ndxList = ndxMed.SelectNodes("./child::p")
          ' The first <p><font> node must contain the head word
          If (ndxList.Count > 0) Then
            ' Get the head word node, which contains its part-of-speech and its variants
            ndxHtml = ndxList(0)
            ' The first <strong> child of the first <p> in the list must contain the lemma and may contain an alternative to it
            ndxChild = ndxHtml.FirstChild : strAlt = ""
            strLemma = MedToEnglish(ndxChild.InnerText, strAlt)
            ' Show where we are
            Dim intPtc As Integer = (100 * intEntryId) \ 54590
            Status("Looking at: " & strMed & " = " & Left(strLemma, 3) & " " & intPtc & "%", intPtc)
            ' Compare with the entry lemma
            If (strLemma <> ndxEntry.Attributes("l").Value) Then
              ' The Lemma I find and the lemma in the current lexicon is not equal
              ' Stop
              If (strLemma = MedToEnglish(ndxEntry.Attributes("l").Value, strAlt)) Then
                ' Adapt the lexeme value
                ndxEntry.Attributes("l").Value = strLemma
              Else
                ' TODO: decide what to do...
                ' Stop
                strLemma = ndxEntry.Attributes("l").Value
              End If
            End If
            ' Add a lexFun for the lemma itself
            If (Not AddLexFunItem(ndxEntry, intEntryId, intLexFunId, strEntryPos, strLemma)) Then Logging("Problem with AddLexFun") : Return False
            ' Possibly add this to the VbInf entry
            If (strEntryPos = "VAG") AndAlso (ndxVbInf IsNot Nothing) Then
              ' Add the alternative as a separate <LexFun> entry
              If (Not AddLexFunItem(ndxVbInf, ndxVbInf.Attributes("EntryId").Value, intLexFunId, strEntryPos, strLemma)) Then Logging("Problem with AddLexFun") : Return False
            End If
            ' Does this lemma have an alternative, e.g. [abil(l)ement] --> abilement is the alternative
            If (strAlt <> "") Then
              ' Add the alternative as a separate <LexFun> entry
              If (Not AddLexFunItem(ndxEntry, intEntryId, intLexFunId, strEntryPos, strAlt)) Then Logging("Problem with AddLexFun") : Return False
              ' Possibly add this to the VbInf entry
              If (strEntryPos = "VAG") AndAlso (ndxVbInf IsNot Nothing) Then
                If (Not AddLexFunItem(ndxVbInf, ndxVbInf.Attributes("EntryId").Value, intLexFunId, strEntryPos, strAlt)) Then Logging("Problem with AddLexFun") : Return False
              End If
            End If
            ' Figure out of the lexeme is a name
            bIsName = (AscW(Left(strLemma, 1)) <> AscW(LCase(Left(strLemma, 1))))
            ' Get the text following the lemma, which may consist of a number of alternatives
            ndxChild = ndxChild.NextSibling
            If (ndxChild IsNot Nothing) Then
              ' Get the text of this entry
              strText = Trim(ndxChild.InnerText)
              ' Get the part-of-speech
              strPos = Regex.Match(strText, "\w+\.(\w+\.)*").Value
              ' Get the text following the POS
              strText = Trim(Mid(strText, InStrRev(strText, ")") + 1))
              ' strText = Trim(Mid(strText, InStr(strText, strPos) + strPos.Length + 1))
              ' Find out what is left
              If (InStr(ndxChild.InnerText.ToLower, " also") > 0) Then
                ' Get the alternatives for the lemma
                colAlt.Clear()
                While (ndxChild IsNot Nothing) AndAlso (Left(ndxChild.InnerText, 1) <> ".") _
                  AndAlso (Left(ndxChild.InnerText, 1) <> ";")
                  ' Can we add this as alternative?
                  If (ndxChild.Name = "strong") Then
                    ' Get the alternative
                    strAlt2 = ""
                    strAlt = MedToEnglish(ndxChild.InnerText, strAlt2)
                    ' Check for alternative prefixes -- I cannot handle them...
                    If (Right(strAlt, 1) <> "-") Then
                      colAlt.Add(strAlt)
                      ' Add this alternative as a <LexFun> element
                      If (Not AddLexFunItem(ndxEntry, intEntryId, intLexFunId, strEntryPos, strAlt)) Then Logging("Problem with AddLexFun") : Return False
                      ' Possibly add the second alternative
                      If (strAlt2 <> "") Then
                        If (Not AddLexFunItem(ndxEntry, intEntryId, intLexFunId, strEntryPos, strAlt2)) Then Logging("Problem with AddLexFun") : Return False
                      End If
                    End If
                  End If
                  ' Go to next child
                  ndxChild = ndxChild.NextSibling
                End While
                ' Check if the stop was caused by the presence of a period
                If (ndxChild IsNot Nothing) AndAlso (ndxChild.NextSibling IsNot Nothing) AndAlso _
                  (Left(ndxChild.InnerText, 1) = "." OrElse Left(ndxChild.InnerText, 1) = ";") Then
                  ' Strip off the period, and see if something remains...
                  strText = Trim(Mid(ndxChild.InnerText, 2))
                  If (strText <> "") Then
                    ' Possibly remove "Forms"
                    With Regex.Match(strText.ToLower, "(forms:|forms|form:|form|froms)")
                      If (.Success) Then
                        strText = Trim(Mid(strText, .Index + .Length + 1))
                      End If
                    End With
                    ' avoid continuations
                    If (Not DoLike(strText.ToLower, strMedContinue)) Then
                      ' ================== DEBUG =====================
                      If (InStr(strText, "also") > 0 OrElse InStr(strText, "contract") > 0) Then Stop
                      ' ==============================================
                      arPosList = GetMEDpos(strText, bIsName, strPos)
                    End If
                  End If
                End If
              ElseIf DoLike(strText, "*)|,") OrElse DoLike(strText.ToLower, strMedContinue) Then
                ' no worries
              ElseIf (ndxChild.NextSibling IsNot Nothing) AndAlso (strText <> "") Then
                ' It is very likely that the text actually is one particular form of the lexeme
                ' Possibly remove "Forms"
                ' Possibly remove "Forms"
                With Regex.Match(strText.ToLower, "(forms:|forms|form:|form|froms)")
                  If (.Success) Then
                    strText = Trim(Mid(strText, .Index + .Length + 1))
                  End If
                End With
                ' So: process this as a form
                ' avoid continuations
                If (Not DoLike(strText.ToLower, strMedContinue)) Then
                  ' ================== DEBUG =====================
                  If (InStr(strText.ToLower, "also") > 0 OrElse InStr(strText.ToLower, "contract") > 0) Then Stop
                  ' ==============================================
                  arPosList = GetMEDpos(strText, bIsName, strPos)
                End If
              Else
                ' We can skip empty stuff
              End If
              ' Add the variant forms of the lemma to the <Entry>

              ' There may be inflections coming up now...
              If (ndxChild IsNot Nothing) AndAlso (ndxChild.NextSibling IsNot Nothing) AndAlso _
                   ((Regex.Match(ndxChild.InnerText, "(forms:|forms|form:|form|froms)").Success) OrElse _
                     (arPosList IsNot Nothing AndAlso arPosList.Length > 0)) Then
                ' Yes, inflections are coming up -- process the first one
                ' Stop
                If (arPosList Is Nothing OrElse arPosList.Length = 0) Then
                  ' Need to retrieve the first form from the text after "forms"
                  strText = Trim(Mid(ndxChild.InnerText, InStrRev(ndxChild.InnerText, ":") + 1))
                  ' ================== DEBUG =====================
                  If (InStr(strText.ToLower, "also") > 0 OrElse InStr(strText.ToLower, "contract") > 0) Then Stop
                  ' ==============================================
                  arPosList = GetMEDpos(strText, bIsName, strPos)
                End If
                ' Get the first lexical value
                ndxChild = ndxChild.NextSibling
                While (ndxChild IsNot Nothing)
                  ' is this text or a node?
                  If (ndxChild.Name = "#text") Then
                    ' Get the text
                    strText = Trim(ndxChild.InnerText)
                    ' Skip indicators like [also] that indicate a continuation of the list
                    If (Not DoLike(strText.ToLower, strMedContinue)) Then
                      Select Case Left(strText, 1)
                        Case ",", "&"
                          ' Another form is following
                        Case ";"
                          ' A different POS follows: get this POS
                          strText = Trim(Mid(strText, 2))
                          ' Skip indicators like [also] that indicate a continuation of the list
                          If (Not DoLike(strText.ToLower, strMedContinue)) Then
                            ' ================== DEBUG =====================
                            If (InStr(strText.ToLower, "also") > 0 OrElse InStr(strText.ToLower, "contract") > 0) Then Stop
                            ' ==============================================
                            arPosList = GetMEDpos(strText, bIsName, strPos)
                          End If
                        Case "."
                          ' End marker
                      End Select
                    End If

                  ElseIf (ndxChild.Name = "strong") Then
                    ' This is an entry - get and process it
                    strAlt = MedToEnglish(ndxChild.InnerText, strAlt2)
                    If (arPosList IsNot Nothing) AndAlso (Left(strAlt, 1) <> "-") Then
                      ' Process this alternative
                      For intI = 0 To arPosList.Length - 1
                        ' Add this variant with the indicated POS
                        If (Not AddLexFunItem(ndxEntry, intEntryId, intLexFunId, arPosList(intI).Replace("|", ";"), strAlt)) Then Logging("Problem with AddLexFun") : Return False
                        intLexFunId += 1
                        ' Check if there is a second alternative
                        If (strAlt2 <> "") Then
                          If (Not AddLexFunItem(ndxEntry, intEntryId, intLexFunId, arPosList(intI).Replace("|", ";"), strAlt2)) Then Logging("Problem with AddLexFun") : Return False
                        End If
                      Next intI
                    End If
                  Else
                    ' What else could it be??
                    ' Stop
                  End If

                  ' Next child
                  ndxChild = ndxChild.NextSibling
                End While

                ' Process the following lexical values
              End If

              ' Process the other <p> elements -- or not? 
              ' ( These are the definitions, and what would I do with them??)
              strDef = ""
              For intI = 1 To ndxList.Count - 1
                ' If there is only one child, and it has a <strong> tag, then it is a number of a definition
                If (ndxList(intI).ChildNodes.Count = 1 AndAlso ndxList(intI).FirstChild.Name = "strong") Then
                  ' There is no need to get the number of the definition: we will use our own number
                  strDef = ndxList(intI).InnerText
                ElseIf (Left(ndxList(intI).InnerText, 1) = "[") Then
                  ' This is a comment that can be added to the whole entry
                  SetXmlDocument(pdxCurrentFile)
                  AddAttribute(ndxEntry, "c", ndxList(intI).InnerText)
                Else
                  ' This is a definition - take it
                  colDef.Add(ndxList(intI).InnerText)
                  ' Add the definition
                  SetXmlDocument(pdxCurrentFile)
                  ndxSense = AddXmlChild(ndxEntry, "Sense", "EntryId", intEntryId, "attribute", _
                                         "SenseId", intSenseId, "attribute", _
                                         "l", strLemma, "attribute", _
                                         "N", colDef.Count, "attribute", _
                                         "Pos", strEntryPos, "attribute", _
                                         "Def", strDef & ndxList(intI).InnerText, "attribute")
                  intSenseId += 1 : strDef = ""
                End If
              Next intI

            End If
          End If
          Debug.Print(strMed & " Lemma = [" & strLemma & "] Alt = [" & colAlt.Semi & "] Def = [" & colDef.Semi & "]")
        Else
          ' We are skipping those that have already been done as well as non-verbs
        End If
        ' Go to the next entry
        ndxEntry = ndxEntry.SelectSingleNode("./following-sibling::Entry[1]")
      End While
      ' Save what we have in a different file
      strFile = IO.Path.GetDirectoryName(strFile) & "\" & IO.Path.GetFileNameWithoutExtension(strFile) & _
        "_New.xml"
      pdxCurrentFile.Save(strFile)
      Logging("Saved result: " & strFile)
      ' Show which POS parts are still missing
      frmMain.wbReport.DocumentText = colMis.Html
      frmMain.TabControl1.SelectedTab = frmMain.tpReport

      ' MsgBox("Missing POS:" & vbCrLf & colMis.Text)
      ' Stop

      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MEDextend error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MedToEnglish
  ' Goal:   Convert MED English to my style of writing English
  '         Pass on bracketed alternatives and remove brackets ultimately
  ' History:
  ' 16-04-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function MedToEnglish(ByVal strIn As String, ByRef strAlt As String) As String
    Dim intLbr As Integer ' Left bracket pos
    Dim intRbr As Integer ' Right bracket pos

    Try
      strIn = strIn.Replace("æ", "ae")
      strIn = strIn.Replace("ð", "dh")
      strIn = strIn.Replace("þ", "th")
      strIn = strIn.Replace("&eth", "dh")
      strIn = strIn.Replace("&thorn", "th")
      ' Are there any brackets?
      intLbr = InStr(strIn, "(") : intRbr = InStr(strIn, ")")
      If (intLbr > 0) AndAlso (intRbr > 0) AndAlso (intRbr > intLbr) Then
        ' Get alternative and regular form
        strAlt = Regex.Replace(strIn, "\(\w+\)", "")
      Else
        strAlt = ""
      End If
      ' Move away prefix and affix indicators (what the brackets are used for
      strIn = strIn.Replace("(", "")
      strIn = strIn.Replace(")", "")
      'Return the result
      Return strIn
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MedToEnglish error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphDictXmlToTdlMED
  ' Goal:   Add the MED dictionary to the MorphDict for ME
  ' History:
  ' 21-01-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MorphDictXmlToTdlMED(ByVal strEngPer As String) As Boolean
    Dim dtrFound() As DataRow         ' Sorted version
    Dim dtrMorph() As DataRow         ' Searching in [Morph]
    Dim dtrThis As DataRow            ' New element
    Dim strLemma As String            ' Current lemma
    Dim strLabel As String            ' Label
    Dim strFeats As String            ' Features
    Dim strPos As String              ' POS
    Dim strType As String             ' Is this a lemma or a derived entry?
    Dim intEntryId As Integer         ' Index into MED_combined
    Dim intPtc As Integer             ' Percentage
    Dim intI As Integer               ' Counter
    Dim intJ As Integer               ' Counter
    Dim intCount As Integer = 0       ' Number of added items
    Dim strMEDfile As String = "U:\Data Files\Research\2013\V2_ME\Data\MED_combined.xml"
    Dim strFileIn As String = "d:\data files\corpora\dictionaries\" & strEngPer & "dict_out.xml"

    Try
      ' We need to create the [tdlMorphDict]
      Status("Opening morphdict...")
      MorphDictIni(False, "MorphDict" & strEngPer & ".xml")
      If (tdlMorphDict Is Nothing) Then Return False
      ' Get the number of current entries in Morph table
      Debug.Print("Table [Morph] size: " & tdlMorphDict.Tables("Morph").Rows.Count)
      ' We also need to read the MED combined dictionary
      Status("Opening MED combined...")
      If (Not ReadDataset("OEdict.xsd", strMEDfile, tdlOEdict)) Then Return False
      ' Walk through all the VERB entries of the MED_combined file
      dtrFound = tdlOEdict.Tables("Entry").Select("Pos LIKE 'V*'")
      For intI = 0 To dtrFound.Length - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ dtrFound.Length
        Status("MorphDict combine MED into Morph " & intPtc & "%", intPtc)
        ' Get the lemma, POS and features
        With dtrFound(intI)
          strLemma = .Item("l").ToString : strPos = .Item("Pos").ToString : strFeats = "s=" & .Item("s").ToString
          intEntryId = .Item("EntryId")
        End With
        ' Any adaptations
        strLemma = strLemma.Replace("æ", "ae")
        Select Case strPos
          Case "VB"
            strType = "l" : strLabel = "Vb"
          Case Else
            strType = "d" : strLabel = "Vb"
        End Select
        If (strType = "l") Then
          ' Check if this entry is available in table [Morph] of tdlMorphDict
          dtrMorph = tdlMorphDict.Tables("Morph").Select("l='" & strLemma & "' AND Pos='" & strPos & "'")
          If (dtrMorph.Length = 0) Then
            ' it is not there yet, so add it
            dtrThis = AddOneDataRow(tdlMorphDict, "Morph", "MorphId", "MorphList")
            With dtrThis
              .Item("EntryId") = intEntryId : .Item("Vern") = strLemma : .Item("l") = strLemma
              .Item("t") = strType : .Item("f") = strFeats : .Item("Label") = strPos : .Item("Pos") = strLabel
            End With
            ' Adapt counter
            intCount += 1
          Else
            ' It is already there
            ' Stop
          End If
        End If
      Next intI
      ' Make sure the result is saved
      Status("Saving...")
      tdlMorphDict.WriteXml(strMorphDictFile)
      Logging("Added " & intCount & " lemma's from MED to MorphDict")
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MorphDictXmlToTdlMED error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphDictXmlToTdl
  ' Goal:   Convert the Gutenberg .xml file into a MorphDict
  ' History:
  ' 03-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MorphDictXmlToTdl(ByVal strEngPer As String) As Boolean
    Dim strText As String = ""        ' Text
    Dim strFeats As String            ' Features
    Dim strLemma As String            ' Lemma
    Dim strLemmaPos As String         ' The POS of the Lemma
    Dim strFileIn As String = "d:\data files\corpora\dictionaries\" & strEngPer & "dict_out.xml"
    Dim strFileSrt As String = ""     ' Sorted file
    Dim strDef As String = ""         ' Sense definition
    Dim strPos As String              ' Part of speech
    Dim strQ As String                ' Query
    Dim strDouble As String           ' Combined double entry
    Dim arLf() As String              ' Content of LF field
    Dim rdThis As XmlReader = Nothing ' An XML reader: I will try to read this in an XML way
    Dim xmlDoc As New XmlDocument     ' One piece of xml we work with
    Dim ndxEntry As XmlNode           ' One entry
    Dim ndxSense As XmlNode           ' One sense
    Dim ndxWork As XmlNode            ' Working node
    Dim ndxList As XmlNodeList        ' List of nodes
    Dim colSrt As New StringColl      ' Sorted version
    Dim colDouble As New StringColl   ' Collection of double entries
    Dim dtrMorph As DataRow           ' One row
    Dim dtrFound() As DataRow         ' Sorted version
    Dim dtrThere() As DataRow         ' Check to see if an entry is already there in terms of Vern/l/Label/f
    Dim setXmlRd As New XmlReaderSettings()
    Dim intEntryId As Integer         ' Index within the [OEdict.xsd] type xml source
    Dim intMorphId As Integer         ' Index
    Dim intDoId As Integer            ' MorphId that should be processed
    Dim intPos As Integer             ' Counter
    Dim intPtc As Integer             ' Percentage
    Dim intI As Integer               ' Counter
    Dim intJ As Integer               ' Counter

    Try
      ' We need to create the [tdlMorphDict]
      MorphDictIni(True, "MorphDict" & strEngPer & ".xml")
      If (tdlMorphDict Is Nothing) Then Return False
      ' Initialise XML reader
      setXmlRd.ProhibitDtd = False : setXmlRd.IgnoreComments = True : setXmlRd.IgnoreWhitespace = True
      rdThis = XmlReader.Create(strFileIn, setXmlRd) : ndxEntry = Nothing
      ' Read the Gutenberg XML dictionary piece by piece
      Status("Converting period " & strEngPer & " XML into MorphDict...")
      While (rdThis.Read)
        ' I only need to read all the <p> sections
        If (rdThis.LocalName = "Entry") AndAlso (rdThis.IsStartElement) Then
          While (XmlReaderGetNextNode(rdThis, xmlDoc, "Entry", ndxEntry))
            If (ndxEntry Is Nothing) Then Logging("modLemmatize/MorphDictXmlToTdl: empty [Entry]") : Return False
            ' Extract some obligatory values from this <Entry>
            intEntryId = ndxEntry.Attributes("EntryId").Value
            strLemma = ndxEntry.Attributes("l").Value
            ' ================ DEBUG ===========
            ' If (strLemma = "note") Then Stop
            ' ==================================
            ' The main entry already contains information that is needed, but I first need to have the <sense> -- take the *FIRST* sense??
            ndxSense = ndxEntry.SelectSingleNode("./descendant::Sense")
            If (ndxSense Is Nothing) Then
              Logging("modLemmatize/MorphDictXmlToTdl: skip empty [Sense] at [" & strLemma & "]")
            ElseIf (ndxSense.Attributes("Def") IsNot Nothing) AndAlso (ndxSense.Attributes("Def").Value Like "See [A-Z]*") Then
              strDef = ndxSense.Attributes("Def").Value
              Logging("modLemmatize/MorphDictXmlToTdl: skip *see* [Sense] at [" & strLemma & "] Def=[" & strDef & "]")
            Else
              ' Get the POS straight away
              strPos = ndxEntry.Attributes("Pos").Value : intPos = InStr(strPos, ";") : strFeats = ""
              If (intPos > 0) Then strFeats = Mid(strPos, intPos + 1) : strPos = Left(strPos, intPos - 1)
              ' Add any features in the @f value
              AddSemiStack(strFeats, GetAttrValue(ndxEntry, "f"), True, ";")
              ' Store the lemma pos
              strLemmaPos = strPos
              ' Certain POS types are *not* allowed as lemma: finite verb forms and plural nouns
              If (DoLike(strPos, "AX*|VB*|DO*|BE*") AndAlso (strPos.Length > 2)) OrElse (DoLike(strPos, "NS")) Then
                ' Skip this one
                Logging("Skipping l=" & strLemma & " POS=" & strPos)
              Else
                ' Process this information into a Morph entry
                dtrMorph = AddOneDataRow(tdlMorphDict, "Morph", "MorphId", "MorphList")
                With dtrMorph
                  .Item("Vern") = ndxEntry.Attributes("l").Value
                  .Item("l") = strLemma
                  ' .Item("Pos") = ndxEntry.Attributes("Pos").Value
                  .Item("Pos") = strPos
                  If (strPos = "Vb") Then
                    .Item("Label") = GetVerbLemmaPos(strLemma, "VB")
                  Else
                    .Item("Label") = GetVerbLemmaPos(strLemma, strPos)
                  End If
                  ' Need any feature?
                  If (strFeats <> "") Then .Item("f") = strFeats
                  ' Add link to the entry in the XML source
                  .Item("EntryId") = intEntryId
                End With
                ' Show where we are
                Status(Left(ndxEntry.Attributes("l").Value, 3))
                ' Get all the <LexFun> entries
                ndxList = ndxEntry.SelectNodes("./child::LexFun")
                For intI = 0 To ndxList.Count - 1
                  ' Get this item
                  ndxWork = ndxList(intI)
                  arLf = Split(GetAttrValue(ndxWork, "lf"), ";")
                  strFeats = ""
                  For intJ = 1 To arLf.Length - 1
                    If (intJ > 1) Then strFeats &= ";"
                    strFeats &= arLf(intJ)
                  Next intJ
                  If (arLf.Length = 0) Then Logging("modLemmatize/MorphDictXmlToTdl: empty [lf]") : Return False
                  ' Get the POS straight away
                  strPos = ndxWork.Attributes("lf").Value : intPos = InStr(strPos, ";") : strFeats = ""
                  If (intPos > 0) Then strFeats = Mid(strPos, intPos + 1) : strPos = Left(strPos, intPos - 1)
                  ' Process it into a new Morph entry
                  dtrMorph = AddOneDataRow(tdlMorphDict, "Morph", "MorphId", "MorphList")
                  With dtrMorph
                    .Item("Vern") = ndxWork.Attributes("lv").Value
                    .Item("l") = strLemma
                    .Item("Pos") = strPos
                    .Item("Label") = GetVerbLemmaPos(strLemma, strPos)
                    '.Item("Pos") = ndxEntry.Attributes("Pos").Value
                    ' .Item("Label") = GetVerbLemmaPos(strLemma, arLf(0))
                    .Item("f") = strFeats
                    ' Add link to the entry in the XML source
                    .Item("EntryId") = intEntryId
                  End With
                Next intI
              End If

            End If

          End While
          'While Not rdThis.EOF
          '  ' Read the <Entry> section
          '  strText = rdThis.ReadOuterXml
          '  ' Check for content
          '  If (strText <> "") Then
          '    xmlDoc.LoadXml(strText)
          '    ndxEntry = xmlDoc.SelectSingleNode("./descendant::Entry")

          '  End If
          'End While
        End If
      End While
      Status("Sorting...")
      '' Save the resulting morphdict
      'tdlMorphDict.WriteXml(strMorphDictFile)
      ' Make a sorted version
      dtrFound = tdlMorphDict.Tables("Morph").Select("", "Vern ASC, l ASC")
      '' Save this sorted version under a different name
      'strFileSrt = GetSetDir() & "\MorphDict-tmp.xml"
      colSrt.Clear() : colSrt.Add("<MorphDict>") : colSrt.Add(" <MorphList>") : intJ = 0
      For intI = 0 To dtrFound.Length - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ dtrFound.Length
        Status("Sorting " & intPtc & "%", intPtc)
        ' Process this entry
        With dtrFound(intI)
          ' ========= DEBUG ============
          ' If (.Item("l").ToString = "note") AndAlso (.Item("Pos").ToString = "VB") Then Stop
          ' ============================
          ' Check against double entries
          intMorphId = dtrFound(intI).Item("MorphId")
          strQ = "MorphId <> " & intMorphId & _
                     " AND Vern='" & .Item("Vern").ToString.Replace("'", "''") & "'" & _
                     " AND l='" & .Item("l").ToString.Replace("'", "''") & "'" & _
                     " AND Pos='" & .Item("Pos").ToString & "'" & _
                     " AND Label='" & .Item("Label").ToString & "'" & _
                     " AND f='" & .Item("f").ToString & "'"
          dtrThere = tdlMorphDict.Tables("Morph").Select(strQ)
          ' Check for doubles
          If (dtrThere.Length > 0) Then
            ' Make sure only one entry is in [colDouble]
            strDouble = .Item("Vern").ToString & "_" & .Item("l").ToString & "_" & .Item("Pos").ToString & _
                        "_" & .Item("Label").ToString & "_" & .Item("f").ToString
            ' Make sure only ONE item gets stored, with ONE particular [MorphId] value
            colDouble.AddUnique(strDouble, intMorphId)
            ' Get the [MorphId] value that ought to be processed
            intDoId = colDouble.Exmp(colDouble.Find(strDouble))
          End If
          If (dtrThere.Length = 0) OrElse (intMorphId = intDoId) Then
            intJ += 1
            colSrt.Add("  <Morph MorphId=" & """" & intJ & """" & " EntryId=" & """" & .Item("EntryId").ToString & """" & _
                       " Vern=" & """" & XmlString(.Item("Vern").ToString) & """" & _
                       " l=" & """" & XmlString(.Item("l").ToString) & """" & _
                       " Pos=" & """" & .Item("Pos").ToString & """" & _
                       " Label=" & """" & .Item("Label").ToString & """" & _
                       " f=" & """" & .Item("f").ToString & """" & _
                       " />")
          End If

        End With
      Next intI
      colSrt.Add(" </MorphList>") : colSrt.Add("</MorphDict>")
      IO.File.WriteAllText(strMorphDictFile, colSrt.Text)
      ' Show we are ready
      Logging("Dictionary saved at: " & strMorphDictFile)
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MorphDictXmlToTdl error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphDictHtmlToXml
  ' Goal:   Convert .html (Gutenberg) dictionary into .xml one
  '         The dictionary, even though it is ME or eModE, will be put in the [tdlOEdict] dataset
  ' History:
  ' 01-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MorphDictHtmlToXml() As Boolean
    Dim strFileIn As String = "d:\data files\corpora\dictionaries\GutenbergMEdict.htm"
    Dim strFileOut As String = "d:\data files\corpora\dictionaries\MEdict_out.xml"
    Dim strPos As String      ' Part-of-speech category
    Dim strLexeme As String   ' The lexeme we are treating
    Dim strLemma As String    ' The current lemma
    Dim strFeats As String    ' Features for an <Entry>
    Dim strText As String     ' Remainder
    Dim strDef As String      ' Definition
    Dim strForm As String     ' One form
    Dim strFormName As String ' Name of form
    Dim strNext As String     ' Following text
    Dim strCategory As String ' Category
    ' Dim arCf() As String      ' Array of cross-references
    Dim intI As Integer       ' Counter
    Dim intJ As Integer       ' Counter
    ' Dim intPtc As Integer     ' Percentage
    Dim intPos As Integer     ' Position
    Dim intForm As Integer    ' Form index
    Dim intSense As Integer   ' Current sense number
    Dim intEntryId As Integer ' Current entry
    Dim intLfId As Integer    ' LexFunId
    Dim arAlt() As String     ' Alternative forms for one lf
    Dim bIsVerb As String     ' Is this a verb entry?
    Dim ndxP As XmlNode       ' The <p> node
    Dim ndxA As XmlNode       ' The anchor node
    Dim ndxB As XmlNode       ' <b> node
    Dim ndxLem As XmlNode     ' The lemma node
    Dim ndxWork As XmlNode    ' Working node
    Dim dtrNew As DataRow     ' Current datarow we are working on
    Dim dtrSense As DataRow   ' Sense
    Dim colLine As New StringColl     ' Output collection
    Dim colForm As New StringColl     ' Collection with forms
    Dim colCat As New StringColl      ' Collection with entry categories
    Dim rdThis As XmlReader = Nothing ' An XML reader: I will try to read this in an XML way
    Dim xmlDoc As New XmlDocument     ' One piece of xml we work with
    Dim encThis As System.Text.Encoding
    Dim setXmlRd As New XmlReaderSettings()

    Try
      ' Validate
      If (Not IO.File.Exists(strFileIn)) Then Return False
      If (Not InitGutenberg()) Then Return False
      ' Create a new dataset
      If (Not CreateDataSet("OEdict.xsd", tdlOEdict)) Then Return False
      ' Read the SFM dictionary
      Status("Reading SFM dictionary...")
      ' encThis = System.Text.Encoding.GetEncoding(1252)
      encThis = System.Text.Encoding.UTF8
      ' Start reading the input file
      setXmlRd.ProhibitDtd = False : setXmlRd.IgnoreComments = True : setXmlRd.IgnoreWhitespace = True
      rdThis = XmlReader.Create(strFileIn, setXmlRd)
      ' Initialisations
      dtrNew = Nothing : dtrSense = Nothing : strPos = "" : strLexeme = "" : intEntryId = 0 : intSense = 0 : intLfId = 0
      IO.File.WriteAllText(strFileOut, "", encThis) : colLine.Clear() : ndxP = Nothing
      colLine.Add("<?xml version='1.0' standalone='yes'?>")
      colLine.Add("<OEdict>")
      colLine.Add(" <EntryList>")
      ' Loop reading
      While (rdThis.Read) AndAlso (Not rdThis.EOF)
        ' I only need to read all the <p> sections
        If (rdThis.LocalName = "p") AndAlso (rdThis.IsStartElement) Then
          Exit While
        End If
      End While
      ' Keep on reading <p> sections/chunks
      While (XmlReaderGetNextNode(rdThis, xmlDoc, "p", ndxP))
        ' Remove all <ins> nodes
        If (Not GutenbergRemoveIns(ndxP)) Then Return False
        If (ndxP IsNot Nothing) Then
          ndxA = ndxP.SelectSingleNode("./child::a")
          ' Is this a full entry?
          If (ndxA Is Nothing) Then
            ' This may be a "Cf" entry
            ndxB = ndxP.SelectSingleNode("./child::b")
            If (ndxB IsNot Nothing) AndAlso (InStr(ndxB.InnerText, "<i>v.") > 0) Then
              ' TODO: Process the "Cf" entry

            End If
          Else
            ' This is a full entry
            ' Check if this is the correct <p> node
            If (ndxA.Attributes("name") IsNot Nothing AndAlso ndxA.Attributes("id") IsNot Nothing) Then
              ' Yes, we have the right section. Now get the lemma
              ndxLem = ndxP.SelectSingleNode("./child::a/child::b")
              If (ndxLem IsNot Nothing) Then
                ' Get the lemma string into proper notation
                strLemma = StringOEtoTagged(LCase(ndxLem.InnerText))
                ' ================= DEBUG
                ' If (strLemma = "striken") Then Stop
                ' ========================
                ' Get the part of speech: the first <i> node
                ndxWork = ndxP.SelectSingleNode("./child::i")
                If (ndxWork IsNot Nothing) Then
                  ' Save this in the collection of entries
                  colCat.AddUnique(Trim(ndxWork.InnerText), strLemma)
                  If DoLike(Trim(ndxWork.InnerText), "vb.|v.|v. *") Then
                    strCategory = "Verb"
                  Else
                    strCategory = "Other"
                  End If
                  Select Case strCategory
                    Case "Verb"
                      ' Initialisations
                      intEntryId += 1 : strPos = "Vb" : colForm.Clear() : strFormName = "Inf" : colForm.Add(strFormName)
                      strFeats = "" : strDef = "" : strNext = ""
                      ' Possible features
                      If (Not ExtractFeats(ndxWork, strFeats)) Then Return False
                      ' Get the definition
                      If (Not GutenBergDefinition(ndxWork, strDef)) Then Return False
                      ' Verb definitions usually start with "to"
                      If (Left(strDef, 3) <> "to ") AndAlso (InStr(strDef, " to ") = 0) Then
                        ' Stop
                        If (DoLike(LCase(strDef), "* see|see *")) Then
                          ' This is not a verb, so leave!
                          Exit Select
                        Else
                          ' Stop
                        End If
                      End If

                      ' Process the main part of this entry
                      If (strFeats = "") Then
                        colLine.Add("  <Entry EntryId=" & """" & intEntryId & """" & " l=" & """" & strLemma & """" & " Pos=" & """" & strPos & """" & ">")
                      Else
                        colLine.Add("  <Entry EntryId=" & """" & intEntryId & """" & " l=" & """" & strLemma & """" & _
                                    " f=" & """" & strFeats & """" & " Pos=" & """" & strPos & """" & ">")
                      End If
                      ' Process the definition
                      If (strDef <> "") Then
                        ' Adapt the definition slightly, if needed
                        strDef = MyTrim(strDef) : intPos = InStrRev(strDef, ";")
                        If (intPos > 0) Then strDef = Trim(Left(strDef, intPos - 1))
                        ' Add a sense entry for the definition
                        intSense += 1
                        colLine.Add("   <Sense EntryId=" & """" & intEntryId & """" & " SenseId=" & """" & intSense & """" & _
                                    " l=" & """" & strLemma & """" & " N=" & """" & "1" & """" & " Pos=" & """" & strPos & """" & _
                                    " Def=" & """" & strDef & """" & " />")
                        ' Check for mdash
                        If DoLike(strDef, loc_strEndIndicator) OrElse (InStr(strDef, "variant of") > 0) Then ndxWork = Nothing
                        ' Possibly eat following <i>
                        If (Not GutenbergSkipTag(ndxWork, "i")) Then Return False
                        ' Possibly eat next text
                        While (ndxWork IsNot Nothing) AndAlso (ndxWork.NodeType = XmlNodeType.Text)
                          ' Check content
                          If (DoLike(ndxWork.InnerText, loc_strEndIndicator)) Then
                            ndxWork = Nothing
                          Else
                            ndxWork = ndxWork.NextSibling
                          End If
                        End While
                      End If
                      ' Get alternative forms
                      While (ndxWork IsNot Nothing)
                        ' Check the kind of sibling we have
                        If (ndxWork.NodeType = XmlNodeType.Element) AndAlso ((ndxWork.Name = "b") OrElse _
                            (ndxWork.Name = "span" AndAlso ndxWork.FirstChild IsNot Nothing AndAlso ndxWork.FirstChild.Name = "b")) Then
                          ' This is an alternative form: get it
                          If (ndxWork.Name = "b") Then
                            strForm = CleanText(ndxWork.InnerText)
                          ElseIf (ndxWork.Name = "span") Then
                            strForm = CleanText(ndxWork.FirstChild.InnerText)
                          Else
                            strForm = ""
                            Stop
                          End If
                          ' Get the following text node
                          ndxWork = ndxWork.NextSibling
                          If (ndxWork Is Nothing) Then Stop
                          strNext = CleanText(ndxWork.InnerText)
                          ' Can we simply check for a period here?
                          If (DoLike(strNext, loc_strEndIndicator) AndAlso Not (DoLike(strNext, loc_strNoEnd))) Then
                            ' We should leave the while-loop
                            Exit While
                          End If
                          ' Skip several possible items
                          If (Not GutenbergSkipItems(ndxWork, strNext)) Then Return False
                          ' Check what the text is that is following
                          If (Right(strNext, 1) = ",") AndAlso (ndxWork.NextSibling IsNot Nothing) AndAlso _
                             (ndxWork.NextSibling.Name = "i") Then
                            ' Expect <i> with alternative form
                            ndxWork = ndxWork.NextSibling
                            ' Validate
                            If (ndxWork Is Nothing) OrElse (ndxWork.Name <> "i") Then Stop
                            ' Get the name of the form
                            strFormName = CleanText(ndxWork.InnerText)
                            colForm.Add(strFormName)
                            ' Add the form to this new collection
                            colForm.Exmp(colForm.Count - 1) &= strForm
                            ' Eat away any following sibling with just text
                            If (ndxWork.NextSibling IsNot Nothing) AndAlso (ndxWork.NextSibling.NodeType = XmlNodeType.Text) Then
                              ' Eat away without looking at content
                              ndxWork = ndxWork.NextSibling
                              ' Skip several possible items
                              If (Not GutenbergSkipItems(ndxWork, strNext)) Then Return False
                              ' Check for mdash or other end indicator
                              If DoLike(ndxWork.InnerText, loc_strEndIndicator) AndAlso Not (DoLike(strNext, loc_strNoEnd)) Then Exit While
                              ' Check for additional grammatical identifier
                              If (ndxWork.NextSibling IsNot Nothing) AndAlso (ndxWork.NextSibling.Name = "i") AndAlso _
                                  (GutenbergIndex(CleanText(ndxWork.NextSibling.InnerText)) >= 0) Then
                                ' Stop
                                ' This is an additional grammatical identifier
                                ndxWork = ndxWork.NextSibling
                                strFormName = CleanText(ndxWork.InnerText)
                                colForm.Add(strFormName)
                                ' Add the form to this new collection
                                colForm.Exmp(colForm.Count - 1) &= strForm
                                ' Eat away following text
                                If (ndxWork.NextSibling IsNot Nothing) AndAlso (ndxWork.NextSibling.NodeType = XmlNodeType.Text) Then
                                  ' Eat awat
                                  ndxWork = ndxWork.NextSibling
                                  ' Check for mdash
                                  If DoLike(ndxWork.InnerText, loc_strEndIndicator) AndAlso Not (DoLike(strNext, loc_strNoEnd)) Then Exit While
                                End If
                              End If
                            End If
                          ElseIf (DoLike(strNext, loc_strEndIndicator) AndAlso Not (DoLike(strNext, loc_strNoEnd))) Then
                            ' Add the form to the previous collection
                            If (colForm.Exmp(colForm.Count - 1) <> "") Then colForm.Exmp(colForm.Count - 1) &= ";"
                            colForm.Exmp(colForm.Count - 1) &= strForm
                            ' Escape from the while loop
                            Exit While
                          Else
                            ' Add the form to the previous collection
                            If (colForm.Exmp(colForm.Count - 1) <> "") Then colForm.Exmp(colForm.Count - 1) &= ";"
                            colForm.Exmp(colForm.Count - 1) &= strForm
                            ' Debug.Print(AscW(Mid(Trim(ndxWork.InnerText), 2, 1)) & AscW(vbLf))
                          End If
                        ElseIf (ndxWork.NodeType = XmlNodeType.Element) AndAlso (ndxWork.Name = "i") Then
                          ' Action depends on the kind of <i> we get here
                          strForm = CleanText(ndxWork.InnerText)
                          ' There is the possibility that this is a form-indicator of the previous one
                          intForm = GutenbergIndex(strForm)
                          If (intForm > 0) AndAlso (colForm.Count > 0) Then
                            ' Add the form to the previous collection
                            colForm.Item(colForm.Count - 1) = strForm
                          Else
                            ' Check for other possibilities
                            Select Case strForm
                              Case "Phr.", "Comb."
                                ' |Time to leave the scene
                                Exit While
                              Case "intr."
                                ' Okay, continue
                              Case "for"
                                ' Read and dismiss the next form in <b> as well as any following text
                                If (ndxWork.NextSibling IsNot Nothing) AndAlso (ndxWork.NextSibling.NodeType = XmlNodeType.Element) _
                                  AndAlso (ndxWork.NextSibling.Name = "b") Then
                                  ' Skip the next form
                                  ndxWork = ndxWork.NextSibling
                                  ' Check for possible additional clutter
                                  If (ndxWork.NextSibling IsNot Nothing) AndAlso (ndxWork.NextSibling.NodeType = XmlNodeType.Text) Then
                                    ' Skip this text
                                    ndxWork = ndxWork.NextSibling
                                  End If
                                End If
                              Case Else
                                If (Not DoLike(strLemma, "bi-steken|flowten")) Then
                                  ' Stop
                                End If
                            End Select
                          End If
                        End If
                        ' Go to the next sibling
                        ndxWork = ndxWork.NextSibling
                      End While
                      ' Process all the forms we received
                      For intI = 0 To colForm.Count - 1
                        ' Walk all the possibilities
                        arAlt = Split(colForm.Exmp(intI), ";")
                        For intJ = 0 To arAlt.Length - 1
                          If (arAlt(intJ) <> "") Then
                            ' Try to determine the correct form name
                            intForm = GutenbergIndex(colForm.Item(intI))
                            If (intForm < 0) Then
                              Debug.Print("Cannot locate form-name " & colForm.Item(intI) & vbTab & _
                                          "Lexeme=[" & strLemma & "]")
                              'Stop
                            Else
                              ' Check if this needs processing
                              strFormName = arFormPos(intForm)
                              If (strFormName <> "") Then
                                ' Yes, needs processing
                                If (arFormFeat(intForm) <> "") Then
                                  strFormName &= ";" & arFormFeat(intForm)
                                End If
                                ' Process this one
                                intLfId += 1 : arAlt(intJ) = StringOEtoTagged(LCase(arAlt(intJ)))
                                colLine.Add("   <LexFun EntryId=" & """" & intEntryId & """" & " LexFunId=" & """" & intLfId & """" & _
                                            " lf=" & """" & strFormName & """" & " lv=" & """" & arAlt(intJ) & """" & " />")
                              End If
                            End If
                          End If
                        Next intJ
                      Next intI
                      ' Finish this entry
                      colLine.Add("  </Entry>")
                      ' Possibly flush
                      If (colLine.Count > 2047) Then
                        colLine.Flush(strFileOut) : colLine.Clear()
                      End If
                    Case "Other"
                    Case Else
                      '' We do not handle other forms yet
                      'If (InStr(Trim(ndxWork.InnerText), "v.") = 1) Then Stop
                  End Select
                End If
              End If
            End If
          End If      ' ndxA Is NOT Nothing
        End If        ' ndxP IsNot Nothing
      End While
      ' Close the reader
      rdThis.Close() : rdThis = Nothing
      ' Flush and output
      colLine.Add(" </EntryList>")
      colLine.Add("</OEdict>")
      colLine.Flush(strFileOut)
      ' Show all the different entries
      frmMain.wbReport.DocumentText = colCat.Html
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MorphDictHtmlToXml error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      If (rdThis IsNot Nothing) Then rdThis.Close()
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   CleanText
  ' Goal:   Clean the text
  ' History:
  ' 01-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function CleanText(ByVal strIn As String) As String
    Try
      strIn = strIn.Replace(vbCrLf, " ")
      strIn = strIn.Replace(vbCr, " ")
      strIn = strIn.Replace(vbLf, " ")
      strIn = strIn.Replace(Chr(160), " ")
      '' Double check
      'If (InStr(strIn, vbLf) > 0) Then Stop
      Return Trim(strIn)
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/CleanText error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   ExtractFeats
  ' Goal:   Try to extract features 
  ' History:
  ' 02-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function ExtractFeats(ByRef ndxWork As XmlNode, ByRef strFeats As String) As Boolean
    Try
      strFeats = ""
      If (DoLike(ndxWork.InnerText, "* refl.|* reflex.")) Then
        strFeats = "arg=refl"
      ElseIf (DoLike(ndxWork.InnerText, "* tr.|* trans.")) Then
        strFeats = "arg=trans"
      ElseIf (DoLike(ndxWork.InnerText, "* int.|* intrans.")) Then
        strFeats = "arg=intrans"
      End If
      ' Check if there is another argument structure indicator
      If (ndxWork.NextSibling IsNot Nothing) AndAlso (ndxWork.NextSibling.NodeType = XmlNodeType.Text) AndAlso _
         (CleanText(ndxWork.NextSibling.InnerText) = "and") AndAlso (ndxWork.NextSibling.NextSibling IsNot Nothing) AndAlso _
         (ndxWork.NextSibling.NextSibling.NodeType = XmlNodeType.Element) AndAlso _
         (ndxWork.NextSibling.NextSibling.Name = "i") Then
        ' This entry has a double argument structure: process it
        ndxWork = ndxWork.NextSibling : ndxWork = ndxWork.NextSibling
        'strFeats &= ";arg=" & ndxWork.InnerText
        Select Case ndxWork.InnerText
          Case "intr."
            strFeats &= ";arg=intrans"
          Case "tr."
            strFeats &= ";arg=trans"
          Case Else
            Stop
        End Select
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/ExtractFeats error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GutenbergSkipItems
  ' Goal:   Skip bracketed information and <ins> nodes
  ' History:
  ' 02-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function GutenbergSkipItems(ByRef ndxWork As XmlNode, ByRef strNext As String) As Boolean
    Try
      ' (1) Skip what is between brackets
      If ((Right(strNext, 1) = "(") OrElse (Right(CleanText(ndxWork.InnerText), 1) = "(")) AndAlso (Not DoLike(ndxWork.InnerText, loc_strEndIndicator)) Then
        ' Eat away until the first ) appears
        While (ndxWork.NextSibling IsNot Nothing) AndAlso (InStr(ndxWork.NextSibling.InnerText, ")") = 0)
          ndxWork = ndxWork.NextSibling
        End While
        ' Make sure we include the ) in eating away
        If (ndxWork.NextSibling IsNot Nothing) Then
          ' Eat
          ndxWork = ndxWork.NextSibling
          ' Get the correct [strNext]
          strNext = CleanText(ndxWork.InnerText)
          strNext = Mid(strNext, InStr(strNext, ")") + 1)
        End If
      End If
      ' (2) Skip <ins>
      If (ndxWork.NextSibling IsNot Nothing) AndAlso (ndxWork.NextSibling.NodeType = XmlNodeType.Element) AndAlso _
         (ndxWork.NextSibling.Name = "ins") Then
        ' Skip this insertion
        ndxWork = ndxWork.NextSibling
        ' Get the following text node
        ndxWork = ndxWork.NextSibling
        If (ndxWork Is Nothing) Then Stop
        strNext = CleanText(ndxWork.InnerText)
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/GutenbergSkipItems error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GutenbergSkipTag
  ' Goal:   Skip a following node of name <strTag>
  ' History:
  ' 02-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function GutenbergSkipTag(ByRef ndxWork As XmlNode, ByVal strTag As String) As Boolean
    Try
      ' (2) Skip <ins>
      If (ndxWork IsNot Nothing) AndAlso (ndxWork.NodeType = XmlNodeType.Element) AndAlso _
         (ndxWork.Name = strTag) Then
        ' Skip this alement
        ndxWork = ndxWork.NextSibling
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/GutenbergSkipTag error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GutenbergRemoveIns
  ' Goal:   Remove <ins> nodes
  ' History:
  ' 02-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function GutenbergRemoveIns(ByRef ndxRoot As XmlNode) As Boolean
    Dim ndxList As XmlNodeList  ' List of nodes
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (ndxRoot Is Nothing) Then Return False
      ' Get alist of ins nodes
      ndxList = ndxRoot.SelectNodes("./descendant::ins")
      ' Walk them
      For intI = ndxList.Count - 1 To 0 Step -1
        ' Delete it
        ndxList(intI).RemoveAll()
        ndxList(intI).ParentNode.RemoveChild(ndxList(intI))
      Next intI
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/GutenbergRemoveIns error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GutenBergDefinition
  ' Goal:   Get the definition
  ' History:
  ' 02-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function GutenBergDefinition(ByRef ndxWork As XmlNode, ByRef strDef As String) As Boolean
    Dim strNext As String = ""

    Try
      ' Get the definition
      ndxWork = ndxWork.NextSibling
      ' possibly process first an <i> entry
      If (ndxWork IsNot Nothing) AndAlso (ndxWork.NodeType = XmlNodeType.Element) AndAlso (ndxWork.Name = "i") Then
        ' Skip this one
        ndxWork = ndxWork.NextSibling
      End If
      If (ndxWork.NodeType = XmlNodeType.Text) Then
        strDef = Trim(ndxWork.InnerText)
        ' Possibly skip (...) stuff
        If (Not GutenbergSkipItems(ndxWork, strNext)) Then Return False
        ' Combine
        If (Right(strDef, 1) = "(") Then strDef = Left(strDef, strDef.Length - 1)
        strDef &= strNext
        ndxWork = ndxWork.NextSibling
      Else
        strDef = ""
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/GutenBergDefinition error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphDictEnrichVernPos
  ' Goal:   Evaluate the forms of each verb lemma in [VernPos] and add the standard derivable forms
  ' Note:   This procedure does not work correctly AVOID USING IT
  ' History:
  ' 04-10-2013  ERK Created
  ' 14-02-2014  ERK Added language-dependance
  ' 27-02-2014  ERK Rewrite of the procedure for eModE, LmodE
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MorphDictEnrichVernPos(ByVal strLngAbbr As String, ByRef bChanged As Boolean) As Boolean
    Dim strVern As String = ""    ' Vernacular
    Dim strV As String = ""       ' The "v" form may contain hyphens for preverb breaks
    Dim strPos As String = ""     ' POS
    Dim strLemma As String = ""   ' Lemma
    Dim intEntryId As Integer = 0 ' Link to Gutenberg lemma Id
    Dim strFeats As String = ""   ' Features
    Dim strFeatsP As String = ""  ' extended features
    Dim strSfx As String = ""     ' Suffix
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim dtrMorph() As DataRow     ' Check availability of particular form in [Morph] table
    Dim bHasImp As Boolean        ' This lemma has an imperative form
    Dim bHasPst As Boolean        ' Lemma has a past tense form
    Dim bHasPrs As Boolean        ' Lemma has a present tense form
    Dim intPtc As Integer         ' Percentage
    Dim intDouble As Integer = 0   ' Number of double entries
    Dim intI As Integer           ' Counter

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      Select Case MsgBox("This procedure is known to work incorrectly. Continue?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.No, MsgBoxResult.Cancel
          Return False
      End Select
      ' Walk through "Morph" entries
      ' Each lemma has a unique EntryId and a unique lemma-form, which distinguishes it from others
      ' (This assumes that the very first label is VB, but what if the first label is VA[G]?? )
      dtrFound = tdlMorphDict.Tables("Morph").Select("", "l ASC, EntryId ASC, Label ASC")
      bHasImp = False : bHasPst = False : bHasPrs = False
      For intI = 0 To dtrFound.Length - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ dtrFound.Length
        Status("Enriching VernPos " & intPtc & "%", intPtc)
        ' Process this entry
        With dtrFound(intI)
          ' Get the POS and the lemma
          strPos = .Item("Label").ToString : strLemma = .Item("l").ToString
          ' Is this an infinitive?
          If (strPos = "VB") Then
            ' This is the infinitive form -- check for the availability of other forms derived from this one
            Select Case LCase(strLngAbbr)
              Case "emode", "lmode"
                ' Capture features
                strFeats = .Item("f").ToString
                ' Check if a VBI form based on this entryId is available
                dtrMorph = tdlMorphDict.Tables("Morph").Select("EntryId=" & intEntryId & " AND Pos='VBI'")
                If (dtrMorph.Length = 0) Then
                  ' Create and add a VBI form in [Morph]
                  'If (Not MorphDictAdd(strVern, "VBI", "VBI", strLemma, "")) Then Return False
                  strVern = strLemma
                  If (Not MorphDictAdd(strVern, "VBI", "VBI", strLemma, strFeats, intEntryId:=intEntryId)) Then Return False
                  ' Create and add a VBI form in [VernPos]
                  If (Not MorphAddOneVernPos(strVern, strLemma, "VBI", strFeats, "Enrich VernPos at [" & Format(Now, "G") & "]")) Then Return False
                  ' Show we have made changes
                  bChanged = True
                Else
                  Stop
                End If
                ' Check if a VBP form based on this entryId is available
                dtrMorph = tdlMorphDict.Tables("Morph").Select("EntryId=" & intEntryId & " AND Pos='VBP'")
                If (dtrMorph.Length = 0) Then
                  ' Create a VBP from the infinitive
                  strVern = strLemma
                  ' Create and add two VBP forms in [Morph]
                  If (Not MorphDictAdd(strVern, "VBP", "VBP", strLemma, strFeats, intEntryId:=intEntryId)) Then Return False
                  ' Create and add two VBP forms in [VernPos]
                  If (Not MorphAddOneVernPos(strVern, strLemma, "VBP", strFeats, "Enrich VernPos at [" & Format(Now, "G") & "]")) Then Return False
                  ' Show we have made changes
                  bChanged = True
                  ' Can we make 3rd person singular variants?
                  If (Not DoLike(Right(strLemma, 1), "[aeiuoy]")) Then
                    ' Adapt features
                    strFeatsP = strFeats : AddSemiStack(strFeatsP, "person=3;number=sg", True)
                    ' Form of suffix depends on how the lemma ends
                    If (DoLike(strLemma, "*sh|*z|*ch|*s")) Then
                      strSfx = "es"
                    Else
                      strSfx = "s"
                    End If
                    If (Not MorphDictAdd(strVern & strSfx, "VBP", "VBP", strLemma, strFeatsP, bUseFeat:=True, intEntryId:=intEntryId)) Then Return False
                    If (Not MorphAddOneVernPos(strVern & strSfx, strLemma, "VBP", strFeatsP, "Enrich VernPos at [" & Format(Now, "G") & "]", True)) Then Return False
                    'Else
                    '  Stop
                  End If
                Else
                  'Stop
                End If
            End Select
          End If
          ' Check if this is a new one
          '       If (strLemma = .Item("l").ToString) AndAlso _
          '           (intEntryId < 0) OrElse ((.Item("EntryId").ToString <> "") AndAlso (intEntryId = CInt(.Item("EntryId").ToString))) Then
          If (strLemma = .Item("l").ToString) Then
            ' This is a form belonging to the lemma under current consideration
            strPos = .Item("Label").ToString
            Select Case strPos
              Case "VBI"
                bHasImp = True
              Case "VBD"
                bHasPst = True
              Case "VBP"
                bHasPrs = True
            End Select
          ElseIf (.Item("Label").ToString Like "VA*") Then
            ' Skip all VA[G] entries
          Else
            ' Do we have an old entry?
            If (strLemma <> "") Then
              ' Debugging
              ' If (strLemma = "leten") Then Stop
              ' Process the old entry
              ' (1) Do we have an imperative form?
              If (Not bHasImp) Then
                ' Try to make an imperative from the infinitive
                If (strLemma.Length >= 5) AndAlso ((DoLike(strLngAbbr, "OE|ME") AndAlso (Right(strLemma, 2) = "en"))) Then
                  ' Create an imperative form
                  strVern = Left(strLemma, strLemma.Length - 2)
                  ' Validate
                  If (InStr(strVern, " ") = 0) Then
                    ' Check last letter
                    If (DoLike(strVern, "*[aeiouy][bcdfgklmnpqrstvwxz]|*[aeiouy]sh|*[aeiouy][chlnrs]k")) Then
                      ' This is a conservative estimate...
                      If (Not MorphAddOneVernPos(strVern, strLemma, "VBI", "", "Derived at [" & Format(Now, "G") & "]")) Then Return False
                      ' Show we have made changes
                      bChanged = True
                    End If
                  End If
                End If
              End If
            End If
            ' Start a new entry
            strFeats = .Item("f").ToString
            ' Reset the flags
            bHasImp = False : bHasPst = False : bHasPrs = False
          End If
          ' Keep track of who I am
          strLemma = .Item("l").ToString
          If (.Item("EntryId").ToString = "") Then
            intEntryId = -1
          Else
            intEntryId = CInt(.Item("EntryId").ToString)
          End If
        End With
      Next intI
      ' Make sure any changes are processed
      Status("processing changes...")
      If (bChanged) Then tdlMorphDict.AcceptChanges()
      Status("Ready")
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MorphDictEnrichVernPos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphAddOneVernPos
  ' Goal:   Add one entry to the vernpos table
  ' History:
  ' 04-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MorphAddOneVernPos(ByVal strVern As String, ByVal strLemma As String, ByVal strPos As String, _
       ByVal strFeats As String, ByVal strSrc As String, Optional ByVal bUseFeat As Boolean = False, _
       Optional ByVal intMorphId As Integer = -1) As Boolean
    Dim dtrFound() As DataRow ' Result of selecting
    Dim dtrThis As DataRow    ' one datarow

    Try
      ' Validate
      If (bUseFeat) Then
        dtrFound = tdlMorphDict.Tables("VernPos").Select("Vern='" & strVern.Replace("'", "''") & _
           "' AND l='" & strLemma.Replace("'", "''") & _
           "' AND Pos='" & strPos & "' AND f='" & strFeats & "'")
      Else
        dtrFound = tdlMorphDict.Tables("VernPos").Select("Vern='" & strVern.Replace("'", "''") & _
           "' AND l='" & strLemma.Replace("'", "''") & "' AND Pos='" & strPos & "'")
      End If
      ' Check if this entry is already present
      If (dtrFound.Length > 0) Then
        ' At least adapt the @f value, if needed
        dtrThis = dtrFound(0)
        With dtrThis
          If (.Item("f").ToString <> strFeats) Then
            .Item("f") = strFeats
            ' Flag changes
            bMorphVernPosChanged = True
          End If
          ' Check on the MorphId
          If (intMorphId >= 0) AndAlso (.Item("MorphId").ToString = "" OrElse .Item("MorphId").ToString = "-1") Then
            ' Adapt the morphid
            .Item("MorphId") = intMorphId
          End If
        End With
      Else
        ' Entry is not yet present, so add it
        dtrThis = AddOneDataRow(tdlMorphDict, "VernPos", "VernPosId", "VernPosList")
        With dtrThis
          .Item("Vern") = strVern.Replace("-", "") : .Item("v") = strVern : .Item("Pos") = strPos
          .Item("Type") = IIf(strFeats = "", "LemmaOnly", "LemmaFeat")
          .Item("l") = strLemma : .Item("f") = strFeats : .Item("Lev") = 1 : .Item("Src") = strSrc
          .Item("MorphId") = intMorphId
        End With
        ' Flag changes
        bMorphVernPosChanged = True
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MorphAddOneVernPos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphVernPosChanged
  ' Goal:   Check if any VernPos additions have been made
  ' History:
  ' 04-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MorphVernPosChanged() As Boolean
    Try
      ' Return result
      Return bMorphVernPosChanged
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MorphAddOneVernPos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphRemoveEmptyLemmas
  ' Goal:   Remove empty lemma's from tdl MorphDict
  ' History:
  ' 17-02-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MorphRemoveEmptyLemmas(ByRef intNum As Integer) As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim intI As Integer       ' Counter

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Initialise
      intNum = 0
      ' Find empty lemma's
      dtrFound = tdlMorphDict.Tables("VernPos").Select("l=' '")
      For intI = dtrFound.Length - 1 To 0 Step -1
        ' Remove this one
        dtrFound(intI).Delete()
      Next intI
      ' Adapt it
      tdlMorphDict.AcceptChanges()
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MorphRemoveEmptyLemmas error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphDelPreVb
  ' Goal:   Return a form of [strForm] without preverb
  ' History:
  ' 27-06-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function MorphDelPreVb(ByVal strForm As String) As String
    Dim intI As Integer     ' Counter

    Try
      ' Validate
      If (strForm = "") Then Return ""
      ' Check all forms in order
      For intI = 0 To arPreVb.Length - 1
        ' Try to match
        If (InStr(strForm, arPreVb(intI)) = 1) Then
          ' This is a match - make it complete
          Return Mid(strForm, arPreVb(intI).Length + 1)
        End If
      Next intI
      ' Nothing found, so return empty
      Return ""
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MorphDelPreVb error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphDictAdaptBT
  ' Goal:   Adapt the tdlMorphDict with information from the BT, stored in [d:\data files\dbase\BTlemma.xml]
  ' History:
  ' 05-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MorphDictAdaptBT(ByVal strType As String) As Boolean
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim dtrVP() As DataRow        ' Entry from [VernPos] table
    Dim dtrNew As DataRow         ' One new entry for [Morph] table
    Dim dtrEntry() As DataRow     ' Walk through entries
    Dim dtrLexFun() As DataRow    ' Walk through LexFuns
    Dim dtrCat() As DataRow       ' Entry in [Cat]
    Dim dtrMorph() As DataRow     ' Section from [Morph] with same lemma
    Dim dtrAmbi() As DataRow      ' Potentially ambiguous
    Dim dtrRemv() As DataRow      ' Result of looking for the one that needs to be removed
    Dim arLf() As String          ' array of @lf
    Dim strLf As String = ""      ' Lexical function
    Dim strF As String = ""       ' Semi-separated array of features
    Dim arF(0) As String          ' Array of features
    Dim strLv As String = ""      ' form
    Dim strLemma As String = ""   ' Lemma
    Dim strLemmaOld As String     ' Old lemma
    Dim strPos As String          ' Part of speech
    Dim strPosThis As String      ' Current pos
    Dim strPosMain As String      ' Main category for this pos
    Dim strPosList As String      ' List of POS cateogries
    Dim strVern As String = ""    ' The vernacular form
    Dim strMed As String = ""     ' The MED number
    Dim strFilter As String = ""  ' Filter to be used
    Dim strFile As String = ""    ' Filename (tab-separated)
    Dim strBest As String = ""    ' Best
    Dim strForm As String = ""    ' Form
    Dim strExmp As String = ""    ' Example
    Dim strHasVP As String = ""   ' String telling whether this entry has related entries in [VernPos] or not
    Dim arText() As String        ' Lines of file
    Dim arLine() As String        ' Parts of one line
    Dim colHtml As New StringColl ' Error page
    Dim colEntry As New StringColl ' List of EntryIds that need to be deleted
    Dim tdlSpare As DataSet       ' Shadow collection
    Dim wrThis As IO.StreamWriter = Nothing
    Dim bLemmaAmbi As Boolean     ' Is this lemma ambiguous or not?
    Dim bCorrect As Boolean       ' Do we have the correct lemma in [Morph] already?
    Dim bChanged As Boolean       ' Anything changed?
    Dim bHasTasL As Boolean       ' Does this lemma have an entry where @t='l'?
    Dim intBest As Integer        ' Index of best match
    Dim intConf As Integer        ' Score of best match
    Dim intValue As Integer       ' Comparison value
    Dim intEquals As Integer      ' Number of equal letters
    Dim intEntryId As Integer     ' Index in MED file
    Dim intMorphId As Integer     ' MorphId
    Dim intCorrect As Integer     ' Index of correct lemma
    Dim intMnum As Integer = 0    ' Additions in [Morph] table
    Dim intVPnum As Integer = 0   ' Additions in [VernPos] table
    Dim intPtc As Integer         ' Percentage
    Dim intNum As Integer         ' Number of entries
    Dim intPos As Integer         ' Position in string
    Dim intK As Integer           ' Counter
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intL As Integer           ' counter

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Logging("The morphdict is not opened") : Return False
      ' Initialise the BT dictionary, which contains the 'correct' lemma's, their POS and their BT numbers
      If (Not InitBT()) Then Logging("Could not initialize BT dictionary") : Return False
      ' Get the restriction
      Select Case strType.ToLower
        Case "verbs", "verb"
          strFilter = "Pos = 'VB'"
      End Select
      ' Step #1: make the attribute @t where it is not present yet
      Select Case MsgBox("B&T adaptation for OE step #1: list of un-done @t attributes in [Morph]?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' Exit
          Status("Aborted at #1") : bInterrupt = True : Return False
        Case MsgBoxResult.No
          ' Skip this step
        Case MsgBoxResult.Yes
          ' Sort [Morph] on the @l field
          dtrMorph = tdlMorphDict.Tables("Morph").Select("t Is Null", "l ASC")
          strLemma = "" : bChanged = False : strPosMain = "" : strPosThis = ""
          For intI = 0 To dtrMorph.Length - 1
            ' Where are we?
            intPtc = (intI + 1) * 100 \ dtrMorph.Length
            Status("MorphDictAdaptBT [" & Left(strLemma, 3) & "] " & intPtc & "%", intPtc)
            Debug.Print(dtrMorph(intI).Item("MorphId").ToString & vbTab & _
                        dtrMorph(intI).Item("Vern").ToString & vbTab & _
                        dtrMorph(intI).Item("Pos").ToString & vbTab & _
                        dtrMorph(intI).Item("l").ToString)
          Next intI
          Stop
      End Select

      ' Step #2: make the attribute @t where it is not present yet
      Select Case MsgBox("B&T adaptation for OE step #2: repair @t attribute in [Morph]?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' Exit
          Status("Aborted at #2") : bInterrupt = True : Return False
        Case MsgBoxResult.No
          ' Skip this step
        Case MsgBoxResult.Yes
          ' Sort [Morph] on the @l field
          dtrMorph = tdlMorphDict.Tables("Morph").Select("", "l ASC")
          strLemma = "" : bChanged = False : strPosMain = "" : strPosThis = ""
          For intI = 0 To dtrMorph.Length - 1
            ' Where are we?
            intPtc = (intI + 1) * 100 \ dtrMorph.Length
            Status("MorphDictAdaptBT [" & Left(strLemma, 3) & "] " & intPtc & "%", intPtc)
            ' Need the POS
            strPos = dtrMorph(intI).Item("Pos").ToString
            strPosMain = MorphCatMain(strPos) : intPos = InStr(strPosMain, "+") : If (intPos > 0) Then strPosMain = Mid(strPosMain, intPos + 1)
            If (strPos Like "NR*") andalso (strPosMain <> "NR") Then strPosMain = "NR"
            '  Debug.Print("Head of [" & strPos & "] is [" & strPosMain & "]")
            ' ============ DEbugging =============
            ' If (dtrMorph(intI).Item("l").ToString.Replace("-", "") = "abannan") Then Stop
            '   If (dtrMorph(intI).Item("MorphId") = 113121) Then Stop
            ' ====================================
            ' Is this a new lemma?
            If (strLemma <> dtrMorph(intI).Item("l").ToString) OrElse (strPosThis <> strPosMain) Then
              ' Start processing the new lemma
              strLemma = dtrMorph(intI).Item("l").ToString
              ' ============ DEBUG =============
              ' If (strPos Like "VBN*") Then Stop
              ' ================================
              strPosThis = MorphCatMain(strPos)
              If (strPos Like "NR*") AndAlso (strPosThis <> "NR") Then strPosThis = "NR"
              strPosList = MorphCatList(strPosThis)
              ' Ad-hoc adaptation for RP
              If (strPos Like "RP+*") OrElse ((strPosMain = "VB") AndAlso (InStr(strLemma, "-") > 0)) Then
                strPosList = "( " & strPosList & " OR (Pos LIKE 'RP+*') )"
                 End If
              ' Initialisations
              bHasTasL = False
              ' Make sure we treat hyphens fairly
              If (strLemma.Contains("-")) Then
                dtrFound = tdlMorphDict.Tables("Morph").Select("(l='" & strLemma.Replace("'", "''") & "' OR l='" & _
                            strLemma.Replace("'", "''").Replace("-", "") & "') AND " & strPosList)
              Else
                dtrFound = tdlMorphDict.Tables("Morph").Select("l='" & strLemma.Replace("'", "''") & "' AND " & strPosList)
              End If
              intCorrect = -1 : bCorrect = False
              For intJ = 0 To dtrFound.Length - 1
                ' Check for @t equals "l"
                If (dtrFound(intJ).Item("t").ToString = "l") Then bHasTasL = True
                ' Check for possible lemma
                If (dtrFound(intJ).Item("Vern").ToString.Replace("-", "") = strLemma.Replace("-", "")) Then
                  ' This could be the lemma
                  intCorrect = intJ : bCorrect = True
                End If
              Next intJ
              ' Where are we?
              If (Not bHasTasL) Then
                If (bCorrect) Then
                  ' Marke the correct lemma
                  dtrFound(intCorrect).Item("t") = "l" : bChanged = True
                Else
                  ' Indicate that we need to have a lemma entry for this one
                  Logging("Need @t='l' for:" & vbTab & strLemma & strPos)
                End If
              End If
              ' Make sure all entries have a @t field set
              For intJ = 0 To dtrFound.Length - 1
                With dtrFound(intJ)
                  If (.Item("t").ToString = "") Then .Item("t") = "d" : bChanged = True : intNum += 1
                  If (.Item("l").ToString <> strLemma) Then .Item("l") = strLemma : bChanged = True : intNum += 1
                End With
              Next intJ
              ' Advance [intI]
              intI += dtrFound.Length - 1
            End If
          Next intI
          ' Need saving?
          If (bChanged) Then
            ' Save the result in morphdict
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
          End If

      End Select
      ' Step #3: Remove hyphens from @Vern field in [Morph]
      Select Case MsgBox("B&T adaptation for OE step #3: remove hyphens from @Vern field in [Morph]?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' Exit
          Status("Aborted at #3") : bInterrupt = True : Return False
        Case MsgBoxResult.No
          ' Skip this step
        Case MsgBoxResult.Yes
          ' Sort [Morph] on the @l field
          dtrMorph = tdlMorphDict.Tables("Morph").Select("", "l ASC")
          strLemma = "" : bChanged = False : strPosMain = "" : strPosThis = ""
          intNum = 0
          For intI = 0 To dtrMorph.Length - 1
            ' Where are we?
            intPtc = (intI + 1) * 100 \ dtrMorph.Length
            Status("MorphDictAdaptBT [" & Left(strLemma, 3) & "] " & intPtc & "%", intPtc)
            With dtrMorph(intI)
              ' Check for hyphen
              strVern = .Item("Vern").ToString
              If (InStr(strVern, "-") > 0) Then
                strVern = strVern.Replace("-", "")
                .Item("Vern") = strVern
                intNum += 1 : bChanged = True
              End If
            End With
          Next intI
          ' Report the number of changes
          Logging("MorphDictAdaptBT step #3: " & intNum & " changes made")
          ' Need saving?
          If (bChanged) Then
            ' Save the result in morphdict
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
          End If
      End Select

      ' Step #4: try to link @t=l records to BTlist
      Select Case MsgBox("B&T adaptation for OE step #4: link @t=l records to BTlist in [Morph]?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' Exit
          Status("Aborted at #4") : bInterrupt = True : Return False
        Case MsgBoxResult.No
          ' Skip this step
        Case MsgBoxResult.Yes
          ' Sort [Morph] on the @l field
          dtrMorph = tdlMorphDict.Tables("Morph").Select("t='l'", "l ASC")
          strLemma = "" : bChanged = False : strPosMain = "" : strPosThis = ""
          intNum = 0 : colEntry.Clear()
          strFile = GetLocalDir() & "\BTstep4.txt" : wrThis = New IO.StreamWriter(strFile)
          For intI = 0 To dtrMorph.Length - 1
            ' Get the lemma
            strLemma = dtrMorph(intI).Item("l").ToString : strLemmaOld = strLemma.Replace("-", "")
            ' Where are we?
            intPtc = (intI + 1) * 100 \ dtrMorph.Length
            Status("MorphDictAdaptBT [" & Left(strLemma, 3) & "] " & intPtc & "%", intPtc)
            With dtrMorph(intI)
              ' Get the contents of the @F field
              strF = .Item("f").ToString
              If (Not Regex.IsMatch(strF, "BT\d+")) Then
                ' There is no BT number yet - 
                ' Get the POS
                strPos = .Item("Pos").ToString
                ' Get the main OS this belongs to
                strPosMain = MorphCatMain(strPos, True)
                ' Make sure that for this operation we only take the POS part after a possible + sign into account
                intPos = InStr(strPosMain, "+") : If (intPos > 0) Then strPosMain = Mid(strPosMain, intPos + 1)
                ' Make sure MD, HV and AX are mapped correctly
                If (DoLike(strPosMain, "HV|*+HV|MD|AX|BE")) Then strPosMain = "VB"
                If (strPosMain = "") Then
                  Stop
                Else
                  ' There may be some more obvious changes needed
                  ' Try get the lemma (but use the @b field and no hyphens)
                  dtrFound = tdlBT.Tables("BT").Select("b='" & strLemmaOld & "' AND ps='" & strPosMain & "'")
                  ' If nothing has been found, try finding an entry without preverb
                  If (dtrFound.Length = 0) AndAlso (DoLike(strPosMain, "VB|*+VB|HV|*+HV")) Then
                    If (MorphDelPreVb(strLemmaOld) <> "") Then
                      dtrFound = tdlBT.Tables("BT").Select("b='" & MorphDelPreVb(strLemmaOld) & "' AND ps='" & strPosMain & "'")
                    End If
                  End If
                  ' Found anything?
                  If (dtrFound.Length = 0) Then
                    ' Nothing has been found -- keep track of this
                    wrThis.Write(.Item("MorphId").ToString & vbTab & .Item("Vern").ToString & vbTab & .Item("Pos").ToString & vbTab & _
                            .Item("l").ToString & vbTab & .Item("f").ToString)
                    wrThis.WriteLine()
                  Else
                    ' Found the lemma: get the BT number
                    AddSemiStack(strF, "s=BT" & dtrFound(0).Item("bt").ToString)
                    .Item("f") = strF
                    intNum += 1 : bChanged = True
                  End If
                End If
                'Else
                '  Stop
              End If
            End With
          Next intI
          '' Show the changes
          'Logging("Missing entries are on the report page", False)
          'frmMain.wbReport.DocumentText = colEntry.Html
          'frmMain.TabControl1.SelectedTab = frmMain.tpReport
          wrThis.Close()
          Logging("MorphDictAdaptBT results are in: " & strFile)
          ' Report the number of changes
          Logging("MorphDictAdaptBT step #4: " & intNum & " changes made")
          ' Need saving?
          If (bChanged) Then
            ' Save the result in morphdict
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
          End If
      End Select

      ' Step #4b: propagate BT numbers inside [Morph]
      Select Case MsgBox("B&T adaptation for OE step #4b: try resolving BT numbers in [Morph]?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' Exit
          Status("Aborted at #4b") : bInterrupt = True : Return False
        Case MsgBoxResult.No
          ' Skip this step
        Case MsgBoxResult.Yes
          ' Find all the @t=l entries inside [Morph]
          dtrMorph = tdlMorphDict.Tables("Morph").Select("t='l'", "l ASC")
          strLemma = "" : bChanged = False : strPosMain = "" : strPosThis = ""
          intNum = 0 : colEntry.Clear()
          ' Walk these entries
          For intI = 0 To dtrMorph.Length - 1
            ' Get this lemma and BT number
            With dtrMorph(intI)
              strLemma = .Item("l").ToString : strPos = .Item("Pos").ToString
              strF = .Item("f").ToString : strMed = Regex.Match(strF, "BT\d+").Value
              ' Where are we?
              intPtc = (intI + 1) * 100 \ dtrMorph.Length
              Status("MorphDictAdaptBT [" & Left(strLemma, 3) & "] " & intPtc & "%", intPtc)
              ' Only process those that have a BT number
              If (strMed <> "") Then
                strPosMain = MorphCatMain(strPos, True)
                intPos = InStr(strPosMain, "+")
                If (intPos > 0) Then strPosMain = Mid(strPosMain, intPos + 1)
                strPosList = MorphCatList(strPosMain, True)
                ' Get all the entries in [Morph] with the same lemma
                '  dtrFound = tdlMorphDict.Tables("Morph").Select("l='" & strLemma & "' AND t<>'l'")
                dtrFound = tdlMorphDict.Tables("Morph").Select("l='" & strLemma & "'")
                ' Walk through all of them
                For intJ = 0 To dtrFound.Length - 1
                  ' Only those without a @t=l
                  If (dtrFound(intJ).Item("t").ToString <> "l") Then
                    ' Check this one's POS
                    strPosMain = MorphCatMain(dtrFound(intJ).Item("Pos").ToString, True)
                    intPos = InStr(strPosMain, "+")
                    If (intPos > 0) Then strPosMain = Mid(strPosMain, intPos + 1)
                    If (DoLike(strPosMain, strPosList)) AndAlso (Not Regex.IsMatch(dtrFound(intJ).Item("f").ToString, "BT\d+")) Then
                      ' Okay, process it
                      dtrFound(intJ).Item("l") = strLemma
                      strF = dtrFound(intJ).Item("f").ToString : AddSemiStack(strF, "s=" & strMed) : dtrFound(intJ).Item("f") = strF
                      ' Log this change
                      Logging(dtrFound(intJ).Item("Vern").ToString & vbTab & dtrFound(intJ).Item("Pos").ToString & vbTab & strLemma & vbTab & strMed)
                      ' And keep track
                      bChanged = True : intNum += 1
                    End If
                  End If
                Next intJ
              End If
            End With
          Next intI

          ' Possibly process changes
          If (bChanged) Then
            ' Save the result in morphdict
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
            Logging("Number of changes in OE step #4b: " & intNum)
            Status("ready")
          End If
      End Select

      ' Step #4c: propagate [VernPos] to [Morph]: (1) missing BT numbers in [Morph], (2) missing lemma's in [Morph]
      Select Case MsgBox("B&T adaptation for OE step #4c: propagate [VernPos] info to [Morph]?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' Exit
          Status("Aborted at #4c") : bInterrupt = True : Return False
        Case MsgBoxResult.No
          ' Skip this step
        Case MsgBoxResult.Yes
          ' Visit all entries in Morph
          dtrMorph = tdlMorphDict.Tables("Morph").Select("", "l ASC")
          ' Initialise
          strLemma = "" : bChanged = False : strPosMain = "" : strPosThis = "" : intNum = 0
          For intI = 0 To dtrMorph.Length - 1
            ' Get the lemma
            strLemma = dtrMorph(intI).Item("l").ToString
            ' Where are we?
            intPtc = (intI + 1) * 100 \ dtrMorph.Length
            Status("MorphDictAdaptBT [" & Left(strLemma, 3) & "] " & intPtc & "%", intPtc)
            With dtrMorph(intI)
              ' Get the vernacular
              strVern = .Item("Vern").ToString
              ' Get the contents of the @F field
              strF = .Item("f").ToString
              ' Get the BT number
              strMed = Regex.Match(strF, "BT\d+").Value
              ' Check if we already have the BT number
              If (strMed = "") Then
                ' Get the POS
                strPos = .Item("Pos").ToString
                strPosMain = MorphCatMain(strPos, True)
                If (strPos Like "NR*") Then strPosMain = "NR"
                intPos = InStr(strPosMain, "+")
                If (intPos > 0) Then strPosMain = Mid(strPosMain, intPos + 1)
                strPosList = MorphCatList(strPosMain, True)
                ' Look for all entries with the same lemma in [VernPos] for the MED number
                dtrVP = tdlMorphDict.Tables("VernPos").Select("l='" & strLemma & "'")
                For intJ = 0 To dtrVP.Length - 1
                  ' Check for the right POS
                  If (DoLike(dtrVP(intJ).Item("Pos"), strPosList)) Then
                    ' Check BT number
                    strMed = Regex.Match(dtrVP(intJ).Item("f").ToString, "BT\d+").Value
                    ' Have we got it?
                    If (strMed <> "") Then
                      ' Process it
                      AddSemiStack(strF, "s=" & strMed) : .Item("f") = strF
                      bChanged = True : intNum += 1
                      Exit For
                    End If
                  End If
                Next intJ
              End If
            End With
          Next intI
          ' Show what we have
          Logging("Step #4c, first one, changes: " & intNum)
          intNum = 0
          ' Visit all entries in [VernPos]
          dtrVP = tdlMorphDict.Tables("VernPos").Select("", "l ASC")
          For intI = 0 To dtrVP.Length - 1
            ' Acces this one
            With dtrVP(intI)
              ' Get the contents of the @F field
              strF = .Item("f").ToString
              ' Get the BT number
              strMed = Regex.Match(strF, "BT\d+").Value
              ' Does this one have it?
              If (strMed = "") Then
                ' Get the lemma
                strLemma = .Item("l").ToString
                '  If (strLemma Like "gescieppan") Then Stop
                ' Where are we?
                intPtc = (intI + 1) * 100 \ dtrVP.Length
                Status("MorphDictAdaptBT [" & Left(strLemma, 3) & "] " & intPtc & "%", intPtc)
                ' Get the POS
                strPos = .Item("Pos").ToString
                strPosMain = MorphCatMain(strPos, True)
                If (strPos Like "NR*") Then strPosMain = "NR"
                intPos = InStr(strPosMain, "+")
                If (intPos > 0) Then strPosMain = Mid(strPosMain, intPos + 1)
                strPosList = MorphCatList(strPosMain, True)
                ' Find the lemma in [Morph]
                dtrMorph = tdlMorphDict.Tables("Morph").Select("Vern='" & strLemma & "' AND l='" & strLemma & "'") : bCorrect = False
                For intJ = 0 To dtrMorph.Length - 1
                  ' Check for the right POS
                  If (DoLike(dtrMorph(intJ).Item("Pos"), strPosList)) Then
                    ' The lemma occurs somewhere in [Morph]
                    bCorrect = True
                  End If
                Next intJ
                ' Is the lemma in [Morph] already?
                If (Not bCorrect) AndAlso (strPosMain <> "NR") AndAlso (strPosMain <> "pronoun") Then
                  ' Add the lemma to [Morph] as a lemma
                  If (Not MorphDictAdd(strLemma, strPosMain, "", strLemma, "", bIsLemma:=True)) Then Return False
                  bChanged = True : intNum += 1
                End If
              End If
            End With
            ' 
          Next intI
          ' Show what we have
          Logging("Step #4c, second one, changes: " & intNum)

          If (bChanged) Then
            ' Save the result in morphdict
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
            Status("ready")
          End If
      End Select

      ' Step #5a: try adding BT numbers to items in VernPos
      Select Case MsgBox("B&T adaptation for OE step #5a: try resolving BT numbers in [VernPos]?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' Exit
          Status("Aborted at #5a") : bInterrupt = True : Return False
        Case MsgBoxResult.No
          ' Skip this step
        Case MsgBoxResult.Yes
          If (Not MorphResolveVernPosBT("OE", bChanged)) Then
            Logging("Could not resolve VernPos/BT") : Return False
          End If
          If (bChanged) Then
            ' Save the result in morphdict
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
            Status("ready")
          End If
      End Select

      ' Step #5b: link the BT numbers in @t='l' lemma's in [Morph] to correct entries in [VernPos]
      Select Case MsgBox("B&T adaptation for OE step #5b: link BT numbers from [Morph] to [VernPos]?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' Exit
          Status("Aborted at #5b") : bInterrupt = True : Return False
        Case MsgBoxResult.No
          ' Skip this step
        Case MsgBoxResult.Yes
          ' Visit all @t=l entries in Morph
          dtrMorph = tdlMorphDict.Tables("Morph").Select("", "l ASC")
          ' Initialise
          strLemma = "" : bChanged = False : strPosMain = "" : strPosThis = "" : intNum = 0
          For intI = 0 To dtrMorph.Length - 1
            ' Get the lemma
            strLemma = dtrMorph(intI).Item("l").ToString
            ' Where are we?
            intPtc = (intI + 1) * 100 \ dtrMorph.Length
            Status("MorphDictAdaptBT [" & Left(strLemma, 3) & "] " & intPtc & "%", intPtc)
            With dtrMorph(intI)
              ' Get the vernacular
              strVern = .Item("Vern").ToString
              ' Get the contents of the @F field
              strF = .Item("f").ToString
              If (Regex.IsMatch(strF, "BT\d+")) AndAlso (Not DoLike(strVern, "on|ut|part|part.|participle|past|pl|sg")) Then
                ' There is a BT number, so we can 'propagate' it
                ' (a) Get the POS
                strPos = .Item("Pos").ToString
                ' (b) Get the BT number
                strMed = Regex.Match(strF, "BT\d+").Value
                ' (c) Get the main POS this belongs to 
                Select Case strPos
                  Case "BE"
                    strPosList = "BE*|BA*|*+BE*|*+BA*"
                  Case "MD"
                    strPosList = "MD*|*+MD*"
                  Case "VB"
                    strPosList = "VB*|VA*|*+VB*|*+VA*"
                  Case Else
                    strPosMain = MorphCatMain(strPos, True)
                    intPos = InStr(strPosMain, "+")
                    If (intPos > 0) Then strPosMain = Mid(strPosMain, intPos + 1)
                    strPosList = MorphCatList(strPosMain, True)
                End Select
                If (strPosList <> "") Then
                  ' Get all candicates from VernPos
                  '  dtrVP = tdlMorphDict.Tables("VernPos").Select("l='" & strLemma & "' OR Vern='" & strVern & "'")
                  dtrVP = tdlMorphDict.Tables("VernPos").Select("Vern='" & strVern & "'")
                  For intJ = 0 To dtrVP.Length - 1
                    ' Get the correct POS to look at
                    strPos = dtrVP(intJ).Item("Pos").ToString
                    strPosMain = MorphCatMain(strPos, True)
                    intPos = InStr(strPosMain, "+")
                    If (intPos > 0) Then strPosMain = Mid(strPosMain, intPos + 1)
                    ' Check the pos
                    If (DoLike(strPosMain, strPosList)) Then
                      ' This has the correct POS, so continue
                      If (Not Regex.IsMatch(dtrVP(intJ).Item("f").ToString, "BT\d+")) Then
                        ' ========= Show what we are doing =============
                        Logging(strVern & vbTab & strPos & vbTab & strLemma & vbTab & strMed, False)
                        ' This one needs to get the correct BT number
                        strF = dtrVP(intJ).Item("f").ToString
                        AddSemiStack(strF, "s=" & strMed)
                        dtrVP(intJ).Item("f") = strF
                        ' Keep track of numbers
                        bChanged = True
                        intNum += 1
                      End If
                    End If
                  Next intJ
                  ' if there is a hyphen, there possibly is a problem
                  If (InStr(strLemma, "-") > 0) Then
                    ' Get all candicates from VernPos, but now without hyphen
                    '  dtrVP = tdlMorphDict.Tables("VernPos").Select("l='" & strLemma.Replace("-", "") & "' OR Vern='" & strVern & "'")
                    dtrVP = tdlMorphDict.Tables("VernPos").Select("Vern='" & strVern & "'")
                    For intJ = 0 To dtrVP.Length - 1
                      ' Get the correct POS to look at
                      strPos = dtrVP(intJ).Item("Pos").ToString
                      strPosMain = MorphCatMain(strPos, True)
                      intPos = InStr(strPosMain, "+")
                      If (intPos > 0) Then strPosMain = Mid(strPosMain, intPos + 1)
                      ' Check the pos
                      If (DoLike(strPosMain, strPosList)) Then
                        ' This has the correct POS, so continue
                        If (Not Regex.IsMatch(dtrVP(intJ).Item("f").ToString, "BT\d+")) Then
                          ' ========= Show what we are doing =============
                          Logging(strVern & vbTab & strPos & vbTab & strLemma & vbTab & strMed, False)
                          ' This one needs to get the correct BT number
                          strF = dtrVP(intJ).Item("f").ToString
                          AddSemiStack(strF, "s=" & strMed)
                          dtrVP(intJ).Item("f") = strF
                          ' Also change the lemma to the one with hyphen
                          dtrVP(intJ).Item("l") = strLemma
                          ' Keep track of numbers
                          bChanged = True
                          intNum += 1
                        End If
                      End If
                    Next intJ

                  End If
                End If
              End If
            End With
          Next intI

          If (bChanged) Then
            ' Give the number of changes
            Logging("Changes to VP: " & intNum)
            ' Save the result in morphdict
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
            Status("ready")
          End If
      End Select

      ' Step #5c: try adding BT numbers to items in VernPos
      Select Case MsgBox("B&T adaptation for OE step #5c: replace orphan lemma's in [VernPos] with correct ones from [Morph]?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' Exit
          Status("Aborted at #5c") : bInterrupt = True : Return False
        Case MsgBoxResult.No
          ' Skip this step
        Case MsgBoxResult.Yes
          ' Initialise
          bChanged = False : intNum = 0
          ' Walk through [VernPos] in the correct order
          dtrVP = tdlMorphDict.Tables("VernPos").Select("", "l ASC")
          For intI = 0 To dtrVP.Length - 1
            With dtrVP(intI)
              ' Get necessary information
              strF = .Item("f").ToString
              If (Not Regex.IsMatch(strF, "BT\d+")) Then
                ' This entry does NOT have a BT number yet - get lemma
                strLemmaOld = .Item("l").ToString
                ' Show where we are
                intPtc = (intI + 1) * 100 \ dtrVP.Length
                Status("MorphDictAdaptBT [" & Left(strLemmaOld, 3) & "] " & intPtc & "%", intPtc)
                ' ========= DEBUG ==========
                ' If (strLemmaOld = "abirian") Then Stop
                ' ==========================
                ' Get the major POS we are looking for
                strPos = .Item("Pos").ToString : strPosMain = MorphCatMain(strPos, True) : intPos = InStr(strPosMain, "+")
                ' Look for entry in Morph
                dtrFound = tdlMorphDict.Tables("Morph").Select("Vern='" & strLemmaOld & "' AND Pos='" & strPosMain & "'")
                If (dtrFound.Length = 0) AndAlso (intPos > 0) Then
                  ' Try without "+"
                  strPosMain = Mid(strPosMain, intPos + 1)
                  dtrFound = tdlMorphDict.Tables("Morph").Select("Vern='" & strLemmaOld & "' AND Pos='" & strPosMain & "'")
                End If
                ' Found anything?
                If (dtrFound.Length > 0) Then
                  ' Get the BT number
                  strF = dtrFound(0).Item("f").ToString : strMed = Regex.Match(strF, "BT\d+").Value
                  If (strMed <> "") Then
                    ' Okay, we have a BT number --> now get the new lemma
                    strLemma = dtrFound(0).Item("l").ToString
                    ' Find and change all that is needed in [Morph] and in [VernPos]
                    If (Not MorphDictAdaptOneLemmaBT("Morph", strLemmaOld, strLemma, strMed, strPosMain, bChanged, intNum)) Then Return False
                    If (Not MorphDictAdaptOneLemmaBT("VernPos", strLemmaOld, strLemma, strMed, strPosMain, bChanged, intNum)) Then Return False
                    ' De we need to try a variant RP+VB?
                    If (strPosMain = "VB") Then
                      If (Not MorphDictAdaptOneLemmaBT("Morph", strLemmaOld, strLemma, strMed, "RP+VB", bChanged, intNum)) Then Return False
                      If (Not MorphDictAdaptOneLemmaBT("VernPos", strLemmaOld, strLemma, strMed, "RP+VB", bChanged, intNum)) Then Return False
                    End If
                  Else
                    ' No lemma has been found, so we need to add the lemma to [Morph] for later resolution
                    'If (strPosMain = "VB") Then Stop
                    If (Not MorphDictAdd(strLemmaOld, strPosMain, "", strLemmaOld, "", bIsLemma:=True)) Then Return False
                    bChanged = True : intNum += 1
                  End If
                End If
              End If
            End With
          Next intI

          If (bChanged) Then
            ' Indicate the number of changes
            Logging("Number of changes = " & intNum)
            ' Save the result in morphdict
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
            Status("ready")
          End If
      End Select


      ' Step #5d: Adapt BT numbers for RP+VB
      Select Case MsgBox("B&T adaptation for OE step #5d: adapt BT numbers for RP+VB in [Morph]?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' Exit
          Status("Aborted at #5d") : bInterrupt = True : Return False
        Case MsgBoxResult.No
          ' Skip this step
        Case MsgBoxResult.Yes
          ' Initialise
          bChanged = False : intNum = 0
          ' Walk through Morph
          dtrMorph = tdlMorphDict.Tables("Morph").Select("Pos='RP+VB'", "l ASC")
          For intI = 0 To dtrMorph.Length - 1
            intPtc = (intI + 1) * 100 \ dtrMorph.Length
            Status("Step 5d " & intPtc & "%", intPtc)
            With dtrMorph(intI)
              ' Access the @f attribute
              strF = .Item("f").ToString
              ' Does this have a BT number?
              If (Not Regex.IsMatch(strF, "BT\d+")) Then
                ' This one does not have a BT number -- check if there is an alternative one that has
                strLemma = .Item("l").ToString
                ' ============== DEBUG ===================
                '  If (strLemma.Replace("-", "") = "ablinnan") Then Stop
                ' ========================================
                dtrFound = tdlMorphDict.Tables("Morph").Select("Pos='VB' AND l='" & strLemma & "' AND t='l'")
                If (dtrFound.Length = 0) AndAlso (InStr(strLemma, "-") > 0) Then
                  dtrFound = tdlMorphDict.Tables("Morph").Select("Pos='VB' AND l='" & strLemma.Replace("-", "") & "' AND t='l'")
                  If (dtrFound.Length = 0) AndAlso (InStr(strLemma, "-") > 0) Then
                    dtrFound = tdlMorphDict.Tables("Morph").Select("Pos='VB' AND l='" & strLemma.Replace("-", "") & "'")
                    If (dtrFound.Length = 0) AndAlso (InStr(strLemma, "-") > 0) Then
                      dtrFound = tdlMorphDict.Tables("Morph").Select("Pos='VB' AND l='" & strLemma & "'")
                    End If
                  End If
                End If
                ' Any result?
                If (dtrFound.Length > 0) Then
                  ' Take over the information to the entry we are looking at
                  strMed = Regex.Match(dtrFound(0).Item("f").ToString, "BT\d+").Value
                  If (strMed <> "") Then
                    AddSemiStack(strF, "s=" & strMed)
                    .Item("f") = strF : bChanged = True : intNum += 1
                    ' Pass these changes on into VernPos
                    If (Not MorphDictAdaptOneLemmaBT("VernPos", strLemma, strLemma, strMed, "RP+VB", bChanged, intNum)) Then Return False
                  End If
                End If
              End If
            End With
          Next intI

          If (bChanged) Then
            ' Indicate the number of changes
            Logging("Number of changes in #5d = " & intNum)
            ' Save the result in morphdict
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
            Status("ready")
          End If
      End Select

      ' Step #5e: try adding BT numbers to items in VernPos
      Select Case MsgBox("B&T adaptation for OE step #5e: try add BT numbers in [VernPos] where possible?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' Exit
          Status("Aborted at #5e") : bInterrupt = True : Return False
        Case MsgBoxResult.No
          ' Skip this step
        Case MsgBoxResult.Yes
          ' Initialise
          bChanged = False : intNum = 0
          ' Walk through [VernPos] in the correct order
          dtrVP = tdlMorphDict.Tables("VernPos").Select("", "l ASC, f DESC") : strLemma = ""
          For intI = 0 To dtrVP.Length - 1
            With dtrVP(intI)
              ' Get necessary information
              strF = .Item("f").ToString
              If (Not Regex.IsMatch(strF, "BT\d+")) Then
                ' This entry does NOT have a BT number yet - get lemma
                strLemmaOld = .Item("l").ToString
                ' Get the major POS we are looking for
                strPos = .Item("Pos").ToString : strPosMain = MorphCatMain(strPos, True) : intPos = InStr(strPosMain, "+")
                ' Show where we are
                intPtc = (intI + 1) * 100 \ dtrVP.Length
                Status("MorphDictAdaptBT step #5e [" & Left(strLemmaOld, 3) & "] " & intPtc & "%", intPtc)
                ' ========= DEBUG ==========
                '  If (strLemmaOld.Replace("-", "") = "onhweorfan") OrElse (strPos Like "V*") Then Stop
                ' ==========================
                ' Try find a BT number from the entries in [VernPos] with this lemma
                If (Not MorphDictAdaptFindBT("VernPos", strLemmaOld, strPosMain, strLemma, strMed)) Then Return False
                ' if failed, try Morph
                If (strMed = "") AndAlso (InStr(strLemmaOld, ";") = 0) Then
                  '  If (InStr(strLemmaOld, "-") > 0) Then Stop
                  If (Not MorphDictAdaptFindBT("Morph", strLemmaOld, strPosMain, strLemma, strMed)) Then Return False
                  If (strLemmaOld.Contains("-")) Then
                    ' Try entry without hyphen
                    strLemmaOld = strLemmaOld.Replace("-", "")
                    If (Not MorphDictAdaptFindBT("VernPos", strLemmaOld, strPosMain, strLemma, strMed)) Then Return False
                    If (strMed = "") Then
                      If (Not MorphDictAdaptFindBT("Morph", strLemmaOld, strPosMain, strLemma, strMed)) Then Return False
                      ' If nothing has been found, try finding an entry without preverb
                      If (strMed = "") AndAlso (DoLike(strPosMain, "VB|RP+VB")) Then
                        If (MorphDelPreVb(strLemmaOld) <> "") Then
                          If (Not MorphDictAdaptFindBT("Morph", MorphDelPreVb(strLemmaOld), strPosMain, strLemma, strMed)) Then Return False
                        End If
                      End If
                    End If
                  End If
                  ' if we didn't find anything yet, and this is a verb....
                  If (strMed = "") AndAlso (DoLike(strPosMain, strAnyVerb)) Then
                    ' If this is a AX, HV, BE, MD verb, try looking at [VB]  for a POS
                    If (DoLike(strPos, "AX*|HV*|BE*|MD*")) Then
                      ' Look at [VB]
                      If (Not MorphDictAdaptFindBT("Morph", strLemmaOld, "VB" & Mid(strPosMain, 3), strLemma, strMed)) Then Return False
                    End If
                    ' If this is a BA verb, try looking at [VA]  for a POS
                    If (DoLike(strPos, "BA*")) Then
                      ' Look at [VB]
                      If (Not MorphDictAdaptFindBT("Morph", strLemmaOld, "VA" & Mid(strPosMain, 3), strLemma, strMed)) Then Return False
                    End If
                  End If
                  If (strMed = "") AndAlso (DoLike(strPosMain, strAnyVerb)) Then
                    Select Case Left(strPos, 2)
                      Case "BE", "BA"
                        If (DoLike(strLemmaOld, "w*r*t*")) Then
                          If (Not MorphDictAdaptFindBT("Morph", strLemmaOld, "VB", "weorthan", strMed)) Then Return False
                        Else
                          If (Not MorphDictAdaptFindBT("Morph", strLemmaOld, "BE", "beon", strMed)) Then Return False
                        End If
                      Case "AX"
                        Stop
                      Case "MD"
                        '   Stop
                        If (DoLike(strLemmaOld, "th*r*f*")) Then
                          If (Not MorphDictAdaptFindBT("Morph", strLemmaOld, "VB", "thearfan", strMed)) Then Return False
                        ElseIf (DoLike(strLemmaOld, "sc*l*")) Then
                          If (Not MorphDictAdaptFindBT("Morph", strLemmaOld, "VB", "scellan", strMed)) Then Return False
                        Else
                          If (Not MorphDictAdaptFindBT("Morph", strLemmaOld, "BE", "beon", strMed)) Then Return False
                        End If
                      Case "HV"
                        Stop
                    End Select
                  End If
                End If
                ' Any results?
                If (strMed <> "") Then
                  ' Adapt myself
                  AddSemiStack(strF, "s=" & strMed) : .Item("f") = strF
                  bChanged = True : intNum += 1
                End If
              End If
            End With
          Next intI
          If (bChanged) Then
            ' Indicate the number of changes
            Logging("Number of changes = " & intNum)
            ' Save the result in morphdict
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
            Status("ready")
          End If
      End Select

      ' Step #5f: make room for missing lemma's from [VernPos] into [Morph]
      Select Case MsgBox("B&T adaptation for OE step #5f: add missing lemma's from [VernPos] to [Morph]?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' Exit
          Status("Aborted at #5f") : bInterrupt = True : Return False
        Case MsgBoxResult.No
          ' Skip this step
        Case MsgBoxResult.Yes
          ' Initialise
          bChanged = False : intNum = 0
          ' Walk through [VernPos] in the correct order
          dtrVP = tdlMorphDict.Tables("VernPos").Select("", "l ASC, f DESC") : strLemma = ""
          For intI = 0 To dtrVP.Length - 1
            With dtrVP(intI)
              ' Get necessary information
              strF = .Item("f").ToString
              If (Not Regex.IsMatch(strF, "BT\d+")) Then
                ' This entry does NOT have a BT number yet - get lemma
                strLemmaOld = .Item("l").ToString
                ' Do we need to process this lemma any further?
                If (strLemmaOld.Contains("=")) Then
                  Dim arThis() As String
                  ' Format: [2;acwelan|sfx=lth|frq=19;acwellan|sfx=lth|frq=3]
                  arThis = Split(strLemmaOld, ";")
                  If (arThis.Length > 1) Then
                    arThis = Split(arThis(1), "|")
                    If (arThis.Length > 1) Then
                      strLemmaOld = arThis(0)
                      ' Also change this at the entry
                      .Item("l") = strLemmaOld
                    End If
                  End If
                End If
                ' Get the major POS we are looking for
                strPos = .Item("Pos").ToString : strPosMain = MorphCatMain(strPos, True) : intPos = InStr(strPosMain, "+")
                ' Adaptation...
                If (strPos Like "NR*") AndAlso (strPosMain <> "NR") Then strPosMain = "NR"
                ' Show where we are
                intPtc = (intI + 1) * 100 \ dtrVP.Length
                Status("MorphDictAdaptBT step #5f [" & Left(strLemmaOld, 3) & "] " & intPtc & "%", intPtc)
                ' ========= DEBUG ==========
                '  If (strLemmaOld.Replace("-", "") = "onhweorfan") OrElse (strPos Like "V*") Then Stop
                ' ==========================
                ' Skip the NR entries
                If (strPosMain = "NR") Then
                  ' SKip it...
                Else
                  ' Try find the lemma in [Morph]
                  dtrMorph = tdlMorphDict.Tables("Morph").Select("l='" & strLemmaOld & _
                                "' AND Pos='" & strPosMain & "'")
                  If (dtrMorph.Length = 0) Then
                    If (strLemmaOld.Contains("-")) Then
                      ' Try finding without hyphen
                      dtrMorph = tdlMorphDict.Tables("Morph").Select("l='" & _
                          strLemmaOld.Replace("-", "") & "' AND Pos='" & strPosMain & "'")
                    End If
                    If (dtrMorph.Length = 0) AndAlso (intPos > 0) Then
                      ' Look for a lemma with slightly different POS
                      dtrMorph = tdlMorphDict.Tables("Morph").Select("l='" & strLemmaOld & _
                                  "' AND Pos='" & Mid(strPosMain, intPos + 1) & "'")
                    End If
                    If (dtrMorph.Length = 0) AndAlso (intPos > 0) AndAlso (strLemmaOld.Contains("-")) Then
                      ' Look for a lemma with slightly different POS
                      dtrMorph = tdlMorphDict.Tables("Morph").Select("l='" & _
                        strLemmaOld.Replace("-", "") & "' AND Pos='" & Mid(strPosMain, intPos + 1) & "'")
                    End If
                  End If
                  ' Any results?
                  If (dtrMorph.Length = 0) Then
                    ' There is no result, so add lemma into [Morph]
                    MorphDictAdd(strLemmaOld, strPosMain, "", strLemmaOld, strF)
                  Else
                    ' There is a result, so process it!
                    strMed = Regex.Match(dtrMorph(0).Item("f").ToString, "BT\d+").Value
                    If (strMed <> "") Then
                      ' Yes, got it!!
                      AddSemiStack(strF, "s=" & strMed)
                      .Item("f") = strF : bChanged = True : intNum += 1
                    End If
                  End If

                End If
              End If
            End With
          Next intI

          If (bChanged) Then
            ' Indicate the number of changes
            Logging("Number of changes = " & intNum)
            ' Save the result in morphdict
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
            Status("ready")
          End If
      End Select


      ' Step #6a: make a list of unlinked lemma's in [Morph]
      Select Case MsgBox("B&T adaptation for OE step #6a: list of unlinked lemma's in [Morph]?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' Exit
          Status("Aborted at #6a") : bInterrupt = True : Return False
        Case MsgBoxResult.No
          ' Skip this step
        Case MsgBoxResult.Yes
          ' Sort [Morph] on the @l field
          dtrMorph = tdlMorphDict.Tables("Morph").Select("t='l'", "l ASC")
          strLemma = "" : bChanged = False : strPosMain = "" : strPosThis = ""
          intNum = 0 : colEntry.Clear()
          strFile = GetLocalDir() & "\BTstep6a.txt" : wrThis = New IO.StreamWriter(strFile)
          For intI = 0 To dtrMorph.Length - 1
            ' Get the lemma
            strLemma = dtrMorph(intI).Item("l").ToString
            ' ======= DEBUG ============
            '  If (strLemma = "fyrdhrian") Then Stop
            ' ==========================
            ' Where are we?
            intPtc = (intI + 1) * 100 \ dtrMorph.Length
            Status("MorphDictAdaptBT [" & Left(strLemma, 3) & "] " & intPtc & "%", intPtc)
            With dtrMorph(intI)
              ' Get the vernacular
              strVern = .Item("Vern").ToString
              ' Get the contents of the @F field
              strF = .Item("f").ToString
              ' (a) Get the POS
              strPos = .Item("Pos").ToString
              If (Not Regex.IsMatch(strF, "BT\d+")) AndAlso (DoLike(strPos, strAnyVerb)) Then
                ' There is no BT number yet - we are going to try and predict it...
                ' (b) Get the main POS this belongs to (the "ps" field used in the BT list
                strPosMain = MorphCatMain(strPos, True)
                If (DoLike(strPos, strAnyVerb)) Then strPosMain = "VB"
                If (strPosMain = "") Then
                  ' Safeguarding bad data
                  Stop
                Else
                  ' Look at the BT list, but only at those that have the correct PS field
                  dtrFound = tdlBT.Tables("BT").Select("ps='" & strPosMain & "'")
                  ' ======= DEBUG ============
                  '  If (strPosMain = "VB") Then Stop
                  ' ==========================
                  ' Initialisations
                  intBest = -1 : intConf = -1 : strBest = "" : strForm = ""
                  ' Compare all the possible lemma's with myself
                  For intJ = 0 To dtrFound.Length - 1
                    ' Compare this form with [strVern]
                    intValue = objSim.LevenDist(dtrFound(intJ).Item("b").ToString, strLemma, intEquals)
                    ' Adjust value for initial letter mismatch
                    If (Left(dtrFound(intJ).Item("b").ToString, 1) <> Left(strLemma, 1)) Then intValue -= 3
                    ' Beware of negatives
                    If (intValue < 0) Then intValue = 0
                    ' Process best value
                    If (intValue > intConf) Then
                      intConf = intValue : strBest = dtrFound(intJ).Item("b").ToString
                      strForm = "BT" & Format(CInt(dtrFound(intJ).Item("bt").ToString), "000000")
                    End If
                  Next intJ
                  If (InStr(strLemma, "dh") > 0) Then
                    ' Try replacing DH with TH
                    strLemmaOld = strLemma.Replace("dh", "th")
                    For intJ = 0 To dtrFound.Length - 1
                      ' Compare this form with [strVern]
                      intValue = objSim.LevenDist(dtrFound(intJ).Item("b").ToString, strLemmaOld, intEquals)
                      ' Adjust value for initial letter mismatch
                      If (Left(dtrFound(intJ).Item("b").ToString, 1) <> Left(strLemmaOld, 1)) Then intValue -= 3
                      ' Beware of negatives
                      If (intValue < 0) Then intValue = 0
                      ' Process best value
                      If (intValue > intConf) Then
                        intConf = intValue : strBest = dtrFound(intJ).Item("b").ToString
                        strForm = "BT" & Format(CInt(dtrFound(intJ).Item("bt").ToString), "000000")
                      End If
                    Next intJ
                  End If
                  ' if we haven't found anything, and this is a verb, then try form without pre-verb
                  If (intConf <= 0) AndAlso (strPosMain = "VB") AndAlso (MorphDelPreVb(strLemma) <> "") Then
                    strLemmaOld = MorphDelPreVb(strLemma)
                    For intJ = 0 To dtrFound.Length - 1
                      ' Compare this form with [strVern]
                      intValue = objSim.LevenDist(dtrFound(intJ).Item("b").ToString, strLemmaOld, intEquals)
                      ' Adjust value for initial letter mismatch
                      If (Left(dtrFound(intJ).Item("b").ToString, 1) <> Left(strLemmaOld, 1)) Then intValue -= 3
                      ' Beware of negatives
                      If (intValue < 0) Then intValue = 0
                      ' Process best value
                      If (intValue > intConf) Then
                        intConf = intValue : strBest = dtrFound(intJ).Item("b").ToString
                        strForm = "BT" & Format(CInt(dtrFound(intJ).Item("bt").ToString), "000000")
                      End If
                    Next intJ
                  End If
                  ' if we haven't found anything, and this is a verb, then try form starting from hyphen
                  If (strLemma.Contains("-")) Then
                    strLemmaOld = Mid(strLemma, InStr(strLemma, "-") + 1)
                    For intJ = 0 To dtrFound.Length - 1
                      ' Compare this form with [strVern]
                      intValue = objSim.LevenDist(dtrFound(intJ).Item("b").ToString, strLemmaOld, intEquals)
                      ' Adjust value for initial letter mismatch
                      If (Left(dtrFound(intJ).Item("b").ToString, 1) <> Left(strLemmaOld, 1)) Then intValue -= 3
                      ' Beware of negatives
                      If (intValue < 0) Then intValue = 0
                      ' Process best value
                      If (intValue > intConf) Then
                        intConf = intValue : strBest = Left(strLemma, InStr(strLemma, "-")) & dtrFound(intJ).Item("b").ToString
                        strForm = "BT" & Format(CInt(dtrFound(intJ).Item("bt").ToString), "000000")
                      End If
                    Next intJ
                  End If
                  ' Also compare with existing lemma's in [Morph], for which the BT number is known
                  dtrFound = tdlMorphDict.Tables("Morph").Select("Pos='" & strPosMain & "' AND t='l'")
                  For intJ = 0 To dtrFound.Length - 1
                    strMed = Regex.Match(dtrFound(intJ).Item("f").ToString, "BT\d+").Value
                    If (strMed <> "") Then
                      ' Get potential lemma
                      strLemmaOld = dtrFound(intJ).Item("l").ToString : strLemmaOld = strLemmaOld.Replace("-", "")
                      ' Compare lemma's
                      intValue = objSim.LevenDist(strLemmaOld, strLemma, intEquals)
                      ' Adjust value for initial letter mismatch
                      If (Left(strLemmaOld, 1) <> Left(strLemma, 1)) Then intValue -= 3
                      ' Beware of negatives
                      If (intValue < 0) Then intValue = 0
                      ' Process best value
                      If (intValue > intConf) Then
                        intConf = intValue : strBest = strLemmaOld
                        strForm = strMed
                      End If
                    End If
                  Next intJ
                  ' Add more for RP+VB
                  If (strPosMain = "VB") Then
                    dtrFound = tdlMorphDict.Tables("Morph").Select("Pos='RP+VB' AND t='l'")
                    For intJ = 0 To dtrFound.Length - 1
                      strMed = Regex.Match(dtrFound(intJ).Item("f").ToString, "BT\d+").Value
                      If (strMed <> "") Then
                        ' Get potential lemma
                        strLemmaOld = dtrFound(intJ).Item("l").ToString : strLemmaOld = strLemmaOld.Replace("-", "")
                        ' Compare lemma's
                        intValue = objSim.LevenDist(strLemmaOld, strLemma, intEquals)
                        ' Adjust value for initial letter mismatch
                        If (Left(strLemmaOld, 1) <> Left(strLemma, 1)) Then intValue -= 3
                        ' Beware of negatives
                        If (intValue < 0) Then intValue = 0
                        ' Process best value
                        If (intValue > intConf) Then
                          intConf = intValue : strBest = strLemmaOld
                          strForm = strMed
                        End If
                      End If
                    Next intJ
                  End If

                  ' Correction for bad ones
                  If (intConf <= 0) Then
                    strBest = "-"
                    'Else
                    '  Stop
                  Else
                    ' Look for ambiguous lemma's
                    dtrFound = tdlBT.Tables("BT").Select("ps='" & strPosMain & "' AND b='" & strBest & "'")
                    ' Add any additional lemma's
                    For intJ = 1 To dtrFound.Length - 1
                      ' Add the BT number; the @b feature is the same
                      AddSemiStack(strForm, "BT" & Format(CInt(dtrFound(intJ).Item("bt").ToString), "000000"), strDivider:="|")
                    Next intJ
                  End If
                  ' Look for examples of Vern
                  dtrVP = tdlMorphDict.Tables("VernPos").Select("l='" & strLemma & "'")
                  If (dtrVP.Length = 0) Then
                    strExmp = "(none)" : strHasVP = "M"
                  Else
                    strExmp = "" : strHasVP = "M_VP"
                    For intJ = 0 To dtrVP.Length - 1
                      ' =============== Double checking =======
                      '  If (dtrVP(intJ).Item("Pos").ToString Like "*+*") Then Stop
                      ' =======================================
                      ' Check that the POS is fitting
                      If (DoLike(MorphCatMain(dtrVP(intJ).Item("Pos").ToString), strPos & "|*+" & strPos)) Then
                        ' Add it to the examples
                        AddSemiStack(strExmp, dtrVP(intJ).Item("Vern").ToString, strDivider:="|")
                      End If
                    Next intJ
                  End If
                  ' Process this into a prediction
                  wrThis.Write(.Item("MorphId").ToString & vbTab & strVern & vbTab & .Item("Pos").ToString & vbTab & _
                          strLemma & vbTab & strF & vbTab & intConf & vbTab & strBest & vbTab & strHasVP & vbTab & strForm & _
                           vbTab & strExmp)
                  wrThis.WriteLine()
                  intNum += 1 : bChanged = False
                End If
              End If
            End With
          Next intI
          wrThis.Close()
          Logging("MorphDictAdaptBT results are in: " & strFile)
          ' Report the number of changes
          Logging("MorphDictAdaptBT step #6a: " & intNum & " BT numbers estimated")
          ' Need saving?
          If (bChanged) Then
            ' Save the result in morphdict
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
          End If
      End Select

      ' Step #7: make the attribute @t where it is not present yet
      Select Case MsgBox("B&T adaptation for OE step #6b: list of unlinked lemma's in [VernPos]?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' Exit
          Status("Aborted at #6b") : bInterrupt = True : Return False
        Case MsgBoxResult.No
          ' Skip this step
        Case MsgBoxResult.Yes
          ' Sort [Morph] on the @l field
          dtrMorph = tdlMorphDict.Tables("VernPos").Select("", "l ASC")
          strLemma = "" : bChanged = False : strPosMain = "" : strPosThis = ""
          intNum = 0 : colEntry.Clear()
          strFile = GetLocalDir() & "\BTstep6b.txt" : wrThis = New IO.StreamWriter(strFile)
          For intI = 0 To dtrMorph.Length - 1
            ' Get the lemma
            strLemma = dtrMorph(intI).Item("l").ToString
            ' Where are we?
            intPtc = (intI + 1) * 100 \ dtrMorph.Length
            Status("MorphDictAdaptBT [" & Left(strLemma, 3) & "] " & intPtc & "%", intPtc)
            With dtrMorph(intI)
              ' Get the vernacular
              strVern = .Item("Vern").ToString
              ' Get the contents of the @F field
              strF = .Item("f").ToString
              If (Not Regex.IsMatch(strF, "BT\d+")) Then
                ' There is no BT number yet - we are going to try and predict it...
                ' (a) Get the POS
                strPos = .Item("Pos").ToString
                ' (b) Get the main POS this belongs to (the "ps" field used in the BT list
                strPosMain = MorphCatMain(strPos, True)
                intPos = InStr(strPosMain, "+")
                If (intPos > 0) Then strPosMain = Mid(strPosMain, intPos + 1)
                If (strPosMain = "") Then
                  ' Safeguarding bad data
                  Stop
                Else
                  ' Look at the BT list, but only at those that have the correct PS field
                  dtrFound = tdlBT.Tables("BT").Select("ps='" & strPosMain & "'")
                  ' ======= DEBUG ============
                  '  If (strPosMain = "VB") Then Stop
                  ' ==========================
                  ' Initialisations
                  intBest = -1 : intConf = -1 : strBest = "" : strForm = ""
                  ' Compare all the possible lemma's with myself
                  For intJ = 0 To dtrFound.Length - 1
                    ' Compare this form with [strVern]
                    intValue = objSim.LevenDist(dtrFound(intJ).Item("b").ToString, strLemma, intEquals)
                    ' Adjust value for initial letter mismatch
                    If (Left(dtrFound(intJ).Item("b").ToString, 1) <> Left(strLemma, 1)) Then intValue -= 3
                    ' Beware of negatives
                    If (intValue < 0) Then intValue = 0
                    ' Process best value
                    If (intValue > intConf) Then
                      intConf = intValue : strBest = dtrFound(intJ).Item("b").ToString
                      strForm = "BT" & Format(CInt(dtrFound(intJ).Item("bt").ToString), "000000")
                    End If
                  Next intJ
                  If (InStr(strLemma, "dh") > 0) Then
                    ' Try replacing DH with TH
                    strLemma = strLemma.Replace("dh", "th")
                    For intJ = 0 To dtrFound.Length - 1
                      ' Compare this form with [strVern]
                      intValue = objSim.LevenDist(dtrFound(intJ).Item("b").ToString, strLemma, intEquals)
                      ' Adjust value for initial letter mismatch
                      If (Left(dtrFound(intJ).Item("b").ToString, 1) <> Left(strLemma, 1)) Then intValue -= 3
                      ' Beware of negatives
                      If (intValue < 0) Then intValue = 0
                      ' Process best value
                      If (intValue > intConf) Then
                        intConf = intValue : strBest = dtrFound(intJ).Item("b").ToString
                        strForm = "BT" & Format(CInt(dtrFound(intJ).Item("bt").ToString), "000000")
                      End If
                    Next intJ
                  End If
                  ' Can we process it straight away?
                  If (intConf >= 85) Then
                    ' Process it
                    AddSemiStack(strF, "s=" & strForm)
                    With dtrMorph(intI)
                      .Item("f") = strF
                      .Item("l") = strLemma
                    End With
                  Else
                    ' Correction for bad ones
                    If (intConf <= 0) Then
                      strBest = "-" : strForm = "BT"
                      'Else
                      '  Stop
                    Else
                      ' Look for ambiguous lemma's
                      dtrFound = tdlBT.Tables("BT").Select("ps='" & strPosMain & "' AND b='" & strBest & "'")
                      ' Add any additional lemma's
                      For intJ = 1 To dtrFound.Length - 1
                        ' Add the BT number; the @b feature is the same
                        AddSemiStack(strForm, "BT" & Format(CInt(dtrFound(intJ).Item("bt").ToString), "000000"), strDivider:="|")
                      Next intJ
                    End If
                    ' This is all only from VP (VernPos)
                    strHasVP = "VP"
                    ' Process this into a prediction
                    wrThis.Write(.Item("VernPosId").ToString & vbTab & strVern & vbTab & .Item("Pos").ToString & vbTab & _
                            strLemma & vbTab & strF & vbTab & intConf & vbTab & strBest & vbTab & strHasVP & vbTab & strForm & _
                             vbTab & strExmp)
                    wrThis.WriteLine()
                    intNum += 1
                    ' Show what we are doing for verbs
                    If (strPosMain = "VB") Then
                      Logging(strVern & vbTab & .Item("Pos").ToString & vbTab & strLemma & vbTab & intConf & vbTab & strBest & vbTab & strForm, False)
                    End If

                  End If
                End If
              End If
            End With
          Next intI
          wrThis.Close()
          Logging("MorphDictAdaptBT results are in: " & strFile)
          ' Report the number of changes
          Logging("MorphDictAdaptBT step #6b: " & intNum & " BT numbers estimated")
          ' Need saving?
          If (bChanged) Then
            ' Save the result in morphdict
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
          End If
      End Select

      ' Step #7: Read a tab-separated file 
      Select Case MsgBox("B&T adaptation for OE step #7: read tab-separated list of lemma modifications + BT numbers [VernPos+Morph]?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' Exit
          Status("Aborted at #7") : bInterrupt = True : Return False
        Case MsgBoxResult.No
          ' Skip this step
        Case MsgBoxResult.Yes
          ' Get the file from the user
          If (Not GetFileName(frmMain.OpenFileDialog1, GetDocDir, strFile, "Tab-separated file (*.csv)|*.csv")) Then bInterrupt = True : Return False
          ' Read the file into memory
          arText = IO.File.ReadAllLines(strFile) : intNum = 0
          ' Process the file line-by-line
          For intI = 0 To arText.Length - 1
            ' Read the tab-separated list of thsi line
            arLine = Split(arText(intI), vbTab)
            ' Skip empty lines and skip lines with an empty lemma or a hyphen-lemma
            If (arLine.Length > 8) AndAlso (arLine(6) <> "") AndAlso (arLine(6) <> "-") Then
              ' Find the important components of this line
              strVern = arLine(1) : strPos = arLine(2) : strLemmaOld = arLine(3) : strLemma = arLine(6) : strMed = arLine(8)
              ' Where are we?
              intPtc = (intI + 1) * 100 \ arText.Length
              Status("MorphDictAdaptBT [" & Left(strLemma, 3) & "] " & intPtc & "%", intPtc)
              ' Find the entry in [VernPos] to be corrected
              '  dtrVP = tdlMorphDict.Tables("VernPos").Select("Vern='" & strVern & "' AND Pos='" & strPos & "' AND l='" & strLemmaOld & "'")
              dtrVP = tdlMorphDict.Tables("VernPos").Select("Vern='" & strVern & "' AND (Pos='" & strPos & "')")
              Select Case dtrVP.Length
                Case 0
                  ' This is unexpected, but can happen...
                  ' Stop
                Case Else
                  ' Why are there more than one entries?
                  '  Stop
                  ' Okay process 1 to many entries
                  For intJ = 0 To dtrVP.Length - 1
                    ' Process this entry
                    With dtrVP(intJ)
                      .Item("l") = strLemma : strF = .Item("f").ToString
                      If (Regex.IsMatch(strF, "BT\d+")) Then
                        ' How do I deal with an entry that already has a BT?
                        If (InStr(strF, strMed) = 0) Then
                          ' The existing BT number does not match the new one - replace it
                          strF = Regex.Replace(strF, "BT\d+", strMed)
                        End If
                      Else
                        ' Add the BT
                        AddSemiStack(strF, "s=" & strMed)
                      End If
                      .Item("f") = strF
                    End With
                  Next intJ
                  bChanged = True : intNum += 1
              End Select
              ' Find out if there is a matching [Morph] entry for @Vern
              '  dtrMorph = tdlMorphDict.Tables("Morph").Select("Vern='" & strVern & "' AND Pos='" & strPos & "' AND l='" & strLemmaOld & "'")
              dtrMorph = tdlMorphDict.Tables("Morph").Select("Vern='" & strVern & "' AND Pos='" & strPos & "'")
              Select Case dtrMorph.Length
                Case 0
                  ' This is unexpected
                  '   Stop
                Case Else
                  ' There may be one or more entries
                  For intJ = 0 To dtrMorph.Length - 1
                    ' Process this entry
                    With dtrMorph(intJ)
                      .Item("l") = strLemma : strF = .Item("f").ToString
                      If (Regex.IsMatch(strF, "BT\d+")) Then
                        ' How do I deal with an entry that already has a BT?
                        If (InStr(strF, strMed) = 0) Then
                          ' Replace it
                          strF = Regex.Replace(strF, "BT\d+", strMed)
                        End If
                      Else
                        ' Add the BT
                        AddSemiStack(strF, "s=" & strMed)
                      End If
                      .Item("f") = strF
                    End With
                  Next intJ
                  bChanged = True : intNum += 1
              End Select
              ' Get the lemma part-of-speech
              strPosMain = MorphCatMain(strPos, True)
              intPos = InStr(strPosMain, "+")
              If (intPos > 0) Then strPosMain = Mid(strPosMain, intPos + 1)
              strPosList = MorphCatList(strPosMain, True)
              ' Check out other entries in [Morph] with the same lemma
              dtrMorph = tdlMorphDict.Tables("Morph").Select("l='" & strLemmaOld & "'")
              For intJ = 0 To dtrMorph.Length - 1
                With dtrMorph(intJ)
                  ' Is this the correct POS?
                  If (DoLike(.Item("Pos").ToString, strPosList)) Then
                    ' Change the lemma and so on
                    .Item("l") = strLemma : strF = .Item("f").ToString
                    If (Regex.IsMatch(strF, "BT\d+")) Then
                      ' How do I deal with an entry that already has a BT?
                      If (InStr(strF, strMed) = 0) Then
                        ' Replace it
                        strF = Regex.Replace(strF, "BT\d+", strMed)
                      End If
                    Else
                      ' Add the BT
                      AddSemiStack(strF, "s=" & strMed) : .Item("f") = strF
                    End If
                  End If
                End With
              Next intJ
              ' Check if there (now) is an entry in [Morph] for the correct lemma
              If (tdlMorphDict.Tables("Morph").Select("Vern='" & strLemma & "' AND Pos='" & strPosMain & "' AND l='" & strLemma & "'").Length = 0) Then
                If (Not MorphDictAdd(strLemma, strPosMain, "", strLemma, "s=" & strMed)) Then Return False
              End If
            End If
          Next intI
          ' Report the number of changes
          Logging("MorphDictAdaptBT step #7: " & intNum & " BT numbers processed")
          ' Need saving?
          If (bChanged) Then
            ' Save the result in morphdict
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
          End If
      End Select

      ' Step #8: Read a tab-separated file 
      Select Case MsgBox("B&T adaptation for OE step #8: deep-process csv lemma-modification list + BT numbers [VernPos+Morph]?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' Exit
          Status("Aborted at #8") : bInterrupt = True : Return False
        Case MsgBoxResult.No
          ' Skip this step
        Case MsgBoxResult.Yes
          ' Get the file from the user
          If (Not GetFileName(frmMain.OpenFileDialog1, GetDocDir, strFile, "Tab-separated file (*.csv)|*.csv")) Then bInterrupt = True : Return False
          ' Read the file into memory
          arText = IO.File.ReadAllLines(strFile) : intNum = 0
          ' Process the file line-by-line
          For intI = 0 To arText.Length - 1
            ' Read the tab-separated list of thsi line
            arLine = Split(arText(intI), vbTab)
            ' Skip empty lines and skip lines with an empty lemma or a hyphen-lemma
            If (arLine.Length > 8) AndAlso (arLine(6) <> "") AndAlso (arLine(6) <> "-") Then
              ' Find the important components of this line
              strVern = arLine(1) : strPos = arLine(2) : strLemmaOld = arLine(3) : strLemma = arLine(6) : strMed = arLine(8)
              ' Where are we?
              intPtc = (intI + 1) * 100 \ arText.Length
              Status("MorphDictAdaptBT [" & Left(strLemma, 3) & "] " & intPtc & "%", intPtc)
              ' Get the lemma part-of-speech
              strPosMain = MorphCatMain(strPos, True)
              intPos = InStr(strPosMain, "+")
              If (intPos > 0) Then strPosMain = Mid(strPosMain, intPos + 1)
              strPosList = MorphCatList(strPosMain, True)
              ' Find the entry in [VernPos] to be corrected
              '  dtrVP = tdlMorphDict.Tables("VernPos").Select("Vern='" & strVern & "' AND Pos='" & strPos & "' AND l='" & strLemmaOld & "'")
              dtrVP = tdlMorphDict.Tables("VernPos").Select("l='" & strLemmaOld & "'")
              ' Okay process 1 to many entries
              For intJ = 0 To dtrVP.Length - 1
                ' Process this entry
                With dtrVP(intJ)
                  ' Debug.Print("Pos=" & dtrVP(intJ).Item("Pos").ToString)
                  '   If (InStr(dtrVP(intJ).Item("Pos").ToString, "+") > 0) OrElse (InStr(strPos, "+") > 0) Then Stop
                  ' Reduce the POS
                  strPosMain = MorphCatMain(strPos, True)
                  intPos = InStr(strPosMain, "+")
                  If (intPos > 0) Then strPosMain = Mid(strPosMain, intPos + 1)
                  ' Is this the correct POS?
                  If (DoLike(strPosMain, strPosList)) Then
                    .Item("l") = strLemma : strF = .Item("f").ToString
                    If (Regex.IsMatch(strF, "BT\d+")) Then
                      ' How do I deal with an entry that already has a BT?
                      If (InStr(strF, strMed) = 0) Then strF = Regex.Replace(strF, "BT\d+", strMed)
                    Else
                      ' Add the BT
                      AddSemiStack(strF, "s=" & strMed)
                    End If
                    ' Put the combined F-number back
                    .Item("f") = strF : bChanged = True : intNum += 1
                  End If
                End With
              Next intJ
              ' Check out other entries in [Morph] with the same lemma
              dtrMorph = tdlMorphDict.Tables("Morph").Select("l='" & strLemmaOld & "'")
              For intJ = 0 To dtrMorph.Length - 1
                With dtrMorph(intJ)
                  '   Debug.Print("Pos=" & dtrVP(intJ).Item("Pos").ToString)
                  '  If (InStr(dtrMorph(intJ).Item("Pos").ToString, "+") > 0) OrElse (InStr(strPos, "+") > 0) Then Stop
                  ' Reduce the POS
                  strPosMain = MorphCatMain(strPos, True)
                  intPos = InStr(strPosMain, "+")
                  If (intPos > 0) Then strPosMain = Mid(strPosMain, intPos + 1)
                  ' Is this the correct POS?
                  If (.Item("l").ToString <> strLemma) AndAlso (DoLike(strPosMain, strPosList)) Then
                    ' Change the lemma and so on
                    .Item("l") = strLemma : strF = .Item("f").ToString
                    If (Regex.IsMatch(strF, "BT\d+")) Then
                      ' How do I deal with an entry that already has a BT?
                      If (InStr(strF, strMed) = 0) Then strF = Regex.Replace(strF, "BT\d+", strMed)
                    Else
                      ' Add the BT
                      AddSemiStack(strF, "s=" & strMed)
                    End If
                    ' Put the combined F-number back
                    .Item("f") = strF : bChanged = True : intNum += 1
                  End If
                End With
              Next intJ
              ' Check if there (now) is an entry in [Morph] for the correct lemma
              If (tdlMorphDict.Tables("Morph").Select("Vern='" & strLemma & "' AND Pos='" & strPosMain & "' AND l='" & strLemma & "'").Length = 0) Then
                If (Not MorphDictAdd(strLemma, strPosMain, "", strLemma, "s=" & strMed)) Then Return False
              End If
            End If
          Next intI
          ' Report the number of changes
          Logging("MorphDictAdaptBT step #8: " & intNum & " BT numbers processed")
          ' Need saving?
          If (bChanged) Then
            ' Save the result in morphdict
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
          End If
      End Select


      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MorphDictAdaptBT error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphDictAdaptFindBT
  ' Goal:   Try find a BT number for [strLemmaOld] with pos within [strPosMain]
  ' History:
  ' 24-06-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function MorphDictAdaptFindBT(ByVal strTable As String, ByVal strLemmaOld As String, ByVal strPosMain As String, _
                                        ByRef strLemma As String, ByRef strMed As String) As Boolean
    Dim dtrFound() As DataRow ' Results
    Dim strPosList As String  ' All POS that are valid
    Dim intJ As Integer       ' Counter
    Dim strF As String        ' value of @f
    Dim strPos As String      ' Current POS

    Try
      ' Get a list of appropriate POS
      strPosList = MorphCatList(strPosMain, True) : strMed = ""
      ' Find and change all that is needed in [Morph]
      '   dtrFound = tdlMorphDict.Tables(strTable).Select("l='" & MyTrim(strLemmaOld) & "'")
      dtrFound = tdlMorphDict.Tables(strTable).Select("l='" & MyTrim(strLemmaOld) & "'")
      '  Debug.Print(Hex(AscW(Left(strLemmaOld, 1))))
      For intJ = 0 To dtrFound.Length - 1
        With dtrFound(intJ)
          strPos = .Item("Pos").ToString
          ' Are we valid?
          If (DoLike(strPos, strPosList)) Then
            strF = .Item("f").ToString
            If (Regex.IsMatch(strF, "BT\d+")) Then
              ' We found it!!
              strLemma = .Item("l")
              strMed = Regex.Match(strF, "BT\d+").Value
              Return True
            End If
          End If
        End With
      Next intJ
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MorphDictAdaptFindBT error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphDictAdaptOneLemmaBT
  ' Goal:   Adapt all entries in [strTable] that have lemma [strLemmaOld] and pos = [strPosMain]
  '         Change their lemma to [strLemma] and add the BT number in [strMed] to their @f feature
  ' History:
  ' 24-06-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function MorphDictAdaptOneLemmaBT(ByVal strTable As String, ByVal strLemmaOld As String, ByVal strLemma As String, ByVal strMed As String, _
                                           ByVal strPosMain As String, ByRef bChanged As Boolean, ByRef intNum As Integer) As Boolean
    Dim dtrFound() As DataRow ' Results
    Dim strPosList As String  ' All POS that are valid
    Dim intJ As Integer       ' Counter
    Dim strF As String        ' value of @f
    Dim strPos As String      ' Current POS

    Try
      ' Get a list of appropriate POS
      strPosList = MorphCatList(strPosMain, True)
      ' Find and change all that is needed in [Morph]
      dtrFound = tdlMorphDict.Tables(strTable).Select("l='" & strLemmaOld & "'")
      For intJ = 0 To dtrFound.Length - 1
        With dtrFound(intJ)
          strPos = .Item("Pos").ToString
          ' Are we valid?
          If (DoLike(strPos, strPosList)) Then
            .Item("l") = strLemma : strF = .Item("f").ToString
            If (Not Regex.IsMatch(strF, "BT\d+")) Then
              AddSemiStack(strF, "s=" & strMed) : .Item("f") = strF
              bChanged = True : intNum += 1
            End If
          End If
        End With
      Next intJ
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MorphDictAdaptOneLemmaBT error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphDictAdaptMed
  ' Goal:   Adapt the tdlMorphDict with information from the MED, stored in [MEDcombinedRes.xml]
  ' History:
  ' 25-04-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MorphDictAdaptMed(ByVal strType As String) As Boolean
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim dtrVP() As DataRow        ' Entry from [VernPos] table
    Dim dtrNew As DataRow         ' One new entry for [Morph] table
    Dim dtrEntry() As DataRow     ' Walk through entries
    Dim dtrLexFun() As DataRow    ' Walk through LexFuns
    Dim dtrMorph() As DataRow     ' Section from [Morph] with same lemma
    Dim dtrAmbi() As DataRow      ' Potentially ambiguous
    Dim dtrRemv() As DataRow      ' Result of looking for the one that needs to be removed
    Dim arLf() As String          ' array of @lf
    Dim strLf As String = ""      ' Lexical function
    Dim strF As String = ""       ' Semi-separated array of features
    Dim arF(0) As String          ' Array of features
    Dim strLv As String = ""      ' form
    Dim strLemma As String = ""   ' Lemma
    Dim strLemmaOld As String     ' Old lemma
    Dim strVern As String = ""    ' The vernacular form
    Dim strMed As String = ""     ' The MED number
    Dim strFilter As String = ""  ' Filter to be used
    Dim strFile As String = ""    ' Filename (tab-separated)
    Dim strMEDfile As String = "\Data Files\Research\2013\V2_ME\Data\MED_combinedRes.xml"
    Dim arText() As String        ' Lines of file
    Dim arLine() As String        ' Parts of one line
    Dim colHtml As New StringColl ' Error page
    Dim colEntry As New StringColl ' List of EntryIds that need to be deleted
    Dim tdlSpare As DataSet       ' Shadow collection
    Dim bLemmaAmbi As Boolean     ' Is this lemma ambiguous or not?
    Dim bCorrect As Boolean       ' Do we have the correct lemma in [Morph] already?
    Dim bChanged As Boolean       ' Anything changed?
    Dim intEntryId As Integer     ' Index in MED file
    Dim intMorphId As Integer     ' MorphId
    Dim intCorrect As Integer     ' Index of correct lemma
    Dim intMnum As Integer = 0    ' Additions in [Morph] table
    Dim intVPnum As Integer = 0   ' Additions in [VernPos] table
    Dim intPtc As Integer         ' Percentage
    Dim intNum As Integer         ' Number of entries
    Dim intPos As Integer         ' Position in string
    Dim intK As Integer           ' Counter
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intL As Integer           ' counter

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Logging("The morphdict is not opened") : Return False
      ' Find the MED-dict
      If (IO.File.Exists("U:" & strMEDfile)) Then
        strMEDfile = "U:" & strMEDfile
      ElseIf (IO.File.Exists("D:" & strMEDfile)) Then
        strMEDfile = "D:" & strMEDfile
      Else
        ' Problem
        Logging("Could not find MED dictionary: " & strMEDfile) : Return False
      End If
      ' Open the dictionary
      Status("Opening MED combined...")
      If (Not ReadDataset("OEdict.xsd", strMEDfile, tdlOEdict)) Then Logging("Could not open MED combined") : Return False
      Debug.Print(tdlOEdict.Tables("LexFun").Rows.Count)
      ' Get the restriction
      Select Case strType.ToLower
        Case "verbs", "verb"
          strFilter = "Pos = 'VB'"
      End Select
      ' Delete superfluous entries in [Morph]
      Select Case MsgBox("1. Would you like to delete superfluous entries from [Morph]?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' We leave
          Return False
        Case MsgBoxResult.No
          ' No action
        Case MsgBoxResult.Yes
          ' Create a shadow table
          tdlSpare = Nothing
          If (Not CreateDataSet("MorphDict.xsd", tdlSpare)) Then Return False
          ' Copy table [Morph] to the shadow, but without any hyphens
          dtrMorph = tdlMorphDict.Tables("Morph").Select("", "l ASC")
          Status("Shadow copy...")
          For intI = 0 To dtrMorph.Length - 1
            dtrNew = AddOneDataRow(tdlSpare, "Morph", "MorphId", "MorphList", False)
            With dtrMorph(intI)
              ' Copy features
              dtrNew.Item("MorphId") = .Item("MorphId").ToString
              dtrNew.Item("Vern") = .Item("Vern").ToString.Replace("-", "")
              dtrNew.Item("Pos") = .Item("Pos").ToString
              dtrNew.Item("Label") = .Item("Label").ToString
              dtrNew.Item("l") = .Item("l").ToString.Replace("-", "")
              dtrNew.Item("f") = .Item("f").ToString
              ' Vern items with a hyphen should be adjusted too -- you never meet them in real life!
              If (InStr(dtrMorph(intI).Item("Vern").ToString, "-") > 0) Then
                dtrMorph(intI).Item("Vern") = dtrMorph(intI).Item("Vern").ToString.Replace("-", "")
              End If
            End With
          Next intI
          ' Walk through the entries of the [Morph] table
          dtrMorph = tdlMorphDict.Tables("Morph").Select("Label='VB' AND f LIKE 's=MED*'", "l ASC")
          For intI = 0 To dtrMorph.Length - 1
            ' access this one
            dtrNew = dtrMorph(intI) : strF = dtrNew.Item("f").ToString
            strLemma = dtrNew.Item("l").ToString
            ' =================== DEBUG ======================
            ' If (strLemma = "leten") Then Stop
            ' ================================================
            ' Where are we?
            intPtc = (intI + 1) * 100 \ dtrMorph.Length
            Status("MorphDictAdaptMed " & Left(strLemma, 3) & " " & intPtc & "%", intPtc)
            ' Get the MED string, if present
            If (Regex.IsMatch(strF, "MED\d+")) Then
              strMed = Regex.Match(strF, "MED\d+").Value
            Else
              strMed = ""
            End If
            strLemmaOld = ""
            ' The MED must be present
            If (strMed <> "") Then
              ' Look for all entries with the same Vernacular Lemma (or differing only in a hyphen...)
              If (InStr(strLemma, "-") = 0) Then
                dtrFound = tdlSpare.Tables("Morph").Select("Label='VB' AND Vern='" & strLemma & "'")
              Else
                dtrFound = tdlSpare.Tables("Morph").Select("Label='VB' AND " & _
                          "(Vern='" & strLemma & "' OR Vern='" & strLemma.Replace("-", "") & "')")
              End If
              intCorrect = -1 : bCorrect = False
              ' Try find one with the MED field
              For intJ = dtrFound.Length - 1 To 0 Step -1
                If (InStr(dtrFound(intJ).Item("f").ToString, strMed) > 0) Then
                  bCorrect = True : intCorrect = intJ : Exit For
                End If
              Next intJ
              ' Found anything? - if not, then keep element [0], and delete the remainder
              If (Not bCorrect) Then intCorrect = 0
              ' Get the lemma that needs to be deleted
              strLemmaOld = dtrMorph(intI).Item("l").ToString
              ' Delete all superfluous ones
              For intJ = dtrFound.Length - 1 To 0 Step -1
                ' Keep or delete?
                If (intJ <> intCorrect) Then
                  ' Mark it as needing to be deleted
                  dtrRemv = tdlMorphDict.Tables("Morph").Select("MorphId = " & dtrFound(intJ).Item("MorphId").ToString)
                  If (dtrRemv.Length > 0) Then
                    dtrRemv(0).Item("f") = "REMOVE" : intMnum += 1
                    strLemmaOld = dtrRemv(0).Item("l").ToString
                  Else
                    strLemmaOld = ""
                  End If
                End If
              Next intJ
              ' ================= HIER GAAT HET MIS =========================
              If (strLemmaOld = "") Then
                ' Unexpected...
                Stop
              ElseIf (strLemma <> strLemmaOld) AndAlso (InStr(strLemmaOld, "-") > 0) Then
                ' ======================== DEBUG ==============
                ' If (strLemmaOld = "leten") Then Stop
                ' =============================================
                ' Change bad lemma into the correct one
                dtrFound = tdlMorphDict.Tables("Morph").Select("l='" & strLemmaOld & "'")
                For intJ = 0 To dtrFound.Length - 1
                  ' Change the lemma
                  dtrFound(intJ).Item("l") = strLemma
                Next intJ
                ' Change bad lemma into the correct one - also for the [VernPos]
                dtrFound = tdlMorphDict.Tables("VernPos").Select("l='" & strLemmaOld & "'")
                For intJ = 0 To dtrFound.Length - 1
                  ' Change the lemma
                  dtrFound(intJ).Item("l") = strLemma
                Next intJ
              End If
              ' ============================================================
            End If
          Next intI
          ' Get all those that need to be removed
          dtrMorph = tdlMorphDict.Tables("Morph").Select("f='REMOVE'", "l ASC")
          For intI = dtrMorph.Length - 1 To 0 Step -1
            ' Remove this one
            dtrMorph(intI).Delete()
          Next intI
          ' Update
          tdlMorphDict.AcceptChanges()
          ' Save the result in morphdict
          Logging("Deleted items: " & intMnum)
          Status("Saving: " & strMorphDictFile)
          tdlMorphDict.WriteXml(strMorphDictFile)
      End Select
      ' Establish the link from [VernPos] entries back to [Morph]
      Select Case MsgBox("2. Would you like to set @MorphId links in [VernPos]?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' We leave
          Return False
        Case MsgBoxResult.No
          ' No action
        Case MsgBoxResult.Yes
          ' Walk through all the entries of the [VernPos] table
          dtrVP = tdlMorphDict.Tables("VernPos").Select("", "l ASC")
          strLemma = "" : intNum = dtrVP.Length : intMorphId = -1 : strMed = ""
          For intI = 0 To intNum - 1
            ' What is the current lemma?
            If (strLemma <> dtrVP(intI).Item("l").ToString) OrElse _
               (strMed <> Regex.Match(dtrVP(intI).Item("f").ToString, "(M|O)ED\d+").Value) Then
              ' Adapt lemma and MED number
              strLemma = dtrVP(intI).Item("l").ToString
              strMed = Regex.Match(dtrVP(intI).Item("f").ToString, "(M|O)ED\d+").Value
              ' Get the new morphid
              dtrMorph = tdlMorphDict.Tables("Morph").Select("Label='VB' AND l='" & strLemma & "' AND t='l'")
              If (dtrMorph.Length = 0) Then
                ' This should not be possible
                ' Stop
                ' Set a default one??
                intMorphId = -1
                ' Or: create an entry in [Morph] with the lemma??
              ElseIf (dtrMorph.Length > 1) Then
                ' There are multiple entries
                '' Get the MED number intended
                'strF = dtrVP(intI).Item("f").ToString : strMed = Regex.Match(strF, "(M|O)ED\d+").Value
                If (strMed <> "") Then
                  bCorrect = False
                  For intJ = 0 To dtrMorph.Length - 1
                    strF = dtrMorph(intJ).Item("f").ToString
                    If (Regex.Match(strF, "(M|O)ED\d+").Value = strMed) Then
                      intMorphId = CInt(dtrMorph(intJ).Item("MorphId").ToString) : bCorrect = True : Exit For
                    End If
                  Next intJ
                  If (Not bCorrect) Then
                    ' This MED number still needs to be added to the table [Morph]
                    ' Access this lemma in [tdlOEdict]
                    dtrEntry = tdlOEdict.Tables("Entry").Select("Pos='VB' AND s='" & strMed & "' AND l='" & strLemma & "'")
                    If (dtrEntry.Length = 0) Then
                      ' Try with only the MED number
                      dtrEntry = tdlOEdict.Tables("Entry").Select("Pos='VB' AND s='" & strMed & "'")
                      If (dtrEntry.Length = 0) AndAlso (Not strLf Like "VB*") Then
                        dtrEntry = tdlOEdict.Tables("Entry").Select("s='" & strMed & "'")
                      End If
                    End If
                    If (dtrEntry.Length > 0) Then
                      ' Get the entry id
                      intEntryId = dtrEntry(0).Item("EntryId") : strLf = "VB" : strLv = strLemma
                      ' Create an entry for this lemma
                      dtrNew = AddOneDataRow(tdlMorphDict, "Morph", "MorphId", "MorphList")
                      With dtrNew
                        .Item("EntryId") = intEntryId : .Item("Vern") = strLv : .Item("Pos") = strLf : .Item("Label") = strLf
                        .Item("l") = strLemma : .Item("f") = strF : .Item("t") = "l"
                        intMorphId = .Item("MorphId").ToString
                      End With
                      intMnum += 1
                    Else
                      '  Check if this is an OED number
                      If (Not strMed Like "OED*") Then
                        ' No it is not OED, so complain
                        ' Show in Tab-separated format what we cannot add
                        Logging(dtrVP(intI).Item("Vern").ToString & vbTab & "VB" & vbTab & strLemma & vbTab & strLemma & vbTab & strMed)
                      End If
                      ' Process it anyway
                      intMorphId = -1
                    End If
                  End If
                End If
              Else
                ' Get the MorphId (choose the first by default)
                intMorphId = dtrMorph(0).Item("MorphId").ToString
              End If
            End If
            ' Where are we?
            intPtc = (intI + 1) * 100 \ intNum
            Status("MorphDictAdaptMed " & Left(strLemma, 3) & " " & intPtc & "%", intPtc)
            ' Process this one?
            If (dtrVP(intI).Item("MorphId").ToString = "") _
              OrElse (dtrVP(intI).Item("MorphId").ToString = "-1") _
              OrElse (dtrVP(intI).Item("MorphId").ToString = "0") Then
              ' Set the link to the morph id
              dtrVP(intI).Item("MorphId") = intMorphId
              ' Possibly adapt the [src]
              If (dtrVP(intI).Item("Src").ToString Like "*MorphId*") Then
                dtrVP(intI).Item("Src") = "Auto"
              End If
            End If
          Next intI
          ' Save the result in morphdict
          Status("Saving: " & strMorphDictFile)
          tdlMorphDict.WriteXml(strMorphDictFile)
      End Select
      ' See if we can get improvements for [Morph] from here
      If (MsgBox("3. Load csv file with matches between [Morph] and MED?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes) Then
        While (GetFileName(frmMain.OpenFileDialog1, GetDocDir(), strFile, "tab-separated text file (*.csv)|*.csv"))
          intMnum = 0 : intVPnum = 0
          Logging("Processing " & strFile)
          ' Read the file as array
          arText = IO.File.ReadAllLines(strFile)
          For intI = 0 To arText.Length - 1
            intPtc = (1 + intI) * 100 \ arText.Length
            Status("MorphDictAdaptMed " & intPtc & "%", intPtc)
            If (Trim(arText(intI)) <> "") Then
              ' DIvide the line into parts
              arLine = Split(arText(intI), vbTab)
              ' Get the information
              strLemmaOld = arLine(2) : strLf = arLine(1) : strLv = arLine(0) : strLemma = arLine(3) : strMed = arLine(4)
              ' Look for an entry in [Morph] that contains the 'old' spelling of the lemma
              dtrMorph = tdlMorphDict.Tables("Morph").Select("t='l' AND l='" & strLemmaOld & "'")
              bCorrect = False
              If (dtrMorph.Length > 0) Then
                ' Visit all these entries
                For intJ = 0 To dtrMorph.Length - 1
                  ' See if this entry does NOT contain the MED number
                  strF = dtrMorph(intJ).Item("f").ToString
                  If (Not DoLike(strF, "*[OM]ED*")) Then
                    ' Access it
                    dtrNew = dtrMorph(intJ)
                    ' Add the MED number
                    AddSemiStack(strF, "s=" & strMed) : dtrNew.Item("f") = strF
                    ' Change the lemma to the new one
                    dtrNew.Item("l") = strLemma
                    intMorphId = dtrNew.Item("MorphId").ToString
                    ' Adapt all the necessary entries in [Morph] with the old lemma
                    bCorrect = True
                  End If
                Next intJ
              End If
              If (bCorrect) Then
                ' Look for all entries in [Morph] that need repairing
                dtrMorph = tdlMorphDict.Tables("Morph").Select("l='" & strLemmaOld & "' AND t<>'l'")
                For intJ = 0 To dtrMorph.Length - 1
                  ' Make sure the [f] is changed here
                  With dtrMorph(intJ)
                    strF = .Item("f").ToString
                    ' Check if this one has MED or not
                    If (Not DoLike(strF, "*[OM]ED*")) Then
                      AddSemiStack(strF, "s=" & strMed) : .Item("f") = strF
                      ' Change the lemma too
                      .Item("l") = strLemma
                      intMnum += 1
                    End If
                  End With
                Next intJ
                ' Look for entries in [VernPos] that need repairing
                dtrVP = tdlMorphDict.Tables("VernPos").Select("l='" & strLemma & "'")
                For intJ = 0 To dtrVP.Length - 1
                  ' Make sure the [f] is changed here
                  With dtrVP(intJ)
                    strF = .Item("f").ToString
                    If (Not DoLike(strF, "*[OM]ED*")) Then
                      AddSemiStack(strF, "s=" & strMed) : .Item("f") = strF
                      ' Change the lemma too
                      .Item("l") = strLemma
                      ' Repair link to MorphId
                      If (.Item("MorphId").ToString <> intMorphId) Then
                        .Item("MorphId") = intMorphId
                      End If
                      intVPnum += 1
                    End If
                  End With
                Next intJ
              End If
            End If
          Next intI
          If (intMnum > 0) Then
            ' Show how many we have added
            Logging("Changes to [Morph]: " & intMnum)
            Logging("Changes to [VernPos]: " & intVPnum)
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
          End If
        End While
      End If
      ' Ask user if first part is required
      Select Case MsgBox("4. Would you like to retrieve missing MED numbers in [Morph] from MED_combined?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' We leave
          Return False
        Case MsgBoxResult.No
        Case MsgBoxResult.Yes
          ' Walk all the lemma's in [Morph]
          intMnum = 0 : intVPnum = 0
          ' Walk all the lemma's in the [Morph] table
          dtrFound = tdlMorphDict.Tables("Morph").Select("Label LIKE 'V*' AND t<>'l'")
          For intI = 0 To dtrFound.Length - 1
            ' Show where we are
            intPtc = (intI + 1) * 100 \ dtrFound.Length
            Status("MorphDictAdapt " & intPtc & "%", intPtc)
            ' Is this one that has a MED number?
            With dtrFound(intI)
              strF = .Item("f").ToString : strLemmaOld = .Item("l").ToString
              ' Check if this entry already has an MED number
              If (Not DoLike(strF, "*s=[MO]ED*")) AndAlso (InStr(strLemmaOld, "-") > 0) Then
                ' Try to find the lemma entry by looking for a lemma without hyphen
                strLemma = strLemmaOld.Replace("-", "")
                dtrMorph = tdlMorphDict.Tables("Morph").Select("Label = 'VB' AND t='l' AND l='" & strLemma & "'")
                If (dtrMorph.Length > 0) Then
                  ' Get this new one's MED number if possible
                  strF = dtrMorph(0).Item("f").ToString : strMed = Regex.Match(strF, "(O|M)ED\d+").Value
                  ' Change all the appropriate old lemma's to this new one
                  dtrMorph = tdlMorphDict.Tables("Morph").Select("Label LIKE 'V*' AND l='" & strLemmaOld & "'")
                  For intJ = 0 To dtrMorph.Length - 1
                    With dtrMorph(intJ)
                      ' Change the lemma
                      .Item("l") = strLemma
                      ' Check the MED number
                      strF = dtrMorph(intJ).Item("f").ToString
                      If (Not strF Like "*[OM]ED*") Then AddSemiStack(strF, "s=" & strMed) : dtrMorph(intJ).Item("f") = strF
                      intMnum += 1
                    End With
                  Next intJ
                End If
              End If
            End With
          Next intI

          ' Walk all the lemma's in the [Morph] table
          dtrFound = tdlMorphDict.Tables("Morph").Select("Label='VB' AND t='l'")
          For intI = 0 To dtrFound.Length - 1
            ' Show where we are
            intPtc = (intI + 1) * 100 \ dtrFound.Length
            Status("MorphDictAdapt " & intPtc & "%", intPtc)
            ' Is this one that has a MED number?
            With dtrFound(intI)
              strF = .Item("f").ToString : strLemmaOld = .Item("l").ToString
              ' Check if this entry already has an MED number
              If (Not DoLike(strF, "*s=[MO]ED*")) Then
                ' No MED number yet, so try to get a match in [LexFun]
                dtrLexFun = tdlOEdict.Tables("LexFun").Select("lf='VB' AND lv='" & strLemmaOld & "'")
                If (dtrLexFun.Length = 0) AndAlso (InStr(strLemmaOld, "-") > 0) Then
                  strLemmaOld = strLemmaOld.Replace("-", "")
                  dtrLexFun = tdlOEdict.Tables("LexFun").Select("lf='VB' AND lv='" & strLemmaOld & "'")
                End If
                ' Got any results?>
                If (dtrLexFun.Length > 0) Then
                  ' Get all the MED numbers together
                  strMed = ""
                  For intJ = 0 To dtrLexFun.Length - 1
                    AddSemiStack(strMed, dtrLexFun(intJ).GetParentRow("Entry_LexFun").Item("s").ToString, True)
                  Next intJ
                  ' The lemma is that of the first we meet
                  strLemma = dtrLexFun(0).GetParentRow("Entry_LexFun").Item("l").ToString
                  ' Replace old with new lemma in current place in [Morph]
                  strF = .Item("f").ToString : AddSemiStack(strF, "s=" & strMed) : .Item("f") = strF
                  .Item("l") = strLemma
                  intMnum += 1
                  ' Replace old with new lemma in all places in [Morph]
                  dtrMorph = tdlMorphDict.Tables("Morph").Select("l='" & strLemmaOld & "' AND t<>'l'")
                  For intJ = 0 To dtrMorph.Length - 1
                    ' Make sure the [f] is changed here
                    With dtrMorph(intJ)
                      strF = .Item("f").ToString
                      ' Check if this one has MED or not
                      If (Not DoLike(strF, "*[OM]ED*")) Then
                        AddSemiStack(strF, "s=" & strMed) : .Item("f") = strF
                        ' Change the lemma too
                        .Item("l") = strLemma
                        intMnum += 1
                      End If
                    End With
                  Next intJ
                  ' Look for entries in [VernPos] that need repairing
                  dtrVP = tdlMorphDict.Tables("VernPos").Select("l='" & strLemma & "'")
                  For intJ = 0 To dtrVP.Length - 1
                    ' Make sure the [f] is changed here
                    With dtrVP(intJ)
                      strF = .Item("f").ToString
                      If (Not DoLike(strF, "*[OM]ED*")) Then
                        AddSemiStack(strF, "s=" & strMed) : .Item("f") = strF
                        ' Change the lemma too
                        .Item("l") = strLemma
                        ' Repair link to MorphId
                        If (.Item("MorphId").ToString <> intMorphId) Then
                          .Item("MorphId") = intMorphId
                        End If
                        intVPnum += 1
                      End If
                    End With
                  Next intJ
                Else
                  ' Indicate that we are missing this value
                  Logging(.Item("Vern").ToString & vbTab & .Item("Label").ToString & vbTab & _
                          .Item("l").ToString & vbTab & "-" & vbTab & "-")
                End If

              End If
            End With
          Next intI
          If (intMnum > 0) Then
            ' Show how many we have added
            Logging("Changes to [Morph]: " & intMnum)
            Logging("Changes to [VernPos]: " & intVPnum)
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
          End If
      End Select
      ' Ask user if first part is required
      Select Case MsgBox("5. Would you like to assign the appropriate MED numbers to lemma's in [Morph]?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' We leave
          Return False
        Case MsgBoxResult.No
        Case MsgBoxResult.Yes
          ' Walk all the lemma's in the [Entry] table 
          dtrEntry = tdlOEdict.Tables("Entry").Select("Pos='VB'")
          For intI = 0 To dtrEntry.Length - 1
            ' Get this lemma and the MED number
            strLemma = dtrEntry(intI).Item("l").ToString : strMed = dtrEntry(intI).Item("s").ToString
            ' Where are we?
            intPtc = (intI + 1) * 100 \ dtrEntry.Length
            Status("MorphDictAdaptMed " & Left(strLemma, 3) & " " & intPtc & "%", intPtc)
            ' Find this lemma in the [Morph] table
            dtrMorph = tdlMorphDict.Tables("Morph").Select("l='" & strLemma & "' AND t='l'")
            ' Got it?
            If (dtrMorph.Length > 0) Then
              ' Check if this lemma has the MED number
              strF = dtrMorph(0).Item("f").ToString
              If (Not DoLike(strF, "*[OM]ED*")) Then
                AddSemiStack(strF, "s=" & strMed) : dtrMorph(0).Item("f") = strF : intMorphId = CInt(dtrMorph(0).Item("MorphId").ToString)
                ' Look for all entries in [Morph] that need repairing
                dtrMorph = tdlMorphDict.Tables("Morph").Select("l='" & strLemma & "' AND t<>'l'")
                For intJ = 0 To dtrMorph.Length - 1
                  ' Make sure the [f] is changed here
                  With dtrMorph(intJ)
                    strF = .Item("f").ToString
                    If (Not DoLike(strF, "*[OM]ED*")) Then AddSemiStack(strF, "s=" & strMed) : .Item("f") = strF
                  End With
                Next intJ
                ' Look for entries in [VernPos] that need repairing
                dtrVP = tdlMorphDict.Tables("VernPos").Select("l='" & strLemma & "'")
                For intJ = 0 To dtrVP.Length - 1
                  ' Make sure the [f] is changed here
                  With dtrVP(intJ)
                    strF = .Item("f").ToString
                    If (Not DoLike(strF, "*[OM]ED*")) Then AddSemiStack(strF, "s=" & strMed) : .Item("f") = strF
                    ' Repair link to MorphId
                    If (.Item("MorphId").ToString <> intMorphId) Then
                      .Item("MorphId") = intMorphId
                    End If
                  End With
                Next intJ
              End If
            End If
          Next intI
          ' Walk all the lemma's in the [Morph] table
          dtrMorph = tdlMorphDict.Tables("Morph").Select("Label='VB' AND t='l'")
          For intI = 0 To dtrMorph.Length - 1
            ' Show where we are
            intPtc = (intI + 1) * 100 \ dtrMorph.Length
            Status("MorphDictAdapt " & intPtc & "%", intPtc)
            ' Is this one that has a MED number?
            With dtrMorph(intI)
              strF = .Item("f").ToString : strLemma = .Item("l").ToString
              If (DoLike(strF, "*s=MED*")) Then
                ' This is a lemma with an MED/OED number - get this number
                strMed = Regex.Match(strF, "MED\d+").Value
                ' Look for other entries in [Morph] pointing to the same lemma
                dtrFound = tdlMorphDict.Tables("Morph").Select("l='" & strLemma & "' AND t<>'l'")
                For intJ = 0 To dtrFound.Length - 1
                  ' Make sure the [f] is changed here
                  With dtrFound(intJ)
                    strF = .Item("f").ToString
                    If (Not DoLike(strF, "*[OM]ED*")) Then AddSemiStack(strF, "s=" & strMed) : .Item("f") = strF
                  End With
                Next intJ
                ' Look for entries in [VernPos] that need repairing
                dtrVP = tdlMorphDict.Tables("VernPos").Select("l='" & strLemma & "'")
                For intJ = 0 To dtrVP.Length - 1
                  ' Make sure the [f] is changed here
                  With dtrVP(intJ)
                    strF = .Item("f").ToString
                    If (Not DoLike(strF, "*[OM]ED*")) Then AddSemiStack(strF, "s=" & strMed) : .Item("f") = strF
                    ' Repair link to MorphId
                    If (.Item("MorphId").ToString <> intMorphId) Then
                      .Item("MorphId") = intMorphId
                    End If
                  End With
                Next intJ
              End If
            End With
          Next intI
          ' Save the result in morphdict
          Status("Saving: " & strMorphDictFile)
          tdlMorphDict.WriteXml(strMorphDictFile)
      End Select
      ' Ask user if this part is required
      Select Case MsgBox("6a. Would you like to adapt the lemma's in [VernPos] (using L)?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' We leave
          Return False
        Case MsgBoxResult.No
        Case MsgBoxResult.Yes
          ' Walk through all the VernPos items
          dtrVP = tdlMorphDict.Tables("VernPos").Select("Pos='VB'")
          intNum = dtrVP.Length : intVPnum = 0
          For intI = 0 To intNum - 1
            ' Get this one's lemma
            dtrNew = dtrVP(intI)
            strLemmaOld = dtrNew.Item("l").ToString
            ' Where are we?
            intPtc = (intK + 1) * 100 \ intNum
            Status("MorphDictAdaptMed " & Left(strLemma, 3) & " " & intPtc & "%", intPtc)
            ' Get the F
            strF = dtrNew.Item("f").ToString
            ' Does it contain a MED?
            If (Not DoLike(strF, "*s=[MO]ED*")) Then
              ' Try to retrieve the correct lemma
              dtrMorph = tdlMorphDict.Tables("Morph").Select("Vern='" & strLemmaOld & "' AND Label = 'VB'")
              If (dtrMorph.Length = 0) AndAlso (InStr(strLemmaOld, "-") > 0) Then
                dtrMorph = tdlMorphDict.Tables("Morph").Select("Vern='" & strLemmaOld.Replace("-", "") & "' AND Label = 'VB'")
              End If
              ' Do we get a lemma?
              If (dtrMorph.Length > 0) Then
                ' Expand the first lemma
                strLemma = dtrMorph(0).Item("l").ToString : intMorphId = dtrMorph(0).Item("MorphId").ToString
                strF = dtrMorph(0).Item("f").ToString
                strMed = Regex.Match(strF, "MED\d+").Value
                Logging(strLemmaOld & vbTab & strLemma & vbTab & strMed)
                ' Find all entries in [VernPos] marked with this lemma
                dtrFound = tdlMorphDict.Tables("VernPos").Select("l='" & strLemmaOld & "'")
                For intJ = 0 To dtrFound.Length - 1
                  ' Get this one's F
                  strF = dtrFound(intJ).Item("f").ToString
                  ' Does it have a MED?
                  If (Not DoLike(strF, "*s=[MO]ED*")) Then
                    ' Give it the MED and the lemma and the MorphId
                    AddSemiStack(strF, "s=" & strMed)
                    With dtrFound(intJ)
                      .Item("f") = strF : .Item("l") = strLemma : .Item("MorphId") = intMorphId
                      intVPnum += 1
                    End With
                  End If
                Next intJ
              End If
            End If
          Next intI
          If (intVPnum > 0) Then
            ' Show how many we have added
            Logging("Changes to [VernPos]: " & intVPnum)
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.AcceptChanges()
            tdlMorphDict.WriteXml(strMorphDictFile)
          End If
      End Select
      ' Ask user if this part is required
      Select Case MsgBox("6b. Would you like to adapt the lemma's in [VernPos] (Using VB)?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' We leave
          Return False
        Case MsgBoxResult.No
        Case MsgBoxResult.Yes
          ' Walk through all the VernPos items
          dtrVP = tdlMorphDict.Tables("VernPos").Select("Pos='VB'")
          intNum = dtrVP.Length : intVPnum = 0
          For intI = 0 To intNum - 1
            ' Get this one's lemma
            dtrNew = dtrVP(intI)
            strLemmaOld = dtrNew.Item("Vern").ToString
            ' If (strLemmaOld = "abuyen") Then Stop
            ' Where are we?
            intPtc = (intK + 1) * 100 \ intNum
            Status("MorphDictAdaptMed " & Left(strLemma, 3) & " " & intPtc & "%", intPtc)
            ' Get the F
            strF = dtrNew.Item("f").ToString
            ' Does it contain a MED?
            If (Not DoLike(strF, "*s=[MO]ED*")) Then
              ' Try to retrieve the correct lemma
              dtrMorph = tdlMorphDict.Tables("Morph").Select("Vern='" & strLemmaOld & "' AND Label = 'VB'")
              ' Do we get a lemma?
              If (dtrMorph.Length > 0) Then
                ' Expand the first lemma
                strLemma = dtrMorph(0).Item("l").ToString : intMorphId = dtrMorph(0).Item("MorphId").ToString
                strF = dtrMorph(0).Item("f").ToString
                strMed = Regex.Match(strF, "(O|M)ED\d+").Value
                Logging(strLemmaOld & vbTab & strLemma & vbTab & strMed)
                ' Find all entries in [VernPos] marked with this lemma
                dtrFound = tdlMorphDict.Tables("VernPos").Select("l='" & dtrNew.Item("l").ToString & "'")
                For intJ = 0 To dtrFound.Length - 1
                  ' Get this one's F
                  strF = dtrFound(intJ).Item("f").ToString
                  ' Does it have a MED?
                  If (Not DoLike(strF, "*s=[MO]ED*")) Then
                    ' Give it the MED and the lemma and the MorphId
                    AddSemiStack(strF, "s=" & strMed)
                    With dtrFound(intJ)
                      .Item("f") = strF : .Item("l") = strLemma : .Item("MorphId") = intMorphId
                      intVPnum += 1
                    End With
                  End If
                Next intJ
              End If
            End If
          Next intI
          If (intVPnum > 0) Then
            ' Show how many we have added
            Logging("Changes to [VernPos]: " & intVPnum)
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
          End If
      End Select
      ' See if we can get improvements for [Morph] from here
      If (MsgBox("6c. Load csv file with matches between [VernPos] and MED?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes) Then
        While (GetFileName(frmMain.OpenFileDialog1, GetDocDir(), strFile, "tab-separated text file (*.csv)|*.csv"))
          intMnum = 0 : intVPnum = 0
          Logging("Processing " & strFile)
          ' Read the file as array
          arText = IO.File.ReadAllLines(strFile)
          For intI = 0 To arText.Length - 1
            intPtc = (1 + intI) * 100 \ arText.Length
            Status("MorphDictAdaptMed " & intPtc & "%", intPtc)
            If (Trim(arText(intI)) <> "") Then
              ' DIvide the line into parts
              arLine = Split(arText(intI), vbTab)
              ' Get the information
              strLemmaOld = arLine(2) : strLf = arLine(1) : strLv = arLine(0) : strLemma = arLine(3) : strMed = arLine(4)
              ' Skip empty lemma's and stuff
              If (strLemma <> "" AndAlso strMed <> "" AndAlso strLemma <> "-" AndAlso strMed <> "") Then
                ' Look for an entry in [VernPos] that contains this @Vern as well as the 'old' spelling of the @lemma
                dtrVP = tdlMorphDict.Tables("VernPos").Select("Vern = '" & strLv & "' AND l='" & strLemmaOld & "'")
                If (dtrVP.Length > 0) Then
                  ' Visit all these entries
                  For intJ = 0 To dtrVP.Length - 1
                    ' See if this entry does NOT contain the MED number
                    strF = dtrVP(intJ).Item("f").ToString
                    If (Not DoLike(strF, "*[OM]ED*")) Then
                      ' Access it
                      dtrNew = dtrVP(intJ)
                      ' Add the MED number
                      AddSemiStack(strF, "s=" & strMed) : dtrNew.Item("f") = strF
                      ' Change the lemma to the new one
                      dtrNew.Item("l") = strLemma
                      ' Indicate correction
                      intVPnum += 1
                      ' Look for the Lemma/MED combination in [Morph]
                      bCorrect = False
                      dtrMorph = tdlMorphDict.Tables("Morph").Select("l='" & strLemma & "'")
                      For intK = 0 To dtrMorph.Length - 1
                        strF = dtrMorph(intK).Item("f").ToString
                        If (DoLike(strF, "*[OM]ED*")) Then
                          bCorrect = True : Exit For
                        End If
                      Next intK
                      If (Not bCorrect) Then
                        ' Access this lemma in [tdlOEdict]
                        dtrEntry = tdlOEdict.Tables("Entry").Select("Pos='VB' AND s='" & strMed & "' AND l='" & strLemma & "'")
                        If (dtrEntry.Length = 0) Then
                          ' Try with only the MED number
                          dtrEntry = tdlOEdict.Tables("Entry").Select("Pos='VB' AND s='" & strMed & "'")
                          If (dtrEntry.Length = 0) AndAlso (Not strLf Like "VB*") Then
                            dtrEntry = tdlOEdict.Tables("Entry").Select("s='" & strMed & "'")
                          End If
                        End If
                        If (dtrEntry.Length > 0) Then
                          ' Get the entry id
                          intEntryId = dtrEntry(0).Item("EntryId") : strLf = "VB" : strLv = strLemma
                          ' Create an entry for this lemma
                          dtrNew = AddOneDataRow(tdlMorphDict, "Morph", "MorphId", "MorphList")
                          With dtrNew
                            .Item("EntryId") = intEntryId : .Item("Vern") = strLv : .Item("Pos") = strLf : .Item("Label") = strLf
                            .Item("l") = strLemma : .Item("f") = strF : .Item("t") = "l"
                            intMorphId = .Item("MorphId").ToString
                          End With
                          intMnum += 1
                          ' Also adapt the VP entry
                          dtrVP(intJ).Item("MorphId") = intMorphId
                        Else
                          '  Check if this is an OED number
                          If (strMed Like "OED*") Then
                            ' Process it anyway
                            dtrVP(intJ).Item("MorphId") = -1
                          Else
                            ' No it is not OED, so complain
                            ' Show in Tab-separated format what we cannot add
                            Logging(arText(intI))
                          End If
                        End If
                      Else
                        ' Recover the MorphId number
                        intMorphId = CInt(dtrMorph(intK).Item("MorphId").ToString)
                        dtrVP(intJ).Item("MorphId") = intMorphId
                      End If
                    End If
                  Next intJ
                End If
              End If
            End If
          Next intI
          If (intMnum > 0) Then
            ' Show how many we have added
            Logging("Changes to [Morph]: " & intMnum)
            Logging("Changes to [VernPos]: " & intVPnum)
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
          End If
        End While
      End If
      ' Ask user if first part is required
      Select Case MsgBox("7. Would you like to extrapolate MED to MorphDict?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' We leave
          Return False
        Case MsgBoxResult.No
        Case MsgBoxResult.Yes
          ' Do it!
          ' Walk through all the entries in the MED file
          dtrEntry = tdlOEdict.Tables("Entry").Select(strFilter)
          intNum = dtrEntry.Length
          For intK = 0 To intNum - 1
            ' Get the lemma and the MED number
            With dtrEntry(intK)
              strLemma = .Item("l").ToString : strMed = .Item("s").ToString : intEntryId = .Item("EntryId").ToString
            End With
            ' Where are we?
            ' intK += 1
            intPtc = (intK + 1) * 100 \ intNum
            Status("MorphDictAdaptMed " & Left(strLemma, 3) & " " & intPtc & "%", intPtc)
            ' Find all the lexical functions belonging to this particular [EntryId]
            dtrLexFun = tdlOEdict.Tables("LexFun").Select("EntryId=" & intEntryId)
            ' Walk all of these [LexFun] items
            For intI = 0 To dtrLexFun.Length - 1
              arLf = Split(dtrLexFun(intI).Item("lf").ToString, ";") : strLv = dtrLexFun(intI).Item("lv").ToString
              ' ============= DEBUG =================
              ' If (strLv = "folwen") Then Stop
              ' =====================================
              ' Do not take into account LV entries that: 
              ' - start with a hyphen
              ' - contain a comma
              If (Left(strLv, 1) <> "-") AndAlso (InStr(strLv, ",") = 0) Then
                ' Get the actual POS, which is the beginning of the @lf field
                strLf = arLf(0)
                ' Look for potential ambiguity
                dtrAmbi = tdlOEdict.Tables("LexFun").Select("lv = '" & strLv & "'")
                bLemmaAmbi = False
                For intJ = 0 To dtrAmbi.Length - 1
                  With dtrAmbi(intJ)
                    If (.Item("lf").ToString = strLf) OrElse (.Item("lf").ToString Like strLf & ";*") Then
                      If (.Item("EntryId").ToString <> intEntryId) Then
                        bLemmaAmbi = True : Exit For
                      End If
                    End If
                  End With
                Next intJ
                ' Look for features, which are further on in the @lf field
                If (arLf.Length > 1) Then
                  ReDim arF(0 To arLf.Length - 1)
                  arF(0) = "s=" & strMed
                  For intJ = 1 To arLf.Length - 1
                    arF(intJ) = arLf(intJ)
                  Next intJ
                  strF = Join(arF, ";")
                Else
                  strF = "s=" & strMed
                End If
                ' Exclude other categories
                If (strLf Like "V*") Then
                  ' Find the combination Lf/Lv in the [Morph] table
                  If (strLf = "VBN") Then
                    strFilter = "(Label = 'VBN' OR Label = 'VAN') AND Vern='" & strLv & "'"
                  Else
                    strFilter = "Label = '" & strLf & "' AND Vern='" & strLv & "'"
                  End If
                  dtrFound = tdlMorphDict.Tables("Morph").Select(strFilter, "Vern ASC, Pos ASC, l ASC")
                  ' Action depends on the number of entries
                  Select Case dtrFound.Length
                    Case 0
                      ' We can add this form and this lemma to the morph table
                      dtrNew = AddOneDataRow(tdlMorphDict, "Morph", "MorphId", "MorphList")
                      With dtrNew
                        .Item("EntryId") = intEntryId : .Item("Vern") = strLv : .Item("Pos") = strLf : .Item("Label") = strLf
                        .Item("l") = strLemma : .Item("f") = strF : .Item("t") = IIf(strLemma = strLv, "l", "d")
                        intMorphId = CInt(.Item("MorphId").ToString)
                      End With
                      intMnum += 1
                      ' Is this a VBN?
                      If (strLf = "VBN") Then
                        ' Add a similar entry, but now for VAN
                        dtrNew = AddOneDataRow(tdlMorphDict, "Morph", "MorphId", "MorphList")
                        With dtrNew
                          .Item("EntryId") = intEntryId : .Item("Vern") = strLv : .Item("Pos") = "VAN" : .Item("Label") = "VAN"
                          .Item("l") = strLemma : .Item("f") = strF : .Item("t") = IIf(strLemma = strLv, "l", "d")
                          intMorphId = CInt(.Item("MorphId").ToString)
                        End With
                        intMnum += 1
                      End If
                    Case Else
                      ' The combination @Label/@Vern/@Lemma is there, but is it there with the same MED number?
                      ' Visit them all
                      bCorrect = False
                      For intJ = 0 To dtrFound.Length - 1
                        ' Compare the MED number of the LexFun item
                        If (strMed = Regex.Match(dtrFound(intJ).Item("f").ToString, "(M|O)ED\d+").Value) Then
                          bCorrect = True : Exit For
                        End If
                      Next intJ
                      ' Did I find the correct MED number?
                      If (Not bCorrect) Then
                        ' This combination @Label/@Vern/@Lemma is not yet in [Morph], so add it
                        dtrNew = AddOneDataRow(tdlMorphDict, "Morph", "MorphId", "MorphList")
                        With dtrNew
                          .Item("EntryId") = intEntryId : .Item("Vern") = strLv : .Item("Pos") = strLf : .Item("Label") = strLf
                          .Item("l") = strLemma : .Item("f") = strF : .Item("t") = IIf(strLemma = strLv, "l", "d")
                          intMorphId = CInt(.Item("MorphId").ToString)
                        End With
                        intMnum += 1
                        ' Is this a VBN?
                        If (strLf = "VBN") Then
                          ' Add a similar entry, but now for VAN
                          dtrNew = AddOneDataRow(tdlMorphDict, "Morph", "MorphId", "MorphList")
                          With dtrNew
                            .Item("EntryId") = intEntryId : .Item("Vern") = strLv : .Item("Pos") = "VAN" : .Item("Label") = "VAN"
                            .Item("l") = strLemma : .Item("f") = strF : .Item("t") = IIf(strLemma = strLv, "l", "d")
                            intMorphId = CInt(.Item("MorphId").ToString)
                          End With
                          intMnum += 1
                        End If
                      End If

                      'Case 1
                      '  ' An entry already exists, but may need to be updated and checked
                      '  dtrNew = dtrFound(0) : intMorphId = dtrNew.Item("MorphId").ToString
                      '  ' Keep the lemma for a second
                      '  strLemmaOld = dtrNew.Item("l").ToString
                      '  ' Is the lemma the same?
                      '  If (strLemmaOld <> strLemma) Then
                      '    ' Change this one's lemma
                      '    dtrNew.Item("l") = strLemma
                      '    ' Check for ambiguity
                      '    If (Not bLemmaAmbi) Then
                      '      ' Change all lemma's that need to be changed in the [Morph] table
                      '      dtrMorph = tdlMorphDict.Tables("Morph").Select("Label LIKE 'V*' AND l='" & strLemmaOld & "'")
                      '      For intJ = 0 To dtrMorph.Length - 1
                      '        ' Change the lemma and the feature
                      '        With dtrMorph(intJ)
                      '          .Item("l") = strLemma
                      '          ' Does this entry have a MED indicator?
                      '          strF = .Item("f").ToString
                      '          If (Not DoLike(strF, "*[OM]ED*")) Then AddSemiStack(strF, "s=" & strMed) : .Item("f") = strF
                      '        End With
                      '      Next intJ
                      '    End If
                      '  End If
                      '  ' Does this entry have a MED indicator?
                      '  strF = dtrNew.Item("f").ToString
                      '  If (Not DoLike(strF, "*[OM]ED*")) Then AddSemiStack(strF, "s=" & strMed) : dtrNew.Item("f") = strF
                      '  ' Also make changes in the [VernPos] table
                      '  If (Not MorphDictAdaptVernPos(strLv, strLf, strLemma, strLemmaOld, bLemmaAmbi, strMed, intMorphId, strF, intVPnum)) Then Return False
                      'Case Else
                      '  ' There are multiple entries
                      '  ' Stop
                      '  ' Check if there is an entry with the correct lemma
                      '  bCorrect = False : intCorrect = -1 : dtrNew = Nothing : strLemmaOld = strLemma
                      '  For intJ = 0 To dtrFound.Length - 1
                      '    ' Is this the correct lemma?
                      '    With dtrFound(intJ)
                      '      If (.Item("l").ToString.Replace("-", "") = strLemma) Then
                      '        ' This is the correct lemma
                      '        strLemmaOld = .Item("l").ToString
                      '        intMorphId = .Item("MorphId").ToString
                      '        bCorrect = True : intCorrect = intJ : dtrNew = dtrFound(intJ) : Exit For
                      '      End If
                      '    End With
                      '  Next intJ
                      '  ' If there is no correct one, then we'll take the first one and delete the others
                      '  If (Not bCorrect) Then
                      '    intCorrect = 0 : dtrNew = dtrFound(0) : intMorphId = dtrNew.Item("MorphId").ToString
                      '  End If
                      '  ' We have a correct lemma --> delete all that is not correct
                      '  For intJ = 0 To dtrFound.Length - 1
                      '    If (intJ <> intCorrect) Then
                      '      ' Find corresponding entry in [VernPos]
                      '      dtrVP = tdlMorphDict.Tables("VernPos").Select("Vern = '" & strLv & "' AND Pos = '" & strLf & _
                      '                                                    "' AND l = '" & dtrFound(intJ).Item("l").ToString & "'")
                      '      For intL = dtrVP.Length - 1 To 0 Step -1
                      '        ' Delete this entry
                      '        dtrVP(intL).Delete()
                      '      Next intL
                      '      ' Check correspondences for VBN/VAN
                      '      If (strLf = "VBN") Then
                      '        dtrVP = tdlMorphDict.Tables("VernPos").Select("Vern = '" & strLv & "' AND Pos = 'VAN'" & _
                      '                                                      " AND l = '" & dtrFound(intJ).Item("l").ToString & "'")
                      '        For intL = dtrVP.Length - 1 To 0 Step -1
                      '          ' Delete this entry
                      '          dtrVP(intL).Delete()
                      '        Next intL
                      '      End If
                      '      ' Delete entry from [Morph]
                      '      dtrFound(intJ).Delete()
                      '    End If
                      '  Next intJ
                      '  ' Make sure the table is refreshed
                      '  tdlMorphDict.AcceptChanges()
                      '  ' Is the lemma the same?
                      '  If (strLemmaOld <> strLemma) Then
                      '    ' Change this one's lemma
                      '    dtrNew.Item("l") = strLemma
                      '    ' Change all lemma's that need to be changed
                      '    dtrMorph = tdlMorphDict.Tables("Morph").Select("Label LIKE 'V*' AND l='" & strLemmaOld & "'")
                      '    For intJ = 0 To dtrMorph.Length - 1
                      '      ' Change the lemma and the feature
                      '      With dtrMorph(intJ)
                      '        .Item("l") = strLemma
                      '        ' Does this entry have a MED indicator?
                      '        strF = .Item("f").ToString
                      '        If (Not DoLike(strF, "*[OM]ED*")) Then AddSemiStack(strF, "s=" & strMed) : .Item("f") = strF
                      '      End With
                      '    Next intJ
                      '  End If
                      '  ' Does this entry have a MED indicator?
                      '  strF = dtrNew.Item("f").ToString
                      '  If (Not DoLike(strF, "*[OM]ED*")) Then AddSemiStack(strF, "s=" & strMed) : dtrNew.Item("f") = strF
                      '  ' Also make changes in the [VernPos] table
                      '  If (Not MorphDictAdaptVernPos(strLv, strLf, strLemma, strLemmaOld, bLemmaAmbi, strMed, intMorphId, strF, intVPnum)) Then Return False
                  End Select
                End If
              End If
            Next intI
            ' Go to next entry
            ' ndxEntry = ndxEntry.SelectSingleNode("./following-sibling::Entry[@Pos='VB'][1]")
            ' End While
          Next intK
          ' Show how many we have added
          Logging("Additions to [Morph]: " & intMnum)
          Logging("Additions to [VernPos]: " & intVPnum)
          Status("Saving: " & strMorphDictFile)
          tdlMorphDict.WriteXml(strMorphDictFile)
      End Select
      ' Second step?
      Select Case MsgBox("8. Would you like to try retrieve [Morph] unsolved entries from MED?", vbYesNoCancel)
        Case MsgBoxResult.Cancel
          ' We leave
          Return False
        Case MsgBoxResult.No
        Case MsgBoxResult.Yes
          ' Now do the reverse: start out with the [Morph] table, check for missing entries, and try to retrieve them from the [tdlOEdict]
          dtrFound = tdlMorphDict.Tables("Morph").Select("(f NOT LIKE '*s=MED*') AND (f NOT LIKE '*s=OED*') ")
          intNum = dtrFound.Length
          For intI = 0 To intNum - 1
            ' This is my entry
            dtrNew = dtrFound(intI)
            ' Get the lemma that is being used here
            strLemmaOld = dtrNew.Item("l").ToString : strLemma = "" : ReDim dtrEntry(0)
            ' Get the vernacular
            strLv = dtrNew.Item("Vern").ToString
            ' Where are we?
            intPtc = (intI + 1) * 100 \ intNum
            Status("MorphDictAdaptMed " & Left(strLemmaOld, 3) & " " & intPtc & "%", intPtc)
            If (Not TryGetMEDlemma(strLv, strLemmaOld, strLemma, strMed, dtrEntry)) Then Return False
            ' Did we find anything?
            If (dtrEntry.Length > 0) Then
              ' =============== DEBUG =================
              ' Stop
              ' =======================================
              ' Okay, we have the correct lemma
              strLemma = dtrEntry(0).Item("l").ToString
              strMed = dtrEntry(0).Item("s").ToString
              ' Adapt the settings in [Morph]
              ' (1) Change this one's lemma
              dtrNew.Item("l") = strLemma : intMnum += 1
              ' (2) Does this entry have a MED indicator?
              strF = dtrNew.Item("f").ToString
              If (Not DoLike(strF, "*[OM]ED*")) Then AddSemiStack(strF, "s=" & strMed) : dtrNew.Item("f") = strF
              ' Also make changes in the [VernPos] table
              bLemmaAmbi = False : intMorphId = dtrNew.Item("MorphId").ToString
              strLv = dtrNew.Item("Vern").ToString : strLf = dtrNew.Item("Label").ToString
              If (Not MorphDictAdaptVernPos(strLv, strLf, strLemma, strLemmaOld, bLemmaAmbi, strMed, intMorphId, strF, intVPnum)) Then Return False
            End If
          Next intI
          If (intMnum > 0) Then
            ' Show how many we have added
            Logging("Changes to [Morph]: " & intMnum)
            Logging("Changes to [VernPos]: " & intVPnum)
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
          End If
      End Select
      ' See if we can get improvements for [Morph] from here
      If (MsgBox("9. Load csv file with improvements for [Morph] table?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes) Then
        If (GetFileName(frmMain.OpenFileDialog1, GetDocDir(), strFile, "tab-separated text file (*.csv)|*.csv")) Then
          ' Read the file as array
          arText = IO.File.ReadAllLines(strFile)
          For intI = 0 To arText.Length - 1
            If (Trim(arText(intI)) <> "") Then
              ' DIvide the line into parts
              arLine = Split(arText(intI), vbTab)
              ' Get the information
              strLemmaOld = arLine(2) : strLf = arLine(1) : strLv = arLine(0) : strLemma = arLine(3) : strMed = arLine(4)
              ' Get the entry in [Morph]
              dtrMorph = tdlMorphDict.Tables("Morph").Select("Vern = '" & strLv & "' AND Label = '" & strLf & "' AND l = '" & strLemmaOld & "'")
              For intJ = 0 To dtrMorph.Length - 1
                dtrNew = dtrMorph(intJ)
                ' Adapt the settings in [Morph]
                ' (1) Change this one's lemma
                dtrNew.Item("l") = strLemma : intMnum += 1
                ' (2) Does this entry have a MED indicator?
                strF = dtrNew.Item("f").ToString
                If (Not DoLike(strF, "*[OM]ED*")) Then AddSemiStack(strF, "s=" & strMed) : dtrNew.Item("f") = strF
                ' Also make changes in the [VernPos] table
                bLemmaAmbi = False : intMorphId = dtrNew.Item("MorphId").ToString
                ' strLv = dtrNew.Item("Vern").ToString : strLf = dtrNew.Item("Label").ToString
                If (Not MorphDictAdaptVernPos(strLv, strLf, strLemma, strLemmaOld, bLemmaAmbi, strMed, intMorphId, strF, intVPnum)) Then Return False
              Next intJ
            End If
          Next intI
          If (intMnum > 0) Then
            ' Show how many we have added
            Logging("Changes to [Morph]: " & intMnum)
            Logging("Changes to [VernPos]: " & intVPnum)
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
          End If
        End If
      End If
      ' See if we can get improvements for [Morph] from here
      If (MsgBox("Load csv file with improvements for [VernPos] table?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes) Then
        If (GetFileName(frmMain.OpenFileDialog1, GetDocDir(), strFile, "tab-separated text file (*.csv)|*.csv")) Then
          ' Read the file as array
          arText = IO.File.ReadAllLines(strFile)
          intVPnum = 0
          For intI = 0 To arText.Length - 1
            intPtc = (intI + 1) * 100 \ arText.Length
            Status("MorphDictAdaptMed " & intPtc & "%", intPtc)
            If (Trim(arText(intI)) <> "") Then
              ' DIvide the line into parts
              arLine = Split(arText(intI), vbTab)
              ' Get the information
              strLemmaOld = arLine(2) : strLf = arLine(1) : strLv = arLine(0) : strLemma = arLine(3)
              strMed = arLine(4) : bLemmaAmbi = False : strF = "s=" & strMed
              ' Is this an addition or deletion
              If (strLemma = "-") Then
                ' This is a deletion: find all that need to be deleted
                dtrVP = tdlMorphDict.Tables("VernPos").Select("Vern='" & strLv & "' AND Pos='" & strLf & "'" & _
                                                              " AND l='" & strLemmaOld & "'")
                If (dtrVP.Length > 0) Then
                  For intJ = dtrVP.Length - 1 To 0 Step -1
                    ' Delete this vernpos entry
                    dtrVP(intJ).Delete()
                  Next intJ
                  ' Make sure we are resetting
                  tdlMorphDict.AcceptChanges()
                End If
              Else
                ' Find out the [MorphId], if possible, for this item
                dtrMorph = tdlMorphDict.Tables("Morph").Select("Label='" & strLf & "' AND l='" & strLemma & "'" & _
                                                               " AND Vern='" & strLv & "'")
                If (dtrMorph.Length > 0) Then
                  intMorphId = dtrMorph(0).Item("MorphId")
                Else
                  intMorphId = -1
                End If
                ' Possibly correct this item
                If (Not MorphDictAdaptVernPos(strLv, strLf, strLemma, strLemmaOld, bLemmaAmbi, strMed, intMorphId, strF, intVPnum)) Then Return False
              End If
            End If
          Next intI
          If (intVPnum > 0) Then
            ' Show how many we have added
            Logging("Changes to [VernPos]: " & intVPnum)
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
          End If
        End If
      End If
      ' Check for any remaining puzzles
      Select Case MsgBox("Would you like to check for remaining entries?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Cancel
          Return False
        Case MsgBoxResult.No
          ' Don't do anything
        Case MsgBoxResult.Ok, MsgBoxResult.Yes
          colHtml.Add("<html><body><table><tr><td>Vern</td><td>Pos</td><td>Lemma</td><td>Suggestion</td><td>MED</td></tr>")
          ' First check the table [Morph]
          dtrFound = tdlMorphDict.Tables("Morph").Select("(f IS Null) OR ( (f NOT LIKE '*s=MED*') AND (f NOT LIKE '*s=OED*') )")
          intNum = dtrFound.Length
          If (intNum > 0) Then
            For intI = 0 To intNum - 1
              ' This is my entry
              dtrNew = dtrFound(intI)
              With dtrNew
                ' Try get a suggestion
                ' Get the lemma that is being used here
                strLemmaOld = .Item("l").ToString : strLemma = "" : ReDim dtrEntry(0)
                ' Get the vernacular
                strLv = dtrNew.Item("Vern").ToString
                ' =========== DEBUG ==============
                ' If (.Item("Vern").ToString = "begaet") Then Stop
                ' ================================
                ' Where are we?
                intPtc = (intI + 1) * 100 \ intNum
                Status("MorphDictAdaptMed " & Left(strLemmaOld, 3) & " " & intPtc & "%", intPtc)
                If (Not TryGetMEDlemma(strLv, strLemmaOld, strLemma, strMed, dtrEntry)) Then Return False
                ' Did we find anything?
                If (dtrEntry.Length > 0) AndAlso (dtrEntry(0) IsNot Nothing) Then
                  colHtml.Add("<tr><td>" & .Item("Vern").ToString & "</td><td>" & .Item("Label").ToString & "</td><td>" & strLemmaOld & _
                              "</td><td>" & strLemma & "</td><td>" & strMed & "</td></tr>")
                Else
                  colHtml.Add("<tr><td>" & .Item("Vern").ToString & "</td><td>" & .Item("Label").ToString & "</td><td>" & strLemmaOld & _
                              "</td><td>-</td><td>-</td></tr>")
                End If
              End With
            Next intI
          End If
          ' Next check the table [VernPos]
          dtrFound = tdlMorphDict.Tables("VernPos").Select("(f IS Null) OR ( (f NOT LIKE '*s=MED*') AND (f NOT LIKE '*s=OED*') )")
          intNum = dtrFound.Length
          If (intNum > 0) Then
            For intI = 0 To intNum - 1
              ' This is my entry
              dtrNew = dtrFound(intI)
              With dtrNew
                ' Try get a suggestion
                ' Get the lemma that is being used here
                strLemmaOld = .Item("l").ToString : strLemma = "" : ReDim dtrEntry(0)
                ' Get the vernacular
                strLv = dtrNew.Item("Vern").ToString
                ' =========== DEBUG ==============
                ' If (.Item("Vern").ToString = "begaet") Then Stop
                ' ================================
                ' Where are we?
                intPtc = (intI + 1) * 100 \ intNum
                Status("MorphDictAdaptMed " & Left(strLemmaOld, 3) & " " & intPtc & "%", intPtc)
                If (Not TryGetMEDlemma(strLv, strLemmaOld, strLemma, strMed, dtrEntry)) Then Return False
                ' Did we find anything?
                If (dtrEntry.Length > 0) AndAlso (dtrEntry(0) IsNot Nothing) Then
                  colHtml.Add("<tr><td>" & .Item("Vern").ToString & "</td><td>" & .Item("Pos").ToString & "</td><td>" & strLemmaOld & _
                              "</td><td>" & strLemma & "</td><td>" & strMed & "</td></tr>")
                Else
                  colHtml.Add("<tr><td>" & .Item("Vern").ToString & "</td><td>" & .Item("Pos").ToString & "</td><td>" & strLemmaOld & _
                              "</td><td>-</td><td>-</td></tr>")
                End If
              End With
            Next intI
          End If
          ' Finish it off
          colHtml.Add("</table></body></html>")
          frmMain.wbReport.DocumentText = colHtml.Text
          frmMain.TabControl1.SelectedTab = frmMain.tpReport
      End Select
      ' Copy Morph information to VernPos
      Select Case MsgBox("Would you like to copy [Morph] information to [VernPos]?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Cancel
          Return False
        Case MsgBoxResult.No
          ' Don't do anything
        Case MsgBoxResult.Ok, MsgBoxResult.Yes
          ' Select all the VernPos tables that have not been treated yet
          dtrVP = tdlMorphDict.Tables("VernPos").Select("(f IS Null) OR ( (f NOT LIKE '*s=MED*') AND (f NOT LIKE '*s=OED*') )")
          intNum = dtrVP.Length : intVPnum = 0
          For intI = 0 To intNum - 1
            ' This is my entry
            dtrNew = dtrVP(intI) : strVern = dtrNew.Item("Vern").ToString : strLf = dtrNew.Item("Pos").ToString
            strLemma = dtrNew.Item("l").ToString : strF = dtrNew.Item("f").ToString
            ' Where are we?
            intPtc = (intI + 1) * 100 \ intNum
            Status("MorphDictAdaptMed " & Left(strLemma, 3) & " " & intPtc & "%", intPtc)
            ' Find the corresponding entry in the [Morph] table
            dtrMorph = tdlMorphDict.Tables("Morph").Select("Vern='" & strVern & "' AND Label='" & strLf & "'")
            If (dtrMorph.Count > 0) Then
              ' Adapt the vernpos value
              With dtrNew
                .Item("l") = strLemma : .Item("f") = strF
                intVPnum += 1
              End With
            End If
          Next intI
          If (intVPnum > 0) Then
            ' Show how many we have added
            Logging("Changes to [VernPos]: " & intVPnum)
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
          End If
      End Select
      ' Check [Morph] against [MED_combined]
      Select Case MsgBox("Would you like to check the [Morph] section against [MED_combined]?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Cancel
          Return False
        Case MsgBoxResult.No
          ' Don't do anything
        Case MsgBoxResult.Ok, MsgBoxResult.Yes
          ' Walk through all the entries of the [Morph] section
          intMnum = 0
          dtrMorph = tdlMorphDict.Tables("Morph").Select("", "f ASC")
          For intI = 0 To dtrMorph.Length - 1
            ' Get this entry
            dtrNew = dtrMorph(intI)
            ' Get this entry's info
            With dtrNew
              strLv = .Item("Vern").ToString : strLf = .Item("Label").ToString : strLemmaOld = .Item("l").ToString
              strF = .Item("f").ToString
              If (Not Regex.Match(strF, "MED\d+").Success) Then
                ' This should not occur -- the MED string is missing...
                strMed = ""
              Else
                ' Retrieve the MED string
                strMed = Regex.Match(strF, "MED\d+").Value
              End If
            End With
            ' Show where we are
            Status("Looking at:" & vbTab & strMed & vbTab & strLv)
            ' Process this entry
            If (strMed <> "") Then
              ' Get all the entries within [MED_combinedRes.xml] with the correct LV and LF
              dtrLexFun = tdlOEdict.Tables("LexFun").Select("lv = '" & strLv & "'" & _
                                                             " AND lf = '" & strLf & "'")
              ' Work through them and check if any of them has the correct MED number in the parent
              bCorrect = False
              For intJ = 0 To dtrLexFun.Length - 1
                If (dtrLexFun(intJ).GetParentRow("Entry_LexFun").Item("s").ToString = strMed) Then
                  bCorrect = True : Exit For
                End If
              Next intJ
              ' Do we have the correct [s] row?
              If (bCorrect) Then
                ' Get its lemma
                strLemma = dtrLexFun(intJ).GetParentRow("Entry_LexFun").Item("l").ToString
                ' Check this lemma against the other one
                If (strLemma <> strLemmaOld) Then
                  ' Now we have a problem...
                  Logging(strLv & vbTab & strLf & vbTab & strLemmaOld & vbTab & strLemma & vbTab & strMed & vbTab & "a")
                  ' Change the lemma in [Morph] in accordance with the MED number
                  dtrNew.Item("l") = strLemma : intMnum += 1
                End If
              Else
                ' We cannot find the correct combination lv/lf/Parent.s
                ' Get the lemma that belongs to [strMed]
                dtrEntry = tdlOEdict.Tables("Entry").Select("s='" & strMed & "'")
                If (dtrEntry.Length > 0) Then
                  ' Check if the lemma here equals what we have
                  strLemma = dtrEntry(0).Item("l").ToString
                  If (strLemma <> strLemmaOld) Then
                    ' Don't bother about those lemma's that have a hyphen
                    If (InStr(strLemmaOld, "-") = 0) Then
                      Logging(strLv & vbTab & strLf & vbTab & strLemmaOld & vbTab & strLemma & vbTab & strMed & vbTab & "b")
                      ' Change the lemma in [Morph] in accordance with the MED number
                      dtrNew.Item("l") = strLemma : intMnum += 1
                    End If
                  End If
                Else
                  Logging(strLv & vbTab & strLf & vbTab & strLemmaOld & vbTab & "-" & vbTab & strMed)
                End If
              End If

            End If
          Next intI
          If (intMnum > 0) Then
            ' Show how many we have added
            Logging("Changes to [Morph]: " & intMnum)
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
          End If
      End Select
      ' Check [VernPos] against [MED_combined]
      Select Case MsgBox("Would you like to check the [VernPos] section against [MED_combined]?", MsgBoxStyle.YesNoCancel)
        Case MsgBoxResult.Cancel
          Return False
        Case MsgBoxResult.No
          ' Don't do anything
        Case MsgBoxResult.Ok, MsgBoxResult.Yes
          ' Walk through all the entries of the [Morph] section
          intVPnum = 0
          dtrVP = tdlMorphDict.Tables("VernPos").Select("", "Vern ASC")
          For intI = 0 To dtrVP.Length - 1
            ' Get this entry
            dtrNew = dtrVP(intI)
            ' Get this entry's info
            With dtrNew
              strLv = .Item("Vern").ToString : strLf = .Item("Pos").ToString : strLemmaOld = .Item("l").ToString
              strF = .Item("f").ToString
              If (Not Regex.Match(strF, "MED\d+").Success) Then
                ' This should not occur -- the MED string is missing...
                strMed = ""
              Else
                ' Retrieve the MED string
                strMed = Regex.Match(strF, "MED\d+").Value
              End If
            End With
            ' Show where we are
            Status("Looking at:" & strMed & " " & strLv)
            ' Process this entry
            If (strMed <> "") Then
              ' Get the lemma that belongs to [strMed]
              dtrEntry = tdlOEdict.Tables("Entry").Select("s='" & strMed & "'")
              If (dtrEntry.Length > 0) Then
                ' Check if the lemma here equals what we have
                strLemma = dtrEntry(0).Item("l").ToString
                If (strLemma <> strLemmaOld) Then
                  ' Don't bother about those lemma's that have a hyphen
                  If (InStr(strLemmaOld, "-") = 0) Then
                    Logging(strLv & vbTab & strLf & vbTab & strLemmaOld & vbTab & strLemma & vbTab & strMed & vbTab & "b")
                    ' Change the lemma in [Morph] in accordance with the MED number
                    dtrNew.Item("l") = strLemma : intVPnum += 1
                  End If
                End If
              Else
                Logging(strLv & vbTab & strLf & vbTab & strLemmaOld & vbTab & "-" & vbTab & strMed)
              End If
            End If
          Next intI
          If (intVPnum > 0) Then
            ' Show how many we have added
            Logging("Changes to [VernPos]: " & intVPnum)
            Status("Saving: " & strMorphDictFile)
            tdlMorphDict.WriteXml(strMorphDictFile)
          End If
      End Select
      ' Return positively
      Status("ready")
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MorphDictAdaptMed error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   CheckMedNumber
  ' Goal:   Check if the combination @Vern/@Label is correctly coded with the MED number found in [strF]
  ' History:
  ' 19-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function CheckMedNumber(ByVal strVern As String, ByVal strLabel As String, ByVal strLemma As String, ByVal strF As String) As Boolean
    Dim dtrMorph() As DataRow ' Part of [Morph]
    Dim dtrVP() As DataRow    ' Part of [VernPos]
    Dim dtrNew As DataRow     '  Onen new row
    Dim arF() As String       ' Array
    Dim intI As Integer       ' Counter
    Dim intJ As Integer       ' Counter
    Dim bFound As Boolean     ' Found MED or not
    Dim strMed As String = GetMEDorOED(strF)

    Try
      ' Validate
      If (strVern = "") OrElse (strLabel = "") OrElse (strF = "") OrElse (strMed = "") Then Return False
      ' Find the combination in [dtrMorph]
      dtrMorph = tdlMorphDict.Tables("Morph").Select("Vern='" & strVern & "' AND Label='" & strLabel & "'")
      If (dtrMorph.Length = 0) Then
        dtrMorph = tdlMorphDict.Tables("Morph").Select("Vern='" & StringOEtoTagged(strVern, True) & "' AND Label='" & strLabel & "'")
        If (dtrMorph.Length = 0) Then
          dtrMorph = tdlMorphDict.Tables("Morph").Select("Vern='" & StringOEtoTagged(strVern, False) & "' AND Label='" & strLabel & "'")
        End If
      End If
      bFound = False
      For intI = 0 To dtrMorph.Length - 1
        If (GetMEDorOED(dtrMorph(intI).Item("f").ToString = strMed) _
          AndAlso (dtrMorph(intI).Item("l").ToString = strLemma)) Then bFound = True : Exit For
      Next intI
      If (Not bFound) Then
        ' This combination is not found, so ...
        Select Case dtrMorph.Length
          Case 0
            ' Needs to be added
            ' Stop
            ' Create an entry for this lemma
            dtrNew = AddOneDataRow(tdlMorphDict, "Morph", "MorphId", "MorphList")
            With dtrNew
              .Item("EntryId") = -1 : .Item("Vern") = strVern : .Item("Pos") = strLabel : .Item("Label") = strLabel
              .Item("l") = strLemma : .Item("f") = strF : .Item("t") = "l"
            End With
          Case 1
            ' Just change the lemma and the MED
            With dtrMorph(0)
              .Item("l") = strLemma
              ' .Item("f") = strF
              ' Adapt the [strF]


              arF = Split(dtrMorph(0).Item("f").ToString, ";")
              For intJ = 0 To arF.Length - 1
                If (DoLike(arF(intJ), "s=[OM]ED*|s=BT*")) Then
                  ' Change it
                  arF(intJ) = "s=" & strMed
                End If
              Next intJ
              ' Remake it
              strF = Join(arF, ";")
              .Item("f") = strF
            End With
          Case Else
            ' There is ambiguity
            Stop
        End Select
      End If

      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/CheckMedNumber error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetMEDorOED
  ' Goal:   Get the MED or OED number from this string
  ' History:
  ' 19-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function GetMEDorOED(ByVal strIn As String) As String
    Dim strMed As String = ""
    Try
      strMed = Regex.Match(strIn, "(M|O)ED\d+").Value
      If (strMed = "") Then
        strMed = Regex.Match(strIn, "BT\d+").Value
      End If
      ' Return the result
      Return strMed
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/GetMEDorOED error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MatchEnv
  ' Goal:   Check if environment [strIn] matches pattern [strPattern]
  ' History:
  ' 01-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function MatchEnv(ByVal strIn As String, ByVal strPattern As String) As Boolean
    ' Devinition of environments: Vowel, Consonant
    Dim strVowel As String = "aeiouy"
    Dim strCons As String = "bcdfghjklmnpqrstvwxz"
    Dim strNasal As String = "bfmp"

    Try
      ' Validate
      If (strIn = "") OrElse (strPattern = "") Then Return True
      ' Look for C matching
      If (strPattern = "C") Then Return (InStr(strCons, strIn) > 0)
      ' Look for V matching
      If (strPattern = "V") Then Return (InStr(strVowel, strIn) > 0)
      ' Look for N matching
      If (strPattern = "N") Then Return (InStr(strNasal, strIn) > 0)
      ' Look for other matching
      Return (strIn = strPattern)
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MatchEnv error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   TryGetMEDlemma
  ' Goal:   Try to get a proper MED lemma for the old [strLemmaOld]
  ' History:
  ' 01-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function TryGetMEDlemma(ByVal strVern As String, ByVal strLemmaOld As String, ByRef strLemma As String, _
                                  ByRef strMed As String, ByRef dtrEntry() As DataRow) As Boolean
    Dim dtrVP() As DataRow          ' LexFun entry
    Dim colLemma As New StringColl  ' Collection of possible lemma's
    Dim intI As Integer             ' Counter

    Try
      ' Get all kinds of variants
      If (Not MakeMEDlemmaList(strVern, strLemmaOld, colLemma)) Then Return False
      ' Check all alternative lemma's in order of preference
      For intI = 0 To colLemma.Count - 1
        ' Try this lemma
        strLemma = colLemma.Item(intI)
        ' Try find this entry in [Entry]
        dtrEntry = tdlOEdict.Tables("Entry").Select("l='" & strLemma & "' AND Pos='VB'")
        If (dtrEntry.Length = 0) Then
          ' Try find this entry in [VernPos]
          dtrVP = tdlOEdict.Tables("LexFun").Select("lv='" & strLemma & "' AND lf='VB'")
          If (dtrVP.Length > 0) Then
            ' Find the [Entry] of me
            dtrEntry = tdlOEdict.Tables("Entry").Select("EntryId = " & dtrVP(0).Item("EntryId").ToString)
          End If
        End If
        ' Found anything?
        If (dtrEntry.Length > 0) Then
          ' =============== DEBUG =================
          ' Stop
          ' =======================================
          ' Okay, we have the correct lemma
          strLemma = dtrEntry(0).Item("l").ToString
          strMed = dtrEntry(0).Item("s").ToString
          Exit For
        End If
      Next intI
      ' Nothing has been found -- try looking for the vernacular
      dtrVP = tdlOEdict.Tables("LexFun").Select("lv='" & strVern & "'")
      If (dtrVP.Length > 0) Then
        ' Check until we get the first correct entry
        For intI = 0 To dtrVP.Length - 1
          ' Is this a verb derivative?
          If (dtrVP(intI).Item("lf").ToString Like "V*") Then
            ' Find the [Entry] of me
            dtrEntry = tdlOEdict.Tables("Entry").Select("EntryId = " & dtrVP(intI).Item("EntryId").ToString)
            ' Is this one correct?
            If (dtrEntry(0).Item("Pos").ToString = "VB") Then
              ' Okay, we have the correct lemma
              strLemma = dtrEntry(0).Item("l").ToString
              strMed = dtrEntry(0).Item("s").ToString
              Logging("Replace lemma [" & strLemmaOld & "] by [" & strLemma & "] " & strMed)
            Else
              ' Show what we found
              strLemma = dtrEntry(0).Item("l").ToString
              strMed = dtrEntry(0).Item("s").ToString
              Logging("Look for lemma [" & strLemmaOld & "] at [" & strLemma & "/" & dtrEntry(0).Item("Pos").ToString & "] " & strMed)
            End If
          End If
        Next intI
        '' Find the [Entry] of me
        'dtrEntry = tdlOEdict.Tables("Entry").Select("EntryId = " & dtrVP(0).Item("EntryId").ToString)
        'If (dtrEntry.Length > 0) Then
        '  ' Okay, we have the correct lemma
        '  strLemma = dtrEntry(0).Item("l").ToString
        '  strMed = dtrEntry(0).Item("s").ToString
        '  Logging("Change lemma [" & strLemmaOld & "] into [" & strLemma & "]")
        'End If
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/TryGetMEDlemma error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MakeMEDlemmaList
  ' Goal:   Try to get a proper MED lemma for the old [strLemmaOld]
  ' History:
  ' 20-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MakeMEDlemmaList(ByVal strVern As String, ByVal strLemmaOld As String, _
                                   ByRef colLemma As StringColl) As Boolean
    Dim strPrec As String           ' Preceding context
    Dim strFoll As String           ' Following context
    Dim strMatch As String          ' Match
    Dim strAlt As String            ' Alternative
    Dim strLemma As String = ""     ' Alternative lemma
    Dim intPos As Integer           ' Position in string
    Dim intI As Integer             ' Counter
    Dim intJ As Integer             ' Counter
    ' Changes that have actually been attested, within their environments
    Dim arChange() As String = { _
      "w", "u", "V", "V", _
      "y", "i", "", "C", _
      "ie", "y", "C", "#", _
      "y", "i", "V", "V", _
      "y", "i", "C", "#", _
      "y", "3", "", "", _
      "ye", "ie", "C", "C", _
      "ey", "ei", "C", "", _
      "ey", "ai", "C", "C", _
      "ay", "ei", "C", "", _
      "ay", "ai", "C", "", _
      "u", "v", "V", "V", _
      "u", "v", "C", "V", _
      "u", "ou", "C", "C", _
      "ow", "ou", "C", "C", _
      "oy", "oi", "C", "C", _
      "dh", "th", "", "", _
      "ck", "kk", "", "", _
      "sch", "sh", "", "", _
      "sch", "ssh", "", "", _
      "s", "ss", "V", "V", _
      "ns", "nc", "V", "V", _
      "ss", "sh", "V", "V", _
      "k", "c", "#", "", _
      "cw", "qu", "", "V", _
      "ien", "en", "C", "", _
      "in", "en", "C", "#", _
      "yn", "en", "C", "#", _
      "y", "3", "#", "V", _
      "neh", "neigh", "#", "", _
      "vm", "um", "#", "", _
      "n", "m", "", "N"}

    Try
      ' ============ Debug ===============
      ' If (strLemmaOld = "sprenkelin") Then Stop
      ' ==================================\
      ' Initialise
      colLemma.Clear()
      ' First preference: the old lemma
      colLemma.Add(strLemmaOld)
      ' Check what kind of trick we can use
      intPos = InStr(strLemmaOld, "-")
      ' ============= DEBUG ===============
      ' Other changes that need to be considered:
      ' ewe > eue // owe > oue ==> w > u / V_V
      ' y   > i  / C_C
      ' u   > v  / V_V
      ' sch > sh
      ' s   > ss / V_V
      ' ns  > nc / V_V
      ' ss  > sh / V_V
      ' cw  > qu
      ' ien > en / C_
      ' u   > ou / C_C
      ' n   > m  / _p
      ' ===================================
      If (intPos > 0) Then
        ' add a variant without hyphen
        strLemma = strLemmaOld.Replace("-", "") : colLemma.Add(strLemma)
      Else
        If (Left(strLemmaOld, 2) = "y-") Then
          ' Try take the lemma with "i" starting, because that's what the MED does
          strLemma = "i" & Mid(strLemmaOld, 3) : colLemma.Add(strLemma)
        End If
        ' There is a hyphen inside -- try taking the lemma *after* the hyphen
        strLemma = Mid(strLemmaOld, intPos + 1) : colLemma.Add(strLemma)
      End If
      ' Look at all kinds of other possible changes, within their possible environments
      For intI = 0 To arChange.Length - 1 Step 4
        ' Get all the components in place
        strPrec = arChange(intI + 2) : strFoll = arChange(intI + 3) : strMatch = arChange(intI)
        ' ===================================
        ' If (strMatch = "in") Then Stop
        ' ===================================
        ' Visit all possible variants
        For intJ = 0 To colLemma.Count - 1
          ' Visit all the possible letters
          strLemma = colLemma.Item(intJ)
          For intPos = 1 To strLemma.Length
            ' Check if preceding context is okay
            If (Mid(strLemma, intPos, strMatch.Length) = strMatch) Then
              ' Check preceding context
              If (strPrec = "" OrElse (intPos - 1 >= strPrec.Length _
                                       AndAlso MatchEnv(Mid(strLemma, intPos - strPrec.Length, strPrec.Length), strPrec)) _
                               OrElse (intPos = 1 AndAlso strPrec = "#")) Then
                ' Check following context
                If (strFoll = "" OrElse (strLemma.Length - intPos - strMatch.Length >= strFoll.Length _
                                         AndAlso MatchEnv(Mid(strLemma, intPos + strMatch.Length, strFoll.Length), strFoll)) _
                                 OrElse (intPos + strMatch.Length - 1 = strLemma.Length AndAlso strFoll = "#")) Then
                  ' We have a match, so add a replacement of this one
                  strAlt = Left(strLemma, intPos - 1) & arChange(intI + 1) & Mid(strLemma, intPos + strMatch.Length)
                  colLemma.Add(strAlt)
                End If
              End If
            End If
          Next intPos
        Next intJ
      Next intI
      For intJ = 0 To colLemma.Count - 1
        ' Do we end in a vowel?
        If (InStr("aeiuoy", Right(colLemma.Item(intJ), 1)) > 0) Then
          ' Add variants with added "n"
          strLemma = colLemma.Item(intJ) & "n" : colLemma.Add(strLemma)
        End If
      Next intJ

      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MakeMEDlemmaList error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphDictAdaptVernPos
  ' Goal:   Check and possibly add an item into the VernPos table
  ' History:
  ' 29-04-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function MorphDictAdaptVernPos(ByVal strLv As String, ByVal strLf As String, ByVal strLemma As String, ByVal strLemmaOld As String, _
           ByVal bLemmaAmbi As Boolean, ByVal strMed As String, ByVal intMorphId As Integer, ByVal strF As String, ByRef intVPnum As Integer) As Boolean
    Dim dtrVP() As DataRow          ' Result of SELECT
    Dim dtrMorph() As DataRow       ' Result of SELECT
    Dim strFilter As String         ' Selection
    Dim dtrNew As DataRow = Nothing ' New row
    Dim intI As Integer             ' Counter
    Dim intJ As Integer             ' Counter

    Try
      ' Also make changes in the [VernPos] table
      If (strLf = "VBN") Then
        strFilter = "Vern = '" & strLv & "' AND (Pos = 'VBN' OR Pos = 'VAN') AND l = '" & strLemmaOld & "'"
      Else
        strFilter = "Vern = '" & strLv & "' AND Pos = '" & strLf & "' AND l = '" & strLemmaOld & "'"
      End If
      dtrVP = tdlMorphDict.Tables("VernPos").Select(strFilter)
      Select Case dtrVP.Length
        Case 0
          ' We can add this form right now...
          If (Not MorphAddOneVernPos(strLv, strLemma, strLf, strF, "AutoMED", intMorphId:=intMorphId)) Then Logging("MorphDictAdaptMED: could not add VernPos entry") : Return False
          intVPnum += 1
          ' Is this VBN?
          If (strLf = "VBN") Then
            If (Not MorphAddOneVernPos(strLv, strLemma, "VAN", strF, "AutoMED", intMorphId:=intMorphId)) Then Logging("MorphDictAdaptMED: could not add VernPos entry") : Return False
            intVPnum += 1
          End If
          ' Show what we have added
          Logging("Added VernPos [" & strLv & ", " & strLemma & ", " & strLf & ", " & strF & "]")
        Case Else
          ' There are multiple entries for this in the [VernPos] table
          ' Check and change the FIRST entry
          If (strLemma <> strLemmaOld) AndAlso (Not bLemmaAmbi) Then
            ' Is the lemma the same?
            If (strLemmaOld <> strLemma) Then
              ' Change this one's lemma
              dtrVP(0).Item("l") = strLemma
              ' Change all lemma's that need to be changed
              dtrMorph = tdlMorphDict.Tables("VernPos").Select("Pos LIKE 'V*' AND l='" & strLemmaOld & "'")
              For intJ = 0 To dtrMorph.Length - 1
                ' Change the lemma and the feature
                With dtrMorph(intJ)
                  .Item("l") = strLemma
                  ' Does this entry have a MED indicator?
                  strF = .Item("f").ToString
                  If (Not DoLike(strF, "*[OM]ED*")) Then AddSemiStack(strF, "s=" & strMed) : .Item("f") = strF
                End With
              Next intJ
            End If

          End If
          ' Insert the MED as feature
          strF = dtrVP(0).Item("f").ToString
          If (Not DoLike(strF, "*[OM]ED*")) Then
            AddSemiStack(strF, "s=" & strMed, True) : dtrVP(0).Item("f") = strF
          End If
          dtrVP(0).Item("Type") = IIf(bLemmaAmbi, "LemmaAmbi", "LemmaFeat")
          ' (Re)set the link to the MorphId
          dtrVP(0).Item("MorphId") = intMorphId
          '' Check "LinkTo" remarke
          'If (InStr(dtrVP(0).Item("Src").ToString, "LinkToMorphId") = 0) OrElse _
          '  (dtrVP(0).Item("Src").ToString <> "LinkToMorphId=" & intMorphId) Then
          '  ' Add the link here
          '  If (intMorphId < 0) Then
          '    ' Remove the link
          '    dtrVP(0).Item("Src") = "Estimate"
          '  Else
          '    ' Set the link
          '    dtrVP(0).Item("Src") = "LinkToMorphId=" & intMorphId
          '  End If
          'End If
          If (dtrVP.Length > 1) Then
            ' Delete all subsequent entries
            For intI = dtrVP.Length - 1 To 1 Step -1
              ' Delete this one
              dtrVP(intI).Delete()
            Next intI
            ' Accept changes
            tdlMorphDict.AcceptChanges()
          End If
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MorphDictAdaptVernPos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphDictCreateVernPos
  ' Goal:   Convert the "Morph" table to "VernPos"
  ' History:
  ' 04-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MorphDictCreateVernPos() As Boolean
    Dim strVern As String = ""    ' Vernacular
    Dim strPos As String = ""     ' POS
    Dim strLemma As String = ""   ' Lemma
    Dim intEntryId As Integer = 0 ' Link to Gutenberg lemma Id
    Dim strFeats As String = ""   ' Features
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim dtrMorph() As DataRow     ' Double checking
    Dim dtrThis As DataRow        ' One datarow
    Dim intPtc As Integer         ' Percentage
    Dim intDouble As Integer = 0   ' Number of double entries
    Dim intI As Integer           ' Counter

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Walk through "Morph" entries
      dtrFound = tdlMorphDict.Tables("Morph").Select("", "Vern ASC, Pos ASC, l ASC, EntryId ASC")
      For intI = 0 To dtrFound.Length - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ dtrFound.Length
        Status("Creating VernPos " & intPtc & "%", intPtc)
        ' Process this entry
        With dtrFound(intI)
          ' Check if this is a new one
          If (strVern = .Item("Vern").ToString.Replace("-", "")) AndAlso (strPos = .Item("Label").ToString) AndAlso _
             (strLemma = .Item("l").ToString) AndAlso _
             ((intEntryId < 0) OrElse .Item("EntryId").ToString = "" OrElse (intEntryId = CInt(.Item("EntryId").ToString))) Then
            ' Do not process this -- this is a double entry
            ' Stop
            intDouble += 1
          ElseIf (strVern = .Item("Vern").ToString.Replace("-", "")) AndAlso (strPos = .Item("Label").ToString) AndAlso _
           (strLemma = .Item("l").ToString) Then
            ' Do not even process if the entryid differs
            intDouble += 1
          Else
            ' We can copy the entry from [Morph] to [VernPos]...
            ' What are the characteristics?
            strVern = .Item("Vern").ToString : strPos = .Item("Label").ToString : strLemma = .Item("l").ToString
            strFeats = .Item("f").ToString
            ' Double check
            If (InStr(strVern, " ") = 0) AndAlso (InStr(strVern, ",") = 0) Then
              If (.Item("EntryId").ToString = "") Then
                intEntryId = -1
              Else
                intEntryId = CInt(.Item("EntryId").ToString)
              End If
              ' Check for these characteristics
              dtrMorph = tdlMorphDict.Tables("VernPos").Select("v='" & strVern.Replace("'", "''") & _
                                                               "' AND Pos='" & strPos & "' AND l='" & strLemma.Replace("'", "''") & "'")
              If (dtrMorph.Length = 0) Then
                ' Checking...
                'If (strPos <> "VB") Then
                '  Stop
                'End If
                dtrThis = AddOneDataRow(tdlMorphDict, "VernPos", "VernPosId", "VernPosList")
                With dtrThis
                  .Item("Vern") = strVern.Replace("-", "") : .Item("Pos") = strPos : .Item("Type") = IIf(strFeats = "", "LemmaOnly", "LemmaFeat")
                  .Item("l") = strLemma : .Item("f") = strFeats : .Item("Lev") = 1
                  ' Only provide "src" information if there is an EntryId link
                  If (intEntryId >= 0) Then .Item("Src") = "LinkToMorphId=" & intEntryId
                  .Item("v") = strVern
                End With
              Else
                ' Stop
              End If

            End If
          End If
        End With
      Next intI
      Status("Saving: " & strMorphDictFile)
      tdlMorphDict.WriteXml(strMorphDictFile)
      ' Show number of double entries
      If (intDouble > 0) Then
        Logging("MorphDictCreateVernPos: there were " & intDouble & " double entries skipped")
      End If
      Status("Ready")
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MorphDictCreateVernPos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphSpreadVernPos
  ' Goal:   Spread [VernPos] information into individual files
  ' Note:   Use 3rd argument of [OneMorphPropaDict()] to specify the particular forms that are being spread
  ' History:
  ' 04-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MorphSpreadVernPos(ByVal strDirIn As String, ByVal strDirOut As String, ByVal strLngAbbr As String, ByVal strDoTags As String) As Boolean
    Dim strInFile As String   ' Inmput file
    Dim strOutFile As String  ' Output file 
    Dim arInFile() As String  ' List of input files to be processed
    Dim intPtc As Integer     ' Percentage
    Dim intI As Integer       ' Counter

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
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
        Logging("Morphdict [VernPos] propagation [" & IO.Path.GetFileNameWithoutExtension(strInFile) & "] (" & intI + 1 & "/" & arInFile.Count & ")")
        ' Process this one file
        If (Not OneMorphPropaDict(strInFile, strOutFile, strLngAbbr, strDoTags)) Then
          ' Inform user
          Status("Error in processing: " & strInFile)
          Return False
        End If
      Next intI
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MorphDictCreateVernPos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   MorphPosToLemma
  ' Goal:   Given the vernacular and the POS, derive lemma's that should work (depending on the language period)
  ' History:
  ' 04-10-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function MorphPosToLemma(ByVal strPeriod As String, ByVal strVern As String, ByVal strPos As String, ByRef strLemma As String, _
                                  ByRef strFeat As String) As Boolean
    'Dim strPeriod As String   ' Language period

    Try
      '' Validate
      'If (pdxCurrentFile Is Nothing) Then Return False
      '' Try get epriod
      'strPeriod = GetPeriod(pdxCurrentFile)
      ' Initialise
      strFeat = ""
      ' Continue, depending on period
      Select Case strPeriod
        Case "ME"
          If (DoLike(strPos, strVerbBe)) Then strLemma = "be" : Return True
          If (DoLike(strPos, strVerbHv)) Then strLemma = "hauen" : Return True
          If (DoLike(strPos, strVerbDo)) Then strLemma = "don" : Return True
          If (DoLike(strPos, "MD*")) Then
            ' Depends on the particular modal we are looking at
            If (DoLike(strVern, "cu*|ca*|co*|k*")) Then strLemma = "kunnen" : Return True
            If (DoLike(strVern, "au*|ac*|ag*|ah*|aw*|ow*")) Then strLemma = "owen" : Return True
            If (DoLike(strVern, "d*")) Then strLemma = "dar" : Return True
            If (DoLike(strVern, "l*")) Then strLemma = "let" : Return True
            If (DoLike(strVern, "mo*t*")) Then strLemma = "mot" : Return True
            If (DoLike(strVern, "m*")) Then strLemma = "maei" : Return True
            If (DoLike(strVern, "o*t|ou*|og*")) Then strLemma = "oht" : Return True
            If (DoLike(strVern, "s*|x*")) Then strLemma = "schal" : Return True
            If (DoLike(strVern, "w*")) Then strLemma = "wille" : Return True
          End If
          If (DoLike(strPos, "NEG+MD*")) Then
            ' Depends on the particular modal we are looking at
            If (DoLike(strVern, "n*l*")) Then strLemma = "nelle" : Return True
            If (DoLike(strVern, "n*b*")) Then strLemma = "nabben" : Return True
          End If
        Case "eModE", "LmodE", "lModE", "ModE", "lmode", "emode"
          If (DoLike(strPos, strVerbBe)) Then strLemma = "be" : Return True
          If (DoLike(strPos, strVerbHv)) Then strLemma = "have" : Return True
          If (DoLike(strPos, strVerbDo)) Then strLemma = "do" : Return True
          If (DoLike(strPos, "MD*")) Then
            ' Depends on the particular modal we are looking at
            If (DoLike(strVern, "cu*|ca*|co*|k*")) Then strLemma = "can" : Return True
            If (DoLike(strVern, "l*")) Then strLemma = "let" : Return True
            If (DoLike(strVern, "mu*s*|m*st*")) Then strLemma = "must" : Return True
            If (DoLike(strVern, "sh*|s*a*l*|s*h*[ou]ld*")) Then strLemma = "shall" : Return True
            If (DoLike(strVern, "wi*|[yi]l|[iy]ll|'ll|vi*l|w[iy]*l*|'d|'ld|w[ou]*ld*|old*|owld*|ould*|oud*|twil*|twoul*|wo|wol|woll|wou*")) Then strLemma = "will" : Return True
            If (DoLike(strVern, "a*t*|ou*t*|ow*t*")) Then strLemma = "ought" : Return True
            If (DoLike(strVern, "mi*t*|m*[iy]*|mot*|mought*")) Then strLemma = "may" : Return True
            If (DoLike(strVern, "n*d*")) Then strLemma = "need" : Return True
            If (DoLike(strVern, "d*r*")) Then strLemma = "dare" : Return True
          End If

      End Select
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/MorphDictCreateVernPos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetMorphDictMatch
  ' Goal:   Get the best match for [strVern] with POS = [strPos]
  '         Method is defined by [strType]
  '         Return [strLemma] with accompanying [intScore] (between 0...100)
  ' History:
  ' 09-08-2013  ERK Created
  ' 10-03-2014  ERK Added [intFreq] parameter
  ' ------------------------------------------------------------------------------------------------------------
  Public Function GetMorphDictMatch(ByVal strVern As String, ByVal strPos As String, ByVal strType As String, ByRef strLemma As String, _
                                      ByRef intScore As Integer, ByRef strDerivation As String, ByRef strFeats As String) As Boolean
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim colRewr As New StringColl ' Collection of rewritten forms
    Dim strBest As String         ' Best lemma so far
    Dim strForm As String         ' Best form
    Dim strRewr As String         ' Rewriting
    Dim strFeatL As String        ' Lemma feature
    Dim strFeat As String = ""    ' String of features belonging to the best match
    Dim bValid As Boolean         ' Valid comparison
    Dim intBest As Integer        ' Index of best match
    Dim intConf As Integer        ' Score of best match
    Dim intFreq As Integer        ' Frequency of current one
    Dim intValue As Integer       ' Comparison value
    Dim intEquals As Integer      ' Number of equal letters
    Dim intPtc As Integer         ' Percentage
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter

    Try
      ' Validate
      If (tdlMorphDict Is Nothing) Then Return False
      ' Adapt POS
      If (strPos Like "*VBP*") Then
        strPos = "VBP"
      ElseIf (strPos Like "*VBD*") Then
        strPos = "VBD"
      ElseIf (strPos Like "*VBI*") Then
        strPos = "VBI"
      End If
      ' Start the derivation string
      strDerivation = "(no information)"
      ' Make sure we only accept vernaculars without hyphen
      If (InStr(strVern, "-") > 0) Then strVern = strVern.Replace("-", "")
      ' Check which type
      Select Case strType
        Case "Form"
          ' Consider all forms with the same POS
          ' Note: better to use feature @Pos than @Label
          ' dtrFound = tdlMorphDict.Tables("Morph").Select("Label = '" & strPos & "'", "")
          dtrFound = tdlMorphDict.Tables("Morph").Select("Pos = '" & strPos & "'", "")
          ' Find the one that matches best with [strVern]
          intBest = -1 : intConf = -1 : strBest = "" : strForm = "" : strFeat = ""
          For intI = 0 To dtrFound.Length - 1
            ' Check the validity of this entry: 
            bValid = True
            If (DoLike(strPos, strAnyVerb)) Then
              strFeatL = dtrFound(intI).Item("l").ToString
              ' Check that this lemma occurs as VB in [Morph]
              bValid = (tdlMorphDict.Tables("Morph").Select("Pos='VB' AND Vern='" & strFeatL.Replace("'", "''") & "'").Length > 0)
              ' ======= DEBUG ========
              'If (Not bValid) Then Stop
              ' ======================
            End If
            ' Only continue if valid
            If (bValid) Then
              ' Compare this form with [strVern]
              intValue = objSim.LevenDist(dtrFound(intI).Item("Vern").ToString, strVern, intEquals)
              ' Adjust the value if it is an obsolete form
              If (InStr(dtrFound(intI).Item("f").ToString, "=obs") > 0) Then intValue -= 3
              ' Adjust value for initial letter mismatch
              If (Left(dtrFound(intI).Item("l").ToString, 1) <> Left(strVern, 1)) Then intValue -= 3
              ' Beware of negatives
              If (intValue < 0) Then intValue = 0
              ' Process best value
              If (intValue > intConf) Then
                intConf = intValue : strBest = dtrFound(intI).Item("l")
                strForm = dtrFound(intI).Item("Vern") : strFeat = dtrFound(intI).Item("f").ToString
              End If
            End If
          Next intI
          ' Prepare the result
          If (strBest = "") Then
            ' There is no result
            intScore = 0 : strLemma = "" : strFeats = "(none)" : strDerivation = "(none)"
          Else
            ' We have a result
            intScore = intConf : strLemma = strBest : strDerivation = "MatchForm= " & strPos & "[" & strForm & "]"
            strFeats = strFeat
          End If
        Case "Rewrite"
          ' Get a collection of rewritten forms
          If (Not FormToLemma(strVern, strPos, colRewr)) Then Return False
          ' Walk the collection and try finding the best match with lemma's in the [Morph] table
          ' Try them all
          intBest = -1 : intConf = -1 : strBest = "" : strForm = "" : strRewr = "" : strFeat = ""
          For intI = 0 To colRewr.Count - 1
            ' Can we find this rewritten lemma as lemma in the Morph table?
            ' dtrFound = tdlMorphDict.Tables("Morph").Select("Label = 'VB'")
            dtrFound = tdlMorphDict.Tables("Morph").Select("Pos = 'VB'")
            For intJ = 0 To dtrFound.Length - 1
              ' Compare this form with [strVern]
              intValue = objSim.LevenDist(dtrFound(intJ).Item("Vern").ToString, colRewr.Item(intI), intEquals)
              ' Adjust the value of this form with the factor in colRewr.Exmp(intI)
              intValue -= (10 - colRewr.Exmp(intI))
              ' Adjust the value if it is an obsolete form
              If (InStr(dtrFound(intJ).Item("f").ToString, "=obs") > 0) Then intValue -= 3
              ' Adjust value for initial letter mismatch
              If (Left(dtrFound(intJ).Item("l").ToString, 1) <> Left(strVern, 1)) Then intValue -= 3
              ' Beware of negatives
              If (intValue < 0) Then intValue = 0
              ' Check which is better
              If (intValue > intConf) Then
                ' highest value is best - first come, first go
                intConf = intValue : strBest = dtrFound(intJ).Item("l").ToString : strForm = dtrFound(intJ).Item("Vern") : strRewr = colRewr.Item(intI) : strFeat = dtrFound(intJ).Item("f").ToString
              ElseIf (intValue = intConf) AndAlso (InStr(dtrFound(intJ).Item("f").ToString, "=obs") = 0) Then
                ' Give preference to those with the same value, but without "obs" feature
                ' Take this to be the best one
                intConf = intValue : strBest = dtrFound(intJ).Item("l").ToString : strForm = dtrFound(intJ).Item("Vern") : strRewr = colRewr.Item(intI) : strFeat = dtrFound(intJ).Item("f").ToString
              End If
            Next intJ
          Next intI
          ' Prepare the result
          If (strBest = "") Then
            ' There is no result
            intScore = 0 : strLemma = "" : strFeats = "(none)" : strDerivation = "(none)"
          Else
            ' We have a result
            intScore = intConf : strLemma = strBest : strDerivation = "MatchRewr=[" & strRewr & "/" & strForm & "]"
            strFeats = strFeat
            ' ======== DEBUG =======
            ' If (strForm <> strBest) Then Stop
            ' ======================
          End If
        Case Else
          ' Return failure
          Return False
      End Select
      ' Return success anyway
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLemmatize/GetMorphDictMatch error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   FormToLemma
  ' Goal:   Get all possible prefix + suffix rewrites for me
  ' History:
  ' 09-05-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function FormToLemma(ByVal strVern As String, ByVal strPos As String, ByRef colRew As StringColl) As Boolean
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
      If (strVern = "") OrElse (strPos = "") Then Return False
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
          dtrRewri = tdlMorphDict.Tables("Derive").Select("Pos LIKE '" & Left(strPos, 3) & "*' AND Freq > 2", "Freq DESC")
        ElseIf (strPos Like "WAD[JV]*") Then
          ' Use adapted POS search
          dtrRewri = tdlMorphDict.Tables("Derive").Select("Pos LIKE '" & Left(strPos, 4) & "*' AND Freq > 2", "Freq DESC")
        Else
          ' Find all possible suffix-rewrite rules for this POS
          dtrRewri = tdlMorphDict.Tables("Derive").Select("Pos='" & strPos & "' AND Freq > 2", "Freq DESC")
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
          colRew.AddUnique(strRewr, REW_BEST)
          ' Possibly add dh <> th rewrites
          If (InStr(strRewr, "dh") > 0) Then
            colRew.AddUnique(strRewr.Replace("dh", "th"), REW_SPCHG)
          ElseIf (InStr(strRewr, "th") > 0) Then
            colRew.AddUnique(strRewr.Replace("th", "dh"), REW_SPCHG)
          ElseIf (InStr(strRewr, "ui") > 0) Then
            colRew.AddUnique(strRewr.Replace("ui", "wi"), REW_SPCHG)
            colRew.AddUnique(strRewr.Replace("ui", "fi"), REW_SPCHG)
            colRew.AddUnique(strRewr.Replace("ui", "vi"), REW_SPCHG)
          ElseIf (InStr(strRewr, "yng") > 0) Then
            colRew.AddUnique(strRewr.Replace("yng", "ing"), REW_SPCHG)
          ElseIf (InStr(strRewr, "uing") > 0) Then
            colRew.AddUnique(strRewr.Replace("uing", "ving"), REW_SPCHG)
          ElseIf (InStr(strRewr, "gg") > 0) Then
            colRew.AddUnique(strRewr.Replace("gg", "cg"), REW_SPCHG)
          ElseIf (InStr(strRewr, "k") > 0) Then
            colRew.AddUnique(strRewr.Replace("k", "c"), REW_SPCHG)
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
              colRew.AddUnique(strRewr2, REW_SPCHG)
              ' Possibly add dh <> th rewrites
              If (InStr(strRewr2, "dh") > 0) Then
                colRew.AddUnique(strRewr2.Replace("dh", "th"), REW_SPCHG)
              ElseIf (InStr(strRewr2, "th") > 0) Then
                colRew.AddUnique(strRewr2.Replace("th", "dh"), REW_SPCHG)
              ElseIf (InStr(strRewr2, "ui") > 0) Then
                colRew.AddUnique(strRewr2.Replace("ui", "wi"), REW_SPCHG)
                colRew.AddUnique(strRewr2.Replace("ui", "fi"), REW_SPCHG)
                colRew.AddUnique(strRewr2.Replace("ui", "vi"), REW_SPCHG)
              ElseIf (InStr(strRewr2, "yng") > 0) Then
                colRew.AddUnique(strRewr2.Replace("yng", "ing"), REW_SPCHG)
              ElseIf (InStr(strRewr2, "uing") > 0) Then
                colRew.AddUnique(strRewr2.Replace("uing", "ving"), REW_SPCHG)
              ElseIf (InStr(strRewr2, "gg") > 0) Then
                colRew.AddUnique(strRewr2.Replace("gg", "cg"), REW_SPCHG)
              ElseIf (InStr(strRewr2, "k") > 0) Then
                colRew.AddUnique(strRewr2.Replace("k", "c"), REW_SPCHG)
              End If
            End If
          Next intI
        End If
      Next intJ
      ' Add possible prefix rewrites for the unchanged form
      strRewr = strVern
      ' Possibly add dh <> th rewrites
      If (InStr(strRewr, "dh") > 0) Then
        colRew.AddUnique(strRewr.Replace("dh", "th"), REW_NOCHG)
      ElseIf (InStr(strRewr, "th") > 0) Then
        colRew.AddUnique(strRewr.Replace("th", "dh"), REW_NOCHG)
      ElseIf (InStr(strRewr, "ui") > 0) Then
        colRew.AddUnique(strRewr.Replace("ui", "wi"), REW_NOCHG)
        colRew.AddUnique(strRewr.Replace("ui", "fi"), REW_NOCHG)
        colRew.AddUnique(strRewr.Replace("ui", "vi"), REW_NOCHG)
      ElseIf (InStr(strRewr, "yng") > 0) Then
        colRew.AddUnique(strRewr.Replace("yng", "ing"), REW_NOCHG)
      ElseIf (InStr(strRewr, "uing") > 0) Then
        colRew.AddUnique(strRewr.Replace("uing", "ving"), REW_NOCHG)
      ElseIf (InStr(strRewr, "gg") > 0) Then
        colRew.AddUnique(strRewr.Replace("gg", "cg"), REW_NOCHG)
      ElseIf (InStr(strRewr, "k") > 0) Then
        colRew.AddUnique(strRewr.Replace("k", "c"), REW_NOCHG)
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
          colRew.AddUnique(strRewr2, REW_WORSE)
          ' Possibly add dh <> th rewrites
          If (InStr(strRewr2, "dh") > 0) Then
            colRew.AddUnique(strRewr2.Replace("dh", "th"), REW_WORSE)
          ElseIf (InStr(strRewr2, "th") > 0) Then
            colRew.AddUnique(strRewr2.Replace("th", "dh"), REW_WORSE)
          ElseIf (InStr(strRewr2, "ui") > 0) Then
            colRew.AddUnique(strRewr2.Replace("ui", "wi"), REW_WORSE)
            colRew.AddUnique(strRewr2.Replace("ui", "fi"), REW_WORSE)
            colRew.AddUnique(strRewr2.Replace("ui", "vi"), REW_WORSE)
          ElseIf (InStr(strRewr2, "yng") > 0) Then
            colRew.AddUnique(strRewr2.Replace("yng", "ing"), REW_WORSE)
          ElseIf (InStr(strRewr2, "uing") > 0) Then
            colRew.AddUnique(strRewr2.Replace("uing", "ving"), REW_WORSE)
          ElseIf (InStr(strRewr2, "gg") > 0) Then
            colRew.AddUnique(strRewr2.Replace("gg", "cg"), REW_WORSE)
          ElseIf (InStr(strRewr2, "k") > 0) Then
            colRew.AddUnique(strRewr2.Replace("k", "c"), REW_WORSE)
          End If
        End If
      Next intI
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/FormToLemma error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitWebsterPos
  ' Goal:   Read the "WebsterPos.csv" list of translations from Cat --> POS
  ' History:
  ' 24-12-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitWebsterPos() As Boolean
    Dim arText() As String  ' File contents
    Dim arLine() As String  ' One line
    Dim intI As Integer     ' Counter

    Try
      If (bWebsterPos) Then Return True
      ' Validate
      If (Not IO.File.Exists(loc_strWebsterPosFile)) Then Return False
      ' Read the file
      arText = IO.File.ReadAllLines(loc_strWebsterPosFile)
      ' Create space
      ReDim arWebster(0 To arText.Length - 1)
      ' Fill the array
      For intI = 0 To arText.Length - 1
        ' Read this entry
        arLine = Split(arText(intI), vbTab)
        ' Fill entry
        With arWebster(intI)
          .cat = arLine(0) : .examp = arLine(1) : .pos = arLine(2) : .feats = arLine(3)
        End With
      Next intI
      ' Return positively
      bWebsterPos = True
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/InitWebsterPos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetWebsterPosEntry
  ' Goal:   Retrieve the category
  ' History:
  ' 24-12-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetWebsterPosEntry(ByVal strCat As String) As Integer
    Dim intI As Integer   ' Counter

    Try
      ' Validate
      If (Not bWebsterPos) Then Return -1
      ' Reduce number of spaces
      strCat = Trim(strCat)
      While (InStr(strCat, "  ") > 0)
        strCat = strCat.Replace("  ", " ")
      End While
      ' Find this entry
      For intI = 0 To arWebster.Length - 1
        If (arWebster(intI).cat = strCat) Then Return intI
      Next intI
      ' Return failure
      Return -1
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/GetWebsterPosEntry error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DoWebsterPos
  ' Goal:   Given the category in [strPos], get the Webstertable equivalent 
  '           in combinations of POS and features
  ' History:
  ' 24-12-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function DoWebsterPos(ByVal strPos As String, ByRef arPos() As String) As Boolean
    Dim strFeats As String  ' Features
    Dim intW As Integer     ' Index
    Dim intI As Integer     ' Counter

    Try
      ' Validate
      If (Not bWebsterPos) Then Return False
      ' Get index
      intW = GetWebsterPosEntry(strPos)
      If (intW < 0) Then Return False
      ' Process entry
      With arWebster(intW)
        strFeats = .feats
        arPos = Split(.pos, ";")
        ' Add the feature everywhere
        For intI = 0 To arPos.Length - 1
          arPos(intI) &= IIf(strFeats = "", "", ";" & strFeats)
        Next intI
        ' NO, don't do this!!
        ' strPos = arPos(0)
      End With
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/DoWebsterPos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  LemmaInitDict
  ' Goal :  Try initialise/load dictionary
  ' History:
  ' 31-01-2014  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function LemmaInitDict() As Boolean
    Dim strLang As String = "OE"      ' Language
    Dim strUrl As String = "http://erwinkomen.ruhosting.nl/software/lng/eng_hist/"
    Dim strDict As String             ' Name of dictionary file

    Try
      ' Check presence of dictionary
      If (tdlOEdict Is Nothing) Then
        ' Make dictionary file name
        ' strDict = "d:\data files\corpora\dictionaries\" & strLang & "dict_out.xml"
        strDict = GetSetDir() & "\" & strLang & "dict_out.xml"
        ' Does it exist??
        If (Not IO.File.Exists(strDict)) Then
          ' Try to download it
          strUrl &= strLang & "dict_out.xml"
          Status("Downloading dictionary (for the first time)...")
          If (Not DownloadFile(strUrl, strDict)) Then Return False
        End If
        Status("Reading dictionary of " & strLang)
        If (Not ReadDataset("OEdict.xsd", strDict, tdlOEdict)) Then Return False
      End If
      Return (tdlOEdict IsNot Nothing)
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/LemmaInitDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  LemmaOneWord
  ' Goal :  Find information for one word
  ' History:
  ' 31-01-2014  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function LemmaOneWord(ByRef ndxThis As XmlNode, ByRef strLemma As String, _
                               ByRef strDef As String) As Boolean
    Dim strLabel As String        ' Label (POS)
    Dim strLang As String = ""    ' Language
    Dim strDict As String = ""    ' Dictionary entry pointer
    Dim ndxLang As XmlNode        ' Language definition node
    Dim strUrl As String = ""     ' URL we want to go to
    Dim strDname As String = ""   ' Dictionary name
    Dim strNum As String = ""     ' Number
    Dim strPage As String = ""    ' Web page
    Dim strLine As String = ""    ' One line
    Dim strWord As String = ""    ' The word
    Dim ndxWord As XmlNode = Nothing
    Dim data As IO.Stream = Nothing  '
    Dim rdThis As IO.StreamReader = Nothing
    Dim intI As Integer           ' Counter
    Dim elcThis As HtmlElementCollection
    Dim bContent As Boolean = False ' Have content or not?

    Try
      ' Validate
      If (ndxThis Is Nothing) OrElse (ndxThis.Name <> "eTree") Then Return False
      strWord = ""
      ndxWord = ndxThis.SelectSingleNode("./descendant::eLeaf[@Type='Vern'][1]")
      If (ndxWord IsNot Nothing) Then
        strWord = ndxWord.Attributes("Text").Value
      End If
      ' Try to determine the language
      ndxLang = pdxCurrentFile.SelectSingleNode("./descendant::language")
      If (ndxLang IsNot Nothing) Then
        strLang = ndxLang.Attributes("ident").Value
      End If
      ' Find out which dictionary to load
      Select Case LCase(strLang)
        Case "nld"
        Case "eng_hist", "oe"
          ' Dictionary...
          If (Not LemmaInitDict()) Then Return False
      End Select
      ' Get the label
      strLabel = ndxThis.Attributes("Label").Value
      ' Try to get lemma
      strLemma = GetFeature(ndxThis, "M", "l")
      ' Try get definition
      strDef = GetFeature(ndxThis, "M", "d")
      If (strDef = "") Then strDef = LemmaToDef(strLang, strLemma, False)
      ' Get dictionary entry pointer
      strDict = GetFeature(ndxThis, "M", "s")
      ' Get the first element (if more than one)
      strDict = Split(strDict, "|")(0)
      ' Derive the URL
      strDname = Regex.Match(strDict, "[A-Z]+").Value
      If (strDname <> "") Then
        strNum = Mid(strDict, strDname.Length + 1)
        ' What is this?
        Select Case strDname
          Case "BT"
            ' Check the number of elements
            If (strNum.Length <> 6) AndAlso (IsNumeric(strNum)) Then
              strNum = Format(CInt(strNum), "000000")
            End If
        End Select
      End If
      intI = Array.FindIndex(arDictAbbr, Function(strName As String) strName = strDname)
      If (strDict <> "") AndAlso (strNum <> "") AndAlso (intI >= 0) Then
        ' Find and go to this entry
        strUrl = arDictUrl(intI) & strNum
        ' Navigate to it
        With loc_wbDom
          .ScriptErrorsSuppressed = True
          .Navigate(strUrl)
          While (.ReadyState <> System.Windows.Forms.WebBrowserReadyState.Complete)
            Status("Looking for " & strLemma)
            Application.DoEvents()
          End While
        End With
        ' Access the page
        With loc_wbDom.Document
          elcThis = .GetElementsByTagName("div")
          For intI = 0 To elcThis.Count - 1
            ' Debug.Print("Div #=" & intI & " class=" & elcThis(intI).GetAttribute("classname"))
            If (elcThis(intI).GetAttribute("id") = "entry-content") Then
              ' Stop
              ' End If
              ' If (elcThis(intI).GetAttribute("classname") Like "content *") Then
              ' We've got the content element

              strPage = "<html><head><style>" & loc_strCss & "</style></head><body>" & elcThis(intI).OuterHtml & "</body></html>"
              bContent = True : Exit For
            End If
          Next intI
        End With
        ' Do we have content?
        If (bContent) Then
          frmMain.wbEdtDict.DocumentText = strPage
          Status(strDict & ": " & strLemma)
        ElseIf (strLemma = "") Then
          frmMain.wbEdtDict.DocumentText = "Lemma is not known"
          Status("Lemma is not known")
        Else
          frmMain.wbEdtDict.DocumentText = "No entry: " & strLemma
          Status("No entry: " & strLemma)
        End If
      Else
        Status("Lemma not known: " & strWord)
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/LemmaOneWord error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  GetDictEntryHtml
  ' Goal :  Given a "s" feature value, get the text of the source URL
  ' History:
  ' 30-01-2014  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function GetDictEntryHtml(ByVal strLemma As String, ByVal strDict As String) As String
    Dim strLang As String = ""    ' Language
    ' Dim strDict As String = ""    ' Dictionary entry pointer
    Dim strUrl As String = ""     ' URL we want to go to
    Dim strDname As String = ""   ' Dictionary name
    Dim strNum As String = ""     ' Number
    Dim strPage As String = ""    ' Web page
    Dim strLine As String = ""    ' One line
    Dim strWord As String = ""    ' The word
    Dim data As IO.Stream = Nothing  '
    Dim rdThis As IO.StreamReader = Nothing
    Dim intI As Integer           ' Counter
    Dim elcThis As HtmlElementCollection
    Dim bContent As Boolean = False ' Have content or not?

    Try
      ' Get the first element (if more than one)
      strDict = Split(strDict, "|")(0)
      ' Derive the URL
      strDname = Regex.Match(strDict, "[A-Z]+").Value
      If (strDname <> "") Then
        strNum = Mid(strDict, strDname.Length + 1)
        ' What is this?
        Select Case strDname
          Case "BT"
            ' Check the number of elements
            If (strNum.Length <> 6) AndAlso (IsNumeric(strNum)) Then
              strNum = Format(CInt(strNum), "000000")
            End If
        End Select
      End If
      intI = Array.FindIndex(arDictAbbr, Function(strName As String) strName = strDname)
      If (strDict <> "") AndAlso (strNum <> "") AndAlso (intI >= 0) Then
        ' Find and go to this entry
        strUrl = arDictUrl(intI) & strNum
        ' Navigate to it
        With loc_wbDom
          .ScriptErrorsSuppressed = True
          .Navigate(strUrl)
          While (.ReadyState <> System.Windows.Forms.WebBrowserReadyState.Complete)
            If (strLemma = "") Then
              Status("Looking for the lemma of " & strDict)
            Else
              Status("Looking for " & strLemma)
            End If
            Application.DoEvents()
          End While
        End With
        ' Access the page
        With loc_wbDom.Document
          elcThis = .GetElementsByTagName("div")
          For intI = 0 To elcThis.Count - 1
            ' Debug.Print("Div #=" & intI & " class=" & elcThis(intI).GetAttribute("classname"))
            If (elcThis(intI).GetAttribute("id") = "entry-content") Then
              ' Stop
              ' End If
              ' If (elcThis(intI).GetAttribute("classname") Like "content *") Then
              ' We've got the content element

              strPage = "<html><head><style>" & loc_strCss & "</style></head><body>" & elcThis(intI).OuterHtml & "</body></html>"
              bContent = True : Exit For
            End If
          Next intI
        End With
        ' Do we have content?
        If (bContent) Then
          Status(strDict & ": " & strLemma)
        ElseIf (strLemma = "") Then
          strPage = "Lemma is not known"
          Status("Lemma is not known")
        Else
          strPage = "No entry: " & strLemma & "=" & strDict
          Status("No entry: " & strLemma & "=" & strDict)
        End If
      Else
        strPage = "Lemma not known"
        Status("Lemma not known: " & strWord & "=" & strDict)
      End If
      ' Return the html document text
      Return strPage
    Catch ex As Exception
      ' Show error
      HandleErr("modMorph/GetDictEntryHtml error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  LemmaFileView
  ' Goal :  Show lemma's of the current file
  ' History:
  ' 30-01-2014  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function LemmaFileView(ByVal bGiveGloss As Boolean, ByRef strFile As String) As Boolean
    Dim strLoc As String = ""         ' Location ID
    Dim strSeg As String = ""         ' Text of one line
    Dim strGls As String = ""         ' Gloss of one line
    Dim strText As String = ""        ' Textof the file
    Dim strDict As String             ' Name of dictionary file
    Dim strPos As String              ' POS
    Dim ndxFor As XmlNode = Nothing   ' <forest> node
    Dim ndxThis As XmlNode            ' <eTree> node above the <eLeaf>
    Dim ndxList As XmlNodeList        ' List of nodes
    Dim strLemma As String            ' Lemma
    Dim strFeats As String            ' Features
    Dim strLabel As String            ' Label
    Dim strDef As String              ' Definition (first few words)
    Dim strLang As String = "OE"      ' Language
    Dim intI As Integer               ' Counter
    Dim intPtc As Integer             ' Percentage
    Dim intCount As Integer           ' Number of <forest> lines
    Dim intSct As Integer = 0         ' Section number
    Dim intJ As Integer               ' Counter
    Dim colBack As New StringColl

    Try
      ' Dictionary...
      If (Not LemmaInitDict()) Then Return False
      ' Adapt the file to the actual one
      strFile = IO.Path.GetDirectoryName(strCurrentFile) & "\" & _
        IO.Path.GetFileNameWithoutExtension(strCurrentFile) & "-Lemmas.html"
      ' Find first forest element - depending on the project type
      If (Not GetFirstForest(pdxCurrentFile, ndxFor)) Then Status("ViewFile error: cannot find first line") : Exit Function
      ' Get number of nodes
      intCount = ndxFor.ParentNode.ChildNodes.Count
      ' Start html output
      colBack.Add("<html><body>")
      ' Step through them
      intI = 1
      While (Not ndxFor Is Nothing)
        ' Show where we are
        intPtc = intI * 100 \ intCount
        Status("Loading " & strFile & " " & intPtc & "%", intPtc)
        ' Get the current location
        If (Not GetForestLoc(ndxFor, strLoc)) Then Status("ViewFile error: cannot find location of line " & intI) : Exit Function
        ' Do we need to add a section break?
        If (Not ndxFor.Attributes("Section") Is Nothing) Then
          ' Increment section number
          intSct += 1
          ' Add a section break
          colBack.Add("<p> --- Section " & intSct & " --- <p>")
        End If
        ' Find the 'varnacular' text of this node
        strSeg = GetSeg(ndxFor, "org")
        ' Add line to collection
        colBack.Add("<div><font color='blue' size='1'>[" & strLoc & "] </font>" & strSeg & " </div>")
        ' Go through all the words
        ndxList = ndxFor.SelectNodes("./descendant::eLeaf")
        colBack.Add("<div><table><tr><td>Word</td><td>POS</td><td>Lemma</td><td>Features</td></td>Definition</td></tr>")
        For intJ = 0 To ndxList.Count - 1
          ndxThis = ndxList(intJ).ParentNode : strLabel = ndxThis.Attributes("Label").Value
          ' Skip certain labels
          If (Not DoLike(strLabel, "CODE|META")) AndAlso (Not DoLike(ndxList(intJ).Attributes("Type").Value, "Star|Zero")) Then
            strFeats = GetFeatVect(ndxThis, "l", "h") : strLemma = GetFeature(ndxThis, "M", "l")
            ' Try get definition
            strDef = LemmaToDef(strLang, strLemma, True)
            ' Put it all together
            colBack.Add("<tr><td valign='top'>" & VernToEnglish(ndxList(intJ).Attributes("Text").Value) & "</td>" & _
                        "<td valign='top'><font size='1' color='blue'>" & strLabel & "</font></td><td valign='top'><font color='red'>" & strLemma & "</font></td>" & _
                        "<td valign='top'><font size='1' color='green'>" & strFeats.Replace("|", ", ") & "</font></td><td valign='top'>" & strDef & "</td></tr>")
          End If
        Next intJ
        ' Finish table
        colBack.Add("</table></div>")
        ' GO to next forest
        ndxFor = ndxFor.NextSibling : intI += 1
      End While
      ' Finish HTML output
      colBack.Add("</body></html>")
      ' Save the html output
      strText = colBack.Text
      IO.File.WriteAllText(strFile, strText, System.Text.Encoding.UTF8)
      Status("File has been saved: " & strFile)
      Return True
    Catch ex As Exception
      ' Warn user
      HandleErr("modLemmatize/LemmaFileView error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Reset the initialisation
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  LemmaToDef
  ' Goal :  Retrieve the definition of the lemma given
  ' History:
  ' 30-01-2014  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function LemmaToDef(ByVal strLang As String, ByVal strLemma As String, ByVal bHtml As Boolean) As String
    Dim dtrFound() As DataRow ' Result of selecting
    Dim colBack As New StringColl
    Dim intI As Integer       ' Counter

    Try
      ' Validate
      If (tdlOEdict Is Nothing) Then Return ""
      ' Try find it
      dtrFound = tdlOEdict.Tables("Sense").Select("l='" & strLemma & "'")
      If (dtrFound.Length > 0) Then
        ' Get all the definitions
        For intI = 0 To dtrFound.Length - 1
          If (bHtml) Then
            colBack.Add("<b>" & intI + 1 & "</b>: " & Left(dtrFound(intI).Item("Def").ToString, 80))
          Else
            colBack.Add(intI + 1 & ": " & Left(dtrFound(intI).Item("Def").ToString, 80))
          End If
        Next intI
        ' Return the result
        Return colBack.Text
      End If
      ' Default returning
      Return "-"
    Catch ex As Exception
      ' Warn user
      HandleErr("modLemmatize/LemmaToDef error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return faiulre
      Return "-"
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   UnaccLemmaCollect
  ' Goal:   Collect lemma's of verbs with VbType = "unacc"
  ' History:
  ' 13-02-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function UnaccLemmaCollect(ByVal strDirIn As String, ByRef colLemma As StringColl) As Boolean
    Dim arInFile() As String      ' List of files
    Dim strInFile As String       ' Oone file
    Dim strHtml As String = ""    ' Output
    Dim strVbType As String = ""  ' Verb type feature value
    Dim strLemma As String = ""   ' Lemma feature
    Dim strS As String = ""       ' The source number of the lemma 
    Dim pdxThis As XmlDocument    ' One document
    Dim ndxForest As XmlNode      ' One forest
    Dim ndxList As XmlNodeList    ' List of nodes
    Dim colHtml As New StringColl ' Gather what we return
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intD As Integer = 0       ' Number done
    Dim intDone As Integer = 0    ' Total done
    Dim intPtc As Integer         ' Percentage

    Try
      ' Consider all *.psdx input files
      arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories) : pdxThis = Nothing : ndxForest = Nothing
      ' Prepare tag-report anyway
      Status("Preparing unacc lemma collection...")
      For intI = 0 To arInFile.Count - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ arInFile.Count
        Status("Processing " & intPtc & "%", intPtc)
        ' Get the input file name
        strInFile = arInFile(intI) : intD = 0
        ' Also show it on the main form
        Logging("Unacc lemma collecting [" & IO.Path.GetFileNameWithoutExtension(strInFile) & "] (" & intI + 1 & "/" & arInFile.Count & ")")
        ' Load this document
        If (Not ReadXmlDoc(strInFile, pdxThis)) Then Logging("Could not load [" & strInFile & "]") : Return False
        ' Walk through forests
        If (Not GetFirstForest(pdxThis, ndxForest)) Then Logging("Could not find forest") : Return False
        While (ndxForest IsNot Nothing)
          ' Find all constituents here that need treating
          ndxList = ndxForest.SelectNodes("./descendant::eTree[tb:matches(@Label, '" & strAnyVerb & "')]", conTb)
          Application.DoEvents()
          For intJ = 0 To ndxList.Count - 1
            ' Get the value of its VbType feature
            strVbType = GetFeature(ndxList(intJ), "Vb", "VbType")
            ' Is this unacc?
            If (strVbType = "unacc") Then
              ' Get its lemma feature
              strLemma = GetFeature(ndxList(intJ), "M", "l")
              strS = GetFeature(ndxList(intJ), "M", "s")
              ' Only accept non-empty lemma's that are NOT ambiguous (if they are, they have a ; inside)
              If (strLemma <> "") AndAlso (InStr(strLemma, ";") = 0) Then
                ' Add lemma to list
                If (strS = "") Then
                  colLemma.AddUnique(strLemma) : intD += 1
                Else
                  colLemma.AddUnique(strLemma & "_" & strS) : intD += 1
                End If
              End If
            End If
          Next intJ
          ' Find next forest
          ndxForest = ndxForest.NextSibling
        End While
        ' Add totals
        intDone += intD
      Next intI
      ' Return success
      Logging("The number of lemma's with feature [unacc] is: " & colLemma.Count)
      Status("Collected unacc lemma's")
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/UnaccLemmaCollect error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   UnaccLemmaSpread
  ' Goal:   Spread "unacc" feature to all forms of verbs with lemma's on the [colLemma] list
  ' History:
  ' 13-02-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function UnaccLemmaSpread(ByVal strDirIn As String, ByRef colLemma As StringColl) As Boolean
    Dim arInFile() As String      ' List of files
    Dim strInFile As String       ' Oone file
    Dim strHtml As String = ""    ' Output
    Dim strVbType As String = ""  ' Verb type feature value
    Dim strLemma As String = ""   ' Lemma feature
    Dim strS As String = ""       ' Source feature
    Dim strLemmaS As String = ""  ' Combination of lemma and s
    Dim pdxThis As XmlDocument    ' One document
    Dim ndxForest As XmlNode      ' One forest
    Dim ndxList As XmlNodeList    ' List of nodes
    Dim colHtml As New StringColl ' Gather what we return
    Dim bChanged As Boolean       ' File has changed and needs saving
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intD As Integer = 0       ' Number done
    Dim intDone As Integer = 0    ' Total done
    Dim intPtc As Integer         ' Percentage

    Try
      ' Consider all *.psdx input files
      arInFile = IO.Directory.GetFiles(strDirIn, "*.psdx", IO.SearchOption.AllDirectories) : pdxThis = Nothing : ndxForest = Nothing
      ' Prepare tag-report anyway
      Status("Preparing unacc lemma spreading...")
      For intI = 0 To arInFile.Count - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ arInFile.Count
        Status("Processing " & intPtc & "%", intPtc)
        ' Get the input file name
        strInFile = arInFile(intI) : intD = 0 : bChanged = False
        ' Also show it on the main form
        Logging("Unacc lemma spreading [" & IO.Path.GetFileNameWithoutExtension(strInFile) & "] (" & intI + 1 & "/" & arInFile.Count & ")")
        ' Load this document
        If (Not ReadXmlDoc(strInFile, pdxThis)) Then Logging("Could not load [" & strInFile & "]") : Return False
        ' Walk through forests
        If (Not GetFirstForest(pdxThis, ndxForest)) Then Logging("Could not find forest") : Return False
        While (ndxForest IsNot Nothing)
          ' Find all constituents with a verb form
          ndxList = ndxForest.SelectNodes("./descendant::eTree[tb:matches(@Label, '" & strAnyVerb & "')]", conTb)
          For intJ = 0 To ndxList.Count - 1
            ' Get the value of its VbType feature
            strVbType = GetFeature(ndxList(intJ), "Vb", "VbType")
            ' Is this not (yet) unacc?
            If (strVbType <> "unacc") Then
              ' Get its lemma feature
              strLemma = GetFeature(ndxList(intJ), "M", "l")
              strS = GetFeature(ndxList(intJ), "M", "s")
              If (strLemma <> "") AndAlso (InStr(strLemma, ";") = 0) Then
                ' Combine lemma and s
                If (strS = "") Then
                  strLemmaS = strLemma
                Else
                  strLemmaS = strLemma & "_" & strS
                End If
                ' Check if lemma is on the unacc lemma list
                If (colLemma.Exists(strLemmaS)) Then
                  ' The lemma is on the "unacc" lemma list, so the verb type should be changed
                  AddFeature(pdxThis, ndxList(intJ), "Vb", "VbType", "unacc")
                  intD += 1 : bChanged = True
                End If
              End If
            End If
          Next intJ
          ' Find next forest
          ndxForest = ndxForest.NextSibling
        End While
        ' Add totals
        intDone += intD
        ' Need changing?
        If (bChanged) Then
          ' Save the file
          Status("Saving changed: " & strInFile & " (" & intD & " changes)")
          Logging(vbTab & " " & intD & " changes")
          pdxThis.Save(strInFile)
        End If
      Next intI
      ' Return success
      Logging("The verb type [unacc] has been spread to " & intDone & " verb forms")
      Status("Finished spreading unacc lemma's")
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/UnaccLemmaSpread error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
End Module
