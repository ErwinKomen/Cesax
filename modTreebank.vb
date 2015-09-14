Imports System.Xml
Imports System.Text.RegularExpressions
Module modTreebank
  ' ===================================== LOCAL STRUCTURE ===========================================================
  Private Structure Corresp
    Dim CorId As Integer    ' The coreference ID
    Dim eTreeId As Integer  ' The ID of the <eTree> node
  End Structure
  ' ================================= GLOBAL VARIABLES =============================
  Public objError As New StringColl   ' Gathering of error messages
  ' ===================================== LOCAL VARIABLES ===========================================================
  Private strLastFile As String = ""      ' Last file read
  Private strLastText As String = ""      ' Text of last file read
  Private strLastLocation As String = ""  ' Last location ID
  Private intLevel As Integer = 0         ' The level of embedding
  Private objParsed As New StringColl     ' The output of the parsing
  Private intForId As Integer = 0         ' ID of <forest> element
  Private intTreId As Integer = 0         ' ID of <eTree> element
  Private arTransl() As Corresp           ' Array of correspondances
  Private loc_strSeg As String            ' The <seg> part
  Private loc_intPos As Integer           ' Local position within <seg>
  Private strTbkScheme As String = "Treebank.xsd"  ' Schema file for the treebank dataset
  Private loc_tdlTbk As DataSet = Nothing ' Local copy of treebank
  Private loc_intEtreeId As Integer       ' ID for <eTree> elements
  Private loc_intN As Integer             ' Word number in this sentence
  Private loc_strLetPrev As String        ' Text of preceding METADATA/LETTER text
  ' ===================================== LOCAL CONSTANTS ===========================================================
  Private strSpace As String = " " & vbTab & vbCrLf ' What should be regarded as white space
  Private strComSt As String = "<+ "                ' How a comment starts on one line
  Private strNoWord As String = ")" & strSpace      ' Characters that are not in a word
  Private strdelim As String = ""                   ' Line delimiter in text
  Private loc_strLineId As String = ""              ' Global line ID
  Private SPEC_CHAR_IN As String = "tadegTADEG"
  Private SPEC_CHAR_OUT As String = "þæđëġÞÆĐËĠ"
  Private SPEC_CHAR_ERR As String = " .,"
  Private SPEC_CHAR_EOUT As String = "þþþ"
  Private Const PUNCTUATION As String = ".,<>;:'[]{}-_=+\|)(*&^%$#@!~`" & """"
  '----------------------------------------------------------------------------------------
  ' Name:       ConvertOnePsdToPsdx()
  ' Goal:       Convert one file from source to destination
  ' History:
  ' 25-12-2009  ERK Created
  ' 17-03-2014  ERK New approach, using XML straight forward
  '----------------------------------------------------------------------------------------
  Public Function ConvertOnePsdToPsdx(ByVal strSrcFile As String, ByVal strDstFile As String, _
                             ByVal strLang As String, Optional ByVal bDoAsk As Boolean = False) As Boolean
    Dim tdlTbk As DataSet = Nothing ' The dataset that should contain the treebank structure
    Dim pdxConv As New XmlDocument  ' Document to hold the PSDX
    Dim strShort As String = ""     ' Short filename
    Dim strSent As String = ""      ' One sentence
    Dim strTextId As String = ""    ' ID for this text
    Dim strLineId As String         ' ID for this line
    Dim strEthno As String          ' Ethno code
    Dim rdThis As IO.StreamReader = Nothing  ' Reader for source file
    Dim ndxForGrp As XmlNode        ' Forest group
    Dim ndxFor As XmlNode           ' Forest node
    Dim ndxWork As XmlNode          ' Working node
    Dim intForestId As Integer = 1  ' Keep track of forest id
    Dim intPtc As Integer           ' Position in the input file
    Dim bEmpty As Boolean = False   ' Is sentence empty?

    Try
      ' Initialisation
      rdThis = Nothing
      ' Does the destination already exist
      If (IO.File.Exists(strDstFile)) Then
        ' Check the date of the destination compared to the source
        If (IO.File.GetLastWriteTime(strDstFile) > IO.File.GetLastWriteTime(strSrcFile)) Then
          ' Ask if we need to overwrite
          If (bDoAsk) Then
            Select Case MsgBox("Would you like to overwrite the existing psdx file?", MsgBoxStyle.YesNoCancel)
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
      ' Initialise the error object
      objError.Clear()
      ' Determine the ethno code
      strEthno = LangToEthno(strLang)
      ' Create XML document
      If (Not CreatePsdxHeader(pdxConv, strDstFile, strShort)) Then Return False
      ' Get forest group and do other initialisations
      ndxForGrp = pdxConv.SelectSingleNode("./descendant::forestGrp")
      strTextId = strShort
      loc_intEtreeId = 1
      loc_strLetPrev = ""
      ' Open source file for reading
      rdThis = New IO.StreamReader(strSrcFile)
      ' Loop through reading this file sentence by sentence
      Do
        ' If (intForestId = 420) Then Stop
        ' Initialise
        strLineId = "" : loc_strLineId = ""
        ' Read a sentence
        ReadOnePsdSentence(rdThis, strSent)
        ' Calculate where we are
        intPtc = (rdThis.BaseStream.Position / rdThis.BaseStream.Length) * 100
        Status("Converting " & intPtc & "%", intPtc)
        ' Be aware of interrupt
        If (bInterrupt) Then rdThis.Close() : Return False
        ' Is this something?
        ' (a) line must not be empty
        ' (b) line must contain at least one <node>
        If (strSent <> "") AndAlso (InStr(strSent, "<node") > 0) Then
          ' Create a new <forest> element and add its properties
          ndxFor = AddXmlChild(ndxForGrp, "forest", "forestId", intForestId, "attribute", _
                               "TextId", LCase(strTextId), "attribute", _
                               "File", strShort & ".psdx", "attribute", _
                               "Location", strTextId & "." & Format(intForestId, "0000"), "attribute")
          ' Add the <divs>: one for Spanish (org) and one for English (to be added)
          ndxWork = AddXmlChild(ndxFor, "div", "lang", "org", "attribute", "seg", "", "child")
          ndxWork = AddXmlChild(ndxFor, "div", "lang", "eng", "attribute", "seg", "", "child")
          ' Convert this sentence to Psdx
          If (Not OneTreebankToForest(ndxFor, strSent)) Then Logging("ConvertOnePsdToPsdx error") : rdThis.Close() : Return False
          ' Adapt the location attribute
          ndxFor.Attributes("Location").Value = MakeForestLoc(strShort, strLineId, intForestId)
          ' Check if a section should be added here
          If (intForestId = 1) OrElse (IsMetaLetterChange(ndxFor)) Then
            ' Add a section at this line
            AddAttribute(ndxFor, "Section", "")
          End If
          ' Calculate @from, @to
          eTreeSentence(ndxFor, ndxWork, False, True)
          '' Adapt the "org"/seg for eng_hist
          'ndxWork = ndxFor.SelectSingleNode("./child::div[@lang='org']/seg")
          'ndxWork.InnerText = VernToEnglish(Trim(ndxWork.InnerText))

          ' Make sure forestId is changed
          intForestId += 1
        End If
        ' Try to read on
      Loop Until (rdThis.EndOfStream)
      ' Close input
      rdThis.Close()
      ' Set the Language and the Ethno code
      ndxWork = pdxConv.SelectSingleNode("./descendant::langUsage/child::language")
      If (ndxWork IsNot Nothing) Then
        ndxWork.Attributes("ident").Value = strEthno
        ndxWork.Attributes("name").Value = strLang
      End If
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

      ' Write the result
      pdxConv.Save(strDstFile)
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modTreebank/ConvertOne error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Close input
      If (rdThis IsNot Nothing) Then rdThis.Close()
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       IsMetaLetterChange()
  ' Goal:       Convert one file from source to destination
  ' History:
  ' 18-03-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function IsMetaLetterChange(ByRef ndxForest As XmlNode) As Boolean
    'Dim ndxPrev As XmlNode    ' Previous <forest>
    Dim ndxLetThis As XmlNode ' Currnet letter
    'Dim ndxLetPrev As XmlNode ' Previous letter value
    Dim strLetThis As String  ' Current text
    'Dim strLetPrev As String  ' Previous text
    'Dim bChange As Boolean    ' There is change

    Try
      ' Validate
      If (ndxForest Is Nothing) OrElse (ndxForest.Name <> "forest") Then Return False
      ' Check for METADATA/LETTER in here
      ndxLetThis = ndxForest.SelectSingleNode("./descendant::eTree[@Label = 'METADATA']/child::eTree[@Label = 'LETTER']/child::eLeaf")
      If (ndxLetThis Is Nothing) Then Return False
      ' Get the text of this one
      strLetThis = ndxLetThis.Attributes("Text").Value
      ' Test
      If (loc_strLetPrev = "") Then
        loc_strLetPrev = strLetThis
        Return False
      End If
      ' Comparison
      If (loc_strLetPrev <> strLetThis) Then
        loc_strLetPrev = strLetThis
        Return True
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Give error
      HandleErr("modTreebank/IsMetaLetterChange error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       ReadOnePsdSentence()
  ' Goal:       Convert one file from source to destination
  ' History:
  ' 25-12-2009  ERK Created
  ' 17-03-2014  ERK New approach, using XML straight forward
  '----------------------------------------------------------------------------------------
  Private Function ReadOnePsdSentence(ByRef rdThis As IO.StreamReader, ByRef strBack As String) As Boolean
    Dim strLine As String         ' One line
    Dim colLine As New StringColl ' Collection of lines
    Dim intBracket As Integer = 0 ' Number of brackets
    Dim bDone As Boolean = False  ' flag

    Try
      ' Validate
      If (rdThis Is Nothing) Then Return False
      If (rdThis.EndOfStream) Then Return False
      ' Initialise
      strBack = ""
      ' Start reading
      While (Not bDone) AndAlso (Not rdThis.EndOfStream)
        ' Read a string
        strLine = rdThis.ReadLine
        ' Get the number of opening and closing brackets
        intBracket += Regex.Matches(strLine, "\(").Count - Regex.Matches(strLine, "\)").Count
        ' Check how fare we are
        If (intBracket <= 0) Then bDone = True
        ' Do preliminary conversion into xml
        strLine = PsdToXmlChunk(strLine)
        If (strLine <> "") Then
          ' Add to input 
          colLine.Add(strLine)
        End If
      End While
      ' Check the result
      If (colLine.Count = 0) Then Return True
      ' Combine with preamble and post
      strBack = "<sent>" & vbCrLf & colLine.Text & vbCrLf & "</sent>"
      ' Return succes
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modTreebank/ReadOnePsdSentence error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       PsdToXmlChunk()
  ' Goal:       Convert string into <node> and </node> elements containing xml string
  ' History:
  ' 18-03-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function PsdToXmlChunk(ByVal strIn As String) As String
    Dim mcLeft As MatchCollection ' Collection of matches
    Dim intI As Integer           ' Counter
    Dim strBack As String = ""
    Dim strLabelChar As String = "(\w|-|_|\+|\=|\$)"

    Try
      ' Do preliminary conversion into xml
      ' (1) Prepare for xml
      strBack = XmlEscape(strIn)
      ' (2) All left-brackets with a label
      '      Label characters: \w, hyphen, underscore, plus, equal, dollar
      '      (Other label characters should be added here!!!)
      mcLeft = Regex.Matches(strBack, "\([^\s\(\)]+(\s|$)")
      For intI = mcLeft.Count - 1 To 0 Step -1
        With mcLeft(intI)
          ' Look what quotes are called for
          If (InStr(.Value, "'") = 0) Then
            strBack = Left(strBack, .Index) & "<node pos='" & Trim(Mid(.Value, 2)) & "'>" & _
              Mid(strBack, .Index + .Length + 1)
          Else
            strBack = Left(strBack, .Index) & "<node pos=" & """" & Trim(Mid(.Value, 2)) & """" & ">" & _
              Mid(strBack, .Index + .Length + 1)
          End If
        End With
      Next intI
      ' (2) All other left-brackets
      mcLeft = Regex.Matches(strBack, "\(")
      For intI = mcLeft.Count - 1 To 0 Step -1
        With mcLeft(intI)
          strBack = Left(strBack, .Index) & "<node pos='ROOT'>" & _
            Mid(strBack, .Index + .Length + 1)
        End With
      Next intI
      ' (3) All right brackets
      mcLeft = Regex.Matches(strBack, "\)")
      For intI = mcLeft.Count - 1 To 0 Step -1
        With mcLeft(intI)
          strBack = Left(strBack, .Index) & "</node>" & _
            Mid(strBack, .Index + .Length + 1)
        End With
      Next intI
      ' Return the result
      Return strBack
    Catch ex As Exception
      ' Give error
      HandleErr("modTreebank/PsdToXmlChunk error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return ""
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       OneTreebankToForest()
  ' Goal:       Convert one sentence into a forest
  ' History:
  ' 17-03-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function OneTreebankToForest(ByRef ndxForest As XmlNode, ByRef strSent As String) As Boolean
    Dim pdxSent As New XmlDocument  ' Sentence
    Dim ndxThis As XmlNode          ' Working node

    Try
      ' Validate
      If (ndxForest Is Nothing) Then Return False
      ' Initialisation
      loc_intN = 1
      ' Convert sentence into xml code
      pdxSent.LoadXml(strSent)
      ' Walk all the elements
      ndxThis = pdxSent.FirstChild.FirstChild
      While (ndxThis IsNot Nothing)
        ' Check if it is a ROOT or other main node to be discarded
        If (DoLike(ndxThis.Attributes("pos").Value, "ROOT")) Then
          ' Go one step deeper
          ndxThis = ndxThis.FirstChild
        End If
        ' Process this node and its children recursively
        If (Not OneTreebankNode(ndxForest, ndxThis)) Then Return False
        ' Go to next node
        ndxThis = ndxThis.NextSibling
      End While
      ' Return succes
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modTreebank/OneTreebankToForest error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       OneTreebankNode()
  ' Goal:       Process node [ndxNode] into a child under [ndxParent]
  ' History:
  ' 17-03-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function OneTreebankNode(ByRef ndxParent As XmlNode, ByRef ndxNode As XmlNode) As Boolean
    Dim ndxWork As XmlNode      ' Working node
    Dim ndxUnkn As XmlNode      ' Unknown node
    Dim ndxLeaf As XmlNode      ' <eLeaf> node
    Dim ndxChild As XmlNode     ' One child node
    Dim ndxForest As XmlNode    ' The forest
    Dim strText As String = ""  ' Vernacular text
    Dim strType As String = ""  ' Type of node
    Dim strValue As String = "" ' Value of <eLeaf>
    Dim strComment As String    ' Comment

    Try
      ' Validate
      If (ndxParent Is Nothing) OrElse (ndxNode Is Nothing) Then Return False
      ' Check for "ID" nodes, which need separate treatment
      If (ndxNode.Attributes("pos").Value = "ID") Then
        ' Retrieve the ID value
        loc_strLineId = ndxNode.InnerText
        ' This node needs no further processing
        Return True
      End If
      ' Check if this is an end-node
      If (ndxNode.SelectNodes("./child::node").Count = 0) Then
        ' Default values
        strType = "Vern" : strValue = Trim(ndxNode.InnerText)
        ' Determine the value of this 'word'
        Select Case strValue
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
        End Select
        ' We are CREATING, not synchronizing: add one XML child under [ndxFor]
        If (strType = "Punct") AndAlso (strValue <> "") Then
          ndxWork = AddEtreeChild(ndxParent, loc_intEtreeId, strValue, 0, 0)
        Else
          ndxWork = AddEtreeChild(ndxParent, loc_intEtreeId, ndxNode.Attributes("pos").Value, 0, 0)
          ' Check what kind of <eLeaf> this is: Vern, Star or Zero
          strType = GetLeafType(strValue)
        End If
        loc_intEtreeId += 1
        ' Add <eLeaf> to the <eTree> node
        ndxLeaf = AddEleafChild(ndxWork, strType, strValue, 0, 0)
        ndxLeaf.Attributes("n").Value = loc_intN
        loc_intN += 1
      Else
        ' Not an end node
        ndxWork = AddEtreeChild(ndxParent, loc_intEtreeId, ndxNode.Attributes("pos").Value, 0, 0)
        loc_intEtreeId += 1
        ' Access the children
        ndxChild = ndxNode.FirstChild
        While (ndxChild IsNot Nothing)
          ' Get next child
          If (ndxChild.Name = "node") Then
            ' Process this child
            If (Not OneTreebankNode(ndxWork, ndxChild)) Then Return False
          Else
            'Stop
            If (ndxChild.Name = "#text") Then
              Dim strLeaf As String = ""    ' Text of leaf

              ' Add this as an UNKNOWN node
              ndxUnkn = AddEtreeChild(ndxWork, loc_intEtreeId, "UNKNOWN", 0, 0)
              loc_intEtreeId += 1
              ' Calculate text of leaf
              strLeaf = ndxChild.InnerText
              strLeaf = Regex.Replace(strLeaf, "\s+", "")
              ' Add the content as an <eLeaf> under it
              ndxLeaf = AddEleafChild(ndxUnkn, "Vern", strLeaf, 0, 0)
              ndxLeaf.Attributes("n").Value = loc_intN
              loc_intN += 1
              ' Show what we've done
              ndxForest = ndxParent.SelectSingleNode("./ancestor-or-self::forest")
              strComment = ndxForest.Attributes("File").Value & ": " & _
                      "Added [UNKNOWN " & strLeaf & "] node at @forestId=" & _
                      ndxForest.Attributes("forestId").Value
              ' Add this in the revision information
              AddRevDesc(pdxCurrentFile, "Cesax", Format(Now, "g"), strComment)
              Logging(strComment)
            Else
              Stop
            End If
          End If
          ' Go to next child
          ndxChild = ndxChild.NextSibling
        End While
      End If
      ' Return succes
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modTreebank/OneTreebankNode error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       ConvertOnePsdToPsdx_Org()
  ' Goal:       Convert one file from source to destination
  ' History:
  ' 25-12-2009  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function ConvertOnePsdToPsdx_Org(ByVal strSrcFile As String, ByVal strDstFile As String, _
                             Optional ByVal bDoAsk As Boolean = False) As Boolean
    Dim tdlTbk As DataSet = Nothing ' The dataset that should contain the treebank structure

    Try
      ' Does the destination already exist
      If (IO.File.Exists(strDstFile)) Then
        ' Check the date of the destination compared to the source
        If (IO.File.GetLastWriteTime(strDstFile) > IO.File.GetLastWriteTime(strSrcFile)) Then
          ' Ask if we need to overwrite
          If (bDoAsk) Then
            Select Case MsgBox("Would you like to overwrite the existing psdx file?", MsgBoxStyle.YesNoCancel)
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
      ' Create a new dataset
      If (CreateDataSet(strTbkScheme, tdlTbk)) Then
        ' Set the local copy
        loc_tdlTbk = tdlTbk
        ' Initialise the error object
        objError.Clear()
        ' Convert text into datastructure
        If (PsdToDataset(strSrcFile, tdlTbk, False, strDstFile)) Then
          ' Save datastructure as file
          tdlTbk.WriteXml(strDstFile)
          ' Were there any error messages?
          If (objError.Count > 0) Then
            ' Save the error messages
            IO.File.WriteAllText(Left(strDstFile, _
                InStrRev(strDstFile, ".") - 1) & "-Error.txt", objError.Text)
            Logging("------ there were errors in " & strDstFile & " ----")
          End If
          ' Return success
          Return True
        End If
      Else
        ' Unable to create the correct scheme!
        Status("modTreebank/ConvertOnePsdToPsdx_Org: Was not able to create a file from scheme " & strTbkScheme)
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Give error
      HandleErr("modMain/ConvertOnePsdToPsdx_Org error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  PsdToDataset
  ' Goal :  Convert PSD text to dataset
  ' History:
  ' 25-12-2009  ERK Created
  ' 15-03-2011  ERK Adapted slightly for Cesax
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function PsdToDataset(ByRef strSrcFile As String, ByRef tdlThis As DataSet, ByVal bIsWSJ As Boolean, _
                                Optional ByVal strUseFile As String = "") As Boolean
    Dim strText As String = ""    ' Return string
    Dim strIn As String = ""      ' Text of the source file
    Dim intPos As Integer = 1     ' Position to start at

    Try
      ' Initialise 
      If (Not InitTbk(tdlThis)) Then
        ' Return failure
        Return False
      End If
      ' Read the source file
      strIn = IO.File.ReadAllText(strSrcFile)
      ' Do we need WSJ conversion?
      If (bIsWSJ) Then
        If (Not DoAddWSJid(IO.Path.GetFileNameWithoutExtension(strSrcFile), strIn)) Then
          ' Return failure
          Return False
        End If
      End If
      ' Check if we don't have a file with (FS-  structures
      If (InStr(strIn, "(FS-") > 0) Then
        ' We cannot process this one
        MsgBox("This file has been produced by Cesax File/Export, and cannot be imported into Cesax again")
        Return False
      End If
      ' Determine the delimiter
      strdelim = GetDelim(strIn, vbCrLf, vbCr, vbLf)
      ' Show that we are busy -- HOW???
      ' Try to parse the PSD text
      If (strUseFile = "") Then
        If (bnf_Text(tdlThis, strSrcFile, strIn, intPos, strText)) Then Return True
      Else
        If (bnf_Text(tdlThis, strUseFile, strIn, intPos, strText)) Then Return True
      End If
      ' Parse failure?
      MsgBox("Could not parse PSD text:" & vbCrLf & strIn)
      Return False
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/PsdToDataset error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  InitTbk
  ' Goal :  Initialize parsing of treebank to a dataset
  ' History:
  ' 26-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function InitTbk(ByRef tdlThis As DataSet) As Boolean
    Dim tblThis As DataTable  ' One table
    Dim intI As Integer       ' Counter

    Try
      '' Set the local copy of the dataset
      'tdlTbk = tdlThis
      ' Do we have something?
      If (Not tdlThis Is Nothing) Then
        ' Clear all rows for all tables
        For Each tblThis In tdlThis.Tables
          ' Access this table
          With tblThis
            ' Walk the rows from last to first
            For intI = .Rows.Count - 1 To 0 Step -1
              ' Delete this row
              .Rows(intI).Delete()
            Next intI
          End With
        Next tblThis
        ' Reset
        tdlThis.AcceptChanges()
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/InitTbk error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  GetPsdText
  ' Goal :  Get the PSD text from [intPrecNum] lines preceding [strTextId] out of the PSD file [strInFile]
  ' Note :  The return INCLUDES the line with the given ID itself
  ' History:
  ' 14-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function GetPsdText(ByVal strInFile As String, ByVal strTextId As String, ByVal intLines As Integer) As String
    Dim intPos As Integer   ' Position in file
    Dim intPosSt As Integer ' Start position in file
    Dim bFound As Boolean   ' Whether the text has been found

    Try
      ' Was this the last file read?
      If (strLastFile = strInFile) Then
        ' Re-use the text that has already been read
        bFound = True
      ElseIf (IO.File.Exists(strInFile)) Then
        ' Read the text
        strLastText = IO.File.ReadAllText(strInFile)
        ' Set the name of the last file read
        strLastFile = strInFile
        ' Signal that we have found it
        bFound = True
      Else
        ' not found the text
        bFound = False
      End If
      ' Check if file exists
      If (bFound) Then
        ' Get the text of the file
        ' Find the position in the file
        intPos = InStr(strLastText, strTextId & ")", CompareMethod.Text)
        ' Found something?
        If (intPos > 0) Then
          ' Advance two ) lines
          intPos = InStr(intPos + 1, strLastText, ")")
          If (intPos > 0) Then intPos = InStr(intPos + 1, strLastText, ")")
          If (intPos > 0) Then
            ' Add 1 position, in order to skip the ) sign
            intPos += 1
            ' We should add 1 to include ourselves...
            intLines += 1
            ' Count backwards the (ID ... signs
            intPosSt = InStrRev(strLastText, "(ID ", intPos)
            While (intLines > 0) AndAlso (intPosSt > 0)
              ' Decrement lines
              intLines -= 1
              ' Try get the next one
              intPosSt = InStrRev(strLastText, "(ID ", intPosSt - 1)
            End While
            ' Did we reach the start of the file?
            If (intPos = 0) Then
              ' Pass back from start of file
              Return Left(strLastText, intPos)
            Else
              ' Skip until the next two ) signs
              intPosSt = InStr(intPosSt + 1, strLastText, ")")
              If (intPosSt > 0) Then intPosSt = InStr(intPosSt + 1, strLastText, ")")
              ' Are we okay?
              If (intPosSt > 0) Then
                ' Advance 1 more place
                intPosSt += 1
                ' Pass back from here
                Return Mid(strLastText, intPosSt, intPos - intPosSt)
              End If
            End If
          End If
        End If
      End If
      '' Warn the user
      'HandleErr("Could not find preceding lines in file " & strInFile & vbCrLf & _
      '       "This concerns the context of [" & strTextId & "]")
      ' Pass back nothing
      GetPsdText = ""
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/GetPsdText error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  bnf_Text
  ' Goal :  Parse the input PSD string starting from position [intPos]
  ' Note :  Here is the BNF for the PSD input:
  '   <Text>        ::= <TxtLine>*
  '   <TxtLine>     ::= <SrcLine> | <ComLine>
  '   <ComLine>     ::= CHAR*
  '   <SrcLine>     ::= "(" sp* <Node>$ sp* ")"
  '   <Node>        ::= "(" sp* <Label> " " sp* <NodeContent> sp* ")"
  '   <NodeContent> ::= <NodeLex> | <Node>$
  '   <NodeLex>     ::= <Lexeme> <Node>*
  '   <Label>       ::= <word>
  '   <Lexeme>      ::= <word>
  '   <word>        ::= CHAR*
  ' History:
  ' 14-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function bnf_Text(ByRef tdlThis As DataSet, ByRef strSrcFile As String, ByVal strIn As String, _
                            ByRef intPos As Integer, ByRef strBack As String) As Boolean
    Dim strText As String = ""          ' Start with empty return
    Dim intBack As Integer = intPos     ' Roll back position
    Dim intPtc As Integer               ' Percentage
    Dim intId As Integer                ' Dummy ID, since a forestgrp doesn't have one
    Dim intI As Integer                 ' Counter
    Dim dtrForGrp As DataRow = Nothing  ' forestGrp Datarow
    Dim dtrCoref() As DataRow           ' Set of coreference nodes

    Try
      ' Initialise the output string collection
      objParsed.Clear() : intForId = 1 : intTreId = 1
      ' We should also at least have a <forestGrp> node
      ' Create a new <forestGrp> element
      If (Not CreateNewRow(tdlThis, "forestGrp", "", intId, dtrForGrp)) Then
        ' There was an error
        Return False
      End If
      ' Set the File parameter of the <forestGrp>
      dtrForGrp.Item("File") = IO.Path.GetFileNameWithoutExtension(strSrcFile)
      ' There should at least be ONE <Node>
      If (bnf_SrcLine(tdlThis, strSrcFile, strIn, intPos, strText)) Then
        ' Try get as many <Node> as possible
        Do
          ' Show progress
          intPtc = (intPos * 100) \ strIn.Length
          Status("Parsing " & strSrcFile & " " & intPtc & "%", intPtc)
          ' Add to the [strBack]
          strBack &= strText
          ' Reset [strNodeLex]
          strText = ""
          ' Try get another <Node>
        Loop While (bnf_SrcLine(tdlThis, strSrcFile, strIn, intPos, strText))
      End If
      ' Now visit all COREF nodes: they have RefType set to something
      dtrCoref = tdlThis.Tables("eTree").Select("RefType <>''")
      For intI = 0 To dtrCoref.Count - 1
        ' Show where we are
        intPtc = (intI * 100) \ dtrCoref.Count
        Status("Coref info " & strSrcFile & " " & intPtc & "%", intPtc)
        ' Add the coreference information to this node
        AddCorefXml(tdlThis, dtrCoref(intI))
      Next intI
      ' Next visit all COREF nodes: they have MyId set to a value larger than zero
      dtrCoref = tdlThis.Tables("eTree").Select("MyId >= 0")
      For intI = 0 To dtrCoref.Count - 1
        ' Show where we are
        intPtc = (intI * 100) \ dtrCoref.Count
        Status("Coref clean " & strSrcFile & " " & intPtc & "%", intPtc)
        ' Clear garbage information from this node
        CleanCorefXml(tdlThis, dtrCoref(intI))
      Next intI
      ' Return success
      bnf_Text = True
    Catch ex As Exception
      ' Note error
      HandleErr("modPars/bnf_Text error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  AddCorefXml
  ' Goal :  Take the coreference information from the parent's attributes and store it in
  '           the appropriate TEI-P5 elements: 
  '           <ref target=""/>
  '           <fs type="coref">
  '             <f name="RefType" value="Identity"/>
  '             <f name="NdDist" value="2"/>
  '           </fs>
  ' History:
  ' 01-03-2010  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function AddCorefXml(ByRef tdlThis As DataSet, ByRef dtrParent As DataRow) As Boolean
    Dim dtrFound() As DataRow         ' Result of SELECT
    Dim dtrRef As DataRow = Nothing   ' a <ref> datarow
    Dim dtrFs As DataRow = Nothing    ' a <fs> datarow
    Dim dtrF As DataRow = Nothing     ' a <f> datarow
    Dim intDummyId As Integer = -1    ' Dummy ID for the above
    Dim intRefId As Integer = -1      ' The ID we are referring to

    Try
      ' What is the reference ID?
      intRefId = dtrParent.Item("RefId")
      ' Find the reference information
      dtrFound = tdlThis.Tables("eTree").Select("MyId=" & intRefId)
      ' Found something?
      If (dtrFound.Count = 0) Then Return False
      ' Create the <ref> datarow
      If (Not CreateNewRow(tdlThis, "ref", "", intDummyId, dtrRef)) Then Return False
      ' Add the reference information
      dtrRef.SetParentRow(dtrParent)
      dtrRef.Item("target") = dtrFound(0).Item("Id")
      ' Create the <fs> datarow
      If (Not CreateNewRow(tdlThis, "fs", "", intDummyId, dtrFs)) Then Return False
      ' Tie up this <fs> element
      dtrFs.SetParentRow(dtrParent)
      dtrFs.Item("type") = "coref"
      ' Make the <f> element for RefType
      If (Not CreateNewRow(tdlThis, "f", "", intDummyId, dtrF)) Then Return False
      ' Tie up this <f> element
      dtrF.SetParentRow(dtrFs)
      dtrF.Item("name") = "RefType"
      dtrF.Item("value") = dtrParent.Item("RefType")
      ' Make the <f> element for NdDist
      If (Not CreateNewRow(tdlThis, "f", "", intDummyId, dtrF)) Then Return False
      ' Tie up this <f> element
      dtrF.SetParentRow(dtrFs)
      dtrF.Item("name") = "NdDist"
      dtrF.Item("value") = dtrParent.Item("NdDist")
      ' Return success
      Return True
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/AddCorefXml error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  CleanCorefXml
  ' Goal :  Clean the MyId stuff etc
  ' History:
  ' 01-03-2010  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function CleanCorefXml(ByRef tdlThis As DataSet, ByRef dtrParent As DataRow) As Boolean
    Try
      ' Remove the appropriate parent's attributes
      dtrParent.SetField("MyId", DBNull.Value)
      dtrParent.SetField("RefId", DBNull.Value)
      dtrParent.SetField("RefType", DBNull.Value)
      dtrParent.SetField("NdDist", DBNull.Value)
      ' Return success
      Return True
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/CleanCorefXml error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  bnf_SrcLine
  ' Goal :  Parse the input PSD string starting from position [intPos]
  ' Note :  Here is the BNF for the PSD input:
  '   <SrcLine>     ::= sp* <Open> sp* <Node>$ sp* ")"
  '   <Open>        ::= "(" | "(ROOT"
  ' History:
  ' 14-12-2009  ERK Created
  ' 27-08-2013  ERK Added recognition of (NODE, which exists because of Dutch input
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function bnf_SrcLine(ByRef tdlThis As DataSet, ByVal strFile As String, ByRef strIn As String, _
                               ByRef intPos As Integer, ByRef strBack As String) As Boolean
    Dim strSrcLine As String = ""       ' Start with empty return
    Dim intBack As Integer = intPos     ' Roll back position
    Dim strType As String = ""          ' The type of node being returned: id or other
    Dim strLineId As String = ""        ' Possible ID value
    Dim strTextId As String             ' The correct textid value
    Dim strLoc As String = ""           ' Location
    Dim strSeg As String = ""           ' The text of this node
    Dim bCreated As Boolean = False     ' Whether a new <forest> element was created
    Dim dtrForest As DataRow = Nothing  ' Created new <forest> element
    Dim dtrDiv As DataRow = Nothing     ' Created new <div> element
    Dim intForestId As Integer = -1     ' ID of the forest
    Dim intDummyId As Integer = -1      ' A dummy ID for the <div> elements
    Dim intI As Integer = 0             ' Position within string

    Try
      ' Skip optional spaces
      bnf_Space(strIn, intPos)
      ' Treat the first line
      If (Mid(strIn, intPos, 1) <> "(") Then
        ' This is a comment line, not one that belongs to the bracketed labelling -- skip till end of line
        bnf_SkipLine(strIn, intPos)
      End If
      ' Try get left bracket
      If (bnf_Expect(strIn, intPos, "(")) Then
        ' Check if this is (ROOT
        If (Mid(strIn, intPos, 4) = "ROOT") OrElse (Mid(strIn, intPos, 4) = "NODE") Then
          ' Some PSD files start out with (ROOT or (NODE in the outer layer
          intPos += 4
        End If
        ' Debug.Print(AscW(Mid(strIn, intPos + 1, 1)))
        ' Skip optional spaces
        bnf_Space(strIn, intPos)
        ' Create a new <forest> element
        If (Not CreateNewRow(tdlThis, "forest", "forestId", intForestId, dtrForest)) Then
          ' There was an error
          Return False
        End If
        ' Set the parent of this <forest> one
        dtrForest.SetParentRow(tdlThis.Tables("forestGrp").Rows(0))
        ' Indicate a new element was created
        bCreated = True
        ' =========== DEBUG ================
        ' If (intForestId = 4783) Then Stop
        ' ==================================
        ' There should at least be ONE <Node>
        If (bnf_Node(tdlThis, strIn, intPos, strSrcLine, strType, dtrForest)) Then
          ' Try get as many <Node> as possible
          Do
            ' What is the type being returned?
            Select Case strType
              Case "id"
                ' Save the ID for me
                strLineId = strSrcLine
              Case "code"
                ' Don't add anything?? --> TODO: CHECK THIS OUT!!!
              Case Else
                ' Add to the [strBack]
                strBack &= strSrcLine
            End Select
            ' Reset [strSrcLine]
            strSrcLine = ""
            ' Try get another <Node>
          Loop While (bnf_Node(tdlThis, strIn, intPos, strSrcLine, strType, dtrForest))
          ' Skip optional spaces
          bnf_Space(strIn, intPos)
          ' Also we should find a closing bracket
          If (bnf_Expect(strIn, intPos, ")")) Then
            ' Make sure we have the correct textid value
            strTextId = GetShortFileName(strFile)
            ' Get the @Location value
            strLoc = MakeForestLoc(strTextId, strLineId, intForestId)
            '' Did we get a LineId element?
            'If (strLineId = "") Then strLineId = loc_strLineId
            'If (strLineId = "") Then
            '  ' There is no LineId, so we have to make up one ourselves!
            '  strLoc = "s" & Format(intForestId, "0000")
            '  strLineId = IO.Path.GetFileNameWithoutExtension(strFile)
            'Else
            '  ' Check on comma usage
            '  If (Split(strLineId, ",").Count > 1) Then
            '    ' Is there a period?
            '    If (InStr(strLineId, ".") = 0) Then
            '      ' No period, so comma is used as division
            '      intI = InStr(strLineId, ",")
            '      strLoc = Mid(strLineId, intI + 1)
            '      strLineId = Left(strLineId, intI - 1)
            '    Else
            '      ' The comma is used for something else
            '      strLoc = strLineId.Replace(",", ".") & "." & intForestId
            '      strLineId = IO.Path.GetFileNameWithoutExtension(strFile)
            '    End If
            '  Else
            '    ' Split the LineId
            '    intI = InStr(strLineId, ",")
            '    If (intI > 0) Then
            '      strLoc = Mid(strLineId, intI + 1)
            '      strLineId = Left(strLineId, intI - 1)
            '    ElseIf (InStr(strLineId, ".") > 0) Then
            '      ' Use period
            '      intI = InStr(strLineId, ".")
            '      strLoc = Mid(strLineId, intI + 1)
            '      strLineId = Left(strLineId, intI - 1)
            '    ElseIf (InStr(strLineId, "_") > 0) Then
            '      ' Use underscore
            '      intI = InStr(strLineId, "_")
            '      strLoc = Mid(strLineId, intI + 1)
            '      strLineId = Left(strLineId, intI - 1)
            '    Else
            '      strLoc = strLineId
            '    End If
            '  End If
            'End If
            ' Create the next <div> element for the "Translation" into present day English
            If (Not CreateNewRow(tdlThis, "div", "", intDummyId, dtrDiv)) Then
              ' There was an error
              Return False
            End If
            ' Set the values for this <div> element
            dtrDiv.Item("lang") = "eng"
            dtrDiv.Item("seg") = ""
            dtrDiv.SetParentRow(dtrForest)
            ' Create one new <div> element for the "Original" OE/ME
            If (Not CreateNewRow(tdlThis, "div", "", intDummyId, dtrDiv)) Then
              ' There was an error
              Return False
            End If
            ' Set the values for this <div> element
            dtrDiv.Item("lang") = "org"
            dtrDiv.Item("seg") = VernToEnglish(Trim(strBack))
            dtrDiv.SetParentRow(dtrForest)
            ' Add this to the [dtrForest]
            dtrForest.Item("Location") = strLoc
            dtrForest.Item("TextId") = LCase(strTextId) ' LCase(strLineId)
            dtrForest.Item("File") = IO.Path.GetFileNameWithoutExtension(strFile)
            ' Now everything is calculated -- proceed by calculating all the From/To positions
            DoSegPositions(dtrForest)
            ' Also store the location in the global variable
            strLastLocation = strLineId
            ' Append a CRLF
            strBack &= vbCrLf
            ' Return success
            Return True
          End If
        End If
      End If
      ' Was a new element created?
      If (bCreated) Then
        ' Show the error
        ShowError(strIn, intPos)
        ' Remove this element
        dtrForest.Delete()
        ' Process the deletion
        tdlThis.AcceptChanges()
      End If
      ' Wind back and return failure
      intPos = intBack
      bnf_SrcLine = False
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/bnf_SrcLine error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  MakeForestLoc
  ' Goal :  Calculate the @Location element for this forest
  ' History:
  ' 18-03-2014  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function MakeForestLoc(ByVal strShort As String, ByVal strLineId As String, ByVal intForestId As Integer) As String
    Dim intPos As Integer ' Position in string
    Dim strLoc As String  ' Location we return

    Try
      ' Did we get a LineId element?
      If (strLineId = "") Then strLineId = loc_strLineId
      If (strLineId = "") Then
        ' There is no LineId, so we have to make up one ourselves!
        strLoc = "s" & Format(intForestId, "0000")
        strLineId = strShort
      Else
        ' Check on comma usage
        If (Split(strLineId, ",").Count > 1) Then
          ' Is there a period?
          If (InStr(strLineId, ".") = 0) Then
            ' No period, so comma is used as division
            intPos = InStr(strLineId, ",")
            strLoc = Mid(strLineId, intPos + 1)
            strLineId = Left(strLineId, intPos - 1)
          Else
            ' The comma is used for something else
            strLoc = strLineId.Replace(",", ".") & "." & intForestId
            strLineId = strShort
          End If
        Else
          ' Split the LineId
          intPos = InStr(strLineId, ",")
          If (intPos > 0) Then
            strLoc = Mid(strLineId, intPos + 1)
            strLineId = Left(strLineId, intPos - 1)
          ElseIf (InStr(strLineId, ".") > 0) Then
            ' Use period
            intPos = InStr(strLineId, ".")
            strLoc = Mid(strLineId, intPos + 1)
            strLineId = Left(strLineId, intPos - 1)
          ElseIf (InStr(strLineId, "_") > 0) Then
            ' Use underscore
            intPos = InStr(strLineId, "_")
            strLoc = Mid(strLineId, intPos + 1)
            strLineId = Left(strLineId, intPos - 1)
          Else
            strLoc = strLineId
          End If
        End If
      End If
      ' Return the location
      Return strLoc
    Catch ex As Exception
      ' Note error
      HandleErr("modTreebank/MakeForestLoc error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return ""
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  DoSegPositions
  ' Goal :  Determine the positions of all the words and the punctuation within the <seg> element
  ' History:
  ' 27-04-2010  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function DoSegPositions(ByRef dtrForest As DataRow) As Boolean
    Dim strRel1 As String = "forest_eTree"  ' Branch from <forest> to <eTree>
    Dim strDiv As String = "forest_div"     ' Branch from <forest> to <div>
    Dim dtrTree As DataRow                  ' Each of the forest's children
    Dim dtrDiv() As DataRow                 ' The two <div> sections of <forest>
    Dim intFrom As Integer = 0              ' Starting point for @From attribute
    Dim intTo As Integer = 0                ' Starting point for @To attribute

    Try
      ' Get the <div> sections
      dtrDiv = dtrForest.GetChildRows(strDiv)
      ' Check whether the number of children is correct
      If (dtrDiv.Length <> 2) Then
        ' Incorrect number of <div> children
        MsgBox("modParse/DoSegPositions: incorrect number of <div> children = " & dtrDiv.Length)
        Return False
      End If
      ' Determine the <seg> part
      loc_strSeg = dtrDiv(1).Item("seg") : loc_intPos = 1
      ' Don't know how to walk a tree within a datatable...
      For Each dtrTree In dtrForest.GetChildRows(strRel1)
        ' Walk this <eTree> recursively
        If (Not DoOneEtree(dtrTree, intFrom, intTo, False)) Then
          ' Something went wrong
          Return False
        End If
      Next dtrTree
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/DoSegPositions error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  DoOneEtree
  ' Goal :  Recursively walk through a tree
  ' History:
  ' 27-04-2010  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function DoOneEtree(ByRef dtrTree As DataRow, ByRef intFrom As Integer, ByRef intTo As Integer, _
                              ByVal bIsCode As Boolean) As Boolean
    Dim strEtree As String = "eTree_eTree"  ' Branch from <eTree> to <eTree>
    Dim strEleaf As String = "eTree_eLeaf"  ' Branch from <eTree> to <eLeaf>
    Dim strText As String = ""              ' The text of the <eLeaf>
    Dim strType As String = ""              ' The type of the <eLeaf>
    Dim dtrChild As DataRow                 ' Each of the <eTree> children
    Dim bLexical As Boolean = False         ' Whether the node being processed is lexical or not
    Dim bTreeCode As Boolean = False        ' Whether the current <dtrTree> is a CODE node
    Dim intMyFrom As Integer = 0            ' Local copy of @From attribute
    Dim intMyTo As Integer = 0              ' Local copy of @To attribute
    'Dim intFrom As Integer = 0              ' Value of the @From attribute
    'Dim intTo As Integer = 0                ' Value of the @To attribute

    Try
      ' Is this <eTree> a code node?
      If (dtrTree.Item("Label") = "CODE") Then bTreeCode = True
      ' Determine whether this tree has <eTree> children or <eLeaf> (it's either or!!!)
      If (dtrTree.GetChildRows(strEtree).Length = 0) Then
        ' There are no <eTree> children; there should be an <eLeaf>
        If (dtrTree.GetChildRows(strEleaf).Length = 0) Then
          ' Something is wrong!!!
          MsgBox("There is an <eTree> without a proper child in: " & loc_strSeg)
          Return False
        Else
          ' Okay there is an <eLeaf> child - process it...
          dtrChild = dtrTree.GetChildRows(strEleaf)(0)
          strText = VernToEnglish(dtrChild.Item("Text"))
          strType = dtrChild.Item("Type")
          ' Determine the kind of node
          Select Case dtrTree.Item("Label")
            Case "E_S", "CODE", "ID"
              bLexical = False
            Case Else
              Select Case strType
                Case "Punct", "Vern"
                  ' Regard punctuation as lexical...
                  bLexical = True
                Case "Star", "Zero"
                  ' Traces, *con* and zero's are not normally regarded as lexical (but see further down!!)
                  bLexical = False
                Case Else
                  ' Warn user
                  MsgBox("modParse/DoOneEtree unknown <eLeaf> type: " & strType)
                  bLexical = False
              End Select
          End Select
          ' Only fully lexical words or punctuation is processed with positions
          If (strText <> "") And (bLexical) And (Not bIsCode) Then
            ' Find the text in <seg>
            loc_intPos = InStr(loc_intPos, loc_strSeg, strText, CompareMethod.Binary)
            ' Found it?
            If (loc_intPos = 0) Then
              ' Something went wrong...
              HandleErr("modParse/DoOneTree fatal problem: Could not find [" & strText & "] in:" & vbCrLf & "[" & loc_strSeg & "]")
              ' Return False
              loc_tdlTbk.WriteXml("d:/temp.psdx")
              Return False
            End If
            ' Kol b-seder...
            dtrChild.Item("from") = loc_intPos
            loc_intPos += strText.Length
            dtrChild.Item("to") = loc_intPos - 1
            '' Note the end position for the parent
            'dtrTree.Item("to") = loc_intPos - 1
          ElseIf (strType = "Star") Then
            ' The position of traces and empty subjects elided under conjunction (*con*) is noted, but with zero length
            dtrChild.Item("from") = loc_intPos
            dtrChild.Item("to") = loc_intPos
            '' Note the end position for the parent
            'dtrTree.Item("to") = loc_intPos
          Else
            ' Pay lip service
            dtrChild.Item("from") = 0
            dtrChild.Item("to") = 0
          End If
        End If
        ' Set local copies of @From and @To
        intMyFrom = dtrChild("from")
        intMyTo = dtrChild("to")
        ' I am an <eTree> with <eLeaf> child, so I get the same @From and @To as my child does
        ' Add the values to the <eTree> element
        With dtrTree
          .Item("from") = intMyFrom
          .Item("to") = intMyTo
        End With
        ' Adapt the global copies of @From and @To
        If (intMyFrom > 0) Then
          ' Check if [intFrom] has already been set
          If (intFrom = 0) Then
            ' Set [intFrom] for the first time
            intFrom = intMyFrom
          ElseIf (intMyFrom < intFrom) Then
            ' Adapt [intFrom] to a lower value (should never be necessary actually!)
            intFrom = intMyFrom
          End If
        End If
        ' Adapt the global copy of @To
        If (intMyTo > 0) Then
          ' Check if [intTo] was already set
          If (intTo = 0) Then
            ' Initially take this local copy as starting point
            intTo = intMyTo
          ElseIf (intMyTo > intTo) Then
            ' Adapt [intTo] to a larger size
            intTo = intMyTo
          End If
        End If
      Else
        '' Note the value of the intPos now
        'intFrom = loc_intPos
        ' There are <eTree> children; Process each separately
        For Each dtrChild In dtrTree.GetChildRows(strEtree)
          ' Start with new @From and @To attributes
          intMyFrom = 0 : intMyTo = 0
          ' Process this child
          If (Not DoOneEtree(dtrChild, intMyFrom, intMyTo, bTreeCode)) Then
            ' Something went wrong...
            Return False
          End If
          ' Adapt the @From and @To parameters
          If (intFrom = 0) Then
            ' set it to the current outcome
            intFrom = intMyFrom
          ElseIf (intMyFrom > 0) AndAlso (intMyFrom < intFrom) Then
            ' Set the from parameter
            intFrom = intMyFrom
          End If
          ' Adapt the @To parameter
          If (intTo = 0) Then
            ' Set it the the current outcome
            intTo = intMyTo
          ElseIf (intMyTo > 0) AndAlso (intMyTo > intTo) Then
            ' Adapt the @To parameter
            intTo = intMyTo
          End If
        Next dtrChild
        ' Add the values to the <eTree> element
        With dtrTree
          .Item("from") = intFrom
          .Item("to") = intTo
        End With
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/DoOneTree error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  bnf_Node
  ' Goal :  Parse the input PSD string starting from position [intPos]
  '         Return the type of the node in [strType]
  ' Note :  Here is the BNF for the PSD input:
  '   <Node>        ::= sp* "(" sp* <Label> " " sp* <NodeContent> sp*        ")"
  '   <Node>        ::= sp* "(" sp* <Label> " " sp* <NodeContent> sp* <CODE> ")"
  ' History:
  ' 14-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function bnf_Node(ByRef tdlThis As DataSet, ByRef strIn As String, ByRef intPos As Integer, ByRef strBack As String, _
                            ByRef strType As String, ByRef dtrParent As DataRow) As Boolean
    Dim strNode As String = ""        ' Start with empty return
    Dim strLabel As String = ""       ' Text of the label
    Dim strLexeme As String = ""      ' Room for a (faulty) lexeme
    Dim intNodeId As Integer = -1     ' ID for this <eTree> element
    Dim dtrNode As DataRow = Nothing  ' The datarow we are adding
    Dim bIsLeaf As Boolean = False    ' Parameter for [NodeContent]
    Dim bMade As Boolean = False      ' Indicates that a new <eTree> was created
    Dim bIsIDnode As Boolean = False  ' Whether this node has the ID label
    Dim bIsValid As Boolean = False   ' Whether we find a valid node: ( + <Label> + space
    Dim bIsCode As Boolean = False    ' Whether this node has the CODE label
    Dim intBack As Integer = intPos   ' Roll back position
    Dim intMyId As Integer = -1       ' Coref: My own Id
    Dim intRefId As Integer = -1      ' Coref: ID I point to
    Dim intNdDist As Integer = 0      ' Coref: nodal distance
    Dim strRefType As String = ""     ' Coref: type of coreference link

    Try
      ' Give some space...
      Application.DoEvents()
      ' Skip optional spaces
      bnf_Space(strIn, intPos)
      ' Try get left bracket
      If (bnf_Expect(strIn, intPos, "(")) Then
        ' Skip optional spaces
        bnf_Space(strIn, intPos)
        ' In principle a <Label> should follow, but if it doesn't, we insert a NOLABEL label
        If (bnf_Label(strIn, intPos, strLabel)) Then
          ' Skip at least one space
          If (bnf_Space(strIn, intPos) > 0) Then
            'This is a valid node
            bIsValid = True
          End If
        Else
          ' No label followed, so we think of a label
          strLabel = "UNKNOWN"
          ' We make the node valid
          bIsValid = True
        End If
        ' Did we find a valid node opening?
        If (bIsValid) Then
          ' Only create a new datarow for non-ID stuff
          If (strLabel = "ID") Then
            ' Set the IDnode flag
            bIsIDnode = True
          Else
            ' Do note if this is a CODE node
            If (strLabel = "CODE") Then bIsCode = True
            ' Create a new <eTree> element
            If (Not CreateNewRow(tdlThis, "eTree", "Id", intNodeId, dtrNode)) Then
              ' There was an error
              Return False
            End If
            ' Indicate a new tree was made
            bMade = True
            ' Access the new row
            With dtrNode
              ' Set the parent for this new row
              .SetParentRow(dtrParent)
              ' Add the label for this new row
              .Item("Label") = strLabel
            End With
          End If
          ' Expect to find <NodeContent>
          If (bnf_NodeContent(tdlThis, strIn, intPos, bIsLeaf, strNode, dtrNode, bIsIDnode, bIsCode)) Then
            ' ============= Check for empty node ============
            ' If (strNode = "") Then Stop
            ' ===============================================
            ' Actions here COULD depend on the label name/type
            Select Case strLabel
              Case "CODE"
                ' Check the content of the node for <Coref>
                If (IsCoref(strNode, intMyId, intRefId, strRefType, intNdDist)) Then
                  ' This is a <Coref> node -- take the information and process it
                  With dtrParent
                    .Item("MyId") = intMyId
                    .Item("RefId") = intRefId
                    .Item("RefType") = strRefType
                    .Item("NdDist") = intNdDist
                  End With
                  ' Delete the <eTree> CODE node
                  dtrNode.Delete()
                End If
                ' Don't return the content of a CODE node
                strNode = ""
                ' Set the type of the node
                strType = "code"
              Case ".", ",", "E_S"
                ' Don't prepend a space - keep [strNode] as is
                ' Set the type of the node
                strType = "punct"
              Case "ID"
                ' Set the type of the node
                strType = "id"
                loc_strLineId = strNode
                ' ================ DEBUG ============
                ' If (InStr(loc_strLineId, "480") > 0) Then Stop
                ' ===================================
              Case Else
                ' Return the content of this node added with space
                If (Trim(strNode) <> "") AndAlso (bIsLeaf) Then
                  ' Add a space to the word
                  strNode = " " & strNode
                End If
                ' Set the node type
                strType = "other"
            End Select
            ' Skip optional spaces
            bnf_Space(strIn, intPos)
            '' Check if something else than ) is following
            'If (Not bnf_Test(strIn, intPos, ")")) AndAlso (intPos < strIn.Length) AndAlso (Not bnf_Test(strIn, intPos, "(")) Then
            '  ' No ")" is following, so try to process the text as CODE
            '  Stop
            'End If
            ' Expect to find closing bracket ")"
            If (bnf_Expect(strIn, intPos, ")")) Then
              ' We're okay, return ...
              strBack &= strNode
              ' Was our label "UNKNOWN"??
              If (strLabel = "UNKNOWN") Then
                ' Perhaps just mention it?
                objError.Add("Added <UNKNOWN> at: " & strLastLocation & _
                                           "Content=" & strNode)
              End If
              ' Return success
              Return True
            End If
          Else
            ' This node appears to have no content
            ' Skip optional spaces
            bnf_Space(strIn, intPos)
            ' Expect to find closing bracket ")"
            If (bnf_Expect(strIn, intPos, ")")) Then
              ' We're okay, return ...
              strBack &= strNode
              ' Was our label "UNKNOWN"??
              If (strLabel = "UNKNOWN") Then
                ' Perhaps just mention it?
                objError.Add("Added <UNKNOWN> at: " & strLastLocation & _
                                           "Content=" & strNode)
              End If
              ' Return success
              Return True
            End If
          End If
        End If
      End If
      ' If a new element was made, clear it up
      If (bMade) Then
        ' Show the error
        ShowError(strIn, intBack, intPos)
        ' Delete this element
        dtrNode.Delete()
        ' Make sure deletion is processed
        tdlThis.AcceptChanges()
      End If
      ' Wind back and return failure
      intPos = intBack
      bnf_Node = False
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/bnf_Node error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  IsCoref
  ' Goal :  If the node's content is <Coref>, then return TRUE, and also return the values found
  ' History:
  ' 01-03-2010  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function IsCoref(ByVal strIn As String, ByRef intMyId As Integer, ByRef intRefId As Integer, _
                           ByRef strRefType As String, ByRef intNdDist As Integer) As Boolean
    Dim docThis As New Xml.XmlDocument  ' Document to load the XML content
    Dim ndThis As Xml.XmlNode           ' The <Coref> node

    ' test if this is Coref
    If (InStr(strIn, "coref", CompareMethod.Text) > 0) Then
      ' This is coref, so get it!
      docThis.LoadXml(strIn.Replace("_", " "))
      ' Get the top node
      ndThis = docThis.SelectSingleNode("//Coref")
      ' At least get the Id value
      intMyId = ndThis.Attributes("Id").Value
      ' Try get the other attributes one by one
      If (ndThis.Attributes("Ref") Is Nothing) Then
        intRefId = -1
      Else
        intRefId = ndThis.Attributes("Ref").Value
      End If
      If (ndThis.Attributes("Type") Is Nothing) Then
        strRefType = ""
      Else
        strRefType = ndThis.Attributes("Type").Value
      End If
      If (ndThis.Attributes("NdDist") Is Nothing) Then
        intNdDist = 0
      Else
        intNdDist = ndThis.Attributes("NdDist").Value
      End If
      ' Return success
      Return True
    Else
      ' Return failure
      Return False
    End If
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  ShowError
  ' Goal :  Show where there is an error
  ' History:
  ' 14-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Sub ShowError(ByRef strIn As String, ByVal intPos As Integer, Optional ByVal intPos2 As Integer = 0)
    Dim intLine As Integer = 0    ' Line where the error occurs

    ' Switch to the SYNTAX page
    frmMain.TabControl1.SelectedTab = frmMain.tpSyntax
    ' Show where the error has occurred
    With frmMain.tbEdtSyntax
      ' Set the string
      .Text = strIn
      ' Determine the mode
      If (intPos2 = 0) Then
        If (intPos2 > intPos) Then
          .SelectionStart = intPos2
          .SelectionLength = intPos - intPos2
        Else
          .SelectionStart = intPos
          .SelectionLength = intPos2 - intPos
        End If
      Else
        ' Select the error place
        .SelectionStart = intPos
        .SelectionLength = 10
      End If
      .SelectionColor = Color.Red
      ' Scroll to this position
      .ScrollToCaret()
      ' Find the line where the error is
      intLine = .GetLineFromCharIndex(intPos)
    End With
    ' Show error message
    With frmErrWait
      ' Show it
      .Show()
      ' Wait until the correct button is pressed
      While (Not .ErrorViewed)
        Application.DoEvents()
      End While
      ' Reset the errorviewed flag
      .ErrorViewed = False
    End With
  End Sub
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  bnf_NodeContent
  ' Goal :  Parse the input PSD string starting from position [intPos]
  ' Note :  Here is the BNF for the PSD input:
  '   <NodeContent> ::= ( <Node>$ [<FaultLex>] ) | <NodeLex>
  ' History:
  ' 14-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function bnf_NodeContent(ByRef tdlThis As DataSet, ByRef strIn As String, ByRef intPos As Integer, _
          ByRef bIsLeaf As Boolean, ByRef strBack As String, ByRef dtrParent As DataRow, _
          ByVal bIsIDnode As Boolean, ByVal bIsCode As Boolean) As Boolean
    Dim strNodeContent As String = ""   ' Start with empty return
    Dim strType As String = ""          ' Type of the node

    Try
      ' There should at least be ONE <Node>
      If (bnf_Node(tdlThis, strIn, intPos, strNodeContent, strType, dtrParent)) Then
        ' Try get as many <Node> as possible
        Do
          ' Add to the [strBack]
          strBack &= strNodeContent
          ' Reset [strNodeLex]
          strNodeContent = ""
          ' Try get another <Node>
        Loop While (bnf_Node(tdlThis, strIn, intPos, strNodeContent, strType, dtrParent))
        ' It could be that we have a parentless lexeme: (LABEL (NODE aap) noot)
        If (bnf_FaultLex(tdlThis, strIn, intPos, strNodeContent, dtrParent)) Then
          ' Okay, this can happen - we don't NEED to do anything here!
          ' Perhaps just mention it?
          objError.Add("Added <UNKNOWN> at: " & strLastLocation & _
                                     "Content=" & strNodeContent)
        End If
        ' Indicate we are NO leaf
        bIsLeaf = False
        ' Return success
        bnf_NodeContent = True
      ElseIf (bnf_NodeLex(tdlThis, strIn, intPos, strNodeContent, dtrParent, bIsIDnode, bIsCode)) Then
        ' Indicate we are a leaf
        bIsLeaf = True
        ' The <NodeContent> consists of a <NodeLex>
        strBack = strNodeContent
        ' Return success
        Return True
      Else
        ' Indicate we are NO leaf (for what it's worth)
        bIsLeaf = False
        ' No, it doesn't work out - return failure
        bnf_NodeContent = False
      End If
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/bnf_NodeContent error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  GetLeafType
  ' Goal :  Determine the kind of leaf this is
  ' History:
  ' 15-02-2010  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Public Function GetLeafType(ByVal strIn As String) As String
    ' Does it contain an asterisk?
    If (InStr(strIn, "*") > 0) Then
      ' It is an asterisk type
      Return "Star"
    ElseIf (strIn = "0") Then
      ' It is a zero
      Return "Zero"
    ElseIf (InStr(PUNCTUATION, strIn) > 0) Then
      ' It is punctuation
      Return "Punct"
    Else
      ' It is a lexeme
      Return "Vern"
    End If
  End Function

  '------------------------------------------------------------------------------------------------------------------
  ' Name :  bnf_NodeLex
  ' Goal :  Parse the input PSD string starting from position [intPos]
  ' Note :  Here is the BNF for the PSD input:
  '   <NodeLex>     ::= sp* <Lexeme> <Node>*
  ' History:
  ' 14-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function bnf_NodeLex(ByRef tdlThis As DataSet, ByRef strIn As String, ByRef intPos As Integer, _
        ByRef strBack As String, ByRef dtrParent As DataRow, ByVal bIsIDnode As Boolean, _
        ByVal bIsCode As Boolean) As Boolean
    Dim intBack As Integer = intPos   ' Roll back position
    Dim intLeafId As Integer = -1     ' ID of the leaf
    Dim dtrLeaf As DataRow = Nothing  ' The new Leaf element we create here
    Dim strNodeLex As String = ""     ' Start with empty return
    Dim strType As String = ""        ' Type of the node
    Dim bCanCreate As Boolean         ' Whether we can create a new <eLeaf> element

    Try
      ' Skip possible spaces
      bnf_Space(strIn, intPos)
      ' There should at least be one <Lexeme>
      If (bnf_Lexeme(strIn, intPos, strNodeLex)) Then
        ' Try get as many <Node> as possible
        Do
          ' ONly create a row for non-ID node stuff
          bCanCreate = (Not bIsIDnode) AndAlso _
            (Not (bIsCode And (InStr(strNodeLex, "coref", CompareMethod.Text) > 0))) AndAlso _
            (strNodeLex <> "")
          If (bCanCreate) Then
            ' Create a new <eLeaf> element
            If (Not CreateNewRow(tdlThis, "eLeaf", "", intLeafId, dtrLeaf)) Then
              ' There was an error
              Return False
            End If
            ' Access the new element
            With dtrLeaf
              ' Set the parent of this leaf
              .SetParentRow(dtrParent)
              ' Set the lexical content of this leaf
              .Item("Text") = strNodeLex
              ' Set the type of the leaf
              .Item("Type") = GetLeafType(strNodeLex)
            End With
          End If
          ' Only certain lexemes may be taken into account
          If (Left(strNodeLex, 1) <> "*") AndAlso (strNodeLex <> "0") Then
            ' Add to the [strBack]
            strBack &= Trim(strNodeLex)
          End If
          ' Reset [strNodeLex]
          strNodeLex = ""
          ' Try get another <Node> - this goes under our own parent
        Loop While (bnf_Node(tdlThis, strIn, intPos, strNodeLex, strType, dtrParent))
        ' Yes, return success
        Return True
      Else
        ' No, wind back
        intPos = intBack
        ' No return failure
        Return False
      End If
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/bnf_NodeLex error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  CorefAdapt
  ' Goal :  Convert a string into correct XML format if it is a Coref stuff
  ' History:
  ' 28-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function CorefAdapt(ByVal strIn As String) As String
    ' Does it contain the keyword "Coref"?
    If (InStr(strIn, "coref", CompareMethod.Text) > 0) Then
      ' Convert the string
      strIn = strIn.Replace("_", " ")
    End If
    ' Return the input string, possibly adapted
    Return strIn
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  bnf_FaultLex
  ' Goal :  Parse the input PSD string starting from position [intPos]
  ' Note :  Here is the BNF for the PSD input:
  '   <NodeLex>     ::= sp* <Lexeme>
  ' History:
  ' 28-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function bnf_FaultLex(ByRef tdlThis As DataSet, ByRef strIn As String, ByRef intPos As Integer, _
        ByRef strBack As String, ByRef dtrParent As DataRow) _
        As Boolean
    Dim intBack As Integer = intPos   ' Roll back position
    Dim intTreeId As Integer = -1     ' ID of the eTree element
    Dim intLeafId As Integer = -1     ' ID of the leaf
    Dim dtrLeaf As DataRow = Nothing  ' The new Leaf element we create here
    Dim dtrNode As DataRow = Nothing  ' The new Tree element we create here
    Dim strNodeLex As String = ""     ' Start with empty return
    Dim strType As String = ""        ' Type of the node

    Try
      ' Skip possible spaces
      bnf_Space(strIn, intPos)
      ' There should at least be one <Lexeme>
      If (bnf_Lexeme(strIn, intPos, strNodeLex)) Then
        ' Create a new <eTree> element
        If (Not CreateNewRow(tdlThis, "eTree", "Id", intTreeId, dtrNode)) Then
          ' There was an error
          Return False
        End If
        ' Access the new element
        With dtrNode
          ' Set the parent
          .SetParentRow(dtrParent)
          ' Set the label
          ' Used to be: .Item("Label") = "UNKNOWN"
          .Item("Label") = "CODE"
        End With
        ' Create a new <eLeaf> element
        If (Not CreateNewRow(tdlThis, "eLeaf", "", intLeafId, dtrLeaf)) Then
          ' There was an error
          Return False
        End If
        ' Access the new element
        With dtrLeaf
          ' Set the parent of this leaf
          .SetParentRow(dtrNode)
          ' Set the lexical content of this leaf
          .Item("Text") = strNodeLex
          ' Set the type of the leaf
          .Item("Type") = GetLeafType(strNodeLex)
        End With
        ' Add to the [strBack]
        strBack &= Trim(strNodeLex)
        ' Reset [strNodeLex]
        strNodeLex = ""
        ' Yes, return success
        Return True
      Else
        ' No, wind back
        intPos = intBack
        ' No return failure
        Return False
      End If
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/bnf_FaultLex error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  bnf_Lexeme
  ' Goal :  Try to get a lexeme and pass it back in [strBack]
  ' History:
  ' 14-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function bnf_Lexeme(ByRef strIn As String, ByRef intPos As Integer, ByRef strBack As String) As Boolean
    Dim strLex As String = ""   ' Text of the lexeme

    Try
      ' Try get the lexeme
      If (bnf_Word(strIn, intPos, "()", strLex)) Then
        ' We have a lexeme
        strBack = Trim(strLex)
        ' Return success
        bnf_Lexeme = True
      Else
        ' Set back string to ""
        strBack = ""
        ' return failure
        bnf_Lexeme = False
      End If
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/bnf_Lexeme error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  bnf_Label
  ' Goal :  Try to get a label and pass it back in [strBack]
  ' History:
  ' 14-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function bnf_Label(ByRef strIn As String, ByRef intPos As Integer, ByRef strBack As String) As Boolean
    Dim strLabel As String = ""   ' Text of the label

    Try
      ' A label may NOT start with a (!!!
      If (Mid(strIn, intPos, 1) = "(") Then
        ' Return failure
        strBack = ""
        Return False
      End If
      ' Try get the label
      If (bnf_Word(strIn, intPos, strNoWord, strLabel)) Then
        ' We have a label
        strBack = strLabel
        ' Return success
        Return True
      Else
        ' Set back string to ""
        strBack = ""
        ' return failure
        Return False
      End If
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/bnf_Lexeme error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  bnf_Word
  ' Goal :  Anything can go in a word, except for a space sign and a left bracket (
  ' History:
  ' 14-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function bnf_Word(ByRef strIn As String, ByRef intPos As Integer, _
          ByVal strBreakChar As String, ByRef strBack As String) As Boolean
    Dim intBack As Integer = intPos   ' Roll back position

    Try
      ' Try to get anything except SPACE signs and )
      While (InStr(strBreakChar, Mid(strIn, intPos, 1)) = 0) AndAlso (intPos < Len(strIn))
        ' Go to the next position
        intPos += 1
      End While
      ' Return the result
      strBack = Mid(strIn, intBack, intPos - intBack)
      ' Return success
      bnf_Word = (intPos > intBack)
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/bnf_Word error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  IsComment
  ' Goal :  Check whether this is a comment
  ' History:
  ' 26-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function IsComment(ByRef strIn As String, ByRef intPos As Integer) As Boolean
    Dim intBack As Integer = intPos ' Roll back position

    Try
      ' Check comment start out
      If (Mid(strIn, intPos, Len(strComSt)) = strComSt) Then
        ' Skip until the first delimiter
        intPos = InStr(intPos + 1, strIn, strdelim)
        ' Found something
        If (intPos > 0) Then
          ' Succeeded
          Return True
        End If
      End If
      ' Roll back
      intPos = intBack
      ' Return failure
      Return False
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/IsComment error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  bnf_SkipLine
  ' Goal :  Skip until the end of line
  ' History:
  ' 14-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function bnf_SkipLine(ByRef strIn As String, ByRef intPos As Integer) As Boolean
    Dim intI As Integer ' Dummy

    Try
      ' Skip everything until we reach the first of CR or LF
      While ((InStr(vbCrLf, Mid(strIn, intPos, 1)) = 0) AndAlso (intPos < Len(strIn)))
        ' Skip one space
        intPos += 1
      End While
      ' Skip spaces
      intI = bnf_Space(strIn, intPos)
      ' Return positively
      Return True
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/bnf_SkipLine error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  bnf_Space
  ' Goal :  Skip any number of spaces
  ' History:
  ' 14-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function bnf_Space(ByRef strIn As String, ByRef intPos As Integer) As Integer
    Dim intNum As Integer = 0 ' Number of spaces skipped

    Try
      ' Try skip spaces
      While ((InStr(strSpace, Mid(strIn, intPos, 1)) > 0) OrElse (IsComment(strIn, intPos))) _
             AndAlso (intPos < Len(strIn))
        ' Skip one space
        intPos += 1
        intNum += 1
      End While
      ' Return the number of spaces actually skipped
      bnf_Space = intNum
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/bnf_Space error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  bnf_Expect
  ' Goal :  Expect to find the string [strExpect] here
  ' History:
  ' 14-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function bnf_Expect(ByRef strIn As String, ByRef intPos As Integer, ByVal strExpect As String) As Boolean
    Dim intBack As Integer = intPos   ' Roll back position

    Try
      ' Test whether the string is here
      If (Mid(strIn, intPos, Len(strExpect)) = strExpect) Then
        ' Skip this number of places
        intPos += Len(strExpect)
        ' Return success
        bnf_Expect = True
      Else
        ' Return failure
        bnf_Expect = False
      End If
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/bnf_Expect error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  bnf_Test
  ' Goal :  Expect to find the string [strExpect] here
  ' History:
  ' 14-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Private Function bnf_Test(ByRef strIn As String, ByRef intPos As Integer, ByVal strExpect As String) As Boolean
    Dim intBack As Integer = intPos   ' Roll back position

    Try
      ' Test whether the string is here
      Return (Mid(strIn, intPos, Len(strExpect)) = strExpect)
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/bnf_Test error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  ' -------------------------------------------------------------------------------------------------------
  ' Name: GetTbSpecChar
  ' Goal: Given a special character with a + sign, convert into "normal" English
  ' History:
  ' 27-11-2008    ERK Created
  ' -------------------------------------------------------------------------------------------------------
  Private Function GetTbSpecChar(ByVal strIn As String) As String
    Dim intPos As Integer
    Dim strEval As String   ' The one character we are evaluating

    ' Get the evaluation character
    strEval = Left(strIn, 1)
    ' Get position in input string
    intPos = InStr(SPEC_CHAR_IN, strEval)
    If (intPos > 0) Then
      GetTbSpecChar = Mid(SPEC_CHAR_OUT, intPos, 1)
    Else
      ' Try an error character
      intPos = InStr(SPEC_CHAR_ERR, strEval)
      If (intPos > 0) Then
        GetTbSpecChar = Mid(SPEC_CHAR_EOUT, intPos, 1) & strEval
      Else
        ' Output the input with + sign
        GetTbSpecChar = "+" & strIn
      End If
    End If
  End Function
  ' -------------------------------------------------------------------------------------------------------
  ' Name: DoAddWSJid
  ' Goal: Add an ID element to the WSJ file
  ' History:
  ' 03-03-2010    ERK Created
  ' -------------------------------------------------------------------------------------------------------
  Public Function DoAddWSJid(ByVal strFile As String, ByRef strText As String) As Boolean
    Dim strAdd As String = ""     ' What needs to be added
    Dim intPos As Integer = 1     ' Position within the string
    Dim intHook As Integer = 0    ' Number of brackets
    Dim intLineId As Integer = 0  ' The line ID number

    Try
      ' Go into the loop
      While (NextHook(strText, intPos, intHook)) AndAlso (intPos < strText.Length)
        ' Have we reached ZERO?
        If (intHook = 0) Then
          ' Increment the lineid
          intLineId += 1
          ' What needs to be added?
          strAdd = " (ID " & strFile & ",s" & Format(intLineId, "0000") & ") "
          ' We need to insert the ID tag
          strText = Left(strText, intPos - 1) & strAdd & Mid(strText, intPos)
          ' The position should at least skip [strAdd]
          intPos += strAdd.Length + 1
        Else
          ' Increment our position
          intPos += 1
        End If
      End While
      ' Return success
      Return True
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/DoAddWSJid error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' -------------------------------------------------------------------------------------------------------
  ' Name:   NextHook
  ' Goal:   Find the first following bracket ( or )
  ' Return: TRUE if success
  ' History:
  ' 03-03-2010    ERK Created
  ' -------------------------------------------------------------------------------------------------------
  Private Function NextHook(ByRef strText As String, ByRef intPos As Integer, ByRef intHook As Integer) As Boolean
    Dim intLeft As Integer = 0    ' Left hook
    Dim intRight As Integer = 0   ' Right hook

    Try
      ' Determine where the hooks are
      intLeft = InStr(intPos, strText, "(")
      intRight = InStr(intPos, strText, ")")
      ' Which is closer by?
      If (intLeft = 0) AndAlso (intRight = 0) Then
        ' Return failure
        Return False
      ElseIf (intLeft < intRight) AndAlso (intLeft > 0) Then
        ' Take the left hook
        intPos = intLeft
        ' Increase
        intHook += 1
      ElseIf (intRight > 0) Then
        ' Take the right hook
        intPos = intRight
        ' Decrement
        intHook -= 1
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Note error
      HandleErr("modParse/NextHook error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
End Module
