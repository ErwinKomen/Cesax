Imports System.Xml
Module modCorpus

  '---------------------------------------------------------------------------------------------------------
  ' Name:       LoadOneDbFile()
  ' Goal:       Load one Corpus-database file
  ' History:
  ' 26-09-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function LoadOneDbFile(ByVal strFile As String) As Boolean
    Dim arFile() As String  ' Array of files
    Dim intI As Integer     ' Counter

    Try
      ' Validate
      Select Case XmlFileType(strFile)
        Case "TEI"      ' Proper
          Status("Use File/Open to open a psdx text file") : Return False
        Case "CrpOview" ' Result file or database
        Case Else       ' Unknown file type
          Status("Unknown file type") : Return False
      End Select
      ' Remove any old temporary database files
      arFile = IO.Directory.GetFiles(GetSetDir, "*-tmp.xml")
      For intI = arFile.Length - 1 To 0 Step -1
        Try
          ' Delete temporary file
          IO.File.Delete(arFile(intI))
        Catch ex As Exception
          ' No problem if there are any problems here
        End Try
      Next intI
      ' Load the database as an XmlDocument
      Status("Loading database as xmlDocument...")
      If (Not ReadXmlDoc(strFile, pdxCrpDbase)) Then Return False
      ' Convert part of this XML document into an xml file that contains just enough information for the [dgvDbResult]
      'strCurrentDbFile = IO.Path.GetDirectoryName(strFile) & "\" & _
      '                    IO.Path.GetFileNameWithoutExtension(strFile) & "-tmp.xml"
      strCurrentDbFile = GetSetDir() & "\" & IO.Path.GetFileNameWithoutExtension(strFile) & "-tmp.xml"
      If (Not CrpDbaseExtract(pdxCrpDbase, strCurrentDbFile)) Then Return False
      ' Load the database
      Status("Loading part of database as xmlDataDocument...")
      If (Not XmlToDataset(strCrpResults, strCurrentDbFile, tdlResults, pdxResults)) Then Status("Could not open corpus database") : Return False
      '' Load the database
      'If (Not ReadDataset(strCrpResults, strResultDb, tdlResults)) Then Status("Could not open dataset") : Exit Function
      ' Return positively
      Status("Database is loaded")
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/LoadOneDbFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       CrpDbaseExtract()
  ' Goal:       Extract the part from a corpus databasefile that is needed for [dgvDbResult]
  ' History:
  ' 26-09-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Function CrpDbaseExtract(ByRef pdxThis As XmlDocument, ByVal strOutFile As String) As Boolean
    Dim ndxRes As XmlNode         ' Current result node
    Dim ndxThis As XmlNode        ' Working node
    Dim ndxList As XmlNodeList    ' List of nodes
    Dim intResNum As Integer      ' Number of results
    Dim strAttrList As String     ' List of attributes and their values
    Dim strAttrVal As String      ' Value of attribute
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage
    Dim arGenName() As String = {"ProjectName", "Created", "DstDir", "SrcDir", "Notes", "Analysis"}
    Dim arGenVal() As String = {"-", Format(Now, "G"), "", "", "-", ""}
    ' Dim arResAttr() As String = {"ResId", "eTreeId", "File", "Search", "Period", "TextId", "forestId", "Cat", "Notes", "Select", "Status"}
    Dim arResAttr() As String = {"ResId", "Period", "TextId", "forestId", "Cat", "Notes", "Select", "Status"}

    Try
      ' Validate
      If (pdxThis Is Nothing) OrElse (strOutFile = "") Then Return False
      ' TODO: validation of the type of database we have
      ' ================================================
      ' Initialise the collection for the results
      colCrpResults.Clear() : intCrpResultsId = 0
      ' Initialise database and add <general> section
      colCrpResults.Add("<?xml version='1.0' standalone='yes'?>")
      colCrpResults.Add("<CrpOview>")
      colCrpResults.Add(" <General>")
      ' Get to the general part of [pdxThis]
      ndxThis = pdxThis.SelectSingleNode("./descendant::General")
      If (ndxThis Is Nothing) Then Return False
      ' Get the necessary general information into [colCrpResults]
      ndxList = ndxThis.ChildNodes
      For intI = 0 To ndxList.Count - 1
        With ndxList(intI)
          intJ = Array.FindIndex(arGenName, Function(strName As String) strName = .Name)
          If (intJ >= 0) Then
            ' Add the value of this item on the correct place
            arGenVal(intJ) = XmlEscape(.InnerText)
          End If
        End With
      Next intI
      ' Add results to [colCrpResult]
      For intI = 0 To arGenVal.Length - 1
        colCrpResults.Add("  <" & arGenName(intI) & ">" & arGenVal(intI) & "</" & arGenName(intI) & ">")
      Next intI
      colCrpResults.Add(" </General>")
      ' Start output file with this general section
      IO.File.WriteAllText(strOutFile, "")
      FlushResultsDbase(strOutFile)
      ' Find the first result
      ndxRes = ndxThis.SelectSingleNode("./following-sibling::Result")
      intResNum = ndxRes.SelectNodes("./following-sibling::Result").Count + 1
      intJ = 1
      While (ndxRes IsNot Nothing)
        ' Show where we are
        intPtc = (intJ * 100) \ intResNum
        Status("Dbase extraction " & intPtc & "%", intPtc)
        ' Check for interrupt
        If (bInterrupt) Then Return False
        ' Gather the <result> information
        strAttrList = "<Result "
        ' Process the attributes of this result
        For intI = 0 To arResAttr.Length - 1
          ' Determine the value of the attribute...
          If (ndxRes.Attributes(arResAttr(intI)) Is Nothing) Then
            strAttrVal = ""
          Else
            strAttrVal = XmlEscape(ndxRes.Attributes(arResAttr(intI)).Value)
          End If
          ' Add this attribute + value to the list
          strAttrList &= arResAttr(intI) & "=" & """" & strAttrVal & """" & " "
        Next intI
        ' Finish attributes list
        colCrpResults.Add(strAttrList & "/>")
        ' Flush results every XX items
        If (colCrpResults.Count > 2047) Then
          ' Flush results
          FlushResultsDbase(strOutFile)
        End If
        ' Find next result
        ndxRes = ndxRes.NextSibling : intJ += 1
      End While
      ' Finish the CrpOview
      FinishResultsDbase(strOutFile)
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/CrpDbaseExtract error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   FlushResultsDbase
  ' Goal:   Save corpus results database and reset collection
  ' History:
  ' 25-09-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function FlushResultsDbase(ByVal strFile As String) As Boolean
    Try
      ' Append what we have
      colCrpResults.Flush(strFile)
      'IO.File.AppendAllText(strFile, colCrpResults.Text)
      ' Reset collection
      colCrpResults.Clear()
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modCorpus/FlushResultsDbase error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   FinishResultsDbase
  ' Goal:   Finish the corpus results database
  ' History:
  ' 25-09-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function FinishResultsDbase(ByVal strFile As String) As Boolean
    Try
      colCrpResults.Add("</CrpOview>")
      ' Append final section
      colCrpResults.Flush(strFile)
      'IO.File.AppendAllText(strFile, colCrpResults.Text)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modCorpus/FinishResultsDbase error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   QuoteIfString
  ' Goal:   If this is a string, put it in quotation marks
  ' History:
  ' 07-04-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function QuoteIfString(ByVal strIn As String) As String
    Try
      If (Not IsNumeric(strIn)) Then strIn = """" & strIn & """"
      ' Make sure CR and LF are replaced by something else
      strIn = strIn.Replace(vbCrLf, vbLf)
      strIn = strIn.Replace(vbCr, vbLf)
      strIn = strIn.Replace(vbLf, "<br>")
      Return strIn
    Catch ex As Exception
      ' Show error
      HandleErr("modCorpus/QuoteIfString error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   UnQuote
  ' Goal:   If this is a string, take the quotation marks away
  ' History:
  ' 21-04-2014  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function UnQuote(ByVal strIn As String) As String
    Try
      If (Left(strIn, 1) = """") AndAlso (Right(strIn, 1) = """") Then
        strIn = Mid(strIn, 2)
        strIn = Left(strIn, strIn.Length - 1)
      End If
      Return strIn
    Catch ex As Exception
      ' Show error
      HandleErr("modCorpus/UnQuote error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetResultsDbaseFeatures
  ' Goal:   Get names for the features in the results dbase, by checking the first <Result> record
  ' History:
  ' 26-09-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetResultsDbaseFeatures(ByVal strAttrList As String, ByRef arName() As String, ByVal strFname As String, _
                                          ByRef ndxThis As XmlNode) As Boolean
    ' Dim ndxRes As XmlNode       ' One result node
    Dim ndxF As XmlNode         ' One feature node
    Dim strText As String = ""  ' Gather results here
    Dim strExmp As String = ""  ' Example
    Dim arAttr(0) As String      ' Attribute list
    Dim intPos As Integer       ' Position within string
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      'If (pdxCrpDbase Is Nothing) Then Return False
      'ndxRes = pdxCrpDbase.SelectSingleNode("./descendant::Result")
      If (ndxThis Is Nothing) Then Return False
      If (ndxThis.Name <> "Result") Then Return False
      ' Initialise strText
      If (strFname = "Name") Then
        strText = strAttrList.Replace(";", vbTab)
      ElseIf (strFname = "Value") Then
        If (strAttrList = "") Then
          strText = ""
        Else
          ' Need the values
          arAttr = Split(strAttrList, ";")
          ' Any attribute?
          If (arAttr.Length > 0) Then
            ' Add first attribute
            strText = QuoteIfString(XmlEscape(GetAttrValue(ndxThis, arAttr(0))))
            For intI = 1 To arAttr.Length - 1
              ' Add this attribute
              strText &= vbTab & QuoteIfString(XmlEscape(GetAttrValue(ndxThis, arAttr(intI))))
            Next intI
          End If
        End If
      Else
        Return False
      End If
      ' Find out which Feature columns there are
      For Each ndxF In ndxThis.SelectNodes("./child::Feature")
        ' Possibly add semicolon
        If (strText <> "") Then strText &= vbTab
        ' Add the heading for this child
        If (strFname = "Name") Then
          strText &= XmlEscape(GetAttrValue(ndxF, strFname))
        Else
          strText &= QuoteIfString(XmlEscape(GetAttrValue(ndxF, strFname)))
        End If
      Next ndxF
      ' Add Example
      If (strFname = "Name") Then
        ' Add the name
        strText &= vbTab & "Example"
      Else
        ' Add the example itself
        strExmp = ndxThis.SelectSingleNode("./descendant::Text").InnerText
        intPos = InStr(strExmp, "<Font")
        If (intPos > 0) Then
          intPos = InStr(intPos, strExmp, ">")
          If (intPos > 0) Then
            strExmp = Mid(strExmp, intPos + 1)
            intPos = InStr(strExmp, "</Font")
            If (intPos > 0) Then
              strExmp = Left(strExmp, intPos - 1)
            End If
          End If
        End If
        strText &= vbTab & strExmp
      End If
      ' Make the array to return
      arName = Split(strText, vbTab)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modCorpus/GetResultsDbaseFeatures error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoCorpusFeaturesFromFile
  ' Goal:   Copy features from a .xml file to the current database
  ' History:
  ' 10-07-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function DoCorpusFeaturesFromFile(ByVal strImpFtFile As String, ByVal strCurrentFilter As String, ByVal strFeatDbase As String, _
                      ByVal strMethod As String, ByVal strFeatValues As String, ByVal strFeatBlackList As String, ByVal bUseStatus As Boolean, _
                      ByVal bUseNotes As Boolean, ByVal bUsePde As Boolean) As Boolean
    Dim strFile As String       ' Current file we are opening
    Dim strLastFile As String   ' Last file we were working on
    Dim arFile() As String      ' Array of files
    Dim strFeatValue As String  ' Value for the feature
    Dim strExpFeatFile As String  ' File name we export the features to
    Dim strStatus As String     ' Status
    Dim strNotes As String      ' Notes
    Dim strPde As String        ' PDE (back translation)
    Dim intForestId As Integer  ' Forest id
    Dim intEtreeId As Integer   ' Etree Id
    Dim intResId As Integer     ' Recovered original resid
    Dim ndxFeat As XmlNode      ' Node containing feature
    Dim ndxForest As XmlNode    ' Forest node
    Dim ndxEtree As XmlNode     ' Constituent
    Dim ndxList As XmlNodeList  ' List of nodes in the corpus database we are visiting
    Dim ndxSpec As XmlNodeList  ' List for [bSpecial]
    Dim dtrThis As DataRow      ' One datarow
    Dim dtrFeat() As DataRow    ' Datarow for this feature
    Dim dtrFound() As DataRow   ' Walking through the dataset
    Dim dtrRes() As DataRow     ' Location in tdlResults
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter
    Dim intK As Integer         ' Counter
    Dim intPtc As Integer       ' Percentage
    Dim intCount As Integer     ' Number of changes
    Dim intRecOk As Integer     ' Number of well recovered entries
    Dim intRecNo As Integer     ' Number of failed recovers
    Dim bSpecial As Boolean = False  ' SPecial treatment -- AmbiLemma problem
    Dim tdlFtDbase As DataSet = Nothing   ' Dataset to hold the features

    Try
      ' Validate
      If (Not IO.File.Exists(strImpFtFile)) Then Return False
      ' Read the file into a database
      If (Not ReadDataset("CrpResult.xsd", strImpFtFile, tdlFtDbase)) Then Return False
      ' Now get on with it...
      If (strCurrentFilter = "") Then
        ndxList = pdxCrpDbase.SelectNodes("./descendant::Result")
      Else
        ndxList = pdxCrpDbase.SelectNodes(strCurrentFilter, conTb)
      End If
      SetXmlDocument(pdxCrpDbase)
      strLastFile = "" : intCount = 0 : intRecOk = 0 : intRecNo = 0
      ' Visit them all
      For intI = 0 To ndxList.Count - 1
        ' Access the information in this node
        With ndxList(intI)
          ' Where are we?
          intPtc = (intI + 1) * 100 \ ndxList.Count
          Status("FeaturesToFile " & intPtc & "%", intPtc)
          CorefProgress(.Attributes("TextId").Value)
          ' Get forest id etc
          intForestId = .Attributes("forestId").Value
          intEtreeId = .Attributes("eTreeId").Value
          ' Make sure we get the file-name without the path
          strFile = IO.Path.GetFileName(.Attributes("File").Value)
          ' Recover the original ResID as stored in OviewId
          intResId = .Attributes("ResId").Value
          ' ============ DEBUG ==========
          '  If (strFile Like "cobede*" AndAlso intForestId = 9) Then Stop
          ' =============================
          ' Get the feature value
          ndxFeat = .SelectSingleNode("./descendant::Feature[@Name='" & strFeatDbase & "']")
          If (ndxFeat IsNot Nothing) Then
            ' Action depends on the method for searching we use
            Select Case strMethod
              Case "ResId"
                ' Each entry is uniquely determined by its ResId
                dtrFound = tdlFtDbase.Tables("Result").Select("ResId=" & intResId)
              Case Else
                ' Use text/forest/etree
                ' Look for the corresponding entry in the file we have read
                dtrFound = tdlFtDbase.Tables("Result").Select("File = '" & strFile & "' AND forestId=" & intForestId & _
                                                              " AND eTreeId=" & intEtreeId)
                ' Make sure we get results
                If (dtrFound.Length = 0) Then
                  ' Look for fuller path
                  dtrFound = tdlFtDbase.Tables("Result").Select("File LIKE '*\" & strFile & "' AND forestId=" & intForestId & _
                                                                " AND eTreeId=" & intEtreeId)
                End If
                ' Be on the watch for ambiguity (too much results)
                If (dtrFound.Length > 1) Then
                  ' Try reduce ambiguity by taking ResId into account
                  dtrFound = tdlFtDbase.Tables("Result").Select("File = '" & strFile & "' AND forestId=" & intForestId & _
                                                                " AND eTreeId=" & intEtreeId & " AND ResId=" & intResId)
                  ' If this has failed, try two more variants
                  If (dtrFound.Length = 0) Then
                    ' Look for fuller path
                    dtrFound = tdlFtDbase.Tables("Result").Select("File LIKE '*\" & strFile & "' AND forestId=" & intForestId & _
                                                                  " AND eTreeId=" & intEtreeId & " AND ResId=" & intResId)
                  End If
                  ' Message depends on the outcome
                  If (dtrFound.Length = 0) Then
                    ' No solution
                    Logging("WARNING: failed to recover ambiguous location at: " & strFile & ":" & intForestId & ":" & intEtreeId)
                    intRecNo += 1
                  Else
                    ' Ambiguous - stop
                    Logging("WARNING: recovered ambiguous location at: " & strFile & ":" & intForestId & ":" & intEtreeId)
                    intRecOk += 1
                  End If
                End If
            End Select
            Select Case dtrFound.Length
              Case 0
                ' Not found, so skip
                Logging("WARNING: could not find entry at: " & strFile & ":" & intForestId & ":" & intEtreeId)
                intRecNo += 1
              Case 1
                ' Perfectly okay, process
                If (bUseStatus) Then
                  AddXmlAttribute(pdxCrpDbase, ndxList(intI), "Status", dtrFound(0).Item("Status").ToString)
                  ' Update [Status] in results (for direct visibility in the DGV)
                  dtrRes = tdlResults.Tables("Result").Select("ResId=" & .Attributes("ResId").Value)
                  If (dtrRes.Length = 1) Then
                    ' Only adapt if there is a difference
                    If (dtrRes(0).Item("Status") <> dtrFound(0).Item("Status").ToString) Then
                      dtrRes(0).Item("Status") = dtrFound(0).Item("Status").ToString
                    End If
                  Else
                    Logging("WARNING: failed to identify [@Status] at: " & strFile & ":" & intForestId & ":" & intEtreeId)
                  End If

                End If
                If (bUseNotes) Then AddXmlAttribute(pdxCrpDbase, ndxList(intI), "Notes", dtrFound(0).Item("Notes").ToString)
                If (bUsePde) Then
                  ndxFeat = .SelectSingleNode("./child::Pde[0]")
                  If (ndxFeat Is Nothing) Then ndxFeat = AddXmlChild(ndxList(intI), "Pde")
                  ndxFeat.InnerText = dtrFound(0).Item("Pde").ToString
                End If
                ' Find the feature 
                dtrFeat = tdlFtDbase.Tables("Feature").Select("Parent.ResId=" & dtrFound(0).Item("ResId").ToString & _
                                                              " AND Name = '" & strFeatDbase & "'")
                If (dtrFeat.Length = 1) Then
                  ndxFeat = .SelectSingleNode("./child::Feature[@Name = '" & strFeatDbase & "']")
                  If (ndxFeat Is Nothing) Then ndxFeat = AddXmlChild(ndxList(intI), "Feature", "Name", strFeatDbase, "attribute",
                                                                                               "Value", "", "attribute")
                  ndxFeat.Attributes("Value").Value = dtrFeat(0).Item("Value").ToString
                End If
                ' Keep track of correctly imported features
                intCount += 1
              Case Else
                ' Ambiguous - give error message
                Logging("WARNING: unable to process  ambiguous location at: " & strFile & ":" & intForestId & ":" & intEtreeId, True)
            End Select
          End If
        End With
      Next intI
      ' Give overall report
      If (intRecNo > 0) OrElse (intRecOk > 0) Then
        Logging("============================= " & vbCrLf & _
                "Recovery report: " & intRecOk & " entries were recovered; " & intRecNo & " entries failed." & vbCrLf & _
                "============================= ")
      End If
      ' Give report of how many succeeeded
      If (intCount = 0) Then
        Logging("Importing features: nothing has been imported")
      Else
        Logging("Importing features: " & intCount & " entries have been imported successfully")
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modCorpus/DoCorpusFeaturesFromFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoCorpusFeaturesFromCsv
  ' Goal:   Copy features from a csv file to the current database
  ' History:
  ' 31-12-2014  ERK Created - derived from [DoCorpusFeaturesFromFile]
  ' ------------------------------------------------------------------------------------------------------------
  Public Function DoCorpusFeaturesFromCsv(ByVal strImpFtFile As String, ByVal strCurrentFilter As String, ByVal strFeatDbase As String, _
                      ByVal strMethod As String, ByVal strFeatValues As String, ByVal strFeatBlackList As String, ByVal bUseStatus As Boolean, _
                      ByVal bUseNotes As Boolean, ByVal bUsePde As Boolean) As Boolean
    Dim strTextId As String     ' Current file we are opening
    Dim strLastFile As String   ' Last file we were working on
    Dim arFile() As String      ' Array of files
    Dim strFeatValue As String  ' Value for the feature
    Dim strExpFeatFile As String  ' File name we export the features to
    Dim strStatus As String     ' Status
    Dim strNotes As String      ' Notes
    Dim strPde As String        ' PDE (back translation)
    Dim intForestId As Integer  ' Forest id
    Dim intEtreeId As Integer   ' Etree Id
    Dim intResId As Integer     ' Recovered original resid
    Dim ndxFeat As XmlNode      ' Node containing feature
    Dim ndxForest As XmlNode    ' Forest node
    Dim ndxEtree As XmlNode     ' Constituent
    Dim ndxList As XmlNodeList  ' List of nodes in the corpus database we are visiting
    Dim ndxSpec As XmlNodeList  ' List for [bSpecial]
    Dim tblCsv As DataTable     ' Datatable representing the CSV file
    Dim dtrThis As DataRow      ' One datarow
    Dim dtrFeat() As DataRow    ' Datarow for this feature
    Dim dtrFound() As DataRow   ' Walking through the dataset
    Dim dtrRes() As DataRow     ' Location in tdlResults
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter
    Dim intK As Integer         ' Counter
    Dim intPtc As Integer       ' Percentage
    Dim intCount As Integer     ' Number of changes
    Dim intRecOk As Integer     ' Number of well recovered entries
    Dim intRecNo As Integer     ' Number of failed recovers
    Dim bSpecial As Boolean = False  ' SPecial treatment -- AmbiLemma problem

    Try
      ' Validate
      If (Not IO.File.Exists(strImpFtFile)) Then Return False
      ' Try read the csv file as a datatable
      If (Not CsvToDataTable(strImpFtFile, tblCsv, "forestId", "System.Int32", "eTreeId", "System.Int32", _
                             "TextId", "System.String")) Then Return False
      ' Start addressing the corpus research database...
      If (strCurrentFilter = "") Then
        ndxList = pdxCrpDbase.SelectNodes("./descendant::Result")
      Else
        ndxList = pdxCrpDbase.SelectNodes(strCurrentFilter, conTb)
      End If
      SetXmlDocument(pdxCrpDbase)
      strLastFile = "" : intCount = 0 : intRecOk = 0 : intRecNo = 0
      ' Visit them all
      For intI = 0 To ndxList.Count - 1
        ' Access the information in this node
        With ndxList(intI)
          ' Where are we?
          intPtc = (intI + 1) * 100 \ ndxList.Count
          Status("FeaturesToFile " & intPtc & "%", intPtc)
          CorefProgress(.Attributes("TextId").Value)
          ' Get forest id etc
          intForestId = .Attributes("forestId").Value
          intEtreeId = .Attributes("eTreeId").Value
          ' Get the correct @TextId
          strTextId = .Attributes("TextId").Value
          ' Recover the original ResID as stored in OviewId
          intResId = .Attributes("ResId").Value
          ' ============ DEBUG ==========
          '  If (strTextId Like "cobede*" AndAlso intForestId = 9) Then Stop
          ' =============================
          ' Get the feature value
          ndxFeat = .SelectSingleNode("./descendant::Feature[@Name='" & strFeatDbase & "']")
          If (ndxFeat IsNot Nothing) Then
            ' Action depends on the method for searching we use
            Select Case strMethod
              Case "ResId"
                ' Each entry is uniquely determined by its ResId
                dtrFound = tblCsv.Select("ResId=" & intResId)
              Case Else
                ' Use text/forest/etree
                ' Look for the corresponding entry in the file we have read
                dtrFound = tblCsv.Select("TextId = '" & strTextId & "' AND forestId=" & intForestId & _
                                                              " AND eTreeId=" & intEtreeId)
                ' Make sure we get results
                If (dtrFound.Length = 0) Then
                  ' Look for fuller path
                  dtrFound = tblCsv.Select("TextId LIKE '*\" & strTextId & "' AND forestId=" & intForestId & _
                                                                " AND eTreeId=" & intEtreeId)
                End If
                ' Be on the watch for ambiguity (too much results)
                If (dtrFound.Length > 1) Then
                  ' Try reduce ambiguity by taking ResId into account
                  dtrFound = tblCsv.Select("TextId = '" & strTextId & "' AND forestId=" & intForestId & _
                                                                " AND eTreeId=" & intEtreeId & " AND ResId=" & intResId)
                  ' If this has failed, try two more variants
                  If (dtrFound.Length = 0) Then
                    ' Look for fuller path
                    dtrFound = tblCsv.Select("TextId LIKE '*\" & strTextId & "' AND forestId=" & intForestId & _
                                                                  " AND eTreeId=" & intEtreeId & " AND ResId=" & intResId)
                  End If
                  ' Message depends on the outcome
                  If (dtrFound.Length = 0) Then
                    ' No solution
                    Logging("WARNING: failed to recover ambiguous location at: " & strTextId & ":" & intForestId & ":" & intEtreeId)
                    intRecNo += 1
                  Else
                    ' Ambiguous - stop
                    Logging("WARNING: recovered ambiguous location at: " & strTextId & ":" & intForestId & ":" & intEtreeId)
                    intRecOk += 1
                  End If
                End If
            End Select
            Select Case dtrFound.Length
              Case 0
                ' Not found, so skip
                Logging("WARNING: could not find entry at: " & strTextId & ":" & intForestId & ":" & intEtreeId)
                intRecNo += 1
              Case 1
                ' Perfectly okay, process
                If (bUseStatus) Then
                  AddXmlAttribute(pdxCrpDbase, ndxList(intI), "Status", dtrFound(0).Item("Status").ToString)
                  ' Update [Status] in results (for direct visibility in the DGV)
                  dtrRes = tdlResults.Tables("Result").Select("ResId=" & .Attributes("ResId").Value)
                  If (dtrRes.Length = 1) Then
                    ' Only adapt if there is a difference
                    If (dtrRes(0).Item("Status") <> dtrFound(0).Item("Status").ToString) Then
                      dtrRes(0).Item("Status") = dtrFound(0).Item("Status").ToString
                    End If
                  Else
                    Logging("WARNING: failed to identify [@Status] at: " & strTextId & ":" & intForestId & ":" & intEtreeId)
                  End If

                End If
                If (bUseNotes) Then AddXmlAttribute(pdxCrpDbase, ndxList(intI), "Notes", dtrFound(0).Item("Notes").ToString)
                If (bUsePde) Then
                  ndxFeat = .SelectSingleNode("./child::Pde[0]")
                  If (ndxFeat Is Nothing) Then ndxFeat = AddXmlChild(ndxList(intI), "Pde")
                  ndxFeat.InnerText = dtrFound(0).Item("Pde").ToString
                End If
                ' Check if this feature is available in this row
                If (dtrFound(0).Item(strFeatDbase) IsNot Nothing) Then
                  ' Get the new feature value
                  strFeatValue = dtrFound(0).Item(strFeatDbase).ToString
                  ' Get the target location 
                  ndxFeat = .SelectSingleNode("./child::Feature[@Name = '" & strFeatDbase & "']")
                  ' If non-existent, create it
                  If (ndxFeat Is Nothing) Then ndxFeat = AddXmlChild(ndxList(intI), "Feature", "Name", strFeatDbase, "attribute",
                                                                                               "Value", "", "attribute")
                  ' Give target location the correct value
                  ndxFeat.Attributes("Value").Value = strFeatValue
                Else
                  Stop
                End If
              Case Else
                ' Ambiguous - give error message
                Logging("WARNING: unable to process  ambiguous location at: " & strTextId & ":" & intForestId & ":" & intEtreeId, True)
            End Select
          End If
        End With
      Next intI
      ' Give overall report
      If (intRecNo > 0) OrElse (intRecOk > 0) Then
        Logging("============================= " & vbCrLf & _
                "Recovery report: " & intRecOk & " entries were recovered; " & intRecNo & " entries failed." & vbCrLf & _
                "============================= ")
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modCorpus/DoCorpusFeaturesFromCsv error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   CsvToDataTable
  ' Goal:   Read a CSV file as a datatable
  ' History:
  ' 31-12-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function CsvToDataTable(ByVal strFile As String, ByRef tblBack As DataTable, _
                                 ByVal ParamArray arColTypes As String()) As Boolean
    Dim SR As IO.StreamReader = Nothing   ' Open a stream reader for the file
    Dim line As String = ""               ' First line
    Dim strArray As String()              ' Array to hold the first line's elements
    Dim row As DataRow                    ' General purpose datarow
    Dim intI As Integer                   ' Counter
    Dim bFound As Boolean                 ' Found match

    Try
      ' Validate
      If (Not IO.File.Exists(strFile)) Then Return False
      ' Open a streamreader for the file
      SR = New IO.StreamReader(strFile)
      ' Create a new datatable
      tblBack = New DataTable()
      ' Read first line and split into array
      line = SR.ReadLine()
      strArray = line.Split(vbTab)
      ' Walk all column headings in the array
      For Each s As String In strArray
        ' Check if datatype is needed for this one
        bFound = False
        For intI = 0 To arColTypes.Count - 1 Step 2
          If (arColTypes(intI) = s) Then bFound = True : Exit For
        Next intI
        If (bFound) Then
          tblBack.Columns.Add(New DataColumn(s, System.Type.GetType(arColTypes(intI + 1))))
        Else
          tblBack.Columns.Add(New DataColumn(s))
        End If
      Next s
      ' Walk all the lines of the csv file
      Do
        line = Trim(SR.ReadLine)
        ' Check if it is not empty
        If (line <> "") Then
          ' Create a new datarow
          row = tblBack.NewRow()
          ' Copy all elements to the new row
          row.ItemArray = line.Split(vbTab)
          ' Add the row to the table
          tblBack.Rows.Add(row)
        End If
      Loop While Not (line = String.Empty)
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modCorpus/CsvToDataTable error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoCorpusFeaturesToFile
  ' Goal:   Copy features from database to a file
  ' History:
  ' 10-07-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function DoCorpusFeaturesToFile(ByVal strTextDir As String, ByVal strCurrentFilter As String, ByVal strFeatDbase As String, _
                      ByVal strFeatValues As String, ByVal strFeatBlackList As String, ByVal bUseStatus As Boolean, _
                      ByVal bUseNotes As Boolean, ByVal bUsePde As Boolean, ByVal strExpDir As String) As Boolean
    Dim strFile As String       ' Current file we are opening
    Dim strLastFile As String   ' Last file we were working on
    Dim arFile() As String      ' Array of files
    Dim strFeatValue As String  ' Value for the feature
    Dim strExpFeatFile As String  ' File name we export the features to
    Dim strStatus As String     ' Status
    Dim strNotes As String      ' Notes
    Dim strPde As String        ' PDE (back translation)
    Dim intForestId As Integer  ' Forest id
    Dim intEtreeId As Integer   ' Etree Id
    Dim ndxFeat As XmlNode      ' Node containing feature
    Dim ndxForest As XmlNode    ' Forest node
    Dim ndxEtree As XmlNode     ' Constituent
    Dim ndxList As XmlNodeList  ' List of nodes in the corpus database we are visiting
    Dim ndxSpec As XmlNodeList  ' List for [bSpecial]
    Dim dtrThis As DataRow      ' One datarow
    Dim dtrFeat As DataRow      ' Datarow for this feature
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter
    Dim intK As Integer         ' Counter
    Dim intPtc As Integer       ' Percentage
    Dim intCount As Integer     ' Number of changes
    Dim bSpecial As Boolean = False  ' SPecial treatment -- AmbiLemma problem
    Dim tdlFtDbase As DataSet = Nothing   ' Dataset to hold the features

    Try
      ' Validate
      If (Not IO.Directory.Exists(strTextDir)) Then Return False
      If (pdxCrpDbase Is Nothing) Then Logging("DoCorpusFeaturesToFile error: the corpus database seems to be invisible", True) : Return False
      ' Determine feature output file name
      strExpFeatFile = IO.Path.GetFileNameWithoutExtension(strCurrentDbFile)
      If (Right(strExpFeatFile, 4) = "-tmp") Then strExpFeatFile = Left(strExpFeatFile, strExpFeatFile.Length - 4)
      strExpFeatFile = strExpDir & "\" & strExpFeatFile & "-" & strFeatDbase & ".xml"
      ' Create a database to hold the feature values
      If (Not CreateDataSet("CrpResult.xsd", tdlFtDbase)) Then Return False
      ' Now get on with it...
      If (strCurrentFilter = "") Then
        ndxList = pdxCrpDbase.SelectNodes("./descendant::Result")
      Else
        ndxList = pdxCrpDbase.SelectNodes(strCurrentFilter, conTb)
      End If
      strLastFile = "" : intCount = 0
      ' Visit them all
      For intI = 0 To ndxList.Count - 1
        ' Access the information in this node
        With ndxList(intI)
          ' Where are we?
          intPtc = (intI + 1) * 100 \ ndxList.Count
          Status("FeaturesToFile " & intPtc & "%", intPtc)
          CorefProgress(.Attributes("TextId").Value)
          ' Get forest id etc
          If (.Attributes("forestId") Is Nothing) Then Logging("DoCorpusFeaturesToFile: node misses @forestId") : Return False
          intForestId = .Attributes("forestId").Value
          If (.Attributes("eTreeId") Is Nothing) Then Logging("DoCorpusFeaturesToFile: node misses @eTreeId") : Return False
          intEtreeId = .Attributes("eTreeId").Value
          If (.Attributes("File") Is Nothing) Then Logging("DoCorpusFeaturesToFile: node misses @File") : Return False
          strFile = .Attributes("File").Value
          ' Get the feature value
          ndxFeat = .SelectSingleNode("./descendant::Feature[@Name='" & strFeatDbase & "']")
          If (ndxFeat IsNot Nothing) Then
            ' Get the value for the feature
            strFeatValue = ndxFeat.Attributes("Value").Value
            ' Is this feature value acceptable?
            If ((strFeatValues = "") OrElse (DoLike(strFeatValue, strFeatValues))) AndAlso _
               ((strFeatBlackList = "") OrElse (Not DoLike(strFeatValue, strFeatBlackList))) Then
              strStatus = GetAttrValue(ndxList(intI), "Status")
              strNotes = GetAttrValue(ndxList(intI), "Notes")
              ' Does not work properly:
              '    strStatus = .Attributes("Status").Value : strNotes = .Attributes("Notes").Value
              If (.SelectSingleNode("./child::Pde") Is Nothing) Then
                strPde = ""
              Else
                strPde = .SelectSingleNode("./child::Pde").InnerText
              End If
              ' Add a row in the database
              dtrThis = AddOneDataRow(tdlFtDbase, "Result", "ResId", "OviewList")
              With dtrThis
                .Item("forestId") = intForestId : .Item("eTreeId") = intEtreeId : .Item("File") = strFile
                If (bUseStatus) Then .Item("Status") = strStatus
                If (bUseNotes) Then .Item("Notes") = strNotes
                If (bUsePde) Then .Item("Pde") = strPde
                ' Also in all cases export the original ResId number as feature
                .Item("OviewId") = GetAttrValue(ndxList(intI), "ResId")
              End With
              ' Process feature name and value
              dtrFeat = AddOneDataRowWithParent(tdlFtDbase, "Feature", "", dtrThis)
              With dtrFeat
                .Item("Name") = strFeatDbase : .Item("Value") = strFeatValue
              End With
              intCount += 1
            Else
              ' There is no need to process this feature, because it does not have the correct value
              Logging("Skipped feature value [" & strFeatValue & "] in " & .Attributes("TextId").Value & ":" & intForestId & "." & intEtreeId, False)
              ' ====== DEBUG ============
              '  Debug.Print("Unacceptable feature value: [" & strFeatValue & "]")
              '  If (strFeatValue = "unacc") OrElse (strFeatValue = "other") Then Stop
              '              Stop
              ' =========================
            End If
          End If
        End With
      Next intI
      ' Save the feature file
      Status("Saving features to " & strExpFeatFile)
      tdlFtDbase.WriteXml(strExpFeatFile)
      ' Reset progress indicator
      frmMain.tsProgress.Text = ""
      Logging("Changes: " & intCount)
      Status("Ready")
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modCorpus/DoCorpusFeaturesToFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoCorpusFeaturesToTexts
  ' Goal:   Copy the annotation from one record to similar records
  ' History:
  ' 02-05-2014  ERK Created
  ' 09-06-2014  ERK Added bUseAttribute
  ' ------------------------------------------------------------------------------------------------------------
  Public Function DoCorpusFeaturesToTexts(ByVal strTextDir As String, ByVal strCurrentFilter As String, ByVal strFeatDbase As String, _
                      ByVal strFeatCat As String, ByVal strFeatName As String, ByVal strFeatValues As String, _
                      ByVal strFeatBlackList As String, ByVal bUseAttribute As Boolean) As Boolean
    Dim strFile As String       ' Current file we are opening
    Dim strLastFile As String   ' Last file we were working on
    Dim arFile() As String      ' Array of files
    Dim strFeatValue As String  ' Value for the feature
    Dim intForestId As Integer  ' Forest id
    Dim intEtreeId As Integer   ' Etree Id
    Dim ndxFeat As XmlNode      ' Node containing feature
    Dim ndxForest As XmlNode    ' Forest node
    Dim ndxEtree As XmlNode     ' Constituent
    Dim ndxList As XmlNodeList  ' List of nodes in the corpus database we are visiting
    Dim ndxSpec As XmlNodeList  ' List for [bSpecial]
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter
    Dim intK As Integer         ' Counter
    Dim intPtc As Integer       ' Percentage
    Dim intCount As Integer     ' Number of changes
    Dim bSpecial As Boolean = False  ' SPecial treatment -- AmbiLemma problem

    Try
      ' Validate
      If (Not IO.Directory.Exists(strTextDir)) Then Return False
      If (strFeatCat = "") OrElse (strFeatName = "") Then Return False
      ' Now get on with it...
      If (strCurrentFilter = "") Then
        ndxList = pdxCrpDbase.SelectNodes("./descendant::Result")
      Else
        ndxList = pdxCrpDbase.SelectNodes(strCurrentFilter, conTb)
      End If
      strLastFile = "" : intCount = 0
      ' Visit them all
      For intI = 0 To ndxList.Count - 1
        ' Access the information in this node
        With ndxList(intI)
          ' Where are we?
          intPtc = (intI + 1) * 100 \ ndxList.Count
          Status("FeaturesToText " & intPtc & "%", intPtc)
          CorefProgress(.Attributes("TextId").Value)
          ' Get forest id etc
          intForestId = .Attributes("forestId").Value
          intEtreeId = .Attributes("eTreeId").Value
          '' ============= Ad hoc emendment for file that went wrong =====
          'If (.Attributes("TextId").Value.ToLower Like "tobz2*") Then
          '  intEtreeId -= 461
          'End If
          '' =============================================================
          ' Get the feature value
          ndxFeat = .SelectSingleNode("./descendant::Feature[@Name='" & strFeatDbase & "']")
          If (ndxFeat IsNot Nothing) Then
            ' Get the value for the feature
            strFeatValue = ndxFeat.Attributes("Value").Value
            ' Is this feature value acceptable?
            If ((strFeatValues = "") OrElse (DoLike(strFeatValue, strFeatValues))) AndAlso _
               ((strFeatBlackList = "") OrElse (Not DoLike(strFeatValue, strFeatBlackList))) Then
              ' Get the file name and the feature value we need to have
              arFile = IO.Directory.GetFiles(strTextDir, .Attributes("File").Value, IO.SearchOption.AllDirectories)
              If (arFile.Length > 0) Then
                ' We only get the first file
                strFile = arFile(0)
                If (strFile <> strLastFile) Then
                  If (strLastFile <> "") AndAlso (strCurrentFile <> "") AndAlso (pdxCurrentFile IsNot Nothing) Then
                    ' Save the last file
                    Logging("Saving changes to " & strCurrentFile, False)
                    pdxCurrentFile.Save(strCurrentFile)
                  End If
                  ' We need to read this file
                  strCurrentFile = strFile
                  If (Not ReadXmlDoc(strCurrentFile, pdxCurrentFile)) Then Logging("Could not load file: " & strCurrentFile) : Return False
                  ' Keep track of the last file
                  strLastFile = strFile
                  ' Set the forest node
                  If (Not GetFirstForest(pdxCurrentFile, ndxForest)) Then Return False
                End If
                ' Look for the correct forestid/etreeid
                If (ndxForest.Attributes("forestId").Value <> intForestId) Then
                  ndxForest = ndxForest.SelectSingleNode("./following-sibling::forest[@forestId=" & intForestId & "]")
                  If (ndxForest Is Nothing) Then
                    Logging("Could not find " & .Attributes("TextId").Value & ": forest " & intForestId)
                    Return False
                  End If
                End If
                ' Get to the constituent
                ndxEtree = ndxForest.SelectSingleNode("./descendant::eTree[@Id=" & intEtreeId & "]")
                If (bSpecial) Then
                  Dim ndxFtLab As XmlNode = Nothing
                  Dim ndxFtVb As XmlNode = Nothing
                  Dim strVb As String = ""
                  Dim bFound As Boolean = False
                  Dim strDbVb As String = ""  ' Verb in the database

                  ndxFtLab = .SelectSingleNode("./descendant::Feature[@Name='VbLabel']")
                  If (ndxFtLab Is Nothing) Then Stop
                  ndxSpec = ndxEtree.SelectNodes("./child::eTree[@Label = '" & _
                                                 ndxFtLab.Attributes("Value").Value & "']")
                  If (ndxSpec.Count = 0) Then
                    Stop
                  Else
                    ' Get the verb form
                    ndxFtVb = .SelectSingleNode("./descendant::Feature[@Name='Verb']")
                    If (ndxFtVb Is Nothing) Then Stop
                    strVb = ndxFtVb.Attributes("Value").Value
                    ' Find the correct child
                    For intK = 0 To ndxSpec.Count - 1
                      strDbVb = NodeText(ndxSpec(intK))
                      ' strDbVb = ndxSpec(intK).Attributes("Text").Value
                      ' Convert as if we are doing RuConv(clean)
                      strDbVb = VernToEnglish(LCase(strDbVb.Replace(";", " ")))
                      strDbVb = strDbVb.Replace("$", "")
                      strDbVb = strDbVb.Replace("'", "")
                      ' Compare the two
                      If (strDbVb = strVb) Then bFound = True : Exit For
                    Next intK
                    If (bFound) Then
                      ' We found the correct one
                      ndxEtree = ndxSpec(intK)
                    Else
                      Stop
                    End If
                  End If
                End If
                ' Validate
                If (ndxEtree Is Nothing) Then
                  ' Problem
                  Logging("Could not find " & .Attributes("TextId").Value & ": forest " & intForestId & " / Etree=" & intEtreeId)
                  Return False
                End If
                ' Change feature or attribute?
                If (bUseAttribute) Then
                  If (ndxEtree.Attributes(strFeatName).Value <> strFeatValue) Then
                    ' Adapt the feature value
                    ndxEtree.Attributes(strFeatName).Value = strFeatValue
                    intCount += 1
                  End If
                Else
                  If (GetFeature(ndxEtree, strFeatCat, strFeatName) <> strFeatValue) Then
                    ' Adapt the feature value
                    If (Not AddFeature(pdxCurrentFile, ndxEtree, strFeatCat, strFeatName, strFeatValue)) Then Logging("Could not add feature value") : Return False
                    intCount += 1
                  End If
                End If
                ' Refresh the changes count
                frmMain.SectionText("Changes=" & intCount)
              Else
                ' What to do if we miss a file?
                Logging("Could not find file: " & .Attributes("File").Value)
              End If
            Else
              ' There is no need to process this feature, because it does not have the correct value
              Logging("Skipped feature value [" & strFeatValue & "] in " & .Attributes("TextId").Value & ":" & intForestId & "." & intEtreeId, False)
              ' ====== DEBUG ============
              '  Debug.Print("Unacceptable feature value: [" & strFeatValue & "]")
              '  If (strFeatValue = "unacc") OrElse (strFeatValue = "other") Then Stop
              '              Stop
              ' =========================
            End If
          End If
        End With
      Next intI
      ' The last file should be saved!!!
      If (strCurrentFile <> "") AndAlso (pdxCurrentFile IsNot Nothing) Then
        ' Save the last file
        Status("Saving changes to " & strCurrentFile)
        pdxCurrentFile.Save(strCurrentFile)
      End If
      ' Reset progress indicator
      frmMain.tsProgress.Text = ""
      Logging("Changes: " & intCount)
      Status("Ready")
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modCorpus/DoCorpusFeaturesToTexts error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   FeatureExists
  ' Goal:   Check if the feature with the specified name exists already in the database
  ' History:
  ' 25-08-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function FeatureExists(ByVal strFeatName As String) As Boolean
    Dim ndxList As XmlNodeList  ' List of features
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (pdxCrpDbase Is Nothing) Then Return False
      ' Get a list of the features by looking at the first result
      ndxList = pdxCrpDbase.SelectNodes("./descendant::Result[1]/descendant::Feature")
      For intI = 0 To ndxList.Count - 1
        ' Compare this feature's name
        If (LCase(strFeatName) = LCase(ndxList(intI).Attributes("Name").Value)) Then Return True
      Next intI
      ' Feature has not been found, so it does not yet exist
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modCorpus/FeatureExists error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   FeatureAdd
  ' Goal:   Add one feature to all the records in the database
  ' History:
  ' 25-08-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function FeatureAdd(ByVal strFeatName, ByVal strFeatValue) As Boolean
    Dim ndxList As XmlNodeList  ' List of features
    Dim intI As Integer         ' Counter
    Dim intPtc As Integer       ' Percentage

    Try
      ' Validate
      If (pdxCrpDbase Is Nothing) Then Return False
      ' Get a list of the records
      ndxList = pdxCrpDbase.SelectNodes("./descendant::Result")
      SetXmlDocument(pdxCrpDbase)
      ' Walk the records
      For intI = 0 To ndxList.Count - 1
        ' Where are we?
        intPtc = (intI + 1) * 100 \ ndxList.Count
        Status("Adding feature [" & strFeatName & "] " & intPtc & "%", intPtc)
        ' Add the feature to the current record
        If (AddXmlChild(ndxList(intI), "Feature", "Name", strFeatName, "attribute", _
                         "Value", strFeatValue, "attribute") Is Nothing) Then Return False
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modCorpus/FeatureAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
End Module
