Imports System.Xml
Imports System.Text.RegularExpressions
Module modMetaData
  Private loc_arMeta() As StringColl      ' Array of string collections
  Private loc_strMetaFile As String = ""  ' Name of meta file used
  Private loc_dlgFld As New FolderBrowserDialog ' The dialog
  Private loc_strFileExt As String = ".psdx"
  ' ------------------------------------------------------------------------------------
  ' Name:   InitMetaEd
  ' Goal:   Initialise editor(s) connected with metadata
  ' History:
  ' 16-10-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function InitMetaEd() As Boolean
    Try
      ' Do away with possible previous handlers
      DgvClear(objMetaEd)   ' Metadata handler for file headers
      ' Initialise the metadata editor for file handlers
      frmMain.InitTeiMetaEditor()
      ' Return success
      Return True
    Catch ex As Exception
      ' Show user that there is an error
      HandleErr("frmMain/InitMetaEd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   TeiMetaDataInit
  ' Goal:   Initialise the tei metadata stuff
  ' History:
  ' 11-08-2014  ERK Created
  ' 19-06-2015  ERK Lak metadata initialisations
  ' ------------------------------------------------------------------------------------------------------------
  Public Function TeiMetaDataInit() As Boolean
    Dim ndxList As XmlNodeList        ' List of nodes
    Dim intI As Integer               ' Counter
    Dim dtrThis As DataRow            ' one datarow

    Try
      ' Validate
      If (tdlMeta Is Nothing) Then Return False
      ' Initialise metadata editing
      InitMetaEd()
      ' Clear any previous information
      ClearTable(tdlMeta.Tables("meta"))
      ' Copy all meta information from psdx file to [tdlMeta]
      ndxList = pdxCurrentFile.SelectNodes("./descendant::metadata[1]/child::meta")
      For intI = 0 To ndxList.Count - 1
        dtrThis = AddOneDataRow(tdlMeta, "meta", "metaId", "metadata")
        dtrThis.Item("name") = ndxList(intI).Attributes("id").Value
        dtrThis.Item("value") = ndxList(intI).InnerText
      Next intI
      If (Not TeiMetaDataSonar()) Then Return False
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/TeiMetaDataInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return False
    End Try
  End Function
  Public Function TeiMetaDataSonar() As Boolean
    Dim intI As Integer               ' Counter
    Dim arSonarField() As String = {"Age", "AuthorNameOrPseudonym",
      "CollectionName", "Country", "LicenseCode", "OriginalLanguage", "PublicationName",
      "PublicationPlace", "Published", "Publisher", "Sex", "SourceName", "TextDescription",
      "TextKeyword", "TextType", "Town", "Translated", "TranslatorName"}
    Dim arSonarControl() As Control = {frmMain.tbSonarAge, frmMain.tbSonarAuthorNameOrPseudonym,
      frmMain.tbSonarCollectionName, frmMain.tbSonarCountry,
      frmMain.tbSonarLicenseCode, frmMain.tbSonarOriginalLanguage, frmMain.tbSonarPublicationName,
      frmMain.tbSonarPublicationPlace, frmMain.tbSonarPublished, frmMain.tbSonarPublisher,
      frmMain.tbSonarSex, frmMain.tbSonarSourceName, frmMain.tbSonarTextDescription,
      frmMain.tbSonarTextKeyword, frmMain.tbSonarTextType, frmMain.tbSonarTown,
      frmMain.tbSonarTranslated, frmMain.tbSonarTranslatorName}

    Try

      ' Show the metadata fields that need showing
      For intI = 0 To arSonarField.Length - 1
        ' Show this field on the indicated control
        arSonarControl(intI).Text = TeiMetaDataGet(arSonarField(intI))
      Next intI
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/TeiMetaDataInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   TeiLakInit
  ' Goal:   Initialise the tei metadata stuff for LAK (lbe)
  ' History:
  ' 19-06-2015  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function TeiLakInit() As Boolean
    Dim intI As Integer               ' Counter
    Dim bChanges As Boolean = False
    Dim arLakFieldName() As String = {"CollectionName", "Published", "Publisher", "Translated"}
    Dim arLakFieldValue() As String = {"PCMLBE", "yes", "Novolunie", "no"}

    Try

      ' Lak (lbe) automatic metadata apatations for Sonar compatibility
      If (TeiGetLanguage("ident").ToLower = "lbe-latn") Then
        ' Check for some obvious values
        For intI = 0 To arLakFieldName.Length - 1
          If (Not TeiMetaDataExists(arLakFieldName(intI))) Then
            TeiMetaDataAddOne(arLakFieldName(intI), arLakFieldValue(intI), True)
            bChanges = True
          End If
        Next intI
        ' Possibly copy the genre
        TeiMetaDataAddOne("TextType", TeiGetGenre(), True)
      End If
      If (bChanges) Then
        ' Set the flag
        frmMain.SetDirty(True)
        ' Show again
        If (Not TeiMetaDataSonar()) Then Return False
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/TeiLakInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   TeiCheInit
  ' Goal:   Initialise the tei metadata stuff for CHECHEN
  ' History:
  ' 19-06-2015  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function TeiCheInit() As Boolean
    Dim intI As Integer               ' Counter
    Dim bChanges As Boolean = False
    Dim arLakFieldName() As String = {"CollectionName", "Published", "Translated"}
    Dim arLakFieldValue() As String = {"NPCMC", "yes", "no"}

    Try

      ' Lak (lbe) automatic metadata apatations for Sonar compatibility
      If (TeiGetLanguage("ident").ToLower Like "che*") Then
        ' Set the language correctly
        ' TODO
        ' Check for some obvious values
        For intI = 0 To arLakFieldName.Length - 1
          If (Not TeiMetaDataExists(arLakFieldName(intI))) Then
            TeiMetaDataAddOne(arLakFieldName(intI), arLakFieldValue(intI), True)
            bChanges = True
          End If
        Next intI
        ' Possibly copy the genre
        TeiMetaDataAddOne("TextType", TeiGetGenre(), True)
      End If
      If (bChanges) Then
        ' Set the flag
        frmMain.SetDirty(True)
        ' Show again
        If (Not TeiMetaDataSonar()) Then Return False
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/TeiCheInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   TeiGetCorpusTable
  ' Goal:   Retrieve all the meta data from the .psdx files in the indiected directory
  ' History:
  ' 26-06-2015  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function TeiGetCorpusTable(ByRef strCsv As String, ByRef strHtml As String) As Boolean
    Dim strDir As String = ""   ' Directory where to look
    Dim strFile As String       ' One file
    Dim strPath As String       ' Current path
    Dim strCat As String        ' One category
    Dim strRow As String        ' One row
    Dim strFull As String       ' Second row in table
    Dim arFile() As String      ' Files in directory
    Dim arCat() As String = {"fileDesc", "profileDesc"}
    Dim pdxFile As XmlDocument  ' One header
    Dim ndxList As XmlNodeList  ' List of nodes
    Dim attrThis As XmlAttribute
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter
    Dim intK As Integer         ' Counter
    Dim intL As Integer         ' Counter
    ' One dicFile per .psdx file in the corpus
    Dim dicFile As Dictionary(Of String, String)
    ' All dicFiles together into the dicCorpus
    Dim dicCorpus As New Dictionary(Of String, Dictionary(Of String, String))
    ' Each unique category into the colCats
    Dim colCats As New Collection   ' Full category names
    Dim colShort As New Collection  ' SHort category names
    Dim sbOut As New StringColl     ' CSV output
    Dim sbHtml As New StringColl    ' HTML output

    Try
      ' Initialise directory
      If (strCurrentFile <> "") Then
        strDir = IO.Path.GetDirectoryName(strCurrentFile)
      Else
        ' Try most recent
        strFile = GetTableSetting(tdlSettings, "Recent1")
        If (strFile = "") Then
          strDir = strWorkDir
        Else
          strDir = IO.Path.GetDirectoryName(strFile)
        End If
      End If
      ' Initialize
      dicCorpus.Clear()
      ' Ask for directory
      If (Not GetDirName(loc_dlgFld, strDir, "Directory containing your corpus", strDir)) Then Return False
      ' Look for files in this directory
      arFile = IO.Directory.GetFiles(strDir, "*" & loc_strFileExt)
      For intI = 0 To arFile.Length - 1
        ' Get this file
        strFile = arFile(intI)
        ' Initialize
        dicFile = New Dictionary(Of String, String)
        ' Read the metadata from this file
        If (Not ReadPsdxHeader(strFile, pdxFile)) Then Return False
        If (pdxFile Is Nothing) Then Return False
        ' Get the meta-data from the different sections: <metadata>, <fileDesc>, <profileDesc>
        For intJ = 0 To arCat.Length - 1
          ndxList = pdxFile.SelectNodes("./descendant::" & arCat(intJ) & "/descendant::*")
          For intK = 0 To ndxList.Count - 1
            ' Review this item
            strPath = getXmlPath(ndxList(intK))
            ' Get all the attributes
            For intL = 0 To ndxList(intK).Attributes.Count - 1
              ' Review this attribute
              attrThis = ndxList(intK).Attributes(intL)
              strCat = strPath & ":" & attrThis.Name
              ' Add attribute name and value to the list
              If (dicFile.ContainsKey(strCat)) Then
                ' This should not take place: give warning
                Logging("Multiple use of key [" & strCat & "]")
              Else
                dicFile.Add(strCat, attrThis.Value)
                ' Check for category
                If (Not colCats.Contains(strCat)) Then colCats.Add(strCat, strCat) : Debug.Print("Adding cat=[" & strCat & "]")
              End If
            Next (intL)
          Next intK
        Next intJ
        ' Do the <metadata> section if it exists
        ndxList = pdxFile.SelectNodes("./descendant::metadata/child::meta")
        For intK = 0 To ndxList.Count - 1
          ' Review this item
          strPath = getXmlPath(ndxList(intK))
          strCat = strPath & ":" & ndxList(intK).Attributes("id").Value
          ' Add this item
          dicFile.Add(strCat, ndxList(intK).InnerText)
          ' Check for category
          If (Not colCats.Contains(strCat)) Then colCats.Add(strCat, strCat) : Debug.Print("Adding cat=[" & strCat & "]")
        Next intK
        ' Add this to the corpus dictionary
        If (dicFile.Count > 0) Then dicCorpus.Add(IO.Path.GetFileNameWithoutExtension(strFile), dicFile)
      Next intI
      ' Convert the dictionary into a CSV and HTML table:
      ' First row:  short category name
      ' Second row: full category path
      ' Next rows:  [filename] [cat value 1] [cat value 2] ...
      ' (1) First row: category short names
      strRow = "File"
      sbHtml.Add("<table><thead><th>" & strRow & "</th>")
      For intI = 1 To colCats.Count
        ' Get the name of the category (full)
        strCat = colCats.Item(intI)
        Dim strShort As String = Strings.Mid(strCat, InStrRev(strCat, ":") + 1)
        strRow &= vbTab & strShort
        sbHtml.Add("<th>" & strShort & "</th>")
      Next intI
      sbOut.Add(strRow)
      sbHtml.Add("</thead><tbody>")
      ' (2) Second row: full category paths
      strRow = "(Category)"
      sbHtml.Add("<tr><td>" & strRow & "</td>")
      For intI = 1 To colCats.Count
        ' Get the name of the category (full)
        strCat = colCats.Item(intI)
        strRow &= vbTab & strCat
        sbHtml.Add("<td>" & strCat & "</td>")
      Next intI
      sbOut.Add(strRow)
      sbHtml.Add("</tr>")
      ' (3) Walk all files
      For Each kvp As KeyValuePair(Of String, Dictionary(Of String, String)) In dicCorpus
        Dim strFthis As String = kvp.Key  ' Which file are we working on?
        ' Start this row
        strRow = strFthis
        sbHtml.Add("<tr><td>" & strRow & "</td>")
        ' Get the dictionary stored for this file
        Dim dicFthis As Dictionary(Of String, String) = kvp.Value
        ' Walk all the categories we have
        For intI = 1 To colCats.Count
          Dim strValue As String = "-"      ' The value of the current category
          ' Get this category
          strCat = colCats.Item(intI)
          ' Does this file have
          If (dicFthis.ContainsKey(strCat)) Then
            ' Get the value of this feature
            strValue = dicFthis.Item(strCat)
            ' Process CR and LF (as well as tab)
            strValue = strValue.Replace(vbCr, "").Replace(vbLf, "\n").Replace(vbTab, "\t")
          End If
          ' Add the value to this row
          strRow &= vbTab & strValue
          sbHtml.Add("<td>" & strValue & "</td>")
        Next intI
        ' Finish this row
        sbOut.Add(strRow)
        sbHtml.Add("</tr>")
      Next kvp
      ' This is the csv table
      strCsv = sbOut.Text
      ' Also provide the html table
      sbHtml.Add("</tbody></table>")
      strHtml = sbHtml.Text
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/TeiGetCorpusTable error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   TeiImportCorpusTable
  ' Goal:   Import the metadata provided by the caller in a .csv file into the file
  '           (the .psdx ones) located in the directory the user indicates
  '         Adapt the metadata of the indicated files accordingly
  ' History:
  ' 26-06-2015  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function TeiImportCorpusTable() As Boolean
    Dim strDir As String = ""   ' Directory where to look
    Dim strFile As String = ""  ' Name of the file containing the csv metadata
    Dim strPath As String       ' Path to the metadata
    Dim strAttrName As String   ' Name of attribute
    Dim strName As String       ' Name
    Dim strValue As String      ' Value
    Dim arFile() As String      ' Result of looking for files
    Dim arLine() As String      ' Array of lines
    Dim arField() As String     ' Array of fields
    Dim arFieldName() As String ' Array of field names
    Dim arShort() As String     ' Short names
    Dim arMeta() As String      ' Division of metadata
    Dim pdxFile As XmlDocument  ' One file
    Dim ndxNode As XmlNode      ' One node
    Dim intI As Integer         ' COunter
    Dim intJ As Integer         ' COunter

    Try
      ' Ask for directory where the files are located and so forth
      With frmMetaInfo
        .grpFormat.Visible = False
        Select Case .ShowDialog
          Case MsgBoxResult.Cancel
            Status("Aborted")
            Return False
        End Select
        ' Get the results
        strFile = .FileName
        strDir = .PsdxDir
      End With
      ' Read the file with the metadata stuff
      arLine = IO.File.ReadAllLines(strFile)
      ' Line '0' contains the short field names
      arShort = arLine(0).Split(vbTab)
      ' Line '1' contains the actual field names
      arFieldName = arLine(1).Split(vbTab)
      ' Walk through the lines
      For intI = 2 To arLine.Length - 1
        If (Trim(arLine(intI).Length > 0)) Then
          ' Process this line
          arField = Split(arLine(intI), vbTab)
          ' Get the file name from the first field
          strName = arField(0) & loc_strFileExt
          ' Find the file from the directory
          arFile = IO.Directory.GetFiles(strDir, strName, IO.SearchOption.AllDirectories)
          If (arFile.Length > 0) Then
            Logging("Processing: " & strFile, True)
            ' Get the file
            strFile = arFile(0)
            ' Read the file into pdx
            If (Not ReadXmlDoc(strFile, pdxFile)) Then
              Logging("Could not read file: " & strFile)
              Return False
            End If
            SetXmlDocument(pdxFile)
            ' Visit all the fields
            For intJ = 1 To arField.Length - 1
              ' Get the value
              strValue = arField(intJ).Replace("\n", vbCrLf)
              ' Process this field's value
              strPath = arFieldName(intJ)
              strAttrName = Mid(strPath, InStrRev(strPath, ":") + 1)
              strPath = Left(strPath, InStrRev(strPath, ":") - 1)
              arMeta = Split(strPath, "/")
              If (arMeta(2) = "metadata") Then
                ndxNode = pdxFile.SelectSingleNode("./descendant::metadata/child::meta[@id='" & strAttrName & "']")
                If (ndxNode Is Nothing) Then
                  If (pdxFile.SelectSingleNode("./descendant::metadata") Is Nothing OrElse
                      pdxFile.SelectSingleNode("./descendant::metadata/child::meta") Is Nothing) Then
                    ' The node does not exist, so we need to make it
                    ndxNode = SetNodePath(pdxFile, Mid(strPath, 2))
                  Else
                    ' The node is there, but we need to add one
                    ndxNode = pdxFile.SelectSingleNode("./descendant::metadata")
                    ndxNode = AddXmlChildAfter(ndxNode,
                                     pdxFile.SelectSingleNode("./descendant::metadata/child::meta"),
                                     "meta")
                  End If
                  ' Also make the correct attribute
                  AddXmlAttribute(pdxFile, ndxNode, "id", strAttrName)
                End If
                ndxNode.InnerText = strValue
              Else
                ' Calculate the node
                ndxNode = pdxFile.SelectSingleNode("/descendant::*" & Join(arMeta, "/child::"))
                If (ndxNode Is Nothing) Then
                  ' Create the path
                  ndxNode = SetNodePath(pdxFile, Mid(strPath, 2))
                End If
                If (ndxNode.Attributes(strAttrName) Is Nothing) Then
                  ' Create the attribute
                  AddXmlAttribute(pdxFile, ndxNode, strAttrName)
                End If
                ndxNode.Attributes(strAttrName).Value = strValue
              End If
            Next intJ
            ' Save changes
            pdxFile.Save(strFile)
            ' Log what we've done
            Logging("Processed file: " & strFile, True)
          Else
            Logging("skipping file: " & strFile)
          End If
        End If
      Next intI

      ' Return positively
      Status("Ready")
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/TeiImportCorpusTable error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   TeiGetLanguage
  ' Goal:   Initialise the tei metadata stuff
  ' History:
  ' 11-08-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function TeiGetLanguage(ByVal strType As String) As String
    Dim ndxWork As XmlNode
    Try
      ndxWork = pdxCurrentFile.SelectSingleNode("./descendant::langUsage/child::language[1]")
      If (ndxWork Is Nothing OrElse ndxWork.Attributes(strType) Is Nothing) Then
        Return ""
      Else
        Return ndxWork.Attributes(strType).Value
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/TeiGetLanguage error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   TeiGetGenre
  ' Goal:   Initialise the tei metadata stuff
  ' History:
  ' 11-08-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function TeiGetGenre() As String
    Dim ndxWork As XmlNode
    Try
      ndxWork = pdxCurrentFile.SelectSingleNode("./descendant::langUsage/child::creation[1]")
      If (ndxWork Is Nothing OrElse ndxWork.Attributes("genre") Is Nothing) Then
        Return ""
      Else
        Return ndxWork.Attributes("genre").Value
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/TeiGetGenre error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   TeiMetaDataAdd
  ' Goal:   Add one item to the tei metadata in the file as well as in [tdlMeta]
  ' History:
  ' 11-08-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function TeiMetaDataAdd() As Boolean
    Dim strId As String = ""  ' Name of node

    Try
      ' Validate
      If (tdlMeta Is Nothing) Then Return False
      ' Ask for the name
      strId = GetName("Give the name for this metadata")
      ' Check result
      If (strId = "") Then Return False
      ' Add the name/value pair
      Return TeiMetaDataAddOne(strId, "")
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/TeiMetaDataAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return False
    End Try
  End Function
  Public Function TeiMetaDataAddOne(ByVal strMetaId As String, ByVal strMetaValue As String,
                                    Optional ByVal bOnlyPdx As Boolean = False) As Boolean
    Dim dtrFound() As DataRow ' Result of looking for it
    Dim dtrThis As DataRow    ' one datarow
    Dim ndxThis As XmlNode = Nothing

    Try
      ' Need to do the dataset part?
      If (Not bOnlyPdx) Then
        ' Validate
        If (tdlMeta Is Nothing) Then Return False
        ' Try find meta id
        dtrFound = tdlMeta.Tables("meta").Select("name='" & strMetaId & "'")
        If (dtrFound.Length = 0) Then
          ' Add one element to [tdlMeta]
          dtrThis = AddOneDataRow(tdlMeta, "meta", "metaId", "metadata")
        Else
          dtrThis = dtrFound(0)
        End If
        ' Correct the record
        With dtrThis
          .Item("name") = strMetaId
          .Item("value") = strMetaValue
        End With
      End If
      ' Find the correct node
      ndxThis = pdxCurrentFile.SelectSingleNode("./descendant::metadata[1]/child::meta[@id='" & strMetaId & "']")
      If (ndxThis Is Nothing) Then
        ' Find the <metadata> node
        ndxThis = SetNodePath(pdxCurrentFile, "teiHeader/metadata")
        ' Add the node
        SetXmlDocument(pdxCurrentFile)
        ndxThis = AddXmlChild(ndxThis, "meta", "id", strMetaId, "attribute")
        If (ndxThis Is Nothing) Then Return False
      End If
      ' Adapt the value
      ndxThis.InnerText = strMetaValue
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/TeiMetaDataAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   TeiMetaDataDel
  ' Goal:   Delete one item from the tei metadata stuff
  ' History:
  ' 11-08-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function TeiMetaDataDel(ByVal strId As String, ByVal strValue As String) As Boolean
    Dim ndxList As XmlNodeList        ' Candidates for the correct node
    Dim ndxThis As XmlNode = Nothing  ' The actually correct node
    Dim dtrFound() As DataRow         ' one datarow
    Dim intI As Integer               ' Counter
    Dim bFound As Boolean = False     ' Flag

    Try
      ' Validate
      If (tdlMeta Is Nothing) Then Return False
      ' Find the item
      ndxList = pdxCurrentFile.SelectNodes("./descendant::metadata/child::meta[@id='" & strId & "']")
      For intI = 0 To ndxList.Count - 1
        If (ndxList(intI).InnerText = strValue) Then
          bFound = True : Exit For
        End If
      Next intI
      ' Check if we found the value
      If (Not bFound) Then Return False
      ' Set the entry
      ndxThis = ndxList(intI)
      If (ndxThis Is Nothing) Then Return False
      ' Delete it from the [tdlMeta]
      dtrFound = tdlMeta.Tables("meta").Select("name='" & strId & "' AND value='" & strValue.Replace("'", "''") & "'")
      If (dtrFound.Length > 0) Then
        dtrFound(0).Delete()
        tdlMeta.AcceptChanges()
      End If
      ' Delete it from the file
      ndxThis.RemoveAll()
      ndxThis.ParentNode.RemoveChild(ndxThis)
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/TeiMetaDataDel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   TeiMetaDataEdit
  ' Goal:   Edit one item from the tei metadata stuff
  ' History:
  ' 11-08-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function TeiMetaDataEdit(ByVal strId As String, ByVal strValue As String) As Boolean
    Dim dtrFound() As DataRow            ' one datarow
    Dim ndxThis As XmlNode = Nothing

    Try
      ' Validate
      If (tdlMeta Is Nothing) Then Return False
      ' Get the item
      ndxThis = pdxCurrentFile.SelectSingleNode("./descendant::metadata[1]/child::meta[@id='" & strId & "']")
      If (ndxThis Is Nothing) Then
        ' Add one afresh
        If (Not TeiMetaDataAddOne(strId, strValue)) Then Return False
      Else
        ' Change the value
        ndxThis.InnerText = strValue
        ' Edit [tdlMeta]
        dtrFound = tdlMeta.Tables("meta").Select("name='" & strId & "'")
        If (dtrFound.Length > 0) Then
          dtrFound(0).Item("value") = strValue
        End If
      End If
      ' Set the dirty flag
      frmMain.SetDirty(True)
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/TeiMetaDataEdit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return False
    End Try
  End Function
  Public Function TeiMetaDataExists(ByVal strId As String) As Boolean
    Dim ndxThis As XmlNode = Nothing

    Try
      ' Validate
      If (tdlMeta Is Nothing) Then Return False
      ' Get the item
      ndxThis = pdxCurrentFile.SelectSingleNode("./descendant::metadata[1]/child::meta[@id='" & strId & "']")
      Return (ndxThis IsNot Nothing)
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/TeiMetaDataExists error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return False
    End Try
  End Function
  Public Function TeiMetaDataGet(ByVal strId As String) As String
    Dim ndxThis As XmlNode = Nothing

    Try
      ' Validate
      If (tdlMeta Is Nothing) Then Return False
      ' Get the item
      ndxThis = pdxCurrentFile.SelectSingleNode("./descendant::metadata[1]/child::meta[@id='" & strId & "']")
      If (ndxThis Is Nothing) Then
        Return ""
      Else
        Return ndxThis.InnerText
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/TeiMetaDataGet error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   TeiMetaDataOneFile
  ' Goal:   Take file [strFile] and process the metadata of [strName] in it
  ' History:
  ' 14-08-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function TeiMetaDataOneFile(ByVal strFile As String, ByVal strMethod As String, ByRef bChanged As Boolean) As Boolean
    Dim strMetaId As String     ' Meta id
    Dim strMetaVal As String    ' meta value
    Dim strName As String       ' Name of the current file
    Dim strLineName As String   ' Name in the line
    Dim strStrip As String      ' Completely stripped name
    Dim strFldName As String    ' Field name
    Dim strFldValue As String   ' Field value
    Dim arText() As String      ' Content
    Dim arLine() As String      ' One line
    Dim mcThis As MatchCollection
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter
    Dim intIdx As Integer       ' Index
    Dim intCol As Integer = -1  ' Column
    Dim intColumns As Integer   ' Number of columns
    Dim intSpare As Integer = -1  ' Spare column

    Try
      ' Initialise changes
      bChanged = False
      ' Has a file been loaded?
      If (pdxCurrentFile Is Nothing) Then Logging("First load a text") : Return False
      ' Get our name
      strName = LCase(IO.Path.GetFileName(strCurrentFile))
      ' Possibly adapt the name
      If (DoLike(strName, "*.psdx|*.psd")) Then strName = IO.Path.GetFileNameWithoutExtension(strName)
      ' Get completely stripped version of the name
      strStrip = IO.Path.GetFileNameWithoutExtension(strName)
      ' Early modern English: strip off -p1, -p2 or -h
      strStrip = Regex.Replace(strStrip, "(\-(p1|p2|h))", "")
      'If (DoLike(Right(strStrip, 3), "-p1|-p2")) Then
      '  ' Strip it off
      '  strStrip = Left(strStrip, strStrip.Length - 3)
      'ElseIf (Right(strStrip, 2) = "-h") Then
      '  ' Strip it off
      '  strStrip = Left(strStrip, strStrip.Length - 2)
      'End If
      ' Read the file into an array
      arText = IO.File.ReadAllLines(strFile)
      ' Action depends on method
      Select Case strMethod
        Case "columns", "Columns"
          ' The first column contains the names of the features; the first row contains the names of the files
          arLine = Split(arText(0), vbTab)
          ' Look for the column with the correct filename
          For intI = 1 To arLine.Length - 1
            ' Get the name in the line
            'strLineName = LCase(Trim(arLine(intI)))
            intCol = -1
            ' We have the file name column -- extract all the filenames in this cell
            mcThis = Regex.Matches(LCase(Trim(arLine(intI))), "\w+(\-\w+)?(\-\d+)?\.e[123]")
            For intJ = 0 To mcThis.Count - 1
              ' Get this match
              strLineName = mcThis(intJ).Value
              ' Process this match for comparison
              strLineName = Regex.Replace(strLineName, "(?<period>\.)(?<emode>e[123]$)", "-$2")
              ' Are we there?
              ' Debug.Print(strLineName)
              If (strName = strLineName) OrElse (strStrip = strLineName) Then
                ' This is the one!
                intCol = intI : Exit For
              End If
            Next intJ
            If (intCol >= 0) Then Exit For

            '' EmodE: change .e1, .e2, .e3 to -e1, -e2, -e3 etc
            'strLineName = Regex.Replace(strLineName, "(?<period>\.)(?<emode>e[123]$)", "-$2")
            'If (strName Like strLineName) Then
            '  ' This is the column
            '  intCol = intI : Exit For
            'ElseIf (strStrip Like strLineName) Then
            '  intSpare = intI
            'End If
          Next intI
          ' Found anything?
          If (intCol < 0) Then
            If (intSpare < 0) Then Logging("Skipping [" & strName & "]" & _
              " Could not locate it within [" & IO.Path.GetFileNameWithoutExtension(strFile) & "]") : Return True
            intCol = intSpare
          End If
          ' Walk the rows
          For intI = 1 To arText.Length - 1
            ' Get this row
            arLine = Split(arText(intI), vbTab)
            ' Get name and value
            strMetaId = Trim(arLine(0)) : strMetaVal = Trim(arLine(intCol))
            ' Add combination of name and value
            If (Not TeiMetaDataAddOne(strMetaId, strMetaVal)) Then Logging("Could not add data") : Return False
          Next intI
        Case "alternated", "Alternated", "alternating", "Alternating"
          ' The data for each text is stored in alternating columns: field name - field value
          If (loc_strMetaFile <> strFile) Then
            ' Transform the array into a better accessable data structure
            intColumns = Split(arText(0), vbTab).Length \ 2
            ReDim loc_arMeta(0 To intColumns - 1)
            ' Process the matter column by column
            strFldName = ""
            For intI = 0 To intColumns - 1
              ' Create a new string collection for this one
              loc_arMeta(intI) = New StringColl
              ' Visit all the rows
              For intJ = 0 To arText.Length - 1
                ' Get this line into an array
                arLine = Split(arText(intJ), vbTab)
                ' Find the field name
                If (Trim(arLine(intI * 2)) <> "") AndAlso (strFldName <> Trim(arLine(intI * 2))) Then
                  ' This is a new field name
                  strFldName = Trim(arLine(intI * 2))
                  ' Get the value (for the moment)
                  strFldValue = Trim(arLine(intI * 2 + 1))
                  ' Store them away
                  loc_arMeta(intI).Add(strFldName, strFldValue)
                Else
                  ' Get the field value
                  strFldValue = Trim(arLine(intI * 2 + 1))
                  ' Is this something?
                  If (strFldValue <> "") Then
                    intIdx = loc_arMeta(intI).Find(strFldName)
                    ' Add it to the existing one
                    loc_arMeta(intI).Exmp(intIdx) &= vbCrLf & strFldValue
                  End If
                End If
              Next intJ
            Next intI
            ' Keep track of what we've done
            loc_strMetaFile = strFile
          End If
          ' Find the correct file in loc_arMeta
          intCol = -1
          For intI = 0 To loc_arMeta.Length - 1
            ' Find the filename index in this file
            intIdx = loc_arMeta(intI).Find("Filename")
            If (intIdx >= 0) Then
              ' We have the file name column -- extract all the filenames in this cell
              mcThis = Regex.Matches(loc_arMeta(intI).Exmp(intIdx), "\w+(\-\w+)?(\-\d+)?\.e[123]")
              For intJ = 0 To mcThis.Count - 1
                ' Get this match
                strLineName = mcThis(intJ).Value
                ' Process this match for comparison
                strLineName = Regex.Replace(strLineName, "(?<period>\.)(?<emode>e[123]$)", "-$2")
                ' Are we there?
                ' Debug.Print(strLineName)
                If (strName = strLineName) OrElse (strStrip = strLineName) Then
                  ' This is the one!
                  intCol = intI : Exit For
                End If
              Next intJ
              ' Found the column?
              If (intCol >= 0) Then Exit For
            End If
          Next intI
          ' Found the column?
          If (intCol >= 0) Then
            ' Yes, check out this column
            For intI = 0 To loc_arMeta(intCol).Count - 1
              ' Get name and value
              strMetaId = loc_arMeta(intCol).Item(intI) : strMetaVal = loc_arMeta(intCol).Exmp(intI)
              ' Add combination of name and value
              If (Not TeiMetaDataAddOne(strMetaId, strMetaVal)) Then Logging("Could not add data") : Return False
            Next intI
          Else
            Logging("Skipping [" & strName & "]" & _
              " Could not locate it within [" & IO.Path.GetFileNameWithoutExtension(strFile) & "]") : Return True
          End If
      End Select
      ' Return positively
      bChanged = True
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/TeiMetaDataOneFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   TeiMetaDataFiles
  ' Goal:   Take file [strFile] and process the metadata of all the files found in [strDir] in it
  ' History:
  ' 14-08-2014  ERK Created
  ' 03-10-2014  ERK Added [strMethod]
  ' ------------------------------------------------------------------------------------------------------------
  Public Function TeiMetaDataFiles(ByVal strFile As String, ByVal strDir As String, ByVal strMethod As String) As Boolean
    Dim arFile() As String  ' Files in directory
    Dim intI As Integer     ' Counter
    Dim intPtc As Integer   ' Percentage
    Dim bChanged As Boolean = False

    Try
      ' Validate
      If (Not IO.Directory.Exists(strDir)) Then Logging("Cannot find directory " & strDir) : Return False
      ' Get all the psdx files here
      arFile = IO.Directory.GetFiles(strDir, "*.psdx", System.IO.SearchOption.TopDirectoryOnly)
      For intI = 0 To arFile.Length - 1
        ' Show where we are
        intPtc = (intI + 1) * 100 \ arFile.Length
        Status("Processing " & intPtc & "%", intPtc)
        ' Try read this file into an XML structure
        If (ReadXmlDoc(arFile(intI), pdxCurrentFile)) Then
          ' Set the current file
          strCurrentFile = arFile(intI)
          ' Process this file
          If (Not TeiMetaDataOneFile(strFile, strMethod, bChanged)) Then Logging("Could not process file [" & strCurrentFile & "]") : Return False
        End If
        ' Saving the changes...
        If (bChanged) Then
          pdxCurrentFile.Save(strCurrentFile)
          ' Show this one is processd
          Logging("Added metadata to: " & IO.Path.GetFileNameWithoutExtension(strCurrentFile))
        Else
          ' Show skipping this
          Logging("No changes to: " & IO.Path.GetFileNameWithoutExtension(strCurrentFile))
        End If
      Next intI
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/TeiMetaDataFiles error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return false
      Return False
    End Try
  End Function
End Module
