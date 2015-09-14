Option Strict On
Option Explicit On
Option Compare Binary

Imports System.Xml

Module modXML
  ' ========================= PUBLIC CONSTANTS =============================
  Public tdlSettings As DataSet = Nothing         ' Dataset with settings
  Public tdlShadow As DataSet = Nothing           ' Shadow dataset with settings
  ' ========================= LOCAL VARIABLES ==============================
  Private xmlSettings As New Xml.XmlDataDocument  ' Settings information
  Private xmlShadow As New Xml.XmlDataDocument    ' Shadow settings
  Private strShadowSchema As String = "CesaxSettings.xsd.txt"
  Private strShadowSettings As String = "CesaxSettings.xml.txt"
  Private arXmlIn() As String = {"&", "<", ">", """"}
  ' WERKT NIET ALTIJD: Private arXmlOut() As String = {"&amp;", "&lt;", "&gt;", "&quote;"}
  ' Check site: http://www.htmlhelp.com/reference/html40/entities/special.html
  Private arXmlOut() As String = {"&#38;", "&#60;", "&#62;", "&#34;"}
  Private arXmlNamed() As String = {"&amp;", "&lt;", "&gt;", "&quot;"}
  '---------------------------------------------------------------
  ' Name:       ReadSettingsXML()
  ' Goal:       Read the Settings XML file into a dataobject
  ' History:
  ' 04-12-2008  ERK Created
  '---------------------------------------------------------------
  Public Function ReadSettingsXML(ByVal strIn As String) As Boolean
    Dim strSource As String   ' XML data file
    Dim strScheme As String   ' XML scheme file

    ' Set correct file names
    strSource = strIn
    strScheme = Left(strIn, InStrRev(strIn, ".")) & "xsd"
    If (Not XmlToDataset(strScheme, strSource, tdlSettings, xmlSettings)) Then
      ' An error has occurred
      Status("modXml/ReadSettingsXML: unable to read settings: " & strSource)
      Return False
    End If
    ' Also try reading the shadow settings
    strSource = strShadowSettings
    strScheme = strShadowSchema
    If (Not XmlToDataset(strScheme, strSource, tdlShadow, xmlShadow)) Then
      ' An error has occurred
      Status("modXml/ReadSettingsXML: unable to read (shadow): " & strSource)
      Return False
    End If
    ' Return success
    ReadSettingsXML = True
  End Function

  '---------------------------------------------------------------
  ' Name:       XmlToDataset()
  ' Goal:       Given a Scheme (XSD) and File (XML), produce:
  '             - an XMLDataDocument
  '             - a Dataset
  ' History:
  ' 04-12-2008  ERK Created
  '---------------------------------------------------------------
  Public Function XmlToDataset(ByVal strScheme As String, ByVal strFile As String, _
    ByRef tdlData As DataSet, ByRef xmlFile As Xml.XmlDataDocument) As Boolean
    ' Dim strLocalScheme As String  ' Local scheme file

    ' Check if the filename needs adaptation
    If (InStr(strFile, "\") = 0) Then
      ' Need to adapt it
      strFile = IO.Path.GetDirectoryName(Application.ExecutablePath) & "\" & strFile
    End If
    ' Check existence of XML file
    If (Not IO.File.Exists(strFile)) Then
      ' Return failure
      XmlToDataset = False
      Exit Function
    End If
    ' Make new datadocument
    xmlFile = New Xml.XmlDataDocument
    ' Make sure errors are caught
    Try
      ' If a previous dataset existed, then cancel it
      If (Not tdlData Is Nothing) Then tdlData = Nothing
      ' Try to get the scheme
      If (FindScheme(strScheme)) Then
        ' Read the scheme
        xmlFile.DataSet.ReadXmlSchema(strScheme)
        ' Read the file
        xmlFile.DataSet.ReadXml(strFile, XmlReadMode.IgnoreSchema)
      Else
        ' Infer the scheme
        xmlFile.DataSet.ReadXml(strFile, XmlReadMode.InferSchema)
      End If
      '' Derive a local scheme file first
      'strLocalScheme = IO.Path.GetDirectoryName(Application.ExecutablePath) & "\" & IO.Path.GetFileNameWithoutExtension(strScheme) & ".xsd.txt"
      '' Does a scheme file exist?
      'If (IO.File.Exists(strScheme)) Then
      '  ' Read the scheme
      '  xmlFile.DataSet.ReadXmlSchema(strScheme)
      '  ' Read the file
      '  xmlFile.DataSet.ReadXml(strFile, XmlReadMode.IgnoreSchema)
      'ElseIf (IO.File.Exists(strLocalScheme)) Then
      '  ' Read the scheme
      '  xmlFile.DataSet.ReadXmlSchema(strLocalScheme)
      '  ' Read the file
      '  xmlFile.DataSet.ReadXml(strFile, XmlReadMode.IgnoreSchema)
      'Else
      '  ' Infer the scheme
      '  xmlFile.DataSet.ReadXml(strFile, XmlReadMode.InferSchema)
      'End If
      ' Assign data to dataset
      tdlData = xmlFile.DataSet
      ' Return success
      XmlToDataset = True
    Catch ex As Exception
      ' Show error to user
      Status(ex.Message)
      HandleErr("Error reading " & strFile & " (scheme=" & strScheme & "):" & vbCrLf & _
        ex.Message)
      ' Return failure
      XmlToDataset = False
    End Try
  End Function
  '---------------------------------------------------------------
  ' Name:       FindScheme()
  ' Goal:       Given a Scheme (XSD) and File (XML), produce:
  '             - a Dataset
  ' History:
  ' 24-02-2014  ERK Created
  '---------------------------------------------------------------
  Public Function FindScheme(ByRef strScheme As String) As Boolean
    Dim arFile() As String                  ' Results of searching a directory

    Try
      ' Check existence of Scheme
      If (Not IO.File.Exists(strScheme)) Then
        ' If the Schema file (xsd) is not on the indicated location, try reading it
        '  from the location of the current executable (where it is stored as .xsd.txt)
        arFile = IO.Directory.GetFiles(My.Application.Info.DirectoryPath, IO.Path.GetFileName(strScheme), IO.SearchOption.AllDirectories)
        ' Got this one?
        If (arFile.Length = 0) Then
          ' Try something else
          arFile = IO.Directory.GetFiles(My.Application.Info.DirectoryPath, IO.Path.GetFileName(strScheme) & ".txt", IO.SearchOption.AllDirectories)
          ' Anything there?
          If (arFile.Length = 0) Then Return False
        End If
        ' Take the first one as the scheme
        strScheme = arFile(0)
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modXml/FindScheme error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------
  ' Name:       ReadDataset()
  ' Goal:       Given a Scheme (XSD) and File (XML), produce:
  '             - a Dataset
  ' History:
  ' 04-12-2008  ERK Created
  '---------------------------------------------------------------
  Public Function ReadDataset(ByVal strScheme As String, ByVal strFile As String, _
    ByRef tdlData As DataSet) As Boolean
    Dim xmlFile As New Xml.XmlDataDocument  ' Needed
    'Dim arFile() As String                  ' Results of searching a directory

    ' Check existence of Scheme
    If (Not FindScheme(strScheme)) Then Return False
    'If (Not IO.File.Exists(strScheme)) Then
    '  ' If the Schema file (xsd) is not on the indicated location, try reading it
    '  '  from the location of the current executable (where it is stored as .xsd.txt)
    '  arFile = IO.Directory.GetFiles(My.Application.Info.DirectoryPath, IO.Path.GetFileName(strScheme), IO.SearchOption.AllDirectories)
    '  ' Got this one?
    '  If (arFile.Length = 0) Then
    '    ' Try something else
    '    arFile = IO.Directory.GetFiles(My.Application.Info.DirectoryPath, IO.Path.GetFileName(strScheme) & ".txt", IO.SearchOption.AllDirectories)
    '    ' Anything there?
    '    If (arFile.Length = 0) Then Return False
    '  End If
    '  ' Take the first one as the scheme
    '  strScheme = arFile(0)
    '  'strScheme = IO.Path.GetFullPath(My.Application.Info.DirectoryPath) & "\" & IO.Path.GetFileName(strScheme)
    '  '' Try and open this one
    '  'If (Not IO.File.Exists(strScheme)) Then
    '  '  ' It may still be there, with the added extension ".txt"...
    '  '  strScheme &= ".txt"
    '  '  If (Not IO.File.Exists(strScheme)) Then
    '  '    ' Return failure
    '  '    Return False
    '  '  End If
    '  'End If
    'End If
    ' Check existence of XML file
    If (Not IO.File.Exists(strFile)) Then
      ' Return failure
      ReadDataset = False
      Exit Function
    End If
    ' Make new datadocument
    xmlFile = Nothing
    xmlFile = New Xml.XmlDataDocument
    ' Make sure errors are caught
    Try
      ' Open scheme
      xmlFile.DataSet.ReadXmlSchema(strScheme)
      ' Load XML data
      xmlFile.Load(strFile)
      ' Delete previous dataset
      tdlData = Nothing
      ' Assign data to dataset
      tdlData = xmlFile.DataSet
      ' tdlData = createfromstring(Xml)
      ' Accept changes?
      tdlData.AcceptChanges()
      ' Return success
      ReadDataset = True
    Catch ex As Exception
      ' Show error to user
      Status(ex.Message)
      HandleErr("Error reading " & strFile & " (scheme=" & strScheme & "):" & vbCrLf & _
        ex.Message)
      ' Return failure
      ReadDataset = False
    End Try
  End Function
  '---------------------------------------------------------------
  ' Name:       ReadDataScheme()
  ' Goal:       Given a Scheme (XSD) produce a Dataset
  ' History:
  ' 27-11-2009  ERK Created
  '---------------------------------------------------------------
  Public Function ReadDataScheme(ByVal strScheme As String, ByRef tdlData As DataSet) As Boolean
    Dim xmlFile As New Xml.XmlDataDocument  ' Needed

    ' Check existence of Scheme
    If (Not FindScheme(strScheme)) Then Return False
    'If (Not IO.File.Exists(strScheme)) Then
    '  ' If the Schema file (xsd) is not on the indicated location, try reading it
    '  '  from the location of the current executable (where it is stored as .xsd.txt)
    '  strScheme = IO.Path.GetFullPath(My.Application.Info.DirectoryPath) & "\" & IO.Path.GetFileName(strScheme)
    '  ' Try and open this one
    '  If (Not IO.File.Exists(strScheme)) Then
    '    ' It may still be there, with the added extension ".txt"...
    '    strScheme &= ".txt"
    '    If (Not IO.File.Exists(strScheme)) Then
    '      ' Return failure
    '      Return False
    '    End If
    '  End If
    'End If
    ' Make new datadocument
    xmlFile = New Xml.XmlDataDocument
    ' Make sure errors are caught
    Try
      ' Open scheme
      xmlFile.DataSet.ReadXmlSchema(strScheme)
      ' Delete previous dataset
      tdlData = Nothing
      ' Assign data to dataset
      tdlData = xmlFile.DataSet
      ' Accept changes?
      tdlData.AcceptChanges()
      ' Return success
      ReadDataScheme = True
    Catch ex As Exception
      ' Show error to user
      Status(ex.Message)
      HandleErr("Error reading " & strScheme & " (scheme=" & strScheme & "):" & vbCrLf & _
        ex.Message)
      ' Return failure
      ReadDataScheme = False
    End Try
  End Function
  '---------------------------------------------------------------
  ' Name:       CreateDataSet()
  ' Goal:       Given a Scheme (XSD) produce a Dataset
  ' History:
  ' 02-10-2009  ERK Created
  '---------------------------------------------------------------
  Public Function CreateDataSet(ByVal strScheme As String, ByRef tdlData As DataSet) As Boolean
    Dim xmlFile As New Xml.XmlDataDocument  ' Needed

    ' Check existence of Scheme
    If (Not FindScheme(strScheme)) Then Return False
    'If (Not IO.File.Exists(strScheme)) Then
    '  ' If the Schema file (xsd) is not on the indicated location, try reading it
    '  '  from the location of the current executable (where it is stored as .xsd.txt)
    '  strScheme = IO.Path.GetFullPath(My.Application.Info.DirectoryPath) & "\" & IO.Path.GetFileName(strScheme)
    '  ' Try and open this one
    '  If (Not IO.File.Exists(strScheme)) Then
    '    ' It may still be there, with the added extension ".txt"...
    '    strScheme &= ".txt"
    '    If (Not IO.File.Exists(strScheme)) Then
    '      ' Return failure
    '      Return False
    '    End If
    '  End If
    'End If
    ' Make new datadocument
    xmlFile = New Xml.XmlDataDocument
    ' Make sure errors are caught
    Try
      ' Open scheme
      xmlFile.DataSet.ReadXmlSchema(strScheme)
      ' Delete the previous dataset
      tdlData = Nothing
      tdlData = New DataSet
      ' Assign data to dataset
      tdlData = xmlFile.DataSet
      ' Return success
      CreateDataSet = True
    Catch ex As Exception
      ' Show error to user
      Status(ex.Message)
      HandleErr("CreateDataset error (scheme=" & strScheme & "):" & vbCrLf & _
        ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      CreateDataSet = False
    End Try
  End Function
  ' --------------------------------------------------------------------------------------
  ' Name :  GetTableID
  ' Goal :  Return the integer value of the strIdField from table strTable
  '         Select the datarow where strNameField has the value strNameValue
  ' History:
  ' 22-03-2008  ERK Created
  ' --------------------------------------------------------------------------------------
  Public Function GetTableID(ByRef tdlThis As DataSet, ByVal strTable As String, _
    ByVal strNameField As String, ByVal strNameValue As String, ByVal strIdField As String, Optional ByVal bExact As Boolean = False) As Integer
    Dim arRow() As DataRow      ' Local datarow array

    Try
      arRow = tdlThis.Tables(strTable).Select(strNameField & " = '" & strNameValue & "'")
      ' See if at least one with this path is found
      If (arRow.Length = 0) Then
        ' Do we need exact comparison?
        If (Not bExact) Then
          ' Try again with partial comparison
          arRow = tdlThis.Tables(strTable).Select(strNameField & " >= '" & strNameValue & "'")
          ' Has something been found now?
          If (arRow.Length > 0) Then
            ' Yes, so return the first row fulfilling criteria
            Return CType(arRow(0).Item(strIdField), Integer)
          End If
        End If
      Else
        ' Return the ID of the first row with this path
        Return CType(arRow(0).Item(strIdField), Integer)
      End If
      ' Default: return -1
      Return -1
    Catch ex As Exception
      ' Show error
      HandleErr("modXml/GetTableId error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      Return -1
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  GetNewId
  ' Goal :  Get a new ID in table [tblThis] for index [strItem]
  ' History:
  ' 23-09-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function GetNewId(ByRef tblThis As DataTable, ByVal strItem As String) As Integer
    'Dim intI As Integer       ' Counter
    'Dim intThis As Integer    ' Index of current one
    Dim intMax As Integer     ' Maximum ID so far
    Dim intSize As Integer    ' Size of the table

    Try
      ' Does the table exist already?
      If (tblThis Is Nothing) OrElse (tblThis.Rows.Count = 0) Then
        ' Return 1
        Return 1
      End If
      ' How large a table is this anyway?
      intSize = tblThis.Rows.Count
      If (intSize > 10) Then
        ' Table is large, so get the maximum ID by just looking at the last member
        intMax = CInt(tblThis.Rows(intSize - 1).Item(strItem).ToString)
        ' Double check
        If (intMax < intSize) Then
          ' Calculate it differently anyway...
          intMax = CInt(tblThis.Select("", strItem & " DESC")(0).Item(strItem).ToString)
        End If
      Else
        ' Table is small, so get the maximum ID by using SELECT
        intMax = CInt(tblThis.Select("", strItem & " DESC")(0).Item(strItem).ToString)
      End If
      '' Table exists, so access it
      'With tblThis
      '  ' Get the value of the last row
      '  intMax = CInt(.Rows(.Rows.Count - 1).Item(strItem).ToString)
      '  'intMax = CInt(.Select("", strItem & " DESC")(0).Item(strItem).ToString)
      '  '' There is at least one row - get it
      '  'intMax = CInt(.Rows(0).Item(strItem).ToString)
      '  '' Check all other rows
      '  'For intI = 1 To .Rows.Count - 1
      '  '  ' Get this index
      '  '  intThis = CInt(.Rows(intI).Item(strItem).ToString)
      '  '  ' Is this a better one?
      '  '  If (intThis > intMax) Then intMax = intThis
      '  'Next intI
      'End With
      ' Return the result
      GetNewId = intMax + 1
    Catch ex As Exception
      ' Show there is something wrong
      MsgBox("modXml/GetNewId error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return 0
    End Try
  End Function

  '' ----------------------------------------------------------------------------------------------------------
  '' Name :  GetNewId
  '' Goal :  Get a new ID in table [tblThis] for index [strItem]
  '' History:
  '' 23-09-2009  ERK Created
  '' ----------------------------------------------------------------------------------------------------------
  'Public Function GetNewId(ByRef tblThis As DataTable, ByVal strIdField As String) As Integer
  '  Dim intI As Integer         ' Counter
  '  Dim intThis As Integer  ' Index of current one
  '  Dim intMax As Integer   ' Maximum ID so far

  '  Try
  '    ' Does the table exist already?
  '    If (tblThis Is Nothing) OrElse (tblThis.Rows.Count = 0) Then
  '      ' Return 1
  '      Return 1
  '    End If
  '    ' Table exists, so access it
  '    With tblThis
  '      ' There is at least one row - get it
  '      intMax = CInt(.Rows(0).Item(strIdField).ToString)
  '      ' Check all other rows
  '      For intI = 1 To .Rows.Count - 1
  '        ' Get this index
  '        intThis = CInt(.Rows(intI).Item(strIdField).ToString)
  '        ' Is this a better one?
  '        If (intThis > intMax) Then intMax = intThis
  '      Next intI
  '    End With
  '    ' Return the result
  '    GetNewId = intMax + 1
  '  Catch ex As Exception
  '    ' Note error
  '    HandleErr("modXML/GetNewId error: " & ex.Message & vbCr & ex.StackTrace & vbCrLf)
  '    ' Return failure
  '    Return 0
  '  End Try
  'End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  AddOneDataRow
  ' Goal :  Create one new datarow and return it
  ' History:
  ' 11-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Public Function AddOneDataRow(ByRef tdlThis As DataSet, ByVal strTable As String, ByVal strSelId As String, _
                             ByVal strParentTable As String, Optional ByVal bMakeId As Boolean = True) As DataRow
    Dim dtrThis As DataRow              ' General purpose datarow
    Dim dtrParent As DataRow = Nothing  ' The parent table
    Dim intNewId As Integer             ' New ID for this row

    Try
      ' Care about a parent?
      If (strParentTable <> "") Then
        ' Check existence of parent table
        If (tdlThis.Tables(strParentTable) Is Nothing) OrElse (tdlThis.Tables(strParentTable).Rows.Count = 0) Then
          ' Add one row
          dtrThis = tdlThis.Tables(strParentTable).NewRow
          tdlThis.Tables(strParentTable).Rows.Add(dtrThis)
        End If
        ' Set parent table
        dtrParent = tdlThis.Tables(strParentTable).Rows(0)
      End If
      ' Add a new datarow
      With tdlThis.Tables(strTable)
        ' Need to think of a new id?
        If (bMakeId) Then
          ' Determine the new ID
          intNewId = GetNewId(tdlThis.Tables(strTable), strSelId)
        Else
          ' Give a default ID value
          intNewId = 1
        End If
        'intNewId = .Rows.Count + 1
        ' Get a new row
        dtrThis = .NewRow
        ' Fill inthe values
        dtrThis.Item(strSelId) = intNewId
        If (strParentTable <> "") Then
          dtrThis.SetParentRow(dtrParent)
        End If
        ' Add this row
        .Rows.Add(dtrThis)
      End With
      ' Return the new datarow
      Return dtrThis
    Catch ex As Exception
      ' Show error
      HandleErr("modXml/AddOneDataRow error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' --------------------------------------------------------------------------------------
  ' Name :  GetFieldValue
  ' Goal :  Get the correct status name
  ' History:
  ' 13-03-2008  ERK Created
  ' --------------------------------------------------------------------------------------
  Public Function GetFieldValue(ByRef tdlThis As DataSet, ByVal strTable As String, _
    ByVal strIdName As String, _
    ByVal intIdValue As Integer, ByVal strField As String) As String
    Dim arRow() As DataRow

    ' Select all rows from the given table where the ID name has the ID value
    arRow = tdlThis.Tables(strTable).Select(strIdName & " = " & intIdValue)
    If (arRow.Length = 0) Then
      GetFieldValue = ""
    Else
      ' Return the value of the selected field
      GetFieldValue = arRow(0).Item(strField).ToString
    End If
  End Function
  ' --------------------------------------------------------------------------------------
  ' Name :  XmlSaveSave
  ' Goal :  Do a save storing of a file in XML form
  ' History:
  ' 15-05-2008  ERK Created
  ' --------------------------------------------------------------------------------------
  Public Sub XmlSaveSave(ByRef xmlThis As Xml.XmlDataDocument, ByVal strFile As String)
    Dim strBkUp As String   ' Backup file

    Try
      ' Make file saveable
      IO.File.SetAttributes(strFile, IO.File.GetAttributes(strFile) And Not IO.FileAttributes.ReadOnly)
      ' Make copy to backup file
      strBkUp = IO.Path.GetFileNameWithoutExtension(strFile) & "_" & Format(Today, "ddd") & ".bak"
      strBkUp = IO.Path.GetDirectoryName(strFile) & "\" & strBkUp
      IO.File.Copy(strFile, strBkUp, True)
      ' save file
      xmlThis.Save(strFile)
      ' make file read only
      IO.File.SetAttributes(strFile, IO.File.GetAttributes(strFile) Or IO.FileAttributes.ReadOnly)
    Catch ex As Exception
      ' Signal user
      HandleErr("There is a problem saving the XML file " & strFile & vbCrLf & _
        "The error is: " & ex.Message)
    End Try
  End Sub
  ' --------------------------------------------------------------------------------------
  ' Name :  XmlSaveSettings
  ' Goal :  Do a save storing of the settings in XML form
  ' History:
  ' 04-12-2008  ERK Created
  ' --------------------------------------------------------------------------------------
  Public Sub XmlSaveSettings(ByVal strFile As String)
    Call XmlSaveSave(xmlSettings, strFile)
  End Sub
  ' --------------------------------------------------------------------------------------
  ' Name :  XmlString
  ' Goal :  Convert certain characters into XML okay characters
  ' History:
  ' 04-12-2008  ERK Created
  ' --------------------------------------------------------------------------------------
  Public Function XmlString(ByVal strIn As String) As String
    Dim intI As Integer   ' Counter

    ' Check all input possibilities
    For intI = 0 To UBound(arXmlIn)
      ' Change input to output
      strIn = strIn.Replace(arXmlIn(intI), arXmlOut(intI))
    Next intI
    ' Return the string
    XmlString = strIn
  End Function
  ' --------------------------------------------------------------------------------------
  ' Name :  XmlEscape
  ' Goal :  Convert certain characters into XML okay characters
  ' History:
  ' 04-12-2008  ERK Created
  ' --------------------------------------------------------------------------------------
  Public Function XmlEscape(ByVal strIn As String) As String
    Dim intI As Integer   ' Counter

    ' Check all input possibilities
    For intI = 0 To UBound(arXmlIn)
      ' Change input to output
      strIn = strIn.Replace(arXmlIn(intI), arXmlNamed(intI))
    Next intI
    ' Return the string
    Return strIn
  End Function
  ' --------------------------------------------------------------------------------------
  ' Name :  XmlUnescape
  ' Goal :  Convert certain XML escape characters back into normal ones
  ' History:
  ' 04-12-2008  ERK Created
  ' --------------------------------------------------------------------------------------
  Public Function XmlUnescape(ByVal strIn As String) As String
    Dim intI As Integer   ' Counter

    ' Check all input possibilities
    For intI = 0 To UBound(arXmlOut)
      ' Change input to output
      strIn = strIn.Replace(arXmlOut(intI), arXmlIn(intI))
      ' Also consider named XML escape characters
      strIn = strIn.Replace(arXmlNamed(intI), arXmlIn(intI))
    Next intI
    ' Return the string
    XmlUnescape = strIn
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  GetFirstRow
  ' Goal :  Get (and possibly create) the first row of the [strName] table from the corpus file
  ' History:
  ' 23-09-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function GetFirstRow(ByRef tdlThis As DataSet, ByVal strName As String, _
                              ByVal ParamArray arParam() As String) As DataRow
    ' Does it already exist?
    If (tdlThis.Tables(strName) Is Nothing) Then
      ' Create it
      tdlThis.Tables.Add(strName)
    End If
    ' Are there any rows?
    If (tdlThis.Tables(strName).Rows.Count = 0) Then
      ' Add at least one row... (with empty values)
      tdlThis.Tables(strName).Rows.Add(arParam)
    End If
    ' Return the row
    GetFirstRow = tdlThis.Tables(strName).Rows(0)
  End Function
  '---------------------------------------------------------------
  ' Name:       ReadXmlDoc()
  ' Goal:       Given a Scheme (XSD) and File (XML), produce:
  '             - an XML document
  ' History:
  ' 23-04-2010  ERK Created
  ' 24-04-2014  ERK Allow using schema
  '---------------------------------------------------------------
  Public Function ReadXmlDoc(ByVal strFile As String, _
    ByRef xmlFile As Xml.XmlDocument, Optional ByVal strSchema As String = "") As Boolean
    Dim rdThis As Xml.XmlReader
    Dim setThis As New XmlReaderSettings

    ' Check existence of XML file
    If (Not IO.File.Exists(strFile)) Then
      ' Return failure
      ReadXmlDoc = False
      Exit Function
    End If
    ' Any schema?
    If (strSchema <> "") Then
      ' Set the schema
      If (Not FindScheme(strSchema)) Then Return False
      setThis.Schemas.Add("http://www.ru.nl/" & IO.Path.GetFileNameWithoutExtension(strSchema), strSchema)
    End If
    ' Make new datadocument
    xmlFile = Nothing
    xmlFile = New Xml.XmlDocument
    ' Make sure errors are caught
    Try
      ' Create reader
      If (strSchema = "") Then
        rdThis = New Xml.XmlTextReader(strFile)
      Else
        rdThis = XmlReader.Create(strFile, setThis)
      End If
      ' Load XML data
      xmlFile.Load(rdThis)
      rdThis.Close()
      ' Return success
      Return True
    Catch ex As Exception
      ' Check the kind of exception
      If (InStr(ex.Message, "OutOfMemory") > 0) Then
        HandleErr("modXml/ReadXmlDoc error..." & vbCrLf & _
               "Sorry, but your document is too large to be loaded." & vbCrLf & _
               "The problem is with: [" & IO.Path.GetFileName(strFile) & "]" & vbCrLf & _
               "The system-generated error is:" & vbCrLf & _
               ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      Else
        ' Show error to user
        HandleErr("modXml/ReadXmlDoc: Error reading " & strFile & vbCrLf & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      End If
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  CreateNewRow
  ' Goal :  Create a new row of the [strTable] table in dataset [tdlThis]
  '         The ID field is [strIdField], and the new ID value is returned in [intId]
  ' History:
  ' 26-12-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function CreateNewRow(ByRef tdlThis As DataSet, ByVal strTable As String, _
         ByVal strIdField As String, ByRef intId As Integer, _
         ByRef dtrNewRow As DataRow, ByVal ParamArray arParam() As String) As Boolean
    Try
      ' Does the table already exist?
      If (tdlThis.Tables(strTable) Is Nothing) Then
        ' Create it
        tdlThis.Tables.Add(strTable)
      End If
      ' Access the table
      With tdlThis.Tables(strTable)
        ' Should we make a new ID?
        If (strIdField <> "") Then
          ' Make new ID for this row
          intId = GetNewId(tdlThis.Tables(strTable), strIdField)
          ' Create a new datarow
          dtrNewRow = .Rows.Add(arParam)
          ' Save the ID value
          dtrNewRow.Item(strIdField) = intId
        Else
          ' At least make a new datarow
          dtrNewRow = .Rows.Add(arParam)
        End If
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Note error
      HandleErr("modXml/CreateNewRow error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  CreateListRow
  ' Goal :  Add one [...List] row to the dataset [tdlThis]
  ' History:
  ' 11-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Public Sub CreateListRow(ByRef tdlThis As DataSet, ByVal strTable As String)
    Dim dtrThis As DataRow      ' General purpose datarow

    Try
      ' Access the correct table
      With tdlThis.Tables(strTable)
        ' Create a new row
        dtrThis = .NewRow
        ' Add this new row
        .Rows.Add(dtrThis)
      End With
    Catch ex As Exception
      ' Note error
      HandleErr("modXml/CreateListRow error: " & ex.Message & vbCrLf & "Context: " & vbCrLf & ex.StackTrace)
    End Try
  End Sub
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  AddOneDataRowWithParent
  ' Goal :  Create one new datarow and return it
  ' History:
  ' 11-12-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Public Function AddOneDataRowWithParent(ByRef tdlThis As DataSet, ByVal strTable As String, ByVal strSelId As String, _
                             ByRef dtrParent As DataRow) As DataRow
    Dim dtrThis As DataRow    ' General purpose datarow

    Try
      ' Add a new datarow
      With tdlThis.Tables(strTable)
        '' Determine the new ID
        'intNewId = .Rows.Count + 1
        ' Get a new row
        dtrThis = .NewRow
        ' Do we have an ID to make?
        If (strSelId <> "") Then
          ' Set the ID
          dtrThis.Item(strSelId) = GetNewId(tdlThis.Tables(strTable), strSelId)
        End If
        ' Do we have a parent?
        If (Not dtrParent Is Nothing) Then
          ' Set the parent row properly
          dtrThis.SetParentRow(dtrParent)
        End If
        ' Add this row
        .Rows.Add(dtrThis)
      End With
      ' Return the new datarow
      Return dtrThis
    Catch ex As Exception
      ' Show error
      MsgBox("modXml/AddOneDataRowWithParent error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
End Module
