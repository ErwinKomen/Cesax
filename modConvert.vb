Imports System.Xml
Imports System.Text.RegularExpressions
Module modConvert
  ' ===================================== LOCAL VARIABLES ===========================================================
  Private pdxConv As XmlDocument        ' One XML document new 
  Private loc_intEtreeId As Integer     ' Highest <eTree> @Id value assigned for this file so far
  Private loc_arLang() As String = {"Dutch", "German", "English", "Spanish", "French", "Welsh", "Vlaams", "Lezgi", "Lak", "Chechen"}
  Private loc_arEthno() As String = {"nld", "deu", "eng", "spa", "fra", "cym", "vls", "lez", "lak", "che"}
  Private loc_colCgnInfo As New StringColl  ' Information collection
  Private loc_colTigInfo As New StringColl  ' Information collection
  Private loc_colTigMain As New StringColl  ' Main collection
  Private loc_nsDf As XmlNamespaceManager = Nothing ' For folia processing
  Private loc_colFoliaId As New StringColl  ' Collection of FoLiA xml:id elements relevant for coreference
  Private loc_colPosInfo As New StringColl  ' Collectin of POS types gathered in Flex reading
  Private loc_colLezgiToPos As New StringColl
  ' ===================================== Public structures ==================================================
  Public Structure HeaderInfo
    Dim Type As String    ' Main type
    Dim SubType As String ' Sub type
    Dim Value As String   ' Value
    Dim Box As TextBox    ' Reference to textbox
    Dim Rich As RichTextBox
    Public Sub New(t As String, s As String, b As TextBox, r As RichTextBox)
      Type = t : SubType = s : Value = "" : Box = b : Rich = r
    End Sub
  End Structure
  ' ==========================================================================================================
  '----------------------------------------------------------------------------------------
  ' Name:       LangToEthno()
  ' Goal:       Convert a language name into an ethnologue code
  ' History:
  ' 27-01-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function LangToEthno(ByVal strLang As String) As String
    Dim intI As Integer

    Try
      ' Validate
      strLang = LCase(Trim(strLang))
      If (strLang = "") Then Return ""
      ' Loop
      For intI = 0 To loc_arLang.Length - 1
        If (LCase(loc_arLang(intI)) = strLang) Then Return loc_arEthno(intI)
      Next intI
      ' Return the language as given
      Return strLang
    Catch ex As Exception
      ' Show error
      HandleErr("modConvert/LangToEthno error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       EthnoToName()
  ' Goal:       Convert an ethnologue code into a language name
  ' History:
  ' 27-01-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function EthnoToName(ByVal strEthno As String) As String
    Dim intI As Integer

    Try
      ' Validate
      strEthno = LCase(Trim(strEthno))
      If (strEthno = "") Then Return ""
      ' Loop
      For intI = 0 To loc_arLang.Length - 1
        If (LCase(loc_arEthno(intI)) = strEthno) Then Return loc_arLang(intI)
      Next intI
      ' Return failure
      Return ""
    Catch ex As Exception
      ' Show error
      HandleErr("modConvert/EthnoToName error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name:   BatchConvertToFolia
  ' Goal:   Produces Folia files from all the psdx files in the selected directory
  ' History:
  ' 12-03-2014  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function BatchConvertToFolia(ByVal strCnvType As String, ByVal strEthno As String) As Boolean
    Dim strDirIn As String = ""     ' input directory
    Dim strDirOut As String = ""    ' output directory
    Dim strDirApp As String = ""    ' Part of input directory to be appended to output directory for this file
    Dim strInFile As String         ' One input file
    Dim strOutFile As String        ' One output file
    Dim strExt As String = ""       ' Allowed input extensions
    Dim strLanguage As String = ""  ' Language
    Dim strShort As String = ""     ' Short file name
    Dim arInFile() As String        ' Array of input files
    Dim arCnv() As String
    Dim intPtc As Integer           ' Percentage
    Dim intI As Integer             ' Counter
    Dim intJ As Integer             ' Counter
    Dim bFound As Boolean = False   ' Flag
    Dim bOverwrite As Boolean = False    ' Whether we may overwrite or not
    Dim arCnvType() As String = {"psdx", "Flex"}
    Dim arExtens() As String = {"psdx;psdx.xml", "flextext"}

    Try
      ' Perform language initialisation -- only needed if possible
      If (Not DoLangInit(strEthno)) Then
        Logging("Warning: no language initialisation specified for [" & strEthno & "]")
      End If
      ' Let user select an input and output directory
      With frmDirs
        ' Initialise the directories myself
        .SrcDir = IIf(strLastFoliaSrc = "", strWorkDir, strLastFoliaSrc)
        .DstDir = .SrcDir
        .DstDir = IIf(strLastFoliaDst = "", .SrcDir, strLastFoliaDst)
        '' If the directory does not exist, then make it
        'If (Not IO.Directory.Exists(.DstDir)) Then IO.Directory.CreateDirectory(.DstDir)
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Get the source and destination directories
            strDirIn = .SrcDir
            strDirOut = .DstDir
          Case Else
            Return False
        End Select
      End With
      ' Save this directory, if it is different from what we have
      If (strLastFoliaSrc <> strDirIn) Then
        strLastFoliaSrc = strDirIn
        SetTableSetting(tdlSettings, "LastFoliaSrc", strLastFoliaSrc)
        XmlSaveSettings(strSetFile)
      End If
      If (strLastFoliaDst <> strDirOut) Then
        strLastFoliaDst = strDirOut
        SetTableSetting(tdlSettings, "LastFoliaDst", strLastFoliaDst)
        XmlSaveSettings(strSetFile)
      End If
      ' Check whether correct directories were selected
      If (IO.Directory.Exists(strDirIn)) AndAlso (IO.Directory.Exists(strDirOut)) Then
        ' Get the conversion index
        bFound = False
        For intJ = 0 To arCnvType.Length - 1
          arCnv = Split(arCnvType(intJ), ";")
          For intI = 0 To arCnv.Length - 1
            If (strCnvType = arCnv(intI)) Then
              strExt = arExtens(intJ) : bFound = True
              Exit For
            End If
          Next intI
          If (bFound) Then Exit For
        Next intJ
        If (Not bFound) Then Return False
        ' Find the input files by looking for the file extensions (in order)...
        arCnv = Split(strExt, ";") : ReDim arInFile(0)
        For intI = 0 To arCnv.Length - 1
          ' Also look in subdirectories!!
          arInFile = IO.Directory.GetFiles(strDirIn, "*." & arCnv(intI), IO.SearchOption.AllDirectories)
          If (arInFile.Count > 0) Then Exit For
        Next intI
        ' Visit all files
        For intI = 0 To arInFile.Count - 1
          ' Show where we are
          intPtc = (intI + 1) * 100 \ arInFile.Count
          Status("Processing " & intPtc & "%", intPtc)
          ' Get the input file name
          strInFile = arInFile(intI)
          ' Find name without dots
          strShort = GetShortFileName(strInFile, arCnv)
          ' Determine the output file name, but WITHOUT EXTENSION yet
          strDirApp = Mid(IO.Path.GetDirectoryName(strInFile), strDirIn.Length + 1)
          strOutFile = strDirOut & strDirApp & "\" & strShort
          ' Check existence of directory
          If (Not IO.Directory.Exists(strDirOut & strDirApp)) Then
            IO.Directory.CreateDirectory(strDirOut & strDirApp)
          End If
          ' Show where we are
          Logging("Producing FoLiA " & intI + 1 & "/" & arInFile.Count & " - " & _
                  strShort)
          ' Check existence of final PSDX result
          If (Not bOverwrite) AndAlso (IO.File.Exists(strOutFile & ".folia.xml")) Then
            ' Ask if overwriting is okay
            If (MsgBox("Can I overwrite files like: [" & strOutFile & ".folia.xml" & "]?", _
                       MsgBoxStyle.YesNoCancel) <> MsgBoxResult.Yes) Then
              ' No overwriting, so exit 
              Return False
            End If
            ' We can overwrite, so put to true
            bOverwrite = True
          End If
          ' Call the actual conversion function
          Select Case strCnvType
            Case "psdx"
              If (Not ConvertOnePsdxToFolia(strInFile, strOutFile & ".folia.xml")) Then
                ' Check for interrupt
                If (bInterrupt) Then Status("Interrupted with error...")
                ' There was an error, so stop
                Return False
              End If
            Case "Flex"
              If (Not ConvertOneFlexToFolia(strInFile, strOutFile & ".folia.xml")) Then
                ' Check for interrupt
                If (bInterrupt) Then Status("Interrupted with error...")
                ' There was an error, so stop
                Return False
              End If
          End Select
          ' Check for interrupt
          If (bInterrupt) Then Status("Interrupted by user...") : Return False
        Next intI
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modConvert/BatchConvertToFolia error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name:   BatchConvertToPsdx
  ' Goal:   Produces PSDX files from a variety of sources:
  '         (1) From CONLL-X  all the CON or CONLL files in the selected directory
  ' Note:   If a conversion type is language dependent, the language will be asked from the user
  ' History:
  ' 27-01-2014  ERK Created
  ' 17-02-2014  ERK Added "folia" type
  ' 29-06-2015  ERK Started extending "alpino" type
  ' ---------------------------------------------------------------------------------------------------------
  Public Function BatchConvertToPsdx(ByVal strCnvType As String, ByVal strEthno As String) As Boolean
    Dim strDirIn As String = ""     ' input directory
    Dim strDirOut As String = ""    ' output directory
    Dim strDirApp As String = ""    ' Part of input directory to be appended to output directory for this file
    Dim strInFile As String         ' One input file
    Dim strOutFile As String        ' One output file
    Dim strExt As String = ""       ' Extension
    Dim strLanguage As String = ""  ' Language
    Dim strShort As String = ""     ' Short file name
    Dim arInFile() As String        ' Array of input files
    Dim arCnv() As String
    Dim intPtc As Integer           ' Percentage
    Dim intI As Integer             ' Counter
    Dim intJ As Integer             ' Counter
    Dim bFound As Boolean = False   ' Flag
    Dim bOverwrite As Boolean = False    ' Whether we may overwrite or not
    Dim arCnvType() As String = {"con;conll", "tiger;tig", "folia", "alpino;alp"}
    Dim arExtens() As String = {"con;conll", "tig;tig", "folia.xml", "xml;xml"}

    Try
      ' Make sure type is lower case
      strCnvType = LCase(strCnvType)
      ' Let user select an input and output directory
      With frmDirs
        ' Initialise the directories myself
        .SrcDir = IIf(strLastConvert = "", strWorkDir, strLastConvert)
        .DstDir = .SrcDir
        .DstDir = IIf(strLastConvertDst = "", .SrcDir, strLastConvertDst)
        '' If the directory does not exist, then make it
        'If (Not IO.Directory.Exists(.DstDir)) Then IO.Directory.CreateDirectory(.DstDir)
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK
            ' Get the source and destination directories
            strDirIn = .SrcDir
            strDirOut = .DstDir
          Case Else
            Return False
        End Select
      End With
      ' Save this directory, if it is different from what we have
      If (strLastConvert <> strDirIn) Then
        strLastConvert = strDirIn
        SetTableSetting(tdlSettings, "LastConvertDir", strLastConvert)
        XmlSaveSettings(strSetFile)
      End If
      If (strLastConvertDst <> strDirOut) Then
        strLastConvertDst = strDirOut
        SetTableSetting(tdlSettings, "LastConvertDirDst", strLastConvertDst)
        XmlSaveSettings(strSetFile)
      End If

      ' Initialisations (of language for instance) depends on the conversion type
      Select Case strCnvType
        Case "conll", "con"
          ' Derive the language somehow
          If (Not DeriveLanguage(strDirIn, strLanguage)) Then Return False
          ' Validate language
          Select Case strLanguage
            Case "Spanish", "es"  ' Okay, implemented
              strLanguage = "es"
            Case "Dutch", "nl"    ' Okay, implemented
              strLanguage = "nl"
            Case Else       ' Not implemented
              Status("Not implemented for language [" & strLanguage & "]")
              Return False
          End Select
        Case "tiger", "tig", "folia", "alpino", "alp"
          ' Conversion to .psdx is language-independent
          strLanguage = strEthno
      End Select
      ' Check whether correct directories were selected
      If (IO.Directory.Exists(strDirIn)) AndAlso (IO.Directory.Exists(strDirOut)) Then
        ' Get the conversion index
        bFound = False
        For intJ = 0 To arCnvType.Length - 1
          arCnv = Split(arCnvType(intJ), ";")
          For intI = 0 To arCnv.Length - 1
            If (strCnvType = arCnv(intI)) Then
              strExt = arExtens(intJ) : bFound = True
              Exit For
            End If
          Next intI
          If (bFound) Then Exit For
        Next intJ
        If (Not bFound) Then Return False
        ' Find the input files by looking for the file extensions (in order)...
        arCnv = Split(strExt, ";") : ReDim arInFile(0)
        Select Case strCnvType
          Case "alpino", "alp"
            ' Get all directories under me
            arInFile = IO.Directory.GetDirectories(strDirIn, "*", IO.SearchOption.AllDirectories)
            ' Put all directories that have .xml files under them into a list
            Dim lstDirs As New List(Of String)
            For intI = arInFile.Length - 1 To 0 Step -1
              ' Does this directory contain .xml files?
              If (IO.Directory.GetFiles(arInFile(intI), "*.xml", IO.SearchOption.TopDirectoryOnly).Length > 0) Then
                ' This directory has .xml files, so add the DIRECTORY to the list
                lstDirs.Add(arInFile(intI))
              End If
            Next intI
            ' Convert the list to the array
            arInFile = lstDirs.ToArray
          Case Else
            For intI = 0 To arCnv.Length - 1
              ' Also look in subdirectories!!
              arInFile = IO.Directory.GetFiles(strDirIn, "*." & arCnv(intI), IO.SearchOption.AllDirectories)
              ' Have we found *any* file with this extension at all? If so, we can leave the for-loop
              If (arInFile.Count > 0) Then Exit For
            Next intI
        End Select
        ' Visit all files
        For intI = 0 To arInFile.Count - 1
          ' Show where we are
          intPtc = (intI + 1) * 100 \ arInFile.Count
          Status("Processing " & intPtc & "%", intPtc)
          ' Get the input file name
          strInFile = arInFile(intI)
          ' Find name without dots
          strShort = GetShortFileName(strInFile)
          ' Is this numeric?
          If (IsNumeric(strShort)) Then
            strShort = IO.Directory.GetParent(strInFile).Name & "_" & strShort
            strDirApp = "\" & IO.Directory.GetParent(strInFile).Name
          Else
            ' Determine the output file name, but WITHOUT EXTENSION yet
            strDirApp = Mid(IO.Path.GetDirectoryName(strInFile), strDirIn.Length + 1)
          End If
          strOutFile = strDirOut & strDirApp & "\" & strShort
          ' Check existence of directory
          If (Not IO.Directory.Exists(strDirOut & strDirApp)) Then
            IO.Directory.CreateDirectory(strDirOut & strDirApp)
          End If
          ' Show where we are
          Logging("Producing PSDX " & intI + 1 & "/" & arInFile.Count & " - " & _
                  strShort)
          ' Check existence of final PSDX result
          If (Not bOverwrite) AndAlso (IO.File.Exists(strOutFile & ".psdx")) Then
            ' Ask if overwriting is okay
            If (MsgBox("Can I overwrite files like: [" & strOutFile & ".psdx" & "]?", _
                       MsgBoxStyle.YesNoCancel) <> MsgBoxResult.Yes) Then
              ' No overwriting, so exit 
              Return False
            End If
            ' We can overwrite, so put to true
            bOverwrite = True
          End If
          ' Call the actual conversion function
          Select Case strCnvType
            Case "conll", "con"
              ' Convert from Text to Conll7 to Psdx
              If (Not PrepareOneTxtToPsdx(strInFile, strOutFile, strLanguage, False)) Then
                ' Check for interrupt
                If (bInterrupt) Then Status("Interrupted with error...")
                ' There was an error, so stop
                Return False
              End If
            Case "tiger", "tig"
              If (Not ConvertOneTigerToPsdx(strInFile, strOutFile, strLanguage)) Then
                ' Check for interrupt
                If (bInterrupt) Then Status("Interrupted with error...")
                ' There was an error, so stop
                Return False
              End If
            Case "alpino", "alp"
              ' TODO: implement...
              If (Not ConvertOneAlpinoToPsdx(strInFile, strOutFile, strLanguage)) Then
                ' Check for interrupt
                If (bInterrupt) Then Status("Interrupted with error...")
                ' There was an error, so stop
                Return False
              End If
            Case "folia"
              If (Not ConvertOneFoliaToPsdx(strInFile, strOutFile, strLanguage)) Then
                ' Check for interrupt
                If (bInterrupt) Then Status("Interrupted with error...")
                ' There was an error, so stop
                Return False
              End If
          End Select
          ' Check for interrupt
          If (bInterrupt) Then Status("Interrupted by user...") : Return False
        Next intI
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modConvert/BatchConvertToPsdx error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       LoadDutchCgnInfo()
  ' Goal:       Load appropriate information for the header
  ' History:
  ' 04-02-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function LoadDutchCgnInfo() As Boolean
    Dim strLine As String                   ' One line
    Dim rdThis As IO.StreamReader = Nothing ' File reader
    Dim strCgnInfoFile As String = "L:\CGN\data\meta\text\recordings.txt"

    Try
      ' Should we proceed?
      If (loc_colCgnInfo.Count > 0) Then Return True
      ' Can we access it?
      If (Not IO.File.Exists(strCgnInfoFile)) Then Return False
      ' Read it
      rdThis = New IO.StreamReader(strCgnInfoFile)
      Status("Reading " & strCgnInfoFile & "...")
      While Not (rdThis.EndOfStream)
        ' Read a string
        strLine = rdThis.ReadLine
        ' Validate
        If (strLine <> "") AndAlso (InStr(strLine, vbTab) > 0) AndAlso (InStr(strLine, "SYNTAC") > 0) Then
          ' process it
          loc_colCgnInfo.Add(strLine)
        End If
      End While
      ' Close file
      rdThis.Close()
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modConvert/LoadDutchCgnInfo error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       LoadGermanTigerInfo()
  ' Goal:       Load appropriate information for the header of the Tiger corpus
  ' History:
  ' 04-02-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function LoadGermanTigerInfo() As Boolean
    Dim strLine As String                   ' One line
    Dim arLine() As String                  ' Line divided into parts
    Dim bOrigin As Boolean = False          ' We are within ORIGIN section
    Dim rdThis As IO.StreamReader = Nothing ' File reader
    Dim strTigInfoFile As String = "D:\Download\Software\Treebank\tigercorpus2.1\corpus\tiger_release_aug07.export"
    Dim strTigInfoAlt As String = "D:\Data Files\Corpora\German\Tiger\corpus\tiger_release_aug07.export"

    Try
      ' Should we proceed?
      If (loc_colTigInfo.Count > 0) Then Return True
      ' Can we access it?
      If (Not IO.File.Exists(strTigInfoFile)) Then
        strTigInfoFile = strTigInfoAlt
        If (Not IO.File.Exists(strTigInfoFile)) Then Return False
      End If
      ' Initialise
      loc_colTigInfo.Clear() : loc_colTigMain.Clear()
      ' Read it
      rdThis = New IO.StreamReader(strTigInfoFile)
      Status("Reading " & strTigInfoFile & "...")
      While Not (rdThis.EndOfStream)
        ' Read a string
        strLine = rdThis.ReadLine
        ' Check where we are
        If (bOrigin) Then
          ' Validate
          If (strLine <> "") Then
            If (InStr(strLine, vbTab) > 0) Then
              arLine = Split(strLine, vbTab)
              If (arLine(0) > 0) Then
                ' process it
                loc_colTigMain.Add(arLine(0) & vbTab & arLine(1) & vbTab & Split(arLine(2), " ")(3))
              End If
            ElseIf (InStr(strLine, "#EOT ORIGIN") > 0) Then
              bOrigin = False
            End If
          End If
        Else
          ' Validate
          If (strLine <> "") Then
            If (InStr(strLine, "#BOS ") > 0) Then
              arLine = Split(strLine, " ")
              ' process it
              loc_colTigInfo.Add(arLine(1), arLine(4))
            ElseIf (InStr(strLine, "#BOT ORIGIN") > 0) Then
              bOrigin = True
            End If
          End If
        End If
      End While
      ' Close file
      rdThis.Close()
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modConvert/LoadGermanTigerInfo error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       GetFoliaTextStart()
  ' Goal:       If sentence [strSntNum] starts a new text, then provide some information
  ' History:
  ' 17-02-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function GetFoliaTextStart(ByVal strSntNum As String, ByRef pdxMeta As XmlDocument, ByRef strInfo As String, _
                    ByRef strYear As String, ByRef strEthno As String, ByRef strSubType As String, _
                    ByRef strAuthor As String, ByRef strTitle As String, ByRef strSource As String) As Boolean
    Dim ndxThis As XmlNode      ' One node
    Dim ndxList As XmlNodeList  ' List of nodes
    ' Dim nsDf As XmlNamespaceManager ' Obligtory NS manager
    Dim intI As Integer       ' Counter

    Try
      ' Validate
      If (pdxMeta Is Nothing) Then Return False
      ' Initialise
      strInfo = "" : strYear = "" : strEthno = "" : strSubType = "" : strAuthor = "" : strTitle = "" : strSource = ""
      ' Set the namespace manager
      loc_nsDf = New XmlNamespaceManager(pdxMeta.NameTable)
      loc_nsDf.AddNamespace("df", pdxMeta.DocumentElement.NamespaceURI)
      ' Check what kind it is
      Select Case pdxMeta.FirstChild.Name
        Case "metadata"   ' This is inline folia metadata information
          ' Try to find the different possibilities
          ndxList = pdxMeta.SelectNodes("./descendant::df:meta", loc_nsDf)
          ' Process all <meta> nodes
          For intI = 0 To ndxList.Count - 1
            Select Case ndxList(intI).Attributes("id").Value
              Case "title"
                strTitle = ndxList(intI).InnerText
              Case "language"
                ' This may not completely be the ethno code
                strEthno = ndxList(intI).InnerText
            End Select
          Next intI
          ' Return positively
          Return True
        Case "Components" ' This is accompanying CMD-format .cmdi.xml file information
          ' Find the title
          ndxThis = pdxMeta.SelectSingleNode("./descendant::df:TextTitle", loc_nsDf)
          If (ndxThis IsNot Nothing) Then strTitle = ndxThis.InnerText
          ' Find the source/publisher
          ndxThis = pdxMeta.SelectSingleNode("./descendant::df:Publisher", loc_nsDf)
          If (ndxThis IsNot Nothing) Then strSource = ndxThis.InnerText
          ' Find the year
          ndxThis = pdxMeta.SelectSingleNode("./descendant::df:PublicationDate", loc_nsDf)
          If (ndxThis IsNot Nothing) Then strYear = ndxThis.InnerText
          ' Find the author
          ndxThis = pdxMeta.SelectSingleNode("./descendant::df:Author/df:Name", loc_nsDf)
          If (ndxThis IsNot Nothing) Then strAuthor = ndxThis.InnerText
          ' Find the ethno code
          ndxThis = pdxMeta.SelectSingleNode("./descendant::df:iso-639-3-code", loc_nsDf)
          If (ndxThis IsNot Nothing) Then strEthno = ndxThis.InnerText
          ' Get some additional information - if available
          ' (1) project name
          ndxThis = pdxMeta.SelectSingleNode("./descendant::df:ProjectName", loc_nsDf)
          If (ndxThis IsNot Nothing) Then strInfo &= "Project: " & ndxThis.InnerText & vbCrLf
          ' (2) collection name
          ndxThis = pdxMeta.SelectSingleNode("./descendant::df:CollectionName", loc_nsDf)
          If (ndxThis IsNot Nothing) Then strInfo &= "Collection: " & ndxThis.InnerText & vbCrLf
          ' (3) collection code
          ndxThis = pdxMeta.SelectSingleNode("./descendant::df:CollectionCode", loc_nsDf)
          If (ndxThis IsNot Nothing) Then strInfo &= "Code: " & ndxThis.InnerText & vbCrLf
          ' (3) collection code
          ndxThis = pdxMeta.SelectSingleNode("./descendant::df:TextType", loc_nsDf)
          If (ndxThis IsNot Nothing) Then strInfo &= "TextType: " & ndxThis.InnerText & vbCrLf
          ' Return positively
          Return True
      End Select
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modConvert/GetFoliaTextStart error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       GetTigTextStart()
  ' Goal:       If sentence [strSntNum] starts a new text, then provide some information
  ' History:
  ' 06-02-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function GetTigTextStart(ByVal strSntNum As String, ByRef strInfo As String, ByRef strYear As String, _
                                ByRef strEthno As String, ByRef strSubType As String, ByRef strAuthor As String, _
                                ByRef strTitle As String, ByRef strSource As String) As Boolean
    Dim arLine() As String    ' Content of one line
    Dim intSentNum As Integer ' The sentence as a number
    Dim intMain As Integer    ' Main number
    Dim intPrev As Integer    ' Previous number
    Dim intI As Integer       ' Counter
    Dim intJ As Integer       ' Counter
    Dim bFound As Boolean     ' Positive
    Dim bNew As Boolean       ' New section

    Try
      ' Validate
      If (loc_colTigInfo.Count = 0) OrElse (loc_colTigMain.Count = 0) Then Return False
      ' Get sentence number
      intSentNum = CInt(Mid(strSntNum, 2)) : bFound = False : bNew = False : intPrev = 0
      ' Find the matching section number text
      For intI = 0 To loc_colTigInfo.Count - 1
        ' Is this a new section?
        If (intPrev <> loc_colTigInfo.Exmp(intI)) Then
          ' If (intPrev <> 0) Then Stop
          bNew = True : intPrev = loc_colTigInfo.Exmp(intI)
        Else
          bNew = False
        End If
        ' Get main number, if not zero
        If (loc_colTigInfo.Exmp(intI) <> "0") Then
          intMain = loc_colTigInfo.Exmp(intI)
        End If
        If (loc_colTigInfo.Item(intI) = intSentNum) AndAlso (bNew) Then
          ' Exact match!
          bFound = True
          If (intMain = loc_colTigInfo.Exmp(intI)) Then
            ' We should find the information in loc_colTigMain
            For intJ = 0 To loc_colTigMain.Count - 1
              ' Access the information for this line
              arLine = Split(loc_colTigMain.Item(intJ), vbTab)
              If (arLine(0) = intMain) Then
                ' Translate into needed info
                strSource = "Frankfurter Rundschau " & arLine(2) : strInfo = "Tiger corpus"
                strYear = Left(arLine(2), 4) : strSubType = "G1" : strEthno = "deu" : strTitle = arLine(1)
                Return True
              End If
              '' Translate into needed info
              'strSource = "Frankfurter Rundschau " & arLine(2) : strInfo = "Tiger corpus"
              'strYear = Left(arLine(2), 4) : strSubType = "G1" : strEthno = "deu" : strTitle = arLine(1)
              'Return True
            Next intJ
          Else
            ' There is no information in loc_colTigMain
            strSource = "Frankfurter Rundschau (unknown)" : strInfo = "Tiger corpus"
            strYear = "" : strSubType = "G1" : strEthno = "deu" : strTitle = strSntNum
            Return True
          End If
        ElseIf (loc_colTigInfo.Item(intI) > intSentNum) Then
          Exit For
        End If
      Next intI
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modConvert/GetTigTextStart error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       GetCgnInfo()
  ' Goal:       Find the information associated with text [strTextid]
  ' History:
  ' 04-02-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function GetCgnInfo(ByVal strFile As String, ByRef strInfo As String, ByRef strYear As String, _
                              ByRef strEthno As String, ByRef strSubType As String, ByRef strAuthor As String, _
                              ByRef strTitle As String, ByRef strSource As String) As Boolean
    Dim strLine As String   ' One line
    Dim strTextId As String ' Id
    Dim arLine() As String  ' Line in parts
    Dim intPos As Integer   ' Position
    Dim intI As Integer     ' Counter

    Try
      ' Validate
      If (loc_colCgnInfo.Count = 0) Then Return False
      strTextId = IO.Path.GetFileNameWithoutExtension(strFile)
      ' Walk all elements
      For intI = 0 To loc_colCgnInfo.Count - 1
        ' Get this one
        strLine = loc_colCgnInfo.Item(intI)
        arLine = Split(strLine, vbTab)
        ' Is this the one?
        If (arLine(0) Like "*" & strTextId) Then
          ' Return the info
          strInfo = arLine(5)
          strAuthor = arLine(32) : strTitle = arLine(33)
          strYear = arLine(37) : strSource = arLine(39) & " - " & arLine(40) & " (loc=" & arLine(47) & ")"
          ' Check the language code
          Select Case Left(strTextId, 2)
            Case "fv"
              strEthno = "vls"
            Case "fn"
              strEthno = "nld"
            Case Else
              strEthno = "-"
          End Select
          ' Try to find the subtype for this file
          intPos = InStr(strFile, "\comp-")
          If (intPos > 0) Then
            strSubType = Mid(strFile, intPos + 6, 1)
            If (strEthno = "vls") Then
              strSubType = "VL-" & UCase(strSubType)
            ElseIf (strEthno = "nld") Then
              strSubType = "NL-" & UCase(strSubType)
            End If
          End If
          ' Return positively
          Return True
        End If
      Next intI
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modConvert/GetCgnInfo error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       GetShortFileName()
  ' Goal:       Get a short file name
  ' History:
  ' 17-03-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function GetShortFileName(ByVal strFile As String) As String
    Dim strShort As String = ""

    Try
      ' Initialise
      strShort = IO.Path.GetFileNameWithoutExtension(strFile)
      ' Take away any additional dot extensions
      While (InStr(strShort, ".") > 0)
        strShort = IO.Path.GetFileNameWithoutExtension(strShort)
      End While
      ' Return the result
      Return strShort
    Catch ex As Exception
      ' Show error
      HandleErr("modConvert/GetShortFileName error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  Public Function GetShortFileName(ByVal strFile As String, ByRef arExt() As String) As String
    Dim strShort As String = ""
    Dim intI As Integer         ' Counter
    Dim intIdx As Integer = -1  ' Max fit
    Dim intLen As Integer = 0   ' Max size

    Try
      ' Initialise
      strShort = IO.Path.GetFileNameWithoutExtension(strFile)
      ' Check possible file extensions and take the maximum fit
      For intI = 0 To arExt.Length - 1
        ' Validate
        If (arExt(intI) <> "") Then
          ' Check this fit
          If (Right(strShort, arExt(intI).Length) = arExt(intI)) Then
            ' This is a candidate - get its size and compare it
            If (arExt(intI).Length > intLen) Then
              ' Okay we take this one
              intIdx = intI : intLen = arExt(intI).Length
            End If
          End If
        End If
      Next intI
      ' Did we find an extension?
      If (intIdx < 0) Then
        ' No extension needs to be taken off anymore
      Else
        ' An extension was found, so take it off
        strShort = Left(strShort, strShort.Length - intLen)
      End If
      ' Return the result
      Return strShort
    Catch ex As Exception
      ' Show error
      HandleErr("modConvert/GetShortFileName error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       ReadPsdxHeader()
  ' Goal:       Read the <teiHeader> of a .psdx file
  ' History:
  ' 26-06-2015  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function ReadPsdxHeader(ByVal strSrcFile As String, ByRef pdxDoc As XmlDocument) As Boolean
    Dim rdXml As XmlReader
    Dim rdText As IO.StreamReader
    Dim setThis As New XmlReaderSettings
    Dim strSchemePsdx As String = "psdx.xsd"

    Try
      ' Validate
      If (Not IO.File.Exists(strSrcFile)) Then Return False
      ' Create a new xml document 
      pdxDoc = New XmlDocument
      ' Read using the streamreader
      rdText = New IO.StreamReader(strSrcFile)
      ' Set the schema
      If (Not FindScheme(strSchemePsdx)) Then Return False
      setThis.Schemas.Add("http://www.ru.nl/psdx", strSchemePsdx)
      rdXml = XmlReader.Create(rdText, setThis)
      ' Find the <teiHeader>
      rdXml.ReadToFollowing("teiHeader")
      ' Load it into an xmldocument
      pdxDoc.LoadXml(rdXml.ReadOuterXml)
      ' CLose the file
      rdXml.Close()
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modConvert/ReadPsdxHeader error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       CreatePsdxHeader()
  ' Goal:       Create a body for a new psdx file
  ' History:
  ' 17-03-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function CreatePsdxHeader(ByRef pdxConv As XmlDocument, ByVal strSrcFile As String, _
                                   ByRef strShort As String) As Boolean
    Dim strDefault As String  ' Header default
    Dim rdXml As XmlReader
    Dim rdText As IO.StringReader
    Dim setThis As New XmlReaderSettings
    Dim strSchemePsdx As String = "psdx.xsd"

    Try
      ' Initialise
      strShort = GetShortFileName(strSrcFile)
      ' Make a validating XML reader

      ' Create a new xml document and set the header to initial values
      pdxConv = New XmlDocument
      strDefault = "<TEI><teiHeader><fileDesc>" & _
                       "<publicationStmt distributor='Radboud University Nijmegen' />" & _
                       "<titleStmt title='' author='' />" & _
                       "<sourceDesc bibl='' />" & _
                       "</fileDesc>" & _
                       "<profileDesc>" & _
                       "<langUsage><language ident='' name='' />" & _
                       "<creation original='' manuscript='' subtype='' genre='' /></langUsage>" & _
                       "</profileDesc>" & _
                       "</teiHeader>" & _
                       "<forestGrp File='" & strShort & "'>" & _
                       "</forestGrp></TEI>"
      ' Read from the string
      rdText = New IO.StringReader(strDefault)
      ' Set the schema
      If (Not FindScheme(strSchemePsdx)) Then Return False
      setThis.Schemas.Add("http://www.ru.nl/psdx", strSchemePsdx)
      rdXml = XmlReader.Create(rdText, setThis)
      pdxConv.Load(rdXml)
      ' pdxConv.LoadXml(strDefault)
      ' Make sure the XML functions access the right document
      pdxCurrentFile = pdxConv
      SetXmlDocument(pdxConv)
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modConvert/OneFoliaToPsdxHeader error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       OneFoliaToPsdxHeader()
  ' Goal:       Produce a header for a psdx file that is being made from a folia file
  ' History:
  ' 17-02-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function OneFoliaToPsdxHeader(ByRef pdxConv As XmlDocument, ByRef pdxMeta As XmlDocument, ByVal strSrcFile As String, _
                                        ByVal strLang As String, Optional ByVal strSntNum As String = "") As Boolean
    Dim strInfo As String = ""      ' CGN bibliographical information
    Dim strEthno As String = ""     ' Actual ethnologue code
    Dim strAuthor As String = ""    ' Author information
    Dim strBibl As String = ""      ' Bibliographical source info
    Dim strYear As String = ""      ' Year
    Dim strSubType As String = ""   ' Subtype
    Dim strMetaFile As String = ""  ' File with meta information
    Dim strTitle As String = ""     ' Title
    Dim strOuter As String = ""     ' Content of outer xml
    Dim ndxWork As XmlNode          ' Working node
    Dim strShort As String = ""     ' Short file name
    Dim pdxCmdi As New XmlDocument  ' Reading cmdi
    Dim rdMeta As XmlReader         ' Meta file reader
    Dim settings As XmlReaderSettings = New XmlReaderSettings()

    Try
      ' Create body
      If (Not CreatePsdxHeader(pdxConv, strSrcFile, strShort)) Then Return False
      ' Check of [pdxMeta] contains enough information
      If (pdxMeta Is Nothing) OrElse (pdxMeta.SelectSingleNode("./descendant::meta") Is Nothing) Then
        ' Try to find the cmdi file with the correct information
        If (InStr(strSrcFile, ".folia.") > 0) Then
          strMetaFile = strSrcFile.Replace(".folia.", ".cmdi.")
          ' Is it here?
          If (IO.File.Exists(strMetaFile)) Then
            ' read the metadata
            settings.ProhibitDtd = False
            rdMeta = XmlReader.Create(strMetaFile, settings)
            rdMeta.ReadToFollowing("Components")
            ' Load this one
            'pdxMeta = New XmlDocument
            ' pdxMeta.LoadXml(rdMeta.ReadOuterXml)
            strOuter = rdMeta.ReadOuterXml
            pdxMeta.LoadXml(strOuter)
            'Debug.Print(pdxMeta.SelectSingleNode("./descendant::TextTitle").InnerText)
          End If
        End If
      End If
      ' Get information about this text
      If (Not GetFoliaTextStart(strSntNum, pdxMeta, strInfo, strYear, strEthno, strSubType, strAuthor, strTitle, strBibl)) Then Return False

      ndxWork = pdxConv.SelectSingleNode("./descendant::titleStmt")
      If (ndxWork IsNot Nothing) Then
        ndxWork.Attributes("title").Value = strTitle
        ndxWork.Attributes("author").Value = strAuthor
      End If
      ndxWork = pdxConv.SelectSingleNode("./descendant::sourceDesc")
      If (ndxWork IsNot Nothing) Then ndxWork.Attributes("bibl").Value = strInfo & vbCrLf & strBibl
      ndxWork = pdxConv.SelectSingleNode("./descendant::language")
      If (ndxWork IsNot Nothing) Then
        ndxWork.Attributes("ident").Value = IIf(strEthno = "", strLang, strEthno)
        ndxWork.Attributes("name").Value = EthnoToName(strEthno)
      End If
      ndxWork = pdxConv.SelectSingleNode("./descendant::creation")
      If (ndxWork IsNot Nothing) Then
        ndxWork.Attributes("original").Value = strYear
        ndxWork.Attributes("manuscript").Value = strYear
        ndxWork.Attributes("subtype").Value = strSubType
        ndxWork.Attributes("genre").Value = "-"
      End If
      ndxWork = pdxConv.SelectSingleNode("./descendant::publicationStmt")
      Select Case strLang
        Case "deu"
          If (ndxWork IsNot Nothing) Then ndxWork.Attributes("distributor").Value = "Tiger (Corpus of a German Newspaper)"
        Case "nld"
          If (ndxWork IsNot Nothing) Then ndxWork.Attributes("distributor").Value = "SoNaR"
      End Select
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modConvert/OneFoliaToPsdxHeader error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       OneAlpinoToPsdxHeader()
  ' Goal:       Produce a header for a psdx file that is being made from a folia file
  ' History:
  ' 17-02-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function OneAlpinoToPsdxHeader(ByRef pdxConv As XmlDocument, ByVal strSrcFile As String, _
                                        ByVal strLang As String) As Boolean
    Dim strInfo As String = ""      ' CGN bibliographical information
    Dim strEthno As String = ""     ' Actual ethnologue code
    Dim strAuthor As String = ""    ' Author information
    Dim strBibl As String = ""      ' Bibliographical source info
    Dim strYear As String = ""      ' Year
    Dim strSubType As String = ""   ' Subtype
    Dim strMetaFile As String = ""  ' File with meta information
    Dim strTitle As String = ""     ' Title
    Dim strOuter As String = ""     ' Content of outer xml
    Dim ndxWork As XmlNode          ' Working node
    Dim strShort As String = ""     ' Short file name
    Dim pdxCmdi As New XmlDocument  ' Reading cmdi
    Dim settings As XmlReaderSettings = New XmlReaderSettings()

    Try
      ' Create body
      If (Not CreatePsdxHeader(pdxConv, strSrcFile, strShort)) Then Return False

      ndxWork = pdxConv.SelectSingleNode("./descendant::titleStmt")
      If (ndxWork IsNot Nothing) Then
        ndxWork.Attributes("author").Value = "-"
      End If
      ndxWork = pdxConv.SelectSingleNode("./descendant::sourceDesc")
      If (ndxWork IsNot Nothing) Then ndxWork.Attributes("bibl").Value = "-"
      ndxWork = pdxConv.SelectSingleNode("./descendant::language")
      If (ndxWork IsNot Nothing) Then
        ndxWork.Attributes("ident").Value = IIf(strEthno = "", strLang, strEthno)
        ndxWork.Attributes("name").Value = EthnoToName(strEthno)
      End If
      ndxWork = pdxConv.SelectSingleNode("./descendant::creation")
      If (ndxWork IsNot Nothing) Then
        ndxWork.Attributes("original").Value = Year(Now)
        ndxWork.Attributes("manuscript").Value = Year(Now)
        ndxWork.Attributes("subtype").Value = "-"
        ndxWork.Attributes("genre").Value = "narrative"
      End If
      ndxWork = pdxConv.SelectSingleNode("./descendant::publicationStmt")
      Select Case strLang
        Case "deu"
          If (ndxWork IsNot Nothing) Then ndxWork.Attributes("distributor").Value = "Radboud University Nijmegen"
        Case "nld"
          If (ndxWork IsNot Nothing) Then ndxWork.Attributes("distributor").Value = "Radboud University Nijmegen"
      End Select
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modConvert/OneAlpinoToPsdxheader error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       OnePsdxToFoliaHeader()
  ' Goal:       Produce a header for a folia file that is being made from a psdx file
  ' History:
  ' 01-03-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function OnePsdxToFoliaHeader(ByRef pdxConv As XmlDocument, ByRef pdxSrc As XmlDocument) As Boolean
    Dim strInfo As String = ""      ' CGN bibliographical information
    Dim strEthno As String = ""     ' Actual ethnologue code
    Dim strAuthor As String = ""    ' Author information
    Dim strBibl As String = ""      ' Bibliographical source info
    Dim strYear As String = ""      ' Year
    Dim strSubType As String = ""   ' Subtype
    Dim strMetaFile As String = ""  ' File with meta information
    Dim strTitle As String = ""     ' Title
    Dim strOuter As String = ""     ' Content of outer xml
    Dim ndxWork As XmlNode          ' Working node
    Dim ndxMeta As XmlNode          ' Node in the target document
    Dim ndxMdata As XmlNode         ' Target document <metadata>
    Dim ndxMchild As XmlNode        ' Children of source psdx <metadata>
    Dim strShort As String = ""     ' Short file name
    Dim strSrcFile As String        ' File name
    Dim pdxCmdi As New XmlDocument  ' Reading cmdi
    Dim ndxFor As XmlNode           ' Forest 
    Dim rdMeta As XmlReader         ' Meta file reader
    Dim strValue As String          ' Content
    Dim settings As XmlReaderSettings = New XmlReaderSettings()

    Try
      ' Validate
      If (pdxSrc Is Nothing) Then Return False
      ' Initialise
      ndxFor = pdxSrc.SelectSingleNode("./descendant::forest")
      If (ndxFor Is Nothing) Then Return False
      strSrcFile = ndxFor.Attributes("File").Value
      strShort = GetShortFileName(strSrcFile)
      ' Create a new folia xml document and set the header to initial values
      pdxConv = New XmlDocument
      pdxConv.LoadXml("<?xml version='1.0' encoding='utf-8'?>" & vbCrLf & _
        "<?xml-stylesheet type='text/xsl' href='foliaviewer.xsl'?>" & vbCrLf & _
        "<FoLiA xmlns:xlink='http://www.w3.org/1999/xlink' " & _
        "xmlns='http://ilk.uvt.nl/folia' " & _
        "xml:id='" & strShort & "' version='0.0.1' " & _
        "generator='Cesax'>" & vbCrLf & _
        "<metadata type='native'>" & vbCrLf & _
        "<annotations>" & vbCrLf & _
        " <token-annotation annotator='Cesax' annotatortype='auto'/>" & vbCrLf & _
        " <pos-annotation annotator='Cesax' annotatortype='auto'/>" & vbCrLf & _
        " <syntax-annotation annotator='Cesax' annotatortype='auto'/>" & vbCrLf & _
        "</annotations>" & vbCrLf & _
        "</metadata>" & vbCrLf & _
        "<text xml:id='" & strShort & ".text'>" & vbCrLf & _
        "</text></FoLiA>" & vbCrLf)
      ' Set the namespace manager
      loc_nsDf = New XmlNamespaceManager(pdxConv.NameTable)
      loc_nsDf.AddNamespace("df", pdxConv.DocumentElement.NamespaceURI)
      ' Get <metadata> node
      ndxMdata = pdxConv.SelectSingleNode("./descendant::df:metadata", loc_nsDf)
      ' Make sure we are working on this document
      SetXmlDocument(pdxConv, pdxConv.DocumentElement.NamespaceURI)
      ' Find all the meta information stored in the document
      ndxWork = pdxSrc.SelectSingleNode("./descendant::titleStmt")
      If (ndxWork IsNot Nothing) Then
        ' Create meta.title in target
        ndxMeta = AddXmlChild(ndxMdata, "meta", "id", "title", "attribute")
        ndxMeta.InnerText = GetAttrValue(ndxWork, "title")
        ' Create meta.author in target
        ndxMeta = AddXmlChild(ndxMdata, "meta", "id", "author", "attribute")
        ndxMeta.InnerText = GetAttrValue(ndxWork, "author")
      End If
      ndxWork = pdxSrc.SelectSingleNode("./descendant::sourceDesc")
      If (ndxWork IsNot Nothing) Then
        ' Create meta.bibl in target
        ndxMeta = AddXmlChild(ndxMdata, "meta", "id", "bibl", "attribute")
        ndxMeta.InnerText = GetAttrValue(ndxWork, "bibl")
      End If
      ndxWork = pdxSrc.SelectSingleNode("./descendant::language")
      If (ndxWork IsNot Nothing) Then
        ' Create meta.language_ident in target
        ndxMeta = AddXmlChild(ndxMdata, "meta", "id", "language", "attribute")
        ndxMeta.InnerText = GetAttrValue(ndxWork, "ident")
        ' Create meta.language_name in target
        ndxMeta = AddXmlChild(ndxMdata, "meta", "id", "language_name", "attribute")
        ndxMeta.InnerText = GetAttrValue(ndxWork, "name")
      End If
      ndxWork = pdxSrc.SelectSingleNode("./descendant::creation")
      If (ndxWork IsNot Nothing) Then
        ' Create meta.date_original in target
        ndxMeta = AddXmlChild(ndxMdata, "meta", "id", "date_original", "attribute")
        ndxMeta.InnerText = GetAttrValue(ndxWork, "original")
        ' Create meta.date_manuscript in target
        ndxMeta = AddXmlChild(ndxMdata, "meta", "id", "date_manuscript", "attribute")
        ndxMeta.InnerText = GetAttrValue(ndxWork, "manuscript")
        ' Create meta.subtype in target
        ndxMeta = AddXmlChild(ndxMdata, "meta", "id", "subtype", "attribute")
        ndxMeta.InnerText = GetAttrValue(ndxWork, "subtype")
        ' Create meta.genre in target
        ndxMeta = AddXmlChild(ndxMdata, "meta", "id", "genre", "attribute")
        ndxMeta.InnerText = GetAttrValue(ndxWork, "genre")
      End If
      ndxWork = pdxSrc.SelectSingleNode("./descendant::publicationStmt")
      If (ndxWork IsNot Nothing) Then
        ' Create meta.subtype in target
        ndxMeta = AddXmlChild(ndxMdata, "meta", "id", "distributor", "attribute")
        ndxMeta.InnerText = GetAttrValue(ndxWork, "distributor")
      End If
      ' Add some metadata tags for free
      ndxMeta = AddXmlChild(ndxMdata, "meta", "id", "mayView", "attribute")
      ndxMeta.InnerText = "true"
      ' Find all the meta information stored in the document
      ndxWork = pdxSrc.SelectSingleNode("./descendant::metadata")
      If (ndxWork IsNot Nothing) Then
        ' Walk through all the <meta> children
        ndxMchild = ndxWork.FirstChild
        While (ndxMchild IsNot Nothing)
          ' Is it a Meta?
          If (ndxMchild.Name = "meta") Then
            ' Process this one
            ndxMeta = AddXmlChild(ndxMdata, "meta", "id", ndxMchild.Attributes("id").Value, "attribute")
            ndxMeta.InnerText = MyTrim(ndxMchild.InnerText)
          End If
          ' Try next child
          ndxMchild = ndxMchild.NextSibling
        End While
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modConvert/OnePsdxToFoliaHeader error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       OneFlexToFoliaHeader()
  ' Goal:       Produce a header for a folia file that is being made from a FLEX file
  ' History:
  ' 27-03-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function OneFlexToFoliaHeader(ByRef pdxConv As XmlDocument, ByRef pdxSrc As XmlDocument, _
                                        ByVal strShort As String, ByRef strTxtLang As String) As Boolean
    Dim strInfo As String = ""      ' CGN bibliographical information
    Dim strEthno As String = ""     ' Actual ethnologue code
    Dim strAuthor As String = ""    ' Author information
    Dim strBibl As String = ""      ' Bibliographical source info
    Dim strYear As String = ""      ' Year
    Dim strSubType As String = ""   ' Subtype
    Dim strMetaFile As String = ""  ' File with meta information
    Dim strTitle As String = ""     ' Title
    Dim strOuter As String = ""     ' Content of outer xml
    Dim arLang() As String          ' Choices in languages
    Dim ndxWork As XmlNode          ' Working node
    Dim ndxWord As XmlNodeList      ' List of <word/item/@txt> nodes
    Dim ndxMeta As XmlNode          ' Node in the target document
    Dim ndxMdata As XmlNode         ' Target document <metadata>
    Dim pdxCmdi As New XmlDocument  ' Reading cmdi
    Dim colAsk As New StringColl    ' Question to be asked
    Dim intI As Integer             ' Counter

    Try
      ' Validate
      If (pdxSrc Is Nothing) Then Return False
      ' Create a new folia xml document and set the header to initial values
      pdxConv = New XmlDocument
      pdxConv.LoadXml("<?xml version='1.0' encoding='utf-8'?>" & vbCrLf & _
        "<?xml-stylesheet type='text/xsl' href='foliaviewer.xsl'?>" & vbCrLf & _
        "<FoLiA xmlns:xlink='http://www.w3.org/1999/xlink' " & _
        "xmlns='http://ilk.uvt.nl/folia' " & _
        "xml:id='" & strShort & "' version='0.0.1' " & _
        "generator='Cesax'>" & vbCrLf & _
        "<metadata type='native'>" & vbCrLf & _
        "<annotations>" & vbCrLf & _
        " <token-annotation annotator='Cesax' annotatortype='auto'/>" & vbCrLf & _
        " <pos-annotation annotator='Cesax' annotatortype='auto'/>" & vbCrLf & _
        " <syntax-annotation annotator='Cesax' annotatortype='auto'/>" & vbCrLf & _
        "</annotations>" & vbCrLf & _
        "</metadata>" & vbCrLf & _
        "<text xml:id='" & strShort & ".text'>" & vbCrLf & _
        "</text></FoLiA>" & vbCrLf)
      ' Set the namespace manager
      loc_nsDf = New XmlNamespaceManager(pdxConv.NameTable)
      loc_nsDf.AddNamespace("df", pdxConv.DocumentElement.NamespaceURI)
      ' Get <metadata> node
      ndxMdata = pdxConv.SelectSingleNode("./descendant::df:metadata", loc_nsDf)
      ' Make sure we are working on this document
      SetXmlDocument(pdxConv, pdxConv.DocumentElement.NamespaceURI)
      ' Find all the meta information stored in the document
      ndxWork = pdxSrc.SelectSingleNode("./descendant::interlinear-text/child::item[@type='title' and @lang='en']")
      If (ndxWork IsNot Nothing) Then
        ' Create meta.title in target
        ndxMeta = AddXmlChild(ndxMdata, "meta", "id", "title", "attribute")
        ndxMeta.InnerText = ndxWork.InnerText
      End If
      ' Find out which language should be used in the text
      ndxWord = pdxSrc.SelectNodes("./descendant::word[count(child::item[@type='txt'])>0][1]/child::item[@type = 'txt']")
      ' Is this more than one?
      Select Case ndxWord.Count
        Case 0
          ' We will use the heading instead
          ndxWork = pdxSrc.SelectSingleNode("./descendant::interlinear-text/child::item[@type='title' and @lang!='en']")
          If (ndxWork IsNot Nothing) Then
            ' Create meta.language_ident in target
            ndxMeta = AddXmlChild(ndxMdata, "meta", "id", "language", "attribute")
            ndxMeta.InnerText = GetAttrValue(ndxWork, "lang")
            strTxtLang = ndxMeta.InnerText
          End If
        Case 1
          ' There is just ONE candidate - take it!
          ' Create meta.language_ident in target
          ndxMeta = AddXmlChild(ndxMdata, "meta", "id", "language", "attribute")
          ndxMeta.InnerText = GetAttrValue(ndxWord(0), "lang")
          strTxtLang = ndxMeta.InnerText
        Case Else
          ' User will need to choose
          ReDim arLang(0 To ndxWord.Count - 1)
          colAsk.Add("Please select a language (give its number):")
          For intI = 0 To ndxWord.Count - 1
            arLang(intI) = ndxWord(intI).Attributes("lang").Value
            colAsk.Add(" " & intI + 1 & ": " & arLang(intI))
          Next intI
          ' Get the user to choose a language
          strTxtLang = InputBox(colAsk.Text, "Main language")
          If (strTxtLang = "") OrElse (Not (IsNumeric(strTxtLang))) Then Return False
          strTxtLang = arLang(CInt(strTxtLang) - 1)
          ' Add this choice to the file
          ndxMeta = AddXmlChild(ndxMdata, "meta", "id", "language", "attribute")
          ndxMeta.InnerText = strTxtLang
      End Select
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modConvert/OneFlexToFoliaHeader error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       OneTigerToPsdxHeader()
  ' Goal:       Produce a header for a psdx file that is being made from a tiger file
  ' History:
  ' 27-01-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function OneTigerToPsdxHeader(ByRef pdxConv As XmlDocument, ByVal strSrcFile As String, ByVal strLang As String, _
                                        Optional ByVal strSntNum As String = "") As Boolean
    Dim strInfo As String = ""    ' CGN bibliographical information
    Dim strEthno As String = ""   ' Actual ethnologue code
    Dim strAuthor As String = ""  ' Author information
    Dim strBibl As String = ""    ' Bibliographical source info
    Dim strYear As String = ""    ' Year
    Dim strSubType As String = "" ' Subtype
    Dim strTitle As String = ""   ' Title
    Dim ndxWork As XmlNode        ' Working node
    'Dim ndxForGrp As XmlNode      ' Pointer to <forestGrp>
    Dim strShort As String = ""   ' Short file name

    Try
      ' Create body
      If (Not CreatePsdxHeader(pdxConv, strSrcFile, strShort)) Then Return False
      '' Initialise
      'strShort = IO.Path.GetFileNameWithoutExtension(strSrcFile)
      '' Create a new xml document and set the header to initial values
      'pdxConv = New XmlDocument
      'pdxConv.LoadXml("<TEI><teiHeader><fileDesc>" & _
      '                "<publicationStmt distributor='Radboud University Nijmegen' />" & _
      '                "<titleStmt title='' author='' />" & _
      '                "<sourceDesc bibl='' />" & _
      '                "</fileDesc>" & _
      '                "<profileDesc>" & _
      '                "<langUsage><language ident='' name='' />" & _
      '                "<creation original='' manuscript='' subtype='' /></langUsage>" & _
      '                "</profileDesc>" & _
      '                "</teiHeader>" & _
      '                "<forestGrp File='" & strShort & "'>" & _
      '                "</forestGrp></TEI>")
      '' Make sure the XML functions access the right document
      'pdxCurrentFile = pdxConv
      'SetXmlDocument(pdxConv)
      'ndxForGrp = pdxConv.SelectSingleNode(".//forestGrp")
      ' Action depends on language
      Select Case strLang
        Case "deu"
          If (Not GetTigTextStart(strSntNum, strInfo, strYear, strEthno, strSubType, strAuthor, strTitle, strBibl)) Then Return False
        Case "nld"
          If (Not GetCgnInfo(strSrcFile, strInfo, strYear, strEthno, strSubType, strAuthor, strTitle, strBibl)) Then Return False
      End Select

      ndxWork = pdxConv.SelectSingleNode("./descendant::titleStmt")
      If (ndxWork IsNot Nothing) Then
        ndxWork.Attributes("title").Value = strTitle
        ndxWork.Attributes("author").Value = strAuthor
      End If
      ndxWork = pdxConv.SelectSingleNode("./descendant::sourceDesc")
      If (ndxWork IsNot Nothing) Then ndxWork.Attributes("bibl").Value = strInfo & vbCrLf & strBibl
      ndxWork = pdxConv.SelectSingleNode("./descendant::language")
      If (ndxWork IsNot Nothing) Then
        ndxWork.Attributes("ident").Value = strEthno
        ndxWork.Attributes("name").Value = EthnoToName(strEthno)
      End If
      ndxWork = pdxConv.SelectSingleNode("./descendant::creation")
      If (ndxWork IsNot Nothing) Then
        ndxWork.Attributes("original").Value = strYear
        ndxWork.Attributes("manuscript").Value = strYear
        ndxWork.Attributes("subtype").Value = strSubType
      End If
      ndxWork = pdxConv.SelectSingleNode("./descendant::publicationStmt")
      Select Case strLang
        Case "deu"
          If (ndxWork IsNot Nothing) Then ndxWork.Attributes("distributor").Value = "Tiger (Corpus of a German Newspaper)"
        Case "nld"
          If (ndxWork IsNot Nothing) Then ndxWork.Attributes("distributor").Value = "CGN (Corpus of Spoken Dutch)"
      End Select
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modConvert/OneTigerToPsdxHeader error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       ConvertOneTigerToPsdx()
  ' Goal:       Convert one file in .tig format to destination (psdx)
  ' History:
  ' 27-01-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function ConvertOneTigerToPsdx(ByVal strSrcFile As String, ByVal strDstFile As String, _
      Optional ByVal strLang As String = "none") As Boolean
    Dim strText As String = ""    ' String of file content
    Dim strBrk As String = ""     ' Break symbol used
    Dim strShort As String = ""   ' Short file name
    Dim strTextId As String = ""  ' Text ID
    Dim strType As String = ""    ' Type of eLeaf
    Dim strValue As String = ""   ' Value of <eLeaf>
    Dim strLabel As String = ""   ' Label value
    Dim strFunct As String = ""   ' Function
    Dim strInfo As String = ""    ' CGN bibliographical information
    Dim strEthno As String = ""   ' Actual ethnologue code
    Dim strAuthor As String = ""  ' Author information
    Dim strBibl As String = ""    ' Bibliographical source info
    Dim strYear As String = ""    ' Year
    Dim strSubType As String = "" ' Subtype
    Dim strTitle As String = ""   ' Title
    Dim strOneFile As String = "" ' Name of one output file
    Dim strSntNum As String       ' Sentence number
    Dim ndxWork As XmlNode        ' Working node
    Dim ndxFor As XmlNode         ' Current forest
    Dim ndxForGrp As XmlNode      ' Pointer to <forestGrp>
    Dim pdxSrc As New XmlDocument ' One <s> from the source
    Dim intForestId As Integer = 1  ' ID of the forest
    Dim intEtreeId As Integer = 1 ' ID of <eTree>
    Dim intI As Integer           ' Counter
    Dim intId As Integer = 0      ' DUmmy
    Dim xrdThis As XmlReader = Nothing  ' A reader of the Tiger xml input
    Dim settings As XmlReaderSettings = New XmlReaderSettings()

    Try
      ' Validate
      If (Not IO.File.Exists(strSrcFile)) Then Return False
      ' Start reading
      settings.ProhibitDtd = False
      xrdThis = XmlReader.Create(strSrcFile, settings)

      ' Initialise
      strShort = GetShortFileName(strSrcFile)
      strTextId = strShort.Replace(" ", "")
      loc_intEtreeId = 1 : ndxForGrp = Nothing
      If (Not DoLangInit(strLang)) AndAlso (strLang <> "none") Then
        Logging("Warning: no parsing information found for language [" & strLang & "]")
      End If
      ' If this is NOT german, then make a header
      If (strLang = "deu") Then
        strOneFile = ""
      Else
        ' Create a new xml document and set the header to initial values
        If (Not OneTigerToPsdxHeader(pdxConv, strSrcFile, strLang)) Then Return False
        intForestId = 1
        ' Set the forestgroup entry
        ndxForGrp = pdxConv.SelectSingleNode(".//forestGrp")
        ' Make output file name
        strOneFile = strDstFile & ".psdx"
      End If
      ' Set the point to the first following <s> element
      xrdThis.ReadToFollowing("s")
      ' Walk through the XmlReader of the text
      While (Not xrdThis.EOF)
        ' Read one <s> element into an XmlDocument
        pdxSrc.LoadXml(xrdThis.ReadOuterXml)
        ' Try reordering
        If (Not OneTigerOrder(pdxSrc)) Then Stop
        ' Depending on the language, a new file needs to be started
        If (strLang = "deu") Then
          ' Get the sentence number from the tiger <s> node
          strSntNum = pdxSrc.SelectSingleNode("./descendant::s[1]").Attributes("id").Value
          ' =========== DEBUG ====================
          ' If (strSntNum = "s6004") Then Stop
          ' ======================================
          If (GetTigTextStart(strSntNum, strInfo, strYear, strEthno, strSubType, strAuthor, strTitle, "")) Then
            ' Possibly finish a previous file
            If (strOneFile <> "") Then
              ' Set the current one
              InitCurrentFile()
              ' Re-calculate the <eTree>@Id values for the whole file starting from number 1
              AdaptEtreeId(1)
              ' Save dataset as file
              pdxConv.Save(strOneFile)
              Logging("Saved [" & IO.Path.GetFileName(strOneFile) & "]")
            End If
            ' Set the new output file name
            'strOneFile = strDstFile & "-" & strSntNum & ".psdx"
            ' strOneFile = strDstFile & "-s" & Format(Mid(strSntNum, 2), "00000") & ".psdx"
            strOneFile = IO.Path.GetDirectoryName(strDstFile) & "\fr-s" & Format(CInt(Mid(strSntNum, 2)), "00000") & ".psdx"
            ' strOneFile = strDstFile & "-" & Format(intSect, "0000") & ".psdx"
            ' Create a new xml document and set the header to initial values
            If (Not OneTigerToPsdxHeader(pdxConv, strSntNum, strLang, strSntNum)) Then Return False
            ' Set the forestgroup entry
            ndxForGrp = pdxConv.SelectSingleNode(".//forestGrp")
            ' Start numbering again
            intForestId = 1
          End If
        End If
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
        If (Not OneTigerToPsdxForest(ndxFor, pdxSrc, intI, strLang, False)) Then Return False
        ' Move to the next <s>
        xrdThis.ReadToFollowing("s")
      End While
      ' Set the current one
      InitCurrentFile()
      ' Re-calculate the <eTree>@Id values for the whole file starting from number 1
      AdaptEtreeId(1)
      ' Save dataset as file
      pdxConv.Save(strOneFile)
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/ConvertOneTigerToPsdx error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       XrdAdvance()
  ' Goal:       Skip until the next starting tag
  ' History:
  ' 17-02-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function XrdAdvance(ByRef xrdThis As XmlReader) As Boolean
    Try
      ' Validate
      If (xrdThis Is Nothing) Then Return False
      Do
        xrdThis.Read()
      Loop Until (xrdThis.EOF) OrElse (xrdThis.Name <> "" AndAlso xrdThis.IsStartElement)
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/XrdAdvance error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       ConvertOneAlpinoToPsdx()
  ' Goal:       Convert one directory of Alpino files in .xml format to one destination (psdx) file
  '             The files for individual sentences are numbered numerically from 1...n
  ' History:
  ' 01-07-2015  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function ConvertOneAlpinoToPsdx(ByVal strSrcDir As String, ByVal strDstFile As String, _
      Optional ByVal strLang As String = "none") As Boolean
    Dim strText As String = ""    ' String of file content
    Dim strSrcFile As String = "" ' One source file
    Dim strBrk As String = ""     ' Break symbol used
    Dim strShort As String = ""   ' Short file name
    Dim strTextId As String = ""  ' Text ID
    Dim strType As String = ""    ' Type of eLeaf
    Dim strValue As String = ""   ' Value of <eLeaf>
    Dim strLabel As String = ""   ' Label value
    Dim strFunct As String = ""   ' Function
    Dim strInfo As String = ""    ' CGN bibliographical information
    Dim strEthno As String = ""   ' Actual ethnologue code
    Dim strAuthor As String = ""  ' Author information
    Dim strBibl As String = ""    ' Bibliographical source info
    Dim strYear As String = ""    ' Year
    Dim strSubType As String = "" ' Subtype
    Dim strTitle As String = ""   ' Title
    Dim strOneFile As String = "" ' Name of one output file
    Dim strSectId As String = ""  ' Section Id
    Dim strSent As String = ""    ' Text of sentence
    Dim strComm As String = ""    ' COmment text
    Dim arSrcFile() As String     ' Array of source files
    Dim ndxForGrp As XmlNode      ' Pointer to <forestGrp>
    Dim pdxMeta As New XmlDocument  ' One <metadata> from the source
    Dim intForestId As Integer = 1  ' ID of the forest
    Dim intEtreeId As Integer = 1 ' ID of <eTree>
    Dim intI As Integer           ' Counter
    Dim intId As Integer = 0      ' DUmmy
    Dim xrdThis As XmlReader = Nothing  ' A reader of the Tiger xml input
    Dim settings As XmlReaderSettings = New XmlReaderSettings()

    Try
      ' Validate
      If (Not IO.Directory.Exists(strSrcDir)) Then Return False

      ' Initialise
      strShort = GetShortFileName(strSrcDir)
      ' Determine the text ID from short
      strTextId = strShort.Replace(" ", "")
      loc_intEtreeId = 1 : ndxForGrp = Nothing
      If (Not DoLangInit(strLang)) AndAlso (strLang <> "none") Then
        Logging("Warning: no parsing information found for language [" & strLang & "]")
      End If
      ' Initialisations
      strOneFile = "" : strSectId = ""
      ' Create the name of the output file
      strOneFile = strDstFile & ".psdx"
      ' Create a new xml document and set the header to initial values
      If (Not OneAlpinoToPsdxHeader(pdxConv, strSrcFile, strLang)) Then Return False
      ' Initialise forest id for this file
      intForestId = 1
      ' Make sure current file is set correctly
      pdxCurrentFile = pdxConv
      ' Make sure correct document is set
      SetXmlDocument(pdxConv, "")
      ' Set the forestgroup entry
      ndxForGrp = pdxConv.SelectSingleNode(".//forestGrp")

      ' Process all numerical .xml files in this directory
      arSrcFile = IO.Directory.GetFiles(strSrcDir, "*.xml", IO.SearchOption.TopDirectoryOnly)
      For intI = 0 To arSrcFile.Length - 1
        strSrcFile = arSrcFile(intI)
        ' Start reading
        settings.DtdProcessing = DtdProcessing.Ignore
        xrdThis = XmlReader.Create(strSrcFile, settings)
        ' Walk through the XmlReader of the text
        While (Not xrdThis.EOF)
          ' Read the start of this text
          xrdThis.ReadToFollowing("alpino_ds")
          ' Read until we get to <node>
          While (Not xrdThis.EOF) AndAlso (XrdAdvance(xrdThis)) AndAlso (xrdThis.Name <> "node")
            Application.DoEvents()
          End While
          ' Double check
          If (Not xrdThis.EOF) AndAlso (xrdThis.IsStartElement()) AndAlso (xrdThis.Name = "node") Then
            ' <node> elements may contain other <node> elements
            Dim pdxSrc As New XmlDocument ' One <s> from the source
            Dim pdxCom As New XmlDocument ' One <comments> from the source

            ' Read the nested <node> for this sentence
            pdxSrc.LoadXml(xrdThis.ReadOuterXml)
            ' Now we are looking for <sentence> or <comment>
            While (Not xrdThis.EOF) AndAlso (XrdAdvance(xrdThis))
              ' Is this a start element?
              If (Not xrdThis.EOF) AndAlso (xrdThis.IsStartElement) Then
                ' Check what it is
                Select Case xrdThis.Name
                  Case "sentence"
                    ' Read the sentence
                    strSent = xrdThis.ReadInnerXml
                  Case "comments"
                    ' read the comments
                    pdxCom.LoadXml(xrdThis.ReadOuterXml)
                  Case Else
                    ' Read until the following element
                    XrdAdvance(xrdThis)
                End Select
              End If
            End While
            ' Process this sentence
            If (Not OneAlpinoToPsdxForest(ndxForGrp, pdxSrc, intForestId, strTextId, strShort, strSectId, strSent)) Then Return False
          Else
            ' Missing <node> error
            Logging("The alpino-format file has no <node> block")
          End If
        End While
      Next intI
      ' Set the current one
      InitCurrentFile()
      ' Re-calculate the <eTree>@Id values for the whole file starting from number 1
      AdaptEtreeId(1)
      ' Save dataset as file
      pdxConv.Save(strOneFile)
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/ConvertOneAlpinoToPsdx error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function

  '----------------------------------------------------------------------------------------
  ' Name:       ConvertOneFoliaToPsdx()
  ' Goal:       Convert one file in .folia.xml format to one or more destination (psdx) file(s)
  ' History:
  ' 17-02-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function ConvertOneFoliaToPsdx(ByVal strSrcFile As String, ByVal strDstFile As String, _
      Optional ByVal strLang As String = "none") As Boolean
    Dim strText As String = ""    ' String of file content
    Dim strBrk As String = ""     ' Break symbol used
    Dim strShort As String = ""   ' Short file name
    Dim strTextId As String = ""  ' Text ID
    Dim strType As String = ""    ' Type of eLeaf
    Dim strValue As String = ""   ' Value of <eLeaf>
    Dim strLabel As String = ""   ' Label value
    Dim strFunct As String = ""   ' Function
    Dim strInfo As String = ""    ' CGN bibliographical information
    Dim strEthno As String = ""   ' Actual ethnologue code
    Dim strAuthor As String = ""  ' Author information
    Dim strBibl As String = ""    ' Bibliographical source info
    Dim strYear As String = ""    ' Year
    Dim strSubType As String = "" ' Subtype
    Dim strTitle As String = ""   ' Title
    Dim strOneFile As String = "" ' Name of one output file
    Dim strSntNum As String       ' Sentence number
    Dim strSectId As String = ""  ' Section Id
    Dim ndxWork As XmlNode        ' Working node
    Dim ndxFor As XmlNode         ' Current forest
    Dim ndxForGrp As XmlNode      ' Pointer to <forestGrp>
    Dim pdxMeta As New XmlDocument  ' One <metadata> from the source
    Dim intForestId As Integer = 1  ' ID of the forest
    Dim intEtreeId As Integer = 1 ' ID of <eTree>
    Dim intI As Integer           ' Counter
    Dim intId As Integer = 0      ' DUmmy
    Dim xrdThis As XmlReader = Nothing  ' A reader of the Tiger xml input
    Dim settings As XmlReaderSettings = New XmlReaderSettings()

    Try
      ' Validate
      If (Not IO.File.Exists(strSrcFile)) Then Return False
      ' Start reading
      settings.ProhibitDtd = False
      xrdThis = XmlReader.Create(strSrcFile, settings)

      ' Initialise
      strShort = GetShortFileName(strSrcFile)
      'strShort = IO.Path.GetFileNameWithoutExtension(strSrcFile)
      '' Take away any additional dot extensions
      'While (InStr(strShort, ".") > 0)
      '  strShort = IO.Path.GetFileNameWithoutExtension(strShort)
      'End While
      ' Determine the text ID from short
      strTextId = strShort.Replace(" ", "")
      loc_intEtreeId = 1 : ndxForGrp = Nothing
      If (Not DoLangInit(strLang)) AndAlso (strLang <> "none") Then
        Logging("Warning: no parsing information found for language [" & strLang & "]")
      End If
      ' Initialisations
      strOneFile = "" : strSectId = ""

      ' Walk through the XmlReader of the text
      While (Not xrdThis.EOF)
        ' Read the start of this text
        xrdThis.ReadToFollowing("FoLiA")
        ' read the "id" attribute, which is the name of the text
        strSntNum = xrdThis.Item("xml:id")
        ' Read the <metadata> component
        xrdThis.ReadToFollowing("metadata")
        ' NB: this can only be properly read with the namespace manager in [OneFoliaToPsdxHeader]...
        pdxMeta.LoadXml(xrdThis.ReadOuterXml)
        ' Create a new xml document and set the header to initial values
        If (Not OneFoliaToPsdxHeader(pdxConv, pdxMeta, strSrcFile, strLang, strSntNum)) Then Return False
        ' Make sure current file is set correctly
        pdxCurrentFile = pdxConv
        ' Make sure correct document is set
        SetXmlDocument(pdxConv, "")
        ' Create the name of the output file
        If (IO.Path.GetExtension(strDstFile) = ".psdx") Then
          strOneFile = strDstFile
        Else
          strOneFile = IO.Path.GetDirectoryName(strDstFile) & "\" & strSntNum & ".psdx"
        End If
        ' Initialise forest id for this file
        intForestId = 1
        ' Set the forestgroup entry
        ndxForGrp = pdxConv.SelectSingleNode(".//forestGrp")
        ' Read until we get to <text>
        While (Not xrdThis.EOF) AndAlso (XrdAdvance(xrdThis)) AndAlso (xrdThis.Name <> "text")
          Application.DoEvents()
        End While
        ' Double check
        If (Not xrdThis.EOF) AndAlso (xrdThis.IsStartElement()) AndAlso (xrdThis.Name = "text") Then
          ' <text> elements may contain:
          ' (1) <gap>
          ' (2) <div>  -- containing p-s-w-t
          ' (3) <p>    -- containing   s-w-t
          ' (4) <head> -- containing   s-w-t
          ' Find out what is coming
          While (Not xrdThis.EOF) AndAlso (XrdAdvance(xrdThis))
            ' Is this a start element?
            If (Not xrdThis.EOF) AndAlso (xrdThis.IsStartElement) Then
              ' Check what it is
              Select Case xrdThis.Name
                Case "gap"
                  ' Gaps need to be skipped
                  xrdThis.ReadOuterXml()
                  ' Clear section ID
                  strSectId = ""
                Case "div"
                  ' Possibly note that we are starting a <div>
                  strSectId = xrdThis.Item("xml:id")
                Case "p", "head"
                  ' First line in paragraph will be regarded as head
                  If (xrdThis.ReadToFollowing("s")) Then
                    Dim pdxSrc As New XmlDocument ' One <s> from the source

                    ' Read this sentence as first in a paragraph
                    pdxSrc.LoadXml(xrdThis.ReadOuterXml)
                    ' Process this sentence
                    If (Not OneFoliaToPsdxForest(ndxForGrp, pdxSrc, intForestId, strTextId, strShort, strSectId)) Then Return False
                    ' Continue reading other sentences
                    While (xrdThis.ReadToNextSibling("s"))
                      Dim pdxSrc2 As New XmlDocument ' One <s> from the source

                      ' Read this sentence
                      pdxSrc2.LoadXml(xrdThis.ReadOuterXml)
                      ' Process this sentence
                      If (Not OneFoliaToPsdxForest(ndxForGrp, pdxSrc2, intForestId, strTextId, strShort, strSectId)) Then Return False
                    End While
                  End If
                  ' Read skip until the next start of an element
                  XrdAdvance(xrdThis)
                  ' Read the next paragraphs
                  While (xrdThis.Name = "p")
                    ' First line in paragraph
                    If (xrdThis.ReadToFollowing("s")) Then
                      Dim pdxSrc As New XmlDocument ' One <s> from the source

                      ' Read this sentence as first in a paragraph
                      pdxSrc.LoadXml(xrdThis.ReadOuterXml)
                      ' Process this sentence
                      If (Not OneFoliaToPsdxForest(ndxForGrp, pdxSrc, intForestId, strTextId, strShort, strSectId)) Then Return False
                      ' Continue reading other sentences in this paragraph
                      While (xrdThis.ReadToNextSibling("s"))
                        Dim pdxSrc2 As New XmlDocument ' One <s> from the source

                        ' Read this sentence
                        pdxSrc2.LoadXml(xrdThis.ReadOuterXml)
                        ' Process this sentence
                        If (Not OneFoliaToPsdxForest(ndxForGrp, pdxSrc2, intForestId, strTextId, strShort, strSectId)) Then Return False
                      End While
                    End If
                    ' Read skip until the next start of an element
                    XrdAdvance(xrdThis)
                  End While
                  If (Not xrdThis.EOF) Then
                    ' Find out what is following now...
                    Select Case xrdThis.Name
                      Case "gap"
                        ' Gaps should be skipped
                        xrdThis.Skip()
                        ' Read skip until the next start of an element
                        XrdAdvance(xrdThis)
                      Case Else
                        ' Not sure what to do with other elements...
                        Stop

                    End Select

                  End If
                Case Else
                  ' Read until the following element
                  XrdAdvance(xrdThis)
              End Select
            End If
          End While
        Else
          ' Missing <text> error
          Logging("The folia-format file has no <text> block")
        End If

      End While
      ' Set the current one
      InitCurrentFile()
      ' Re-calculate the <eTree>@Id values for the whole file starting from number 1
      AdaptEtreeId(1)
      ' Save dataset as file
      pdxConv.Save(strOneFile)
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/ConvertOneFoliaToPsdx error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function

  '----------------------------------------------------------------------------------------
  ' Name:       OneAlpinoToPsdxForest()
  ' Goal:       Process one alpino top-level <node> element into a psdx <forest>
  ' History:
  ' 01-07-2015  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function OneAlpinoToPsdxForest(ByRef ndxForGrp As XmlNode, ByRef pdxSrc As XmlDocument, _
    ByRef intForestId As Integer, ByVal strTextId As String, ByVal strShort As String, _
    ByRef strSectId As String, ByVal strSent As String) As Boolean
    Dim ndxThis As XmlNode      ' Working node
    Dim ndxFor As XmlNode       ' Forest
    Dim ndxEtree As XmlNode     ' Etree node
    Dim ndxLeaf As XmlNode      ' Eleaf node
    Dim ndxList As XmlNodeList  ' List of items
    Dim ndxWord As XmlNode      ' Word in source
    Dim strWord As String       ' Word
    Dim strEng As String        ' The <t> node taken for the english BT
    Dim strPos As String        ' POS
    Dim strType As String       ' Type 
    Dim strValue As String      ' Valure
    Dim intIchCounter As Integer ' *ICH* node counter
    Dim intI As Integer         ' Counter
    Dim nmsAlpinoa As XmlNamespaceManager ' Needed to browse through [pdxSrc]

    Try
      ' Validate
      If (ndxForGrp Is Nothing) OrElse (pdxSrc Is Nothing) Then Return False
      ' Set the namespace manager for this particular [pdxSrc]
      nmsAlpinoa = New XmlNamespaceManager(pdxSrc.NameTable)
      nmsAlpinoa.AddNamespace("alpino", pdxSrc.DocumentElement.NamespaceURI)
      ' Create a new <forest> element and add its properties
      ndxFor = AddXmlChild(ndxForGrp, "forest", "forestId", intForestId, "attribute", _
                           "TextId", strTextId, "attribute", _
                           "File", strShort & ".psdx", "attribute", _
                           "Location", strTextId & "." & Format(intForestId, "0000"), "attribute")
      ' Do we have a section id?
      If (strSectId <> "" OrElse intForestId = 1) Then
        ' Add it
        AddAttribute(ndxFor, "Section", strSectId)
        ' Reset id
        strSectId = ""
      End If
      intForestId += 1 : strEng = "" : intIchCounter = 0
      ' Add the <divs>: one for the original (org) and one for English (to be added)
      ndxThis = AddXmlChild(ndxFor, "div", "lang", "org", "attribute", "seg", strSent, "child")
      ndxThis = AddXmlChild(ndxFor, "div", "lang", "eng", "attribute", "seg", "", "child")
      ' Get all the <node> elements with attribute @word, but sorted according to @begin
      Dim ndSorted = pdxSrc.SelectNodes("./descendant-or-self::alpino:node[@word]", _
        nmsAlpinoa).Cast(Of XmlNode)().OrderBy(Function(ndNode) CInt(ndNode.Attributes("begin").Value))
      ' Walk through the nodes (which are now ordered)
      intI = 1
      For Each ndxThis In ndSorted
        ' ========== DEbug
        ' Debug.Print(ndOne.Attributes("id").Value & "-" & ndOne.Attributes("begin").Value)
        ' ===================
        ' Default values
        strType = "Vern"
        ' Check for potential problem
        If (ndxThis Is Nothing) Then Stop : Return False
        ' Get the word into the variable [strValue]
        strValue = ndxThis.Attributes("word").Value
        ' Check if this is punctuation or not
        If (ndxThis.Attributes("pos").Value = "punct") Then
          ' we have punctuation
          strType = "Punct" : strPos = ""
          ' Determine the value
          Select Case strValue
            Case "(", "["
              strValue = "("
            Case ")", "]"
              strValue = ")"
            Case ".", "!", "?"
              strValue = "."
            Case ",", ":", ";", "-"
              strValue = ","
            Case """", "'", "«", "»", "''", "``", "/"
              strValue = """"
          End Select
        Else
          ' We have something else - determin the POS
          strPos = ndxThis.Attributes("pt").Value & "-" & ndxThis.Attributes("rel").Value
        End If
        ' We are CREATING, not synchronizing: add one XML child under [ndxFor]
        If (strType = "Punct") AndAlso (strValue <> "") Then
          ndxEtree = AddEtreeChild(ndxFor, loc_intEtreeId, strValue, 0, 0)
        Else
          ndxEtree = AddEtreeChild(ndxFor, loc_intEtreeId, strPos, 0, 0)
        End If
        ' Keep track of the <eTree> id
        loc_intEtreeId += 1
        ' Add <eLeaf> to the <eTree> node
        ndxLeaf = AddEleafChild(ndxEtree, strType, strValue, 0, 0)
        ' The @n attribute is the number of the word inside the current sentence (=forest)
        ndxLeaf.Attributes("n").Value = intI : intI += 1
        ' Walk through all the attributes
        Dim attrThis As Xml.XmlAttribute
        For Each attrThis In ndxThis.Attributes
          ' Check if it is okay
          Select Case attrThis.Name
            Case "pt", "rel", "word"
              ' No action
            Case "lemma"
              ' Add this separately
              AddFeature(pdxCurrentFile, ndxEtree, "M", "l", attrThis.Value)
            Case Else
              ' add the name and the value
              AddFeature(pdxCurrentFile, ndxEtree, "alp", attrThis.Name, attrThis.Value)
          End Select
        Next attrThis
      Next ndxThis
      ' Merk all non-word nodes 
      ndxList = pdxSrc.SelectNodes("./descendant-or-self::alpino:node[not(@word)]", nmsAlpinoa)
      For intI = 0 To ndxList.Count - 1
        AddAttribute(pdxSrc, ndxList(intI), "done", "no")
      Next intI
      SetXmlDocument(pdxCurrentFile)
      ' Process the words one-by-one
      ndxList = ndxFor.SelectNodes("./child::eTree")
      For Each ndxThis In ndxList
        ' Find this node in the Alpino source [pdxSrc]
        ndxWord = pdxSrc.SelectSingleNode("./descendant::alpino:node[@id='" & _
          GetFeature(ndxThis, "alp", "id") & "']", nmsAlpinoa)
        ' Get the parent of this node
        Dim ndxParent As XmlNode
        Dim strLabel As String
        Dim bWordDone As Boolean = False
        ' ========== debug
        strWord = ndxWord.Attributes("word").Value
        ' ======================

        ' Start going upwards in the Alpino source tree
        ndxParent = ndxWord.ParentNode
        ' Words that are directly attached to the 'top' get a special treatment
        If (ndxParent.Attributes("cat").Value = "top") Then
          ' Those with a "top" parent get a special treatment
          ' They will be treated further down
          AddXmlAttribute(pdxCurrentFile, ndxThis, "later", "true")
        Else
          ' Nodes initially get attached to the <forest> element
          ndxEtree = ndxThis
          ' Walk upwards, but keep in <node> structure
          While (ndxParent IsNot Nothing AndAlso (ndxParent.Name = "node") AndAlso _
                 (ndxParent.Attributes("done").Value = "no") AndAlso (ndxParent.Attributes("cat").Value <> "top"))
            ' This node has not yet been processed
            ' (1) Determine label
            strLabel = ndxParent.Attributes("cat").Value
            If (Not DoLike(ndxParent.Attributes("rel").Value, "hd|--")) Then strLabel &= "-" & ndxParent.Attributes("rel").Value
            ' (2) Insert a node (and [ndxEtree] becomes this *new* node!
            If (Not eTreeInsertLevel(ndxThis, ndxEtree)) Then Return False
            loc_intEtreeId += 1
            ' (3) set the label of this new node
            ndxEtree.Attributes("Label").Value = strLabel
            ndxEtree.Attributes("Id").Value = loc_intEtreeId
            ' (4) TODO: copy attributes to this node
            For intI = 0 To ndxParent.Attributes.Count - 1
              Dim strAttrName As String = ndxParent.Attributes(intI).Name
              Dim strAttrValue As String = ndxParent.Attributes(intI).Value
              If (strAttrName <> "cat") Then AddFeature(pdxCurrentFile, ndxEtree, "alp", strAttrName, strAttrValue)
            Next intI
            ' (x) Indicate in the Alpino tree that this node has been processed
            ndxParent.Attributes("done").Value = "yes"
            ' Go one up
            ndxParent = ndxParent.ParentNode
            ndxThis = ndxThis.ParentNode
          End While
          ' Check situation: is parent a "done" one?
          If (ndxParent IsNot Nothing AndAlso (ndxParent.Name = "node") AndAlso _
                 (ndxParent.Attributes("done").Value = "yes")) Then
            ' We need to append [ndxEtree] as child under the PSDX copy of the ALPINO node
            ' (1) Find the PSDX copy
            Dim ndxNewParent As XmlNode = ndxFor.SelectSingleNode( _
              "./descendant::eTree[child::fs/child::f[@name='id' and @value='" & ndxParent.Attributes("id").Value & "']]")
            ' (2) Double checking
            If (ndxNewParent Is Nothing) Then
              HandleErr("modConvert/OneAlpinoToPsdxForest warning: cannot find PSDX equivalent of ALPINO node")
              Stop
            End If
            ' (3) Check for well-formedness: 
            '     - Linear precedence: the last <eLeaf> child under my prospective PSDX parent
            '       must be my immediately linearly preceding neighbour
            '       Means: @end of rightmost ndxNewParent descendant == @begin of leftmost ndxEtree one
            '     - The word preceding me may not come after my parent-to-be in PSDX
            Dim intNprec As Integer = ndxEtree.SelectSingleNode("./descendant::eLeaf[1]").Attributes("n").Value
            ' Dim ndxTest As XmlNode = ndxNewParent.SelectSingleNode("./following::eLeaf[@n=" & intNprec-1 & "]")
            Dim ndxTest As XmlNode = ndxNewParent.SelectSingleNode( _
              "./following::eLeaf[@n<" & intNprec & " and parent::eTree[not(@later)]]")
            If (ndxTest IsNot Nothing) Then
              ' (4) The order would be corrupted
              ' (4.1) Create a new node under the 'ndxNewParent' we determined
              Dim ndxIchNode As XmlNode = Nothing
              Dim ndxIchLeaf As XmlNode = Nothing
              intIchCounter += 1
              eTreeAdd(ndxNewParent, ndxIchNode, "child")
              ' (4.2) This node gets the same @Label as I have, but a *ICH*-n child
              ndxIchNode.Attributes("Label").Value = ndxEtree.Attributes("Label").Value
              eTreeAdd(ndxIchNode, ndxIchLeaf, "leaf")
              ndxIchLeaf.Attributes("Type").Value = "Star"
              ndxIchLeaf.Attributes("Text").Value = "*ICH*-" & intIchCounter
              ' (4.3) The [ndxEtree] get a "-n" attached to my label
              ndxEtree.Attributes("Label").Value = ndxEtree.Attributes("Label").Value & "-" & intIchCounter
              ' (4.4) The [ndxEtree] must append as child under the parent of the node in PSDX
              '         that has the @n minus 1
              Dim intN = ndxEtree.SelectSingleNode("./descendant::eLeaf[1]").Attributes("n").Value
              ' (4.5) So find our new parent-to-be
              Dim intSubtract As Integer = 0
              Do
                intSubtract -= 1
                ndxNewParent = ndxFor.SelectSingleNode( _
                  "./descendant::eTree[child::eLeaf[@n = " & intN + intSubtract & "]]")
                ' Double checking
                If (ndxNewParent Is Nothing) Then
                  HandleErr("modConvert/OneAlpinoToPsdxForest warning: could not find leaf with @n=" & intN - 1)
                  Stop
                End If
                ' Debugging ======
                If (ndxNewParent.ParentNode.Name = "forest") Then Stop
                ' ================
              Loop Until (ndxNewParent.ParentNode.Name <> "forest")
              ' (4.6) we need to go one step upwards and check it still is an <eTree>
              ndxNewParent = ndxNewParent.ParentNode
              If (ndxNewParent Is Nothing OrElse ndxNewParent.Name <> "eTree") Then
                HandleErr("modConvert/OneAlpinoToPsdxForest warning: no suitable <eTree> grandparent for leaf with @n=" & intN - 1)
                Stop
              End If
            End If
            ' (5) The order is okay, so append [ndxEtree] as child
            ndxNewParent.AppendChild(ndxEtree)
          End If
        End If
      Next ndxThis
      ' Find all the end-nodes with an index in ALPINO
      ndxList = pdxSrc.SelectNodes("./descendant::node[@index > 0 and count(child::node) = 0 and not(@cat) and not(@lcat)]", nmsAlpinoa)
      For Each ndxWord In ndxList
        ' Get my index value
        Dim intIndex As Integer = ndxWord.Attributes("index").Value
        ' Determine the referential counter
        Dim intRefIndex As Integer = intIchCounter + intIndex
        ' Get the node in PSDX to which I should point
        Dim ndxIndexTarget As XmlNode = ndxFor.SelectSingleNode( _
          "./descendant::eTree[child::fs/child::f[@name='index' and @value='" & intIndex & "']]")
        If (ndxIndexTarget Is Nothing) Then
          HandleErr("modConvert/OneAlpinoToPsdxForest warning: could not find target for index=" & intIndex)
          Stop
        Else
          Dim ndxIchNode As XmlNode = Nothing
          Dim ndxIchLeaf As XmlNode = Nothing
          ' Adapt the label of this constituent
          ndxIndexTarget.Attributes("Label").Value &= "-" & intRefIndex
          ' Determine my position by getting my parent, preceding sibling and following sibling
          Dim intParIdx As Integer = -1 : Dim intBefIdx As Integer = -1 : Dim intAftIdx As Integer = -1
          Dim ndxPar As XmlNode = ndxWord.ParentNode : If (ndxPar IsNot Nothing) Then intParIdx = ndxPar.Attributes("id").Value
          ' Validate
          If (ndxPar Is Nothing) Then
            HandleErr("modConvert/OneAlpinoToPsdxForest warning: could not find parent of index=" & intIndex)
            Stop
          Else
            ' Find the equivalent parent within Psdx
            ndxPar = ndxFor.SelectSingleNode("./descendant::eTree[child::fs/child::f[@name='id' and @value='" & intParIdx & "']]")
            ' Find the preceding sibling and following sibling within ALP
            Dim ndxBef As XmlNode = ndxWord.PreviousSibling : If (ndxBef IsNot Nothing) Then intBefIdx = ndxBef.Attributes("id").Value
            Dim ndxAft As XmlNode = ndxWord.NextSibling : If (ndxAft IsNot Nothing) Then intAftIdx = ndxAft.Attributes("id").Value
            ' Try find the corresponding nodes in PSDX
            If (intBefIdx >= 0) Then ndxBef = ndxFor.SelectSingleNode("./descendant::eTree[child::fs/child::f[@name='id' and @value='" & intBefIdx & "']]")
            If (intAftIdx >= 0) Then ndxAft = ndxFor.SelectSingleNode("./descendant::eTree[child::fs/child::f[@name='id' and @value='" & intAftIdx & "']]")
            ' Find out where a node should be created in the target PSDX
            If (ndxBef Is Nothing) Then
              If (ndxAft Is Nothing) Then
                ' No before and no after --> only child
                eTreeAdd(ndxPar, ndxIchNode, "child")
              Else
                ' Insert a new child before ndxAft
                eTreeAdd(ndxAft, ndxIchNode, "left")
              End If
            Else
                ' Insert a new child after ndxBef
              eTreeAdd(ndxBef, ndxIchNode, "right")
            End If
            ' Double check
            If (ndxIchNode IsNot Nothing) Then
              ' Adapt the label of the new node
              Dim sLabel As String = ndxIndexTarget.Attributes("Label").Value
              Dim intIdx As Integer = InStr(sLabel, "-")
              If (intIdx > 0) Then sLabel = Left(sLabel, intIdx - 1)
              ndxIchNode.Attributes("Label").Value = sLabel & "-" & ndxWord.Attributes("rel").Value
              ' Copy the attributes from source to target
              For intI = 0 To ndxWord.Attributes.Count - 1
                Dim strAttrName As String = ndxWord.Attributes(intI).Name
                Dim strAttrValue As String = ndxWord.Attributes(intI).Value
                If (strAttrName <> "cat") Then AddFeature(pdxCurrentFile, ndxIchNode, "alp", strAttrName, strAttrValue)
              Next intI
              ' Add an eLeaf node to the ICH
              eTreeAdd(ndxIchNode, ndxIchLeaf, "leaf")
              ndxIchLeaf.Attributes("Type").Value = "Star"
              ndxIchLeaf.Attributes("Text").Value = "*T*-" & intRefIndex
            End If
          End If
        End If
      Next ndxWord
      ' Now do the words marked 'later'
      ndxList = ndxFor.SelectNodes("./child::eTree[@later]")
      For Each ndxThis In ndxList
        ' Find this node in the Alpino source [pdxSrc]
        ndxWord = pdxSrc.SelectSingleNode("./descendant::alpino:node[@id='" & _
          GetFeature(ndxThis, "alp", "id") & "']", nmsAlpinoa)
        ' Find the words between which it has to be plugged
        Dim intN = ndxThis.SelectSingleNode("./child::eLeaf").Attributes("n").Value
        Dim ndxBef As XmlNode = ndxFor.SelectSingleNode("./descendant::eTree[child::eLeaf[@n=" & intN - 1 & "]]")
        Dim ndxAft As XmlNode = ndxFor.SelectSingleNode("./descendant::eTree[child::eLeaf[@n=" & intN + 1 & "]]")
        ' Check which situation we are in
        If (ndxBef Is Nothing OrElse ndxBef.Name = "forest") Then
          ' Must plug in before [ndxAft]
          ' (1) Get first eTree child under forest
          Dim ndxFirst = ndxFor.SelectSingleNode("./child::eTree[not(@later)][1]")
          ' (2) Prepend before this one
          ndxFirst.PrependChild(ndxThis)
        ElseIf (ndxAft Is Nothing OrElse ndxAft.Name = "forest") Then
          ' Must plug in after [ndxBef]
          ' (1) Get last eTree child under forest
          Dim ndxLast = ndxFor.SelectSingleNode("./child::eTree[not(@later)][last()]")
          ' (2) Append after this one
          ndxLast.AppendChild(ndxThis)
        Else
          ' Must plug in between [Bef - Aft]
          Dim strCond As String = ""  ' "[child::fs/child::f[@name = 'cat']/@value != 'top']"
          Dim ndxLeft As XmlNode = Nothing
          Dim ndxRight As XmlNode = Nothing
          Dim ndxCommon As XmlNode = getCommonAncestor(ndxBef, ndxAft, strCond, ndxLeft, ndxRight)
          If (ndxCommon Is Nothing) Then
            ' If this happens, we will take the parent of [ndxBef] as the one after we need to plug in
            ndxBef.ParentNode.InsertAfter(ndxThis, ndxBef)
          Else
            ' Add the node to this ancestor, but after [ndxBef]
            ndxCommon.InsertAfter(ndxThis, ndxLeft)
          End If
        End If
        ' Remove the @later
        ndxThis.Attributes.Remove(ndxThis.Attributes("later"))
      Next ndxThis

 
      ' (5) Make sure the sentence is re-analyzed
      ndxThis = Nothing
      eTreeSentence(ndxFor, ndxThis)
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/OneAlpinoToPsdxForest error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function

  '----------------------------------------------------------------------------------------
  ' Name:       OneFoliaSentence()
  ' Goal:       Process one folia <s> element into a psdx <forest>
  ' History:
  ' 22-02-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function OneFoliaToPsdxForest(ByRef ndxForGrp As XmlNode, ByRef pdxSrc As XmlDocument, _
    ByRef intForestId As Integer, ByVal strTextId As String, ByVal strShort As String, _
    ByRef strSectId As String) As Boolean
    Dim ndxWork As XmlNode      ' Working node
    Dim ndxFor As XmlNode       ' Forest
    Dim ndxEtree As XmlNode     ' Etree node
    Dim ndxLeaf As XmlNode      ' Eleaf node
    Dim ndxList As XmlNodeList  ' List of items
    Dim ndxFeat As XmlNodeList  ' List of features
    Dim ndxSu As XmlNodeList    ' List of syntactic units
    Dim strWord As String       ' Word
    Dim strEng As String        ' The <t> node taken for the english BT
    Dim strPos As String        ' POS
    Dim strType As String       ' Type 
    Dim strValue As String      ' Valure
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter
    Dim nmsFolia As XmlNamespaceManager ' Needed to browse through [pdxSrc]

    Try
      ' Validate
      If (ndxForGrp Is Nothing) OrElse (pdxSrc Is Nothing) Then Return False
      ' Set the namespace manager for this particular [pdxSrc]
      nmsFolia = New XmlNamespaceManager(pdxSrc.NameTable)
      nmsFolia.AddNamespace("folia", pdxSrc.DocumentElement.NamespaceURI)
      ' Create a new <forest> element and add its properties
      ndxFor = AddXmlChild(ndxForGrp, "forest", "forestId", intForestId, "attribute", _
                           "TextId", strTextId, "attribute", _
                           "File", strShort & ".psdx", "attribute", _
                           "Location", strTextId & "." & Format(intForestId, "0000"), "attribute")
      ' Do we have a section id?
      If (strSectId <> "") Then
        ' Add it
        AddAttribute(ndxFor, "Section", strSectId)
        ' Reset id
        strSectId = ""
      End If
      intForestId += 1 : strEng = ""
      ' Add the <divs>: one for Spanish (org) and one for English (to be added)
      ndxWork = AddXmlChild(ndxFor, "div", "lang", "org", "attribute", "seg", "", "child")
      ndxWork = AddXmlChild(ndxFor, "div", "lang", "eng", "attribute", "seg", "", "child")
      ' Check if there is a <s>/<t class='gls'>
      ndxList = pdxSrc.SelectNodes("./descendant::folia:t[@class = 'gls']", nmsFolia)
      If (ndxList.Count = 0) Then
        ndxList = pdxSrc.SelectNodes("./descendant::folia:t[@class = 'lit']", nmsFolia)
      End If
      If (ndxList.Count > 0) Then
        ' Get the translation
        ndxWork.SelectSingleNode("./child::seg").InnerText = ndxList(0).InnerText
        ' Which has been taken for the English?
        strEng = ndxList(0).Attributes("class").Value
      End If
      ' Double check
      If (strEng = "") Then
        ndxList = pdxSrc.SelectNodes("./descendant::folia:t", nmsFolia)
      Else
        ndxList = pdxSrc.SelectNodes("./descendant::folia:t[@class != '" & strEng & "']", nmsFolia)
      End If
      'Add this node
      For intI = 0 To ndxList.Count - 1
        ndxWork = AddXmlChild(ndxFor, "div", "lang", ndxList(intI).Attributes("class").Value, "attribute", _
                              "seg", ndxList(intI).InnerText, "child")
      Next intI
      ' Walk through all the relevant elements of the <s> in the [pdxSrc]
      ' (1) Make a list of all the <w> items
      ndxList = pdxSrc.SelectNodes("./descendant::folia:w", nmsFolia)
      ' (2) Walk all the items
      For intI = 0 To ndxList.Count - 1
        ' Default values
        strType = "Vern"
        ' Retrieve the word
        ndxWork = ndxList(intI).SelectSingleNode("./child::folia:t", nmsFolia)
        ' Check for potential problem
        If (ndxWork Is Nothing) Then Stop : Return False
        ' Get the word into the variable [strValue]
        strValue = ndxWork.InnerText
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
          Case """", "'", "«", "»", "''", "``", "/"
            strType = "Punct" : strValue = """"
        End Select
        ' Get the POS
        ndxWork = ndxList(intI).SelectSingleNode("./child::folia:pos", nmsFolia)
        ' Check for problems
        If (ndxWork Is Nothing) OrElse (ndxWork.Attributes("class") Is Nothing) Then Stop : Return False
        ' Get the POS value (preliminary)
        strPos = ndxWork.Attributes("class").Value
        ' make sure all punctuation is recognized as such
        If (strPos Like "$*" OrElse strPos Like "LET*") Then strType = "Punct"
        ' We are CREATING, not synchronizing: add one XML child under [ndxFor]
        If (strType = "Punct") AndAlso (strValue <> "") Then
          ndxEtree = AddEtreeChild(ndxFor, loc_intEtreeId, strValue, 0, 0)
        Else
          ndxEtree = AddEtreeChild(ndxFor, loc_intEtreeId, strPos, 0, 0)
        End If
        loc_intEtreeId += 1
        ' Add <eLeaf> to the <eTree> node
        ndxLeaf = AddEleafChild(ndxEtree, strType, strValue, 0, 0)
        ' The @n attribute is the number of the word inside the current sentence (=forest)
        ndxLeaf.Attributes("n").Value = intI + 1
        ' Check if the <pos> element has a @head attribut
        If (ndxWork.Attributes("head") IsNot Nothing) Then
          ' make the @class value a feature, and the @head attribute the Label
          AddFeature(pdxCurrentFile, ndxEtree, "folia", "pos", strPos)
          strPos = ndxWork.Attributes("head").Value
          ndxEtree.Attributes("Label").Value = strPos
        End If
        ' Find the list of all the features
        ndxFeat = ndxWork.SelectNodes("./child::folia:feat", nmsFolia)
        For intJ = 0 To ndxFeat.Count - 1
          ' Add this feature
          With ndxFeat(intJ)
            ' Add the feature
            AddFeature(pdxCurrentFile, ndxEtree, "folia", _
                       .Attributes("subset").Value, .Attributes("class").Value)
            ' Check for possibly better POS (sonar)
            If (InStr(strPos, "(") > 0) AndAlso (InStr(strPos, ")") > 0) AndAlso (.Attributes("subset").Value = "head") Then
              ' Take this as POS
              strPos = .Attributes("class").Value
              ndxEtree.Attributes("Label").Value = strPos
            End If
          End With
        Next intJ
        ' Get the lemma
        ndxWork = ndxList(intI).SelectSingleNode("./child::folia:lemma", nmsFolia)
        ' Validate
        If (ndxWork IsNot Nothing) AndAlso (ndxWork.Attributes("class") IsNot Nothing) Then
          ' Process lemma as morphological feature
          AddFeature(pdxCurrentFile, ndxEtree, "M", "l", ndxWork.Attributes("class").Value)
        End If
        ' Also add the folia @xml:id as feature
        AddFeature(pdxCurrentFile, ndxEtree, "folia", "id", ndxList(intI).Attributes("xml:id").Value)
      Next intI
      ' Does this sentence have syntactic annotation?
      ndxList = pdxSrc.SelectNodes("./descendant::folia:syntax", nmsFolia)
      For intI = 0 To ndxList.Count - 1
        ' Process the syntax of this part: look for a list of <su> elements
        ndxSu = ndxList(intI).SelectNodes("./child::folia:su", nmsFolia)
        ' Process each syntactic unit recursively
        For intJ = 0 To ndxSu.Count - 1
          If (Not OneFoliaToPsdxSyntax(ndxSu(intJ), ndxFor, nmsFolia)) Then Return False
        Next intJ
      Next intI
      ' (5) Make sure the sentence is re-analyzed
      ndxWork = Nothing
      eTreeSentence(ndxFor, ndxWork)
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/OneFoliaToPsdxForest error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       OneFoliaToPsdxSyntax()
  ' Goal:       Recursively process node <ndxSu> into forest <ndxFor>
  ' History:
  ' 22-02-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function OneFoliaToPsdxSyntax(ByRef ndxSu As XmlNode, ByRef ndxFor As XmlNode, _
                                        ByRef nmsFolia As XmlNamespaceManager) As Boolean
    Try
      ' Validate
      If (ndxSu Is Nothing) OrElse (ndxFor Is Nothing) Then Return False
      ' TODO: implement code here

      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/OneFoliaToPsdxSyntax error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       ConvertOneFlexToFolia()
  ' Goal:       Convert one file in SIL-Fieldwork xml format to a folia-format xml file
  ' History:
  ' 27-03-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function ConvertOneFlexToFolia(ByVal strInFile As String, ByVal strDstFile As String) As Boolean
    Dim pdxFile As XmlDocument  ' XML file we are currently working on
    Dim pdxConv As XmlDocument  ' Converted folia XML document
    Dim ndxForest As XmlNode    ' One forest node
    Dim ndxPara As XmlNode      ' One <paragraph>
    Dim ndxText As XmlNode      ' the <text> node within folia
    Dim ndxDiv As XmlNode       ' One <div> element
    Dim ndxP As XmlNode         ' One <p> element
    Dim ndxS As XmlNode         ' One <s> element
    Dim strShort As String      ' Short file name
    Dim strTxtLang As String    ' language
    Dim strDivId As String      ' Id of <div>
    Dim strPid As String        ' Id of <p>
    Dim strSid As String        ' Id of <s>
    Dim intSection As Integer   ' Section number
    Dim intP As Integer         ' P number
    Dim intS As Integer         ' S number
    Dim intDiv As Integer       ' Div number
    Dim intNum As Integer       ' Number of forest nodes
    Dim intPtc As Integer       ' Percentage
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (Not IO.File.Exists(strInFile)) Then
        Status("Could not find file " & strInFile)
        Return False
      End If
      ' Show we are reading
      Status("Reading file " & strInFile)
      ' Initialise
      pdxFile = Nothing : pdxConv = Nothing : ndxForest = Nothing
      intDiv = 1 : intSection = 1 : intP = 1 : loc_colPosInfo.Clear()
      strShort = strInFile
      While (InStr(strShort, ".") > 0)
        strShort = IO.Path.GetFileNameWithoutExtension(strShort)
      End While
      ' Try read this file into an XML structure
      Status("Reading the text...")
      If (Not ReadXmlDoc(strInFile, pdxFile)) Then
        ' Could not read the file, so return failure
        Status("Could not read the FLEX file: " & strInFile)
        Return False
      End If
      ' Set the current file
      pdxCurrentFile = pdxFile : strTxtLang = ""
      ' Make initial header
      If (Not OneFlexToFoliaHeader(pdxConv, pdxFile, strShort, strTxtLang)) Then Return False
      ' Get <text> node in folia
      ndxText = pdxConv.SelectSingleNode("./descendant::df:text[last()]", loc_nsDf)
      If (ndxText Is Nothing) Then Return False
      ' Need to create <div> + <p>
      strDivId = strShort & ".d." & intSection : intSection += 1
      ndxDiv = AddXmlChild(ndxText, "div", "xml:id", strDivId, "attribute")
      ' Note the number of <paragraph> items
      ndxPara = pdxFile.SelectSingleNode("./descendant::paragraph[1]")
      intNum = ndxPara.ParentNode.ChildNodes.Count : intI = 0 : ndxP = Nothing
      strPid = "" : loc_colFoliaId.Clear()
      ' Walk all the paragraphs
      While (ndxPara IsNot Nothing)
        ' Note where we are in terms of <paragraph> structures
        intPtc = (intI + 1) * 100 \ intNum : intI += 1
        Status("Fieldwork to folia " & intPtc & "%", intPtc)
        ' Need to create <p>
        strPid = strDivId & ".p." & intP : intP += 1
        ndxP = AddXmlChild(ndxDiv, "p", "xml:id", strPid, "attribute")
        ' Walk all the <phrase> elements in this <paragraph>
        ndxForest = ndxPara.SelectSingleNode("./descendant::phrase[1]")
        ' Walk the foreste
        While (ndxForest IsNot Nothing)
          ' Create the <s> element under <p>
          intS += 1 : strSid = strPid & ".s." & intS
          ndxS = AddXmlChild(ndxP, "s", "xml:id", strSid, "attribute")
          ' Process this forest
          If (Not OneFlexToFoliaForest(ndxForest, ndxS, strSid, strTxtLang, loc_nsDf)) Then Return False
          ' Go to the next phrase
          ndxForest = ndxForest.SelectSingleNode("./following-sibling::phrase[1]")
        End While
        ' Go to the next paragraph
        ndxPara = ndxPara.SelectSingleNode("./following-sibling::paragraph[1]")
      End While

      ' Save the result in [strDstFile]
      pdxConv.Save(strDstFile)
      ' Provide an overview of the POS types that have been used
      frmMain.wbReport.DocumentText = GetFlexPosReport()
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/ConvertOneFlexToFolia error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       GetFlexPosReport()
  ' Goal:       Get an HTML overview of the POS types that have been used
  ' History:
  ' 02-04-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function GetFlexPosReport() As String
    Dim colHtml As New StringColl ' What we return
    Dim intI As Integer           ' Counter

    Try
      ' Validate
      If (loc_colPosInfo.Count = 0) Then Return ""
      ' Initialise
      colHtml.Add("<html><body><table><tr><td>POS</td><td>Example</td><td>Frequency</td></tr>")
      For intI = 0 To loc_colPosInfo.Count - 1
        colHtml.Add("<tr><td>" & loc_colPosInfo.Item(intI) & "</td><td>" & _
                    loc_colPosInfo.Exmp(intI) & "</td><td>" & _
                    loc_colPosInfo.Freq(intI) & "</td></tr>")
      Next intI
      colHtml.Add("</body></html>")
      ' Go to correct tab page
      With frmMain
        .TabControl1.SelectedTab = .tpReport
      End With
      Return colHtml.Text
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/GetFlexPosReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return "(error)"
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       ConvertOnePsdxToFolia()
  ' Goal:       Convert one file in psdx format to a folia-format xml file
  ' History:
  ' 01-03-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function ConvertOnePsdxToFolia(ByVal strInFile As String, ByVal strDstFile As String) As Boolean
    Dim pdxFile As XmlDocument  ' XML file we are currently working on
    Dim pdxConv As XmlDocument  ' Converted folia XML document
    Dim ndxForest As XmlNode    ' One forest node
    Dim ndxText As XmlNode      ' the <text> node within folia
    Dim ndxDiv As XmlNode       ' One <div> element
    Dim ndxP As XmlNode         ' One <p> element
    Dim ndxS As XmlNode         ' One <s> element
    Dim strShort As String      ' Short file name
    Dim strDivId As String      ' Id of <div>
    Dim strPid As String        ' Id of <p>
    Dim strSid As String        ' Id of <s>
    Dim intSection As Integer   ' Section number
    Dim intP As Integer         ' P number
    Dim intS As Integer         ' S number
    Dim intDiv As Integer       ' Div number
    Dim intNum As Integer       ' Number of forest nodes
    Dim intPtc As Integer       ' Percentage
    Dim intI As Integer         ' Counter
    'Dim nsDf As XmlNamespaceManager = Nothing ' Obligtory NS manager

    Try
      ' Validate
      If (Not IO.File.Exists(strInFile)) Then
        Status("Could not find file " & strInFile)
        Return False
      End If
      ' Show we are reading
      Status("Reading file " & strInFile)
      ' Initialise
      pdxFile = Nothing : pdxConv = Nothing : ndxForest = Nothing
      intDiv = 1 : intSection = 1 : intP = 1
      strShort = strInFile
      While (InStr(strShort, ".") > 0)
        strShort = IO.Path.GetFileNameWithoutExtension(strShort)
      End While
      ' Try read this file into an XML structure
      If (Not ReadXmlDoc(strInFile, pdxFile)) Then
        ' Could not read the file, so return failure
        Status("Could not read the XML file: " & strInFile)
        Return False
      End If
      ' Set the current file
      pdxCurrentFile = pdxFile
      ' Make initial header
      If (Not OnePsdxToFoliaHeader(pdxConv, pdxFile)) Then Return False
      ' Get <text> node in folia
      ndxText = pdxConv.SelectSingleNode("./descendant::df:text[last()]", loc_nsDf)
      If (ndxText Is Nothing) Then Return False
      ' Note the number of forests
      If (Not GetFirstForest(pdxFile, ndxForest)) Then Return False
      intNum = ndxForest.ParentNode.ChildNodes.Count : intI = 0 : ndxP = Nothing
      strPid = "" : loc_colFoliaId.Clear()
      ' Walk all forests in original
      While (ndxForest IsNot Nothing)
        ' Note where we are
        intPtc = (intI + 1) * 100 \ intNum : intI += 1
        Status("Psdx to folia " & intPtc & "%", intPtc)
        ' Check whether a <div>+<p> needs to be created for this forest's start
        If (ndxForest.Attributes("Section") IsNot Nothing) Then
          ' Need to create <div> + <p>
          strDivId = strShort & ".d." & intSection : intSection += 1
          ndxDiv = AddXmlChild(ndxText, "div", "xml:id", strDivId, "attribute")
          strPid = strDivId & ".p." & intP : intP += 1
          ndxP = AddXmlChild(ndxDiv, "p", "xml:id", strPid, "attribute")
        ElseIf (ndxP Is Nothing) Then
          ' Check for empty <p>
          strPid = strShort & ".p." & intP : intP += 1
          ndxP = AddXmlChild(ndxText, "p", "xml:id", strPid, "attribute")
        End If
        ' Create the <s> element under <p>
        intS += 1 : strSid = strPid & ".s." & intS
        ndxS = AddXmlChild(ndxP, "s", "xml:id", strSid, "attribute")
        ' Process this forest
        If (Not OnePsdxToFoliaForest(ndxForest, ndxS, strSid, loc_nsDf)) Then Return False
        ' Go to next forest
        ndxForest = ndxForest.NextSibling
      End While

      ' Save the result in [strDstFile]
      pdxConv.Save(strDstFile)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modAdapt/ConvertOnePsdxToFolia error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       OneFlexToFoliaForest()
  ' Goal:       Process one Flex <phrase> element into folia format
  '             The <w> nodes must be placed under the <s> (=[ndxS]) node
  '             The structure must be placed as <su> as child under <s>
  '             The ID of this sentence is passed on in [strSid]
  ' History:
  ' 27-03-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function OneFlexToFoliaForest(ByRef ndxFor As XmlNode, ByRef ndxS As XmlNode, _
              ByVal strSid As String, ByVal strTxtLang As String, ByRef nsDF As XmlNamespaceManager) As Boolean
    Dim ndxList As XmlNodeList  ' List of nodes
    Dim ndxLeaf As XmlNode      ' The leaf
    Dim strPos As String        ' POS
    Dim strPosOrg As String     ' POS as originally assigned
    Dim strEthno As String = "" ' Ethno code
    Dim strTclass As String     ' The class for the <t> element in a sentence
    Dim strMorph As String      ' Morph
    Dim strType As String       ' punctuation or vernacular?
    Dim ndxW As XmlNode         ' One <w> node
    Dim ndxT As XmlNode         ' One <t> node
    Dim ndxPos As XmlNode       ' One <pos> node
    Dim strWid As String        ' ID of this word
    Dim intPos As Integer       ' Position
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (ndxFor Is Nothing) OrElse (ndxFor.Name <> "phrase") Then Logging("OneFlexToFoliaForest: expect <phrase> input") : Return False
      If (ndxS Is Nothing) OrElse (ndxS.Name <> "s") Then Logging("OneFlexToFoliaForest: <s> expected") : Return False
      If (strSid = "") Then Return False
      If (nsDF Is Nothing) Then Return False
      strType = ""
      ' Create <t> nodes onder the sentence where appropriate
      ndxList = ndxFor.SelectNodes("./child::item[@type ='lit' or @type = 'gls']")
      For intI = 0 To ndxList.Count - 1
        ' Determine the class
        strTclass = ndxList(intI).Attributes("type").Value
        ' Make this <t> node
        ndxT = AddXmlChild(ndxS, "t", "class", strTclass, "attribute")
        ndxT.InnerText = ndxList(intI).InnerText
      Next intI
      ' Also make room for the original
      ndxT = AddXmlChild(ndxS, "t", "class", "original", "attribute")
      ' Check for ethno
      If (strTxtLang <> "") Then
        intPos = InStr(strTxtLang, "-")
        If (intPos > 0) Then strEthno = Left(strTxtLang, intPos - 1)
      End If
      ' Get a list of <word> nodes in this phrase
      ndxList = ndxFor.SelectNodes("./descendant::word")
      For intI = 0 To ndxList.Count - 1
        ' Prepare a word ID
        strWid = strSid & "." & intI + 1 : ndxW = Nothing : strPos = "" : strPosOrg = ""
        ' see if this is punctuation
        ndxLeaf = ndxList(intI).SelectSingleNode("./child::item[@type='punct']")
        If (ndxLeaf Is Nothing) Then
          ' The item is not punctuation: find the vernacular 
          If (strTxtLang = "") Then
            ndxLeaf = ndxList(intI).SelectSingleNode("./child::item[1]")
          Else
            ' Check for latin script
            If (strEthno <> "") Then
              ndxLeaf = ndxList(intI).SelectSingleNode("./child::item[tb:matches(@lang,'" & strEthno & "*lat')]", conTb)
            Else
              ndxLeaf = Nothing
            End If
            If (ndxLeaf Is Nothing) Then
              ndxLeaf = ndxList(intI).SelectSingleNode("./child::item[@lang='" & strTxtLang & "']")
            End If
          End If
          If (ndxLeaf IsNot Nothing) Then
            ' This may still be punctuation
            If (ndxList(intI).SelectSingleNode("./child::item[@type='pos']") Is Nothing) Then
              strType = "punct" : strPos = ndxLeaf.InnerText
            Else
              strType = "vern"
            End If
            ndxW = AddXmlChild(ndxS, "w", "xml:id", strWid, "attribute", _
                               "class", strType, "attribute", _
                               "t", ndxLeaf.InnerText, "child")
          End If
        Else
          ' Yes, we are dealing with punctuation
          ndxW = AddXmlChild(ndxS, "w", "xml:id", strWid, "attribute", _
                             "class", "punct", "attribute", _
                             "t", ndxLeaf.InnerText, "child")
          strType = "punct"
          Select Case ndxLeaf.InnerText
            Case "(", "["
              strPos = "("
            Case ")", "]"
              strPos = ")"
            Case ".", "!", "?"
              strPos = "."
            Case ",", ":", ";", "-"
              strPos = ","
            Case """", "'", "«", "»", "''", "``", "/"
              strPos = """"
            Case Else
              strPos = "-"
          End Select
        End If
        ' Validate
        If (ndxW Is Nothing) Then Stop
        ' Find the POS tag
        ndxLeaf = ndxList(intI).SelectSingleNode("./child::item[@type='pos']")
        If (ndxLeaf Is Nothing) Then
          ' Perhaps we already have a pos?
          If (strPosOrg = "") Then
            ' Probably this is punct?
             If (strPos = "") Then
              ' This cannot be punct
              Stop
            End If
          End If
        Else
          ' Get the POS initially
          strPosOrg = ndxLeaf.InnerText.ToUpper
        End If
        ndxLeaf = ndxList(intI).SelectSingleNode("./child::item[@type='gls']")
        If (ndxLeaf IsNot Nothing) Then
          strMorph = ndxLeaf.InnerText
          strMorph = Regex.Match(strMorph, "\-.+").Value
          If (strMorph <> "") Then
            strPos = strPosOrg & "-" & strMorph
          End If
          'Make sure we keep the gloss
          strMorph = ndxLeaf.InnerText
        Else
          strMorph = ""
        End If
        ' Repair
        strPos = Regex.Replace(strPos, "--", "-")
        ' TODO: Translate the POS, depending on the language...
        Select Case strEthno
          Case "lez"
            strPos = MorphToPos(strMorph, strPosOrg)
        End Select
        ' Add the POS tag
        ndxPos = AddXmlChild(ndxW, "pos", "class", strPos, "attribute")
        ' Also add the GLS as a feature
        If (strMorph <> "") Then
          AddXmlChild(ndxPos, "feat", "subset", "gls", "attribute", "class", strMorph, "attribute")
        End If
        ' Keep statistics of the POS that is being used
        If (strType <> "punct") AndAlso (ndxLeaf IsNot Nothing) Then
          loc_colPosInfo.AddUnique(strPos, ndxLeaf.InnerText)
        End If
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/OneFlexToFoliaForest error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function

  '----------------------------------------------------------------------------------------
  ' Name:       OnePsdxToFoliaForest()
  ' Goal:       Process one psdx <forest> element into folia format
  '             The <w> nodes must be placed under the <s> (=[ndxS]) node
  '             The structure must be placed as <su> as child under <s>
  '             The ID of this sentence is passed on in [strSid]
  ' History:
  ' 01-03-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function OnePsdxToFoliaForest(ByRef ndxFor As XmlNode, ByRef ndxS As XmlNode, _
              ByVal strSid As String, ByRef nsDF As XmlNamespaceManager) As Boolean
    Dim ndxList As XmlNodeList  ' List of nodes
    Dim ndxFeat As XmlNodeList  ' List of features
    Dim ndxLeaf As XmlNode      ' The leaf
    Dim strLemma As String      ' Lemma
    Dim strTclass As String     ' The class for the <t> element in a sentence
    Dim ndxW As XmlNode         ' One <w> node
    Dim ndxT As XmlNode         ' One <t> node
    Dim ndxF As XmlNode         ' One <feat> node
    Dim ndxL As XmlNode         ' One <lemma> node
    Dim ndxPos As XmlNode       ' One <pos> node
    Dim ndxSyntax As XmlNode    ' One <syntax> node
    Dim ndxSu As XmlNode        ' One <su> element
    Dim strWid As String        ' ID of this word
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter

    Try
      ' Validate
      If (ndxFor Is Nothing) OrElse (ndxFor.Name <> "forest") Then Logging("OnePsdxToFoliaForest: expect <forest> input") : Return False
      If (ndxS Is Nothing) OrElse (ndxS.Name <> "s") Then Logging("OnePsdxToFoliaForest: <s> expected") : Return False
      If (strSid = "") Then Return False
      If (nsDF Is Nothing) Then Return False
      ' Get the text of this sentence in all applicable languages
      ndxList = ndxFor.SelectNodes("./child::div")
      For intI = 0 To ndxList.Count - 1
        ' Determine the class for this element
        strTclass = ndxList(intI).Attributes("lang").Value
        If (strTclass = "org") Then strTclass = "original"
        ' Create the <t> element under the <s>
        ndxT = AddXmlChild(ndxS, "t", "class", strTclass, "attribute")
        ndxT.InnerText = ndxList(intI).SelectSingleNode("./child::seg").InnerText
      Next intI
      ' Get a list of <eTree> nodes that have an <eLeaf> child
      ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf)>0]")
      For intI = 0 To ndxList.Count - 1
        ' Get to the leaf
        ndxLeaf = ndxList(intI).SelectSingleNode("./child::eLeaf[1]")
        ' Create a new <w> node under <s>
        strWid = strSid & "." & intI + 1
        ndxW = AddXmlChild(ndxS, "w", "xml:id", strWid, "attribute", _
                           "class", ndxLeaf.Attributes("Type").Value, "attribute", _
                           "t", ndxLeaf.Attributes("Text").Value, "child")
        ' Try to find a lemma feature
        strLemma = GetFeature(ndxList(intI), "M", "l")
        If (strLemma <> "") Then
          ' Add a lemma child to <w>
          ndxL = AddXmlChild(ndxW, "lemma", "class", strLemma, "attribute")
        End If
        ' Add the POS tag
        ndxPos = AddXmlChild(ndxW, "pos", "class", ndxList(intI).Attributes("Label").Value, "attribute")
        ' Find any other features
        ndxFeat = ndxList(intI).SelectNodes("./child::fs/child::f[not(@name='l' and parent::fs/@type='M')]")
        For intJ = 0 To ndxFeat.Count - 1
          ' Add this feature
          ndxF = AddXmlChild(ndxPos, "feat", "subset", ndxFeat(intJ).ParentNode.Attributes("type").Value & "/" & _
                             ndxFeat(intJ).Attributes("name").Value, "attribute", _
                             "class", ndxFeat(intJ).Attributes("value").Value, "attribute")
        Next intJ
      Next intI
      ' Create a <syntax> node under the sentence
      ndxSyntax = AddXmlChild(ndxS, "syntax")
      ' Add all children under <forest>
      ndxList = ndxFor.SelectNodes("./child::eTree")
      For intI = 0 To ndxList.Count - 1
        ' Add this child and whatever is under it recursively
        If (Not OnePsdxToFoliaSu(ndxSyntax, ndxList(intI), strSid)) Then Return False
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/OnePsdxToFoliaForest error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       OnePsdxToFoliaSu()
  ' Goal:       Add psdx element [ndxThis] as <su> under [ndxParentFolia]
  ' History:
  ' 12-03-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function OnePsdxToFoliaSu(ByRef ndxParentFolia As XmlNode, ByRef ndxThis As XmlNode, _
                                    ByVal strSid As String) As Boolean
    Dim ndxList As XmlNodeList  ' List of nodes
    Dim ndxFeat As XmlNodeList  ' List of features
    Dim ndxLeaf As XmlNode      ' An <eLeaf> node
    Dim ndxF As XmlNode         ' One <feat> node
    Dim ndxSu As XmlNode        ' One <su> node
    Dim ndxWref As XmlNode      ' One <wref> element
    Dim strWid As String        ' the ID of the word
    Dim strSuId As String       ' ID of this syntactic unit
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter

    Try
      ' Validate
      If (ndxParentFolia Is Nothing) OrElse (ndxThis Is Nothing) Then Return False
      ' Determine the xml:id of this syntactic unit
      strSuId = strSid & ".su." & ndxParentFolia.SelectNodes("./ancestor-or-self::df:syntax/descendant::df:su", loc_nsDf).Count + 1
      ' Mark the connection between this <eTree> and <su>
      loc_colFoliaId.Add(ndxThis.Attributes("Id").Value, strSuId)
      ' Add the [ndxThis] contents
      ndxSu = AddXmlChild(ndxParentFolia, "su", _
                          "xml:id", strSuId, "attribute", _
                          "class", ndxThis.Attributes("Label").Value, "attribute")
      ' Add features
      ndxFeat = ndxThis.SelectNodes("./child::fs/child::f")
      For intJ = 0 To ndxFeat.Count - 1
        ' Add this feature
        ndxF = AddXmlChild(ndxSu, "feat", "subset", ndxFeat(intJ).ParentNode.Attributes("type").Value & "/" & _
                           ndxFeat(intJ).Attributes("name").Value, "attribute", _
                           "class", ndxFeat(intJ).Attributes("value").Value, "attribute")
      Next intJ
      ' Check out if this is an endnode
      ndxLeaf = ndxThis.SelectSingleNode("./child::eLeaf[1]")
      If (ndxLeaf IsNot Nothing) Then
        ' Calculate the Wid
        strWid = strSid & "." & ndxThis.SelectNodes("./ancestor::forest/descendant::eLeaf[@to < " & _
                                                    ndxLeaf.Attributes("from").Value & "]").Count + 1
        ' Add the <wref> element pointing to the <eLeaf> within Folia
        ndxWref = AddXmlChild(ndxSu, "wref", "id", strWid, "attribute", _
                              "t", ndxLeaf.Attributes("Text").Value, "attribute")
      End If
      ' Process all my psdx children
      ndxList = ndxThis.SelectNodes("./child::eTree")
      For intI = 0 To ndxList.Count - 1
        ' Process this child
        If (Not OnePsdxToFoliaSu(ndxSu, ndxList(intI), strSid)) Then Return False
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/OnePsdxToFoliaSu error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       OneTigerOrder()
  ' Goal:       Make sure the <nt> elements inside <nonterminals> are ordered properly
  ' History:
  ' 06-02-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function OneTigerOrder(ByRef pdxSrc As XmlDocument) As Boolean
    Dim ndxList As XmlNodeList    ' List of <nt> nodes
    Dim ndxNtParent As XmlNode    ' Parent of <nt>
    Dim intIdFirst As Integer = 0 ' First @id number in one <nt> section
    Dim intIdLast As Integer = 0  ' Last @id number in one <nt> section
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim bOk As Boolean            ' Flag
    Dim bTonly As Boolean         ' This node has only terminals

    Try
      ' Validate
      If (pdxSrc Is Nothing) Then Return False
      ndxNtParent = pdxSrc.SelectSingleNode("./descendant::nonterminals")
      Do
        ' Get a list of nodes
        ndxList = ndxNtParent.SelectNodes("./child::nt")
        ' Initialise
        intIdFirst = 0 : intIdLast = 0 : bOk = True
        ' Walk all the elements
        For intI = 0 To ndxList.Count - 1
          ' Check if this one should be counted
          If (ndxList(intI).SelectSingleNode("./child::edge[tb:matches(@idref,'*_5??')]", conTb) Is Nothing) Then
            intIdFirst = LowestIdNumber(ndxList(intI))
            If (intIdFirst < intIdLast) Then
              bOk = False
              ' Find out where it should go
              For intJ = 0 To intI - 1
                ' Should it go here?
                intIdLast = LowestIdNumber(ndxList(intJ))
                If (intIdFirst < intIdLast) Then
                  ' This is where it should go
                  ndxNtParent.InsertBefore(ndxList(intI), ndxList(intJ))
                  Exit For
                End If
              Next intJ
              Exit For
            End If

          End If
          ' Adapt last
          intIdLast = intIdFirst
        Next intI
      Loop Until bOk
      ' Step #2: make sure that <nt> gatherings only having end-nodes receive priority
      Do
        ' Get a list of nodes
        ndxList = ndxNtParent.SelectNodes("./child::nt")
        ' Initialise
        intIdFirst = 0 : intIdLast = 0 : bOk = True : bTonly = True
        ' Walk all the elements
        For intI = 0 To ndxList.Count - 1
          ' Does this come in order?
          If (Not bTonly) AndAlso (ndxList(intI).SelectSingleNode("./child::edge[tb:matches(@idref,'*_5??')]", conTb) Is Nothing) Then
            ' This should come earlier, if possible
            If (intIdLast > 0) Then
              ' Move this node after intIdLast
              ndxNtParent.InsertAfter(ndxList(intI), ndxList(intIdLast))
              bOk = False : Exit For
            End If
          End If
          ' Does this one only contain <t> nodes?
          If (ndxList(intI).SelectSingleNode("./child::edge[tb:matches(@idref,'*_5??')]", conTb) IsNot Nothing) Then
            bTonly = False : intIdLast = intI - 1
          End If
        Next intI
      Loop Until bOk

      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/OneTigerOrder error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '-----------------------------------------------------------------------------------------------------------
  ' Name:       OneTigerNtGroup()
  ' Goal:       Process the <nt> node in [ndxNt] (tiger), adding it on the correct place under [ndxFor] (psdx)
  '             Return the newly created node under [ndxFor] for the <nt> in [ndxThis]
  ' Note:       This is a recursive function
  ' History:
  ' 06-02-2014  ERK Created
  '-----------------------------------------------------------------------------------------------------------
  Private Function OneTigerNtGroup(ByRef ndxFor As XmlNode, ByRef pdxSrc As XmlDocument, ByRef ndxNt As XmlNode, _
                                   ByRef ndxThis As XmlNode) As Boolean
    Dim ndxList As XmlNodeList    ' List of nodes
    Dim ndxPrec As XmlNodeList    ' List of preceding nodes
    Dim ndxChild As XmlNode       ' One child
    Dim ndxWork As XmlNode        ' Working node
    Dim ndxWalk As XmlNode        ' Walking node
    Dim ndxOneCh As XmlNode       ' One child node (within tiger)
    Dim ndxNP As XmlNode          ' An NP that needs to be inserted on-the-fly
    Dim intI As Integer           ' Counter
    Dim intL As Integer           ' Counter

    Try
      ' Validate
      If (ndxFor Is Nothing) OrElse (pdxSrc Is Nothing) OrElse (ndxNt Is Nothing) Then Return False
      If (ndxFor.Name <> "forest") OrElse (Not DoLike(ndxNt.Name, "nt|t")) Then Return False
      ' Add this node 
      ' Check if it not yet exists
      ndxThis = ndxFor.SelectSingleNode("./descendant::eTree[child::fs/child::f[" & _
                                            "@name='id' and @value='" & ndxNt.Attributes("id").Value & "']]")
      If (ndxThis Is Nothing) Then
        ' Create a node for this
        ndxThis = CreateNewEtree(pdxCurrentFile) : ndxThis.Attributes("Label").Value = ndxNt.Attributes("cat").Value
        ' Initially put it under the forest
        ndxFor.PrependChild(ndxThis)
      Else
        ' There might be something wrong with the structure of the original Tiger file 
        ' Look for the sentence number
        ' Stop
        ' Logging("Warning at [" & ndxFor.SelectSingleNode("./descendant::fs[@type='tig']/child::f[@name='id']/@value").Value & "]")
      End If
      ' Add appropriate tiger feature
      If (Not AddFeatureXml(ndxThis, "tig", "id", ndxNt.Attributes("id").Value)) Then Return False
      ' Make a list of all <edge> children
      ndxList = ndxNt.SelectNodes("./child::edge")
      ' Walk all the children
      For intI = 0 To ndxList.Count - 1
        ' Access the child
        ndxChild = ndxList(intI)
        ' Get the corresponding <t> or <nt> node within tiger
        ndxOneCh = TigerNodeGet(ndxChild)


        ' See where this is located inside [ndxFor]
        ndxWork = ndxFor.SelectSingleNode("./descendant::eTree[child::fs/child::f[" & _
                                          "@name='id' and @value='" & ndxOneCh.Attributes("id").Value & "']]")
        If (ndxWork Is Nothing) Then
          ' This occurs when the node has not yet been created...
          If (ndxOneCh.Name = "nt") Then
            ' Make sure [ndxWork] is properly created
            If (Not OneTigerNtGroup(ndxFor, pdxSrc, ndxOneCh, ndxWork)) Then Return False
          Else
            ' This should never happen
            Stop
          End If
        End If
        If (ndxThis.Attributes("Label").Value <> "NP") AndAlso (DoLike(ndxWork.Attributes("Label").Value, "N[123456789]*|NN|VNW*")) Then
          ' Adding a "N" under a nonNP requires inserting an NP level
          ndxNP = CreateNewEtree(pdxCurrentFile)
          If (ndxChild.Attributes("label").Value <> "--") Then
            ndxNP.Attributes("Label").Value = "NP-" & ndxChild.Attributes("label").Value
          End If
          ' Position [ndxNP]
          ndxWork.ParentNode.InsertAfter(ndxNP, ndxWork)
          ' Add [ndxWork] under ndxNP
          ndxNP.AppendChild(ndxWork)
          ' Set the new [ndxWork] to the [ndxNP]
          ndxWork = ndxNP
        Else
          ' If (InStr(ndxChild.Attributes("label").Value, "-") > 0) Then Stop
          ' Adapt the label of this node according to the <edge>'s @label value
          If (ndxChild.Attributes("label").Value <> "--") Then
            ' ============= Double check ======================
            If (InStr(ndxChild.Attributes("label").Value, "-") > 0) Then Stop
            ' =================================================
            ndxWork.Attributes("Label").Value &= "-" & ndxChild.Attributes("label").Value
          End If
        End If
        ' ============= Double check ======================
        If (InStr(ndxWork.Attributes("Label").Value, "--") > 0) Then Stop
        ' =================================================
        ' Find out where to put this <t> element inside  [ndxThis] children
        ndxPrec = ndxThis.SelectNodes("./child::eTree") : ndxWalk = Nothing
        ' Walk all the children of [ndxThis] from right to left (highest @n number to lowest)
        For intL = ndxPrec.Count - 1 To 0 Step -1
          ' Check whether the leftmost (lowest) @n of [ndxWork] should precede child number [intL]
          If (LeftmostWordNumber(ndxWork) < LeftmostWordNumber(ndxPrec(intL))) Then
            ' Insert [ndxWork]before child number [intL]
            ndxThis.InsertBefore(ndxWork, ndxPrec(intL))
            ' Indicate that we have a result and exit the for-loop
            ndxWalk = ndxPrec(intL) : Exit For
          End If
        Next intL
        ' Did we find a child where to store [ndxWork]?
        If (ndxWalk Is Nothing) Then
          ' No appropriate child has been found, 
          '   which means [ndxWork] should be appended as a child under [ndxThis]
          ndxThis.AppendChild(ndxWork)
        End If
      Next intI
      ' ============= Double check ======================
      If (InStr(ndxThis.Attributes("Label").Value, "--") > 0) Then Stop
      ' =================================================

      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/OneTigerNtGroup error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       OneTigerToPsdxForest()
  ' Goal:       Convert one forest section in [pdxSrc] and process it in [ndxFor]
  ' History:
  ' 01-02-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function OneTigerToPsdxForest(ByRef ndxFor As XmlNode, ByRef pdxSrc As XmlDocument, ByRef intI As Integer, _
                                         ByVal strLang As String, ByVal bSynchronize As Boolean) As Boolean
    Dim ndxWork As XmlNode        ' Working node
    Dim ndxNP As XmlNode          ' NP level insertion
    Dim ndxThis As XmlNode        ' Working node
    Dim ndxLeaf As XmlNode        ' Leaf child
    Dim ndxRoot As XmlNode        ' Root node (tig)
    Dim ndxHigh As XmlNode        ' Higest ancestor
    Dim ndxWalk As XmlNode        ' Walking node
    Dim ndxNbr As XmlNode         ' Neighbour
    Dim ndxICH As XmlNode         ' New *ICH* node (interpret constituent here)
    Dim ndxOneCh As XmlNode       ' One child node
    Dim ndxNT As XmlNode          ' One NT node
    Dim ndxList As XmlNodeList    ' List of nodes
    Dim ndxChild As XmlNodeList   ' Children of current <nt>
    Dim ndxPrec As XmlNodeList    ' Preceding children
    Dim ndxT As XmlNodeList       ' List of terminal <t> nodes
    Dim strType As String = ""    ' Type of eLeaf
    Dim strValue As String = ""   ' Value of <eLeaf>
    Dim strBack As String = ""    ' syntax
    Dim intEtreeId As Integer = 1 ' ID of <eTree>
    Dim intJ As Integer           ' Counter
    Dim intK As Integer           ' Counter
    Dim intL As Integer           ' Counter
    Dim intM As Integer           ' Counter
    Dim intChild As Integer       ' Which child are we?
    'Dim intIdLast As Integer      ' Last id from here
    Dim intIdLeft As Integer      ' Left id
    Dim intIdRight As Integer     ' Right id
    Dim intICHid As Integer = 1   ' ICH identifier
    Dim intNum As Integer         ' Number of @n
    Dim bIsLast As Boolean        ' Last (topmost) node (is the last <nt> node)
    Dim bChanged As Boolean = False ' Something has been changed

    Try
      ' Validate
      If (ndxFor Is Nothing) OrElse (ndxFor.Name <> "forest") Then Return False
      ' Note where we are
      intChild = 0
      ' Get a list of terminal nodes
      ndxT = pdxSrc.SelectNodes("./descendant::t")
      ' =========== DEBUG ============
      ' If (ndxT(0).Attributes("id").Value Like "s4191*") Then Stop
      ' ==============================
      For intI = 0 To ndxT.Count - 1
        ' Default values
        strType = "Vern" : strValue = ndxT(intI).Attributes("word").Value
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
          Case """", "'", "«", "»", "''", "``", "/"
            strType = "Punct" : strValue = """"
        End Select
        If (ndxT(intI).Attributes("pos").Value Like "$*") Then
          strType = "Punct"
        End If
        ' We are CREATING, not synchronizing: add one XML child under [ndxFor]
        If (strType = "Punct") AndAlso (strValue <> "") Then
          ndxWork = AddEtreeChild(ndxFor, loc_intEtreeId, strValue, 0, 0)
        Else
          ndxWork = AddEtreeChild(ndxFor, loc_intEtreeId, ndxT(intI).Attributes("pos").Value, 0, 0)
        End If
        loc_intEtreeId += 1
        ' Add <eLeaf> to the <eTree> node
        ndxLeaf = AddEleafChild(ndxWork, strType, strValue, 0, 0)
        ndxLeaf.Attributes("n").Value = intI + 1
        ' Add features to the <eTree> node
        If (Not AddFeatureXml(ndxWork, "tig", "morph", ndxT(intI).Attributes("morph").Value)) Then Return False
        If (Not AddFeatureXml(ndxWork, "tig", "id", ndxT(intI).Attributes("id").Value)) Then Return False
        ' Possibly lemma and other features
        If (ndxT(intI).Attributes("lemma") IsNot Nothing) Then
          If (Not AddFeatureXml(ndxWork, "M", "l", ndxT(intI).Attributes("lemma").Value)) Then Return False
        End If
        ' TODO: process attributes @case, @number, @gender, @person, @degree, @tense, @mood
      Next intI
      ' Find the root node in tiger
      ndxWork = pdxSrc.SelectSingleNode("./descendant::graph")
      If (ndxWork Is Nothing) Then Status("Could not find root node") : Return False
      ndxRoot = pdxSrc.SelectSingleNode("./descendant::nt[@id='" & ndxWork.Attributes("root").Value & "']")
      If (ndxRoot Is Nothing) Then
        ndxRoot = pdxSrc.SelectSingleNode("./descendant::t[@id='" & ndxWork.Attributes("root").Value & "']")
        If (ndxRoot Is Nothing) Then
          ' Cannot find root - cannot process this line.
          Stop
        End If
      End If
      ndxThis = Nothing
      ' Process the root node
      If (Not OneTigerNtGroup(ndxFor, pdxSrc, ndxRoot, ndxThis)) Then Status("Could not process NT group") : Return False

      ' Get the root in terms of Psdx now
      ndxRoot = ndxFor.SelectSingleNode("./descendant::eTree[" & _
          "child::fs[@type='tig']/child::f[@name='id' and @value='" & ndxWork.Attributes("root").Value & "']]")


      ' Check for any unresolved remaining elements 
      '  (these elements are dangling freely under [ndxForest])
      ' Get the consecutive word numbers @n between which there might be a gap
      intIdLeft = RightMostWordNumber(ndxThis)
      intIdRight = pdxSrc.SelectNodes("./descendant::t").Count
      ' Process all that is needed
      While (intIdRight - intIdLeft > 0)
        ' Loop 
        intL = intIdLeft + 1
        ' Find highest <eTree> ancestor of [intL]
        ndxHigh = ndxFor.SelectSingleNode("./descendant::eLeaf[@n=" & intL & "]/ancestor::eTree[last()]")
        ' double check
        If (ndxRoot IsNot ndxHigh) Then
          ' Find the location between the children of [ndxRoot] where [ndxHigh] may be placed
          ndxPrec = ndxRoot.SelectNodes("./child::eTree") : ndxWalk = Nothing
          For intL = 0 To ndxPrec.Count - 1
            ' Check if the word number (intIdLeft+1) can be placed after child [intL]
            If (intIdLeft + 1 < RightMostWordNumber(ndxPrec(intL))) Then
              ' Insert it after child [intL]
              ndxRoot.InsertAfter(ndxHigh, ndxPrec(intL))
              ' Indicate we have found an element and leave the for-loop
              ndxWalk = ndxPrec(intL) : Exit For
            End If
          Next intL
          ' Double check
          If (ndxWalk Is Nothing) Then
            ' Just append it
            Try
              ndxRoot.AppendChild(ndxHigh)
            Catch ex As Exception
              ' Only log errors
              Logging("OneTigerToPsdxForest warning at forest=" & ndxFor.Attributes("forestId").Value)
              Stop
            End Try
          End If
        Else
          intIdLeft += 1
        End If
      End While

      ' Postprocessing: re-arrange the constituents that are not in the correct order
      Do
        bChanged = False
        ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf[@Type='Punct' or @Type='Vern'])>0]")
        For intI = 0 To ndxList.Count - 1
          ' Check if something is missing between me and my predecessor
          If (intI = 0) Then
            ' This is the first one, so check if we start out with "1"
            intIdLeft = 0
          Else
            intIdLeft = RightMostWordNumber(ndxList(intI - 1))
          End If
          intIdRight = LeftmostWordNumber(ndxList(intI))
          ndxThis = ndxList(intI).ParentNode
          ' Is there a gap?
          If (intIdRight - intIdLeft > 1) Then
            ' Try process the first number 
            intL = intIdLeft + 1
            ' Find highest <eTree> ancestor of [intL]
            'ndxHigh = ndxFor.SelectSingleNode("./descendant::eTree[descendant::eLeaf[@Type='Punct' or @Type='Vern'][1]/@n>" & _
            '                                  intIdLeft & " and descendant::eLeaf[@Type='Punct' or @Type='Vern'][last()]/@n < " & intIdRight & "]")
            ndxPrec = ndxFor.SelectNodes("./descendant::eTree[descendant::eLeaf[@Type='Punct' or @Type='Vern'][1]/@n>" & _
                                              intIdLeft & " and descendant::eLeaf[@Type='Punct' or @Type='Vern'][last()]/@n < " & intIdRight & "]")
            ' Select the one with the highest number of <eLeaf/@n> elements
            If (ndxPrec.Count = 1) Then
              ndxHigh = ndxPrec(0)
            Else
              intNum = 0 : ndxHigh = Nothing
              For intM = 0 To ndxPrec.Count - 1
                If (ndxPrec(intM).SelectNodes("./descendant::eLeaf[@n>0]").Count > intNum) Then
                  intNum = ndxPrec(intM).SelectNodes("./descendant::eLeaf[@n>0]").Count
                  ndxHigh = ndxPrec(intM)
                End If
              Next intM
            End If

            ' Validate
            If (ndxHigh Is Nothing) Then
              Stop
            Else
              ' Check if an "ICH" node needs to be created or not...
              If (ndxHigh.ParentNode IsNot ndxThis) AndAlso (ndxHigh.SelectSingleNode("./child::eLeaf[@Type='Punct']") Is Nothing) Then
                ' Create a new ICH node 
                ndxICH = CreateNewEtree(pdxCurrentFile)
                ndxICH.Attributes("Label").Value = ndxHigh.Attributes("Label").Value
                ' Add an <eLeaf>
                ndxLeaf = AddEleafChild(ndxICH, "Star", "*ICH*-" & intICHid, 0, 0)
                ' Append this node
                ndxHigh.ParentNode.AppendChild(ndxICH)
                ' Adapt the label of the <ndxWork> node
                ndxHigh.Attributes("Label").Value &= "-" & intICHid
                ' Adapt the ICH id counter
                intICHid += 1
              Else
                ndxICH = Nothing
                ' Stop
              End If
              ' Move the node with the actual information to the location we want it to be
              Try
                ndxThis.InsertBefore(ndxHigh, ndxList(intI))
              Catch ex As Exception
                ' There is a problem...
                strBack = "" : TravNode(ndxFor, "Tiger", strBack)
                Debug.Print("Text=" & ndxFor.Attributes("TextId").Value & " forest=" & ndxFor.Attributes("forestId").Value & vbCrLf & _
                            "Seq=" & GetSeq(ndxFor) & vbCrLf & strBack)
                Stop
              End Try
              ' Walk all my ancestors to update the order -- if needed
              If (ndxICH IsNot Nothing) Then
                ndxWalk = ndxICH.ParentNode
                While (ndxWalk IsNot Nothing) AndAlso (ndxWalk.Name = "eTree")
                  ' Check if my leftmost is still *before* my right neighbour
                  ndxNbr = ndxWalk.SelectSingleNode("./following-sibling::eTree")
                  If (ndxNbr IsNot Nothing) Then
                    If (LeftmostWordNumber(ndxNbr) < LeftmostWordNumber(ndxWalk)) Then
                      ' swap places
                      ndxWalk.ParentNode.InsertAfter(ndxWalk, ndxNbr)
                    End If
                  End If
                  ' Go to parent
                  ndxWalk = ndxWalk.ParentNode
                End While
              End If
              ' Note there has been a change
              bChanged = True : Exit For
            End If
          End If
        Next intI
      Loop While (bChanged)
      ' Post processing: look for extra-syntactical elements
      ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf[@Type='Vern'])>0]")
      For intI = 0 To ndxList.Count - 1
        ' Any leaf that is not part of an <nt> does not join in the syntactic analysis...
        ndxOneCh = pdxSrc.SelectSingleNode("./descendant::nt/child::edge[@idref='" & GetFeature(ndxList(intI), "tig", "id") & "']")
        If (ndxOneCh Is Nothing) Then
          ' Give it a feature indicating that it does not join in the syntactic analysis
          AddFeature(pdxCurrentFile, ndxList(intI), "tig", "syn", "no")
        End If
      Next intI
      ' (5) Make sure the sentence is re-analyzed
      ndxWork = Nothing
      eTreeSentence(ndxFor, ndxWork)
      'Debug.Print(NodeText(ndxFor))
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/OneTigerToPsdxForest error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       OneTigerToPsdxForest_ORG()
  ' Goal:       Convert one forest section in [pdxSrc] and process it in [ndxFor]
  ' History:
  ' 01-02-2014  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function OneTigerToPsdxForest_ORG(ByRef ndxFor As XmlNode, ByRef pdxSrc As XmlDocument, ByRef intI As Integer, _
                                         ByVal strLang As String, ByVal bSynchronize As Boolean) As Boolean
    Dim ndxWork As XmlNode        ' Working node
    Dim ndxNP As XmlNode          ' NP level insertion
    Dim ndxThis As XmlNode        ' Working node
    Dim ndxLeaf As XmlNode        ' Leaf child
    Dim ndxRoot As XmlNode        ' Root node (tig)
    Dim ndxHigh As XmlNode        ' Higest ancestor
    Dim ndxWalk As XmlNode        ' Walking node
    Dim ndxNbr As XmlNode         ' Neighbour
    Dim ndxICH As XmlNode         ' New *ICH* node (interpret constituent here)
    Dim ndxOneCh As XmlNode       ' One child node
    Dim ndxNT As XmlNode          ' One NT node
    Dim ndxList As XmlNodeList    ' List of nodes
    Dim ndxChild As XmlNodeList   ' Children of current <nt>
    Dim ndxPrec As XmlNodeList    ' Preceding children
    Dim ndxT As XmlNodeList       ' List of terminal <t> nodes
    Dim strType As String = ""    ' Type of eLeaf
    Dim strValue As String = ""   ' Value of <eLeaf>
    Dim strBack As String = ""    ' syntax
    Dim intEtreeId As Integer = 1 ' ID of <eTree>
    Dim intJ As Integer           ' Counter
    Dim intK As Integer           ' Counter
    Dim intL As Integer           ' Counter
    Dim intM As Integer           ' Counter
    Dim intChild As Integer       ' Which child are we?
    'Dim intIdLast As Integer      ' Last id from here
    Dim intIdLeft As Integer      ' Left id
    Dim intIdRight As Integer     ' Right id
    Dim intICHid As Integer = 1   ' ICH identifier
    Dim intNum As Integer         ' Number of @n
    Dim bIsLast As Boolean        ' Last (topmost) node (is the last <nt> node)
    Dim bChanged As Boolean = False ' Something has been changed

    Try
      ' Validate
      If (ndxFor Is Nothing) OrElse (ndxFor.Name <> "forest") Then Return False
      ' Note where we are
      intChild = 0
      ' Get a list of terminal nodes
      ndxT = pdxSrc.SelectNodes("./descendant::t")
      ' =========== DEBUG ============
      ' If (ndxT(0).Attributes("id").Value Like "s4191*") Then Stop
      ' ==============================
      For intI = 0 To ndxT.Count - 1
        ' Default values
        strType = "Vern" : strValue = ndxT(intI).Attributes("word").Value
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
          ndxWork = AddEtreeChild(ndxFor, loc_intEtreeId, strValue, 0, 0)
        Else
          ndxWork = AddEtreeChild(ndxFor, loc_intEtreeId, ndxT(intI).Attributes("pos").Value, 0, 0)
        End If
        loc_intEtreeId += 1
        ' Add <eLeaf> to the <eTree> node
        ndxLeaf = AddEleafChild(ndxWork, strType, strValue, 0, 0)
        ndxLeaf.Attributes("n").Value = intI + 1
        ' Add features to the <eTree> node
        If (Not AddFeatureXml(ndxWork, "tig", "morph", ndxT(intI).Attributes("morph").Value)) Then Return False
        If (Not AddFeatureXml(ndxWork, "tig", "id", ndxT(intI).Attributes("id").Value)) Then Return False
      Next intI
      ' New approadh

      ndxRoot = Nothing
      ' '' =========== DEBUG ============
      'If (ndxFor.Attributes("TextId").Value Like "*7324") AndAlso (ndxFor.Attributes("forestId").Value = 44) Then
      '  Stop
      'End If
      ' '' ==============================
      ' We have read one sentence in terms of end-nodes, now analyze and add this to the psdx
      ' (1) Visit all the <nt> sections and process their contents in turn
      ndxList = pdxSrc.SelectNodes("./descendant::nt")
      For intJ = 0 To ndxList.Count - 1
        ' Is this the last?
        bIsLast = (intJ = ndxList.Count - 1)
        ' Access this <nt> node
        ndxNT = ndxList(intJ)
        ' Check if it not yet exists
        ndxThis = ndxFor.SelectSingleNode("./descendant::eTree[child::fs/child::f[" & _
                                              "@name='id' and @value='" & ndxNT.Attributes("id").Value & "']]")
        If (ndxThis Is Nothing) Then
          ' Create a node for this
          ndxThis = CreateNewEtree(pdxCurrentFile) : ndxThis.Attributes("Label").Value = ndxNT.Attributes("cat").Value
          ' Initially put it under the forest
          ndxFor.PrependChild(ndxThis)
        Else
          ' There might be something wrong with the structure of the original Tiger file 
          ' Look for the sentence number
          ' Stop
          Logging("Warning at [" & ndxFor.SelectSingleNode("./descendant::fs[@type='tig']/child::f[@name='id']/@value").Value & "]")
        End If
        ' The last one will be the root...
        ndxRoot = ndxThis
        ' Add appropriate tiger feature
        If (Not AddFeatureXml(ndxThis, "tig", "id", ndxNT.Attributes("id").Value)) Then Return False
        ' If (ndxNT.Attributes("id").Value Like "*510") Then Stop
        ' Get the children of this node
        ndxChild = ndxNT.SelectNodes("./child::edge")
        ' intIdLast = 0
        For intK = 0 To ndxChild.Count - 1
          ' Get the corresponding node
          ndxOneCh = TigerNodeGet(ndxChild(intK))
          ' ========== DEBUGGING =============
          ' If (ndxChild(intK).Attributes("idref").Value Like "*.500") Then Stop
          ' ==================================
          ' Validate
          If (ndxOneCh Is Nothing) Then
            ' Warn
            MsgBox("modConvert/OneTigerToPsdxForest error: could not find <edge> with id= [" & _
                   ndxChild(intK).Attributes("idref").Value & "].")
            Stop
            Return False
          Else
            ' Locate the tiger node identified as <edge> from its location under <ndxFor>
            ndxWork = ndxFor.SelectSingleNode("./descendant::eTree[child::fs/child::f[" & _
                                              "@name='id' and @value='" & ndxOneCh.Attributes("id").Value & "']]")
            If (ndxWork Is Nothing) Then
              ' The node above has NOT yet been made available, so create a dummy node for this
              ndxWork = CreateNewEtree(pdxConv) : ndxWork.Attributes("Label").Value = ndxOneCh.Attributes("cat").Value
              ' Add appropriate tiger feature
              If (Not AddFeatureXml(ndxWork, "tig", "id", ndxOneCh.Attributes("id").Value)) Then Return False
              ' Initially put it under the forest
              ndxFor.PrependChild(ndxWork)
            End If
            If (ndxThis.Attributes("Label").Value <> "NP") AndAlso (DoLike(ndxWork.Attributes("Label").Value, "N[123456789]*|VNW*")) Then
              ' Adding a "N" under a nonNP requires inserting an NP level
              ndxNP = CreateNewEtree(pdxCurrentFile) : ndxNP.Attributes("Label").Value = "NP-" & ndxChild(intK).Attributes("label").Value
              ' Position [ndxNP]
              ndxWork.ParentNode.InsertAfter(ndxNP, ndxWork)
              ' Add [ndxWork] under ndxNP
              ndxNP.AppendChild(ndxWork)
              ' Set the new [ndxWork] to the [ndxNP]
              ndxWork = ndxNP
            Else
              ' Adapt the label of this node according to the <edge>'s @label value
              ndxWork.Attributes("Label").Value &= "-" & ndxChild(intK).Attributes("label").Value
            End If
            ' Find out whether to prepend or append it
            Select Case ndxOneCh.Name
              Case "t"    ' Might demand more processing
                ' Find out where to put this <t> element inside  [ndxThis] children
                ndxPrec = ndxThis.SelectNodes("./child::eTree") : ndxWalk = Nothing
                ' Walk all the children of [ndxThis] from right to left (highest @n number to lowest)
                For intL = ndxPrec.Count - 1 To 0 Step -1
                  ' Check whether the leftmost (lowest) @n of [ndxWork] should precede child number [intL]
                  If (LeftmostWordNumber(ndxWork) < LeftmostWordNumber(ndxPrec(intL))) Then
                    ' Insert [ndxWork]before child number [intL]
                    ndxThis.InsertBefore(ndxWork, ndxPrec(intL))
                    ' Indicate that we have a result and exit the for-loop
                    ndxWalk = ndxPrec(intL) : Exit For
                  End If
                Next intL
                ' Did we find a child where to store [ndxWork]?
                If (ndxWalk Is Nothing) Then
                  ' No appropriate child has been found, 
                  '   which means [ndxWork] should be appended as a child under [ndxThis]
                  ndxThis.AppendChild(ndxWork)
                End If

                '' Append it as child under me
                'ndxThis.AppendChild(ndxWork)
                '' Get its id number
                'intIdLast = GetTigerIdNumber(ndxChild(intK).Attributes("idref").Value)
              Case "nt"   ' This demands detailed processing
                ' Find out where to put this <nt> element inside  [ndxThis] children
                ndxPrec = ndxThis.SelectNodes("./child::eTree") : ndxWalk = Nothing
                ' Walk all the children of [ndxThis] from right to left (highest @n number to lowest)
                For intL = ndxPrec.Count - 1 To 0 Step -1
                  ' Check whether the leftmost (lowest) @n of [ndxWork] should precede child number [intL]
                  If (LeftmostWordNumber(ndxWork) < LeftmostWordNumber(ndxPrec(intL))) Then
                    ' Insert [ndxWork]before child number [intL]
                    ndxThis.InsertBefore(ndxWork, ndxPrec(intL))
                    ' Indicate that we have a result and exit the for-loop
                    ndxWalk = ndxPrec(intL) : Exit For
                  End If
                Next intL
                ' Did we find a child where to store [ndxWork]?
                If (ndxWalk Is Nothing) Then
                  ' No appropriate child has been found, 
                  '   which means [ndxWork] should be appended as a child under [ndxThis]
                  ndxThis.AppendChild(ndxWork)
                End If
                '' Keep track of the last id number
                'intIdLast = RightMostWordNumber(ndxThis)
              Case Else
                ' Cannot happen - should not happen - VERIFY
                Stop
            End Select
            ' Check for any unresolved remaining elements 
            '  (these elements are dangling freely under [ndxForest])
            If (bIsLast) AndAlso (intK = ndxChild.Count - 1) Then
              ' Get the consecutive word numbers @n between which there might be a gap
              intIdLeft = RightMostWordNumber(ndxThis)
              intIdRight = pdxSrc.SelectNodes("./descendant::t").Count
              ' Process all that is needed
              While (intIdRight - intIdLeft > 0)
                ' Loop 
                intL = intIdLeft + 1
                ' Find highest <eTree> ancestor of [intL]
                ndxHigh = ndxFor.SelectSingleNode("./descendant::eLeaf[@n=" & intL & "]/ancestor::eTree[last()]")
                ' double check
                If (ndxRoot IsNot ndxHigh) Then
                  ' Find the location between the children of [ndxRoot] where [ndxHigh] may be placed
                  ndxPrec = ndxRoot.SelectNodes("./child::eTree") : ndxWalk = Nothing
                  For intL = 0 To ndxPrec.Count - 1
                    ' Check if the word number (intIdLeft+1) can be placed after child [intL]
                    If (intIdLeft + 1 < RightMostWordNumber(ndxPrec(intL))) Then
                      ' Insert it after child [intL]
                      ndxRoot.InsertAfter(ndxHigh, ndxPrec(intL))
                      ' Indicate we have found an element and leave the for-loop
                      ndxWalk = ndxPrec(intL) : Exit For
                    End If
                  Next intL
                  ' Double check
                  If (ndxWalk Is Nothing) Then
                    ' Just append it
                    Try
                      ndxRoot.AppendChild(ndxHigh)
                    Catch ex As Exception
                      ' Only log errors
                      Logging("OneTigerToPsdxForest warning at forest=" & ndxFor.Attributes("forestId").Value)
                      Stop
                    End Try
                  End If
                Else
                  intIdLeft += 1
                End If
              End While
            End If
          End If
        Next intK
      Next intJ
      '' =========== DEBUG ============
      'If (ndxFor.Attributes("TextId").Value Like "*800155") AndAlso (ndxFor.Attributes("forestId").Value = 22) Then
      '  strBack = "" : TravNode(ndxFor, "Tiger", strBack)
      '  Clipboard.SetText(strBack)
      '  Stop
      'End If
      '' ==============================

      ' Postprocessing: re-arrange the constituents that are not in the correct order
      Do
        bChanged = False
        ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf[@Type='Punct' or @Type='Vern'])>0]")
        For intI = 0 To ndxList.Count - 1
          ' Check if something is missing between me and my predecessor
          If (intI = 0) Then
            ' This is the first one, so check if we start out with "1"
            intIdLeft = 0
          Else
            intIdLeft = RightMostWordNumber(ndxList(intI - 1))
          End If
          intIdRight = LeftmostWordNumber(ndxList(intI))
          ndxThis = ndxList(intI).ParentNode
          ' Is there a gap?
          If (intIdRight - intIdLeft > 1) Then
            ' Try process the first number 
            intL = intIdLeft + 1
            ' Find highest <eTree> ancestor of [intL]
            'ndxHigh = ndxFor.SelectSingleNode("./descendant::eTree[descendant::eLeaf[@Type='Punct' or @Type='Vern'][1]/@n>" & _
            '                                  intIdLeft & " and descendant::eLeaf[@Type='Punct' or @Type='Vern'][last()]/@n < " & intIdRight & "]")
            ndxPrec = ndxFor.SelectNodes("./descendant::eTree[descendant::eLeaf[@Type='Punct' or @Type='Vern'][1]/@n>" & _
                                              intIdLeft & " and descendant::eLeaf[@Type='Punct' or @Type='Vern'][last()]/@n < " & intIdRight & "]")
            ' Select the one with the highest number of <eLeaf/@n> elements
            If (ndxPrec.Count = 1) Then
              ndxHigh = ndxPrec(0)
            Else
              intNum = 0 : ndxHigh = Nothing
              For intM = 0 To ndxPrec.Count - 1
                If (ndxPrec(intM).SelectNodes("./descendant::eLeaf[@n>0]").Count > intNum) Then
                  intNum = ndxPrec(intM).SelectNodes("./descendant::eLeaf[@n>0]").Count
                  ndxHigh = ndxPrec(intM)
                End If
              Next intM
            End If

            ' Validate
            If (ndxHigh Is Nothing) Then
              Stop
            Else
              ' Check if an "ICH" node needs to be created or not...
              If (ndxHigh.ParentNode IsNot ndxThis) Then
                ' Create a new ICH node 
                ndxICH = CreateNewEtree(pdxCurrentFile)
                ndxICH.Attributes("Label").Value = ndxHigh.Attributes("Label").Value
                ' Add an <eLeaf>
                ndxLeaf = AddEleafChild(ndxICH, "Star", "*ICH*-" & intICHid, 0, 0)
                ' Append this node
                ndxHigh.ParentNode.AppendChild(ndxICH)
                ' Adapt the label of the <ndxWork> node
                ndxHigh.Attributes("Label").Value &= "-" & intICHid
                ' Adapt the ICH id counter
                intICHid += 1
              Else
                ndxICH = Nothing
                ' Stop
              End If
              ' Move the node with the actual information to the location we want it to be
              Try
                ndxThis.InsertBefore(ndxHigh, ndxList(intI))
              Catch ex As Exception
                ' There is a problem...
                strBack = "" : TravNode(ndxFor, "Tiger", strBack)
                Debug.Print("Text=" & ndxFor.Attributes("TextId").Value & " forest=" & ndxFor.Attributes("forestId").Value & vbCrLf & _
                            "Seq=" & GetSeq(ndxFor) & vbCrLf & strBack)
                Stop
              End Try
              ' Walk all my ancestors to update the order -- if needed
              If (ndxICH IsNot Nothing) Then
                ndxWalk = ndxICH.ParentNode
                While (ndxWalk IsNot Nothing) AndAlso (ndxWalk.Name = "eTree")
                  ' Check if my leftmost is still *before* my right neighbour
                  ndxNbr = ndxWalk.SelectSingleNode("./following-sibling::eTree")
                  If (ndxNbr IsNot Nothing) Then
                    If (LeftmostWordNumber(ndxNbr) < LeftmostWordNumber(ndxWalk)) Then
                      ' swap places
                      ndxWalk.ParentNode.InsertAfter(ndxWalk, ndxNbr)
                    End If
                  End If
                  ' Go to parent
                  ndxWalk = ndxWalk.ParentNode
                End While
              End If
              ' Note there has been a change
              bChanged = True : Exit For
            End If
          End If
        Next intI
      Loop While (bChanged)
      ' Post processing: look for extra-syntactical elements
      ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf[@Type='Vern'])>0]")
      For intI = 0 To ndxList.Count - 1
        ' Any leaf that is not part of an <nt> does not join in the syntactic analysis...
        ndxOneCh = pdxSrc.SelectSingleNode("./descendant::nt/child::edge[@idref='" & GetFeature(ndxList(intI), "tig", "id") & "']")
        If (ndxOneCh Is Nothing) Then
          ' Give it a feature indicating that it does not join in the syntactic analysis
          AddFeature(pdxCurrentFile, ndxList(intI), "tig", "syn", "no")
        End If
      Next intI
      ' (5) Make sure the sentence is re-analyzed
      ndxWork = Nothing
      eTreeSentence(ndxFor, ndxWork)
      'Debug.Print(NodeText(ndxFor))
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/OneTigerToPsdxForest_ORG error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  Private Function GetSeq(ByRef ndxThis As XmlNode) As String
    Dim intI As Integer
    Dim ndxList As XmlNodeList
    Dim strBack As String = ""

    Try
      If (ndxThis Is Nothing) Then Return "(empty)"
      ndxList = ndxThis.SelectNodes("./descendant-or-self::eLeaf[@Type='Vern' or @Type='Punct']")
      For intI = 0 To ndxList.Count - 1
        If (strBack <> "") Then strBack &= "-"
        strBack &= ndxList(intI).Attributes("n").Value
      Next intI
      Return strBack
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/GetSeq error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return "(error)"
    End Try
  End Function

  '--------------------------------------------------------------------------------------------------------------------------
  ' Name:       BestFitConst()
  ' Goal:       Find the constituent under [ndxWork] that maximally contains words for which:
  '             1. their @n is larger than IdLeft and 
  '             2. their @n is smaller than IdRight
  ' History:
  ' 01-02-2014  ERK Created
  '--------------------------------------------------------------------------------------------------------------------------
  Private Function BestFitConst(ByVal ndxThis As XmlNode, ByVal intIdLeft As Integer, ByVal intIdRight As Integer) As XmlNode
    Dim ndxChild As XmlNodeList ' Children
    Dim ndxFit As XmlNode       ' Result
    Dim intI As Integer         ' Counter

    Try
      ' Check myself first
      If (LeftmostWordNumber(ndxThis) > intIdLeft) AndAlso (RightMostWordNumber(ndxThis) < intIdRight) Then Return ndxThis
      ' Try all children 
      ndxChild = ndxThis.SelectNodes("./child::eTree")
      For intI = 0 To ndxChild.Count - 1
        ' Check this child
        ndxFit = BestFitConst(ndxChild(intI), intIdLeft, intIdRight)
        ' Got it?
        If (ndxFit IsNot Nothing) Then Return ndxFit
      Next intI
      ' Failure...
      Return Nothing
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/BestFitConst error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return Nothing
    End Try
  End Function
  '--------------------------------------------------------------------------------------------------------------------------
  ' Name:       LowestIdNumber()
  ' Goal:       Get the @id attribute value of the leftmost <eLeaf> under <ndxThis>
  ' History:
  ' 01-02-2014  ERK Created
  '--------------------------------------------------------------------------------------------------------------------------
  Private Function LowestIdNumber(ByRef ndxThis As XmlNode) As Integer
    Dim ndxList As XmlNodeList  ' List of nodes
    Dim intLow As Integer       ' Lowest
    Dim intThis As Integer      ' This ID number
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return 0
      ' Get list of <edge> nodes
      ndxList = ndxThis.SelectNodes("./child::edge")
      ' Validate
      If (ndxList Is Nothing) OrElse (ndxList.Count = 0) Then Return 0
      ' Get first number
      intLow = GetTigerIdNumber(ndxList(0).Attributes("idref").Value)
      For intI = 1 To ndxList.Count - 1
        intThis = GetTigerIdNumber(ndxList(intI).Attributes("idref").Value)
        If (intThis < intLow) Then intLow = intThis
      Next intI
      Return intLow
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/LowestIdNumber error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return 0
    End Try
  End Function
  '--------------------------------------------------------------------------------------------------------------------------
  ' Name:       LeftmostWordNumber()
  ' Goal:       Get the @n attribute value of the leftmost <eLeaf> under <ndxThis>
  ' History:
  ' 01-02-2014  ERK Created
  '--------------------------------------------------------------------------------------------------------------------------
  Private Function LeftmostWordNumber(ByRef ndxThis As XmlNode) As Integer
    Dim ndxLeaf As XmlNode  ' One <eLeaf>

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return 0
      ' Get leaf
      ndxLeaf = ndxThis.SelectSingleNode("./descendant-or-self::eLeaf[@n != 0][1]")
      If (ndxLeaf Is Nothing) Then
        Return 0
      Else
        Return ndxLeaf.Attributes("n").Value
      End If
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/LeftmostWordNumber error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return 0
    End Try
  End Function
  '--------------------------------------------------------------------------------------------------------------------------
  ' Name:       RightMostWordNumber()
  ' Goal:       Get the @n attribute value of the RightMost <eLeaf> under <ndxThis>
  ' History:
  ' 01-02-2014  ERK Created
  '--------------------------------------------------------------------------------------------------------------------------
  Private Function RightMostWordNumber(ByRef ndxThis As XmlNode) As Integer
    Dim ndxLeaf As XmlNode  ' One <eLeaf>

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return 0
      ' Get leaf
      ndxLeaf = ndxThis.SelectSingleNode("./descendant-or-self::eLeaf[@n != 0][last()]")
      If (ndxLeaf Is Nothing) Then
        Return 0
      Else
        Return ndxLeaf.Attributes("n").Value
      End If
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/RightMostWordNumber error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return 0
    End Try
  End Function
  '--------------------------------------------------------------------------------------------------------------------------
  ' Name:       TigerNodeGet()
  ' Goal:       Get the node referred to within [ndxThis]
  ' History:
  ' 28-01-2014  ERK Created
  '--------------------------------------------------------------------------------------------------------------------------
  Private Function TigerNodeGet(ByRef ndxThis As XmlNode) As XmlNode
    Dim strId As String   ' The ID we are looking at
    Dim intPos As Integer ' Location within string
    Dim intNum As Integer ' Number after the last period
    Dim bIsNT As Boolean  '
    Dim ndxS As XmlNode   ' sentence

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return Nothing
      ' Action depends on node type
      Select Case ndxThis.Name
        Case "graph"
          strId = ndxThis.Attributes("root").Value
        Case "nt"
          strId = ndxThis.Attributes("id").Value
        Case "t"
          strId = ndxThis.Attributes("id").Value
        Case "edge"
          strId = ndxThis.Attributes("idref").Value
        Case Else
          Return Nothing
      End Select
      ' Get the id number
      intPos = InStrRev(strId, ".")
      If (intPos = 0) Then intPos = InStrRev(strId, "_")
      If (intPos > 0) Then
        intNum = Mid(strId, intPos + 1)
        ' Size of number determins NT or not
        bIsNT = (intNum >= 500)
        ndxS = ndxThis.SelectSingleNode("./ancestor::s[1]")
        If (bIsNT) Then
          Return ndxS.SelectSingleNode("./descendant::nt[@id='" & strId & "']")
        Else
          Return ndxS.SelectSingleNode("./descendant::t[@id='" & strId & "']")
        End If
      End If
      ' Return failre
      Return Nothing
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/TigerNodeGet error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return Nothing
    End Try
  End Function
  '--------------------------------------------------------------------------------------------------------------------------
  ' Name:       GetTigerIdNumber()
  ' Goal:       Strip off the id number from the id string
  ' History:
  ' 31-01-2014  ERK Created
  '--------------------------------------------------------------------------------------------------------------------------
  Public Function GetTigerIdNumber(ByVal strId As String) As Integer
    Dim intPos As Integer ' Location within string
    Dim intNum As Integer ' Number after the last period

    Try
      ' Get the id number
      intPos = InStrRev(strId, ".")
      If (intPos = 0) Then intPos = InStrRev(strId, "_")
      If (intPos > 0) Then
        If (IsNumeric(Mid(strId, intPos + 1))) Then
          intNum = Mid(strId, intPos + 1)
        Else
          Select Case Mid(strId, intPos + 1)
            Case "VROOT"
              intNum = 600
          End Select
        End If
        Return intNum
      End If
      Return -1
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/GetTigerIdNumber error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return -1
    End Try
  End Function
  '--------------------------------------------------------------------------------------------------------------------------
  ' Name:       TigerNodeIsNT()
  ' Goal:       Is the node [ndxThis] from the tiger xml referring to an nt node or a t node?
  ' History:
  ' 28-01-2014  ERK Created
  '--------------------------------------------------------------------------------------------------------------------------
  Private Function TigerNodeIsNT(ByRef ndxThis As XmlNode) As Boolean
    Dim strId As String   ' The ID we are looking at
    Dim intPos As Integer ' Location within string
    Dim intNum As Integer ' Number after the last period

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Action depends on node type
      Select Case ndxThis.Name
        Case "graph"
          strId = ndxThis.Attributes("root").Value
        Case "nt"
          strId = ndxThis.Attributes("id").Value
        Case "t"
          strId = ndxThis.Attributes("id").Value
        Case "edge"
          strId = ndxThis.Attributes("idref").Value
        Case Else
          Return False
      End Select
      ' Get the id number
      intPos = InStrRev(strId, ".")
      If (intPos > 0) Then
        intNum = Mid(strId, intPos + 1)
        ' Size of number determins NT or not
        Return (intNum >= 500)
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/TigerNodeIsNT error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '--------------------------------------------------------------------------------------------------------------------------
  ' Name:       OneTigerNode()
  ' Goal:       Process node [ndxThis] of the tiger file, adding to [ndxFor] of the psdx file
  ' History:
  ' 28-01-2014  ERK Created
  '--------------------------------------------------------------------------------------------------------------------------
  Private Function OneTigerNode(ByRef ndxThis As XmlNode, ByRef ndxFor As XmlNode, _
                                 ByVal strLang As String) As Boolean
    Dim strLabel As String            ' Label of me
    Dim ndxWork As XmlNode = Nothing  ' Node
    Dim ndxChild As XmlNode = Nothing ' One child
    Dim ndxSave As XmlNode = Nothing  ' Additional working node
    Dim ndxList As XmlNodeList        ' List of nodes
    Dim ndxFoll As XmlNode            ' Element before which I should be inserted
    Dim intL As Integer               ' Counter
    Dim intI As Integer               ' Counter
    Dim strFeatId As String           ' Value of feature "id"

    Dim intHeadId As Integer          ' Value of feature "hd"
    Dim intChildId As Integer         ' Value of feature "id" for ndxChild
    Dim strNodePos As String          ' Value of feature "pos"
    Dim strNodeDrel As String         ' Value of feature "drel"
    Dim strPhrase As String           ' Phrase label

    Try
      ' Validate
      If (ndxThis Is Nothing) OrElse (ndxFor Is Nothing) Then Return False
      ' Capture my ID
      strFeatId = GetFeature(ndxThis, "tig", "id")
      intHeadId = GetFeature(ndxThis, "con", "hd")
      '' Select all nodes that have my id as their head
      'ndxList = ndxFor.SelectNodes("./descendant::eTree[child::fs/child::f[@name='hd' and @value='" & intFeatId & "']]")
      '' Get my own label
      'strLabel = ndxThis.Attributes("Label").Value
      'Select Case strLang
      '  Case "nl", "du"
      '    ' Adapt my own label at any rate
      '    ndxThis.Attributes("Label").Value = GetFeature(ndxThis, "con", "pos")
      'End Select
      '' Additional levels depend on the number of nodes having me as head
      'If (ndxList.Count = 0) Then
      '  ' I am not the head of anyone else!
      '  ' Depending on who I am, I may not need an additional head
      '  ' I need to have my POS type
      '  strNodePos = GetFeature(ndxThis, "con", "pos")
      '  ' Safeguard the @con/drel feature
      '  strNodeDrel = GetFeature(ndxThis, "con", "drel")
      '  ' Determine the phrase label 
      '  strPhrase = GetPhraseRuleRev(strNodePos, strNodeDrel)
      '  ' Got anything?
      '  If (strPhrase <> "") Then
      '    ' Insert a level above me
      '    If (Not eTreeInsertLevel(ndxThis, ndxWork)) Then Return False
      '    ' Make [ndxWork] the new [ndxThis]
      '    ndxThis = ndxWork
      '    ' Set the phrase type
      '    ndxThis.Attributes("Label").Value = strPhrase
      '  End If
      'Else
      '  ' If my own @con/id never figures as head, then I don't need an additional level
      '  ' Check if an additional level needs to be made
      '  Select Case strLang
      '    Case "ch", "che"  ' Chechen
      '      ' Action depends on what the label is
      '      Select Case UCase(strLabel)
      '        Case ".", """", ",", "!", "?", ":", ";", "«", "»"   ' Punctuation: don't insert an additional level
      '        Case Else
      '          ' I need to have my POS type
      '          strNodePos = GetFeature(ndxThis, "con", "pos")
      '          ' Safeguard the @con/drel feature
      '          strNodeDrel = GetFeature(ndxThis, "con", "drel")
      '          ' Determine the phrase label 
      '          strPhrase = GetPhraseRule(strNodePos, strNodeDrel)
      '          ' Did we get any result?
      '          If (strPhrase = "") AndAlso (Not DoLike(strNodePos, "C")) Then
      '            ' Double check if this is what I want
      '            ' Stop
      '            Logging("Warning. No phrase rule found for combination:" & vbCrLf & _
      '                    "POS = " & strNodePos & " Drel=" & strNodeDrel)
      '            ' Take a default phrase above me
      '            strPhrase = "PhraseOver_" & strNodePos
      '          End If
      '          If (strPhrase = "CP-REL") Then
      '            ' We need to insert an IP-SUB between the RC and us
      '            If (Not eTreeInsertLevel(ndxThis, ndxWork)) Then Return False
      '            ' ============ DEBUG: Test =====================
      '            If (Not TestEndNodeIdOrder(ndxWork)) Then
      '              Debug.Print("OneTigerNode: bad order")
      '              Logging("modConLLX/OneTigerNode warning: bad constituent order at IP-SUB forestId=" & _
      '                      ndxFor.Attributes("forestId").Value)
      '            End If
      '            ' ==============================================
      '            ' Make [ndxWork] the new [ndxThis]
      '            ndxThis = ndxWork
      '            ' Set the phrase type
      '            ndxThis.Attributes("Label").Value = "IP-SUB"
      '          End If
      '          ' Insert a level above me
      '          If (Not eTreeInsertLevel(ndxThis, ndxWork)) Then Return False
      '          ' ============ DEBUG: Test =====================
      '          If (Not TestEndNodeIdOrder(ndxWork)) Then
      '            Debug.Print("OneTigerNode: bad order")
      '            Logging("modConLLX/OneTigerNode warning: bad constituent order at IP-MAT(?) forestId=" & _
      '                    ndxFor.Attributes("forestId").Value)
      '          End If
      '          ' ==============================================
      '          ' Make [ndxWork] the new [ndxThis]
      '          ndxThis = ndxWork
      '          ' Set the phrase type
      '          ndxThis.Attributes("Label").Value = strPhrase
      '          ' Check if this is a top node
      '          If (intHeadId = 0) AndAlso (strPhrase Like "IP*") Then
      '            ndxThis.Attributes("Label").Value = "IP-MAT"
      '          End If
      '      End Select

      '  End Select
      'End If

      '' Walk all these nodes that have me (intFeatId) as their head
      'For intI = 0 To ndxList.Count - 1
      '  ' Access this node
      '  ndxChild = ndxList(intI)
      '  ' Get its @con/id feature value before we start meddling with it
      '  intChildId = GetFeature(ndxChild, "con", "id")
      '  ' Process this child
      '  If (Not OneTigerNode(ndxChild, ndxFor, strLang)) Then Return False
      '  ' Does [ndxThis] have children already?
      '  If (ndxThis.SelectSingleNode("./child::eTree") Is Nothing) Then
      '    ' Let this become the FIRST child under [ndxThis]
      '    ndxThis.AppendChild(ndxChild)
      '    ' ============ DEBUG: Test =====================
      '    If (Not TestEndNodeIdOrder(ndxThis)) Then Debug.Print("OneTigerNode: bad order")
      '    ' ==============================================
      '  Else
      '    ' Find the first child under [ndxThis] with an @con/id feature value that is 
      '    '   higher than the one of ndxChild
      '    'ndxFoll = ndxThis.SelectSingleNode("./child::eTree[" & _
      '    '          "child::fs[@type='con']/child::f[@name='id' and (@value > " & intChildId & ")]]")
      '    ndxFoll = ndxThis.SelectSingleNode("./child::eTree[" & _
      '              "descendant::fs[@type='con']/child::f[@name='id' and (@value > " & intChildId & ")]]")
      '    ' validate
      '    If (ndxFoll Is Nothing) Then
      '      ' Let this become LAST child under [ndxThis]
      '      ndxThis.AppendChild(ndxChild)
      '      ' ============ DEBUG: Test =====================
      '      If (Not TestEndNodeIdOrder(ndxThis)) Then Debug.Print("OneTigerNode: bad order")
      '      ' ==============================================
      '    Else
      '      ' Insert [ndxChild] before [ndxFoll]
      '      ndxThis.InsertBefore(ndxChild, ndxFoll)
      '      ' ============ DEBUG: Test =====================
      '      If (Not TestEndNodeIdOrder(ndxThis)) Then Debug.Print("OneTigerNode: bad order")
      '      ' ==============================================
      '    End If
      '  End If
      '  '' Possibly adapt the label of the child
      '  'strNodeDrel = GetFeature(ndxChild, "con", "drel")
      '  'ndxChild.Attributes("Label").Value = DrelToLabel(ndxChild.Attributes("Label").Value, strNodeDrel)
      'Next intI
      ' Return positively 
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConLLX/OneTigerNode error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   RefChainCalc
  ' Goal:   Make an XML overview of ALL coreferential chains in this document
  ' History:
  ' 14-03-2014  ERK Derived from modMain/RefListXML()
  ' ------------------------------------------------------------------------------------
  Public Function RefChainCalc(ByRef tdlThis As DataSet, ByRef intNumChain As Integer) As Boolean
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
    Dim strPer As String = ""         ' Period
    Dim colBack As New StringColl     ' Messages about chains with the same chainroot

    Try
      ' Validate: check if a text has been loaded
      If (pdxCurrentFile Is Nothing) Then Return False
      ' Initialise current file
      If (Not InitCurrentFile()) Then Return False
      ' Create an XML data document
      If (Not CreateDataSet("RefChain.xsd", tdlThis)) Then Status("Unable to create a dataset") : Return False
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
        Status("Processing chainlist " & intPtc & "%", intPtc)
        ' Get the forest node
        ndxThis = pdxList(intI)
        ' Perform action recursively backwards
        If (Not TravEtreeBack(ndxThis, colDone, tdlThis)) Then Return False
      Next intI
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
      ' Provide the number of chains
      intNumChain = tdlThis.Tables("Chain").Rows.Count
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/RefChainCalc error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name:   AdelheitToXml
  ' Goal:   Convert improved 5-column Adelheit to its XML format
  ' Note:   Division of the 5 columns is:
  '         1 - Original
  '         2 - Normalisation
  '         3 - Lemma
  '         4 - POS
  ' History:
  ' 01-04-2014  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function AdelheitToXml(ByVal strSrcFile As String, ByVal strDstFile As String) As Boolean
    Dim pdxThis As XmlDocument  ' What we make
    Dim ndxMan As XmlNode       ' Manuscript node
    Dim ndxSep As XmlNode       ' <sep> node
    Dim ndxToken As XmlNode     ' <token> node
    Dim ndxTlp As XmlNode       ' <tlp> node
    Dim strShort As String      ' Short filename
    Dim strOrg As String
    Dim strLemma As String
    Dim strPos As String
    Dim arInput() As String     ' Input file
    Dim arField() As String     ' Line broken up into fields
    Dim intPosition As Integer  ' Position within the string
    Dim intI As Integer         ' Counter
    Dim intLine As Integer = 1  ' Line counter

    Try
      ' Validate
      If (Not IO.File.Exists(strSrcFile)) Then Return False
      ' Get short filename
      strShort = GetShortFileName(strSrcFile)
      ' Create new XmlDocument
      pdxThis = New XmlDocument
      pdxThis.LoadXml("<?xml version='1.0' encoding='utf-8'?>" & vbCrLf & _
                      "<adelheid-document><body>" & vbCrLf & _
                      "<manuscript manid='1' name='" & strShort & "'>" & _
                      "</manuscript></body></adelheid-document>")
      ' Set the namespace manager
      loc_nsDf = New XmlNamespaceManager(pdxThis.NameTable)
      loc_nsDf.AddNamespace("df", pdxThis.DocumentElement.NamespaceURI)
      SetXmlDocument(pdxThis)
      ' Read the input
      arInput = IO.File.ReadAllLines(strSrcFile)
      ' Initialisations
      intPosition = 0 : ndxMan = pdxThis.SelectSingleNode("./descendant::df:manuscript", loc_nsDf)
      ' Walk through the lines
      For intI = 0 To arInput.Length - 1
        ' Do we have anything?
        If (Trim(arInput(intI)) <> "") Then
          ' Break up into fields
          arField = Split(arInput(intI), vbTab)
          strOrg = arField(0) : strLemma = arField(2) : strPos = arField(3)
          ' Skip <manuscript
          If (strOrg <> "<manuscript") Then
            ' Add the information to the xml file
            ' (1) Create a <sep> node
            ndxSep = AddXmlChild(ndxMan, "sep", "TPos", intLine & "/" & intPosition, "attribute", "MSep", "True", "attribute", _
                                 "MForm", " ", "attribute", _
                                  "TSep", "True", "attribute", "ASep", "True", "attribute", _
                                   "Src", "sys", "attribute", "Conf", "0.999", "attribute")
            ' (2) Create a <token> node
            ndxToken = AddXmlChild(ndxMan, "token", "TForm", strOrg, "attribute", "Tag", strPos, "attribute", _
                                   "Lemma", strLemma, "attribute", _
                                   "TPos", intLine & "/" & intPosition & "-" & intPosition + strOrg.Length, "attribute", _
                                   "MForm", strOrg, "attribute", "AForm", strOrg, "attribute", "Conf", "0.999", "attribute")
            ' (3) This <token> gets a <tlp> child
            ndxTlp = AddXmlChild(ndxToken, "tlp", "PTag", strPos, "attribute", "PLemma", strLemma, "attribute", "PProb", "0.999", "attribute")

            ' Keep track of position
            intPosition += strOrg.Length
            ' intLine += 1
            ' Check if this is the end of a sentence
            If (strOrg = "&period;") Then
              ' This is the end of a sentence, so add...
              AddXmlChild(ndxMan, "line")
            End If
          End If
        End If
      Next intI
      ' Save the result
      pdxThis.Save(strDstFile)
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/AdelheitToXml error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       ConvertOneTxtToPsdx()
  ' Goal:       COnvert a text file into a tokenized psdx file
  ' History:
  ' 22-05-2015  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function ConvertOneTxtToPsdx(ByVal strSrcFile As String, ByVal strDstFile As String, _
                                      Optional ByVal strLang As String = "none") As Boolean
    Dim strText As String = ""    ' String of file content
    Dim strBrk As String = ""     ' Break symbol used
    Dim strShort As String = ""   ' Short file name
    Dim strTextId As String = ""  ' Text ID
    Dim strType As String = ""    ' Type of eLeaf
    Dim strValue As String = ""   ' Value of <eLeaf>
    Dim strLabel As String = ""   ' Label value
    Dim strFunct As String = ""   ' Function
    Dim strSent As String = ""    ' One sentence
    Dim arText() As String        ' File broken up into paragraphs
    Dim arSent() As String        ' Paragraph broken up into sentences
    Dim ndxWork As XmlNode        ' Working node
    Dim ndxFor As XmlNode         ' Current forest
    Dim ndxForGrp As XmlNode      ' Pointer to <forestGrp>
    Dim intForestId As Integer = 1  ' ID of the forest
    Dim intEtreeId As Integer = 1 ' ID of <eTree>
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intId As Integer = 0      ' DUmmy
    Dim intParaId As Integer = 0  ' Id of current paragraph

    Try
      ' Read the file into a string
      strText = IO.File.ReadAllText(strSrcFile)
      ' strBrk = GetDelim(strText, vbCrLf, vbCr, vbLf)
      ' Determining linebreak
      strBrk = GetDelimDeep(strText)
      arText = Split(strText, strBrk)
      ' Create a new xml document and set the header to initial values
      If (Not CreatePsdxHeader(pdxConv, strDstFile, strShort)) Then Return False
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
        ' =============== DEBUG =============
        ' If (intI = 32) Then Stop
        ' ===================================
        ' Each element from [arText] is a paragraph.
        ' Split the paragraph into sentences + end-delimiters
        arSent = SentSegment(arText(intI))
        'arSent = Regex.Split(arText(intI), "([^A-ZА-ЯІ](\.+|\?+|\!+|\;)(\s+|$))")
        'Dim mcThis As MatchCollection
        'mcThis = Regex.Matches(arText(intI), "(?<Sentence>\S.+?(?<Terminator>[.!?]|\Z))(?=\s+|\Z)")
        If (arSent IsNot Nothing) Then
          For intJ = 0 To arSent.Length - 1 ' Step 4
            strSent = Trim(arSent(intJ))
            ' double check
            ' If (strSent = "?") Then Stop
            ' Make sure to include only non-empty sentences
            If (strSent <> "" AndAlso Not DoLike(strSent, ".|[?]|[!]|;")) Then
              ' ========== Debug ============
              If (strSent = "") Then Stop
              ' =============================
              '' Include the end-token in the line or not?
              'If (intJ + 1 < arSent.Length) Then
              '  strSent &= arSent(intJ + 1).Trim
              'End If
              ' ========== Debug ============
              If (strSent = "") Then Stop
              ' If (intForestId = 237) Then Stop
              ' =============================
              ' Create a new <forest> element and add its properties
              If (intJ = 0) Then
                ' Start of paragraph
                intParaId += 1
                ndxFor = AddXmlChild(ndxForGrp, "forest", "forestId", intForestId, "attribute", _
                                     "TextId", strTextId, "attribute", _
                                     "File", strShort & ".psdx", "attribute", _
                                     "Location", strTextId & "." & Format(intForestId, "0000"), "attribute", _
                                     "Para", intParaId, "attribute")
              Else
                ndxFor = AddXmlChild(ndxForGrp, "forest", "forestId", intForestId, "attribute", _
                                     "TextId", strTextId, "attribute", _
                                     "File", strShort & ".psdx", "attribute", _
                                     "Location", strTextId & "." & Format(intForestId, "0000"), "attribute")
              End If
              intForestId += 1
              ' Add the <divs>: one for Spanish (org) and one for English (to be added)
              ndxWork = AddXmlChild(ndxFor, "div", "lang", "org", "attribute", "seg", strSent, "child")
              ndxWork = AddXmlChild(ndxFor, "div", "lang", "eng", "attribute", "seg", "", "child")
            End If
          Next intJ
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
      HandleErr("modConvert/ConvertOneTxtToPsdx error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       SentSegment()
  ' Goal:       Convert a string into sentences
  ' History:
  ' 29-05-2015  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function SentSegment(ByRef strIn As String) As String()
    Dim mcThis As MatchCollection       ' Tokenization collection
    Dim lstBack As New List(Of String)  ' List of sentences
    Dim strSent As String               ' One sentence
    Dim strPrev As String               ' Previous value
    Dim sValue As String                ' Current token's value
    Dim intSent As Integer = 0          ' number of the sentence
    Dim intI As Integer                 ' Counter
    Dim objMatch As Match               ' One match

    Try
      ' Validate
      strIn = strIn.Trim
      If (strIn = "") Then Return Nothing
      ' Split string into tokens
      mcThis = Regex.Matches(strIn, "(((\w|\d|[А-ЯІ])+)|\p{P}+|\s+)")
      strSent = "" : strPrev = "" : sValue = ""
      ' Walk all matches
      For intI = 0 To mcThis.Count - 1
        ' Get this new match
        objMatch = mcThis(intI)
        ' Get current value and store previous value
        strPrev = sValue : sValue = objMatch.Value
        ' Is it a potential end-of-sentence token?
        Select Case sValue
          Case "."
            ' Potential end-of-sentence: check
            ' Stop
            ' Check the previous token
            If (strPrev.Length = 1) Then
              ' This must be an abbreviation
              strSent &= sValue
            Else
              ' This must be sentence-end
              strSent &= sValue
              ' store it
              lstBack.Add(strSent.Trim) : strSent = "" : strPrev = "" : sValue = ""
            End If
          Case "..", "...", "?", "??", "???", "!", "!!", "!!!", "?!", "!?", ":", ";"
            ' This must be end-of-sentence
            strSent &= sValue
            ' store it
            lstBack.Add(strSent.Trim) : strSent = "" : strPrev = "" : sValue = ""
          Case " "
            ' So it was a space: skip
            strSent &= sValue
          Case Else
            ' Check for spaces
            If (objMatch.Value.Trim = "") Then
              strSent &= sValue
            Else
              ' Default behaviour: word
              strSent &= sValue
            End If
        End Select
      Next intI
      ' Do we need more adding?
      If (strSent <> "") Then lstBack.Add(strSent.Trim)

      ' Return the list of sentences as array
      Return lstBack.ToArray
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/SentSegment error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return Nothing
    End Try
  End Function
  '----------------------------------------------------------------------------------------
  ' Name:       OneTokenizePsdx()
  ' Goal:       Perform tokenization on a psdx text and save the result
  ' History:
  ' 25-05-2015  ERK Created
  '----------------------------------------------------------------------------------------
  Public Function OneTokenizePsdx(ByVal strSrcFile As String) As Boolean
    Dim strLang As String   ' Language of this file
    Dim strSent As String   ' One sentence
    Dim ndxList As XmlNodeList
    Dim ndxTree As XmlNodeList
    Dim ndxFor As XmlNode
    Dim intI As Integer     ' Counter
    Dim intJ As Integer     ' Counter

    Try
      ' Validate
      If (Not IO.File.Exists(strSrcFile)) Then Logging("Tokenization: cannot open file") : Return False
      ' Read the file into memory
      pdxConv = New XmlDocument
      pdxConv.Load(strSrcFile)
      ' Make sure subsequent operations are done on this PDX document
      SetXmlDocument(pdxConv)
      ' Get the language
      strLang = GetFileDesc(pdxConv, "ident", "language")
      If (strLang = "") Then Logging("Tokenization: no language specified") : Return False
      ' Walk through all forest elements
      ndxList = pdxConv.SelectNodes("./descendant::forest")
      For intI = 0 To ndxList.Count - 1
        ' Get the forest
        ndxFor = ndxList.Item(intI)
        ' Get the sentence
        strSent = ndxFor.SelectSingleNode("./descendant::div[@lang='org']/seg").InnerText
        ' Delete any <eTree> children it has
        ndxTree = ndxFor.SelectNodes("./child::eTree")
        For intJ = ndxTree.Count - 1 To 0 Step -1
          ' Remove any descendants of me
          ndxTree(intJ).RemoveAll()
          ndxFor.RemoveChild(ndxTree(intJ))
        Next intJ

        ' Process this <forest>
        If (Not OneTextToPsdxForest(ndxFor, pdxConv, strSent, strLang, True)) Then Return False
      Next intI
      ' Save the result
      pdxConv.Save(strSrcFile)
      pdxCurrentFile = pdxConv
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/OneTokenize error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function

  '----------------------------------------------------------------------------------------
  ' Name:       OneTextToPsdxForest()
  ' Goal:       Process one sentence into a psdx <forest> element
  ' History:
  ' 23-05-2015  ERK Created
  '----------------------------------------------------------------------------------------
  Private Function OneTextToPsdxForest(ByRef ndxFor As XmlNode, ByRef pdxSrc As XmlDocument, _
    ByRef strThisLine As String, ByVal strLang As String, Optional ByVal bAuto As Boolean = False) As Boolean
    Dim ndxWork As XmlNode      ' Working node
    Dim intSt As Integer        ' Start of line
    Dim intEn As Integer        ' end of line
    Try
      ' Validate
      If (ndxFor Is Nothing) OrElse (pdxSrc Is Nothing) Then Return False
      ' (1) Break the line into words
      intSt = 1 : intEn = 1
      ' ======= DEBUG
      ' If (ndxFor.Attributes("forestId").Value = 11) Then Stop
      ' ================
      If (Not DoWords(ndxFor, strThisLine, intSt, intEn, bAuto)) Then Status("modMain/ParsePosLine error") : Return False
      ' (2) Make sure the sentence is re-analyzed
      ndxWork = Nothing
      eTreeSentence(ndxFor, ndxWork)
      ' Return success
      Return True
    Catch ex As Exception
      ' Give error
      HandleErr("modConvert/OneTextToPsdxForest error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failre
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoAddFileDesc
  ' Goal:   Add description to TEI header
  ' History:
  ' 05-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function DoAddFileDesc(ByVal strType As String, ByVal strValue As String, Optional ByVal strSub As String = "") As Boolean
    Return DoAddFileDesc(pdxCurrentFile, strType, strValue, strSub)
  End Function
  Public Function DoAddFileDesc(ByRef pdxThis As XmlDocument, ByVal strType As String, ByVal strValue As String, Optional ByVal strSub As String = "") As Boolean
    Dim ndxThis As XmlNode = Nothing ' Working node
    ' Dim ndxMain As XmlNode

    Try
      ' Validate
      ' If (Not bInit) Then Return False
      ' Now the action depends on the specific element we are adding
      If (strSub = "") Then
        ' Get the sub parth
        Select Case strType
          Case "title", "author", "editor"
            ' Get the correct master node
            ndxThis = SetNodePath(pdxThis, "teiHeader/fileDesc")
            ' Get a good return?
            If (ndxThis Is Nothing) Then Return False
            strSub = "titleStmt"
          Case "distributor"
            ' Get the correct master node
            ndxThis = SetNodePath(pdxThis, "teiHeader/fileDesc")
            ' Get a good return?
            If (ndxThis Is Nothing) Then Return False
            strSub = "publicationStmt"
          Case "bibl"
            ' Get the correct master node
            ndxThis = SetNodePath(pdxThis, "teiHeader/fileDesc")
            ' Get a good return?
            If (ndxThis Is Nothing) Then Return False
            strSub = "sourceDesc"
          Case "manuscript", "original", "subtype", "genre"
            ' Get the correct master node
            ndxThis = SetNodePath(pdxThis, "teiHeader/profileDesc/langUsage")
            strSub = "creation"
          Case Else
            ' THis is not possible!
            Return False
        End Select
      Else
        Select Case strSub
          Case "language"
            ' Get the correct master node
            ndxThis = SetNodePath(pdxThis, "teiHeader/profileDesc/langUsage")
          Case Else
            Return False
        End Select
      End If
      ' Add the appropriate child and attribute
      If (SetXmlNodeChild(pdxThis, ndxThis, strSub, "", strType, strValue) Is Nothing) Then Return False
      ' Set the dirty bit
      frmMain.SetDirty(True)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("frmMain/DoAddFileDesc error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Failure...
      Return False
    End Try
  End Function

End Module
