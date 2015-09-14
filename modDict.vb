Imports System.Xml
Module modDict
  ' ================================== LOCAL VARIABLES =========================================================
  Private bInit As Boolean = False
  Private bDirty As Boolean = False
  ' =================================LOCAL CONSTANTS============================================================
  ' ============================================================================================================
  '---------------------------------------------------------------------------------------------------------
  ' Name:       TryReadChainDict()
  ' Goal:       Try to read the chain dictionary dataset
  ' History:
  ' 06-07-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Sub TryReadChainDict()
    Try
      ' We should also try to read the period definitions...
      If (strChainDict = "") Then
        ' Log our complaints
        Logging("No chain dictionary file has been specified.")
      Else
        ' Does the file exist?
        If (IO.File.Exists(strChainDict)) Then
          ' Show what we are doing
          Status("Reading chain dictionary definitions...")
          ' Read the periods
          If (Not ReadDataset(strSchemaCdict, strChainDict, tdlChainDict)) Then
            ' Log the complaint
            Logging("Unable to read valid dataset from chain dictionary file " & strChainDict)
          End If
          ' Show we are ready
          Status("Ready")
        Else
          ' Log the complaint
          Logging("Creating new chain dictionary file: " & strChainDict)
          ' We need to start a new chain dictionary with this file
          If (Not CreateDataSet(strSchemaCdict, tdlChainDict)) Then
            ' Log the complaint
            Logging("Unable to create a new Chain Dictionary dataset")
            Status("Unable to create a new Chain Dictionary dataset")
            ' Make sure no dataset is specified
            tdlChainDict = Nothing
          End If
        End If
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modDict/TryReadChainDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :      ChainDictProcess
  ' Goal:       Add all necessary links from src to antecedent(s)
  ' History:
  ' 06-07-2010  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Sub ChainDictProcess(ByVal strSrc As String, ByVal strAntHead As String, ByVal strAntEqual As String)
    Dim intI As Integer     ' Counter
    Dim arEqual() As String ' Array of equal heads

    Try
      ' Validate
      If (tdlChainDict Is Nothing) OrElse (strSrc = "") Then Exit Sub
      ' Is there any antecedent head?
      If (strAntHead <> "") Then
        ' Make the relation with the head
        If (Not ChainDictAdd(strSrc, strAntHead)) Then Exit Sub
      End If
      ' Get all the elements of [Equal]
      arEqual = Split(strAntEqual, ";")
      For intI = 0 To UBound(arEqual)
        ' Process this relation
        If (Not ChainDictAdd(strSrc, arEqual(intI))) Then Exit Sub
      Next intI
    Catch ex As Exception
      ' Show error
      HandleErr("modDict/ChainDictProcess error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :      TrySaveChainDict
  ' Goal:       Try to save the chain dictionary dataset
  ' History:
  ' 06-07-2010  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Sub TrySaveChainDict()
    Try
      ' Validate + only save when needed!!!
      If (tdlChainDict Is Nothing) OrElse (Not bDirty) Then Exit Sub
      If (strChainDict = "") Then Exit Sub
      ' Save the dataset to the indicated location
      tdlChainDict.WriteXml(strChainDict)
      ' Reset the dirty flag
      bDirty = False
    Catch ex As Exception
      ' Show error
      HandleErr("modDict/TrySaveChainDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :      ChainDictSetDirty
  ' Goal:       Make sure the chain dictionary gets saved...
  ' History:
  ' 06-07-2010  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Sub ChainDictSetDirty()
    bDirty = True
  End Sub
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :      ChainDictAdd
  ' Goal:       Store a combination of source > antecedent (period) into the chain dictionary
  ' History:
  ' 06-07-2010  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function ChainDictAdd(ByVal strSrc As String, ByVal strAnt As String) As Boolean
    Dim dtrNew As DataRow   ' Room for the new element

    Try
      ' Validate
      If (tdlChainDict Is Nothing) OrElse (strSrc = "") Or (strAnt = "") Then Return False
      ' Only accept source>ant, if they are NOT equal
      If (strSrc = strAnt) Then Return True
      ' Check if the entry already exists
      If (ChainDictHas(strSrc, strAnt)) Then Return True
      ' Create a new datarow
      With tdlChainDict.Tables("entry")
        dtrNew = .NewRow
        With dtrNew
          .Item("form") = strSrc
          .Item("ant") = strAnt
          .Item("period") = strCurrentPeriod
          .Item("user") = strUserName
          .Item("date") = Format(Now, "g")
        End With
        ' Add this new row
        .Rows.Add(dtrNew)
      End With
      ' Set the dirty flag
      bDirty = True
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modDict/ChainDictAdd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return faiulre
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :      ChainDictHas
  ' Goal:       Verify whether [strSRc] has [strAnt] as antecedent in the current period
  ' History:
  ' 06-07-2010  ERK Created
  ' 15-09-2011  ERK There is a problem with the LIKE operator and '*ICH*-1'
  ' ----------------------------------------------------------------------------------------------------------
  Public Function ChainDictHas(ByVal strSrc As String, ByVal strAnt As String, _
                               Optional ByVal strAntEqual As String = "") As Boolean
    Dim dtrFound() As DataRow ' Result of select statement
    Dim arEqual() As String   ' Other antecedents
    Dim strEqual As String    ' Temporary string (for speed)
    Dim strSelect As String   ' The select string
    Dim intI As Integer       ' Counter

    Try
      ' Validate
      If (tdlChainDict Is Nothing) OrElse (strSrc = "") Or (strAnt = "") Then Return False
      ' Validation: make sure no forms with a star are compared
      If (InStr(strSrc, "*") > 0) OrElse (InStr(strAnt, "*") > 0) Then Return False
      ' Access the <entry> table
      With tdlChainDict.Tables("entry")
        ' Adapt strings for the select statement
        strSrc = AdaptForSelect(strSrc) : strAnt = AdaptForSelect(strAnt)
        strSelect = "form LIKE '" & strSrc & "' AND ant LIKE '" & strAnt & "' AND period='" & strCurrentPeriod & "'"
        ' Try find
        dtrFound = .Select(strSelect)
        If (dtrFound.Length > 0) Then Return True
        ' Check out other possible antecedents
        arEqual = Split(strAntEqual, ";")
        For intI = 0 To UBound(arEqual)
          ' Adapt strings for the select statement
          strEqual = AdaptForSelect(arEqual(intI))
          ' Double check asterisks
          If (InStr(strEqual, "*") = 0) Then
            strSelect = "form LIKE '" & strSrc & "' AND ant LIKE '" & strEqual & "' AND period='" & strCurrentPeriod & "'"
            ' Try find
            dtrFound = .Select(strSelect)
            If (dtrFound.Length > 0) Then Return True
          End If
        Next intI
      End With
      ' Default: return failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modDict/ChainDictHas error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return faiulre
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :      AdaptForSelect
  ' Goal:       The SELECT statement has some peculiarities, so we need to adapt...
  ' History:
  ' 06-07-2010  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Private Function AdaptForSelect(ByVal strIn As String) As String
    Dim intPos As Integer   ' Position within string
    Try
      ' Validate
      If (strIn = "") Then Return ""
      ' See if any replacement is needed
      intPos = InStr(strIn, "'")
      If (intPos > 0) Then
        ' Replace "'" signs
        strIn = strIn.Replace("'", "*")
        ' Double check
        If (intPos < strIn.Length) Then
          ' We need to adapt it
          strIn = Left(strIn, InStr(strIn, "*"))
        End If
      End If
      ' Return the result 
      Return strIn
    Catch ex As Exception
      ' Show error
      HandleErr("modDict/AdaptForSelect error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return faiulre
      Return ""
    End Try
  End Function
End Module
