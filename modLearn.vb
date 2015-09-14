Imports System.Xml
Module modLearn
  ' ========================================= LOCAL VARIABLES ===========================================
  Private loc_pdxNp As XmlNodeList            ' Selection of all IP's
  Private loc_colGrRoleRel As New StringColl  ' All the relations we encounter, and their frequency
  Private loc_colNPtypeRel As New StringColl  ' All the relations we encounter, and their frequency
  Private loc_colGrRoleSrc As New StringColl  ' All the grammatical roles that can serve as SOURCE
  Private loc_colGrRoleDst As New StringColl  ' All the grammatical roles that can serve as TARGET
  Private loc_colNPtypeSrc As New StringColl  ' All NP types that serve as SOURCE
  Private loc_colNPtypeDst As New StringColl  ' All NP types that serve as TARGET
  Private tblNPfeat As DataTable              ' Table with features
  ' =====================================================================================================
  ' ------------------------------------------------------------------------------------
  ' Name:   DoLearnCoref
  ' Goal:   Learn from previous coreferencing
  ' History:
  ' 10-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function DoLearnCoref(ByVal strInFile As String) As Boolean
    Dim ndxThis As XmlNode    ' Current NP under investigation
    Dim ndxForest As XmlNode  ' The forest for this line
    Dim ndxTarget As XmlNode  ' Destination NP
    Dim strRefType As String  ' The kind of coreference relation (if any)
    Dim strGrRole As String   ' The grammatical role of the source and destination
    Dim strNpType As String   ' The NP type of the source and the destination
    Dim strRelation As String ' NP or GrRole relation between source and target
    Dim intPtc As Integer     ' Percentage
    Dim intNp As Integer      ' Counter for the NP's within one <forest> line
    Dim intF As Integer       ' One forest item

    Try
      ' Validate
      If (Not IO.File.Exists(strInFile)) Then
        Status("Could not find file " & strInFile)
        Return False
      End If
      ' TODO: OTHER INITIALISATIONS
      ' Show we are reading
      Status("Reading file " & strInFile)
      ' Try read this file into an XML structure
      If (ReadXmlDoc(strInFile, pdxCurrentFile)) Then
        ' Initialise [pdxList] and [arLeftId()]
        If (Not InitCurrentFile()) Then Return False
        ' Go through all the [pdxList] items
        For intF = 0 To pdxList.Count - 1
          ' Keep track of where we are
          intPtc = (intF + 1) * 100 \ pdxList.Count
          Status("Learning coreference information " & IO.Path.GetFileNameWithoutExtension(strInFile) & " " & intPtc & "%", intPtc)
          ' Get the forestnode of this line
          ndxForest = pdxList(intF)
          ' Find the NP's in this line
          ' Retrieve the Noun Phrases and PRO$ elements to be reviewed
          loc_pdxNp = ndxForest.SelectNodes(".//eTree[tb:Like(string(@Label),'" & strNPsourceTypes & "')]")
          'loc_pdxNp = ndxForest.SelectNodes(".//eTree[@Label='NP' or " & _
          '    " starts-with(@Label, 'NP-') or starts-with(@Label, 'PRO$') ]")
          ' Visit all the NP elements in this forest one by one
          For intNp = 0 To loc_pdxNp.Count - 1
            ' Get this NP
            ndxThis = loc_pdxNp(intNp)
            ' Find out whether this NP is referring to something
            strRefType = GetFeature(ndxThis, "coref", "RefType")
            ' Action depends on the reference type
            Select Case strRefType
              Case "Identity", "Inferred", "CrossSpeech"
                ' Get the destination node
                ndxTarget = CorefDst(ndxThis)
                ' We want to keep track of source/target grammatical role
                If (GetFeature(ndxThis, "NP", "GrRole") = "") Then
                  strGrRole = Labelmain(ndxThis.Attributes("Label").Value)
                Else
                  strGrRole = GetFeature(ndxThis, "NP", "GrRole")
                End If
                ' Process the source GrRole
                loc_colGrRoleSrc.AddUnique(strGrRole)
                ' Build the relation
                strRelation = strGrRole & " > "
                ' Determine the destination GrRole
                If (GetFeature(ndxTarget, "NP", "GrRole") = "") Then
                  strGrRole = Labelmain(ndxTarget.Attributes("Label").Value)
                Else
                  strGrRole = GetFeature(ndxTarget, "NP", "GrRole")
                End If
                ' Store the target
                loc_colGrRoleDst.AddUnique(strGrRole)
                ' Adapt the relation
                strRelation &= strGrRole
                ' Store this combination, while keeping track of its frequency
                loc_colGrRoleRel.AddUnique(strRelation)
                ' We want to keep track of source/target NP type
                If (GetFeature(ndxThis, "NP", "NPtype") = "") Then
                  strNpType = Labelmain(ndxThis.Attributes("Label").Value)
                Else
                  strNpType = GetFeature(ndxThis, "NP", "NPtype")
                End If
                ' Store the source NPtype
                loc_colNPtypeSrc.AddUnique(strNpType)
                ' Keep track of the relation
                strRelation = strNpType & " > "
                ' Find the target NPtype
                If (GetFeature(ndxTarget, "NP", "NPtype") = "") Then
                  strNpType = Labelmain(ndxTarget.Attributes("Label").Value)
                Else
                  strNpType = GetFeature(ndxTarget, "NP", "NPtype")
                End If
                ' Store the target NPtype
                loc_colNPtypeDst.AddUnique(strNpType)
                ' Build the relation further
                strRelation &= strNpType
                ' Store this combination, while keeping track of its frequency
                loc_colNPtypeRel.AddUnique(strRelation)
              Case strRefNew, strRefAssumed
                ' No need to do anything
              Case Else
                ' Not referential, or unrecognized type -- don't count
            End Select
          Next intNp
        Next intF
      End If
      ' TODO: any clearing up
      ' Default: return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modLearn/DoLearnCoref error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   LearnReport
  ' Goal:   Return the report of learning
  ' History:
  ' 10-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function LearnReport() As String
    Dim strHtml As String = ""  ' The HTML result

    Try
      ' Validate
      If (tdlSettings Is Nothing) Then Return ""
      ' Initialise the NP datatable
      tblNPfeat = tdlSettings.Tables("NPfeat")
      ' Validate
      If (tblNPfeat Is Nothing) Then Return ""
      ' Make the HTML result
      strHtml = "<html><body><h1>Coref learning results</h1>" & vbCrLf & _
        "<h2>Not occurring grammatical role matches</h2>" & vbCrLf & _
        "<p>" & GrRoleMissing() & "<p>" & vbCrLf & _
        "<h2>Not occurring NP type matches</h2>" & vbCrLf & _
        "<p>" & NPtypeMissing() & "<p>" & vbCrLf & _
        "<h2>Grammatical role Sources</h2>" & vbCrLf & _
        "<p>" & loc_colGrRoleSrc.FreqTable & "<p>" & vbCrLf & _
        "<h2>Grammatical role Targets</h2>" & vbCrLf & _
        "<p>" & loc_colGrRoleDst.FreqTable & "<p>" & vbCrLf & _
        "<h2>Grammatical role relations</h2>" & vbCrLf & _
        "<p>" & loc_colGrRoleRel.FreqTable & "<p>" & vbCrLf & _
        "<h2>NP type Sources</h2>" & vbCrLf & _
        "<p>" & loc_colNPtypeSrc.FreqTable & "<p>" & vbCrLf & _
        "<h2>NP type Targets</h2>" & vbCrLf & _
        "<p>" & loc_colNPtypeDst.FreqTable & "<p>" & vbCrLf & _
        "<h2>NP type relations</h2>" & vbCrLf & _
        "<p>" & loc_colNPtypeRel.FreqTable & "<p>" & vbCrLf & _
        "</body></html>" & vbCrLf
      ' Return this result
      Return strHtml
    Catch ex As Exception
      ' Show error
      HandleErr("modLearn/LearnReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   NPtypeMissing
  ' Goal:   Find out which of the NPtype pairs are missing
  ' History:
  ' 10-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function NPtypeMissing() As String
    Dim strSrc As String      ' The source NP type
    Dim strDst As String      ' The destination type
    Dim strSearch As String   ' Search string
    Dim intI As Integer       ' One counter
    Dim intJ As Integer       ' Second counter
    Dim arNPtype() As String  ' Array with NP types
    Dim colMis As New StringColl

    Try
      ' Get the NPtype variants
      arNPtype = GetFeatureArray("NPtype")
      ' Walk all source types
      For intI = 0 To UBound(arNPtype)
        ' Get the source type
        strSrc = arNPtype(intI)
        ' Walk all destination ones
        For intJ = 0 To UBound(arNPtype)
          ' Get the destination type
          strDst = arNPtype(intJ)
          ' Combine into a search string
          strSearch = strSrc & " > " & strDst
          ' Find out if this string is present in the collection
          If (strSrc <> "") AndAlso (strDst <> "") AndAlso (Not loc_colNPtypeRel.Exists(strSearch)) Then
            ' Add to the list
            colMis.AddUnique(strSearch)
          End If
        Next intJ
      Next intI
      ' Return the missing items
      Return colMis.Text.Replace(vbCrLf, "<br>")
    Catch ex As Exception
      ' Show error
      HandleErr("modLearn/NPtypeMissing error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GrRoleMissing
  ' Goal:   Find out which of the GrRole pairs are missing
  ' History:
  ' 10-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GrRoleMissing() As String
    Dim strSrc As String      ' The source NP type
    Dim strDst As String      ' The destination type
    Dim strSearch As String   ' Search string
    Dim intI As Integer       ' One counter
    Dim intJ As Integer       ' Second counter
    Dim arGrRole() As String  ' Array with NP types
    Dim colMis As New StringColl

    Try
      ' Get the GrRole variants
      arGrRole = GetFeatureArray("GrRole")
      ' Validate
      If (arGrRole Is Nothing) Then Return ""
      ' Walk all source types
      For intI = 0 To UBound(arGrRole)
        ' Get the source type
        strSrc = Trim(arGrRole(intI))
        ' Walk all destination ones
        For intJ = 0 To UBound(arGrRole)
          ' Get the destination type
          strDst = Trim(arGrRole(intJ))
          ' Combine into a search string
          strSearch = strSrc & " > " & strDst
          ' Find out if this string is present in the collection
          If (strSrc <> "") AndAlso (strDst <> "") AndAlso (Not loc_colGrRoleRel.Exists(strSearch)) Then
            ' Add to the list
            colMis.AddUnique(strSearch)
          End If
        Next intJ
      Next intI
      ' Return the missing items
      Return colMis.Text.Replace(vbCrLf, "<br>")
    Catch ex As Exception
      ' Show error
      HandleErr("modLearn/GrRoleMissing error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetFeatureArray
  ' Goal:   Find an array of features in the table
  ' History:
  ' 10-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetFeatureArray(ByVal strFeatName As String) As String()
    Dim intI As Integer       ' One counter
    Dim intJ As Integer       ' Counter
    Dim strVariants As String ' ALl variants
    Dim arThis() As String    ' Array to be returned
    Dim arOne() As String     ' Array of one element

    Try
      ' Validate
      If (tblNPfeat Is Nothing) Then
        ' Show user that there is a problem
        Logging("modLearn/GetFeatureArray: the table [tblNPfeat] is missing")
        ' Return failure
        Return Nothing
      End If
      ' Try find the correct member
      For intI = 0 To tblNPfeat.Rows.Count - 1
        ' Is this the correct feature we are looking for?
        If (tblNPfeat.Rows(intI).Item("Name") = strFeatName) Then
          ' We've got it!
          strVariants = tblNPfeat.Rows(intI).Item("Variants")
          ' Turn it into an array
          arThis = Split(strVariants, GetDelim(strVariants, vbCrLf, vbCr, vbLf))
          ' Adapt each entry
          For intJ = 0 To UBound(arThis)
            ' Adapt this one
            arOne = Split(arThis(intJ), ";")
            arThis(intJ) = arOne(0)
          Next intJ
          ' Return the result
          Return arThis
        End If
      Next intI
      ' Return failre
      Return Nothing
    Catch ex As Exception
      ' Show error
      HandleErr("modLearn/GetFeatureArray error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitLearn
  ' Goal:   Initialise the learning collection
  ' History:
  ' 28-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub InitLearn()
    Try
      ' Initialise it
      loc_colNPtypeRel.Clear()
      loc_colGrRoleRel.Clear()
      loc_colGrRoleSrc.Clear()
      loc_colGrRoleDst.Clear()
      loc_colNPtypeSrc.Clear()
      loc_colNPtypeDst.Clear()
    Catch ex As Exception
      ' Show error
      HandleErr("modLearn/InitLearn error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Module
