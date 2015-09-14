Imports System.Xml
Module modCoref
  ' ------------------------------------------------------------------------------------
  ' Name:   TravEtreeBack
  ' Goal:   Recursive function to go backwards, finding <eTree> elements and acting upon them
  ' History:
  ' 16-08-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function TravEtreeBack(ByRef ndxThis As XmlNode, ByRef colDone As NodeColl, ByRef tdlThis As DataSet) As Boolean
    Dim ndxChild As XmlNode ' Child node

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      ' Check for interrupt
      If (bInterrupt) Then Return False
      ' First check all its children - last to first
      ndxChild = ndxThis.LastChild
      While (Not ndxChild Is Nothing)
        ' Process this child
        If (Not TravEtreeBack(ndxChild, colDone, tdlThis)) Then Return False
        ndxChild = ndxChild.PreviousSibling
      End While
      ' Check what we have
      Select Case ndxThis.Name
        Case "eTree"
          ' Perform the action for this item
          If (Not MakeOneChain(ndxThis, colDone, tdlThis)) Then Return False
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modCoref/TravEtreeBack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   MakeOneChain
  ' Goal:   See if this is a chain, and add it to [tdlThis]
  ' History:
  ' 10-08-2011  ERK Created
  ' 09-02-2011  ERK Added RootEtreeId in <Chain>
  ' ------------------------------------------------------------------------------------
  Private Function MakeOneChain(ByRef ndxThis As XmlNode, ByRef colDone As NodeColl, _
                              ByRef tdlThis As DataSet) As Boolean
    Dim ndxWork As XmlNode              ' Working node
    Dim ndxForest As XmlNode = Nothing  ' Forest node
    Dim intLen As Integer = 0           ' Length of this chain
    Dim intNoTraceLen As Integer = 0    ' Length of this chain excluding traces
    Dim strLastText As String = ""      ' Last node text
    Dim strRefType As String            ' The kind of reference we have here
    Dim dtrChain As DataRow = Nothing   ' Datarow for this chain
    Dim strAnim As String = ""          ' Animacy of this chain
    Dim intChainId As Integer = 0       ' ID for this chain
    Dim dtrItem As DataRow = Nothing    ' One datarow for an item
    Dim intItemId As Integer = 0        ' ID for this item
    Dim intIPmatId As Integer           ' the @Id of the IP under which we are 
    Dim intIPclsId As Integer           ' the @Id of the IPmat under which we are 
    Dim bStop As Boolean = False        ' Flag to stop the chain

    Try
      ' Validate
      If (ndxThis Is Nothing) OrElse (colDone Is Nothing) Then Return False
      ' Action depends on the reference type of this node\
      strRefType = CorefAttr(ndxThis, "RefType")
      Select Case strRefType
        Case "NewVar", "Inert"
          ' These are purely hypothetical, theoretical nodes, so we exclude them
        Case "New", "Assumed", "Identity", "CrossSpeech", "Inferred"
          ' Check if this node has already been processed
          If (colDone.Exists(ndxThis)) Then Return True
          ' This node starts a chain!
          intChainNum += 1
          ' (1) Start XML output for this chain
          dtrChain = AddOneDataRow(tdlThis, "Chain", "ChainId", "ChainList")
          intChainId = dtrChain.Item("ChainId")
          ' See if we have the animacy feature for this chain
          strAnim = GetFeature(ndxThis, "NP", "anim")
          dtrChain.Item("Animacy") = IIf(strAnim = "", "-", strAnim)
          ' Go into a loop
          ndxWork = ndxThis
          While (Not ((ndxWork Is Nothing) OrElse (colDone.Exists(ndxWork)) OrElse bStop))
            ' Increment the length of this chain
            intLen += 1
            ' Give information about this one
            strLastText = NodeText(ndxWork, True)
            ' Get the forestnode of me
            If (Not GetForestNode(ndxWork, ndxForest)) Then Return False
            ' Get the @Id of the IPmat under which we are 
            intIPclsId = GetIpMatOrSubId(ndxWork, False)
            intIPmatId = GetIpMatOrSubId(ndxWork, True)
            ' Store the XML information
            If (Not CreateNewRow(tdlThis, "Item", "ItemId", intItemId, dtrItem)) Then Return False
            With dtrItem
              .SetParentRow(dtrChain)
              .Item("ChainId") = intChainId
              .Item("eTreeId") = ndxWork.Attributes("Id").Value
              .Item("forestId") = ndxForest.Attributes("forestId").Value
              .Item("IPmatId") = intIPmatId
              .Item("IPclsId") = intIPclsId
              .Item("IPclsLb") = NodeLabel(IdToNode(intIPclsId))
              If (ndxWork.Attributes("IPnum") Is Nothing) Then
                .Item("IPnum") = .Item("forestId")
              Else
                .Item("IPnum") = ndxWork.Attributes("IPnum").Value
              End If
              .Item("SbjRefAll") = "False"
              .Item("SbjRefMat") = "False"
              .Item("SbjRefMat3") = "False"
              .Item("Loc") = NodeLocation(ndxWork)
              .Item("RefType") = CorefAttr(ndxWork, "RefType")
              .Item("Syntax") = ndxWork.Attributes("Label").Value
              .Item("NPtype") = GetFeature(ndxWork, "NP", "NPtype")
              .Item("GrRole") = GetFeature(ndxWork, "NP", "GrRole")
              .Item("Node") = strLastText
              ' Set initial (default) values of distances
              .Item("DistIPnum") = 0 : .Item("DistForest") = 0
              ' Set PGN value, if known
              .Item("PGN") = GetFeature(ndxWork, "NP", "PGN")
              ' Adapt the NoTraceLen
              If (.Item("NPtype") <> "Trace") Then intNoTraceLen += 1
            End With
            ' File this node
            colDone.AddUnique(ndxWork)
            ' Note our current Reference type
            strRefType = CorefAttr(ndxWork, "RefType")
            ' Calculate whether we need to stop
            bStop = (Not DoLike(strRefType, "Identity|CrossSpeech"))
            ' Check if the Chain number is present as feature
            If (HasFeature(ndxWork, "coref", "ChainId")) Then
              ' Check the value of the feature
              If (GetFeature(ndxWork, "coref", "ChainId") <> intChainId) Then
                ' Change the value of the feature
                If (Not AddFeature(pdxCurrentFile, ndxWork, "coref", "ChainId", intChainId)) Then Return False
                ' Indicate we have changed
                bNeedSaving = True
              End If
            Else
              ' Add the chain number as a Coref type feature
              If (Not AddFeature(pdxCurrentFile, ndxWork, "coref", "ChainId", intChainId)) Then Return False
              ' Indicate we have changed
              bNeedSaving = True
            End If
            ' Go to the next antecedent
            ndxWork = CorefDst(ndxWork)
          End While
          ' Add the length of this chain
          dtrChain.Item("Len") = intLen
          ' Add the length of this chain excluding traces
          dtrChain.Item("NoTraceLen") = intNoTraceLen
          ' Add the chain's root
          ' ======== DEBUG ==========
          ' If (dtrChain.Item("ChainId") = 13) Then Stop
          ' =========================
          ndxWork = GetRootNode(ndxThis)
          If (ndxWork Is Nothing) Then
            dtrChain.Item("RootEtreeId") = -1
          Else
            dtrChain.Item("RootEtreeId") = ndxWork.Attributes("Id").Value
          End If
        Case ""
          ' No need to take action!!
        Case Else
          ' This should not occur, but may happen with other files -- just signal this
          Logging("modCoref/MakeOneChain unknown reference type = " & strRefType & " in: " & _
                  IO.Path.GetFileNameWithoutExtension(GetTableSetting(tdlThis, "File")))
      End Select
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modCoref/MakeOneChain error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   GetIpMatOrSubId
  ' Goal:   Get the @Id number of the IP-MAT or IP-SUB I am part of
  '         Only incorporate IP-SUB, if it is not part of a CP-REL (but of another CP)
  ' History:
  ' 14-09-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetIpMatOrSubId(ByRef ndxThis As XmlNode, ByVal bMatOnly As Boolean) As Integer
    Dim ndxWork As XmlNode  ' Working node
    Dim ndxCp As XmlNode    ' The parent CP of a subordinate clause
    Dim strLabel As String  ' The label of the node we are now considering

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return -1
      If (ndxThis.Name <> "eTree") Then Return -1
      ' Go upwards one by one
      ndxWork = ndxThis.ParentNode
      While (Not ndxWork Is Nothing) AndAlso (ndxWork.Name = "eTree")
        ' Get this one's label
        strLabel = ndxWork.Attributes("Label").Value
        ' Check if this is an IP-MAT
        If (strLabel Like "IP-MAT*") Then
          ' Okay, found my parent
          Return ndxWork.Attributes("Id").Value
        ElseIf (Not bMatOnly) AndAlso (strLabel Like "IP-SUB*") Then
          ' THis is a subordinate clause -- check for the ancestor CP
          ndxCp = ndxWork.SelectSingleNode("./ancestor::eTree[starts-with(@Label, 'CP')]")
          If (ndxCp Is Nothing) Then
            ' We cannot determine what the CP parent of us is--be on the safe side and don't count it in
            ' Stop
          Else
            ' Check the kind of CP parent we have
            If (Not ndxCp.Attributes("Label").Value Like "CP-REL*") Then
              ' This is a subordinate clause, but not a relative clause --> okay
              Return ndxWork.Attributes("Id").Value
            End If
          End If
        End If
        ' Go up one more level
        ndxWork = ndxWork.ParentNode
      End While
      ' Return failure
      Return -1
    Catch ex As Exception
      ' Show error
      HandleErr("modCoref/GetIpMatOrSubId error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
End Module
