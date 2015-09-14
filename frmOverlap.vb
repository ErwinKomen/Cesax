Imports System.Xml
Public Class frmOverlap
  Private bInit As Boolean = False    ' Initialisation flag
  Private bReady As Boolean = False   ' Indicate we are ready
  Private dtrChain() As DataRow       ' Selection of <chain> elements
  Private strState As String = ""     ' The permission status
  Private ndxSrc As XmlNode         ' Source node
  Private ndxDst As XmlNode         ' Destination node
  ' ------------------------------------------------------------------------------------
  ' Name:   Ready
  ' Goal:   Show our readiness
  ' History:
  ' 10-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Public ReadOnly Property Ready() As Boolean
    Get
      ' Return current status
      Return bReady
    End Get
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   frmOverlap_Load
  ' Goal:   Trigger timer
  ' History:
  ' 10-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmOverlap_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Set the parent window
    Me.Owner = frmMain
    ' Trigger initialisation
    Me.Timer1.Enabled = True
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Start Initialisation
  ' History:
  ' 10-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Try
      ' Switch off the timer
      Me.Timer1.Enabled = False
      ' Perform initialisation
      DoInit()
      ' If we are okay, then continue
      If (bInit) Then
        If (CheckOverlap(True)) Then
          ' We are ready
          Status("Ready. You can close this window.")
        Else
          ' There was a problem
          Status("Could not finish checking for overlapping chains")
        End If
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("frmOverlap/Timer1 error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   DoInit
  ' Goal:   Initialise
  ' History:
  ' 10-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub DoInit()
    Try
      ' Validate
      If (tdlRefChain Is Nothing) Then
        ' Signal to the user that we cannot perform
        Status("Close this window and first do Reference/List")
        Exit Sub
      End If
      ' Select ALL the chains in increasing order
      dtrChain = tdlRefChain.Tables("Chain").Select("", "ChainId ASC")
      ' Check result
      If (dtrChain Is Nothing) Then
        Status("Could not produce a set of chains")
        Exit Sub
      End If
      ' Set the init flag
      bInit = True
    Catch ex As Exception
      ' Show error
      HandleErr("frmOverlap/DoInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   CheckOverlap
  ' Goal:   Check for overlap between chains
  ' History:
  ' 09-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function CheckOverlap(ByVal bAuto As Boolean) As Boolean
    Dim dtrItem() As DataRow      ' Selection of <item> elements
    Dim dtrHigh() As DataRow      ' Selection of <item> elements
    Dim ndxHigh As XmlNode        ' First element on the "higher" chain
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intK As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage progress
    Dim intChainId As Integer     ' Chain's id
    Dim intRootEtreeId As Integer ' Current chain's root id
    Dim intTargetId As Integer    ' eTreeId of the target
    Dim intSourceId As Integer    ' eTreeId of the source
    Dim strRefType As String      ' Reference type
    Dim strSrcIp As String
    Dim strDstIp As String

    Try
      ' Walk all the chains
      For intI = 0 To dtrChain.Length - 1
        ' Show progress
        intPtc = (intI + 1) * 100 \ dtrChain.Length
        Status("Chains " & intPtc, intPtc)
        ' Get values from this chain
        intChainId = dtrChain(intI).Item("ChainId")
        intRootEtreeId = dtrChain(intI).Item("RootEtreeId")
        ' Initialise
        intSourceId = -1
        ' Get the first node of this chain
        dtrItem = tdlRefChain.Tables("Item").Select("ChainId=" & intChainId, "eTreeId ASC")
        ' Do we have something?
        If (dtrItem.Length > 0) Then
          ' Double check
          If (Not dtrItem(0).Item("Syntax") Like "*PRN*") Then
            ' Look within all "higher" chains
            For intJ = intI + 1 To dtrChain.Length - 1
              ' We already know that the ChainId value is higher (due to the sorting of [dtrChain])
              ' Check identity of the root id
              If (intRootEtreeId = dtrChain(intJ).Item("RootEtreeId")) Then
                ' Get the first node on this chain
                dtrHigh = tdlRefChain.Tables("Item").Select("ChainId=" & dtrChain(intJ).Item("ChainId"), "ItemId ASC")
                ' Check the @Syntax field
                If (dtrHigh.Length > 0) AndAlso (Not dtrHigh(0).Item("Syntax") Like "*PRN*") Then
                  ' Get the first node of this "higher" chain -- this is the target we suggest
                  intTargetId = dtrHigh(0).Item("eTreeId")
                  ndxHigh = IdToNode(intTargetId)
                  ' Get the first node within [dtrItem] that has a HIGHER eTreeId than [ndxHigh] has
                  For intK = 0 To dtrItem.Length - 1
                    If (dtrItem(intK).Item("eTreeId") > intTargetId) Then
                      intSourceId = dtrItem(intK).Item("eTreeId")
                      Exit For
                    End If
                  Next intK
                  ' Validate
                  If (intSourceId >= 0) Then
                    ' Get source and destination
                    ndxSrc = IdToNode(intSourceId)
                    ndxDst = IdToNode(intTargetId)
                    ' Show suggested link to the user
                    SelectNode(ndxSrc, "source")
                    SelectNode(ndxDst, "target")
                    ' Find out what the reference type should be
                    strSrcIp = GetIpAncestor(ndxSrc).Attributes("Label").Value
                    strDstIp = GetIpAncestor(ndxDst).Attributes("Label").Value
                    ' Determine reference type
                    If ((InStr(strSrcIp, "SPE") > 0) AndAlso (InStr(strDstIp, "SPE") = 0)) OrElse _
                       ((InStr(strSrcIp, "SPE") = 0) AndAlso (InStr(strDstIp, "SPE") > 0)) Then
                      strRefType = strRefCross
                    Else
                      strRefType = strRefIdentity
                    End If
                    ' Give the source and destination information
                    Me.tbSrcNode.Text = NodeInfo(ndxSrc)
                    Me.tbDstNode.Text = NodeInfo(ndxDst)
                    ' Show the two chains in the textboxes
                    Me.tbChainMain.Text = GetRefChain(intChainId)
                    Me.tbChainOverlap.Text = GetRefChain(dtrChain(intJ).Item("ChainId"))
                    ' Make sure back translation is adapted
                    ShowPdeEtree(intSourceId)
                    ' Put selection back to me
                    Me.Focus()
                    ' Wait until we have an acceptable permission state
                    strState = "" : Status("Please choose <Accept> or <Reject>...")
                    While (strState = "")
                      Application.DoEvents()
                    End While
                    ' Action depends on the status
                    Select Case strState
                      Case "Accept"
                        ' Try to make the link
                        If (CorefChange(ndxSrc, ndxDst, strRefType)) Then
                          ' Copy features from antecedent to me
                          If (Not CopyFeaturesNode(ndxSrc, ndxDst, strRefType)) Then
                            ' Well, it is a pity that we cannot copy the features, but not a deadly sin!
                          End If
                          ' Copy the ChainId feature from source to antecedent and further down
                          If (Not CopyChainId(ndxSrc, ndxDst, strRefType)) Then
                            ' It's a pity if this doesn't work -- user will have to recalculate
                          End If
                          ' Add the history
                          If (Not AddHistory(ndxSrc, "coref", "OverlapRepair")) Then Return False
                          ' Show we have changed
                          frmMain.SetDirty(True)
                        Else
                          ' Show error information on error page
                          Status("Could not make this link")
                          ' There is a problem
                          Return False
                        End If
                      Case Else
                        ' Don't accept it
                    End Select
                    ' Switch off the colouring
                    SelectNode(ndxSrc, "blank")
                    SelectNode(ndxDst, "blank")
                  End If
                End If
              End If
            Next intJ
          End If
        End If
      Next intI
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/CheckOverlap error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   cmdReject_Click
  ' Goal:   Indicate we reject this one
  ' History:
  ' 09-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdReject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReject.Click
    Try
      ' Set the state to "Reject"
      strState = "Reject"
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/cmdReject error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   cmdAccept_Click
  ' Goal:   Indicate we accept this one
  ' History:
  ' 09-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdAccept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAccept.Click
    Try
      ' Set the state to "Accept"
      strState = "Accept"
    Catch ex As Exception
      ' Show error
      HandleErr("modMain/cmdAccept error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   cmdExit_Click
  ' Goal:   Quit the form
  ' History:
  ' 09-02-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
    ' Show we are ready now
    bReady = True
    ' Switch off the colouring
    SelectNode(ndxSrc, "blank")
    SelectNode(ndxDst, "blank")
    ' Just quit!
    Me.Close()
  End Sub

End Class