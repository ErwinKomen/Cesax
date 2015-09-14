Imports System.Xml
Module modRefState
  ' ==================================== PRIVATE VARIABLES ==========================================================
  Private loc_strNoun As String = "N|NS|*+N|*+NS|N-*|N^*|N+*"
  Private loc_strPro As String = "PRO|PRO+*|PRO-*|PRO^*|WPRO|*+WPRO|WPRO^*|WPRO+ADV"
  Private loc_strPoss As String = "PRO$*|PRO$^*"
  Private loc_strProp As String = "NR|NR-*|NR^*|NPR*"
  Private loc_strAdj As String = "ADJ|ADJ^*|ADJP|ADJP-*|ADJR|ADJR^*|ADJS|ADJS^*"
  Private loc_strDem As String = "D|D-*|D^*|D+*|DPRO*"
  Private loc_strNum As String = "NUM|NUM-*|NUM^*|NUMP|NUMP-*"
  Private loc_strQuan As String = "Q|Q^*|NEG+Q*|Q+N*|QR|QR^*|QS|QS^*"
  Private loc_strPtc As String = "VAG^*|VBN^*|PTP|PTP-*"
  Private loc_strOth As String = "RP+*|PP|FW|LATIN|ADV|ADV^*|ADV+P|ONE|*+ONE*|OTHER*|SUCH|MAN|MAN^*|QTP"
  Private loc_strConjp As String = "CONJP|CONJP-*"
  Private loc_strPrn As String = "NP*-PRN"
  Private loc_strNP As String = "NP|NP-*"
  Private loc_strNPphrasal As String = "QP|QP-*|WNP|WNP-*"
  Private arChild() As String = {loc_strNoun, loc_strPro, loc_strProp, loc_strAdj, _
                                 loc_strDem, loc_strQuan, loc_strNum, loc_strPtc, _
                                 loc_strOth}
  ' =================================================================================================================
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  GetNPhead
  ' Goal :  Determine what the head node is of an NP
  ' History:
  ' 28-02-2013  ERK Created (derived from tb:NPhead() in CorpusStudio definition file)
  ' -----------------------------------------------------------------------------------------------------------------
  Public Function GetNPhead(ByRef ndxThis As XmlNode, ByRef ndxHead As XmlNode, ByVal bDeep As Boolean) As Boolean
    Dim ndxWork As XmlNode  ' Working node
    Dim ndxPoss As XmlNode  ' One possessive child
    Dim intI As Integer     ' Counter

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      If (ndxThis.Name <> "eTree") Then Return False
      ' Check on different child possibilities in order
      For intI = 0 To arChild.Length - 1
        ' Check on this child
        ndxWork = ndxThis.SelectSingleNode("./child::eTree[tb:Like(string(@Label), '" & _
                    arChild(intI) & "')]", conTb)
        If (ndxWork IsNot Nothing) Then
          ' There is this child, so assume it is the head
          ndxHead = ndxWork
          Return True
        End If
      Next intI
      ' There are no obvious heads, so try different specific possibilities
      ndxWork = ndxThis.SelectSingleNode("./child::eTree[tb:Like(string(@Label), '" & _
                    loc_strNP & "')]", conTb)
      If (ndxWork IsNot Nothing) Then
        ' Could it be a CONJP or PRN?
        If (ndxThis.SelectSingleNode("./child::eTree[tb:Like(string(@Label), '" & _
            loc_strConjp & "|" & loc_strPrn & "')]", conTb) IsNot Nothing) Then
          ' Action depends on deep parsing or shallow
          If (bDeep) Then
            ' Return the head of the NP
            If (GetNPhead(ndxWork, ndxHead, True)) Then Return True
          ElseIf (ndxThis.SelectSingleNode("./child::eTree[tb:Like(string(@Label), '" & _
            loc_strConjp & "|" & loc_strPrn & "') and not(descendant::eLeaf[@Type='Vern'])]", conTb) IsNot Nothing) Then
            ' Return the head of the NP
            If (GetNPhead(ndxWork, ndxHead, True)) Then Return True
          Else
            ' Shallow: we found it!
            ndxHead = ndxWork : Return True
          End If
        End If
      End If
      ' Suppose the whole NP consists of just one possessive?
      ndxPoss = ndxThis.SelectSingleNode("./child::eTree[tb:Like(string(@Label), '" & _
                  loc_strPoss & "')]", conTb)
      If (ndxThis.SelectNodes("./child::eTree").Count = 1) AndAlso (ndxPoss IsNot Nothing) Then
        ' The head is the possessive child
        ndxHead = ndxPoss
        Return (True)
      End If
      ' Suppose I myself am nothing but a possessive pronoun?
      If (DoLike(ndxThis.Attributes("Label").Value, loc_strPoss)) Then
        ' I am my own head
        ndxHead = ndxThis
        Return True
      End If

      ' One last attempt: suppose there is an NP child anyway?
      If (ndxWork IsNot Nothing) Then
        ' Action depends on deep parsing or shallow
        If (bDeep) Then
          ' Return the head of the NP
          If (GetNPhead(ndxWork, ndxHead, True)) Then Return True
        Else
          ' Shallow: we found it!
          ndxHead = ndxWork : Return True
        End If
      End If
      ' Suppose I do not have <eTree> children, but only have an <eLeaf> child?
      ndxWork = ndxThis.SelectSingleNode("child::eLeaf[1]")
      If (ndxWork IsNot Nothing) Then
        ' I myself am the head
        ndxHead = ndxThis
        Return True
      End If
      ' There are no obvious heads, but is there one child or more?
      If (ndxThis.SelectNodes("./child::eTree").Count = 1) Then
        ' There is just one child, so this must be the head
        ndxHead = ndxThis.SelectSingleNode("./child::eTree")
        Return True
      ElseIf (ndxThis.SelectNodes("./child::eTree").Count > 0) Then
        ' Check for some phrasal children
        ndxWork = ndxThis.SelectSingleNode("./child::eTree[tb:matches(@Label,'" & loc_strNPphrasal & "')]", conTb)
        ' There are some phrases we need to check
        If (ndxWork IsNot Nothing) Then
          ' Accept the last child as the head
          ndxHead = ndxWork.SelectSingleNode("./child::eTree[last()]")
          If (ndxHead Is Nothing) Then
            Stop
          End If
          Return True
        ElseIf (ndxThis.SelectSingleNode("./child::eTree[tb:matches(@Label,'CP-REL*')]", conTb) IsNot Nothing) Then
          ' Is there more than one child or not?
          If (ndxThis.SelectNodes("./child::eTree").Count = 1) Then
            ' There is just one CP-REL child...
            ndxWork = ndxThis.SelectSingleNode("./child::eTree[tb:matches(@Label,'CP-REL*')]", conTb)
            ' Return the first descendant <eTree.
            ndxHead = ndxWork.SelectSingleNode("./descendant::eTree")
            If (ndxHead Is Nothing) Then
              Stop
            End If
            Return True
          Else
            ' Take the <eTree> preceding the CP-REL
            ndxHead = ndxThis.SelectSingleNode("./child::eTree[following-sibling::eTree[1][tb:matches(@Label,'CP-REL*')]]", conTb)
            If (ndxHead Is Nothing) Then
              Stop
            End If
            Return True
          End If
        Else
          ' Get the last child
          ndxWork = ndxThis.SelectSingleNode("./child::eTree[last()]")
          ' Action depends on the phrase kind of the only child
          Select Case CatPhrase(ndxWork.Attributes("Label").Value)
            Case "CP", "IP"   ' CP-FRL, IP-PPL etc
              ' Return the first NP child descendant
              ndxHead = ndxWork.SelectSingleNode("./descendant::eTree")
              If (ndxHead Is Nothing) Then
                Stop
              End If
              Return True
            Case "N", "NS"
              ' Return myself
              ndxHead = ndxWork : Return True
              'Case "QP"
              '  ' Return the first Q-kind child
              '  ndxHead = ndxWork.SelectSingleNode("./descendant::eTree[tb:matches(@Label, 'Q*')]", conTb)
              '  If (ndxHead Is Nothing) Then
              '    Stop
              '  End If
              '  Return True
            Case "NX"
              ' Return the head of the NP
              If (GetNPhead(ndxWork, ndxHead, True)) Then Return True
            Case Else
              Debug.Print("modRefState/GetNPhead: don' know what to do with [" & ndxWork.Attributes("Label").Value & "] (children=" & _
                          ndxThis.SelectNodes("./child::eTree").Count & ")")
              Stop
          End Select
        End If
      ElseIf (ndxThis.SelectNodes("./child::eLeaf").Count > 0) Then
        ' Just return myself as head
        ndxHead = ndxThis
        Return True
      Else
        Stop
      End If
      ' Return failure
      Return False
    Catch ex As Exception
      ' Show the error
      HandleErr("modRefState/GetNPhead error: " & ex.Message & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  CleanLabel
  ' Goal :  Get the label of this node, and clean it up
  ' History:
  ' 28-02-2013  ERK Created 
  ' -----------------------------------------------------------------------------------------------------------------
  Public Function CleanLabel(ByRef ndxThis As XmlNode) As String
    Dim strLabel As String = "" ' The label
    ' Dim intPos As Integer       ' Position in string

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      If (ndxThis.Name <> "eTree") Then Return False
      ' Get the label
      strLabel = ndxThis.Attributes("Label").Value
      ' Return the phrasal category
      Return CatPhrase(strLabel)
    Catch ex As Exception
      ' Show the error
      HandleErr("modRefState/CleanLabel error: " & ex.Message & vbCrLf & ex.StackTrace)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CatPhrase
  ' Goal:   Get the category of the label of <eTree> node [ndxThis]:
  '         phrase    - first part of the label
  ' History:
  ' 15-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function CatPhrase(ByVal strLabel As String) As String
    Dim intPos As Integer       ' Position in string

    Try
      ' Validate
      If (strLabel = "") Then Return "-"
      ' Take part before -, ^ or $
      intPos = InStr(strLabel, "-")
      If (intPos = 0) Then intPos = InStr(strLabel, "^")
      ' Strip off
      If (intPos > 0) Then strLabel = Left(strLabel, intPos - 1)
      ' What about +
      intPos = InStr(strLabel, "+")
      If (intPos > 0) Then strLabel = Mid(strLabel, intPos + 1)
      ' And now the $ sign
      intPos = InStr(strLabel, "$")
      If (intPos > 0) Then strLabel = Left(strLabel, intPos - 1)
      ' As well as the = sign
      intPos = InStr(strLabel, "=")
      If (intPos > 0) Then strLabel = Left(strLabel, intPos - 1)
      ' =========== DEBUG ===========
      If (strLabel = "") Then
        ' Stop
        strLabel = "-"
      End If

      ' =============================
      ' Return the label
      Return strLabel
    Catch ex As Exception
      ' Show the error
      HandleErr("modRefState/CatPhrase error: " & ex.Message & vbCrLf & ex.StackTrace)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   CatFunction
  ' Goal:   Get the category of the label of <eTree> node [ndxThis]:
  '         function  - last part of the label
  ' History:
  ' 15-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function CatFunction(ByVal strLabel As String) As String
    Dim intPos As Integer       ' Position in string

    Try
      ' Validate
      If (strLabel = "") Then Return "-"
      '' Eat away any numerals ==> do not do this, since it works counter-productive!
      'While (strLabel.Length > 0) AndAlso (DoLike(Right(strLabel, 2), "-[1-9]"))
      '  strLabel = Left(strLabel, strLabel.Length - 2)
      'End While
      ' Eat away before the $ sign, the ^ sign and the = sign
      intPos = InStr(strLabel, "$")
      If (intPos > 0) Then strLabel = Left(strLabel, intPos - 1)
      intPos = InStr(strLabel, "=")
      If (intPos > 0) Then strLabel = Left(strLabel, intPos - 1)
      ' Find function if there is a ^ or - sign from the right
      intPos = InStrRev(strLabel, "^")
      If (intPos > 0) Then strLabel = Mid(strLabel, intPos + 1) : Return strLabel
      intPos = InStrRev(strLabel, "-")
      If (intPos > 0) Then
        strLabel = Mid(strLabel, intPos + 1)
      Else
        strLabel = "-"
      End If
      ' =========== DEBUG ===========
      If (strLabel = "") Then Stop
      ' =============================
      ' Return the functional category
      Return strLabel
    Catch ex As Exception
      ' Show the error
      HandleErr("modRefState/CatFunction error: " & ex.Message & vbCrLf & ex.StackTrace)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   LabelCategory
  ' Goal:   Get the category of the label of <eTree> node [ndxThis]:
  '         function  - last part of the label
  '         phrase    - first part of the label
  ' History:
  ' 15-04-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function LabelCategory(ByRef ndxThis As XmlNode, ByVal strType As String) As String
    Dim strLabel As String = "" ' This node's label

    Try
      ' validate
      If (ndxThis Is Nothing) Then Return "-"
      If (ndxThis.Name <> "eTree") Then Return "-"
      ' Get label
      strLabel = ndxThis.Attributes("Label").Value
      ' Check the type
      Select Case strType
        Case "phrase", "Phrase", "phr"
          Return CatPhrase(strLabel)
        Case "function", "fun", "Function"
          Return CatFunction(strLabel)
      End Select
      ' Found nothing, so return empty
      Return "-"
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/LabelCategory error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return "-"
    End Try
  End Function
End Module
