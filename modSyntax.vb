Imports System.Xml
Imports System.Text.RegularExpressions
Module modSyntax
  ' ============================== PUBLIC VARIABLES ============================================================
  Public colConst As New StringColl ' Collection of constituents with their frequency of occurrence
  Public tdlSyntax As DataSet = Nothing ' Where we contain the learned information
  Public strSyntaxFile As String = ""   ' Where to store the syntax information
  Public Const SYNT_FILE As String = "ParseInfo.xml"
  ' ============================== LOCAL  VARIABLES ============================================================
  Private pdxFile As XmlDocument    ' The XML document we are currently working on...
  Private Const SYNT_XSD As String = "ParseInfo.xsd"
  Private Const MAX_ANC As Integer = 4  ' maximum number of ancestors reported in <Ptree>
  Private Const MAX_CHI As Integer = 4  ' maximum number of children reported in <Ptree>
  ' ============================================================================================================
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   OneLearnSyntaxFile
  ' Goal:   Learn the constituents from one Psdx file
  ' History:
  ' 31-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function OneLearnSyntaxFile(ByVal strFile As String) As Boolean
    Dim ndxFor As XmlNode = Nothing   ' Forest node
    Dim ndxThis As XmlNode = Nothing  ' Current node
    Dim ndxList As XmlNodeList        ' List of <eLeaf> nodes
    Dim strLabel As String = ""       ' Output label of a constituent
    Dim strChildren As String = ""    ' Children of a constituent in semi-stack
    Dim strPartial As String = ""     ' Partial tree seen from me as center
    Dim intI As Integer               ' Counter
    Dim intPtc As Integer             ' Percentage
    Dim intCount As Integer           ' Number of child nodes
    Dim intNum As Integer             ' Number of child nodes retrieved

    Try
      ' Validate
      If (Not IO.File.Exists(strFile)) Then Return False
      ' Read the file as an XML document
      Status("Reading [" & IO.Path.GetFileNameWithoutExtension(strFile) & "]...")
      ' Try read this file into an XML structure
      If (Not ReadXmlDoc(strFile, pdxFile)) Then
        ' There was an error
        Status("Unable to read file: " & strFile)
        Return False
      End If
      ' Walk all the forests in the file
      ' Set the current file
      pdxCurrentFile = pdxFile
      ' Look for the first <forest> element
      If (Not GetFirstForest(pdxFile, ndxFor)) Then Logging("Cannot find first forest") : Return False
      ' Determine how many elements there are potentially
      If (ndxFor Is Nothing) Then
        intCount = 0
      ElseIf (ndxFor.ParentNode Is Nothing) Then
        intCount = 0
      Else
        intCount = ndxFor.ParentNode.ChildNodes.Count
      End If
      ' Walk through all the <forest> elements FORWARD
      While (Not ndxFor Is Nothing)
        ' Check if we are not being interrupted
        If (bInterrupt) Then Return False
        ' Show where we are (how do we KNOW where we are?)
        intPtc = (ndxFor.Attributes("forestId").Value) * 100 \ intCount
        Status("Learning [" & IO.Path.GetFileNameWithoutExtension(strFile) & "] " & intPtc & "%", intPtc)
        ' Make sure this is a forest and not something else
        If (ndxFor.Name = "forest") Then
          ' Visit all the <eTree> constituents of this forest
          ndxList = ndxFor.SelectNodes("./descendant::eTree")
          For intI = 0 To ndxList.Count - 1
            ' Note the label of this constituent
            strLabel = CleanLabel(ndxList(intI).Attributes("Label").Value)
            ' Get all the children's labels into a semi-stack
            strChildren = GetLabelList(ndxList(intI), "./child::eTree", intNum)
            ' Add the appropriate partial tree
            If (Not AddPartialTree(strLabel, ndxList(intI))) Then Logging("Problem creating partial tree") : Return False
            ' Get the other lists
            strPartial = GetPartialTree(ndxList(intI))
            ' Add the combination of @Label/@Children/@Partial into the collection
            If (Not AddSyntax(strLabel, strChildren, intNum, NodeInfo(ndxList(intI)), strPartial)) Then Return False
            'End If
          Next intI
        End If
        ' Go to the next forst
        ndxFor = ndxFor.NextSibling
      End While
      ' Return success 
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/OneLearnSyntaxFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   FreqSum
  ' Goal:   Get the total number of @Freq attribute values added
  '         Optional [intMin] is the minimum frequency to be added
  ' History:
  ' 20-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function FreqSum(ByRef dtrThis() As DataRow, Optional ByVal intMin As Integer = 0) As Integer
    Dim intI As Integer       ' COunter
    Dim intFreq As Integer    ' One frequency
    Dim intSum As Integer = 0 ' Total

    Try
      ' validate
      If (dtrThis Is Nothing) OrElse (dtrThis.Length = 0) Then Return 0
      ' Count
      For intI = 0 To dtrThis.Length - 1
        ' Get the frequency
        intFreq = dtrThis(intI).Item("Freq")
        ' Is it higher than the minimum?
        If (intFreq > intMin) Then
          ' Add it to the sum-total
          intSum += intFreq
        End If
      Next intI
      ' Return the total
      Return intSum
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/FreqSum error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return 0
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DerivePartialRules
  ' Goal:   Derive the possible rules based on <Part> and <Ptree> items
  ' History:
  ' 20-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function DerivePartialRules() As Boolean
    Dim colRep As New StringColl  ' HTML report table
    Dim colRule As New StringColl ' Collection of rules
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim dtrPart As DataRow        ' Current row
    Dim strPath As String         ' Current path
    Dim strLine As String = ""    ' One line for the collection
    Dim strRule As String = ""    ' One rule line
    ' Dim strText As String         ' Combination
    Dim intPartId As Integer      ' ID of current part
    Dim intCount As Integer       ' Current count
    Dim intThresh As Integer = 4  ' Minimum number of items that should be there
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intM As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage

    Try
      ' Validate
      If (tdlSyntax Is Nothing) Then Return False
      If (tdlSyntax.Tables("Part").Rows.Count = 0) OrElse (tdlSyntax.Tables("Ptree").Rows.Count = 0) Then Return False
      ' Gather results
      dtrFound = tdlSyntax.Tables("Part").Select("Count >= " & intThresh, "Count DESC")
      For intI = 0 To dtrFound.Length - 1
        ' Get it
        dtrPart = dtrFound(intI)
        ' Where are we?
        intPtc = (intI + 1) * 100 \ dtrFound.Length
        Status("Partial Rules " & intPtc & "%", intPtc)
        With dtrPart
          ' Get Id
          intPartId = .Item("PartId") : intCount = .Item("Count")
          ' Find all the ancestors
          For intJ = 1 To MAX_ANC
            ' Set current ancestor path
            strPath = "a[" & intJ & "]"
            ' Add to report
            If (Not DerivePartialRulesTableRow(dtrPart, intPartId, strPath, intCount)) Then Return False
            ' Check preceding siblings of this ancestor
            For intM = 1 To MAX_CHI
              ' Add preceding sibling to report
              If (Not DerivePartialRulesTableRow(dtrPart, intPartId, strPath & "c[-" & intM & "]", intCount)) Then Return False
            Next intM
            ' Check following siblings of this ancestor
            For intM = 1 To MAX_CHI
              ' Add following sibling to report
              If (Not DerivePartialRulesTableRow(dtrPart, intPartId, strPath & "c[+" & intM & "]", intCount)) Then Return False
            Next intM
            ' Add preceding last sibling to report
            If (Not DerivePartialRulesTableRow(dtrPart, intPartId, strPath & "c[-L]", intCount)) Then Return False
            ' Add following last sibling to report
            If (Not DerivePartialRulesTableRow(dtrPart, intPartId, strPath & "c[+R]", intCount)) Then Return False
          Next intJ
        End With
      Next intI
      ' Return positively
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/DerivePartialRules error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DerivePartialRulesTableRow
  ' Goal:   Try to derive one rule and place it under <Part> with PartId=intPartId
  ' History:
  ' 20-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function DerivePartialRulesTableRow(ByRef dtrParent As DataRow, ByVal intPartId As Integer, _
                                              ByVal strRulePath As String, ByVal intCount As Integer) As Boolean
    Dim dtrWork() As DataRow        ' Result of SELECT
    Dim dtrRule As DataRow          ' One new rule datarow
    Dim strRuleLabel As String = "" ' Label for one rule
    Dim intMinFreq As Integer = 1   ' Minimum frequency to count for anything
    Dim intMinPerc As Integer = 10  ' Minimal percentage coverage
    Dim intFreqSum As Integer       ' Total frequency
    Dim intRuleFreq As Integer      ' Frequency for this rule
    Dim intRuleLen As String        ' Length of the initial rule label

    Try
      ' Select relevant part of <Ptree>
      dtrWork = tdlSyntax.Tables("Ptree").Select("PartId= " & intPartId & " AND Path='" & strRulePath & "'", "Freq DESC")
      ' Does this one count?
      intFreqSum = FreqSum(dtrWork)
      If (intFreqSum >= intCount) Then
        ' Get initial rule
        strRuleLabel = dtrWork(0).Item("Label") : intRuleFreq = 0 : intRuleLen = strRuleLabel.Length
        ' This one counts, so display it...
        For intK = 0 To dtrWork.Length - 1
          With dtrWork(intK)
            ' Only note those with frequence higher than 1
            If ((100 * .Item("Freq") \ intFreqSum) > intMinPerc) Then
              ' Adapt the rule
              strRuleLabel = PartialLabels(strRuleLabel, .Item("Label"))
              ' Add a row for this line
              intRuleFreq += .Item("Freq")
              'strLine &= "<tr><td>" & .Item("Path") & "</td><td>" & .Item("Label") & "</td><td>" & .Item("Freq") & _
              '  "</td></tr>" & vbCrLf
            End If
          End With
        Next intK
        ' Add a line for the rule
        If (intRuleFreq > 0) Then
          ' Possibly adapt the rule label
          If (strRuleLabel <> "") AndAlso (intRuleLen > strRuleLabel.Length) Then
            ' Add a question mark to signal that it needs adaptation
            strRuleLabel &= "?"
          End If
          ' Create a new datarow
          dtrRule = AddOneDataRowWithParent(tdlSyntax, "Rule", "RuleId", dtrParent)
          With dtrRule
            .Item("PartId") = intPartId
            .Item("Path") = strRulePath
            .Item("Label") = strRuleLabel
            .Item("Freq") = intRuleFreq
          End With
        End If
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/DerivePartialRulesTableRow error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   PartialReport
  ' Goal:   Give a report of the Trigger elements with completely covering partial trees
  ' History:
  ' 20-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function PartialReport() As String
    Dim colRep As New StringColl  ' HTML report table
    'Dim colRule As New StringColl ' Collection of rules
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim strPath As String         ' Current path
    Dim strLine As String = ""    ' One line for the collection
    Dim strRule As String = ""    ' One rule line
    Dim strText As String         ' Combination
    Dim intPartId As Integer      ' ID of current part
    Dim intCount As Integer       ' Current count
    Dim intThresh As Integer = 4  ' Minimum number of items that should be there
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intM As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage

    Try
      ' Validate
      If (tdlSyntax Is Nothing) Then Return ""
      If (tdlSyntax.Tables("Part").Rows.Count = 0) OrElse (tdlSyntax.Tables("Ptree").Rows.Count = 0) Then Return ""
      ' Gather results
      dtrFound = tdlSyntax.Tables("Part").Select("Count >= " & intThresh, "Count DESC")
      colRep.Add("<h1>Partial report</h1><p>")
      For intI = 0 To dtrFound.Length - 1
        ' Where are we?
        intPtc = (intI + 1) * 100 \ dtrFound.Length
        Status("Partial Report " & intPtc & "%", intPtc)
        With dtrFound(intI)
          ' Get Id
          intPartId = .Item("PartId") : intCount = .Item("Count")
          ' Start a table
          colRep.Add("<h4>Trigger = " & .Item("Trigger") & "</h4><p>Count=" & intCount & _
                     "<table><tr><td>Path</td><td>Label</td><td>Freq</td></tr>")
          ' Find all the ancestors
          For intJ = 1 To MAX_ANC
            ' Set current ancestor path
            strPath = "a[" & intJ & "]"
            ' Add to report
            If (Not PartialReportTableRow(intPartId, strPath, intCount, strLine, strRule)) Then Return False
            colRep.Add(strLine)
            ' If (strRule <> "") Then colRule.Add("<br>" & .Item("Trigger") & ": " & strRule)
            ' Check preceding siblings of this ancestor
            For intM = 1 To MAX_CHI
              ' Add preceding sibling to report
              If (Not PartialReportTableRow(intPartId, strPath & "c[-" & intM & "]", intCount, strLine, strRule)) Then Return False
              colRep.Add(strLine)
              'If (strRule <> "") Then colRule.Add("<br>" & .Item("Trigger") & ": " & strRule)
            Next intM
            ' Check following siblings of this ancestor
            For intM = 1 To MAX_CHI
              ' Add following sibling to report
              If (Not PartialReportTableRow(intPartId, strPath & "c[+" & intM & "]", intCount, strLine, strRule)) Then Return False
              colRep.Add(strLine)
              'If (strRule <> "") Then colRule.Add("<br>" & .Item("Trigger") & ": " & strRule)
            Next intM
            ' Add preceding last sibling to report
            If (Not PartialReportTableRow(intPartId, strPath & "c[-L]", intCount, strLine, strRule)) Then Return False
            colRep.Add(strLine)
            'If (strRule <> "") Then colRule.Add("<br>" & .Item("Trigger") & ": " & strRule)
            ' Add following last sibling to report
            If (Not PartialReportTableRow(intPartId, strPath & "c[+R]", intCount, strLine, strRule)) Then Return False
            colRep.Add(strLine)
            'If (strRule <> "") Then colRule.Add("<br>" & .Item("Trigger") & ": " & strRule)
          Next intJ
          ' Finish table
          colRep.Add("</table>")
        End With
      Next intI
      ' Return the result
      ' strText = colRep.Text & "<h1>Rules</h1>" & colRule.Text
      strText = colRep.Text
      Return strText
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/PartialReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   PartialRulesReport
  ' Goal:   Give a report of the rules for partial trees
  ' History:
  ' 21-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function PartialRulesReport() As String
    Dim colRep As New StringColl  ' HTML report table
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim dtrRule() As DataRow      ' The rules belonging to the current <Part>
    Dim strLine As String = ""    ' One line for the collection
    Dim strRule As String = ""    ' One rule line
    Dim strText As String         ' Combination
    Dim intPartId As Integer      ' ID of current part
    Dim intCount As Integer       ' Current count
    Dim intThresh As Integer = 4  ' Minimum number of items that should be there
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter
    Dim intPtc As Integer         ' Percentage

    Try
      ' Validate
      If (tdlSyntax Is Nothing) Then Return ""
      If (tdlSyntax.Tables("Part").Rows.Count = 0) OrElse (tdlSyntax.Tables("Ptree").Rows.Count = 0) Then Return ""
      ' Gather results
      dtrFound = tdlSyntax.Tables("Part").Select("Count >= " & intThresh, "Trigger ASC, Count DESC")
      colRep.Add("<h1>Partial rules report</h1><p>")
      For intI = 0 To dtrFound.Length - 1
        ' Where are we?
        intPtc = (intI + 1) * 100 \ dtrFound.Length
        Status("Partial Rule Report " & intPtc & "%", intPtc)
        With dtrFound(intI)
          ' Get Id
          intPartId = .Item("PartId") : intCount = .Item("Count")
          ' Find rules belong to this part
          dtrRule = tdlSyntax.Tables("Rule").Select("PartId=" & intPartId, "Path ASC")
          If (dtrRule.Length > 0) Then
            ' Start a table
            colRep.Add("<h4>Trigger = " & .Item("Trigger") & "</h4><p>Count=" & intCount & _
                       "<table><tr><td>Type</td><td>Path</td><td>Label</td><td>Freq</td><td>Perc</td></tr>")
            ' Add the individual rules
            For intJ = 0 To dtrRule.Length - 1
              With dtrRule(intJ)
                If (.Item("Label") = "") Then
                  colRep.Add("<tr><td>Add</td><td>" & .Item("Path") & "</td><td>(any)</td>")
                Else
                  colRep.Add("<tr><td>Connect</td><td>" & .Item("Path") & "</td><td>" & .Item("Label") & "</td>")
                End If
                ' Common part
                colRep.Add("<td>" & .Item("Freq") & "</td><td>" & 100 * .Item("Freq") \ intCount & "</td></tr>")
              End With
            Next intJ
            ' Finish table
            colRep.Add("</table>")
          End If
        End With
      Next intI
      ' Return the result
      strText = colRep.Text
      Return strText
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/PartialRulesReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   PartialReportTableRow
  ' Goal:   Add one row to the HTML output of the partial report
  ' History:
  ' 20-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function PartialReportTableRow(ByVal intPartId As Integer, ByVal strPath As String, ByVal intCount As Integer, _
                                        ByRef strLine As String, ByRef strRule As String) As Boolean
    Dim dtrWork() As DataRow        ' Result of SELECT
    Dim intMinFreq As Integer = 1   ' Minimum frequency to count for anything
    Dim intMinPerc As Integer = 10  ' Minimal percentage coverage
    Dim intFreqSum As Integer       ' Total frequency

    Try
      ' Select relevant part of <Ptree>
      dtrWork = tdlSyntax.Tables("Ptree").Select("PartId= " & intPartId & " AND Path='" & strPath & "'", "Freq DESC")
      strLine = "" : strRule = ""
      ' Does this one count?
      intFreqSum = FreqSum(dtrWork)
      If (intFreqSum >= intCount) Then
        ' Get initial rule
        strRule = dtrWork(0).Item("Label")
        ' This one counts, so display it...
        For intK = 0 To dtrWork.Length - 1
          With dtrWork(intK)
            ' Only note those with frequence higher than 1
            If ((100 * .Item("Freq") \ intFreqSum) > intMinPerc) Then
              ' Adapt the rule
              strRule = PartialLabels(strRule, .Item("Label"))
              ' Add a row for this line
              strLine &= "<tr><td>" & .Item("Path") & "</td><td>" & .Item("Label") & "</td><td>" & .Item("Freq") & _
                "</td></tr>" & vbCrLf
            End If
          End With
        Next intK
        ' Add a line for the rule
        If (strLine <> "") Then
          If (strRule = "") Then
            strLine &= "<tr><td>==RULE==</td><td>(<font color='red'>none</font>)</td><td>--</td></tr>" & vbCrLf
            strRule = "ConnectTo(" & strPath & ")"
          Else
            strLine &= "<tr><td>==RULE==</td><td><font color='blue'>" & strRule & "</font></td><td>--</td></tr>" & vbCrLf
            strRule = "Add(" & strPath & ", " & strRule & ")"
          End If
        End If
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/PartialReportTableRow error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   AddPartialTree
  ' Goal:   Add a partial tree relative to [ndxThis]
  '         The tree consists of one <Part> item containing a list of <Ptree> ones:
  '         @Label  ::= <label>
  '         @Path   ::= <ancestor> [<child>]
  '                     <ancestor> ::= "a" DistanceToAncestor()
  '                     <child>    ::= "c" DistanceToChild()
  ' History:
  ' 20-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function AddPartialTree(ByVal strTrigger As String, ByRef ndxThis As XmlNode) As Boolean
    Dim ndxParent As XmlNode    ' Ancestor
    Dim ndxChild As XmlNodeList ' List of child nodes
    Dim ndxPrev As XmlNode      ' Previous ancestor
    Dim dtrFound() As DataRow   ' Result of SELECT
    Dim dtrThis As DataRow      ' One datarow element
    Dim strPath As String       ' The path to the label
    Dim strLabel As String      ' The label of the node on this path
    Dim intPartId As Integer    ' The ID of our <Part> element
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter

    Try
      ' Validate
      If (strTrigger = "") OrElse (ndxThis Is Nothing) Then Return False
      ' Exclude some trigger(s)
      If (strTrigger = "-") Then Return True
      ' Find an entry for the trigger
      dtrFound = tdlSyntax.Tables("Part").Select("Trigger='" & strTrigger.Replace("'", "''") & "'")
      If (dtrFound.Length = 0) Then
        ' Create a new one
        dtrThis = AddOneDataRow(tdlSyntax, "Part", "PartId", "PartList")
        With dtrThis
          .Item("Count") = 1 : .Item("Trigger") = strTrigger
        End With
      Else
        ' Add information to the existing one, and keep track of the total number
        dtrThis = dtrFound(0) : dtrThis.Item("Count") += 1
      End If
      ' Get the ID
      intPartId = dtrThis.Item("PartId")
      ' Walk all the ancestors, starting with the parent
      ndxParent = ndxThis.ParentNode : intI = 0
      ndxPrev = ndxThis
      ' Walk the parents
      While (ndxParent IsNot Nothing) AndAlso (ndxParent.Name = "eTree") AndAlso (intI < MAX_ANC)
        ' Get path and label
        strPath = "a[" & intI + 1 & "]" : strLabel = CleanLabel(ndxParent.Attributes("Label").Value)
        ' Add the ancestor node information as a <Ptree> element
        If (Not AddPtreeItem(strPath, strLabel, intPartId, dtrThis)) Then Return False
        ' Walk the children left of me
        ndxChild = ndxPrev.SelectNodes("./preceding-sibling::eTree")
        For intJ = 0 To ndxChild.Count - 1
          ' Check where we are
          If (intJ + 1 < MAX_CHI) OrElse (intJ + 1 = ndxChild.Count) Then
            ' Get label
            strLabel = CleanLabel(ndxChild(intJ).Attributes("Label").Value)
            ' Determine the path 
            strPath = "a[" & intI + 1 & "]c[-" & IIf(intJ > 0 And intJ + 1 = ndxChild.Count, "L]", intJ + 1 & "]")
            ' Add the ancestor node information as a <Ptree> element
            If (Not AddPtreeItem(strPath, strLabel, intPartId, dtrThis)) Then Return False
          End If
        Next intJ
        ' Walk the children right of me
        ndxChild = ndxPrev.SelectNodes("./following-sibling::eTree")
        For intJ = 0 To ndxChild.Count - 1
          ' Check where we are
          If (intJ + 1 < MAX_CHI) OrElse (intJ + 1 = ndxChild.Count) Then
            ' Get label
            strLabel = CleanLabel(ndxChild(intJ).Attributes("Label").Value)
            ' Determine the path 
            strPath = "a[" & intI + 1 & "]c[+" & IIf(intJ > 0 And intJ + 1 = ndxChild.Count, "R]", intJ + 1 & "]")
            ' Add the ancestor node information as a <Ptree> element
            If (Not AddPtreeItem(strPath, strLabel, intPartId, dtrThis)) Then Return False
          End If
        Next intJ
        ' Adjust previous
        ndxPrev = ndxParent
        ' Go to next parent
        ndxParent = ndxParent.ParentNode : intI += 1
      End While
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/AddPartialTree error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   AddPtreeItem
  ' Goal:   Add a <Ptree> item, or adjust an existing one's frequency
  ' History:
  ' 20-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function AddPtreeItem(ByVal strPath As String, ByVal strLabel As String, ByVal intPartId As Integer, _
                                ByRef dtrParent As DataRow) As Boolean
    Dim dtrPtree As DataRow = Nothing ' One datarow element
    Dim dtrFound() As DataRow         ' Result of SELECT

    Try
      dtrFound = tdlSyntax.Tables("Ptree").Select("PartId=" & intPartId & " AND Path='" & strPath & _
                                                  "' AND Label='" & strLabel & "'")
      If (dtrFound.Length = 0) Then
        dtrPtree = AddOneDataRowWithParent(tdlSyntax, "Ptree", "PtreeId", dtrParent)
        With dtrPtree
          .Item("Freq") = 1
          .Item("Path") = strPath
          .Item("Label") = strLabel
          .Item("PartId") = intPartId
          .SetParentRow(dtrParent)
        End With
      Else
        dtrPtree = dtrFound(0) : dtrPtree.Item("Freq") += 1
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/AddPtreeItem error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetPartialTree
  ' Goal:   Get a partial tree relative to [ndxThis]
  '         The tree consists of an array of lines, where each line has:
  '         <line>     ::= <location> ":" <label>
  '         <location> ::= <ancestor> [<child>]
  '         <ancestor> ::= "a" DistanceToAncestor()
  '         <child>    ::= "c" DistanceToChild()
  ' History:
  ' 15-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function GetPartialTree(ByRef ndxThis As XmlNode) As String
    Dim ndxList As XmlNodeList  ' List of nodes
    Dim ndxParent As XmlNode    ' Ancestor
    Dim ndxChild As XmlNodeList ' List of child nodes
    Dim ndxPrev As XmlNode      ' Previous ancestor
    Dim strBack As String = ""  ' The list we build
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return ""
      ' Get a list of ancestors and walk them
      ndxList = ndxThis.SelectNodes("./ancestor::eTree")
      ndxParent = ndxThis.ParentNode : intI = 0
      ndxPrev = ndxThis
      ' Walk the parents
      While (ndxParent IsNot Nothing) AndAlso (ndxParent.Name = "eTree")
        ' Strangely enough, this has to be done in reversed order...
        'For intI = ndxList.Count - 1 To 0 Step -1
        ' Add the ancestor node information
        ' AddSemiStack(strBack, "a[" & intI + 1 & "]:" & ndxList(intI).Attributes("Label").Value)
        AddSemiStack(strBack, "a[" & intI + 1 & "]:" & ndxParent.Attributes("Label").Value)
        ' Walk the children left of me
        ndxChild = ndxPrev.SelectNodes("./preceding-sibling::eTree")
        For intJ = 0 To ndxChild.Count - 1
          ' Add this child information
          AddSemiStack(strBack, "a[" & intI + 1 & "]c[-" & intJ + 1 & "]:" & ndxChild(intJ).Attributes("Label").Value)
        Next intJ
        ' Walk the children right of me
        ndxChild = ndxPrev.SelectNodes("./following-sibling::eTree")
        For intJ = 0 To ndxChild.Count - 1
          ' Add this child information
          AddSemiStack(strBack, "a[" & intI + 1 & "]c[" & intJ + 1 & "]:" & ndxChild(intJ).Attributes("Label").Value)
        Next intJ
        ' Adjust previous
        ' ndxPrev = ndxList(intI)
        ndxPrev = ndxParent
        'Next intI
        ' Go to next parent
        ndxParent = ndxParent.ParentNode : intI += 1
      End While
      ' Return what we gathered
      Return strBack
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/GetPartialTree error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetPartialMatch
  ' Goal:   Get the common partial tree match for constituent label [strLabel]
  ' History:
  ' 15-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function GetPartialMatch(ByVal strLabel As String, ByRef strPartial As String) As Boolean
    Dim dtrFound() As DataRow     ' Result of selection
    Dim colBase As New StringColl ' Basic list
    Dim colAdd As New StringColl  ' List to be added
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Counter for colBase
    Dim intK As Integer           ' Counter for colAdd

    Try
      ' Validate
      If (strLabel = "") Then Return False
      ' Get all the datarows
      dtrFound = tdlSyntax.Tables("Const").Select("Label = '" & strLabel & "'")
      ' Are there any matches?
      If (dtrFound.Length = 0) Then Return False
      ' Get the first list
      strPartial = dtrFound(0).Item("Partial")
      ' Is there just one possibility?
      If (dtrFound.Length = 1) Then Return True
      ' Otherwise: make a basic list and put that in a string collection
      If (Not PartialToCollection(strPartial, colBase)) Then Return False
      Debug.Print("Base = " & strPartial)
      ' Walk through all possibilities in the database
      For intI = 1 To dtrFound.Length - 1
        Debug.Print("Comparison #" & intI & "=" & vbCrLf & dtrFound(intI).Item("Partial").ToString)
        ' Retrieve this possibility
        If (Not PartialToCollection(dtrFound(intI).Item("Partial").ToString, colAdd)) Then Return False
        ' Compare this with the basic list
        For intJ = colBase.Count - 1 To 0 Step -1
          ' Find the same path in colAdd
          intK = colAdd.Find(colBase.Item(intJ))
          ' Do we have a path match?
          If (intK < 0) Then
            Debug.Print("Removing " & colBase.Item(intJ))
            ' There is no path match --> remove this item from the basic list
            colBase.DelItem(intJ)
          Else
            ' Both lists have this path! Check how much in the labels overlaps
            strLabel = PartialLabels(colBase.Exmp(intJ), colAdd.Exmp(intK))
            If (strLabel = "") Then
              Debug.Print("Removing " & colBase.Item(intJ))
              ' Remove him anyway
              colBase.DelItem(intJ)
            Else
              Debug.Print("Adjust label from " & colBase.Exmp(intJ) & " to " & strLabel)
              ' Adjust the partial label of the base
              colBase.Exmp(intJ) = strLabel
            End If
          End If
        Next intJ
      Next intI
      ' Reconstruct the result
      strPartial = ""
      For intI = 0 To colBase.Count - 1
        AddSemiStack(strPartial, colBase.Item(intI) & ":" & colBase.Exmp(intI))
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/GetPartialMatch error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   PartialLabels
  ' Goal:   Match two strings and return the overlap
  ' History:
  ' 15-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function PartialLabels(ByVal strLabel As String, ByVal strMatch As String) As String
    Dim intPos As Integer   ' Position in [strLabel]

    Try
      ' Validate
      If (strLabel = "") OrElse (strMatch = "") Then Return ""
      If (strLabel = strMatch) Then Return strLabel
      ' Walk on
      For intPos = 1 To strLabel.Length
        ' Is [strMatch]okay?
        If (intPos > strMatch.Length) Then Return strMatch
        ' Do they still match?
        If (Mid(strLabel, intPos, 1) <> Mid(strMatch, intPos, 1)) Then Return Left(strLabel, intPos - 1)
      Next intPos
      ' Match is complete!
      Return strLabel
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/PartialLabels error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   PartialToCollection
  ' Goal:   Transform array [arThis] into the collection [colThis]
  ' History:
  ' 15-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function PartialToCollection(ByRef strPartial As String, ByRef colThis As StringColl) As Boolean
    Dim arThis() As String  ' Array
    Dim arLine() As String  ' One line
    Dim intI As Integer     ' Counter

    Try
      ' Validate
      If (colThis Is Nothing) Then Return False
      ' Initialize
      arThis = Split(strPartial, ";")
      colThis.Clear()
      ' Walk the result
      For intI = 0 To arThis.Length - 1
        ' Split line in parts
        arLine = Split(arThis(intI), ":")
        ' Add line into collection
        colThis.Add(arLine(0), arLine(1))
      Next intI
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/PartialToCollection error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetLabelList
  ' Goal:   Get a list of all constituents relative to [ndxThis] defined by query [strXpath]
  ' History:
  ' 15-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function GetLabelList(ByRef ndxThis As XmlNode, ByVal strXpath As String, ByRef intNum As Integer) As String
    Dim ndxList As XmlNodeList  ' List of nodes
    Dim strBack As String = ""  ' The list we build
    Dim intI As Integer         ' Counter

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return ""
      ' Get a list of elements and walk them
      ndxList = ndxThis.SelectNodes(strXpath)
      For intI = 0 To ndxList.Count - 1
        ' Add the ancestor node information
        AddSemiStack(strBack, ndxList(intI).Attributes("Label").Value)
      Next intI
      ' Set the amount of children
      intNum = ndxList.Count
      ' Return what we gathered
      Return strBack
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/GetLabelList error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   CleanLabel
  ' Goal:   Strip the label in [strIn] of -1 or -2 and so on
  ' History:
  ' 31-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function CleanLabel(ByVal strIn As String) As String
    Dim strStrip As String  ' The last part of the label

    Try
      ' Validate
      If (Len(strIn) <= 2) Then Return strIn
      ' Check the last part
      strStrip = Right(strIn, 2)
      ' Check this part
      If (Left(strStrip, 1) = "-") AndAlso (IsNumeric(Right(strStrip, 1))) Then
        ' Strip this part 
        Return Left(strIn, strIn.Length - 2)
      End If
      ' Return input
      Return strIn
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/CleanLabel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   AddSyntax
  ' Goal:   Add one parsed constituent to the set of syntax
  ' History:
  ' 31-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function AddSyntax(ByVal strLabel As String, ByVal strChildren As String, ByVal intNum As Integer, _
                             ByVal strExample As String, ByVal strPartial As String) As Boolean
    Dim dtrFound() As DataRow ' Result of select
    Dim dtrNew As DataRow     ' One datarow

    Try
      ' Validate
      ' If (strLabel = "") OrElse (strChildren = "") Then Return False
      If (strLabel = "") Then Return False
      ' Get the datarow with this information
      dtrFound = tdlSyntax.Tables("Const").Select("Label = '" & strLabel.Replace("'", "''") & "' AND Children = '" & strChildren.Replace("'", "''") & "'")
      If (dtrFound.Length = 0) OrElse (strChildren = "") Then
        ' We have to make one, because either (1) there is none yet, or (2) there are no children
        dtrNew = AddOneDataRow(tdlSyntax, "Const", "ConstId", "ConstList")
        dtrNew.Item("Label") = strLabel
        dtrNew.Item("Children") = strChildren
        dtrNew.Item("Example") = strExample
        dtrNew.Item("Partial") = strPartial
        dtrNew.Item("Num") = intNum
        dtrNew.Item("Freq") = 1
      Else
        ' We already have one
        dtrNew = dtrFound(0)
        ' Adapt the frequency
        dtrNew.Item("Freq") += 1
      End If
      ' Return positively 
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/AddSyntax error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   SaveSyntax
  ' Goal:   Save the syntax information to the standard location
  ' History:
  ' 16-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Sub SaveSyntax()
    Try
      ' Validate
      If (strSyntaxFile <> "") AndAlso (tdlSyntax IsNot Nothing) Then
        ' Actually save it
        tdlSyntax.WriteXml(strSyntaxFile)
        ' Give message
        Status("Results saved in: " & strSyntaxFile)
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/SaveSyntax error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   HaveSyntax
  ' Goal:   Initialise syntax and check if we actually have data
  ' History:
  ' 31-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function HaveSyntax() As Boolean
    Try
      If (Not InitSyntax()) Then Return False
      ' Check if data is available
      Return (tdlSyntax.Tables("Const").Rows.Count > 0) AndAlso (tdlSyntax.Tables("Part").Rows.Count > 0)
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/HaveSyntax error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   InitSyntax
  ' Goal:   Initialise syntax processing
  ' History:
  ' 31-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function InitSyntax() As Boolean
    Try
      ' Check
      If (tdlSyntax Is Nothing) Then
        ' Build the name of the syntax file
        strSyntaxFile = GetDocDir() & "\" & SYNT_FILE
        ' Check existence
        If (IO.File.Exists(strSyntaxFile)) Then
          ' Open this file
          If (Not ReadDataset(SYNT_XSD, strSyntaxFile, tdlSyntax)) Then Return False
        Else
          ' Create a dataset
          If (Not CreateDataSet(SYNT_XSD, tdlSyntax)) Then Return False
          ' Clear the <Cons> table
          ClearTable(tdlSyntax.Tables("Const"))
        End If
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/InitSyntax error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   SyntaxReport
  ' Goal:   Give an html report on the constituents found
  ' History:
  ' 31-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function SyntaxReport() As String
    Dim dtrFound() As DataRow     ' Result of select
    Dim colRep As New StringColl  ' Where we store the report
    Dim strText As String         ' One report
    Dim intI As Integer           ' Counter

    Try
      ' Start html
      colRep.Add("<html><body><table><tr><td>Constituent</td><td>NumCh</td><td>Children</td><td>Frequency</td>" & _
                 "<td>Partial</td><td>Example</td></tr>")
      ' Select the elements in right order
      dtrFound = tdlSyntax.Tables("Const").Select("", "Children ASC, Freq DESC")
      ' Walk through all elements
      For intI = 0 To dtrFound.Length - 1
        ' Get this line
        colRep.Add("<tr><td>" & dtrFound(intI).Item("Label") & "</td>" & _
                   "<td align='right'>" & dtrFound(intI).Item("Num") & "</td>" & _
                   "<td>" & dtrFound(intI).Item("Children") & "</td>" & _
                   "<td align='right'>" & dtrFound(intI).Item("Freq") & "</td>" & _
                   "<td align='left'>" & dtrFound(intI).Item("Partial") & "</td>" & _
                   "<td>" & dtrFound(intI).Item("Example") & "</td>" & _
                   "</tr>")
      Next intI
      ' Finish table
      colRep.Add("</table>")
      ' Add partial report
      strText = PartialReport()
      colRep.Add(strText)
      ' Add rules report
      strText = PartialRulesReport()
      colRep.Add(strText)
      ' Finish html
      colRep.Add("</body></html>")
      ' Return the report
      Return colRep.Text
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/SyntaxReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetChildLabels
  ' Goal:   Get the labels of the nodes in [ndxList] starting at [intStart]
  ' History:
  ' 31-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function GetChildLabels(ByRef ndxList As XmlNodeList, ByVal intStart As Integer, ByVal intLen As Integer) As String
    Dim intI As Integer         ' Counter
    Dim strBack As String = ""  ' What we return

    Try
      ' Validate
      If (ndxList Is Nothing) Then Return ""
      If (intStart < 0) OrElse (intStart > ndxList.Count - 1) Then Return ""
      If (intStart + intLen > ndxList.Count) Then Return ""
      ' Construct the list
      For intI = intStart To intStart + intLen - 1
        ' Add this to a semi-stack
        AddSemiStack(strBack, CleanLabel(ndxList(intI).Attributes("Label").Value))
      Next intI
      ' Return the result
      Return strBack
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/GetChildLabels error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   SelectedChildLabels
  ' Goal:   Show the selected labels of the nodes in [ndxList] starting at [intStart]
  ' History:
  ' 31-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function SelectedChildLabels(ByRef ndxList As XmlNodeList, ByVal intStart As Integer, ByVal intLen As Integer) As String
    Dim intI As Integer         ' Counter
    Dim strBack As String = ""  ' What we return

    Try
      ' Validate
      If (ndxList Is Nothing) Then Return ""
      If (intStart < 0) OrElse (intStart > ndxList.Count - 1) Then Return ""
      If (intStart + intLen > ndxList.Count) Then Return ""
      ' Construct the list
      strBack = "["
      For intI = intStart To intStart + intLen - 1
        ' Are we at the beginning?
        If (intI <> intStart) Then strBack &= " - "
        ' Just add the constituent
        strBack &= ndxList(intI).Attributes("Label").Value
      Next intI
      ' Add the last constituent of the selection
      strBack &= "]"
      ' Return the result
      Return strBack
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/SelectedChildLabels error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   SelectedChildLabelsHtml
  ' Goal:   Get html code for the selected labels of the nodes in [ndxList] starting at [intStart]
  ' History:
  ' 31-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function SelectedChildLabelsHtml(ByRef ndxList As XmlNodeList, ByVal intStart As Integer, ByVal intLen As Integer) As String
    Dim colBack As New StringColl ' Html code
    Dim ndxLeaf As XmlNodeList    ' List of leaves
    Dim intI As Integer           ' Counter
    Dim intJ As Integer           ' Another counter

    Try
      ' Validate
      If (ndxList Is Nothing) Then Return ""
      If (intStart < 0) OrElse (intStart > ndxList.Count - 1) Then Return ""
      If (intStart + intLen > ndxList.Count) Then Return ""
      ' Start html
      colBack.Add("<html><body><font color='black'>")
      ' Construct the list
      For intI = 0 To ndxList.Count - 1
        ' Are we at the beginning?
        If (intI <> 0) Then colBack.Add(" ")
        ' Are we at "start"?
        If (intI = intStart) Then
          ' Start bolding and make blue for "selected"
          colBack.Add("<b><font color='blue'>")
        End If
        ' Just add all the <eLeaf> elements under this node
        ndxLeaf = ndxList(intI).SelectNodes("./descendant::eLeaf[@Type='Vern' or @Type='Punct']")
        For intJ = 0 To ndxLeaf.Count - 1
          colBack.Add(ndxLeaf(intJ).Attributes("Text").Value & " ")
        Next intJ
        ' Are we at the end of our selection?
        If (intI = intStart + intLen - 1) Then
          ' Add Switch off bolding
          colBack.Add("</font></b>")
        End If
      Next intI
      ' Finish this line
      colBack.Add("</font></body></html>")
      ' Return the result
      Return colBack.Text
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/SelectedChildLabelsHtml error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   SyntaxMatchesTrigger
  ' Goal:   Check if the indicated node is a trigger for something, and if so, return its datarow
  ' History:
  ' 12-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function SyntaxMatchesTrigger(ByRef ndxThis As XmlNode, ByVal strTrigger As String, _
                                        ByRef arCond() As String) As Boolean
    Dim strLabel As String    ' Label of current node
    Dim strCond As String     ' Condition (optional)
    Dim arPart() As String    ' Two parts of the condition
    Dim intJ As Integer       ' Counter
    Dim bOkay As Boolean      ' COnditions matched?

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      If (ndxThis.Name <> "eTree") Then Return False
      ' Get my label 
      strLabel = ndxThis.Attributes("Label").Value
      ' Is this the trigger?
      If (DoLike(strLabel, strTrigger)) Then
        ' Check all conditions
        bOkay = True
        For intJ = 0 To arCond.Length - 1
          ' What is the condition?
          strCond = MyTrim(arCond(intJ))
          If (strCond <> "") Then
            ' SPlit the condition into two parts
            arPart = Split(strCond, ">>")
            If (arPart.Length <> 2) Then Logging("SyntaxMatchesTrigger: condition must have two parts") : Return False
            arPart(0) = Trim(arPart(0))
            arPart(1) = Trim(arPart(1))
            Try
              ' Make sure the rules are executed inside a tri
              Select Case arPart(0)
                Case "T"
                  ' Try to apply it
                  If (ndxThis.SelectSingleNode(arPart(1), conTb) Is Nothing) Then
                    ' True Condition was not met
                    bOkay = False
                    Return bOkay
                  End If
                Case "F"
                  ' Try to apply it
                  If (ndxThis.SelectSingleNode(arPart(1), conTb) IsNot Nothing) Then
                    ' False Condition was not met
                    bOkay = False
                    Return bOkay
                  End If
                  ' Debug.Print(ndxThis.SelectSingleNode("./following-sibling::eTree[1][tb:matches(@Label, 'NP*')]", conTb) Is Nothing)
              End Select
            Catch ex As Exception
              ' Show that something is wrong
              MsgBox("There is a syntax error in the following line:" & vbCrLf & _
                     arPart(1))
              bInterrupt = True
              Return False
            End Try
          End If
        Next intJ
        ' Return the result
        Return bOkay
        '' Are we still okay?
        'If (bOkay) Then
        '  ' This is the first to fulfill the trigger
        '  Return True
        'End If
      End If
      ' No success
      Return False
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/SyntaxMatchesTrigger error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   SyntaxIsTrigger
  ' Goal:   Check if the indicated node is a trigger for something, and if so, return its datarow
  ' History:
  ' 12-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Function SyntaxIsTrigger(ByRef ndxThis As XmlNode, ByRef intMosId As Integer) As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim strLabel As String    ' Label of current node
    Dim arCond() As String    ' Array of conditions
    Dim strCond As String     ' Condition (optional)
    Dim intI As Integer       ' Counter
    Dim intJ As Integer       ' Counter
    Dim bOkay As Boolean      ' COnditions matched?

    Try
      ' Validate
      If (ndxThis Is Nothing) Then Return False
      If (ndxThis.Name <> "eTree") Then Return False
      If (tdlSettings.Tables("Mos").Rows.Count = 0) Then Return False
      If (tdlSettings.Tables("MosAct").Rows.Count = 0) Then Return False
      ' Get my label 
      strLabel = ndxThis.Attributes("Label").Value
      ' Look
      dtrFound = tdlSettings.Tables("Mos").Select("", "Order ASC, MosId ASC")
      For intI = 0 To dtrFound.Length - 1
        ' ============= DEBUG =============
        If (InStr(dtrFound(intI).Item("Trigger").ToString, "*BAG*") > 0) Then Stop
        ' =================================
        ' Is this the trigger?
        If (DoLike(strLabel, dtrFound(intI).Item("Trigger").ToString)) Then
          ' Check for the condition
          arCond = Split(dtrFound(intI).Item("Cond").ToString, vbLf)
          ' Check all conditions
          bOkay = True
          For intJ = 0 To arCond.Length - 1
            ' What is the condition?
            strCond = MyTrim(arCond(intJ))
            If (strCond <> "") Then
              ' Try to apply it
              If (ndxThis.SelectSingleNode(strCond, conTb) IsNot Nothing) Then
                ' Condition was not met
                bOkay = False : Exit For
              End If
            End If
          Next intJ
          ' Are we still okay?
          If (bOkay) Then
            ' This is the first to fulfill the trigger
            intMosId = dtrFound(intI).Item("MosId")
            Return True
          End If
        End If
      Next intI
      ' No success
      Return False
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/SyntaxIsTrigger error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   SyntaxTryMOS
  ' Goal:   Try to perform the MOS system on the [ndxForest] passed on
  ' Noot:   Let op de volgorde die aangehouden dient te worden:
  '           FOR each $trigger IN Syntax.Triggers ORDER BY PreferredOrder
  '             FOR each $child IN Forest.Children
  '               IF $child.Label MATCHES $trigger THEN
  '                 Perform($trigger.Actions, $child)
  '               END IF
  '             NEXT $child
  '           NEXT $trigger
  ' History:
  ' 12-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function SyntaxTryMOS(ByRef ndxForest As XmlNode) As Boolean
    Dim ndxList As XmlNodeList        ' List of nodes
    Dim ndxJoin As XmlNodeList        ' List of nodes to join
    Dim ndxThis As XmlNode            ' Current main node
    Dim ndxWork As XmlNode = Nothing  ' Working node
    Dim ndxNew As XmlNode = Nothing   ' New node
    Dim dtrMosAct() As DataRow        ' Array of MosAct rows
    Dim dtrFound() As DataRow         ' Result of SELECT
    Dim strOp As String = ""          ' Operation to perform
    Dim strTarget As String = ""      ' Path to target
    Dim strTrigger As String = ""     ' The form of the trigger
    Dim strLabel As String = ""       ' Label for target
    Dim strName As String = ""        ' Name of this operation
    Dim arArg() As String             ' Argument(s) for this operation
    Dim arCond() As String            ' Condition array
    Dim bLeftToRight As Boolean       ' Direction
    Dim bHandled As Boolean = False   ' Obligatory argument
    Dim intI As Integer               ' Counter
    Dim intJ As Integer               ' Counter
    Dim intK As Integer               ' Argument counter
    Dim intM As Integer               ' Counter for triggers
    Dim intMosId As Integer = 0       ' ID of the MOS that matches
    Dim intNum As Integer = 0         ' Number of modifications found
    Dim intDebug As Integer = 0       ' Number for debugging

    Try
      ' Validate
      If (ndxForest Is Nothing) Then Return False
      If (ndxForest.Name <> "forest") Then Return False
      ' Get the list of triggers
      dtrFound = tdlSettings.Tables("Mos").Select("", "Order ASC, MosId ASC")
      For intM = 0 To dtrFound.Length - 1
        ' Get the conditions and the trigger string
        strTrigger = dtrFound(intM).Item("Trigger").ToString
        arCond = Split(dtrFound(intM).Item("Cond").ToString, vbLf)
        intMosId = dtrFound(intM).Item("MosId")
        strName = dtrFound(intM).Item("Name").ToString
        ' Check all the top children until no more modifications can be made
        Do
          ' Walk the top children of [forest] in search for a "trigger"
          ndxList = ndxForest.SelectNodes("./child::eTree")
          ' Reset the number of process
          intNum = 0
          ' ============= DEBUG =============
          intDebug += 1
          Debug.Print("SyntaxTryMOS: reset. Count = " & ndxList.Count & " Debug=" & intDebug)
          ' =================================
          ' Loop through the nodes
          For intI = 0 To ndxList.Count - 1
            ' Get this node
            ndxThis = ndxList(intI)
            ' ====== DEBUG: Show where we are =============
            Debug.Print("SyntaxTryMos: node[" & intI & "]=" & ndxThis.Attributes("Label").Value)
            ' =============================================
            ' Check if this trigger+ conditions match [ndxList]
            If (SyntaxMatchesTrigger(ndxThis, strTrigger, arCond)) Then
              ' Show the latest of what I look like
              ShowTree(frmMain.litTree)
              ' Show what we are doing
              Logging("Perform [" & strName & "] on child #" & intI + 1 & "[" & ndxThis.Attributes("Label").Value & "]")
              'Select Case MsgBox("Perform [" & strName & "] on child #" & intI + 1 & "[" & ndxThis.Attributes("Label").Value & "]", MsgBoxStyle.YesNoCancel)
              '  Case MsgBoxResult.Yes
              '  Case MsgBoxResult.No
              '    Exit For
              '  Case MsgBoxResult.Cancel
              '    Return False
              'End Select
              ' Perform all the necessary actions in order
              dtrMosAct = tdlSettings.Tables("MosAct").Select("MosId=" & intMosId, "Order ASC")
              For intJ = 0 To dtrMosAct.Length - 1
                ' Get the details of this action
                With dtrMosAct(intJ)
                  bLeftToRight = (.Item("Dir") = "LeftToRight")
                  strOp = .Item("Name")
                  arArg = Split(.Item("Arg").ToString, ";")
                End With
                ' Action depends on operator selected
                Select Case strOp
                  Case "Insert"
                    ' There must be two arguments
                    If (arArg.Length <> 2) Then Logging("SyntaxTryMOS: operation [insert] must have two arguments") : Return False
                    ' Get the two arguments
                    strTarget = arArg(0) : strLabel = arArg(1)
                    ' Get the target node
                    ndxWork = ndxThis.SelectSingleNode(strTarget, conTb)
                    If (ndxWork Is Nothing) Then Logging("SyntaxTryMOS: [insert] could not find target at [" & strTarget & "]") : Return False
                    ' Perform the "Insert" operation
                    If (eTreeInsertLevel(ndxThis, ndxNew)) Then
                      ' Process the Ids of this forest
                      If (Not eTreeHandling(ndxForest)) Then Return False
                    End If
                    ' Set the label of the new node
                    If (Not eTreeDoLabel(ndxNew, strLabel, bHandled)) Then Logging("SyntaxTryMOS: problem adding label for [Insert]") : Return False
                    ' Also go up to the newly created <eTree> node with my reference
                    ndxThis = ndxNew
                    ' Adapt processing
                    intNum += 1
                  Case "JoinUnderLeft"
                    ' There must be one argument
                    If (arArg.Length <> 1) Then Logging("SyntaxTryMOS: operation [JoinUnderLeft] must have one argument") : Return False
                    ' Get the argument
                    strTarget = arArg(0)
                    ' Validate
                    If (strTarget = "") Then Logging("SyntaxTryMOS: operation [JoinUnderLeft] must have one argument") : Return False
                    ' Get the target nodes
                    ndxJoin = ndxThis.SelectNodes(strTarget, conTb)
                    If (ndxJoin Is Nothing) Then Logging("SyntaxTryMOS: [JoinUnderLeft] could not find target at [" & strTarget & "]") : Return False
                    ' determine the direction
                    If (bLeftToRight) Then
                      ' Left to right: normal
                      For intK = 0 To ndxJoin.Count - 1
                        ' Get this one
                        ndxWork = ndxJoin(intK)
                        ' Perform the "JoinUnderLeft" operation
                        If (Not eTreeJoinUnder(ndxWork, ndxNew, False)) Then Logging("SyntaxTryMos: [JoinUnderLeft] problem") : Return False
                      Next intK
                    Else
                      ' Right to left: reverse
                      For intK = ndxJoin.Count - 1 To 0 Step -1
                        ' Get this one
                        ndxWork = ndxJoin(intK)
                        ' Perform the "JoinUnderLeft" operation
                        If (Not eTreeJoinUnder(ndxWork, ndxNew, False)) Then Logging("SyntaxTryMos: [JoinUnderLeft] problem") : Return False
                      Next intK
                    End If
                    ' Process the Ids of this forest
                    If (Not eTreeHandling(ndxForest)) Then Return False
                    ' Adapt processing
                    intNum += 1
                  Case "JoinUnderRight"
                    ' There must be one argument
                    If (arArg.Length <> 1) Then Logging("SyntaxTryMOS: operation [JoinUnderRight] must have one argument") : Return False
                    ' Get the argument
                    strTarget = arArg(0)
                    ' Validate
                    If (strTarget = "") Then Logging("SyntaxTryMOS: operation [JoinUnderRight] must have one argument") : Return False
                    ' Get the target nodes
                    ndxJoin = ndxThis.SelectNodes(strTarget, conTb)
                    If (ndxJoin Is Nothing) Then Logging("SyntaxTryMOS: [JoinUnderRight] could not find target at [" & strTarget & "]") : Return False
                    ' determine the direction
                    If (bLeftToRight) Then
                      ' Left to right: normal
                      For intK = 0 To ndxJoin.Count - 1
                        ' Get this one
                        ndxWork = ndxJoin(intK)
                        ' Perform the "JoinUnderRight" operation
                        If (Not eTreeJoinUnder(ndxWork, ndxNew, True)) Then Logging("SyntaxTryMos: [JoinUnderRight] problem") : Return False
                      Next intK
                    Else
                      ' Right to left: reverse
                      For intK = ndxJoin.Count - 1 To 0 Step -1
                        ' Get this one
                        ndxWork = ndxJoin(intK)
                        ' Perform the "JoinUnderRight" operation
                        If (Not eTreeJoinUnder(ndxWork, ndxNew, True)) Then Logging("SyntaxTryMos: [JoinUnderRight] problem") : Return False
                      Next intK
                    End If
                    ' Process the Ids of this forest
                    If (Not eTreeHandling(ndxForest)) Then Return False
                    ' Adapt processing
                    intNum += 1
                  Case "AddLeft"    ' Create a new sibling to my left
                    ' There must be one argument
                    If (arArg.Length <> 1) Then Logging("SyntaxTryMOS: operation [AddLeft] must have one argument") : Return False
                    ' Get the argument: a label for the new node
                    strLabel = arArg(0)
                    ' Perform the "Add" operation
                    If (eTreeAdd(ndxThis, ndxNew, "left")) Then
                      ' Process the Ids of this forest
                      If (Not eTreeHandling(ndxForest)) Then Return False
                    End If
                    ' Set the label of the new node
                    If (Not eTreeDoLabel(ndxNew, strLabel, bHandled)) Then Logging("SyntaxTryMOS: problem adding label for AddLeft") : Return False
                    ' Adapt processing
                    intNum += 1
                  Case "AddRight"   ' Create a new sibling to my right
                    ' There must be one argument
                    If (arArg.Length <> 1) Then Logging("SyntaxTryMOS: operation [AddRight] must have one argument") : Return False
                    ' Get the argument: a label for the new node
                    strLabel = arArg(0)
                    ' Perform the "Add" operation
                    If (eTreeAdd(ndxThis, ndxNew, "right")) Then
                      ' Process the Ids of this forest
                      If (Not eTreeHandling(ndxForest)) Then Return False
                    End If
                    ' Set the label of the new node
                    If (Not eTreeDoLabel(ndxNew, strLabel, bHandled)) Then Logging("SyntaxTryMOS: problem adding label for AddRight") : Return False
                    ' Adapt processing
                    intNum += 1
                  Case "AddChildLast"   ' Create a new child <eTree> and make it the last child
                    ' Adapt processing
                    ' There must be one argument
                    If (arArg.Length <> 1) Then Logging("SyntaxTryMOS: operation [AddChild] must have one argument") : Return False
                    ' Get the argument: a label for the new node
                    strLabel = arArg(0)
                    ' Perform the "Add" operation
                    If (eTreeAdd(ndxThis, ndxNew, "childlast")) Then
                      ' Process the Ids of this forest
                      If (Not eTreeHandling(ndxForest)) Then Return False
                    End If
                    ' Set the label of the new node
                    If (Not eTreeDoLabel(ndxNew, strLabel, bHandled)) Then Logging("SyntaxTryMOS: problem adding label for childlast") : Return False
                    intNum += 1
                  Case "AddChildFirst"   ' Create a new child <eTree> and make it the first child
                    ' Adapt processing
                    ' There must be one argument
                    If (arArg.Length <> 1) Then Logging("SyntaxTryMOS: operation [AddChild] must have one argument") : Return False
                    ' Get the argument: a label for the new node
                    strLabel = arArg(0)
                    ' Perform the "Add" operation
                    If (eTreeAdd(ndxThis, ndxNew, "childfirst")) Then
                      ' Process the Ids of this forest
                      If (Not eTreeHandling(ndxForest)) Then Return False
                    End If
                    ' Set the label of the new node
                    If (Not eTreeDoLabel(ndxNew, strLabel, bHandled)) Then Logging("SyntaxTryMOS: problem adding label for childfirst") : Return False
                    intNum += 1
                  Case "AddEndnode" ' Create a new child <eLeaf>
                    ' There must be two arguments
                    If (arArg.Length <> 2) Then Logging("SyntaxTryMOS: operation [AddLeaf] must have two arguments") : Return False
                    ' Get the two arguments
                    strTarget = arArg(0) : strLabel = arArg(1)
                    ' Get the target node
                    ndxWork = ndxThis.SelectSingleNode(strTarget, conTb)
                    ' Perform the "Add" operation
                    If (eTreeAdd(ndxWork, ndxNew, "leaf")) Then
                      ' Process the Ids of this forest
                      If (Not eTreeHandling(ndxForest)) Then Return False
                    End If
                    ' Set the label of the new node
                    If (Not eTreeDoLabel(ndxNew, strLabel, bHandled)) Then Logging("SyntaxTryMOS: problem adding label for endnode") : Return False
                    ' Adapt processing
                    intNum += 1
                  Case Else
                    ' Give warning 
                    Logging("SyntaxTryMOS: operation [" & strOp & "] is unknown") : Return False
                End Select
              Next intJ
              ' Get out of the [intI] loop, and start over again, until there are no more changes made
              Exit For
            End If
          Next intI
        Loop While intNum > 0

      Next intM
      ' Loop through all possible triggers
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/SyntaxTryMOS error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   SyntaxConst
  ' Goal:   Find a match for [strChildren] (numbered [intNum]) in the learned sample
  ' History:
  ' 31-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function SyntaxConst(ByVal strChildren As String, ByVal intNum As Integer, _
                              ByRef colLabel As StringColl) As Boolean
    Dim dtrFound() As DataRow   ' Result of select
    Dim intI As Integer         ' Counter
    Dim strBack As String = ""  ' What we return

    Try
      ' Validate
      If (strChildren = "") OrElse (intNum = 0) Then Return False
      ' Find the combination
      dtrFound = tdlSyntax.Tables("Const").Select("Num = " & intNum & " AND Children = '" & strChildren & "'", _
                                                  "Freq DESC")
      ' Initialise
      colLabel.Clear()
      ' See how many results we have
      If (dtrFound.Length = 0) Then
        Return False
      Else
        ' Return all the possible labels in decreasing frequency 
        For intI = 0 To dtrFound.Length - 1
          colLabel.Add(dtrFound(intI).Item("Label").ToString, dtrFound(intI).Item("ConstId").ToString)
        Next intI
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/SyntaxConst error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   OneDepFile
  ' Goal:   Learn the constituents from one Psdx file
  ' History:
  ' 31-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function OneDepFile(ByVal strFile As String, ByVal strLang As String, ByRef strResult As String, ByVal bDoCode As Boolean) As Boolean
    Dim ndxFor As XmlNode = Nothing   ' Forest node
    Dim ndxThis As XmlNode = Nothing  ' Current node
    Dim ndxWork As XmlNode            ' Working node
    Dim ndxList As XmlNodeList        ' List of nodes
    Dim strLabel As String = ""       ' Output label of a constituent
    Dim strChildren As String = ""    ' Children of a constituent in semi-stack
    Dim strPartial As String = ""     ' Partial tree seen from me as center
    Dim strBack As String = ""        ' Dummy return string
    Dim strSect As String = ""        ' One section
    Dim intI As Integer               ' Counter
    Dim intPtc As Integer             ' Percentage
    Dim intCount As Integer           ' Number of child nodes
    Dim colResult As New StringColl   ' Result

    Try
      ' Validate
      If (Not IO.File.Exists(strFile)) Then Return False
      ' Read the file as an XML document
      Status("Reading [" & IO.Path.GetFileNameWithoutExtension(strFile) & "]...")
      ' Try read this file into an XML structure
      If (Not ReadXmlDoc(strFile, pdxFile)) Then
        ' There was an error
        Status("Unable to read file: " & strFile)
        Return False
      End If
      ' Walk all the forests in the file
      ' Set the current file
      pdxCurrentFile = pdxFile
      ' Look for the first <forest> element
      If (Not GetFirstForest(pdxFile, ndxFor)) Then Logging("Cannot find first forest") : Return False
      ' Determine how many elements there are potentially
      If (ndxFor Is Nothing) Then
        intCount = 0
      ElseIf (ndxFor.ParentNode Is Nothing) Then
        intCount = 0
      Else
        intCount = ndxFor.ParentNode.ChildNodes.Count
      End If
      ' Walk through all the <forest> elements FORWARD
      While (Not ndxFor Is Nothing)
        ' Check if we are not being interrupted
        If (bInterrupt) Then Return False
        ' Show where we are (how do we KNOW where we are?)
        intPtc = (ndxFor.Attributes("forestId").Value) * 100 \ intCount
        Status("Dependency [" & IO.Path.GetFileNameWithoutExtension(strFile) & "] " & intPtc & "%", intPtc)
        ' Make sure this is a forest and not something else
        If (ndxFor.Name = "forest") Then
          ' =========== DEBUG =============
          ' If (ndxFor.Attributes("forestId").Value = 58) Then Stop
          ' ===============================
          ' Recursively add dependency for this forest
          If (Not TravDep(ndxFor, strLang)) Then
            ' Check for interrupt
            If (bInterrupt) Then Return False
            ' There is a problem!!!
            Logging("modSyntax/OneDepFile: there is a TravNode-problem at @forestId= " & ndxFor.Attributes("forestId").Value)
            Return False
          End If
          ' Add dependency "id" values to all nodes that need it
          ' ndxList = ndxFor.SelectNodes("./descendant::eTree[count(child::eLeaf[not(tb:matches(@Type,'Zero|ZeroSbj|Trace|Star'))])>0]", conTb)
          If (bDoCode) Then
            ' Include CODE
            ndxList = ndxFor.SelectNodes("./descendant::eTree[" & _
              "count(child::eLeaf[tb:matches(@Type,'Vern|Punct')])>0]", conTb)
          Else
            ' Skip CODE
            ndxList = ndxFor.SelectNodes("./descendant::eTree[" & _
              "@Label != 'CODE' and count(child::eLeaf[tb:matches(@Type,'Vern|Punct')])>0]", conTb)
          End If
          For intI = 0 To ndxList.Count - 1
            ' Access this node
            ndxThis = ndxList(intI)
            AddFeature(pdxCurrentFile, ndxThis, "dep", "id", intI + 1)
          Next intI
          ' Remove IDs if they are on the wrong place
          If (bDoCode) Then
            ndxList = ndxFor.SelectNodes("./descendant::eTree[" & _
              "count(child::eLeaf[tb:matches(@Type,'Vern|Punct')])=0]", conTb)
          Else
            ndxList = ndxFor.SelectNodes("./descendant::eTree[" & _
              "@Label = 'CODE' or count(child::eLeaf[tb:matches(@Type,'Vern|Punct')])=0]", conTb)
          End If
          For intI = 0 To ndxList.Count - 1
            ' Access this node
            ndxThis = ndxList(intI).SelectSingleNode("./child::fs[@type='dep']")
            If (ndxThis IsNot Nothing) Then
              ndxWork = ndxThis.SelectSingleNode("./child::f[@name='id']")
              ' Found anything?
              If (ndxWork IsNot Nothing) Then
                ' Remove it
                ndxThis.RemoveChild(ndxWork)
              End If
            End If
          Next intI
          ' Check
          ' Debug.Print("OneDepFile id count=" & ndxFor.SelectNodes("./descendant::eTree[child::fs[@type='dep']/child::f[@name='id']]").Count)
          ' Convert forest to a dependency table in ConLLX format
          If (Not OneForestToConLLX(ndxFor, strSect)) Then
            If (bInterrupt) Then Return False
            Logging("modSyntax/OneDepFIle: could not process [" & IO.Path.GetFileNameWithoutExtension(strFile) & _
                    ":" & ndxFor.Attributes("forestId").Value & "]") : Return False
          End If
          ' Add the result
          colResult.Add(strSect)
        End If
        ' Go to the next forst
        ndxFor = ndxFor.NextSibling
      End While
      ' Save the result
      pdxCurrentFile.Save(strFile)
      ' Add the result to the caller
      strResult = colResult.Text
      ' Return success 
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/OneDepFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoRepairOneDutch
  ' Goal:   Repair Diderot files for the Longdale project (January 2014)
  ' History:
  ' 10-01-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function DoRepairOneDiderot(ByRef ndxFor As XmlNode) As Boolean
    Dim ndxThis As XmlNode      ' Working node
    Dim ndxWork As XmlNode      ' Working node
    Dim ndxId As XmlNodeList    ' ID nodes
    Dim strText As String
    Dim intI As Integer         ' Counter
    Dim intPos As Integer       ' Position
    Dim bChanged As Boolean = False ' Any changes made?

    Try
      ' Validate
      If (ndxFor Is Nothing) Then Return False
      ' ========== Debugging  ==========
      ' Debug.Print("Forest number " & ndxFor.Attributes("forestId").Value)
      ' ================================
      ' Get all 'overlap' <eLeaf> nodes
      ndxId = ndxFor.SelectNodes("./descendant::eLeaf[parent::eTree[@Label = 'META'] and tb:matches(@Text, '*Â*')]", conTb)
      ' Walk all these nodes
      For intI = 0 To ndxId.Count - 1
        ' Get this node
        ndxThis = ndxId(intI)
        ' Check for presence of A-sign 
        strText = ndxThis.Attributes("Text").Value
        intPos = InStr(strText, "Â")
        If (intPos > 0) Then
          While (intPos > 0)
            ' Okay, find the position and remove it
            strText = Left(strText, intPos - 1) & Mid(strText, intPos + 1)
            ' Look for other instance
            intPos = InStr(strText, "Â")
          End While
          ndxThis.Attributes("Text").Value = strText
          ' Indicate changes
          bChanged = True
        End If
      Next intI
      ' Any changes made?
      If (bChanged) Then
        ndxWork = Nothing
        ' Re-calculate forest
        eTreeSentence(ndxFor, ndxWork)
      End If
      ' Return success 
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/DoRepairOneDiderot error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoRepairOneLezgi
  ' Goal:   Repair one <forest> node
  '         - Remove spaces from <eLeaf> and <fs><f> elements
  ' History:
  ' 19-03-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function DoRepairOneLezgi(ByRef ndxFor As XmlNode, ByVal strLang As String) As Boolean
    Dim ndxThis As XmlNode      ' Working node
    Dim ndxLeaf As XmlNodeList  ' List of leaves
    Dim ndxChild As XmlNodeList ' List of children
    Dim strMorph As String      ' The morph
    Dim strPos As String        ' POS
    Dim strCyr As String        ' Cyrillic script
    Dim strLat As String        ' Latin script
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter

    Try
      ' Validate
      If (ndxFor Is Nothing) Then Return False
      ' Get leaves
      ndxLeaf = ndxFor.SelectNodes("./descendant::eLeaf")
      For intI = 0 To ndxLeaf.Count - 1
        ' Initialize
        strMorph = ""
        ' Get the cyrillic text of this one
        ' ndxLeaf(intI).Attributes("Text").Value = Trim(ndxLeaf(intI).Attributes("Text").Value)
        strCyr = Trim(ndxLeaf(intI).Attributes("Text").Value)
        ' Determine the type again
        ndxLeaf(intI).Attributes("Type").Value = GetLeafType(ndxLeaf(intI).Attributes("Text").Value)
        ' Get the node itself
        ndxThis = ndxLeaf(intI).ParentNode : strPos = "X"
        ' Get all the features
        ndxChild = ndxThis.SelectNodes("./descendant::f")
        For intJ = 0 To ndxChild.Count - 1
          ndxChild(intJ).Attributes("value").Value = Trim(ndxChild(intJ).Attributes("value").Value)
          ' Get morph
          If (ndxChild(intJ).Attributes("name").Value = "mrph") Then
            strMorph = ndxChild(intJ).Attributes("value").Value
            strMorph = Regex.Replace(strMorph, "\s+$", "")
          ElseIf (ndxChild(intJ).Attributes("name").Value = "lat") Then
            ' Retrieve the latin text
            strLat = ndxChild(intJ).Attributes("value").Value
            ' Set the text as latin-based
            ndxLeaf(intI).Attributes("Text").Value = Regex.Replace(strLat, "\s+", "")
            ' Set the feature here to cyrillic + give cyrillic text
            ndxChild(intJ).Attributes("name").Value = "cyr"
            ndxChild(intJ).Attributes("value").Value = Regex.Replace(strCyr, "\s+", "")
          End If
        Next intJ
        ' Check if we can guess the POS of this node
        strPos = MorphToPos(strMorph, "")
        ndxThis.Attributes("Label").Value = strPos
      Next intI
      ' Return success 
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/DoRepairOneLezgi error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoRepairOneDutch
  ' Goal:   Repair one <forest> node
  '         - Remove a top "NODE"
  '         - Remove a "VP" layer
  '         - Remove the ID nodes
  '         - Resolve multiple <eTree> children under one <forest>
  ' History:
  ' 27-08-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function DoRepairOneDutch(ByRef ndxFor As XmlNode) As Boolean
    Dim ndxThis As XmlNode      ' Working node
    Dim ndxPar As XmlNode       ' Parent
    Dim ndxVp As XmlNodeList    ' VP nodes
    Dim ndxId As XmlNodeList    ' ID nodes
    Dim ndxChild As XmlNodeList ' List of children
    Dim intI As Integer         ' Counter
    Dim intJ As Integer         ' Counter

    Try
      ' Validate
      If (ndxFor Is Nothing) Then Return False
      ' Check if a "NODE" needs removing
      ndxThis = ndxFor.SelectSingleNode("./child::eTree[@Label = 'NODE']")
      If (ndxThis IsNot Nothing) Then
        ' Walk the children of [ndxThis]
        ndxChild = ndxThis.ChildNodes
        For intI = ndxChild.Count - 1 To 0 Step -1
          ' Append the child to the forest
          ndxFor.PrependChild(ndxChild(intI))
        Next intI
        ' Now remove the node [ndxThis]
        ndxFor.RemoveChild(ndxThis)
      End If
      ' Get all the VP nodes
      ndxVp = ndxFor.SelectNodes("./descendant::eTree[@Label = 'VP']")
      ' Walk all of them
      For intI = 0 To ndxVp.Count - 1
        ' Get the node preceding the VP node as well as the parent of the VP node
        ndxThis = ndxVp(intI).PreviousSibling
        ndxPar = ndxVp(intI).ParentNode
        ' Walk all the child-nodes of the VP
        ndxChild = ndxVp(intI).SelectNodes("./child::eTree")
        For intJ = ndxChild.Count - 1 To 0 Step -1
          ' Move this child
          If (ndxThis Is Nothing) Then
            ' We need to prepend them as children 
            ndxPar.PrependChild(ndxChild(intJ))
          Else
            ' We need to insert them as siblings after [ndxThis]
            ndxPar.InsertAfter(ndxChild(intJ), ndxThis)
          End If
          ' ndxThis.AppendChild(ndxChild(intJ))
        Next intJ
        ' Remove the node
        ndxPar.RemoveChild(ndxVp(intI))
      Next intI
      ' Get all the VP nodes
      ndxId = ndxFor.SelectNodes("./descendant::eTree[@Label = 'ID']")
      ' Walk all of them
      For intI = 0 To ndxId.Count - 1
        ' Get my parent
        ndxPar = ndxId(intI).ParentNode
        ' Remove me
        ndxPar.RemoveChild(ndxId(intI))
      Next intI
      ' Check number of forest children
      ndxChild = ndxFor.SelectNodes("./child::eTree")
      If (ndxChild.Count > 1) Then
        ' Determine parent
        ndxPar = ndxChild(0)
        ' Determine after which node we need to be inserted
        ndxThis = ndxPar.SelectSingleNode("./child::eTree[last()]")
        ' Insert remaining children after [ndxThis]
        For intI = ndxChild.Count - 1 To 1 Step -1
          ' Insert this child
          ndxPar.InsertAfter(ndxChild(intI), ndxThis)
        Next intI
      End If
      ' Return success 
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/DoRepairOneDutch error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   OneUpdateSyntaxFile
  ' Goal:   Update one Chechen file for syntax
  ' History:
  ' 11-02-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function OneUpdateSyntaxFile(ByVal strFile As String) As Boolean
    Dim ndxFor As XmlNode = Nothing   ' Forest node
    Dim ndxThis As XmlNode = Nothing  ' Current node
    Dim ndxList As XmlNodeList        ' List of <eLeaf> nodes
    Dim strLabel As String = ""       ' Output label of a constituent
    Dim strChildren As String = ""    ' Children of a constituent in semi-stack
    Dim strPartial As String = ""     ' Partial tree seen from me as center
    Dim intI As Integer               ' Counter
    Dim intPtc As Integer             ' Percentage
    Dim intCount As Integer           ' Number of child nodes
    Dim intNum As Integer             ' Number of child nodes updated

    Try
      ' Validate
      If (Not IO.File.Exists(strFile)) Then Return False
      ' Read the file as an XML document
      Status("Reading [" & IO.Path.GetFileNameWithoutExtension(strFile) & "]...")
      ' Try read this file into an XML structure
      If (Not ReadXmlDoc(strFile, pdxFile)) Then
        ' There was an error
        Status("Unable to read file: " & strFile)
        Return False
      End If
      ' Walk all the forests in the file
      ' Set the current file
      pdxCurrentFile = pdxFile
      ' Look for the first <forest> element
      If (Not GetFirstForest(pdxFile, ndxFor)) Then Logging("Cannot find first forest") : Return False
      ' Determine how many elements there are potentially
      If (ndxFor Is Nothing) Then
        intCount = 0
      ElseIf (ndxFor.ParentNode Is Nothing) Then
        intCount = 0
      Else
        intCount = ndxFor.ParentNode.ChildNodes.Count
      End If
      ' Initialise
      intNum = 0
      ' Walk through all the <forest> elements FORWARD
      While (Not ndxFor Is Nothing)
        ' Check if we are not being interrupted
        If (bInterrupt) Then Return False
        ' Show where we are (how do we KNOW where we are?)
        intPtc = (ndxFor.Attributes("forestId").Value) * 100 \ intCount
        Status("Learning [" & IO.Path.GetFileNameWithoutExtension(strFile) & "] " & intPtc & "%", intPtc)
        ' Make sure this is a forest and not something else
        If (ndxFor.Name = "forest") Then
          ' (1) Any forest NP children should have the "NP" only label
          ndxList = ndxFor.SelectNodes("./child::eTree[tb:matches(@Label, 'NP-*')]", conTb)
          For intI = 0 To ndxList.Count - 1
            ' Change the label to a simple NP
            ndxList(intI).Attributes("Label").Value = "NP" : intNum += 1
          Next intI
          ' (2) PP-child NPs should have a label of type NP-OB
          ndxList = ndxFor.SelectNodes("./descendant::eTree[tb:matches(@Label, 'NP|NP-*') " & _
                                       " and count(parent::eTree[tb:matches(@Label, 'PP|PP-*')])>0]", conTb)
          For intI = 0 To ndxList.Count - 1
            If (ndxList(intI).Attributes("Label").Value <> "NP-OB") Then
              ' Change the label to a NP-OB label
              ndxList(intI).Attributes("Label").Value = "NP-OB" : intNum += 1
            End If
          Next intI
        End If
        ' Go to the next forst
        ndxFor = ndxFor.NextSibling
      End While
      ' Keep track of the number of changes
      Logging(IO.Path.GetFileNameWithoutExtension(strFile) & ": " & intNum & " changes")
      If (intNum > 0) Then
        ' Save changes
        pdxCurrentFile.Save(strFile)
      End If
      ' Return success 
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/OneUpdateSyntaxFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuSyntaxLezgiPos_Click
  ' Goal:   Perform POS tagging for language [strLang]
  '         (1) Do POS tagging using a <liftpos> dictionary (if available)
  '         (2) Build a list of Vern-POS matches in this text
  '         (3) Check all constituents semi-automatically
  ' History:
  ' 19-03-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function SyntaxPosTag(ByVal strLang As String, Optional ByVal bAuto As Boolean = False, _
                               Optional ByVal bDoAll As Boolean = False) As Boolean
    Dim strThisLangDict As String        ' Name of dictionary
    Dim strEng As String              ' English language
    Dim strOrg As String              ' Original
    Dim strVern As String             ' Vernacular
    Dim strPos As String              ' Part of speech
    Dim strGloss As String            ' Gloss for this word
    Dim ndxList As XmlNodeList        ' List of forest files
    Dim ndxConst As XmlNodeList       ' List of constituents
    Dim tdlThisLang As DataSet = Nothing ' Dictionary for Lezgi
    Dim dtrThis As DataRow            ' One row
    Dim dtrFound() As DataRow         ' Result of selecting
    Dim intCount As Integer = 0       ' Number of additions
    Dim intI As Integer               ' Counter
    Dim intJ As Integer               ' Counter

    Try
      ' Validate
      If (pdxCurrentFile Is Nothing) Then Status("First load a psdx file") : Return False
      ' Initialize the language
      If (Not DoLangInit(strLang)) Then Status("Could not initialize language [" & strLang & "]") : Return False
      ' Create dictionary according to a set scheme
      If (Not CreateDataSet("MorphDict.xsd", tdlThisLang)) Then Return False
      ' (1) Stage 1: Does this language have a <liftpos> dictionary?
      If (tdlLiftPosDic IsNot Nothing) Then
        ' Walk all the forest files
        ndxList = pdxCurrentFile.SelectNodes("//forest")
        For intI = 0 To ndxList.Count - 1
          ' Walk all the X-marked nodes
          If (bDoAll) Then
            ndxConst = ndxList(intI).SelectNodes("./descendant::eTree")
          Else
            ndxConst = ndxList(intI).SelectNodes("./descendant::eTree[@Label = 'X']")
          End If
          ' Walk them all
          For intJ = 0 To ndxConst.Count - 1
            ' Get the type
            Dim strTokenType As String = ndxConst(intJ).SelectSingleNode("./child::eLeaf").Attributes("Type").Value
            Select Case strTokenType
              Case "Vern"
                ' Get the word
                Dim strWord As String = ndxConst(intJ).SelectSingleNode("./child::eLeaf").Attributes("Text").Value
                Dim strDef As String = ""
                Dim strLat As String = ""
                Dim strLemma As String = ""
                ' Look up the word in tdlLiftPosDic
                strPos = getPosFromLift(strWord, InStr(strLang.ToLower, "cyr") > 0, strDef, strLat, strLemma)
                ' Any positive result?
                If (strPos <> "") Then
                  ndxConst(intJ).Attributes("Label").Value = strPos
                  AddFeature(pdxCurrentFile, ndxConst(intJ), "M", "d", strDef)
                  AddFeature(pdxCurrentFile, ndxConst(intJ), "M", "l", strLemma)
                  AddFeature(pdxCurrentFile, ndxConst(intJ), "lift", "lat", strLat)
                  ' Make sure the 'dirty' flag is set for this document
                  frmMain.SetDirty(True)
                  intCount += 1
                End If
            End Select
          Next intJ
        Next intI
      End If

      ' (2) stage 2: look for all words that now HAVE a POS tag
      ' Walk all the forest files
      ndxList = pdxCurrentFile.SelectNodes("//forest")
      For intI = 0 To ndxList.Count - 1
        ' Walk all the non X-marked nodes
        ndxConst = ndxList(intI).SelectNodes("./descendant::eTree[@Label != 'X']")
        ' Walk them all
        For intJ = 0 To ndxConst.Count - 1
          ' What do we have?
          Dim strWord As String = ndxConst(intJ).SelectSingleNode("./child::eLeaf").Attributes("Text").Value
          strPos = ndxConst(intJ).Attributes("Label").Value
          ' Is this in there already?
          If (tdlThisLang.Tables("Morph").Select("Vern='" & strWord.Replace("'", "''") & _
                                                 "' AND Pos='" & strPos & "'").Length = 0) Then
            ' Add this combination
            dtrThis = AddOneDataRow(tdlThisLang, "Morph", "MorphId", "MorphList")
            dtrThis.Item("Vern") = strWord : dtrThis.Item("Pos") = strPos
          End If
        Next intJ
      Next intI
      ' Save the dictionary so far
      strThisLangDict = GetDocDir() & "\" & strLang & "Dict.xml"
      tdlThisLang.WriteXml(strThisLangDict)
      Logging("Dictionary: " & strThisLangDict)
      ' Walk all X-marked ones
      ' Walk all the forest files
      ndxList = pdxCurrentFile.SelectNodes("//forest")
      For intI = 0 To ndxList.Count - 1
        ' Show the line
        Status("ForestId = " & ndxList(intI).Attributes("forestId").Value)
        ' Get this line's original and translation
        strOrg = GetSeg(ndxList(intI), "org")
        strEng = GetSeg(ndxList(intI), "eng")
        ' Walk all the non X-marked nodes
        ndxConst = ndxList(intI).SelectNodes("./descendant::eTree[@Label = 'X' and count(child::eLeaf[@Type='Vern'])>0]")
        ' Walk them all
        For intJ = 0 To ndxConst.Count - 1
          ' Get the vernacular
          strVern = ndxConst(intJ).SelectSingleNode("./child::eLeaf").Attributes("Text").Value
          ' Try get the gloss
          strGloss = GetFeature(ndxConst(intJ), "M", "mrph")
          ' See if we can estimate the POS
          dtrFound = tdlThisLang.Tables("Morph").Select("Vern = '" & strVern.Replace("'", "''") & "'")
          If (dtrFound.Length > 0) Then
            strPos = dtrFound(0).Item("Pos").ToString
            ' Add the POS to the PSDX file
            ndxConst(intJ).Attributes("Label").Value = strPos
            ' Add the POS to the PSDX file
            ndxConst(intJ).Attributes("Label").Value = strPos
            ' Make sure the 'dirty' flag is set for this document
            frmMain.SetDirty(True)
            intCount += 1
          ElseIf (Not bAuto) Then
            strPos = "N"
            ' Ask for this one
            strPos = InputBox("Please provide the POS of [" & strVern & "]=[" & strGloss & "] in:" & vbCrLf & _
                             vbCrLf & strOrg & vbCrLf & vbCrLf & strEng & vbCrLf, "POS provider", strPos)
            If (strPos = "") Then
              ' Ask user for quits
              Select Case MsgBox("Would you like to quit?", MsgBoxStyle.YesNoCancel)
                Case MsgBoxResult.Yes
                  Status("Done (but you need to save the results)")
                  Return False
              End Select
            Else
              ' Add the POS to the PSDX file
              ndxConst(intJ).Attributes("Label").Value = strPos
              ' Dictionary work...
              dtrFound = tdlThisLang.Tables("Morph").Select("Vern = '" & strVern.Replace("'", "''") & "' AND Pos='" & strPos & "'")
              If (dtrFound.Length = 0) Then
                ' Add this combination
                dtrThis = AddOneDataRow(tdlThisLang, "Morph", "MorphId", "MorphList")
                dtrThis.Item("Vern") = strVern
                dtrThis.Item("Pos") = strPos
              End If
              ' Add the POS to the PSDX file
              ndxConst(intJ).Attributes("Label").Value = strPos
            End If
          End If

        Next intJ
      Next intI
      ' Show we are ready
      Logging("Ready - automatically added: " & intCount, True)
      ' Return success 
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSyntax/SyntaxPosTag error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
End Module
