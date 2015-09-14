Imports System.Xml
Module modFullStat
  ' ================================= LOCAL VARIABLES =========================================================
  Private loc_dblNPcount As Double     ' Total number of NPs
  Private loc_dblNPcoref As Double     ' Total number of NPs for which coreference was made
  Private loc_intNP
  Private tblFullStat As DataTable = Nothing  ' The table with the FULL statistics
  Private strAutoStatScheme As String = "AutoStat.xsd"
  Private strHtmlHead As String = "<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8' /></head>"
  ' ============================================================================================================
  ' ------------------------------------------------------------------------------------
  ' Name:   InitFullStat
  ' Goal:   Initialise the statistics of fully automatically resolved coreferences
  ' History:
  ' 21-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub InitFullStat()
    Try
      ' Does a table already exist?
      If (tblFullStat Is Nothing) Then
        ' Does not exist, so create one
        tblFullStat = New DataTable("FullStat")
        ' Create the columns
        tblFullStat.Columns.Add("RefType", Type.GetType("System.String"))
        tblFullStat.Columns.Add("Constraint", Type.GetType("System.String"))
        tblFullStat.Columns.Add("IPdist", Type.GetType("System.Int32"))
        tblFullStat.Columns.Add("LinkType", Type.GetType("System.String"))
        tblFullStat.Columns.Add("Count", Type.GetType("System.Int32"))
      Else
        ' Already exists, so only reset the rows
        tblFullStat.Rows.Clear()
      End If
      ' Other initialisations
      ' (1) total number of NPs
      loc_dblNPcount = 0
      ' (2) Total number of NPs for which coreference was determined
      loc_dblNPcoref = 0
    Catch ex As Exception
      ' Show error
      HandleErr("modFullStat/InitFullStat error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   ReadAutoStat
  ' Goal:   Read the indicated AutoStat file into the [tdlAutoStat] dataset
  ' History:
  ' 21-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function ReadAutoStat(ByVal strFile As String) As Boolean
    Try
      ' Validate
      If (strFile = "") Then Return False
      If (Not IO.File.Exists(strFile)) Then Return False
      ' Read the file into the [tdl] dataset
      Return (ReadDataset(strAutoStatScheme, strFile, tdlAutoStat))
    Catch ex As Exception
      ' Show error
      HandleErr("modFullStat/ReadAutoStat error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   OneFullStatFile
  ' Goal:   Process one -stat.xml file
  ' History:
  ' 21-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function OneFullStatFile(ByVal strFile As String) As Boolean
    Dim tblStat As DataTable    ' The [Stat] table
    Dim intI As Integer         ' Counter
    Dim strRefType As String    ' Reference type
    Dim strConstraint As String ' The last used constraint
    Dim strLinkType As String   ' The kind of link made (User, Auto etc)
    Dim intIPdist As Integer    ' The IP distance used

    Try
      ' Get the file into the dataset
      If (Not ReadAutoStat(strFile)) Then Return False
      ' Get all the [Stat] elements
      tblStat = tdlAutoStat.Tables("Stat")
      ' Double check
      If (tblStat Is Nothing) Then Return False
      ' Go through it
      For intI = 0 To tblStat.Rows.Count - 1
        ' Make sure other processes get time too
        Application.DoEvents()
        ' Process the information in this row
        With tblStat(intI)
          strRefType = .Item("RefType")
          strConstraint = GetConstraint(.Item("Constraint"))
          intIPdist = .Item("IPdist")
          strLinkType = .Item("LinkType")
        End With
        ' Store this information in the necessary statistics row
        AddFullStat(strRefType, strLinkType, strConstraint, intIPdist)
        ' One more NP was found
        loc_dblNPcoref += 1
      Next intI
      ' Also add the total number of NPs
      tblStat = tdlAutoStat.Tables("StatList")
      If (tblStat Is Nothing) Then Return False
      If (tblStat.Rows.Count = 0) Then Return False
      ' We need to have the LAST [StatList] row in order to get the total amount of NPs
      With tblStat
        loc_dblNPcount += .Rows(.Rows.Count - 1).Item("NPcount")
      End With
      'For intI = 0 To tblStat.Rows.Count - 1
      '  ' Add the information from this section
      '  loc_dblNPcount += tblStat.Rows(intI).Item("NPcount")
      'Next
      ' Try to get the current period
      If (Not HasPeriod(tblStat.Rows(0).Item("File"), strCurrentPeriod)) Then
        ' Unable to get the period
        strCurrentPeriod = "Unable to get the current period"
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modFullStat/OneFullStatFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   FullStatReport
  ' Goal:   Give a HTML report of the statistics
  ' History:
  ' 21-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function FullStatReport() As String
    Dim dtrFound() As DataRow   ' Result of SELECT statement
    Dim intI As Integer         ' Counter
    Dim strHtml As String = "<html><body>" & vbCrLf

    Try
      ' Validate
      If (tblFullStat Is Nothing) Then Return ""
      ' Add preamble
      strHtml &= "<table><tr><td>Date:</td><td>" & Format(Now, "f") & "</td></tr>" & vbCrLf
      strHtml &= "<tr><td>User:</td><td>" & strUserName & "</td></tr>" & vbCrLf
      strHtml &= "<tr><td>Period:</td><td>" & strCurrentPeriod & "</td></tr>" & vbCrLf
      strHtml &= "<tr><td>Total number of NPs:</td><td>" & loc_dblNPcount & "</td></tr>" & vbCrLf
      strHtml &= "<tr><td>Number of coref resolutions:</td><td>" & loc_dblNPcoref & "</td></tr>" & vbCrLf
      strHtml &= "<tr><td>Percentage resolved:</td><td>" & Format(100 * loc_dblNPcoref / loc_dblNPcount, "0.00") & "</td></tr>" & vbCrLf
      ' Finish this table
      strHtml &= "</table><p><p>" & vbCrLf
      ' Add a subtable for LinkType
      strHtml &= SubTable("LinkType")
      ' Add a subtable for RefType
      strHtml &= SubTable("RefType")
      ' Add a subtable for Constraint
      strHtml &= SubTable("Constraint")
      ' Add a subtable for IPdist
      strHtml &= SubTable("IPdist")
      ' Add a subtable to see what constraints were used per RefType
      strHtml &= TwoTable("RefType", "Constraint")
      ' Add a subtable to see what constraints were used per LinkType
      strHtml &= TwoTable("LinkType", "Constraint")
      ' Start making a table of the full results
      dtrFound = tblFullStat.Select("", "RefType ASC, Constraint ASC, IPdist DESC, Count DESC")
      strHtml &= "<table><tr><td>RefType</td><td>LinkType</td><td>Constraint</td><td>IPdist</td><td>Frequency</td></tr>" & vbCrLf
      ' Go through all the rows
      For intI = 0 To dtrFound.Length - 1
        ' Add one line in the table
        With dtrFound(intI)
          strHtml &= "<tr><td>" & .Item("RefType") & "</td>" & _
            "<td>" & .Item("LinkType") & "</td>" & _
           "<td>" & .Item("Constraint") & "</td>" & _
            "<td>" & .Item("IPdist") & "</td>" & _
            "<td>" & .Item("Count") & "</td>" & _
            "</tr>" & vbCrLf
        End With
      Next intI
      ' Finish the table
      strHtml &= "</table>" & vbCrLf
      ' Finish the htmla
      strHtml &= "</body></html>" & vbCrLf
      ' Return the result
      Return strHtml
    Catch ex As Exception
      ' Show error
      HandleErr("modFullStat/FullStatReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SubTable
  ' Goal:   One subtable for the statistics summary
  ' History:
  ' 21-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function SubTable(ByVal strType As String) As String
    Dim dtrFound() As DataRow   ' Result of SELECT statement
    Dim intI As Integer         ' Counter
    Dim dblCount As Double      ' Count of result
    Dim strItem As String       ' One item
    Dim strHtml As String = ""  ' Our output

    Try
      ' Start making a table of the <RefType> results
      dtrFound = tblFullStat.Select("", strType & " ASC, Count DESC")
      strItem = ""
      ' Start a table
      strHtml = "<table><tr><td>" & strType & "</td><td>Frequency</td><td>Count</td></tr>" & vbCrLf
      ' Go through all the rows
      For intI = 0 To dtrFound.Length - 1
        ' Access one line in the table
        With dtrFound(intI)
          ' Should we output this line?
          If (strItem <> "") AndAlso (strItem <> .Item(strType)) Then
            ' We should output the previous item
            strHtml &= "<tr><td>" & strItem & "</td>" & _
              "<td>" & Format(100 * dblCount / loc_dblNPcoref, "0.00") & "</td>" & _
              "<td>" & Format(dblCount, "0.00") & "</td>" & _
              "</tr>" & vbCrLf
            ' Reset the counter
            dblCount = 0
          End If
          ' Set the item correct
          strItem = .Item(strType)
          ' Increment the counter appropriately
          dblCount += .Item("Count")
        End With
      Next intI
      ' Add the last row
      strHtml &= "<tr><td>" & strItem & "</td>" & _
        "<td>" & Format(100 * dblCount / loc_dblNPcoref, "0.00") & "</td>" & _
        "<td>" & Format(dblCount, "0.00") & "</td>" & _
        "</tr>" & vbCrLf
      ' Finish the table
      strHtml &= "</table><p><p>" & vbCrLf
      ' Return the result
      Return strHtml
    Catch ex As Exception
      ' Show error
      HandleErr("modFullStat/SubTable error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   TwoTable
  ' Goal:   One table divided over two kinds (strType > strSub) for the statistics summary
  ' History:
  ' 21-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function TwoTable(ByVal strType As String, ByVal strSub As String) As String
    Dim dtrFound() As DataRow   ' Result of SELECT statement
    Dim intI As Integer         ' Counter
    Dim dblCount As Double      ' Count of result
    Dim strItem As String       ' One item
    Dim strHtml As String = ""  ' Our output

    Try
      ' Start making a table of the <RefType> results
      dtrFound = tblFullStat.Select("", strType & " ASC, " & strSub & " ASC, Count DESC")
      strItem = ""
      ' Start a table
      strHtml = "<table><tr><td>" & strType & "/" & strSub & "</td><td>Frequency</td><td>Count</td></tr>" & vbCrLf
      ' Go through all the rows
      For intI = 0 To dtrFound.Length - 1
        ' Access one line in the table
        With dtrFound(intI)
          ' Should we output this line?
          If (strItem <> "") AndAlso (strItem <> .Item(strType) & "_" & .Item(strSub)) Then
            ' We should output the previous item
            strHtml &= "<tr><td>" & strItem & "</td>" & _
              "<td>" & Format(100 * dblCount / loc_dblNPcoref, "0.00") & "</td>" & _
              "<td>" & Format(dblCount, "0.00") & "</td>" & _
              "</tr>" & vbCrLf
            ' Reset the counter
            dblCount = 0
          End If
          ' Set the item correct
          strItem = .Item(strType) & "_" & .Item(strSub)
          ' Increment the counter appropriately
          dblCount += .Item("Count")
        End With
      Next intI
      ' Add the last row
      strHtml &= "<tr><td>" & strItem & "</td>" & _
        "<td>" & Format(100 * dblCount / loc_dblNPcoref, "0.00") & "</td>" & _
        "<td>" & Format(dblCount, "0.00") & "</td>" & _
        "</tr>" & vbCrLf
      ' Finish the table
      strHtml &= "</table><p><p>" & vbCrLf
      ' Return the result
      Return strHtml
    Catch ex As Exception
      ' Show error
      HandleErr("modFullStat/TwoTable error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddFullStat
  ' Goal:   Add one line in statistics and/or increment its freq. of occurrance
  ' History:
  ' 21-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub AddFullStat(ByVal strRefType As String, ByVal strLinkType As String, _
                          ByVal strConstraint As String, ByVal intIPdist As Integer)
    Dim dtrThis As DataRow    ' New datarow
    Dim dtrFound() As DataRow ' Result of SELECT function

    Try
      ' Validate
      If (tblFullStat Is Nothing) Then InitFullStat()
      ' Find out if this combination already exists
      dtrFound = tblFullStat.Select("RefType='" & strRefType & "' AND Constraint='" & strConstraint & _
                                    "' AND IPdist=" & intIPdist & " AND LinkType='" & strLinkType & "'")
      If (dtrFound.Length = 0) Then
        ' Nothing found, so make a new row
        dtrThis = tblFullStat.NewRow
        With dtrThis
          .Item("RefType") = strRefType
          .Item("LinkType") = strLinkType
          .Item("Constraint") = strConstraint
          .Item("IPdist") = intIPdist
          .Item("Count") = 0
        End With
        ' Add the row
        tblFullStat.Rows.Add(dtrThis)
      Else
        ' Found it
        dtrThis = dtrFound(0)
      End If
      ' Increment the frequency
      dtrThis.Item("Count") += 1
    Catch ex As Exception
      ' Show error
      HandleErr("modFullStat/AddFullStat error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   GetConstraint
  ' Goal:   Derive the constraint
  ' History:
  ' 21-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function GetConstraint(ByVal strIn As String) As String
    Dim intPos As Integer ' Position in string
    Dim strCns As String  ' One constraint

    Try
      ' Validate
      If (strIn = "") Then Return "(none)"
      ' Does it contain semicolon?
      intPos = InStrRev(strIn, ";")
      If (intPos = 0) Then
        ' No, so take this as the constraint
        strCns = strIn
      Else
        ' Yes, so get the last constraint
        strCns = Mid(strIn, intPos + 1)
      End If
      ' Does it contain '['
      intPos = InStr(strCns, "[")
      If (intPos > 0) Then
        ' Strip off the [
        strCns = Left(strCns, intPos - 1)
      End If
      ' Double check
      If (InStr(strCns, "[") > 0) Then Stop
      ' Return the result
      Return strCns
    Catch ex As Exception
      ' Show error
      HandleErr("modFullStat/GetConstraint error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function
End Module
