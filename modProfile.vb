Module modProfile
  ' ========================================= LOCAL VARIABLES ===========================================
  Private tblProfile As DataTable = Nothing
  ' ========================================= GLOBAL VARIABLES ==========================================
  ' =====================================================================================================
  ' ------------------------------------------------------------------------------------
  ' Name:   ProfileInit
  ' Goal:   Make the table or clear it
  ' History:
  ' 19-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub ProfileInit()
    Try
      ' Does the table already exist?
      If (tblProfile Is Nothing) Then
        ' Create the table
        tblProfile = New DataTable("Profile")
        ' Make the correct columns
        tblProfile.Columns.Add("Id", System.Type.GetType("System.String"))
        tblProfile.Columns.Add("Count", System.Type.GetType("System.Double"))
        tblProfile.Columns.Add("Start", System.Type.GetType("System.Double"))
        tblProfile.Columns.Add("Stop", System.Type.GetType("System.Double"))
      Else
        ' Clear the rows for this table
        ClearTable(tblProfile)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modProfile/ProfileInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   ProfileGetRow
  ' Goal:   Start counting time in (seconds, milliseconds?) for the indicated
  '           identifier [strId]
  ' History:
  ' 19-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function ProfileGetRow(ByVal strId As String) As DataRow
    Dim dtrFound() As DataRow ' Result of select
    Dim dtrThis As DataRow    ' New row to be added

    Try
      ' Does the table already exist?
      If (tblProfile Is Nothing) Then ProfileInit()
      ' Check whether this constituent exists
      dtrFound = tblProfile.Select("Id='" & strId & "'")
      If (dtrFound.Length = 0) Then
        ' Make a new row for this constituent
        dtrThis = tblProfile.NewRow
        With dtrThis
          .Item("Id") = strId
          .Item("Count") = 0
        End With
        ' Add row to table
        tblProfile.Rows.Add(dtrThis)
        ' Return the result
        Return dtrThis
      Else
        ' Return the correct row
        Return dtrFound(0)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modProfile/ProfileGetRow error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return Nothing
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ProfileAdd
  ' Goal:   Add the amount of [intCount] to the time for identifier [strId]
  ' History:
  ' 19-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub ProfileAdd(ByVal strId As String, ByVal intCount As Integer)
    Dim dtrThis As DataRow  ' Result of SELECT

    Try
      ' Validate
      If (tblProfile Is Nothing) Then ProfileInit()
      ' Try find the correct row
      dtrThis = ProfileGetRow(strId)
      ' Add the amount
      dtrThis.Item("Count") += intCount
    Catch ex As Exception
      ' Show error
      HandleErr("modProfile/ProfileStart error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   ProfileStart
  ' Goal:   Set a start time for identifier [strId]
  ' History:
  ' 19-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub ProfileStart(ByVal strId As String, Optional ByVal intLevel As Integer = 1)
    Dim dtrThis As DataRow  ' Result of SELECT

    Try
      ' Check level
      If (intLevel = 0) Then Exit Sub
      If (intLevel > intProfileLevel) Then Exit Sub
      ' Validate
      If (tblProfile Is Nothing) Then ProfileInit()
      ' Try find the correct row
      dtrThis = ProfileGetRow(strId)
      ' Add the amount
      dtrThis.Item("Start") = DateAndTime.Timer
    Catch ex As Exception
      ' Show error
      HandleErr("modProfile/ProfileStart error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   ProfileEnd
  ' Goal:   Set a stop/end time for identifier [strId]
  ' History:
  ' 19-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub ProfileEnd(ByVal strId As String, Optional ByVal intLevel As Integer = 1)
    Dim dtrThis As DataRow    ' Result of SELECT

    Try
      ' Check level
      If (intLevel = 0) Then Exit Sub
      If (intLevel > intProfileLevel) Then Exit Sub
      ' Validate
      If (tblProfile Is Nothing) Then ProfileInit()
      ' Try find the correct row
      dtrThis = ProfileGetRow(strId)
      ' Add the amount
      dtrThis.Item("Stop") = DateAndTime.Timer
      ' Check if there is a start time >0 and lower than Stop
      If (dtrThis.Item("Start") > 0) AndAlso (dtrThis.Item("Start") < dtrThis.Item("Stop")) Then
        ' Calculate the difference between start and stop in seconds (?)
        dtrThis.Item("Count") += dtrThis.Item("Stop") - dtrThis.Item("Start")
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modProfile/ProfileEnd error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   ProfileOview
  ' Goal:   Make an HTML overview of the profile
  ' History:
  ' 19-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function ProfileOview() As String
    Dim strHtml As String ' HTML return
    Dim intI As Integer   ' Counter

    Try
      ' Start html
      strHtml = "<html><body>" & vbCrLf
      ' Start table
      strHtml &= "<table><tr><td>Identifier</td><td>Time</td></tr>" & vbCrLf
      ' Go through all the rows
      For intI = 0 To tblProfile.Rows.Count - 1
        ' Add this row
        strHtml &= "<tr><td>" & tblProfile(intI).Item("Id") & "</td>" & _
          "<td>" & Format(tblProfile(intI).Item("Count"), "0.00") & "</td>" & _
          "</tr>" & vbCrLf
      Next intI
      ' Finish table
      strHtml &= "</table>"
      ' Finish html
      strHtml &= "</body></html>" & vbCrLf
      ' Return the result
      Return strHtml
    Catch ex As Exception
      ' Show error
      HandleErr("modProfile/ProfileOview error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return ""
    End Try
  End Function

End Module
