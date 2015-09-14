Imports Microsoft.Office.Interop
Module modWord
  '---------------------------------------------------------------
  ' Name:       GetWord()
  ' Goal:       Get a new instance of WORD or a running instance
  ' History:
  ' 17-01-2008  ERK Created
  '---------------------------------------------------------------
  Public Function GetWord(ByRef objThis As Word.Application, Optional ByVal bCreate As Boolean = False) As Boolean
    Try
      ' Catch errors
      Try
        'First try to open an existing instance
        objThis = GetObject(, "Word.Application")
      Catch ex As Exception
        ' No existing instance, so try opening a new one
        If (bCreate) Then
          Try
            objThis = New Word.Application
          Catch ex2 As Exception
            MsgBox("Joozanash/modMain/GetWord error:" & vbCrLf & ex2.Message)
            Return False
          End Try
        Else
          Return False
        End If
      End Try
      objThis.Visible = True
      objThis.WindowState = Word.WdWindowState.wdWindowStateMaximize
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modWord/GetWord error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   DissertationList
  ' Goal:   Produce a list for the dissertation (provided it is open)
  ' History:
  ' 28-11-2012  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub DissertationList(ByRef strBack As String)
    Dim objWord As Word.Application = Nothing

    Try
      ' Try to access the word that is right now open
      If (Not GetWord(objWord)) Then Status("Could not open WORD") : Exit Sub
      ' Get a list of endnote references
      If (Not EndnoteRefs(objWord.ActiveDocument, strBack)) Then Exit Sub
      ' Show we are successfull
      Status("Ready")
    Catch ex As Exception
      ' Show error
      HandleErr("modWord/DissertationList error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' -----------------------------------------------------------------------------------------------
  ' Name: EndnoteRefs
  ' Goal: Provide a list of Endnote references and the sections they occur in
  ' History:
  ' 25-11-2012  ERK Created
  ' -----------------------------------------------------------------------------------------------
  Private Function EndnoteRefs(ByRef docThis As Word.Document, ByRef strBack As String) As Boolean
    'Dim strHead As String     ' Name of the heading we are under
    'Dim strType As String     ' Type of the field
    Dim strText As String     ' Text of the field
    'Dim strSect As String     ' Section (internal numbering)
    Dim strLoc As String      ' Location
    Dim strAut As String      ' AUthor
    Dim strYr As String       ' Year
    Dim strTitle As String    ' Title of the publication
    Dim strFile As String     ' Where we return it
    Dim rngThis As Word.Range ' Selection
    Dim intPage As Integer    ' Page number
    Dim intI As Integer       ' Counter
    Dim intType As Integer    ' Type of field
    Dim objBack As New StringColl

    Try
      ' Set the output file
      strFile = "d:\data files\publications\dissertation\Oview.htm"
      With frmMain.SaveFileDialog1
        .InitialDirectory = IO.Path.GetDirectoryName(docThis.FullName)
        .Filter = "Html file (*.htm)|*.htm"
        .FileName = docThis.Name
        Select Case .ShowDialog
          Case DialogResult.OK, DialogResult.Yes
            strFile = .FileName
          Case Else
            Return False
        End Select
      End With
      ' Start preparing return
      objBack.Add("<html><body><table>")
      ' Access all the fields in the active document
      With docThis.Fields
        ' Walk through all the fields
        For intI = 1 To .Count
          ' Get the field type
          intType = .Item(intI).Type
          ' Show where we are
          ' Debug.Print("Processing " & Format(intI, "0000") & "/" & .Count)
          Status("Processing " & Format(intI, "0000") & "/" & .Count & " type=" & Format("000", intType))
          ' Make sure we are interrupted
          Application.DoEvents()
          ' Get the type of this field
          Select Case intType
            Case Word.WdFieldType.wdFieldAddin
              ' Select the item
              .Item(intI).Select()
              ' Get the text of this style
              strText = Trim(.Item(intI).Code.Text) : strAut = "" : strYr = "" : strTitle = ""
              ' Get the author and the year
              If (Not GetRefDetails(strText, strAut, strYr, strTitle)) Then
                ' Do something
              End If
              ' Find out where this is
              rngThis = .Item(intI).Result
              ' Get the page number
              intPage = rngThis.Information(Word.WdInformation.wdActiveEndPageNumber)
              ' strSect = rngThis.Information(Word.WdInformation.wdActiveEndSectionNumber)
              strLoc = docThis.Bookmarks("\HeadingLevel").Range.Paragraphs(1).Range.ListFormat.ListString
              ' ================== DEBUGGING =====================
              'Debug.Print("Location = " & rngThis.Start & ":" & rngThis.End & " at page=" & _
              '            intPage & " section=" & strSect & " heading=" & strLoc)
              ' ==================================================
              ' Show what we have to the user
              Status("Found p." & intPage & "(s=" & strLoc & ")" & vbCrLf & strText & vbCrLf)
              ' Process this information in the table
              objBack.Add("<tr><td>" & strLoc & "</td><td>" & intPage & "</td><td>" & strAut & "</td><td>" & _
                          strYr & "</td><td>" & strTitle & "</td></tr>")
          End Select
        Next intI
      End With
      ' Finish the HTML 
      objBack.Add("</table></body></html>")
      strBack = objBack.Text
      ' File this information
      IO.File.WriteAllText(strFile, strBack)
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modWord/DissertationList error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' -----------------------------------------------------------------------------------------------
  ' Name: GetRefDetails
  ' Goal: Find the Author and the Year in the ENDNOTE string
  ' History:
  ' 25-11-2012  ERK Created
  ' -----------------------------------------------------------------------------------------------
  Private Function GetRefDetails(ByVal strIn As String, ByRef strAuthor As String, _
                                ByRef strYear As String, ByRef strTitle As String) As Boolean
    Dim pdxDoc As New Xml.XmlDocument ' XML document presentation
    Dim ndxThis As Xml.XmlNode        ' Working node

    Try
      ' Initialise
      strAuthor = "" : strYear = ""
      ' Get the correct part of the input string
      If (InStr(strIn, "<EndNote") = 0) Then
        ' Cannot get the date anyway
        Return True
      End If
      strIn = Mid(strIn, InStr(strIn, "<EndNote"))
      ' Load the document from the string
      pdxDoc.LoadXml(strIn)
      ' Find the author node
      ndxThis = pdxDoc.SelectSingleNode("//Author")
      If (ndxThis Is Nothing) Then ndxThis = pdxDoc.SelectSingleNode("//author")
      If (Not ndxThis Is Nothing) Then
        strAuthor = ndxThis.InnerText
      End If
      ' Find the year node
      ndxThis = pdxDoc.SelectSingleNode("//Year")
      If (Not ndxThis Is Nothing) Then
        strYear = ndxThis.InnerText
      End If
      ' Find the title
      ndxThis = pdxDoc.SelectSingleNode("//Title")
      If (ndxThis Is Nothing) Then ndxThis = pdxDoc.SelectSingleNode("//title")
      If (Not ndxThis Is Nothing) Then
        strTitle = ndxThis.InnerText
      End If
      ' Return positively
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modWord/GetRefDetails error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
End Module
