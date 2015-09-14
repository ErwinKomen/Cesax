Public Class frmReport
  Dim strReportName As String = "MissingPronouns"   ' Default report name
  Dim strReport As String = "(Nothing loaded)"      ' Report to be displayed
  ' ------------------------------------------------------------------------------------
  ' Name:   Report
  ' Goal:   Allow user to set the text of the report
  ' History:
  ' 28-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public WriteOnly Property Report() As String
    Set(ByVal strHtml As String)
      ' Show what we are doing
      Status("Loading text...")
      ' Write the report
      strReport = strHtml
      ' Show we are ready
      Status("Ready")
    End Set
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   ReportName
  ' Goal:   Allow user to specify the name of this report
  ' History:
  ' 28-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public WriteOnly Property ReportName() As String
    Set(ByVal value As String)
      strReportName = value
    End Set
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   frmReport_Load
  ' Goal:   Trigger initialisation
  ' History:
  ' 28-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Set my owner
    Me.Owner = frmMain
    ' Trigger initialisation
    Me.Timer1.Enabled = True
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Start initialisation
  ' History:
  ' 28-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    ' Switch off timer
    Me.Timer1.Enabled = False
    ' Show what we are doing
    Status("Loading ...")
    ' Start initialisation
    Me.wbReport.DocumentText = strReport
    ' Show all is loaded
    Status("Ready")
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuFileSave_Click
  ' Goal:   Allow user to save this report
  ' History:
  ' 28-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuFileSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileSave.Click
    ' Let user save it
    TrySaveReport()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   mnuFileClose_Click
  ' Goal:   Close the report
  ' History:
  ' 28-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub mnuFileClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileClose.Click
    ' Close the form
    Me.Close()
  End Sub
  Private Sub TrySaveReport()
    Dim strFile As String   ' File to save to

    Try
      ' Get the filename from the user
      With Me.SaveFileDialog1
        ' The initial directory is the one we already know
        .InitialDirectory = strWorkDir
        ' The name is derived from the name of this project
        strFile = strReportName & ".psdx"
        ' Set the default extention
        .DefaultExt = "html"
        ' Set the default filter
        .Filter = "Report HTML file|*.html"
        ' Assign default file name to the FileSave dialog
        .FileName = strFile
        ' Show the actual dialog to the user
        Select Case .ShowDialog()
          Case Windows.Forms.DialogResult.OK, Windows.Forms.DialogResult.Yes
            ' Get the filename that the user has finally selected
            strFile = .FileName
          Case Else
            ' Aborted, so exit
            Status("File is not saved")
            Exit Sub
        End Select
      End With
      ' If we come here, then the user has selected a filename properly - do the saving...
      IO.File.WriteAllText(strFile, Me.wbReport.DocumentText)
      ' Show success
      Status("File saved to " & strFile)
    Catch ex As Exception
      ' Show error
      HandleErr("frmReport/TrySaveReport error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Class