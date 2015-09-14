Public Class frmHelp
  ' ------------------------------------------------------------------------------------
  ' Name:   Help
  ' Goal:   Set the text of the help message
  ' History:
  ' 21-10-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public WriteOnly Property Help() As String
    Set(ByVal strFile As String)
      Dim strHtml As String   ' The complete text in HTML

      Try
        ' Is this internet or local?
        If (InStr(strFile, "http:") > 0) Then
          ' Internet
          Me.wbHelp.Navigate(strFile)
          ' Show where we are
          Status(strFile)
        Else
          ' Local file -- Does the file exist?
          If (Not IO.File.Exists(strFile)) Then
            ' Warn user
            Status("Unable to open help file: " & strFile)
          Else
            ' Load it
            Status("Opening " & strFile)
            strHtml = IO.File.ReadAllText(strFile, System.Text.Encoding.GetEncoding(1252))
            ' Perform changes
            'strHtml = strHtml.Replace(ChrW(Convert.ToInt32("201c", 16)), "&quot;")
            'strHtml = strHtml.Replace(ChrW(Convert.ToInt32("201d", 16)), "&quot;")
            strHtml = strHtml.Replace(ChrW(147), "&quot;")
            strHtml = strHtml.Replace(ChrW(148), "&quot;")
            ' Show the document
            Me.wbHelp.DocumentText = strHtml
            Status("ok")
          End If
        End If
      Catch ex As Exception
        ' Show error
        HandleErr("frmHelp/Help error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      End Try
    End Set
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Trigger initialisation
  ' History:
  ' 21-10-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    ' Switch off timer
    Me.Timer1.Enabled = False
    ' Start initialisation
    DoInit()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   frmHelp_KeyDown
  ' Goal:   Detect pressing of ESCape
  ' History:
  ' 21-10-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmHelp_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
    If (e.KeyCode = Keys.Escape) Then
      ' Leave...
      Me.Close()
      e.Handled = True
    End If
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   frmHelp_Load
  ' Goal:   Trigger timer
  ' History:
  ' 21-10-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmHelp_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Set owner
    Me.Owner = frmMain
    ' Set timer
    Me.Timer1.Enabled = True
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   DoInit
  ' Goal:   Perform any initialisation needed
  ' History:
  ' 21-10-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub DoInit()
    ' Show we are ready
    Status("Ready")
  End Sub

  Private Sub cmdBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBack.Click
    If (Me.wbHelp.GoBack()) Then
      Status("One step back")
    End If
  End Sub

  Private Sub cmdForward_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdForward.Click
    If (Me.wbHelp.GoForward()) Then
      Status("One step forward")
    End If
  End Sub
End Class