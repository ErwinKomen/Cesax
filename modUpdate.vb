Imports System.Deployment.Application
Module modUpdate
  ' =========================================== LOCAL VARIABLES =============================================
  Private sizeOfUpdate As Long = 0
  Dim WithEvents ADUpdateAsync As ApplicationDeployment
  Private bMustUpdate As Boolean = False    ' WHether updating MUST be done if anything is available
  ' =========================================================================================================
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  UpdateApplication
  ' Goal :  Try to find an update for this application...
  ' History:
  ' 22-06-2012  ERK Created
  ' 25-06-2012  ERK Added [MandatoryOnly] option
  ' ---------------------------------------------------------------------------------------------------------
  Public Sub UpdateApplication(ByVal bMandatoryOnly As Boolean)
    Try
      ' Check how we are deployed
      If (ApplicationDeployment.IsNetworkDeployed) Then
        ' Get an asynchronous application deployment object
        ADUpdateAsync = ApplicationDeployment.CurrentDeployment
        If (ADUpdateAsync Is Nothing) Then
          Status("UpdateApplication: cannot find a current deployment")
        End If
        ' Pass on the manditorinous
        bMustUpdate = (Not bMandatoryOnly)
        ' Look for updates asynchronously
        ADUpdateAsync.CheckForUpdateAsync()
      End If
    Catch ex As Exception
      ' Show the error
      HandleErr("modUpdate/UpdateApplication error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  SoftwareMessage
  ' Goal :  Check presence of and display content of a software message for this program
  ' History:
  ' 25-03-2014  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Sub SoftwareMessage()
    Dim strFile As String = IO.Path.GetDirectoryName(Application.ExecutablePath) & "\msg.txt"
    Dim strText As String = ""
    Dim bEd As Boolean    ' Temporary

    Try
      If (IO.File.Exists(strFile)) Then
        ' Get the message
        strText = Trim(IO.File.ReadAllText(strFile))
        ' Add the message to frmMain[tbComments]
        With frmMain.tbGenLog
          bEd = .Enabled
          .Enabled = True
          .AppendText("This version: " & strText & vbCrLf)
          .Enabled = bEd
        End With
      End If
    Catch ex As Exception
      ' Show the error
      HandleErr("modUpdate/SoftwareMessage error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  LookForUpdMsg
  ' Goal :  Check presence of and display content of a software message for this program
  ' History:
  ' 25-03-2014  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Sub LookForUpdMsg(ByVal strUrl As String)
    Dim strFile As String = GetDocDir() & "\msg.txt"
    Dim bEn As Boolean        ' Situation of "enabledness"

    Try
      If (DownloadFile(strUrl, strFile, False)) Then
        With frmMain.tbGenLog
          ' Get enabledness
          bEn = .Enabled
          ' Make sure we are enabled
          .Enabled = True
          ' Put the text there
          .AppendText("Software: " & IO.File.ReadAllText(strFile) & vbCrLf)
          ' Restore enabledness state
          .Enabled = bEn
          'MsgBox(strErr)
        End With
      End If
    Catch ex As Exception
      ' Show the error
      HandleErr("modUpdate/LookForUpdMsg error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  ADUpdateAsync_CheckForUpdateProgressChanged
  ' Goal :  Check if the update progress has changed
  ' History:
  ' 22-06-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub ADUpdateAsync_CheckForUpdateProgressChanged(ByVal sender As Object, _
         ByVal e As DeploymentProgressChangedEventArgs) Handles ADUpdateAsync.CheckForUpdateProgressChanged
    Dim intPtc As Integer   ' Percentage

    Try
      ' Validate
      If (e Is Nothing) Then Exit Sub
      If (e.BytesTotal > 0) Then
        ' Are we there already?
        If (e.BytesTotal = e.BytesCompleted) Then
          ' show we are ready
          Status("Ready")
        Else
          ' Calculate the percentage
          intPtc = (100 * e.BytesCompleted) \ e.BytesTotal
          ' Show where we are in the downloading process
          Status("Checking for updates " & intPtc & "%", intPtc)
        End If
      Else
        Status("")
      End If
    Catch ex As Exception
      ' Show the error
      HandleErr("modUpdate/ADUpdateAsync_CheckForUpdateProgressChanged error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  ADUpdateAsync_CheckForUpdateCompleted
  ' Goal :  Check if the update is completed
  ' History:
  ' 22-06-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub ADUpdateAsync_CheckForUpdateCompleted(ByVal sender As Object, _
         ByVal e As CheckForUpdateCompletedEventArgs) Handles ADUpdateAsync.CheckForUpdateCompleted
    Dim strErr As String = "" ' Text of error
    Dim bEn As Boolean        ' Situation of "enabledness"

    Try
      If (e.Error IsNot Nothing) Then
        strErr = "Update Error: Could not check for a possible update." & vbCrLf & _
               "The Reason: " & vbCrLf & e.Error.Message & vbCrLf
        With frmMain.tbGenLog
          ' Get enabledness
          bEn = .Enabled
          ' Make sure we are enabled
          .Enabled = True
          ' Put the text there
          .Text = strErr
          ' Restore enabledness state
          .Enabled = bEn
          'MsgBox(strErr)
        End With
        Return
      Else
        If (e.Cancelled = True) Then
          MsgBox("The update was cancelled.")
        End If
      End If

      ' Ask the user if they would like to update the application now.
      If (e.UpdateAvailable) Then
        sizeOfUpdate = e.UpdateSizeBytes
        ' Is this a required update?
        If (Not e.IsUpdateRequired) Then
          ' Should we bother?
          If (bMustUpdate) Then
            ' Give user the option of updating
            Select Case MsgBox("An update is available. Would you like to update the application now?", _
                                MsgBoxStyle.OkCancel, "Update Available")
              Case MsgBoxResult.Ok
                BeginUpdate()
            End Select
          End If
        Else
          ' This is an update that MUST be accepted
          MsgBox("A mandatory update is available for your application. " & _
                 "We will install the update now, after which we will save all of your in-progress data " & _
                 "and restart your application.")
          BeginUpdate()
        End If
      End If
    Catch ex As Exception
      ' Show the error
      HandleErr("modUpdate/ADUpdateAsync_CheckForUpdateCompleted error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  BeginUpdate
  ' Goal :  Start actually updating
  ' History:
  ' 22-06-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub BeginUpdate()
    Try
      ADUpdateAsync = ApplicationDeployment.CurrentDeployment
      If (ADUpdateAsync Is Nothing) Then
        Status("BeginUpdate: there is no update available")
      End If
      ADUpdateAsync.UpdateAsync()
    Catch ex As Exception
      ' Show the error
      HandleErr("modUpdate/BeginUpdate error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  ADUpdateAsync_UpdateProgressChanged
  ' Goal :  The update is progressing...
  ' History:
  ' 22-06-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub ADUpdateAsync_UpdateProgressChanged(ByVal sender As Object, _
           ByVal e As DeploymentProgressChangedEventArgs) Handles ADUpdateAsync.UpdateProgressChanged
    Dim intPtc As Integer   ' Percentage

    Try
      ' Validate
      If (e.BytesTotal > 0) Then
        ' Are we there already?
        If (e.BytesTotal = e.BytesCompleted) Then
          ' show we are ready
          Status("Ready")
        Else
          ' Calculate the percentage
          intPtc = (100 * e.BytesCompleted) \ e.BytesTotal
          ' Show where we are in the downloading process
          Status("Downloading " & intPtc & "%", intPtc)
        End If
      Else
        Status("")
      End If

      'Dim progressText As String = String.Format("{0:D}K out of {1:D}K downloaded - {2:D}% complete", e.BytesCompleted / 1024, e.BytesTotal / 1024, e.ProgressPercentage)

      'Status(progressText)
    Catch ex As Exception
      ' Show the error
      HandleErr("modUpdate/ADUpdateAsync_UpdateProgressChanged error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  ADUpdateAsync_UpdateCompleted
  ' Goal :  Is reached when the update is completed
  ' History:
  ' 22-06-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub ADUpdateAsync_UpdateCompleted(ByVal sender As Object, _
           ByVal e As System.ComponentModel.AsyncCompletedEventArgs) Handles ADUpdateAsync.UpdateCompleted
    Try
      Dim intCode As Integer = 0    ' Return code

      If (e.Cancelled) Then
        MsgBox("The update of the application's latest version was cancelled.")
        Exit Sub
      Else
        If (e.Error IsNot Nothing) Then
          MsgBox("ERROR: Could not install the latest version of the application. Reason: " & vbCrLf & _
                 e.Error.Message & vbCrLf & "Please report this error to the system administrator.")
          Exit Sub
        End If
      End If
      intCode = MsgBox("The application has been updated. Restart? " & vbCrLf & _
                         "(If you do not restart now, the new version will not take effect until " & _
                         "after you quit and launch the application again.)", _
                         MsgBoxStyle.OkCancel, "Restart Application")
      Select Case intCode
        Case MsgBoxResult.Ok
          Application.Restart()
        Case Else
          MsgBox("Code = " & intCode)
      End Select
      'Dim dr As DialogResult = MessageBox.Show("The application has been updated. Restart? (If you do not restart now, the new version will not take effect until after you quit and launch the application again.)", "Restart Application", MessageBoxButtons.OKCancel)
      'If (dr = System.Windows.Forms.DialogResult.OK) Then
      '  Application.Restart()
      'End If
    Catch ex As Exception
      ' Show the error
      HandleErr("modUpdate/ADUpdateAsync_UpdateCompleted error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
    End Try
  End Sub
End Module
