Option Strict On
Option Explicit On
Option Compare Binary
Public NotInheritable Class frmAbout
  Private strVersionInfo As String = ""
  Private strVersionFile As String = "VersionInfo.txt"
  ' ------------------------------------------------------------------------------------
  ' Name:   frmAbout_Load
  ' Goal:   Fill in what needs to be filled in before showing myself
  ' History:
  ' 06-01-2009  ERK Automatic creation
  ' ------------------------------------------------------------------------------------
  Private Sub frmAbout_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Set the title of the form.
    Dim ApplicationTitle As String    ' The title of the application
    Dim strVersion As String          ' The version number as a string

    ' Get the proper title of the application
    If My.Application.Info.Title <> "" Then
      ApplicationTitle = My.Application.Info.Title
    Else
      ApplicationTitle = System.IO.Path.GetFileNameWithoutExtension(My.Application.Info.AssemblyName)
    End If
    Me.Text = String.Format("About {0}", ApplicationTitle)
    ' Initialize all of the text displayed on the About Box.
    ' TODO: Customize the application's assembly information in the "Application" pane of the project 
    '    properties dialog (under the "Project" menu).
    If (System.Deployment.Application.ApplicationDeployment.IsNetworkDeployed) Then
      strVersion = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
    Else
      strVersion = String.Format("Version {0}", My.Application.Info.Version.ToString)
    End If
    ' Fill in the values on the appropriate place
    Me.LabelProductName.Text = My.Application.Info.ProductName
    Me.LabelVersion.Text = strVersion
    Me.LabelCopyright.Text = My.Application.Info.Copyright
    Me.LabelCompanyName.Text = My.Application.Info.CompanyName
    ' Retrieve the version information (if existent)
    strVersionFile = IO.Path.GetDirectoryName(Application.ExecutablePath) & "\" & strVersionFile
    If (IO.File.Exists(strVersionFile)) Then
      ' Read the version information
      strVersionInfo = IO.File.ReadAllText(strVersionFile)
    End If
    ' Show the version information
    Me.TextBoxDescription.Text = My.Application.Info.Description & vbCrLf & strVersionInfo & vbCrLf
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   OKButton_Click
  ' Goal:   Close form on OK button clicking
  ' History:
  ' 06-01-2009  ERK Automatic creation
  ' ------------------------------------------------------------------------------------
  Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click
    Me.Close()
  End Sub

End Class
