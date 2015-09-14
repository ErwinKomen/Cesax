Imports System.Windows.Forms

Public Class frmDirs
  ' ====================================== LOCAL VARIABLES ======================================================
  Private strSrcDir As String = ""  ' Source directory
  Private strDstDir As String = ""  ' Destination directory
  ' =============================================================================================================

  '---------------------------------------------------------------------------------------------------------------
  ' Name:       SrcDir()
  ' Goal:       Set or get the source directory
  ' History:
  ' 26-05-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------------
  Public Property SrcDir() As String
    Get
      Return strSrcDir
    End Get
    Set(ByVal value As String)
      strSrcDir = value
      ' Show the value on the form
      Me.tbSrcDir.Text = strSrcDir
    End Set
  End Property
  '---------------------------------------------------------------------------------------------------------------
  ' Name:       DstDir()
  ' Goal:       Set or get the destination directory
  ' History:
  ' 26-05-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------------
  Public Property DstDir() As String
    Get
      Return strDstDir
    End Get
    Set(ByVal value As String)
      strDstDir = value
      ' Show the value on the form
      Me.tbDstDir.Text = strDstDir
    End Set
  End Property
  '---------------------------------------------------------------------------------------------------------------
  ' Name:       AskOneDir()
  ' Goal:       Make sure only one directory is asked for
  ' History:
  ' 10-06-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------------
  Public Sub AskOneDir()
    Try
      Me.tbDstDir.Visible = False
      Me.cmdDstDir.Visible = False
      Me.lbDst.Visible = False
      Me.Text = "Specify the directory you would like Cesax to use"
    Catch ex As Exception
      ' Inform user
      HandleErr("frmDirs/AskOneDir error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------------------------------------------------------
  ' Name:       OK_Button_Click()
  ' Goal:       Close with a positive outcome
  ' History:
  ' 26-05-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------------
  Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  '---------------------------------------------------------------------------------------------------------------
  ' Name:       Cancel_Button_Click()
  ' Goal:       Close with a negative outcome
  ' History:
  ' 26-05-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------------
  Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  '---------------------------------------------------------------------------------------------------------------
  ' Name:       cmdSrcDir_Click()
  ' Goal:       Let user select a source directory
  ' History:
  ' 26-05-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------------
  Private Sub frmDirs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Set my owner
    Me.Owner = frmMain
    ' SPecify the text we show
    Me.Text = "Specify the directories you would like Cesax to use"
    ' Start up initialisation
    Me.Timer1.Enabled = True
  End Sub

  '---------------------------------------------------------------------------------------------------------------
  ' Name:       cmdSrcDir_Click()
  ' Goal:       Let user select a source directory
  ' History:
  ' 26-05-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------------
  Private Sub cmdSrcDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSrcDir.Click
    Try
      ' Try get directory
      If (GetDirName(Me.FolderBrowserDialog1, strSrcDir, _
                     "Select the directory where the source *.psdx files are", strWorkDir)) Then
        ' Replace what is shown
        Me.tbSrcDir.Text = strSrcDir
      End If
    Catch ex As Exception
      ' Inform user
      HandleErr("frmDirs/cmdSrcDir error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '---------------------------------------------------------------------------------------------------------------
  ' Name:       cmdDstDir_Click()
  ' Goal:       Let user select a destination directory
  ' History:
  ' 26-05-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------------
  Private Sub cmdDstDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDstDir.Click
    Dim strStartDir As String = "" ' Directory where to start from

    Try
      ' Determine the initial directory
      strDstDir = IIf(strSrcDir = "", strWorkDir, strSrcDir)
      ' Try get directory
      If (GetDirName(Me.FolderBrowserDialog1, strDstDir, _
                     "Select the directory where the destination *.psdx files are", strStartDir)) Then
        ' Replace what is shown
        Me.tbDstDir.Text = strDstDir
      End If
    Catch ex As Exception
      ' Inform user
      HandleErr("frmDirs/cmdDstDir error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '---------------------------------------------------------------------------------------------------------------
  ' Name:       Timer1_Tick()
  ' Goal:       Perform actual initialisations
  ' History:
  ' 26-05-2010  ERK Created
  '---------------------------------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    ' Switch off the timer
    Me.Timer1.Enabled = False
    ' Initialise matters
    If (strSrcDir = "") Then
      ' Set initial source and destination directories
      strSrcDir = strWorkDir
    End If
    If (strDstDir = "") Then
      strDstDir = strWorkDir
    End If
    Me.tbSrcDir.Text = strSrcDir
    Me.tbDstDir.Text = strDstDir
    ' Show we are ready
    Status("Ready")
  End Sub
End Class
