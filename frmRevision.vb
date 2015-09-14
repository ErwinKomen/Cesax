Imports System.Windows.Forms

Public Class frmRevision
  Private loc_strWho As String = ""     ' The person who made the comment
  Private loc_strWhen As String = ""    ' The date
  Private loc_strComment As String = "" ' The comment for the revision itself
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Who
  ' Goal :  Make the "who" value available to the user
  ' History:
  ' 05-07-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public ReadOnly Property Who() As String
    Get
      Return loc_strWho
    End Get
  End Property
  Public ReadOnly Property WhenMade() As String
    Get
      Return loc_strWhen
    End Get
  End Property
  Public ReadOnly Property Comment() As String
    Get
      Return loc_strComment
    End Get
  End Property
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  OK_Button_Click
  ' Goal :  Pass back a positive dialog result
  ' History:
  ' 05-07-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Cancel_Button_Click
  ' Goal :  Pass back a negative dialog result
  ' History:
  ' 05-07-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  tbWho_TextChanged
  ' Goal :  Put changes into the local variables
  ' History:
  ' 05-07-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Private Sub tbWho_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbWho.TextChanged
    loc_strWho = Me.tbWho.Text
  End Sub
  Private Sub dtWhen_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtWhen.ValueChanged
    Try
      ' Get date/time from the datestuff
      Me.tbWhen.Text = Format(Me.dtWhen.Value, "g")
      loc_strWhen = Me.tbWhen.Text
    Catch ex As Exception
      ' Show error
      HandleErr("frmRevision/When error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub tbComment_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbComment.TextChanged
    loc_strComment = Me.tbComment.Text
  End Sub

  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  tbWho_TextChanged
  ' Goal :  Put changes into the local variables
  ' History:
  ' 05-07-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Private Sub frmRevision_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Start timer
    Me.Timer1.Enabled = True
  End Sub

  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Timer1_Tick
  ' Goal :  Perform other initialisations
  ' History:
  ' 05-07-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    ' Switch off timer
    Me.Timer1.Enabled = False
    ' Make an initial guess as to who the user is...
    Me.tbWho.Text = strUserName
    ' Make initial guess as to the date
    Me.dtWhen.Value = Now
    Me.tbWhen.Text = Format(Now, "g")
    ' Set the filename
    Me.tbFile.Text = IO.Path.GetFileNameWithoutExtension(strCurrentFile)
    ' Put the focus on the comment itself
    Me.tbComment.Focus()
    ' Set the selection
    Me.tbComment.SelectionStart = 0
    Me.tbComment.SelectionLength = Me.tbComment.TextLength
    ' Show we are ready
    Status("Ready")
  End Sub
End Class
