Public Class frmFeatAdd
  Private loc_strFeatName As String = ""    ' Name for the feature
  Private loc_strFeatValue As String = ""   ' Default value for the feature
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  FeatureName
  ' Goal :  Make the definition name available
  ' History:
  ' 02-10-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public ReadOnly Property FeatureName() As String
    Get
      FeatureName = loc_strFeatName
    End Get
  End Property
  Public ReadOnly Property FeatureValue() As String
    Get
      Return loc_strFeatValue
    End Get
  End Property

  Public Sub SetOwner(ByRef frmThis As Form)
    Try
      Me.Owner = frmThis
    Catch ex As Exception

    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  cmdOk_Click
  ' Goal :  Close the window and make the name for the query available
  ' History:
  ' 02-10-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    ' Store the name locally
    loc_strFeatName = Trim(Me.tbName.Text)
    loc_strFeatValue = Trim(Me.tbValue.Text)
    ' Give the positive dialog result
    Me.DialogResult = Windows.Forms.DialogResult.OK
    ' Close this form
    Me.Close()
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  cmdCancel_Click
  ' Goal :  Say that the operation was canceled anc close the window
  ' History:
  ' 02-10-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    ' Give the negative dialog result
    Me.DialogResult = Windows.Forms.DialogResult.Cancel
    ' Close this form
    Me.Close()
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  frmResTextId_Load
  ' Goal :  Set the owner of the form, so that it is shown in the center of the owner
  ' History:
  ' 02-10-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub frmResTextId_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Start timer
    Me.Timer1.Enabled = True
  End Sub

  Private Sub tbName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbName.TextChanged
    ' Make sure feature does not contain spaces
    If (InStr(Me.tbName.Text, " ") > 0) Then
      Me.tbName.Text = Me.tbName.Text.Replace(" ", "")
    End If
  End Sub

  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Try
      ' Switch off timer
      Me.Timer1.Enabled = False
      ' Set the owner
      Me.Owner = frmMain
      ' Set tooltip text
      With Me.ToolTip1
        .SetToolTip(Me.tbName, "The name of the feature you would like to add")
      End With
      ' Clear values
      Me.tbName.Text = "" : Me.tbValue.Text = "-"
      ' Set focus correctly
      Me.tbName.Focus()
    Catch ex As Exception

    End Try
  End Sub

  Private Sub frmFeatAdd_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
    ' Has it become visible
    If (Me.Visible) Then
      Me.Timer1.Enabled = True
    End If
  End Sub
End Class