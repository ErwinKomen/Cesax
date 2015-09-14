Public Class frmGetName
  Private strElementName As String = ""   ' Name for the query
  Private strDefaultValue As String = ""  ' Default value for the property
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  ResTextId
  ' Goal :  Make the definition name available
  ' History:
  ' 02-10-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public ReadOnly Property ElementName() As String
    Get
      ElementName = strElementName
    End Get
  End Property

  Public WriteOnly Property DefaultValue() As String
    Set(ByVal value As String)
      strDefaultValue = value
    End Set
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
    strElementName = Trim(Me.tbName.Text)
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
    ' Set the owner
    Me.Owner = frmMain
    ' Set tooltip text
    With Me.ToolTip1
      .SetToolTip(Me.tbName, "The name of the element you would like to add")
    End With
  End Sub

  Private Sub frmGetName_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.VisibleChanged
    Try
      If (Me.Visible = True) Then
        ' We have just become visible!
        Me.tbName.Text = strDefaultValue
      End If
    Catch ex As Exception

    End Try
  End Sub
End Class