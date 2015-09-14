Public Class frmGetResult
  Private strResTextId As String = ""   ' The TextId of the result
  Private strResText As String = ""     ' The actual text of the result
  Private strResPrec As String = ""     ' What precedes the text
  Private strResFoll As String = ""     ' What follows the text
  Private strResPsd As String = ""      ' The syntactic layout of the result in PSD form
  Private strResLoc As String = ""      ' 
  Private strResFile As String = ""     ' 
  Private strResPeriod As String = ""   ' 
  Private strResForestId As String = "" ' 
  Private strResEtreeId As String = ""  ' 
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  TextId
  ' Goal :  Make the value [TextId] available
  ' History:
  ' 02-10-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public ReadOnly Property ResTextId() As String
    Get
      ResTextId = strResTextId
    End Get
  End Property
  Public ReadOnly Property ResText() As String
    Get
      ResText = strResPrec & " <font color='blue'>" & strResText & "</font> " & strResFoll
    End Get
  End Property
  Public ReadOnly Property ResPsd() As String
    Get
      ResPsd = strResPsd
    End Get
  End Property
  Public ReadOnly Property ResLoc() As String
    Get
      ResLoc = strResLoc
    End Get
  End Property
  Public ReadOnly Property ResFile() As String
    Get
      ResFile = strResFile
    End Get
  End Property
  Public ReadOnly Property ResPeriod() As String
    Get
      ResPeriod = strResPeriod
    End Get
  End Property
  Public ReadOnly Property ResForestId() As String
    Get
      ResForestId = strResForestId
    End Get
  End Property
  Public ReadOnly Property ResEtreeId() As String
    Get
      ResEtreeId = strResEtreeId
    End Get
  End Property
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  cmdOk_Click
  ' Goal :  Close the window and make the name for the query available
  ' History:
  ' 02-10-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    ' Store the name locally
    strResTextId = Trim(Me.tbResTextId.Text)
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
  ' Name :  frmElementName_Load
  ' Goal :  Set the owner of the form, so that it is shown in the center of the owner
  ' History:
  ' 02-10-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub frmGetResult_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Try
      ' Set the owner
      Me.Owner = frmMain
      ' Set tooltip text
      With Me.ToolTip1
        .SetToolTip(Me.tbResTextId, "The name of the element you would like to add")
      End With
      ' Set default values
      Me.tbResPeriod.Text = "C3"
      Me.tbResPsd.Text = "(IP-MAT " & vbCrLf & ")" & vbCrLf
      Me.tbResTextId.Text = "(name)"
      Me.tbResForestId.Text = "1"
      Me.tbResEtreeId.Text = "1"
      Me.tbResLoc.Text = "1"
    Catch ex As Exception
      MsgBox("frmGetResult/Load error: " & ex.Message & vbCrLf)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  tbResTextId_TextChanged
  ' Goal :  Set the local values
  ' History:
  ' 27-09-2011  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub tbResTextId_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbResTextId.TextChanged
    strResTextId = Me.tbResTextId.Text
  End Sub
  Private Sub tbResLoc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbResLoc.TextChanged
    strResLoc = Me.tbResLoc.Text
  End Sub
  Private Sub tbResFile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbResFile.TextChanged
    strResFile = Me.tbResFile.Text
  End Sub
  Private Sub tbResPeriod_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbResPeriod.TextChanged
    strResPeriod = Me.tbResPeriod.Text
  End Sub
  Private Sub tbResForestId_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbResForestId.TextChanged
    strResForestId = Me.tbResForestId.Text
  End Sub
  Private Sub tbResEtreeId_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbResEtreeId.TextChanged
    strResEtreeId = Me.tbResEtreeId.Text
  End Sub
  Private Sub tbResPrec_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbResPrec.TextChanged
    strResPrec = Me.tbResPrec.Text
  End Sub
  Private Sub tbResText_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbResText.TextChanged
    strResText = Me.tbResText.Text
  End Sub
  Private Sub tbResFoll_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbResFoll.TextChanged
    strResFoll = Me.tbResFoll.Text
  End Sub
  Private Sub tbResPsd_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbResPsd.TextChanged
    strResPsd = Me.tbResPsd.Text
  End Sub
End Class