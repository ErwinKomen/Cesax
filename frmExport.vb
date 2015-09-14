Imports System.Windows.Forms

Public Class frmExport
  Private strType As String = ""    ' The type of export chosen

  ' ------------------------------------------------------------------------------------
  ' Name:   OK_Button_Click
  ' Goal:   User clicked OK button
  ' History:
  ' 10-03-2009  ERK Created shell
  ' ------------------------------------------------------------------------------------
  Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Cancel_Button_Click
  ' Goal:   User canceled the operation
  ' History:
  ' 10-03-2009  ERK Created shell
  ' ------------------------------------------------------------------------------------
  Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   frmExport_Load
  ' Goal:   Indicate owner of this form
  ' History:
  ' 10-03-2009  ERK Created shell
  ' ------------------------------------------------------------------------------------
  Private Sub frmExport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Indicate who my owner is
    Me.Owner = frmMain
    ' Indicate default type
    Me.rbPsd.Checked = True
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   rbPsd_CheckedChanged
  ' Goal:   Show we want PSD nodes output
  ' History:
  ' 10-03-2009  ERK Created shell
  ' ------------------------------------------------------------------------------------
  Private Sub rbPsd_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbPsd.CheckedChanged, rbPsd.CheckedChanged
    ' Adjust value of [strType]
    strType = "PsdOutput"
  End Sub
  Private Sub rbPsdSimple_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbPsdSimple.CheckedChanged
    ' Adjust value of [strType]
    strType = "PsdSimple"
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   rbPde_CheckedChanged
  ' Goal:   Show we want the psdy output with the English back translation
  ' History:
  ' 01-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub rbPde_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbPde.CheckedChanged
    ' Adjust value of [strType]
    strType = "PdeOutput"
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   ExportType
  ' Goal:   Make the chosen exporttype available for the caller
  ' History:
  ' 10-03-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public ReadOnly Property ExportType() As String
    Get
      ExportType = strType
    End Get
  End Property
End Class
