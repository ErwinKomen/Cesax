Option Strict On
Option Explicit On
Option Compare Binary
Imports System.Windows.Forms

Public Class frmRef
  Private loc_strType As String = ""    ' The type of reference selected
  ' ------------------------------------------------------------------------------------
  ' Name:   RefType
  ' Goal:   Give the reference type selected
  ' History:
  ' 02-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public ReadOnly Property RefType() As String
    Get
      RefType = loc_strType
    End Get
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   OK_Button_Click
  ' Goal:   Act on OK button
  ' History:
  ' 02-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    ' Set the result of this dialog
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    ' Close the form
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Cancel_Button_Click
  ' Goal:   Act on CANCEL button
  ' History:
  ' 02-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
    ' Set the result of this dialog
    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    ' Close the form
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   frmRef_Load
  ' Goal:   Prepare the form for showing necessary information
  ' History:
  ' 02-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmRef_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Enable timer, which triggers initialisation
    Me.Timer1.Enabled = True
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   DoRefTypes
  ' Goal:   Fill the combobox with the appropriate Coreference types
  ' History:
  ' 02-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function DoRefTypes() As Boolean
    Dim intI As Integer       ' Counter
    Dim strRefType As String  ' Reftype in the table
    Dim bSrcOnly As Boolean   ' Whether source only is selected
    Dim bIsOneArg As Boolean  ' Whether the current reftype is a one-argument one

    ' Fill the combobox
    With Me.lboRefType
      ' First clear old information
      .Items.Clear()
      ' Check whether we need source only
      bSrcOnly = (CorefCount() = 1)
      ' Add the appropriate reference types
      With tdlSettings.Tables("RefType")
        ' Visit all Coreference types from the list
        For intI = 0 To .Rows.Count - 1
          ' Get this reftype
          strRefType = .Rows(intI).Item("Name").ToString
          bIsOneArg = DoLike(strRefType, strRefOneArg)
          ' Should this one be included?
          If (bSrcOnly AndAlso bIsOneArg) OrElse _
             ((Not bSrcOnly) AndAlso (Not bIsOneArg)) Then
            ' Add the name for this Coreference type
            Me.lboRefType.Items.Add(strRefType)
          End If
        Next intI
      End With
      ' How many are there now?
      If (.Items.Count = 0) Then
        ' Return failure
        Return False
      Else
        ' Select appropriate default one
        .SelectedIndex = 0
      End If
    End With
    ' Return success
    DoRefTypes = True
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Trigger initialisation to be done
  ' History:
  ' 02-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    ' Disable timer
    Me.Timer1.Enabled = False
    ' Do the initialisation
    If (Not DoRefInit()) Then
      ' Give it a second chance
      Me.Timer1.Enabled = True
    End If
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   DoRefInit
  ' Goal:   Do any initialisation needed
  ' History:
  ' 04-02-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function DoRefInit() As Boolean
    Dim strText As String = ""  ' Text for tbDst

    ' Fill the combobox (again)
    If (Not DoRefTypes()) Then
      ' Return failure
      Return False
    End If
    ' Set cursor to appropriate place
    Me.lboRefType.Focus()
    ' Show refinit is done
    Status("Select a reference type")
    ' Return success
    DoRefInit = True
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   lboRefType_DoubleClick
  ' Goal:   Return to the caller
  ' History:
  ' 05-02-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub lboRefType_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lboRefType.DoubleClick
    ' Set the result of this dialog
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    ' Close the form
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   lboRefType_SelectedIndexChanged
  ' Goal:   Fix the reference type selected
  ' History:
  ' 02-01-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub lboRefType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lboRefType.SelectedIndexChanged
    ' Adjust what reference type is selected
    loc_strType = Me.lboRefType.Text
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   frmRef_Shown
  ' Goal:   Make sure that the form is initialized and everything shown correctly
  ' History:
  ' 23-06-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmRef_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
    ' Set the timer
    Me.Timer1.Enabled = True
  End Sub
End Class
