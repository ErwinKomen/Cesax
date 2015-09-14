Imports System.Windows.Forms

Public Class frmSyntax
  ' ====================================== LOCAL VARIABLES =================================================
  Private bInit As Boolean = False      ' Initialisation flag
  Private loc_intUserId As Integer = -1 ' Selected Id
  ' ========================================================================================================
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   UserId
  ' Goal:   Return the selected User Id
  ' History:
  ' 31-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public ReadOnly Property UserId() As Integer
    Get
      Return loc_intUserId
    End Get
  End Property
  Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  Private Sub frmSyntax_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Trigger timer for initialisation
    Me.Timer1.Enabled = True
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Trigger initialisation
  ' History:
  ' 31-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Try
      ' Disable timer
      Me.Timer1.Enabled = False
      ' Initialise this form
      DoInit()
      ' Show we are ready
      Status("Ready")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSyntax/Timer error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   DoInit
  ' Goal:   Perform initialisation: setup the DGV
  ' History:
  ' 31-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub DoInit()
    Try
      ' Validate
      If (bInit) Then Exit Sub
      ' Do away with possible previous handlers
      DgvClear(objSynEd)    ' Syntax Editor
      ' Initialise the syntax editor (SynED)
      InitSyntaxEditor()
      ' Ready!
      bInit = True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSyntax/Doinit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   InitSyntaxEditor
  ' Goal:   Initialise the editor for Coref Types
  ' History:
  ' 23-09-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitSyntaxEditor() As Boolean
    Try
      ' Initialise the Query DGV handle
      objSynEd = New DgvHandle
      With objSynEd
        .Init(Me, tdlSyntax, "User", "UserId", "UserId", "Label;Children", "", _
              "", "", Me.dgvUser, Nothing)
        '.BindControl(Me.tbCorefName, "Name", "textbox")
        '.BindControl(Me.tbCorefDescr, "Descr", "richtext")
        '.BindControl(Me.cboColor, "Color", "combo")
        ' Set the parent table for the [CorefType] editor
        .ParentTable = "UserList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitSyntaxEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   dgvUser_SelectionChanged
  ' Goal:   Show the currently selected parse in more detail
  ' History:
  ' 23-09-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub dgvUser_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvUser.SelectionChanged
    Dim dtrFound() As DataRow     ' Result of SELECT
    Dim intSelId As Integer = -1  ' Id of selected one

    Try
      ' Get the selected id
      intSelId = objSynEd.SelectedId
      If (intSelId < 0) Then Exit Sub
      ' We have a valid selection!! Find it...
      dtrFound = tdlSyntax.Tables("User").Select("UserId = " & intSelId)
      If (dtrFound.Length > 0) Then
        ' Get the HTML code for this line
        Me.wbSelected.DocumentText = dtrFound(0).Item("html").ToString
        ' Show user that this one is chosen
        Status("You have chosen: " & dtrFound(0).Item("Label").ToString)
        loc_intUserId = intSelId
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSyntax/SelectionChanged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Class
