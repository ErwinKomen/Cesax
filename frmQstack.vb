Public Class frmQstack
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   frmQstack_Load
  ' Goal:   Trigger initialisation
  ' History:
  ' 01-11-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub frmQstack_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Me.Timer1.Enabled = True
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick // DoInit
  ' Goal:   handle initialisation
  ' History:
  ' 01-11-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Try
      ' Switch off timer
      Me.Timer1.Enabled = False
      ' Position the window in the lower-right corner of the screen
      Me.Top = My.Computer.Screen.WorkingArea.Height - Me.Height
      Me.Left = My.Computer.Screen.WorkingArea.Width - Me.Width
      ' Do initialisation actions
      If (DoInit()) Then
        ' Show we are ready
        Status("Ready")
      Else
        Status("Could not initialise")
        ' Trigger initialisation again
        Me.Timer1.Enabled = True
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmQstack/Timer error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Function DoInit() As Boolean
    Try
      ' Clear and initialise editors
      DgvClear(objQsEd)
      If (Not InitQstackEditor()) Then Return False
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmQstack/DoInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return faiulre
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitQstackEditor
  ' Goal:   Initialise the Qstack editor
  ' History:
  ' 01-11-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitQstackEditor() As Boolean
    Try
      ' Initialise the dependency editor's DGV handle
      objQsEd = New DgvHandle
      With objQsEd
        .Init(Me, tdlQstack, "entry", "entryId", "entryId ASC", "Number;Name;Type;Base;Relation;Value;Node", "", _
              "", "", Me.dgvQstack, Nothing)
        '.BindControl(Me.tbDepRel, "Rel", "textbox")
        ' Set the parent table for the [Qstack] editor
        ' .ParentTable = "Qstack"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Switch off multiselect
      Me.dgvQstack.MultiSelect = False
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmQstack/InitQstackEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdCancel_Click
  ' Goal:   user canceled
  ' History:
  ' 01-11-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.DialogResult = Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdExport_Click
  ' Goal:   user canceled
  ' History:
  ' 01-11-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExport.Click
    Try
      ' Try to export
      QstackExport()
      ' Close up nicely
      Me.DialogResult = Windows.Forms.DialogResult.OK
      Me.Close()
      ' Show what we have done
      Status("You can now import the query details in CorpusStudio")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmQstack/Export error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdReset_Click
  ' Goal:   Reset the query stack
  ' History:
  ' 01-11-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
    Try
      If (tdlQstack IsNot Nothing) Then
        ClearTable(tdlQstack.Tables("entry"))
        tdlQstack.AcceptChanges()
        ClearQstack()
        Status("Cleared")
      Else
        Status("Nothing to clear")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmQstack/Reset error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdDelete_Click
  ' Goal:   Delete currently selected condition
  ' History:
  ' 01-11-2013  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
    Try
      ' Double check and execute
      TryRemove(objQsEd)
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmQstack/Delete error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  TryRemove
  ' Goal :  Delete the selected element from the dgv in [objThis]
  ' History:
  ' 02-06-2010  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub TryRemove(ByRef objThis As DgvHandle)
    Dim strQ As String = "Would you really like to delete the selected row?"

    Try
      ' Validate
      If (objThis Is Nothing) Then Status("No QS object available") : Exit Sub
      ' First enquire
      If (MsgBox(strQ, MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes) Then
        ' TRy to remove the selected object
        If (Not objThis.Remove) Then
          ' Unsuccesful...
          Status("Unable to delete this element")
        Else
          ' Show status
          Status("Deleted")
        End If
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmQstack/TryRemove error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Class