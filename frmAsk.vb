Imports System.Windows.Forms

Public Class frmAsk
  Private Const MAX_ANT As Integer = 20   ' Maximum number of antecedents to consider
  Private loc_intDstId As Integer = -1    ' The ID of the selected target
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   Source
  ' Goal:   Set the source NP's text
  ' History:
  ' 15-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public WriteOnly Property Source() As String
    Set(ByVal value As String)
      ' Fill in the source text
      Me.tbSource.Text = value
    End Set
  End Property
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   Message
  ' Goal:   Set the message for the user
  ' History:
  ' 15-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public WriteOnly Property Message() As String
    Set(ByVal value As String)
      ' Set the message for the user
      Me.tbProblem.Text = value
    End Set
  End Property
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   SetTargets
  ' Goal:   Set the listbox from the source and the antecedent stacks
  ' History:
  ' 15-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Sub SetTargets(ByRef tblSrcStack As DataTable, ByRef tblAntStack As DataTable)
    Dim intI As Integer     ' Counter
    Dim dtrThis As DataRow  ' One element

    Try
      ' Validate
      If (tblSrcStack Is Nothing) Then Exit Sub
      ' Clear the listbox
      Me.lboTarget.Items.Clear()
      ' Read the source stack one at a time backwards...
      For intI = tblSrcStack.Rows.Count - 1 To 0 Step -1
        ' Get this element
        dtrThis = tblSrcStack.Rows(intI)
        ' Put the source stack element in the listbox, including its ID
        Me.lboTarget.Items.Add(dtrThis.Item("Id") & ": " & dtrThis.Item("Vern"))
      Next intI
      ' Are there antecedents?
      If (tblAntStack Is Nothing) Then Exit Sub
      ' Put the antecedent stack's elements in there (up to a maximum)
      For intI = tblAntStack.Rows.Count - 1 To 0 Step -1
        ' Get this element
        dtrThis = tblAntStack.Rows(intI)
        ' Put the source stack element in the listbox, including its ID
        Me.lboTarget.Items.Add(dtrThis.Item("Id") & ": " & dtrThis.Item("Vern"))
        ' Check if we are beyond our limit
        If (tblAntStack.Rows.Count - MAX_ANT < intI) Then Exit Sub
      Next intI
    Catch ex As Exception
      ' Show error
      MsgBox("frmAsk/SetTargets error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   GetTarget
  ' Goal:   Get the ID of the selected constituent
  ' History:
  ' 15-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function GetTarget() As Integer
    ' Return the destiination node's id
    GetTarget = loc_intDstId
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   OK_Button_Click
  ' Goal:   Return with a positive result
  ' History:
  ' 15-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    Dim strSelected As String ' What has been selected
    Dim arThis() As String    ' Array

    Try
      ' Set the correct result for the dialogue
      Me.DialogResult = System.Windows.Forms.DialogResult.OK
      ' Read the target's ID
      strSelected = Me.lboTarget.SelectedItem
      ' Do we have anything?
      If (strSelected <> "") Then
        arThis = Split(strSelected, ":")
        ' Get the first number
        If (IsNumeric(arThis(0))) Then
          loc_intDstId = arThis(0)
        End If
      End If
      ' Close the window
      Me.Close()
    Catch ex As Exception
      ' Show error
      MsgBox("frmAsk/OK_Button_Click error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   Cancel_Button_Click
  ' Goal:   Return with a negative result
  ' History:
  ' 15-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
    ' Set the correct result for the dialogue
    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    ' Close the window
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   frmAsk_Load
  ' Goal:   Trigger initialisation
  ' History:
  ' 15-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub frmAsk_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Set the owner form
    Me.Owner = frmMain
    ' Trigger initialisation
    Me.Timer1.Enabled = True
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Do whatever initialisations are possible at this point
  ' History:
  ' 15-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Try
      ' Switch off the timer
      Me.Timer1.Enabled = False
      ' Perform the actual initialisations needed
    Catch ex As Exception
      ' Show error
      MsgBox("frmAsk/Timer1_Tick error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Class
