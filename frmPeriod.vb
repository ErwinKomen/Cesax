Imports System.Windows.Forms

Public Class frmPeriod
  ' ================================= LOCAL VARIABLES =========================================================
  Private tblPeriod As DataTable = Nothing      ' Period definitions: from period Id to period name
  Private tblPeriodFile As DataTable = Nothing  ' From filename to period ID
  Private loc_strFileName As String = ""        ' The file we are looking at
  ' ============================================================================================================
  ' ------------------------------------------------------------------------------------
  ' Name:   FileName
  ' Goal:   Set the filename we are looking at
  ' History:
  ' 14-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public WriteOnly Property FileName() As String
    Set(ByVal value As String)
      loc_strFileName = value
      Me.tbFile.Text = loc_strFileName
      Me.tbName.Text = LCase(IO.Path.GetFileNameWithoutExtension(loc_strFileName))
    End Set
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   Period
  ' Goal:   Return the correct period
  ' History:
  ' 14-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public ReadOnly Property Period() As String
    Get
      ' Return the period as has been selected here
      Period = Me.tbPeriod.Text
    End Get
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   OK_Button_Click
  ' Goal:   Return OK
  ' History:
  ' 14-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Cancel_Button_Click
  ' Goal:   Return failure
  ' History:
  ' 14-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   frmPeriod_Load
  ' Goal:   Trigger initialisation
  ' History:
  ' 14-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmPeriod_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Set my owner
    Me.Owner = frmMain
    ' Start the timer
    Me.Timer1.Enabled = True
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Perform all initialisations needed
  ' History:
  ' 14-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Dim intI As Integer   ' COunter

    Try
      ' Switch off the timer
      Me.Timer1.Enabled = False
      ' Show we are initialising
      Status("Initialising...")
      Me.lboPeriod.Items.Clear()
      ' Do we need to initialize period + periodfile tables?
      If (tblPeriod Is Nothing) OrElse (tblPeriodFile Is Nothing) Then
        ' Initialise them
        tblPeriod = tdlPeriods.Tables("Period")
        tblPeriodFile = tdlPeriods.Tables("PeriodFile")
      End If
      ' Load the listbox with the correct period values
      For intI = 0 To tblPeriod.Rows.Count - 1
        Me.lboPeriod.Items.Add(tblPeriod.Rows(intI).Item("Name"))
      Next intI
      ' Show we are done
      Status("Ready")
    Catch ex As Exception
      ' Show error
      HandleErr("frmPeriod/Timer1 error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   lboPeriod_DoubleClick
  ' Goal:   Exit with success
  ' History:
  ' 14-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub lboPeriod_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lboPeriod.DoubleClick
    OK_Button_Click(sender, e)
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   lboPeriod_SelectedIndexChanged
  ' Goal:   Set the period as shown in the textbox
  ' History:
  ' 14-07-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub lboPeriod_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lboPeriod.SelectedIndexChanged
    ' Check if there are any items selected
    If (Me.Visible) AndAlso (Me.lboPeriod.SelectedItems.Count > 0) Then
      Me.tbPeriod.Text = Me.lboPeriod.SelectedItem
    End If
  End Sub
End Class
