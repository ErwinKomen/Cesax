Imports System.Windows.Forms

Public Class frmSection
  ' ==================================== LOCAL VARIABLES ==============================
  Private loc_intSection As Integer = 0 ' The section selected
  ' ------------------------------------------------------------------------------------
  ' Name:   Section
  ' Goal:   The number of the selection used
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public ReadOnly Property Section() As Integer
    Get
      ' Return the locally selected section
      Return loc_intSection
    End Get
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   OK_Button_Click
  ' Goal:   Return with the correct dialog result
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Cancel_Button_Click
  ' Goal:   Set the parent of this window
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   frmSection_Load
  ' Goal:   Set the parent of this window
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmSection_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Try
      ' Set the owner form
      Me.Owner = frmMain
      ' Start timer for initialisation
      Me.Timer1.Enabled = True
    Catch ex As Exception
      ' Show error
      HandleErr("frmSection/SectionLoad error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Initialisation of the form: load the listbox
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Try
      ' Switch off the timer
      Me.Timer1.Enabled = False
      ' Load the sections
      SectionsLoad(Me.lboSection)
      ' Set the focus to the listbox
      Me.lboSection.Focus()
    Catch ex As Exception
      ' Show error
      HandleErr("frmSection/Timer1_Tick error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   lboSection_DoubleClick
  ' Goal:   Return OK
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub lboSection_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lboSection.DoubleClick
    Try
      ' Do the same as when the OK button has been clicked
      OK_Button_Click(sender, e)
    Catch ex As Exception
      ' Show error
      HandleErr("frmSection/lboSection_DoubleClick error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   lboSection_SelectedIndexChanged
  ' Goal:   Set the local variable holding the section number
  ' History:
  ' 15-05-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub lboSection_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lboSection.SelectedIndexChanged
    Try
      ' Set the section number
      loc_intSection = Me.lboSection.SelectedIndex
    Catch ex As Exception
      ' Show error
      HandleErr("frmSection/lboSection_SelectedIndexChanged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Class
