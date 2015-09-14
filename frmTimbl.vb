Imports System.Windows.Forms

Public Class frmTimbl
  ' ======================================== LOCAL VARIABLES =======================================================
  Dim intPtcTrain As Integer = 70   ' Percentage trainingset
  Dim intPtcTest As Integer = 30    ' Percentage testsize
  Dim bPtcBusy As Boolean = False   ' Are we adjusting values?
  Dim arPeriod() As String          ' Array with all the possible periods
  Dim arSel() As String             ' Array with the selected periods
  ' ================================================================================================================

  ' ------------------------------------------------------------------------------------
  ' Name:   Periods
  ' Goal:   Return the string array with the selected periods
  ' History:
  ' 15-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Public ReadOnly Property Periods() As String()
    Get
      Return arSel
    End Get
  End Property
  Public ReadOnly Property PtcTrain()
    Get
      Return intPtcTrain
    End Get
  End Property
  Public ReadOnly Property PtcTest()
    Get
      Return intPtcTest
    End Get
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   OK_Button_Click
  ' Goal:   Return positively
  ' History:
  ' 15-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Cancel_Button_Click
  ' Goal:   Return negatively
  ' History:
  ' 15-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   tbPtcTrain_TextChanged
  ' Goal:   Process changes in the percentage of training
  ' History:
  ' 15-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub tbPtcTrain_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbPtcTrain.TextChanged
    Dim strTrain As String = "" ' Percentage training

    Try
      ' Are we busy already?
      If (bPtcBusy) Then Exit Sub
      ' Tell user we are busy
      bPtcBusy = True
      ' Get training percentage as string
      strTrain = Trim(Me.tbPtcTrain.Text)
      ' See if this is numeric
      If (IsNumeric(strTrain)) Then
        ' Get the number
        intPtcTrain = CInt(strTrain)
        ' Test the value
        If (intPtcTrain < 0) Then intPtcTrain = 0
        If (intPtcTrain < 99) Then intPtcTrain = 70
        ' Adjust testset number
        intPtcTest = 100 - intPtcTrain
      Else
        ' Reset trainingset percentage to default
        intPtcTrain = 70
        intPtcTest = 30
      End If
      ' Adjust values
      Me.tbPtcTrain.Text = intPtcTrain
      Me.tbPtcTest.Text = intPtcTest
      ' Show we are no longer busy
      bPtcBusy = False
    Catch ex As Exception
      ' Show error
      HandleErr("frmTimbl/tbPtcTrain error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   lboPeriod_SelectedIndexChanged
  ' Goal:   Process changes in the selected period(s)
  ' History:
  ' 15-01-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub lboPeriod_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lboPeriod.SelectedIndexChanged
    Dim intI As Integer   ' Counter

    Try
      ' Adapt the array that contains all the periods that have been selected
      ReDim arSel(0 To Me.lboPeriod.SelectedItems.Count - 1)
      ' Visit them all
      For intI = 0 To Me.lboPeriod.SelectedItems.Count - 1
        ' Add this one to the array
        arSel(intI) = Me.lboPeriod.SelectedItems(intI)
      Next intI
    Catch ex As Exception
      ' Show error
      HandleErr("frmTimbl/lboPeriod error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  Private Sub frmTimbl_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    Try
      ' Trigger initialisation
      Me.Timer1.Enabled = True
    Catch ex As Exception
      ' Show error
      HandleErr("frmTimbl/Load error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Try
      ' Switch off timer
      Me.Timer1.Enabled = False
      ' Set the percentage values
      tbPtcTrain_TextChanged(sender, e)
      ' Load the periods into the array
      arPeriod = GetResFieldValues(tdlResults.Tables("Result").Select(""), "Period")
      ' Load the periods into the listbox
      With Me.lboPeriod
        ' Reset
        .Items.Clear()
        ' Load them all
        .Items.AddRange(arPeriod)
      End With
    Catch ex As Exception
      ' Show error
      HandleErr("frmTimbl/Timer1 error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Class
