Imports System.Windows.Forms

Public Class frmChart
  ' ===================================== LOCAL VARIABLES ============================================
  Private bSubCl As Boolean = False   ' Should we include breakdown of subordinate clauses?
  Private bAllSc As Boolean = False   ' SHould we include all sections?
  Private strOutType As String = ""   ' Kind of output type
  Private intV1ptc As Integer = 0
  Private intV2ptc As Integer = 0
  Private intSptc As Integer = 0
  ' ==================================================================================================
  Public ReadOnly Property SubCl() As Boolean
    Get
      Return bSubCl
    End Get
  End Property
  Public ReadOnly Property AllSections() As Boolean
    Get
      Return bAllSc
    End Get
  End Property
  Public ReadOnly Property OutputType() As String
    Get
      Return strOutType
    End Get
  End Property
  Public Property V1ptc() As Integer
    Get
      Return intV1ptc
    End Get
    Set(ByVal value As Integer)
      intV1ptc = value
      Me.numV1ptc.Value = intV1ptc
    End Set
  End Property
  Public Property V2ptc() As Integer
    Get
      Return intV2ptc
    End Get
    Set(ByVal value As Integer)
      intV2ptc = value
      Me.numV2ptc.Value = intV2ptc
    End Set
  End Property
  Public Property Sptc() As Integer
    Get
      Return intSptc
    End Get
    Set(ByVal value As Integer)
      intSptc = value
      Me.numSptc.Value = intSptc
    End Set
  End Property
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   OK_Button_Click
  ' Goal:   Produce a "chart" for the current section of the text
  ' History:
  ' 25-09-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   Cancel_Button_Click
  ' Goal:   Produce a "chart" for the current section of the text
  ' History:
  ' 25-09-2012  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  Private Sub chbSubCl_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbSubCl.CheckedChanged
    bSubCl = Me.chbSubCl.Checked
  End Sub

  Private Sub chbAllSections_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbAllSections.CheckedChanged
    bAllSc = Me.chbAllSections.Checked
  End Sub

  Private Sub frmChart_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Set default value: chart output
    Me.rbChart.Checked = True
    strOutType = "Chart"
  End Sub

  Private Sub rbWordOrder_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbWordOrder.CheckedChanged
    strOutType = "WordOrder"
  End Sub

  Private Sub rbChart_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbChart.CheckedChanged
    strOutType = "Chart"
  End Sub

  Private Sub numV1ptc_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numV1ptc.ValueChanged
    intV1ptc = Me.numV1ptc.Value
  End Sub

  Private Sub numV2ptc_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numV2ptc.ValueChanged
    intV2ptc = Me.numV2ptc.Value
  End Sub

  Private Sub numSptc_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles numSptc.ValueChanged
    intSptc = Me.numSptc.Value
  End Sub
End Class
