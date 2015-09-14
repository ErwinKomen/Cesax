Imports System.Windows.Forms

Public Class frmSpss
  ' ============================================== LOCAL VARIABLES ===============================================================
  Private bInit As Boolean = False      '
  Private bSelecting As Boolean = False ' Flag
  Private arFeat() As String            ' Array of features
  Private arType() As String            ' Array of types for features: 
  '                                       "Dep" = dependant
  '                                       "Ind" = independant
  '                                       "Nos" = nostatistics
  ' ==============================================================================================================================
  '---------------------------------------------------------------------------------------------------------
  ' Name:       SetFeatures()
  ' Goal:       initialise the features
  ' History:
  ' 15-07-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public WriteOnly Property Features() As String()
    Set(ByVal value As String())
      Dim intI As Integer     ' Counter

      Try
        ' Get the array
        arFeat = value
        ' Make an array of types
        ReDim arType(0 To arFeat.Length - 1)
        ' Initialise the types
        For intI = 0 To arType.Length - 1
          arType(intI) = "Ind"
        Next intI
        ' Load the listbox
        With Me.lboFeatures
          .Items.Clear()
          .Items.AddRange(arFeat)
        End With
        ' Make sure the other listboxes are reloaded
        DoReloadTypes()
      Catch ex As Exception
        ' Show error
        HandleErr("frmSpss/Features error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      End Try
    End Set
  End Property
  '---------------------------------------------------------------------------------------------------------
  ' Name:       FtTypes()
  ' Goal:       Return the types of the features
  ' History:
  ' 15-07-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public ReadOnly Property FtTypes() As String()
    Get
      Return arType
    End Get
  End Property
  '---------------------------------------------------------------------------------------------------------
  ' Name:       UseCatAsDependant()
  ' Goal:       Flag saying that [Cat] should be used as dependant variable
  ' History:
  ' 12-08-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public ReadOnly Property UseCatAsDependant() As Boolean
    Get
      Return (Me.rbUseCat.Checked = True)
    End Get
  End Property
  '---------------------------------------------------------------------------------------------------------
  ' Name:       OK_Button_Click()
  ' Goal:       Close and show we're okay
  ' History:
  ' 15-07-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       Cancel_Button_Click()
  ' Goal:       Close and show we canceled
  ' History:
  ' 15-07-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       frmSpss_Load()
  ' Goal:       Load the SPSS form
  ' History:
  ' 15-07-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub frmSpss_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Trigger initialisation
    Me.Timer1.Enabled = True
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       Timer1_Tick()
  ' Goal:       Perform the actual initialisation
  ' History:
  ' 15-07-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    ' Switch off timer
    Me.Timer1.Enabled = False
    ' Call the init formula
    DoInit()
  End Sub
  Private Sub DoInit()
    Try
      ' Check state
      If (bInit) Then Exit Sub
      Me.rbUseFeature.Checked = True
      Me.lboDep.Visible = True
      ' Show readiness
      Status("Ready")
      ' Indicate we are ready
      bInit = True
    Catch ex As Exception
      ' Show error
      HandleErr("frmSpss/DoInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------------------------------------------------
  ' Name:       DoReloadTypes()
  ' Goal:       Initialise the three listboxes
  ' History:
  ' 15-07-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub DoReloadTypes()
    Dim intI As Integer   ' Counter

    Try
      ' Clear the three listboxes
      Me.lboDep.Items.Clear()
      Me.lboIndep.Items.Clear()
      Me.lboExclude.Items.Clear()
      ' Walk the array of types
      For intI = 0 To arType.Length - 1
        ' Find out which type this goes to
        Select Case arType(intI)
          Case "Ind"
            Me.lboIndep.Items.Add(arFeat(intI))
          Case "Dep"
            Me.lboDep.Items.Add(arFeat(intI))
          Case "Nos"
            Me.lboExclude.Items.Add(arFeat(intI))
        End Select
      Next intI
      ' Ready!
      Status("Features reloaded")
    Catch ex As Exception
      ' Show error
      HandleErr("frmSpss/DoReloadTypes error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       rbNoStat_CheckedChanged()
  ' Goal:       Change the features of the currently selected one
  ' History:
  ' 15-07-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub rbNoStat_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbNoStat.CheckedChanged
    Try
      ' Validate
      If (Not bInit) OrElse (bSelecting) Then Exit Sub
      SetType(Me.lboFeatures.SelectedIndex, "Nos")
    Catch ex As Exception
      ' Show error
      HandleErr("frmSpss/rbNoStat_CheckedChanged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub rbIndep_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbIndep.CheckedChanged
    Try
      ' Validate
      If (Not bInit) OrElse (bSelecting) Then Exit Sub
      SetType(Me.lboFeatures.SelectedIndex, "Ind")
    Catch ex As Exception
      ' Show error
      HandleErr("frmSpss/rbIndep_CheckedChanged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub rbDep_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbDep.CheckedChanged
    Try
      ' Validate
      If (Not bInit) OrElse (bSelecting) Then Exit Sub
      SetType(Me.lboFeatures.SelectedIndex, "Dep")
    Catch ex As Exception
      ' Show error
      HandleErr("frmSpss/rbDep_CheckedChanged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub SetType(ByVal intSelIdx As Integer, ByVal strType As String)
    Try
      ' Validate
      If (intSelIdx < 0) Then Exit Sub
      ' Set the type of the selected one
      arType(intSelIdx) = strType
      ' Relaod
      DoReloadTypes()
    Catch ex As Exception
      ' Show error
      HandleErr("frmSpss/SetType error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '---------------------------------------------------------------------------------------------------------
  ' Name:       lboFeatures_SelectedIndexChanged()
  ' Goal:       Show which features are selected for this one
  ' History:
  ' 15-07-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Private Sub lboFeatures_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lboFeatures.SelectedIndexChanged
    Dim intIdx As Integer ' Index of this one

    Try
      ' Validate
      If (Not bInit) Then Exit Sub
      ' Show which one belongs to this feature
      intIdx = Me.lboFeatures.SelectedIndex
      ' make sure we don't interrupt anything
      bSelecting = True
      ' Show which type is associated with this one
      Select Case arType(intIdx)
        Case "Dep"
          Me.rbDep.Checked = True
        Case "Ind"
          Me.rbIndep.Checked = True
        Case "Nos"
          Me.rbNoStat.Checked = True
      End Select
      ' Reset flag
      bSelecting = False
    Catch ex As Exception
      ' Show error
      HandleErr("frmSpss/lboFeatures_SelectIdx error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  Private Sub rbUseCat_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbUseCat.CheckedChanged
    Dim intI As Integer ' Counter

    Try
      If (Not bInit) Then Exit Sub
      ' Make features invisible
      Me.lboDep.Visible = False
      ' Make sure no type is "Dep"
      For intI = 0 To arType.Length - 1
        If (arType(intI) = "Dep") Then
          arType(intI) = "Ind"
        End If
      Next intI
    Catch ex As Exception
      ' Show error
      HandleErr("frmSpss/rbUseCat_CheckedChanged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  Private Sub rbUseFeature_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbUseFeature.CheckedChanged
    Try
      If (Not bInit) Then Exit Sub
      ' Make features invisible
      Me.lboDep.Visible = True
    Catch ex As Exception
      ' Show error
      HandleErr("frmSpss/rbUseFeature_CheckedChanged error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Class
