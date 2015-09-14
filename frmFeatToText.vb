Imports System.Xml
Public Class frmFeatToText
  ' ====================================== PRIVATE =============================================================
  Private bInit As Boolean = False
  Private loc_strTextDir As String = ""   ' Directory where files are located
  Private loc_strFeatCat As String = ""   ' Category name of the feature (type)
  Private loc_strFeatName As String = ""  ' Name of the feature
  Private loc_strFeatWhite As String = "" ' Restrict features to these
  Private loc_strFeatBlack As String = "" ' Blacklist of feature values
  Private loc_strFeatDbase As String = "" ' Feature in the database we are targeting
  ' ============================================================================================================

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   TextDir
  ' Goal:   Set or change the directory where texts are being kept
  ' History:
  ' 02-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Property TextDir() As String
    Get
      Return loc_strTextDir
    End Get
    Set(ByVal value As String)
      loc_strTextDir = value
      Me.tbTextDir.Text = loc_strTextDir
    End Set
  End Property
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   Filter
  ' Goal:   Set the filter to the necessary text
  ' History:
  ' 02-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public WriteOnly Property Filter() As String
    Set(ByVal value As String)
      Me.tbFeatFilter.Text = value
    End Set
  End Property
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   FeatCat, FeatName, FeatWhiteList
  ' Goal:   Pass the user-selected values back
  ' History:
  ' 02-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public ReadOnly Property FeatCat() As String
    Get
      Return loc_strFeatCat
    End Get
  End Property
  Public ReadOnly Property FeatName() As String
    Get
      Return loc_strFeatName
    End Get
  End Property
  Public ReadOnly Property FeatWhiteList() As String
    Get
      Return loc_strFeatWhite
    End Get
  End Property
  Public ReadOnly Property FeatBlackList() As String
    Get
      Return loc_strFeatBlack
    End Get
  End Property
  Public ReadOnly Property FeatDbase() As String
    Get
      Return loc_strFeatDbase
    End Get
  End Property
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   SetDbaseFeatures
  ' Goal:   Find out which feature names are used in the database, and make these names available to the user
  ' History:
  ' 02-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Function SetDbaseFeatures() As Boolean
    Dim ndxList As XmlNodeList  ' List of nodex
    ' Dim ndxThis As XmlNode      ' Working node
    Dim intI As Integer         ' Counter

    Try
      ' (1) Determine the names of the features from looking at the first record containing them
      ndxList = pdxCrpDbase.SelectNodes("./descendant::Result[1]/child::Feature")
      If (ndxList.Count > 0) Then
        ' Reset the combobox
        With Me.lboFeatDbase
          ' Clear it
          .Items.Clear()
          ' Get a list of features and load the combobox
          For intI = 0 To ndxList.Count - 1
            .Items.Add(ndxList(intI).Attributes("Name").Value)
          Next intI
          ' Set the first one
          .SelectedIndex = 0
        End With
      Else
        ' There are no features, so clear up everything
        Me.lboFeatDbase.Items.Clear()
      End If
      ' Return positively
      Status("okay")
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFeatToText/SetDbaseFeatures error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   mnuCorpusFeaturesToTexts_Click
  ' Goal:   Copy the annotation from one record to similar records
  ' History:
  ' 02-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub frmFeatToText_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Start initialisation and set owner
    Me.Owner = frmMain
    Me.Timer1.Enabled = True
    Status("Initializing...")
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Initialise
  ' History:
  ' 02-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    Try
      ' Switch off timer
      Me.Timer1.Enabled = False
      ' Perform initialisation

      ' Show we are ready
      bInit = True
      ' Wait until we become visible
      While (Not Me.Visible)
        Application.DoEvents()
      End While
      Status("Ready")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFeatToText/Timer1 error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdTextDir_Click
  ' Goal:   Allow changing the directory where texts are located
  ' History:
  ' 02-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdTextDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTextDir.Click
    Dim strDir As String = ""
    Try
      If (GetDirName(Me.FolderBrowserDialog1, strDir, "Directory containing the texts where you want to add features to", Me.tbTextDir.Text)) Then
        loc_strTextDir = strDir
        Me.tbTextDir.Text = strDir
      Else
        Status("The directory has not been changed")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFeatToText/TextDir error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbFeatCat_TextChanged
  ' Goal:   Process changes in the values
  ' History:
  ' 02-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub tbFeatCat_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbFeatCat.TextChanged
    loc_strFeatCat = Me.tbFeatCat.Text
  End Sub
  Private Sub tbFeatname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbFeatname.TextChanged
    loc_strFeatName = Me.tbFeatname.Text
  End Sub
  Private Sub tbFeatValues_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbFeatWhiteList.TextChanged
    loc_strFeatWhite = Me.tbFeatWhiteList.Text
  End Sub
  Private Sub tbFeatBlackList_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbFeatBlackList.TextChanged
    loc_strFeatBlack = Me.tbFeatBlackList.Text
  End Sub
  Private Sub rbRestrNone_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbRestrNone.CheckedChanged
    Try
      If (Me.rbRestrNone.Checked) Then
        Me.tbFeatWhiteList.Text = ""
        Me.tbFeatWhiteList.ReadOnly = True
      Else
        Me.tbFeatWhiteList.ReadOnly = False
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFeatToText/RestrNone error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub lboFeatDbase_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lboFeatDbase.SelectedIndexChanged
    loc_strFeatDbase = Me.lboFeatDbase.SelectedItem
    Me.tbFeatDbase.Text = loc_strFeatDbase
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdOk_Click
  ' Goal:   Close the form positively or negatively
  ' History:
  ' 02-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    Me.DialogResult = DialogResult.OK
    Me.Close()
  End Sub
  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.DialogResult = DialogResult.Cancel
    Me.Close()
  End Sub

End Class