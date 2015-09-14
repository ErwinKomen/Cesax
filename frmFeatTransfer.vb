Imports System.Xml
Public Class frmFeatTransfer
  ' ====================================== PRIVATE =============================================================
  Private bInit As Boolean = False
  Private loc_strExportDir As String = ""   ' Directory where files are located
  Private loc_strImportFile As String = ""  ' File we want to load features from
  Private loc_strFeatWhite As String = ""   ' Restrict features to these
  Private loc_strFeatBlack As String = ""   ' Blacklist of feature values
  Private loc_strFeatDbase As String = ""   ' Feature in the database we are targeting
  Private loc_bStatus As Boolean = False    ' Also copy attribute "Status"
  Private loc_bNotes As Boolean = False     ' Also copy attribute "Notes"
  Private loc_bPde As Boolean = False       ' Also copy "Pde" (not an attribute)
  Private loc_bTransExp As Boolean = True   ' Assume export
  Private loc_strMethod As String = "ResId" ' Method used to import
  ' ============================================================================================================

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   ExportDir
  ' Goal:   Set or change the directory where features are to be saved
  ' History:
  ' 02-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public Property ExportDir() As String
    Get
      Return loc_strExportDir
    End Get
    Set(ByVal value As String)
      loc_strExportDir = value
      Me.tbExportDir.Text = loc_strExportDir
    End Set
  End Property
  Public Property ImportFile() As String
    Get
      Return loc_strImportFile
    End Get
    Set(ByVal value As String)
      loc_strImportFile = value
    End Set
  End Property
  Public ReadOnly Property UseStatus() As Boolean
    Get
      Return loc_bStatus
    End Get
  End Property
  Public ReadOnly Property UseNotes As Boolean
    Get
      Return loc_bNotes
    End Get
  End Property
  Public ReadOnly Property UsePde As Boolean
    Get
      Return loc_bPde
    End Get
  End Property
  Public Property Method() As String
    Get
      Return loc_strMethod
    End Get
    Set(ByVal value As String)
      loc_strMethod = value
      Select Case loc_strMethod
        Case "ResId"
          Me.rbResId.Checked = True : Me.rbTextForEtree.Checked = False
        Case Else
          Me.rbResId.Checked = False : Me.rbTextForEtree.Checked = True
      End Select
    End Set
  End Property


  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   Mode
  ' Goal:   Set the mode: export or import
  ' History:
  ' 10-07-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public WriteOnly Property Mode As String
    Set(strMode As String)
      Select Case strMode
        Case "Export", "exp"
          loc_bTransExp = True
        Case "Import", "imp"
          loc_bTransExp = False
      End Select
      If (loc_bTransExp) Then
        Me.TabControl1.SelectedTab = Me.tpExport
      Else
        Me.TabControl1.SelectedTab = Me.tpImport
      End If
      Application.DoEvents()
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
      HandleErr("frmFeatTransfer/SetDbaseFeatures error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   frmFeatTransfer_Load
  ' Goal:   Export or import features to/from a file
  ' History:
  ' 02-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub frmFeatTransfer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
      ' Me.TabControl1.SelectedTab = Me.tpExport
      Me.rbResId.Checked = (loc_strMethod = "ResId")
      Me.rbTextForEtree.Checked = Not Me.rbResId.Checked
      ' Show we are ready
      bInit = True
      ' Wait until we become visible
      While (Not Me.Visible)
        Application.DoEvents()
      End While
      Status("Ready")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFeatTransfer/Timer1 error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdTextDir_Click
  ' Goal:   Allow changing the directory where texts are located
  ' History:
  ' 02-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdExportDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExportDir.Click
    Dim strDir As String = ""
    Try
      If (GetDirName(Me.FolderBrowserDialog1, strDir, "Directory where you want to save the feature values", Me.tbExportDir.Text)) Then
        loc_strExportDir = strDir
        Me.tbExportDir.Text = strDir
      Else
        Status("The directory has not been changed")
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmFeatTransfer/TextDir error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   tbFeatCat_TextChanged
  ' Goal:   Process changes in the values
  ' History:
  ' 02-05-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
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
      HandleErr("frmFeatTransfer/RestrNone error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub lboFeatDbase_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lboFeatDbase.SelectedIndexChanged
    loc_strFeatDbase = Me.lboFeatDbase.SelectedItem
    Me.tbFeatDbase.Text = loc_strFeatDbase
  End Sub
  Private Sub rbResId_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbResId.CheckedChanged
    loc_strMethod = IIf(Me.rbResId.Checked, "ResId", "TextForEtree")
  End Sub

  Private Sub rbTextForEtree_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbTextForEtree.CheckedChanged
    loc_strMethod = IIf(Me.rbTextForEtree.Checked, "TextForEtree", "ResId")
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

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   chbAttrStatus_CheckedChanged
  ' Goal:   Allow copying status and notes attributes
  ' History:
  ' 10-07-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub chbAttrStatus_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chbAttrStatus.CheckedChanged
    loc_bStatus = Me.chbAttrStatus.Checked
  End Sub
  Private Sub chbAttrNotes_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chbAttrNotes.CheckedChanged
    loc_bNotes = Me.chbAttrNotes.Checked
  End Sub
  Private Sub chbChildPde_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chbChildPde.CheckedChanged
    loc_bPde = Me.chbChildPde.Checked
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdImportFile_Click
  ' Goal:   Handle file import
  ' History:
  ' 10-07-2014  ERK Created
  ' 31-12-2014  ERK Added .csv file option
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdImportFile_Click(sender As System.Object, e As System.EventArgs) Handles cmdImportFile.Click
    If (GetFileName(Me.OpenFileDialog1, loc_strExportDir, loc_strImportFile, "Feature files (*.xml;*.csv)|*.xml;*.csv")) Then
      ' Show the file name
      Me.tbImportFile.Text = loc_strImportFile
      '  Me.rbTransImp.Checked = True
    Else
      Status("There was a problem")
    End If
  End Sub

End Class