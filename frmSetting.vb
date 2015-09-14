Public Class frmSetting
  ' =================================LOCAL VARIABLES===================================
  Private bInit As Boolean = False        ' Initialisation flag
  Private bDirty As Boolean = False       ' Dirty flag
  Private bPrjEdBusy As Boolean = False   ' Whether the project editor is busy...
  Private bPrjEdDirty As Boolean = False  ' Whether project editor is dirty or not
  Private tblSettings As New DataTable    ' Section [Setting]
  Private tblPhrTypeList As New DataTable ' Section [PhrType]
  Private tblRefType As New DataTable     ' Section [RefType]
  Private tblPronounList As New DataTable ' Section [Pronoun]
  Private tblCatList As New DataTable     ' Section [Category]
  Private tblNPfeatList As New DataTable  ' Section [NPfeat]
  Private tblConstraints As New DataTable ' Section [Constraints]
  Private intPrjTypeId As Integer = -1    ' The ProjectTypeId of the selected row...
  Private Const def_DATA_DIR As String = "d:\data files"
  Private Const THIS_APPLICATION As String = "Cesax"

  ' ------------------------------------------------------------------------------------
  ' Name:   frmSetting_Load
  ' Goal:   Trigger initialisation
  ' History:
  ' 30-07-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub frmSetting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    ' Set the owner
    Me.Owner = frmMain
    ' Start timer for initialisation
    Me.Timer1.Enabled = True
  End Sub
  Private Sub frmSetting_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
    ' Start timer
    Me.Timer1.Enabled = True
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   Timer1_Tick
  ' Goal:   Call initialiser
  ' History:
  ' 30-07-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
    ' Reset timer
    Me.Timer1.Enabled = False
    ' Call initialisation
    DoInit()
    ' Show we are ready
    Status("Ready")
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   DoInit
  ' Goal:   Try to initialise
  ' History:
  ' 30-07-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub DoInit()
    Dim bAdded As Boolean = False ' Flag

    Try
      ' Is initialisation needed?
      If (Not bInit) Then
        ' Is there a settings object?
        If (tdlSettings Is Nothing) Then
          ' Exit, no initialisation
          Status("No initialisation yet (Settings object is not yet found) ...")
          Exit Sub
        End If
        ' Set the init flag
        bInit = True
        ' Indicate we are initialised
        Status("Settings taken from: " & strSetFile)
      End If
      ' Initialise the tooltips
      SetToolTips()
      ' Set settings table
      tblSettings = tdlSettings.Tables("Setting")
      tblSettings.PrimaryKey = New DataColumn() {tblSettings.Columns("Name")}
      ' Set Coref table: this contains the list of Phrases + children, Can and Must
      tblPhrTypeList = tdlSettings.Tables("PhrType")
      tblPhrTypeList.PrimaryKey = New DataColumn() {tblPhrTypeList.Columns("Node"), tblPhrTypeList.Columns("Child")}
      ' Set RefType table: this contains all the recognized coreference types (Identity, CrossSpeech etc)
      tblRefType = tdlSettings.Tables("RefType")
      tblRefType.PrimaryKey = New DataColumn() {tblRefType.Columns("Name")}
      ' Set the Pronoun table: 
      tblPronounList = tdlSettings.Tables("Pronoun")
      tblPronounList.PrimaryKey = New DataColumn() {tblPronounList.Columns("Name")}
      ' Set the Category table: 
      tblCatList = tdlSettings.Tables("Pronoun")
      tblCatList.PrimaryKey = New DataColumn() {tblCatList.Columns("Name")}
      ' Set the NP features table: 
      tblNPfeatList = tdlSettings.Tables("NPfeat")
      tblNPfeatList.PrimaryKey = New DataColumn() {tblNPfeatList.Columns("Name")}
      ' Set the constraints table: 
      tblConstraints = tdlSettings.Tables("Constraint")
      ' Initialise comboboxes and the different editors
      If (Not InitCombos()) OrElse (Not InitEditors()) Then
        ' Make sure initialisation is not set
        bInit = False
        ' Failure
        Exit Sub
      End If
      ' Update phrase type names, if necessary (if they have not been defined yet)
      PhrTypeInit()
      ' Show settings values
      ShowSettings()
    Catch ex As Exception
      ' Show user that there is an error
      HandleErr("frmSetting/DoInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Make sure initialisation is not set
      bInit = False
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   PhrTypeInit
  ' Goal:   Initialise the phrase type NAMES
  ' History:
  ' 23-11-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function PhrTypeInit() As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT
    Dim intI As Integer       ' Counter
    Dim bFlag As Boolean = False  ' Dirty or not

    Try
      ' Validate
      If (tdlSettings Is Nothing) OrElse (tdlSettings.Tables("PhrType") Is Nothing) Then Return False
      ' Get the datarow
      dtrFound = tdlSettings.Tables("PhrType").Select("", "PhrTypeId ASC")
      ' Walk through all the entries
      For intI = 0 To dtrFound.Length - 1
        ' Access this one
        With dtrFound(intI)
          ' Is the name defined yet?
          If (.Item("Name").ToString = "") Then
            ' Not yet defined, so make a default name
            .Item("Name") = .Item("Node") & "_" & Format(intI + 1, "00")
            ' Make sure settings are applied
            bFlag = True
          End If
        End With
      Next intI
      ' Need to save changes?
      If (bFlag) Then
        ' Try to save the settings file, and don't ask the user about it
        If (Not SaveSettings(strSetFile, False)) Then
          Status("Unable to change settings to file " & strSetFile)
        End If
      End If
      ' Return okay
      Return True
    Catch ex As Exception
      ' Show user that there is an error
      HandleErr("frmSetting/PhrTypeInit error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitCombos
  ' Goal:   Initialise the combobox with engines
  ' History:
  ' 23-11-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitCombos() As Boolean
    ' Get all the values from the KnownColor enumeration.
    Dim colorsArray As System.Array = [Enum].GetValues(GetType(KnownColor))
    Dim allColors(colorsArray.Length) As KnownColor
    Dim intI As Integer   ' Counter

    Try
      ' Access the combobox
      With Me.cboPhraseTarget
        ' Clear the entries
        .Items.Clear()
        ' Add items by hand
        .Items.Add("Any")
        .Items.Add("Dst")
      End With
      ' Access the combobox
      With Me.cboPhraseType
        ' Clear the entries
        .Items.Clear()
        ' Add items by hand
        .Items.Add("Can")
        .Items.Add("Must")
      End With
      ' Get all the system colors
      Array.Copy(colorsArray, allColors, colorsArray.Length)
      ' Access the combobox for colors
      With Me.cboColor
        ' Clear the entries
        .Items.Clear()
        For intI = 0 To allColors.Length - 1
          ' Copy the name of this color
          .Items.Add(allColors(intI).ToString)
        Next intI
        ' Set the value to "Black"
        .SelectedValue = "Black"
      End With
      ' Set the function/relation types
      With Me.cboXrelType
        ' Clear the entries
        .Items.Clear()
        ' Add items by han d
        .Items.Add("Axis") : .Items.Add("Function") : .Items.Add("Attribute")
      End With
      ' Return ok
      Return True
    Catch ex As Exception
      ' Give warning to user
      HandleErr("InitCombos error: " & ex.Message)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitEditors
  ' Goal:   Initialise the different editors with datagridviews and other controls
  '           that are bound to the dataset
  ' History:
  ' 16-10-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitEditors() As Boolean
    Try
      ' Do away with possible previous handlers
      DgvClear(objCed)    ' Query Editor
      DgvClear(objPed)    ' Period Editor
      DgvClear(objPROed)  ' Pronoun editor
      DgvClear(objNPed)   ' NP feature editor
      DgvClear(objCnsEd)  ' Constraint editor
      DgvClear(objStaEd)  ' Status editor
      DgvClear(objTmpEd)  ' Template editor
      DgvClear(objTCed)   ' Template cell editor
      DgvClear(objXrelEd) ' Xpath relation editor
      ' Initialise the query editor (QED)
      InitCorefTypeEditor()
      ' Initialise the constructor editor (CED)
      InitPhraseTypeEditor()
      ' Initialise the pronoun list editor (PROed)
      InitPronounEditor()
      ' Initialize the category list editor (CatEd)
      InitCategoryEditor()
      ' Initialize the NP feature list editor (NPed)
      InitNPfeatEditor()
      ' Initialize the constraints editor (CnsEd)
      InitConstraintEditor()
      ' Initialize the status editor (StaEd)
      InitStatusEditor()
      ' Initialize the template editor (TmpEd)
      InitTemplateEditor()
      ' Initialize the template-cell editor (TCed)
      InitTcellEditor()
      ' Initialize the Xrelations editor (XrelEd)
      InitXrelEditor()
      ' Initialise the Minimally Obligatory Structure editor(MosEd + MosActEd)
      InitMosEditor()
      InitMosActEditor()
      ' Return success
      InitEditors = True
    Catch ex As Exception
      ' Show user that there is an error
      HandleErr("frmSetting/InitEditors error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      InitEditors = False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitCorefTypeEditor
  ' Goal:   Initialise the editor for Coref Types
  ' History:
  ' 23-09-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitCorefTypeEditor() As Boolean
    Try
      ' Initialise the Query DGV handle
      objCed = New DgvHandle
      With objCed
        .Init(Me, tdlSettings, "RefType", "RefTypeId", "RefTypeId", "Name", "", _
              "", "", Me.dgvCorefType, Nothing)
        .BindControl(Me.tbCorefName, "Name", "textbox")
        .BindControl(Me.tbCorefDescr, "Descr", "richtext")
        .BindControl(Me.cboColor, "Color", "combo")
        ' Set the parent table for the [CorefType] editor
        .ParentTable = "RefTypeList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitCorefTypeEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitPhraseTypeEditor
  ' Goal:   Initialise the editor for Phrase Types
  ' History:
  ' 23-09-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitPhraseTypeEditor() As Boolean
    Try
      ' Initialise the Phrase Type DGV handle
      objPed = New DgvHandle
      With objPed
        .Init(Me, tdlSettings, "PhrType", "PhrTypeId", "PhrTypeId", "Name;Node;Child", "", _
               "", "", Me.dgvPhraseType, Nothing)
        .BindControl(Me.tbPhraseName, "Name", "textbox")
        .BindControl(Me.tbPhraseNode, "Node", "textbox")
        .BindControl(Me.tbPhraseChild, "Child", "textbox")
        .BindControl(Me.cboPhraseTarget, "Target", "combo")
        .BindControl(Me.cboPhraseType, "Type", "combo")
        ' Set the parent table for the [PhraseType] editor
        .ParentTable = "PhrTypeList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitPhraseTypeEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitPronounEditor
  ' Goal:   Initialise the editor for Pronoun Types
  ' History:
  ' 23-09-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitPronounEditor() As Boolean
    Try
      ' Initialise the Phrase Type DGV handle
      objPROed = New DgvHandle
      With objPROed
        .Init(Me, tdlSettings, "Pronoun", "PronounId", "Name", "Name", "", _
               "", "", Me.dgvPronoun, Nothing)
        .BindControl(Me.tbProName, "Name", "textbox")
        .BindControl(Me.tbProDescr, "Descr", "textbox")
        .BindControl(Me.tbProPGN, "PGN", "textbox")
        .BindControl(Me.tbProVarOE, "OE", "richtext")
        .BindControl(Me.tbProVarME, "ME", "richtext")
        .BindControl(Me.tbProVarEmodE, "eModE", "richtext")
        .BindControl(Me.tbProVarMBE, "MBE", "richtext")
        .BindControl(Me.tbProNotes, "Notes", "richtext")
        ' Set the parent table for the [PhraseType] editor
        .ParentTable = "PronounList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitPronounEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitCategoryEditor
  ' Goal:   Initialise the editor for Categories
  ' History:
  ' 04-10-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitCategoryEditor() As Boolean
    Try
      ' Initialise the Phrase Type DGV handle
      objCatEd = New DgvHandle
      With objCatEd
        .Init(Me, tdlSettings, "Category", "CatId", "Name", "Name", "", _
               "", "", Me.dgvCategory, Nothing)
        .BindControl(Me.tbCatName, "Name", "textbox")
        .BindControl(Me.tbCatDescr, "Descr", "textbox")
        .BindControl(Me.tbCatType, "Type", "textbox")
        .BindControl(Me.tbCatOE, "OE", "richtext")
        .BindControl(Me.tbCatME, "ME", "richtext")
        .BindControl(Me.tbCatEmodE, "eModE", "richtext")
        .BindControl(Me.tbCatMBE, "MBE", "richtext")
        .BindControl(Me.tbCatNotes, "Notes", "richtext")
        ' Set the parent table for the [PhraseType] editor
        .ParentTable = "CategoryList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitCategoryEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitNPfeatEditor
  ' Goal:   Initialise the editor for NPfeat Types
  ' History:
  ' 23-09-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitNPfeatEditor() As Boolean
    Try
      ' Initialise the Phrase Type DGV handle
      objNPed = New DgvHandle
      With objNPed
        .Init(Me, tdlSettings, "NPfeat", "NPfeatId", "Name", "Name", "", _
               "", "", Me.dgvNPfeat, Nothing)
        .BindControl(Me.tbNPfeatName, "Name", "textbox")
        .BindControl(Me.tbNPfeatDescr, "Descr", "textbox")
        .BindControl(Me.tbNPfeatVariants, "Variants", "richtext")
        ' Set the parent table for the [NPfeat] editor
        .ParentTable = "NPfeatList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitNPfeatEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitConstraintEditor
  ' Goal:   Initialise the editor for the constraints
  ' History:
  ' 23-09-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitConstraintEditor() As Boolean
    Try
      ' Initialise the Phrase Type DGV handle
      objCnsEd = New DgvHandle
      With objCnsEd
        .Init(Me, tdlSettings, "Constraint", "ConstraintId", "Level ASC, Name ASC", "Name", "", _
               "", "", Me.dgvCons, Nothing)
        .BindControl(Me.tbConsName, "Name", "textbox")
        .BindControl(Me.tbConsLevel, "Level", "textbox")
        .BindControl(Me.tbConsMult, "Mult", "textbox")
        .BindControl(Me.tbConsDescr, "Descr", "richtext")
        ' Set the parent table for the [NPfeat] editor
        .ParentTable = "ConstraintList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitConstraintEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitStatusEditor
  ' Goal:   Initialise the editor for the Status
  ' History:
  ' 12-05-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitStatusEditor() As Boolean
    Try
      ' Initialise the Phrase Type DGV handle
      objStaEd = New DgvHandle
      With objStaEd
        .Init(Me, tdlSettings, "Status", "StatusId", "Name ASC", "Name", "", _
               "", "", Me.dgvStatus, Nothing)
        .BindControl(Me.tbStatusName, "Name", "textbox")
        .BindControl(Me.tbStatusDesc, "Descr", "richtext")
        ' Set the parent table for the [NPfeat] editor
        .ParentTable = "StatusList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitStatusEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitTemplateEditor
  ' Goal:   Initialise the editor for the Template
  ' History:
  ' 12-05-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitTemplateEditor() As Boolean
    Try
      ' Initialise the Phrase Type DGV handle
      objTmpEd = New DgvHandle
      With objTmpEd
        .Init(Me, tdlSettings, "Template", "TemplateId", "Name ASC", "Name", "", _
               "", "", Me.dgvTemplate, Nothing)
        .BindControl(Me.tbTemplateName, "Name", "textbox")
        .BindControl(Me.tbTemplateDescr, "Descr", "richtext")
        ' Set the parent table for the [NPfeat] editor
        .ParentTable = "TemplateList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitTemplateEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitMosEditor
  ' Goal:   Initialise the editor for the MOS (Minimally Obligatory Structures)
  ' History:
  ' 08-02-2013  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitMosEditor() As Boolean
    Try
      ' Initialise the MOS Type DGV handle
      objMosEd = New DgvHandle
      With objMosEd
        .Init(Me, tdlSettings, "Mos", "MosId", "Order ASC, Name ASC", "Order;Name", "", _
               "", "", Me.dgvMos, Nothing)
        .BindControl(Me.tbMosName, "Name", "textbox")
        .BindControl(Me.tbMosTrigger, "Trigger", "textbox")
        .BindControl(Me.tbMosCond, "Cond", "richtext")
        ' Set the parent table for the [MOS] editor
        .ParentTable = "MosList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Make sure only one can be selected
      Me.dgvMos.MultiSelect = False
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitMosEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitMosActEditor
  ' Goal:   Initialise the editor for the MosAct
  ' History:
  ' 12-05-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitMosActEditor() As Boolean
    Try
      ' Initialise the Phrase Type DGV handle
      objMosAcEd = New DgvHandle
      With objMosAcEd
        .Init(Me, tdlSettings, "MosAct", "MosActId", "Order ASC", "Order;Name", "", _
               "", "", Me.dgvMosAct, Nothing)
        ' Create a combobox and fill it
        With Me.cboMosActDir
          .Items.Clear()
          .Items.Add("LeftToRight") : .Items.Add("RightToLeft")
        End With
        .BindControl(Me.tbMosActId, "MosActId", "number")
        .BindControl(Me.tbMosActOp, "Name", "textbox")
        .BindControl(Me.tbMosActArg, "Arg", "richtext")
        .BindControl(Me.tbMosActDescr, "Descr", "richtext")
        .BindControl(Me.cboMosActDir, "Dir", "combo")
        ' Do NOT set the parent table for the MosAct editor, because the parent is determined by the selected template
        ' .ParentTable = "MosActList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Make sure only one can be selected
      Me.dgvMosAct.MultiSelect = False
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitMosActEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   InitXrelEditor
  ' Goal:   Initialise the editor for the Xrel
  ' History:
  ' 12-05-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitXrelEditor() As Boolean
    Try
      ' Initialise the Phrase Type DGV handle
      objXrelEd = New DgvHandle
      With objXrelEd
        .Init(Me, tdlSettings, "Xrel", "XrelId", "Name ASC", "Name", "", _
               "", "", Me.dgvXrel, Nothing)
        .BindControl(Me.tbXrelName, "Name", "textbox")
        .BindControl(Me.cboXrelType, "Type", "combo")
        .BindControl(Me.tbXrelXname, "Xname", "richtext")
        .BindControl(Me.tbXrelDescr, "Descr", "richtext")
        ' Set the parent table for the [NPfeat] editor
        .ParentTable = "XrelList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitXrelEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function

  ' ------------------------------------------------------------------------------------
  ' Name:   InitTcellEditor
  ' Goal:   Initialise the editor for the Tcell
  ' History:
  ' 12-05-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Function InitTcellEditor() As Boolean
    Try
      ' Initialise the Phrase Type DGV handle
      objTCed = New DgvHandle
      With objTCed
        .Init(Me, tdlSettings, "Tcell", "TcellId", "TcellId ASC", "Name", "", _
               "", "", Me.dgvTcell, Nothing)
        .BindControl(Me.tbTcellId, "TcellId", "number")
        .BindControl(Me.tbTcellName, "Name", "textbox")
        .BindControl(Me.tbTcellContent, "Content", "richtext")
        .BindControl(Me.tbTcellEnv, "Env", "textbox")
        .BindControl(Me.tbTcellDescr, "Descr", "richtext")
        ' Do NOT set the parent table for the Tcell editor, because the parent is determined by the selected template
        ' .ParentTable = "TcellList"
        ' Scroll to the first entry?
        .Refresh()
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitTcellEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SetToolTips
  ' Goal:   Set the tooltip texts for the different controls
  ' History:
  ' 21-10-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub SetToolTips()
    With Me.ToolTip1
      ' Textboxes
      .SetToolTip(Me.tbWorkDir, "The base directory where CorpusResearchProjects are stored")
      .SetToolTip(Me.tbDriveChange, "Automatically change drive X to Y when reading a CRPX file. Syntax: X>Y;V>W;Z>W")
      .SetToolTip(Me.tbUserName, "Your name as it for instance appears in history of changes")
      ' Check boxes
      .SetToolTip(Me.chbUserProPGN, "Ask user to supply person/gender/number of a pronoun or demonstrative")
      .SetToolTip(Me.chbShowAllAnt, "Show all possible antecedents rather then the first X")
      .SetToolTip(Me.chbStatusAutoChanged, "Automaticaly change the [status] of a database record to 'Changed' when it is changed")
      .SetToolTip(Me.chbStatusAutoNotes, "Automatically add Date/Time to the [notes] section of a database item when it is changed")
      .SetToolTip(Me.chbShowCODE, "Show the CODE nodes in the syntax tab page")
      .SetToolTip(Me.chbDebugging, "Show the command prompt windows where the queries are executed")
      .SetToolTip(Me.chbCheckCat, "Check and possibly add new categories when a new version of Cesax is installed")
      .SetToolTip(Me.chbCheckCns, "Check and possibly modify the constraint hierarchy when a new versino of Cesax is installed")
      .SetToolTip(Me.chbUpdStartup, "Check for updates on startup of Cesax")
      .SetToolTip(Me.chbGenActionInet, "Keep track of my activities with CorpusStudio")
      ' Date boxes
      ' Data grid views
      ' Other controls
    End With
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   ShowSettings
  ' Goal:   Show the settings
  ' History:
  ' 30-07-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub ShowSettings()

    Try
      ' There is a settings object, so take over its values
      Me.tbWorkDir.Text = GetTableSetting(tdlSettings, "WorkDir")
      Me.tbCorpusStudio.Text = GetTableSetting(tdlSettings, "CrpDir")
      Me.tbPeriodDef.Text = GetTableSetting(tdlSettings, "PeriodDefinition")
      Me.tbChainDict.Text = GetTableSetting(tdlSettings, "ChainDictionary")
      Me.tbDriveChange.Text = GetTableSetting(tdlSettings, "DriveChange")
      Me.tbUserName.Text = GetTableSetting(tdlSettings, "UserName")
      Me.tbMaxIPdist.Text = GetTableSetting(tdlSettings, "MaxIPdist")
      Me.tbLogBase.Text = GetTableSetting(tdlSettings, "LogBase")
      Me.tbAbsMaxIPdist.Text = GetTableSetting(tdlSettings, "AbsMaxIPdist")
      Me.tbProfile.Text = GetTableSetting(tdlSettings, "ProfileLevel")
      Me.tbPhraseInclude.Text = GetTableSetting(tdlSettings, "PhraseInclude")
      Me.tbPhraseExclude.Text = GetTableSetting(tdlSettings, "PhraseExclude")
      Me.chbUserProPGN.Checked = (GetTableSetting(tdlSettings, "UserProPGN") = "True")
      Me.chbShowAllAnt.Checked = (GetTableSetting(tdlSettings, "ShowAllAnt") = "True")
      Me.chbStatusAutoChanged.Checked = (GetTableSetting(tdlSettings, "StatusAutoChanged") = "True")
      Me.chbStatusAutoNotes.Checked = (GetTableSetting(tdlSettings, "StatusAutoNotes") = "True")
      Me.chbShowCODE.Checked = (GetTableSetting(tdlSettings, "ShowCode") = "True")
      Me.chbDebugging.Checked = (GetTableSetting(tdlSettings, "Debugging") = "True")
      Me.chbCheckCns.Checked = (GetTableSetting(tdlSettings, "CheckCns") = "True")
      Me.chbCheckCat.Checked = (GetTableSetting(tdlSettings, "CheckCat") = "True")
      Me.chbUpdStartup.Checked = (My.Settings.UpdStartup = "True")
      Me.chbGenActionInet.Checked = (My.Settings.GenActionInet = "True")
      ' Show where the settings file resides
      Me.tbSettingsFile.Text = strSetFile
    Catch ex As Exception
      ' Give warning to user
      HandleErr("frmSetting/ShowSettings error: " & ex.Message)
      Exit Sub
    End Try
  End Sub

  ' ------------------------------------------------------------------------------------
  ' Name:   tbWorkDir_TextChanged...
  ' Goal:   Set the dirty flag if something is changed or added or removed
  ' History:
  ' 30-07-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub tbWorkDir_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbWorkDir.TextChanged
    MakeDirty()
  End Sub
  Private Sub tbDriveChange_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbDriveChange.TextChanged
    MakeDirty()
  End Sub
  Private Sub tbUserName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbUserName.TextChanged
    MakeDirty()
  End Sub
  Private Sub tbSchema_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    MakeDirty()
  End Sub
  Private Sub dgvProjType_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
    MakeDirty()
  End Sub
  Private Sub dgvProjType_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs)
    MakeDirty()
  End Sub
  Private Sub dgvProjType_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs)
    MakeDirty()
  End Sub
  Private Sub chbDebugging_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbDebugging.CheckedChanged
    MakeDirty()
  End Sub
  Private Sub chbShowCode_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbShowCODE.CheckedChanged
    MakeDirty()
  End Sub
  Private Sub chbSwitchOutput_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    MakeDirty()
  End Sub
  Private Sub chbUserProPGN_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbUserProPGN.CheckedChanged
    MakeDirty()
  End Sub
  Private Sub chbShowAllAnt_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbShowAllAnt.CheckedChanged
    MakeDirty()
  End Sub
  Private Sub chbStatusAutoChanged_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbStatusAutoChanged.CheckedChanged
    MakeDirty()
  End Sub
  Private Sub chbStatusAutoNotes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbStatusAutoNotes.CheckedChanged
    MakeDirty()
  End Sub
  Private Sub chbSyncEntry_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    MakeDirty()
  End Sub
  Private Sub chbSyncExit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    MakeDirty()
  End Sub
  Private Sub tbMaxIPdist_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbMaxIPdist.TextChanged
    MakeDirty()
  End Sub
  Private Sub tbLogBase_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbLogBase.TextChanged
    MakeDirty()
  End Sub
  Private Sub tbAbsMaxIPdist_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbAbsMaxIPdist.TextChanged
    MakeDirty()
  End Sub
  Private Sub tbProfile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbProfile.TextChanged
    MakeDirty()
  End Sub
  Private Sub chbCheckCat_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbCheckCat.CheckedChanged
    MakeDirty()
  End Sub
  Private Sub chbCheckCns_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbCheckCns.CheckedChanged
    MakeDirty()
  End Sub
  Private Sub chbGenActionInet_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbGenActionInet.CheckedChanged
    MakeDirty()
  End Sub
  Private Sub chbUpdStartup_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chbUpdStartup.CheckedChanged
    MakeDirty()
  End Sub
  '---------------------------------------------------------------
  ' Name:       TryExit()
  ' Goal:       See if we can exit gracefully
  ' History:
  ' 04-12-2008  ERK Created
  '---------------------------------------------------------------
  Private Sub TryExit()
    ' Save the settings, and do ask the user if he wants to
    If (Not SaveSettings(strSetFile)) Then
      Status("Unable to change settings to file " & strSetFile)
    End If
    ' We can now safely exit
    Me.Hide()
  End Sub

  '---------------------------------------------------------------
  ' Name:       cmdCancel_Click()
  ' Goal:       Do not apply the indicated changes
  ' History:
  ' 07-01-2009  ERK Created
  '---------------------------------------------------------------
  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    ' Just close this form completely
    Me.Close()
  End Sub

  '---------------------------------------------------------------
  ' Name:       cmdOk_Click()
  ' Goal:       Apply the indicated changes
  ' History:
  ' 07-01-2009  ERK Created
  '---------------------------------------------------------------
  Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
    ' Try to save the settings file, and don't ask the user about it
    If (Not SaveSettings(strSetFile, False)) Then
      Status("Unable to change settings to file " & strSetFile)
    End If
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  TryAddNew
  ' Goal :  Try add a new element
  ' History:
  ' 02-10-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub TryAddNew(ByRef objThis As DgvHandle, ByVal strElement As String, _
                        ByVal ParamArray arFldVal() As String)
    Dim strElName As String = ""      ' Name of the new element
    Dim strText As String = ""        ' Initial text of the query
    Dim bRemNodes As Boolean = False  ' Whether to remove nodes or not
    Dim bPrintIdc As Boolean = False  ' Whether to print indices or not
    Dim arAllFldVal() As String       ' Array that combines all field/value combinations
    Dim intI As Integer               ' Counter

    Try
      ' Validate
      If (objThis Is Nothing) Then Exit Sub
      If (strElement = "") Then strElement = "(element)"
      ' Try get a name for this element
      strElName = GetName("What is the name of the new " & strElement & "?")
      ' Validate
      If (strElName = "") Then
        ' Warn the user
        MsgBox("You should provide the " & strElement & " with a name. Try again!")
        Exit Sub
      End If
      ' Does this name entry already exist?
      If (objThis.Exists("Name='" & strElName & "'")) Then
        ' Warn user and leave!
        MsgBox("A " & strElement & " with this name is already present in the list." & vbCrLf & _
               "Try again with another name!")
        Exit Sub
      End If
      ' Make a unified array
      If (arFldVal.Length > 0) Then
        ReDim arAllFldVal(0 To arFldVal.Length + 1)
        arAllFldVal(0) = "Name"
        arAllFldVal(1) = strElName
        For intI = 0 To arFldVal.Length - 1
          arAllFldVal(2 + intI) = arFldVal(intI)
        Next intI
      Else
        ReDim arAllFldVal(0 To 1)
        arAllFldVal(0) = "Name"
        arAllFldVal(1) = strElName
      End If
      ' Call function to perform the action
      If (objThis.AddNew(arAllFldVal)) Then
        ' Set dirty flag
        MakeDirty()
      Else
        ' Failure: leave
        Exit Sub
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSetting/TryAddNew error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  TryAddNewMosAct
  ' Goal :  Try add a new MosAct element
  ' History:
  ' 08-02-2013  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub TryAddNewMosAct(ByRef objThis As DgvHandle, ByVal intMosId As Integer)
    Dim strElName As String = ""      ' Name of the new element
    Dim strText As String = ""        ' Initial text of the query
    Dim bRemNodes As Boolean = False  ' Whether to remove nodes or not
    Dim bPrintIdc As Boolean = False  ' Whether to print indices or not
    Dim arAllFldVal() As String       ' Array that combines all field/value combinations
    Dim dtrMosAct() As DataRow        ' Result of SELECT
    Dim dtrFound() As DataRow         ' Result of SELECT
    'Dim intI As Integer               ' Counter

    Try
      ' Validate
      If (objThis Is Nothing) Then Exit Sub
      ' Make a unified array
      ReDim arAllFldVal(0 To 3)
      ReDim arAllFldVal(0 To 5)
      arAllFldVal(0) = "Name"
      arAllFldVal(1) = "Insert"
      arAllFldVal(2) = "MosId"
      arAllFldVal(3) = intMosId
      arAllFldVal(4) = "Dir"
      arAllFldVal(5) = "LeftToRight"
      ' Call function to perform the action
      If (objThis.AddNew(arAllFldVal)) Then
        ' Set dirty flag
        MakeDirty()
      Else
        ' Failure: leave
        Exit Sub
      End If
      ' Find out what was made: the first one when listing [MosAct] in decreasing order
      dtrMosAct = tdlSettings.Tables("MosAct").Select("", "MosActId DESC")
      If (dtrMosAct.Length = 0) Then Status("Could not create a MosAct") : Exit Sub
      ' Set the parent of the MOS appropriately
      dtrFound = tdlSettings.Tables("Mos").Select("MosId=" & intMosId)
      If (dtrFound.Length > 0) Then
        dtrMosAct(0).SetParentRow(dtrFound(0))
      End If
      ' Adjust the value of "order" for the latest addition
      dtrMosAct(0).Item("Order") = tdlSettings.Tables("MosAct").Select("MosId=" & intMosId).Length
      ' Ready!
      Status("Okay")
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSetting/TryAddNewMosAct error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  cmdNew_Click
  ' Goal :  Make a new X, where X depends on which tabpage is visible
  ' History:
  ' 02-06-2010  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub cmdNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNew.Click
    Dim intLen As Integer = 0   ' Length

    Try
      ' Action depends on the tabpage we are on
      Select Case Me.TabControl1.SelectedTab.Name
        Case Me.tpCoref.Name
          ' Double check and execute
          TryAddNew(objCed, "Coreference type")
        Case Me.tpNPfeat.Name
          ' Double check and execute
          TryAddNew(objNPed, "NP feature type")
        Case Me.tpPhrase.Name
          ' Double check and execute
          TryAddNew(objPed, "Phrase type", "Node", "-", "Type", "Can", "Target", "Dst", "Child", "*")
        Case Me.tpPronoun.Name
          ' Double check and execute
          TryAddNew(objPROed, "Pronoun class")
        Case Me.tpCons.Name
          ' Double check and execute
          TryAddNew(objCnsEd, "Constraint")
        Case Me.tpCat.Name
          ' Double check and execute
          TryAddNew(objCatEd, "Category")
        Case Me.tpStatus.Name
          ' Double check and execute
          TryAddNew(objStaEd, "Status")
        Case Me.tpCharting.Name
          ' Double check and execute
          TryAddNew(objTmpEd, "Template")
        Case Me.tpXrel.Name
          ' Double check and execute
          TryAddNew(objXrelEd, "Query relation/function")
          ' Make sure we set the combobox in a default position
          Me.cboXrelType.SelectedIndex = 0
          ' Make sure changes are processed
          If (CtlChanged(objXrelEd, Me.cboXrelType, bInit)) Then MakeDirty()
        Case Me.tpMos.Name
          ' Get current length
          intLen = tdlSettings.Tables("Mos").Rows.Count
          ' Double check and execute
          TryAddNew(objMosEd, "Mos", "Order", intLen + 1)
      End Select
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSetting/New error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
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
      ' First enquire
      If (MsgBox(strQ, MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes) Then
        ' TRy to remove the selected object
        If (Not objThis.Remove) Then
          ' Unsuccesful...
          Status("Unable to delete this element")
        Else
          ' Make sure dirty bit is set
          MakeDirty()
          ' Show status
          Status("Deleted")
        End If
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSetting/TryRemove error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  cmdDel_Click
  ' Goal :  Delete selected item X, where X depends on which tabpage is visible
  ' History:
  ' 02-06-2010  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub cmdDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDel.Click
    Dim strQ As String = "Would you really like to delete the selected row?"

    Try
      ' Action depends on the tabpage we are on
      Select Case Me.TabControl1.SelectedTab.Name
        Case Me.tpCoref.Name
          ' Double check and execute
          TryRemove(objCed)
        Case Me.tpNPfeat.Name
          ' Double check and execute
          TryRemove(objNPed)
        Case Me.tpPhrase.Name
          ' Double check and execute
          TryRemove(objPed)
        Case Me.tpPronoun.Name
          ' Double check and execute
          TryRemove(objPROed)
        Case Me.tpCons.Name
          ' Double check and execute
          TryRemove(objCnsEd)
        Case Me.tpCat.Name
          ' Double check and execute
          TryRemove(objCatEd)
        Case Me.tpStatus.Name
          ' Double check and execute
          TryRemove(objStaEd)
        Case Me.tpCharting.Name
          ' Double check and execute
          TryRemove(objTmpEd)
        Case Me.tpXrel.Name
          ' Double check and execute
          TryRemove(objXrelEd)
      End Select
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSetting/Del error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------
  ' Name:       DoSaving()
  ' Goal:       Save the results to the indicated location
  ' History:
  ' 04-12-2008  ERK Created
  '---------------------------------------------------------------
  Private Sub DoSaving(ByVal strFile As String)
    Dim intI As Integer           ' Row number of the project where we are
    Dim dtrParent As DataRow      ' Parent for ProjType lines

    ' Adapt the general values in the dataset
    SetTableSetting(tdlSettings, "WorkDir", Me.tbWorkDir.Text)
    SetTableSetting(tdlSettings, "CrpDir", Me.tbCorpusStudio.Text)
    SetTableSetting(tdlSettings, "PeriodDefinition", Me.tbPeriodDef.Text)
    SetTableSetting(tdlSettings, "ChainDictionary", Me.tbChainDict.Text)
    SetTableSetting(tdlSettings, "DriveChange", Me.tbDriveChange.Text)
    SetTableSetting(tdlSettings, "UserName", Me.tbUserName.Text)
    SetTableSetting(tdlSettings, "MaxIPdist", Me.tbMaxIPdist.Text)
    SetTableSetting(tdlSettings, "AbsMaxIPdist", Me.tbAbsMaxIPdist.Text)
    SetTableSetting(tdlSettings, "LogBase", Me.tbLogBase.Text)
    SetTableSetting(tdlSettings, "ProfileLevel", Me.tbProfile.Text)
    SetTableSetting(tdlSettings, "UserProPGN", IIf(Me.chbUserProPGN.Checked, "True", "False"))
    SetTableSetting(tdlSettings, "ShowAllAnt", IIf(Me.chbShowAllAnt.Checked, "True", "False"))
    SetTableSetting(tdlSettings, "ShowCode", IIf(Me.chbShowCODE.Checked, "True", "False"))
    SetTableSetting(tdlSettings, "Debugging", IIf(Me.chbDebugging.Checked, "True", "False"))
    SetTableSetting(tdlSettings, "CheckCat", IIf(Me.chbCheckCat.Checked, "True", "False"))
    SetTableSetting(tdlSettings, "CheckCns", IIf(Me.chbCheckCns.Checked, "True", "False"))
    SetTableSetting(tdlSettings, "StatusAutoChanged", IIf(Me.chbStatusAutoChanged.Checked, "True", "False"))
    SetTableSetting(tdlSettings, "StatusAutoNotes", IIf(Me.chbStatusAutoNotes.Checked, "True", "False"))
    SetTableSetting(tdlSettings, "PhraseInclude", Me.tbPhraseInclude.Text)
    SetTableSetting(tdlSettings, "PhraseExclude", Me.tbPhraseExclude.Text)
    My.Settings.GenActionInet = IIf(Me.chbGenActionInet.Checked, "True", "False")
    My.Settings.UpdStartup = IIf(Me.chbUpdStartup.Checked, "True", "False")
    ' Accept changes
    tdlSettings.AcceptChanges()
    ' Make sure that all <PhrType> lines have as father the <PhrTypeList> item
    dtrParent = tdlSettings.Tables("PhrTypeList").Rows(0)
    With tblPhrTypeList
      For intI = 1 To .Rows.Count
        .Rows(intI - 1).SetParentRow(dtrParent)
      Next intI
    End With
    ' Make sure that all <RefType> lines have as father the <RefTypeList> item
    dtrParent = tdlSettings.Tables("RefTypeList").Rows(0)
    With tblRefType
      For intI = 1 To .Rows.Count
        .Rows(intI - 1).SetParentRow(dtrParent)
      Next intI
    End With
    ' Save the settings
    Call XmlSaveSettings(strFile)
    My.Settings.Save()
    ' Turn off the dirty bit
    bDirty = False
    ' Make sure the changes become available straight away
    ReadSettings()
    ' Turn off Accept button
    Me.cmdOk.Enabled = False
    ' Change the Cancel button
    Me.cmdCancel.Text = "E&xit"
  End Sub
  '---------------------------------------------------------------
  ' Name:       SaveSettings()
  ' Goal:       Save the results to the indicated location
  ' History:
  ' 04-12-2008  ERK Created
  '---------------------------------------------------------------
  Private Function SaveSettings(ByVal strFile As String, Optional ByVal bAsk As Boolean = True) _
    As Boolean
    ' Is saving needed?
    If (bDirty) AndAlso (bInit) Then
      ' Do we need to ask?
      If (bAsk) Then
        ' Ask user if he wants saving
        Select Case MsgBox("Save settings before continuing?", MsgBoxStyle.YesNoCancel)
          Case MsgBoxResult.Ok, MsgBoxResult.Yes
            Call DoSaving(strFile)
            ' Return success
            SaveSettings = True
            Status("Settings saved")
          Case MsgBoxResult.No
            ' Still okay
            SaveSettings = True
            Status("Settings saved")
          Case MsgBoxResult.Cancel
            ' Return failure
            SaveSettings = False
            Status("Aborted")
        End Select
      Else
        ' Save without asking
        Call DoSaving(strFile)
        Status("Settings saved")
        ' Return success
        SaveSettings = True
      End If
    Else
      ' Return success ( but nothing is actually saved)
      SaveSettings = True
    End If
  End Function

  '---------------------------------------------------------------
  ' Name:       cmdSetFile_Click()
  ' Goal:       Open a settings file, and start using this one from now on
  ' History:
  ' 10-12-2008  ERK Created
  '---------------------------------------------------------------
  Private Sub cmdSetFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSetFile.Click
    Dim strFile As String = ""

    ' Get the settings file name from the user
    If (GetFileName(Me.OpenFileDialog1, def_DATA_DIR, strFile, THIS_APPLICATION & _
                    " settings file (*.xml)|*.xml")) Then
      ' Check whether we already have data
      If (Not tdlSettings Is Nothing) Then
        ' Save the new settings location
        Call SaveSetting(THIS_APPLICATION, "Settings", "Location", strFile)
        ' Tell user to restart
        MsgBox("You need to restart " & THIS_APPLICATION & " in order to make the new settings effective")
      ElseIf (ReadSettingsXML(strFile)) Then
        ' If all went well, then save this new settings file name
        Call SaveSetting(THIS_APPLICATION, "Settings", "Location", strFile)
        ' Assign the GLOBAL variable to remember the correct settings file name
        strSetFile = strFile
        ' Now actually load these settings, so that they are visible on the form
        Call ShowSettings()
      Else
        ' Warn the user
        MsgBox("Unable to read the settings file " & strFile & _
          vbCrLf & "Please repair this file manually, and try again")
      End If
      '' Try to read the settings from this file
      'If (ReadSettingsXML(strFile)) Then
      '  ' If all went well, then save this new settings file name
      '  Call SaveSetting(THIS_APPLICATION, "Settings", "Location", strFile)
      '  ' Assign the GLOBAL variable to remember the correct settings file name
      '  strSetFile = strFile
      '  ' Now actually load these settings, so that they are visible on the form
      '  Call ShowSettings()
      'Else
      '  ' Warn the user
      '  MsgBox("Unable to read the settings file " & strFile & _
      '    vbCrLf & "Please repair this file manually, and try again")
      'End If
    End If
  End Sub
  '---------------------------------------------------------------
  ' Name:       MakeDirty()
  ' Goal:       Signal that something needs to be saved
  ' History:
  ' 04-12-2008  ERK Created
  '---------------------------------------------------------------
  Private Sub MakeDirty()
    ' Are we already set?
    If (Not bDirty) AndAlso (bInit) Then
      ' Set dirty bit
      bDirty = True
      ' Enable Cancel button
      Me.cmdOk.Enabled = True
    End If
  End Sub

  '---------------------------------------------------------------
  ' Name:       cmdWorkDir_Click()
  ' Goal:       Allow user to change the working directory for CorpusStudio
  ' History:
  ' 23-11-2009  ERK Created
  '---------------------------------------------------------------
  Private Sub cmdWorkDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWorkDir.Click
    Dim strSelDir As String = ""  ' The selected directory

    ' Try get directory
    If (GetDirName(Me.FolderBrowserDialog1, strSelDir, _
                   "Select the directory where corpora are located", strWorkDir)) Then
      ' Replace what is shown
      Me.tbWorkDir.Text = strSelDir
      ' Show we are dirty
      MakeDirty()
    End If
  End Sub

  '---------------------------------------------------------------
  ' Name:       cmdPeriodDef_Click()
  ' Goal:       Allow user to change the directory where the period definitions come from
  ' History:
  ' 28-05-2010  ERK Created
  '---------------------------------------------------------------
  Private Sub cmdPeriodDef_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPeriodDef.Click
    Dim strFileName As String = ""  ' The selected filename

    ' Try get directory
    If (GetFileName(Me.OpenFileDialog1, strWorkDir, _
                   strFileName, "English period definitions (*.xml)|*.xml")) Then
      ' Replace what is shown
      Me.tbPeriodDef.Text = strFileName
      strPeriodFile = strFileName
      ' Read this period into memory
      TryReadPeriods()
      ' Show we are dirty
      MakeDirty()
    End If
  End Sub

  '---------------------------------------------------------------
  ' Name:       cmdChainDict_Click()
  ' Goal:       Allow user to change the location of the chain dictionary
  ' History:
  ' 06-07-2010  ERK Created
  '---------------------------------------------------------------
  Private Sub cmdChainDict_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdChainDict.Click
    Dim strFileName As String = ""  ' The selected filename
    Dim objSaveDialog As New SaveFileDialog

    Try
      ' Try get a new place to SAVE the chain dictionary
      With objSaveDialog
        ' Set the existing filename as default
        .FileName = IO.Path.GetFileNameWithoutExtension(strChainDict)
        ' Set initial directory
        .InitialDirectory = strWorkDir
        ' Set filter
        .Filter = "Chain dictionary (*.xml)|*.xml"
        ' Get the result
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.Yes, Windows.Forms.DialogResult.OK
            ' Process the resul;t
            strChainDict = .FileName
            Me.tbChainDict.Text = strChainDict
            ' Make sure it gets saved
            ChainDictSetDirty()
            ' Show we are dirty
            MakeDirty()
        End Select
      End With
      'If (GetFileName(Me.OpenFileDialog1, strWorkDir, _
      '               strFileName, "Chain dictionary (*.xml)|*.xml")) Then
      '  ' Replace what is shown
      '  Me.tbChainDict.Text = strFileName
      '  strChainDict = strFileName
      '  ' Read this chain dictionary into memory
      '  TryReadChainDict()
      '  ' Make sure it gets saved!!!
      '  ChainDictSetDirty()
      '  ' Show we are dirty
      '  MakeDirty()
      'End If
    Catch ex As Exception
      ' Show message
      HandleErr("frmSetting/ChainDict error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  '------------------------------------------------------------------------------------
  ' Name:       cmdCorpusStudio_Click()
  ' Goal:       Allow user to change the location of the CorpusStudio working directory
  ' History:
  ' 21-12-2010  ERK Created
  '------------------------------------------------------------------------------------
  Private Sub cmdCorpusStudio_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCorpusStudio.Click
    Dim strSelDir As String = ""  ' The selected directory

    Try
      ' Try get directory
      If (GetDirName(Me.FolderBrowserDialog1, strSelDir, _
                     "Select the directory where Corpus Research Project files (.crpx) are stored", strWorkDir)) Then
        ' Replace what is shown
        Me.tbCorpusStudio.Text = strSelDir
        ' Show we are dirty
        MakeDirty()
      End If
    Catch ex As Exception
      ' Show message
      HandleErr("frmSetting/CorpusStudio error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  cmdCnsUp_Click
  ' Goal :  Try to move the constraint one step up
  ' History:
  ' 27-07-2010  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub cmdCnsUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCnsUp.Click
    Dim intSelId As Integer   ' ID of selected row
    Dim intPrevId As Integer  ' ID of previous row
    Dim intI As Integer       ' Counter
    Dim intLevel As Integer   ' Swap value
    Dim dtrFound() As DataRow ' Result of select
    Dim dtrThis As DataRow    ' My own datarow
    Dim dtrPrev As DataRow    ' The preceding datarow we need to swap with

    Try
      ' Find out which constraint is selected
      With Me.dgvCons
        For intI = 0 To .Rows.Count - 1
          ' Is this row selected?
          If (.Rows(intI).Selected) Then
            ' Yes, this row is selected - can we move this upwards at all?
            If (intI = 0) Then Exit Sub
            ' Get the ID of the selected row and of the previous row
            intSelId = .Rows(intI).Cells(0).Value
            intPrevId = .Rows(intI - 1).Cells(0).Value
            ' Yes, we can move it upwards - let's get the appropriate IDs of the datastructure
            dtrFound = tdlSettings.Tables("Constraint").Select("ConstraintId=" & intSelId)
            If (dtrFound.Length = 0) Then Exit Sub
            dtrThis = dtrFound(0)
            dtrFound = tdlSettings.Tables("Constraint").Select("ConstraintId=" & intPrevId)
            If (dtrFound.Length = 0) Then Exit Sub
            dtrPrev = dtrFound(0)
            ' Now do the exchange of levels...
            intLevel = dtrThis.Item("Level")
            dtrThis.Item("Level") = dtrPrev.Item("Level")
            dtrPrev.Item("Level") = intLevel
            ' Now exit the subroutine
            Exit Sub
          End If
        Next intI
      End With
    Catch ex As Exception
      ' Show message
      HandleErr("frmSetting/CnsUp error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  cmdCnsDown_Click
  ' Goal :  Try to move the constraint one step down
  ' History:
  ' 27-07-2010  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub cmdCnsDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCnsDown.Click
    Dim intSelId As Integer   ' ID of selected row
    Dim intNextId As Integer  ' ID of following row
    Dim intI As Integer       ' Counter
    Dim intLevel As Integer   ' Swap value
    Dim dtrFound() As DataRow ' Result of select
    Dim dtrThis As DataRow    ' My own datarow
    Dim dtrNext As DataRow    ' The following datarow we need to swap with

    Try
      ' Find out which constraint is selected
      With Me.dgvCons
        For intI = 0 To .Rows.Count - 1
          ' Is this row selected?
          If (.Rows(intI).Selected) Then
            ' Yes, this row is selected - can we move this down at all?
            If (intI = .Rows.Count - 1) Then Exit Sub
            ' Get the ID of the selected row and of the next row
            intSelId = .Rows(intI).Cells(0).Value
            intNextId = .Rows(intI + 1).Cells(0).Value
            ' Yes, we can move it upwards - let's get the appropriate IDs of the datastructure
            dtrFound = tdlSettings.Tables("Constraint").Select("ConstraintId=" & intSelId)
            If (dtrFound.Length = 0) Then Exit Sub
            dtrThis = dtrFound(0)
            dtrFound = tdlSettings.Tables("Constraint").Select("ConstraintId=" & intNextId)
            If (dtrFound.Length = 0) Then Exit Sub
            dtrNext = dtrFound(0)
            ' Now do the exchange of levels...
            intLevel = dtrThis.Item("Level")
            dtrThis.Item("Level") = dtrNext.Item("Level")
            dtrNext.Item("Level") = intLevel
            ' Now exit the subroutine
            Exit Sub
          End If
        Next intI
      End With
    Catch ex As Exception
      ' Show message
      HandleErr("frmSetting/CnsDown error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  tbCorefDescr_TextChanged
  ' Goal :  (1) Process the change immediately in [tdlSettings]
  '         (2) Set dirty flag...
  ' History:
  ' 23-04-2010  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub tbCorefDescr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbCorefDescr.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objCed, Me.tbCorefDescr, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbCorefName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbCorefName.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objCed, Me.tbCorefName, bInit)) Then MakeDirty()
  End Sub
  Private Sub cboColor_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboColor.SelectedIndexChanged
    ' Make sure changes are processed
    If (CtlChanged(objCed, Me.cboColor, bInit)) Then MakeDirty()
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  tbPhraseNode_TextChanged
  ' Goal :  (1) Process the change immediately in [tdlSettings]
  '         (2) Set dirty flag...
  ' History:
  ' 23-04-2010  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub tbPhraseNode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbPhraseNode.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objPed, Me.tbPhraseNode, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbPhraseChild_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbPhraseChild.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objPed, Me.tbPhraseChild, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbPhraseName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbPhraseName.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objPed, Me.tbPhraseName, bInit)) Then MakeDirty()
  End Sub
  Private Sub cboPhraseType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPhraseType.SelectedIndexChanged
    ' Make sure changes are processed
    If (CtlChanged(objPed, Me.cboPhraseType, bInit)) Then MakeDirty()
  End Sub
  Private Sub cboPhraseTarget_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboPhraseTarget.SelectedIndexChanged
    ' Make sure changes are processed
    If (CtlChanged(objPed, Me.cboPhraseTarget, bInit)) Then MakeDirty()
  End Sub


  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  tbProName_TextChanged
  ' Goal :  (1) Process the change immediately in [tdlSettings]
  '         (2) Set dirty flag...
  ' History:
  ' 27-05-2010  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub tbProName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbProName.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objPROed, Me.tbProName, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbProDescr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbProDescr.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objPROed, Me.tbProDescr, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbProPGN_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbProPGN.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objPROed, Me.tbProPGN, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbProVarOE_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbProVarOE.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objPROed, Me.tbProVarOE, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbProVarME_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbProVarME.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objPROed, Me.tbProVarME, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbProVarEmodE_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbProVarEmodE.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objPROed, Me.tbProVarEmodE, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbProVarMBE_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbProVarMBE.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objPROed, Me.tbProVarMBE, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbProNotes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbProNotes.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objPROed, Me.tbProNotes, bInit)) Then MakeDirty()
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  tbNPfeatName_TextChanged
  ' Goal :  (1) Process the change immediately in [tdlSettings]
  '         (2) Set dirty flag...
  ' History:
  ' 02-06-2010  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub tbNPfeatName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbNPfeatName.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objNPed, Me.tbNPfeatName, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbNPfeatDescr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbNPfeatDescr.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objNPed, Me.tbNPfeatDescr, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbNPfeatVariants_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbNPfeatVariants.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objNPed, Me.tbNPfeatVariants, bInit)) Then MakeDirty()
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  tbConsName_TextChanged
  ' Goal :  (1) Process the change immediately in [tdlSettings]
  '         (2) Set dirty flag...
  ' History:
  ' 30-06-2010  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub tbConsName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbConsName.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objCnsEd, Me.tbConsName, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbConsLevel_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbConsLevel.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objCnsEd, Me.tbConsLevel, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbConsMult_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbConsMult.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objCnsEd, Me.tbConsMult, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbConsDescr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbConsDescr.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objCnsEd, Me.tbConsDescr, bInit)) Then MakeDirty()
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  tbCatName_TextChanged
  ' Goal :  (1) Process the change immediately in [tdlSettings]
  '         (2) Set dirty flag...
  ' History:
  ' 04-10-2010  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub tbCatName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbCatName.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objCatEd, Me.tbCatName, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbCatDescr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbCatDescr.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objCatEd, Me.tbCatDescr, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbCatOE_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbCatOE.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objCatEd, Me.tbCatOE, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbCatME_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbCatME.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objCatEd, Me.tbCatME, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbCatEmodE_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbCatEmodE.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objCatEd, Me.tbCatEmodE, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbCatMBE_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbCatMBE.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objCatEd, Me.tbCatMBE, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbCatNotes_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbCatNotes.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objCatEd, Me.tbCatNotes, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbCatType_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbCatType.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objCatEd, Me.tbCatType, bInit)) Then MakeDirty()
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  tbStatusName_TextChanged
  ' Goal :  (1) Process the change immediately in [tdlSettings]
  '         (2) Set dirty flag...
  ' History:
  ' 12-05-2011  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub tbStatusName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbStatusName.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objStaEd, Me.tbStatusName, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbStatusDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbStatusDesc.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objStaEd, Me.tbStatusDesc, bInit)) Then MakeDirty()
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  cmdCellNew_Click
  ' Goal :  Make a new X, where X depends on which tabpage is visible
  ' History:
  ' 15-03-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub cmdCellNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCellNew.Click
    Dim intTemplateId As Integer    ' ID of currently selected template

    Try
      ' Validate
      If (Not objTCed Is Nothing) AndAlso (Not objTmpEd Is Nothing) Then
        ' Get the template id
        intTemplateId = objTmpEd.SelectedId
        ' Double check and execute
        TryAddNew(objTCed, "TemplateCell", "TemplateId", intTemplateId)
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSetting/CellNew error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  cmdCellDel_Click
  ' Goal :  Delete the currently selected cell
  ' History:
  ' 15-03-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub cmdCellDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCellDel.Click
    Try
      ' Double check and execute
      TryRemove(objTCed)
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSetting/CellDel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  cmdCellUp_Click
  ' Goal :  Move this cell one up
  ' History:
  ' 15-03-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub cmdCellUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCellUp.Click
    'Dim intTcellId As Integer   ' ID of the Template Cell

    Try

    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSetting/CellUp error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  cmdCellDown_Click
  ' Goal :  Move this cell one down
  ' History:
  ' 15-03-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub cmdCellDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCellDown.Click
    'Dim intTcellId As Integer   ' ID of the Template Cell

    Try

    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSetting/CellDown error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  cmdMosActNew_Click
  ' Goal :  Make a new X, where X depends on which tabpage is visible
  ' History:
  ' 08-02-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub cmdMosActNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMosActNew.Click
    Dim intMosId As Integer    ' ID of currently selected template

    Try
      ' Validate
      If (Not objMosAcEd Is Nothing) AndAlso (Not objMosEd Is Nothing) Then
        ' Get the MOS id
        intMosId = objMosEd.SelectedId
        ' Double check and execute
        TryAddNewMosAct(objMosAcEd, intMosId)
        ' Put the focus right there where it is needed
        Me.tbMosActOp.Focus()
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSetting/MosActNew error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  cmdMosActDel_Click
  ' Goal :  Delete the currently selected MosAct
  ' History:
  ' 08-02-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub cmdMosActDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMosActDel.Click
    Try
      ' Double check and execute
      TryRemove(objMosAcEd)
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSetting/MosActDel error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  cmdMosActUp_Click
  ' Goal :  Move selected MosAct one up
  ' History:
  ' 08-02-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub cmdMosActUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMosActUp.Click
    Dim intMosId As Integer     ' MosId
    Dim intMosActId As Integer  ' MosActId
    Dim intOrder As Integer     ' Selected order
    Dim dtrFound() As DataRow   ' Result of SELECT
    Dim intI As Integer         ' Counter

    Try
      ' Get the current [MosId] and [MosActId]
      intMosId = objMosEd.SelectedId
      intMosActId = objMosAcEd.SelectedId
      ' Get all the items in order
      dtrFound = tdlSettings.Tables("MosAct").Select("MosId=" & intMosId, "Order ASC")
      ' Walk them all to find me
      For intI = 0 To dtrFound.Length - 1
        ' Check if this is the one
        If (dtrFound(intI).Item("MosActId") = intMosActId) Then
          ' Okay found it -- see what we can do
          If (intI > 0) Then
            ' Yes, we can move up
            intOrder = dtrFound(intI).Item("Order")
            dtrFound(intI).Item("Order") = dtrFound(intI - 1).Item("Order")
            dtrFound(intI - 1).Item("Order") = intOrder
            Exit Sub
          End If
        End If
      Next intI
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSetting/MosActUp error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  cmdMosActDown_Click
  ' Goal :  Move selected MosAct one down
  ' History:
  ' 08-02-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub cmdMosActDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMosActDown.Click
    Dim intMosId As Integer     ' MosId
    Dim intMosActId As Integer  ' MosActId
    Dim intOrder As Integer     ' Selected order
    Dim dtrFound() As DataRow   ' Result of SELECT
    Dim intI As Integer         ' Counter

    Try
      ' Get the current [MosId] and [MosActId]
      intMosId = objMosEd.SelectedId
      intMosActId = objMosAcEd.SelectedId
      ' Get all the items in order
      dtrFound = tdlSettings.Tables("MosAct").Select("MosId=" & intMosId, "Order ASC")
      ' Walk them all to find me
      For intI = 0 To dtrFound.Length - 1
        ' Check if this is the one
        If (dtrFound(intI).Item("MosActId") = intMosActId) Then
          ' Okay found it -- see what we can do
          If (intI < dtrFound.Length - 1) Then
            ' Yes, we can move down
            intOrder = dtrFound(intI).Item("Order")
            dtrFound(intI).Item("Order") = dtrFound(intI + 1).Item("Order")
            dtrFound(intI + 1).Item("Order") = intOrder
            Exit Sub
          End If
        End If
      Next intI
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSetting/MosActDown error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  cmdMosUp_Click
  ' Goal :  Move selected Mos one up
  ' History:
  ' 12-02-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub cmdMosUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMosUp.Click
    Dim intMosId As Integer     ' MosId
    Dim intOrder As Integer     ' Selected order
    Dim dtrFound() As DataRow   ' Result of SELECT
    Dim intI As Integer         ' Counter

    Try
      ' Get the current [MosId] 
      intMosId = objMosEd.SelectedId
      ' Get all the items in order
      dtrFound = tdlSettings.Tables("Mos").Select("", "Order ASC")
      ' Walk them all to find me
      For intI = 0 To dtrFound.Length - 1
        ' Check if this is the one
        If (dtrFound(intI).Item("MosId") = intMosId) Then
          ' Okay found it -- see what we can do
          If (intI > 0) Then
            ' Yes, we can move up
            intOrder = dtrFound(intI).Item("Order")
            dtrFound(intI).Item("Order") = dtrFound(intI - 1).Item("Order")
            dtrFound(intI - 1).Item("Order") = intOrder
            Exit Sub
          End If
        End If
      Next intI
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSetting/MosUp error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  cmdMosDown_Click
  ' Goal :  Move selected Mos one down
  ' History:
  ' 12-02-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub cmdMosDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMosDown.Click
    Dim intMosId As Integer     ' MosId
    Dim intOrder As Integer     ' Selected order
    Dim dtrFound() As DataRow   ' Result of SELECT
    Dim intI As Integer         ' Counter

    Try
      ' Get the current [MosId] 
      intMosId = objMosEd.SelectedId
      ' Get all the items in order
      dtrFound = tdlSettings.Tables("Mos").Select("", "Order ASC")
      ' Walk them all to find me
      For intI = 0 To dtrFound.Length - 1
        ' Check if this is the one
        If (dtrFound(intI).Item("MosId") = intMosId) Then
          ' Okay found it -- see what we can do
          If (intI < dtrFound.Length - 1) Then
            ' Yes, we can move down
            intOrder = dtrFound(intI).Item("Order")
            dtrFound(intI).Item("Order") = dtrFound(intI + 1).Item("Order")
            dtrFound(intI + 1).Item("Order") = intOrder
            Exit Sub
          End If
        End If
      Next intI
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSetting/MosDown error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  tbTemplateName_TextChanged
  ' Goal :  (1) Process the change immediately in [tdlSettings]
  '         (2) Set dirty flag...
  ' History:
  ' 15-03-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub tbTemplateName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbTemplateName.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objTmpEd, Me.tbTemplateName, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbTemplateDescr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbTemplateDescr.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objTmpEd, Me.tbTemplateDescr, bInit)) Then MakeDirty()
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  tbTcellName_TextChanged
  ' Goal :  (1) Process the change immediately in [tdlSettings]
  '         (2) Set dirty flag...
  ' History:
  ' 15-03-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub tbTcellName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbTcellName.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objTCed, Me.tbTcellName, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbTcellContent_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbTcellContent.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objTCed, Me.tbTcellContent, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbTcellEnv_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbTcellEnv.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objTCed, Me.tbTcellEnv, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbTcellDescr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbTcellDescr.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objTCed, Me.tbTcellDescr, bInit)) Then MakeDirty()
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  tbXrelName_TextChanged
  ' Goal :  (1) Process the change immediately in [tdlSettings]
  '         (2) Set dirty flag...
  ' History:
  ' 15-03-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub tbXrelName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbXrelName.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objXrelEd, Me.tbXrelName, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbXrelXname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbXrelXname.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objXrelEd, Me.tbXrelXname, bInit)) Then MakeDirty()
  End Sub
  Private Sub cboXrelType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboXrelType.SelectedIndexChanged
    ' Make sure changes are processed
    If (CtlChanged(objXrelEd, Me.cboXrelType, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbXrelDescr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbXrelDescr.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objXrelEd, Me.tbXrelDescr, bInit)) Then MakeDirty()
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  dgvTemplate_SelectionChanged
  ' Goal :  As soon as we are initialised and a DGV template is selected,
  '           adapt the FILTER for showing/selecting Tcells
  ' History:
  ' 15-03-2012  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub dgvTemplate_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvTemplate.SelectionChanged
    Dim intTemplateId As Integer    ' THe ID of the currently selected template

    Try
      ' Is there a [PF object] already?
      If (Not objTCed Is Nothing) AndAlso (Not objTmpEd Is Nothing) Then
        ' Get the ID of the template
        intTemplateId = objTmpEd.SelectedId
        ' Set the filter for the [Template] section
        objTCed.Filter = "TemplateId=" & intTemplateId
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSetting/dgvTemplate error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  dgvMos_SelectionChanged
  ' Goal :  As soon as we are initialised and a DGV template is selected,
  '           adapt the FILTER for showing/selecting MosActs
  ' History:
  ' 08-02-2013  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub dgvMos_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvMos.SelectionChanged
    Dim intMosId As Integer    ' THe ID of the currently selected template

    Try
      ' Is there a [PF object] already?
      If (Not objMosAcEd Is Nothing) AndAlso (Not objMosEd Is Nothing) Then
        ' Get the ID of the MOS
        intMosId = objMosEd.SelectedId
        ' Validate
        If (intMosId < 0) Then Exit Sub
        ' Set the filter for the [Mos] section
        objMosAcEd.Filter = "MosId=" & intMosId
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("frmSetting/dgvMos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  Private Sub tbPhraseInclude_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbPhraseInclude.TextChanged
    MakeDirty()
  End Sub

  Private Sub tbPhraseExclude_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbPhraseExclude.TextChanged
    MakeDirty()
  End Sub


  Private Sub tbMosName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbMosName.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objMosEd, Me.tbMosName, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbMosTrigger_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbMosTrigger.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objMosEd, Me.tbMosTrigger, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbMosCond_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbMosCond.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objMosEd, Me.tbMosCond, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbMosActOp_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbMosActOp.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objMosAcEd, Me.tbMosActOp, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbMosActArg_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbMosActArg.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objMosAcEd, Me.tbMosActArg, bInit)) Then MakeDirty()
  End Sub
  Private Sub tbMosActDescr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbMosActDescr.TextChanged
    ' Make sure changes are processed
    If (CtlChanged(objMosAcEd, Me.tbMosActDescr, bInit)) Then MakeDirty()
  End Sub
  Private Sub cboMosActDir_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboMosActDir.SelectedIndexChanged
    ' Make sure changes are processed
    If (CtlChanged(objMosAcEd, Me.cboMosActDir, bInit)) Then MakeDirty()
  End Sub


End Class