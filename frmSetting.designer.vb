<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSetting
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSetting))
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.TabControl1 = New System.Windows.Forms.TabControl()
    Me.tbGeneral = New System.Windows.Forms.TabPage()
    Me.cmdCorpusStudio = New System.Windows.Forms.Button()
    Me.tbCorpusStudio = New System.Windows.Forms.TextBox()
    Me.Label57 = New System.Windows.Forms.Label()
    Me.cmdChainDict = New System.Windows.Forms.Button()
    Me.tbChainDict = New System.Windows.Forms.TextBox()
    Me.Label42 = New System.Windows.Forms.Label()
    Me.cmdPeriodDef = New System.Windows.Forms.Button()
    Me.tbPeriodDef = New System.Windows.Forms.TextBox()
    Me.Label27 = New System.Windows.Forms.Label()
    Me.GroupBox3 = New System.Windows.Forms.GroupBox()
    Me.chbUpdStartup = New System.Windows.Forms.CheckBox()
    Me.chbGenActionInet = New System.Windows.Forms.CheckBox()
    Me.GroupBox1 = New System.Windows.Forms.GroupBox()
    Me.chbStatusAutoNotes = New System.Windows.Forms.CheckBox()
    Me.chbStatusAutoChanged = New System.Windows.Forms.CheckBox()
    Me.tbLogBase = New System.Windows.Forms.TextBox()
    Me.Label61 = New System.Windows.Forms.Label()
    Me.chbCheckCns = New System.Windows.Forms.CheckBox()
    Me.chbCheckCat = New System.Windows.Forms.CheckBox()
    Me.tbAbsMaxIPdist = New System.Windows.Forms.TextBox()
    Me.Label45 = New System.Windows.Forms.Label()
    Me.chbShowAllAnt = New System.Windows.Forms.CheckBox()
    Me.tbProfile = New System.Windows.Forms.TextBox()
    Me.Label44 = New System.Windows.Forms.Label()
    Me.tbMaxIPdist = New System.Windows.Forms.TextBox()
    Me.Label43 = New System.Windows.Forms.Label()
    Me.tbUserName = New System.Windows.Forms.TextBox()
    Me.Label41 = New System.Windows.Forms.Label()
    Me.tbDriveChange = New System.Windows.Forms.TextBox()
    Me.Label14 = New System.Windows.Forms.Label()
    Me.chbUserProPGN = New System.Windows.Forms.CheckBox()
    Me.chbDebugging = New System.Windows.Forms.CheckBox()
    Me.chbShowCODE = New System.Windows.Forms.CheckBox()
    Me.cmdWorkDir = New System.Windows.Forms.Button()
    Me.tbWorkDir = New System.Windows.Forms.TextBox()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.tpCoref = New System.Windows.Forms.TabPage()
    Me.cboColor = New System.Windows.Forms.ComboBox()
    Me.Label11 = New System.Windows.Forms.Label()
    Me.tbCorefName = New System.Windows.Forms.TextBox()
    Me.Label5 = New System.Windows.Forms.Label()
    Me.tbCorefDescr = New System.Windows.Forms.RichTextBox()
    Me.Label4 = New System.Windows.Forms.Label()
    Me.dgvCorefType = New System.Windows.Forms.DataGridView()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.tpPhrase = New System.Windows.Forms.TabPage()
    Me.SplitContainer8 = New System.Windows.Forms.SplitContainer()
    Me.SplitContainer9 = New System.Windows.Forms.SplitContainer()
    Me.tbPhraseInclude = New System.Windows.Forms.RichTextBox()
    Me.Label71 = New System.Windows.Forms.Label()
    Me.tbPhraseExclude = New System.Windows.Forms.RichTextBox()
    Me.Label72 = New System.Windows.Forms.Label()
    Me.SplitContainer7 = New System.Windows.Forms.SplitContainer()
    Me.dgvPhraseType = New System.Windows.Forms.DataGridView()
    Me.Label7 = New System.Windows.Forms.Label()
    Me.tbPhraseName = New System.Windows.Forms.TextBox()
    Me.Label70 = New System.Windows.Forms.Label()
    Me.tbPhraseNode = New System.Windows.Forms.TextBox()
    Me.Label10 = New System.Windows.Forms.Label()
    Me.Label6 = New System.Windows.Forms.Label()
    Me.cboPhraseTarget = New System.Windows.Forms.ComboBox()
    Me.Label8 = New System.Windows.Forms.Label()
    Me.cboPhraseType = New System.Windows.Forms.ComboBox()
    Me.tbPhraseChild = New System.Windows.Forms.TextBox()
    Me.Label9 = New System.Windows.Forms.Label()
    Me.tpPronoun = New System.Windows.Forms.TabPage()
    Me.Label26 = New System.Windows.Forms.Label()
    Me.tbProVarMBE = New System.Windows.Forms.RichTextBox()
    Me.Label25 = New System.Windows.Forms.Label()
    Me.tbProVarEmodE = New System.Windows.Forms.RichTextBox()
    Me.Label24 = New System.Windows.Forms.Label()
    Me.tbProVarME = New System.Windows.Forms.RichTextBox()
    Me.Label23 = New System.Windows.Forms.Label()
    Me.tbProNotes = New System.Windows.Forms.RichTextBox()
    Me.Label22 = New System.Windows.Forms.Label()
    Me.tbProPGN = New System.Windows.Forms.TextBox()
    Me.Label21 = New System.Windows.Forms.Label()
    Me.tbProVarOE = New System.Windows.Forms.RichTextBox()
    Me.Label16 = New System.Windows.Forms.Label()
    Me.tbProDescr = New System.Windows.Forms.TextBox()
    Me.Label15 = New System.Windows.Forms.Label()
    Me.tbProName = New System.Windows.Forms.TextBox()
    Me.Label13 = New System.Windows.Forms.Label()
    Me.dgvPronoun = New System.Windows.Forms.DataGridView()
    Me.Label12 = New System.Windows.Forms.Label()
    Me.tpNPfeat = New System.Windows.Forms.TabPage()
    Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
    Me.dgvNPfeat = New System.Windows.Forms.DataGridView()
    Me.Label28 = New System.Windows.Forms.Label()
    Me.tbNPfeatVariants = New System.Windows.Forms.RichTextBox()
    Me.Label31 = New System.Windows.Forms.Label()
    Me.tbNPfeatDescr = New System.Windows.Forms.TextBox()
    Me.Label30 = New System.Windows.Forms.Label()
    Me.tbNPfeatName = New System.Windows.Forms.TextBox()
    Me.Label29 = New System.Windows.Forms.Label()
    Me.tpCons = New System.Windows.Forms.TabPage()
    Me.SplitContainer3 = New System.Windows.Forms.SplitContainer()
    Me.dgvCons = New System.Windows.Forms.DataGridView()
    Me.Label36 = New System.Windows.Forms.Label()
    Me.cmdCnsDown = New System.Windows.Forms.Button()
    Me.cmdCnsUp = New System.Windows.Forms.Button()
    Me.tbConsDescr = New System.Windows.Forms.RichTextBox()
    Me.Label40 = New System.Windows.Forms.Label()
    Me.tbConsMult = New System.Windows.Forms.TextBox()
    Me.Label39 = New System.Windows.Forms.Label()
    Me.tbConsLevel = New System.Windows.Forms.TextBox()
    Me.Label38 = New System.Windows.Forms.Label()
    Me.tbConsName = New System.Windows.Forms.TextBox()
    Me.Label37 = New System.Windows.Forms.Label()
    Me.tpCat = New System.Windows.Forms.TabPage()
    Me.Label56 = New System.Windows.Forms.Label()
    Me.tbCatType = New System.Windows.Forms.TextBox()
    Me.Label55 = New System.Windows.Forms.Label()
    Me.Label47 = New System.Windows.Forms.Label()
    Me.tbCatMBE = New System.Windows.Forms.RichTextBox()
    Me.Label48 = New System.Windows.Forms.Label()
    Me.tbCatEmodE = New System.Windows.Forms.RichTextBox()
    Me.Label49 = New System.Windows.Forms.Label()
    Me.tbCatME = New System.Windows.Forms.RichTextBox()
    Me.Label50 = New System.Windows.Forms.Label()
    Me.tbCatNotes = New System.Windows.Forms.RichTextBox()
    Me.Label51 = New System.Windows.Forms.Label()
    Me.tbCatOE = New System.Windows.Forms.RichTextBox()
    Me.Label52 = New System.Windows.Forms.Label()
    Me.tbCatDescr = New System.Windows.Forms.TextBox()
    Me.Label53 = New System.Windows.Forms.Label()
    Me.tbCatName = New System.Windows.Forms.TextBox()
    Me.Label54 = New System.Windows.Forms.Label()
    Me.dgvCategory = New System.Windows.Forms.DataGridView()
    Me.Label46 = New System.Windows.Forms.Label()
    Me.tpStatus = New System.Windows.Forms.TabPage()
    Me.SplitContainer4 = New System.Windows.Forms.SplitContainer()
    Me.dgvStatus = New System.Windows.Forms.DataGridView()
    Me.Label58 = New System.Windows.Forms.Label()
    Me.tbStatusDesc = New System.Windows.Forms.RichTextBox()
    Me.Label59 = New System.Windows.Forms.Label()
    Me.tbStatusName = New System.Windows.Forms.TextBox()
    Me.Label60 = New System.Windows.Forms.Label()
    Me.tpCharting = New System.Windows.Forms.TabPage()
    Me.SplitContainer5 = New System.Windows.Forms.SplitContainer()
    Me.dgvTemplate = New System.Windows.Forms.DataGridView()
    Me.Label62 = New System.Windows.Forms.Label()
    Me.tbTemplateDescr = New System.Windows.Forms.RichTextBox()
    Me.Label69 = New System.Windows.Forms.Label()
    Me.SplitContainer6 = New System.Windows.Forms.SplitContainer()
    Me.dgvTcell = New System.Windows.Forms.DataGridView()
    Me.cmdCellDel = New System.Windows.Forms.Button()
    Me.cmdCellNew = New System.Windows.Forms.Button()
    Me.cmdCellDown = New System.Windows.Forms.Button()
    Me.cmdCellUp = New System.Windows.Forms.Button()
    Me.Label68 = New System.Windows.Forms.Label()
    Me.tbTcellEnv = New System.Windows.Forms.TextBox()
    Me.tbTcellDescr = New System.Windows.Forms.RichTextBox()
    Me.Label67 = New System.Windows.Forms.Label()
    Me.tbTcellContent = New System.Windows.Forms.RichTextBox()
    Me.Label66 = New System.Windows.Forms.Label()
    Me.tbTcellId = New System.Windows.Forms.TextBox()
    Me.Label65 = New System.Windows.Forms.Label()
    Me.tbTcellName = New System.Windows.Forms.TextBox()
    Me.Label64 = New System.Windows.Forms.Label()
    Me.tbTemplateName = New System.Windows.Forms.TextBox()
    Me.Label63 = New System.Windows.Forms.Label()
    Me.tpXrel = New System.Windows.Forms.TabPage()
    Me.SplitContainer10 = New System.Windows.Forms.SplitContainer()
    Me.Label77 = New System.Windows.Forms.Label()
    Me.dgvXrel = New System.Windows.Forms.DataGridView()
    Me.Label81 = New System.Windows.Forms.Label()
    Me.tbXrelXname = New System.Windows.Forms.RichTextBox()
    Me.cboXrelType = New System.Windows.Forms.ComboBox()
    Me.tbXrelDescr = New System.Windows.Forms.RichTextBox()
    Me.Label78 = New System.Windows.Forms.Label()
    Me.Label79 = New System.Windows.Forms.Label()
    Me.tbXrelName = New System.Windows.Forms.TextBox()
    Me.Label80 = New System.Windows.Forms.Label()
    Me.tpMos = New System.Windows.Forms.TabPage()
    Me.SplitContainer11 = New System.Windows.Forms.SplitContainer()
    Me.cmdMosDown = New System.Windows.Forms.Button()
    Me.cmdMosUp = New System.Windows.Forms.Button()
    Me.dgvMos = New System.Windows.Forms.DataGridView()
    Me.Label82 = New System.Windows.Forms.Label()
    Me.SplitContainer12 = New System.Windows.Forms.SplitContainer()
    Me.tbMosName = New System.Windows.Forms.TextBox()
    Me.Label90 = New System.Windows.Forms.Label()
    Me.tbMosCond = New System.Windows.Forms.RichTextBox()
    Me.Label83 = New System.Windows.Forms.Label()
    Me.tbMosTrigger = New System.Windows.Forms.TextBox()
    Me.Label84 = New System.Windows.Forms.Label()
    Me.SplitContainer13 = New System.Windows.Forms.SplitContainer()
    Me.dgvMosAct = New System.Windows.Forms.DataGridView()
    Me.cboMosActDir = New System.Windows.Forms.ComboBox()
    Me.Label86 = New System.Windows.Forms.Label()
    Me.cmdMosActDel = New System.Windows.Forms.Button()
    Me.cmdMosActNew = New System.Windows.Forms.Button()
    Me.cmdMosActDown = New System.Windows.Forms.Button()
    Me.cmdMosActUp = New System.Windows.Forms.Button()
    Me.Label85 = New System.Windows.Forms.Label()
    Me.tbMosActDescr = New System.Windows.Forms.RichTextBox()
    Me.tbMosActArg = New System.Windows.Forms.RichTextBox()
    Me.Label87 = New System.Windows.Forms.Label()
    Me.tbMosActId = New System.Windows.Forms.TextBox()
    Me.Label88 = New System.Windows.Forms.Label()
    Me.tbMosActOp = New System.Windows.Forms.TextBox()
    Me.Label89 = New System.Windows.Forms.Label()
    Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
    Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.cmdOk = New System.Windows.Forms.Button()
    Me.cmdSetFile = New System.Windows.Forms.Button()
    Me.tbSettingsFile = New System.Windows.Forms.TextBox()
    Me.Label3 = New System.Windows.Forms.Label()
    Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
    Me.Label17 = New System.Windows.Forms.Label()
    Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
    Me.Label18 = New System.Windows.Forms.Label()
    Me.TextBox1 = New System.Windows.Forms.TextBox()
    Me.TextBox2 = New System.Windows.Forms.TextBox()
    Me.Label19 = New System.Windows.Forms.Label()
    Me.DataGridView1 = New System.Windows.Forms.DataGridView()
    Me.Label20 = New System.Windows.Forms.Label()
    Me.cmdNew = New System.Windows.Forms.Button()
    Me.cmdDel = New System.Windows.Forms.Button()
    Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
    Me.DataGridView2 = New System.Windows.Forms.DataGridView()
    Me.Label32 = New System.Windows.Forms.Label()
    Me.RichTextBox2 = New System.Windows.Forms.RichTextBox()
    Me.Label33 = New System.Windows.Forms.Label()
    Me.TextBox3 = New System.Windows.Forms.TextBox()
    Me.Label34 = New System.Windows.Forms.Label()
    Me.TextBox4 = New System.Windows.Forms.TextBox()
    Me.Label35 = New System.Windows.Forms.Label()
    Me.bsProjType = New System.Windows.Forms.BindingSource(Me.components)
    Me.bsProjEdit = New System.Windows.Forms.BindingSource(Me.components)
    Me.DataGridView3 = New System.Windows.Forms.DataGridView()
    Me.Label73 = New System.Windows.Forms.Label()
    Me.RichTextBox3 = New System.Windows.Forms.RichTextBox()
    Me.Label74 = New System.Windows.Forms.Label()
    Me.TextBox5 = New System.Windows.Forms.TextBox()
    Me.Label75 = New System.Windows.Forms.Label()
    Me.TextBox6 = New System.Windows.Forms.TextBox()
    Me.Label76 = New System.Windows.Forms.Label()
    Me.StatusStrip1.SuspendLayout()
    Me.TabControl1.SuspendLayout()
    Me.tbGeneral.SuspendLayout()
    Me.GroupBox3.SuspendLayout()
    Me.GroupBox1.SuspendLayout()
    Me.tpCoref.SuspendLayout()
    CType(Me.dgvCorefType, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.tpPhrase.SuspendLayout()
    CType(Me.SplitContainer8, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer8.Panel1.SuspendLayout()
    Me.SplitContainer8.Panel2.SuspendLayout()
    Me.SplitContainer8.SuspendLayout()
    CType(Me.SplitContainer9, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer9.Panel1.SuspendLayout()
    Me.SplitContainer9.Panel2.SuspendLayout()
    Me.SplitContainer9.SuspendLayout()
    CType(Me.SplitContainer7, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer7.Panel1.SuspendLayout()
    Me.SplitContainer7.Panel2.SuspendLayout()
    Me.SplitContainer7.SuspendLayout()
    CType(Me.dgvPhraseType, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.tpPronoun.SuspendLayout()
    CType(Me.dgvPronoun, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.tpNPfeat.SuspendLayout()
    CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer1.Panel1.SuspendLayout()
    Me.SplitContainer1.Panel2.SuspendLayout()
    Me.SplitContainer1.SuspendLayout()
    CType(Me.dgvNPfeat, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.tpCons.SuspendLayout()
    CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer3.Panel1.SuspendLayout()
    Me.SplitContainer3.Panel2.SuspendLayout()
    Me.SplitContainer3.SuspendLayout()
    CType(Me.dgvCons, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.tpCat.SuspendLayout()
    CType(Me.dgvCategory, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.tpStatus.SuspendLayout()
    CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer4.Panel1.SuspendLayout()
    Me.SplitContainer4.Panel2.SuspendLayout()
    Me.SplitContainer4.SuspendLayout()
    CType(Me.dgvStatus, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.tpCharting.SuspendLayout()
    CType(Me.SplitContainer5, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer5.Panel1.SuspendLayout()
    Me.SplitContainer5.Panel2.SuspendLayout()
    Me.SplitContainer5.SuspendLayout()
    CType(Me.dgvTemplate, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.SplitContainer6, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer6.Panel1.SuspendLayout()
    Me.SplitContainer6.Panel2.SuspendLayout()
    Me.SplitContainer6.SuspendLayout()
    CType(Me.dgvTcell, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.tpXrel.SuspendLayout()
    CType(Me.SplitContainer10, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer10.Panel1.SuspendLayout()
    Me.SplitContainer10.Panel2.SuspendLayout()
    Me.SplitContainer10.SuspendLayout()
    CType(Me.dgvXrel, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.tpMos.SuspendLayout()
    CType(Me.SplitContainer11, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer11.Panel1.SuspendLayout()
    Me.SplitContainer11.Panel2.SuspendLayout()
    Me.SplitContainer11.SuspendLayout()
    CType(Me.dgvMos, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.SplitContainer12, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer12.Panel1.SuspendLayout()
    Me.SplitContainer12.Panel2.SuspendLayout()
    Me.SplitContainer12.SuspendLayout()
    CType(Me.SplitContainer13, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer13.Panel1.SuspendLayout()
    Me.SplitContainer13.Panel2.SuspendLayout()
    Me.SplitContainer13.SuspendLayout()
    CType(Me.dgvMosAct, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SplitContainer2.Panel1.SuspendLayout()
    Me.SplitContainer2.Panel2.SuspendLayout()
    Me.SplitContainer2.SuspendLayout()
    CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.bsProjType, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.bsProjEdit, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 514)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(740, 22)
    Me.StatusStrip1.TabIndex = 0
    Me.StatusStrip1.Text = "StatusStrip1"
    '
    'ToolStripStatusLabel1
    '
    Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
    Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(19, 17)
    Me.ToolStripStatusLabel1.Text = "..."
    '
    'ToolStripProgressBar1
    '
    Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
    Me.ToolStripProgressBar1.Size = New System.Drawing.Size(300, 16)
    '
    'Timer1
    '
    '
    'TabControl1
    '
    Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.TabControl1.Controls.Add(Me.tbGeneral)
    Me.TabControl1.Controls.Add(Me.tpCoref)
    Me.TabControl1.Controls.Add(Me.tpPhrase)
    Me.TabControl1.Controls.Add(Me.tpPronoun)
    Me.TabControl1.Controls.Add(Me.tpNPfeat)
    Me.TabControl1.Controls.Add(Me.tpCons)
    Me.TabControl1.Controls.Add(Me.tpCat)
    Me.TabControl1.Controls.Add(Me.tpStatus)
    Me.TabControl1.Controls.Add(Me.tpCharting)
    Me.TabControl1.Controls.Add(Me.tpXrel)
    Me.TabControl1.Controls.Add(Me.tpMos)
    Me.TabControl1.Location = New System.Drawing.Point(0, 40)
    Me.TabControl1.Name = "TabControl1"
    Me.TabControl1.SelectedIndex = 0
    Me.TabControl1.Size = New System.Drawing.Size(740, 440)
    Me.TabControl1.TabIndex = 1
    '
    'tbGeneral
    '
    Me.tbGeneral.Controls.Add(Me.cmdCorpusStudio)
    Me.tbGeneral.Controls.Add(Me.tbCorpusStudio)
    Me.tbGeneral.Controls.Add(Me.Label57)
    Me.tbGeneral.Controls.Add(Me.cmdChainDict)
    Me.tbGeneral.Controls.Add(Me.tbChainDict)
    Me.tbGeneral.Controls.Add(Me.Label42)
    Me.tbGeneral.Controls.Add(Me.cmdPeriodDef)
    Me.tbGeneral.Controls.Add(Me.tbPeriodDef)
    Me.tbGeneral.Controls.Add(Me.Label27)
    Me.tbGeneral.Controls.Add(Me.GroupBox3)
    Me.tbGeneral.Controls.Add(Me.cmdWorkDir)
    Me.tbGeneral.Controls.Add(Me.tbWorkDir)
    Me.tbGeneral.Controls.Add(Me.Label1)
    Me.tbGeneral.Location = New System.Drawing.Point(4, 22)
    Me.tbGeneral.Name = "tbGeneral"
    Me.tbGeneral.Padding = New System.Windows.Forms.Padding(3)
    Me.tbGeneral.Size = New System.Drawing.Size(732, 414)
    Me.tbGeneral.TabIndex = 0
    Me.tbGeneral.Text = "General"
    Me.tbGeneral.UseVisualStyleBackColor = True
    '
    'cmdCorpusStudio
    '
    Me.cmdCorpusStudio.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCorpusStudio.Location = New System.Drawing.Point(697, 87)
    Me.cmdCorpusStudio.Name = "cmdCorpusStudio"
    Me.cmdCorpusStudio.Size = New System.Drawing.Size(27, 20)
    Me.cmdCorpusStudio.TabIndex = 27
    Me.cmdCorpusStudio.Text = "..."
    Me.cmdCorpusStudio.UseVisualStyleBackColor = True
    '
    'tbCorpusStudio
    '
    Me.tbCorpusStudio.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbCorpusStudio.Location = New System.Drawing.Point(119, 87)
    Me.tbCorpusStudio.Name = "tbCorpusStudio"
    Me.tbCorpusStudio.Size = New System.Drawing.Size(572, 20)
    Me.tbCorpusStudio.TabIndex = 26
    '
    'Label57
    '
    Me.Label57.AutoSize = True
    Me.Label57.Location = New System.Drawing.Point(8, 90)
    Me.Label57.Name = "Label57"
    Me.Label57.Size = New System.Drawing.Size(113, 13)
    Me.Label57.TabIndex = 25
    Me.Label57.Text = "CorpusStudio directory"
    '
    'cmdChainDict
    '
    Me.cmdChainDict.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdChainDict.Location = New System.Drawing.Point(697, 61)
    Me.cmdChainDict.Name = "cmdChainDict"
    Me.cmdChainDict.Size = New System.Drawing.Size(27, 20)
    Me.cmdChainDict.TabIndex = 24
    Me.cmdChainDict.Text = "..."
    Me.cmdChainDict.UseVisualStyleBackColor = True
    '
    'tbChainDict
    '
    Me.tbChainDict.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbChainDict.Location = New System.Drawing.Point(112, 61)
    Me.tbChainDict.Name = "tbChainDict"
    Me.tbChainDict.Size = New System.Drawing.Size(579, 20)
    Me.tbChainDict.TabIndex = 23
    '
    'Label42
    '
    Me.Label42.AutoSize = True
    Me.Label42.Location = New System.Drawing.Point(8, 64)
    Me.Label42.Name = "Label42"
    Me.Label42.Size = New System.Drawing.Size(82, 13)
    Me.Label42.TabIndex = 22
    Me.Label42.Text = "Chain dictionary"
    '
    'cmdPeriodDef
    '
    Me.cmdPeriodDef.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdPeriodDef.Location = New System.Drawing.Point(697, 35)
    Me.cmdPeriodDef.Name = "cmdPeriodDef"
    Me.cmdPeriodDef.Size = New System.Drawing.Size(27, 20)
    Me.cmdPeriodDef.TabIndex = 21
    Me.cmdPeriodDef.Text = "..."
    Me.cmdPeriodDef.UseVisualStyleBackColor = True
    '
    'tbPeriodDef
    '
    Me.tbPeriodDef.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbPeriodDef.Location = New System.Drawing.Point(112, 35)
    Me.tbPeriodDef.Name = "tbPeriodDef"
    Me.tbPeriodDef.Size = New System.Drawing.Size(579, 20)
    Me.tbPeriodDef.TabIndex = 20
    '
    'Label27
    '
    Me.Label27.AutoSize = True
    Me.Label27.Location = New System.Drawing.Point(8, 38)
    Me.Label27.Name = "Label27"
    Me.Label27.Size = New System.Drawing.Size(98, 13)
    Me.Label27.TabIndex = 19
    Me.Label27.Text = "Period definition file"
    '
    'GroupBox3
    '
    Me.GroupBox3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.GroupBox3.Controls.Add(Me.chbUpdStartup)
    Me.GroupBox3.Controls.Add(Me.chbGenActionInet)
    Me.GroupBox3.Controls.Add(Me.GroupBox1)
    Me.GroupBox3.Controls.Add(Me.tbLogBase)
    Me.GroupBox3.Controls.Add(Me.Label61)
    Me.GroupBox3.Controls.Add(Me.chbCheckCns)
    Me.GroupBox3.Controls.Add(Me.chbCheckCat)
    Me.GroupBox3.Controls.Add(Me.tbAbsMaxIPdist)
    Me.GroupBox3.Controls.Add(Me.Label45)
    Me.GroupBox3.Controls.Add(Me.chbShowAllAnt)
    Me.GroupBox3.Controls.Add(Me.tbProfile)
    Me.GroupBox3.Controls.Add(Me.Label44)
    Me.GroupBox3.Controls.Add(Me.tbMaxIPdist)
    Me.GroupBox3.Controls.Add(Me.Label43)
    Me.GroupBox3.Controls.Add(Me.tbUserName)
    Me.GroupBox3.Controls.Add(Me.Label41)
    Me.GroupBox3.Controls.Add(Me.tbDriveChange)
    Me.GroupBox3.Controls.Add(Me.Label14)
    Me.GroupBox3.Controls.Add(Me.chbUserProPGN)
    Me.GroupBox3.Controls.Add(Me.chbDebugging)
    Me.GroupBox3.Controls.Add(Me.chbShowCODE)
    Me.GroupBox3.Location = New System.Drawing.Point(11, 113)
    Me.GroupBox3.Name = "GroupBox3"
    Me.GroupBox3.Size = New System.Drawing.Size(713, 295)
    Me.GroupBox3.TabIndex = 18
    Me.GroupBox3.TabStop = False
    Me.GroupBox3.Text = "Preferences"
    '
    'chbUpdStartup
    '
    Me.chbUpdStartup.AutoSize = True
    Me.chbUpdStartup.Location = New System.Drawing.Point(417, 157)
    Me.chbUpdStartup.Name = "chbUpdStartup"
    Me.chbUpdStartup.Size = New System.Drawing.Size(163, 17)
    Me.chbUpdStartup.TabIndex = 37
    Me.chbUpdStartup.Text = "Check for updates on startup"
    Me.chbUpdStartup.UseVisualStyleBackColor = True
    '
    'chbGenActionInet
    '
    Me.chbGenActionInet.AutoSize = True
    Me.chbGenActionInet.Location = New System.Drawing.Point(417, 134)
    Me.chbGenActionInet.Name = "chbGenActionInet"
    Me.chbGenActionInet.Size = New System.Drawing.Size(81, 17)
    Me.chbGenActionInet.TabIndex = 36
    Me.chbGenActionInet.Text = "Log actions"
    Me.chbGenActionInet.UseVisualStyleBackColor = True
    '
    'GroupBox1
    '
    Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GroupBox1.Controls.Add(Me.chbStatusAutoNotes)
    Me.GroupBox1.Controls.Add(Me.chbStatusAutoChanged)
    Me.GroupBox1.Location = New System.Drawing.Point(405, 42)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(302, 78)
    Me.GroupBox1.TabIndex = 35
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Corpus database processing"
    '
    'chbStatusAutoNotes
    '
    Me.chbStatusAutoNotes.AutoSize = True
    Me.chbStatusAutoNotes.Location = New System.Drawing.Point(12, 46)
    Me.chbStatusAutoNotes.Name = "chbStatusAutoNotes"
    Me.chbStatusAutoNotes.Size = New System.Drawing.Size(206, 17)
    Me.chbStatusAutoNotes.TabIndex = 35
    Me.chbStatusAutoNotes.Text = "Add Date/Time in [Notes] on changes"
    Me.chbStatusAutoNotes.UseVisualStyleBackColor = True
    '
    'chbStatusAutoChanged
    '
    Me.chbStatusAutoChanged.AutoSize = True
    Me.chbStatusAutoChanged.Location = New System.Drawing.Point(12, 19)
    Me.chbStatusAutoChanged.Name = "chbStatusAutoChanged"
    Me.chbStatusAutoChanged.Size = New System.Drawing.Size(184, 17)
    Me.chbStatusAutoChanged.TabIndex = 34
    Me.chbStatusAutoChanged.Text = "Set status [Changed] on changes"
    Me.chbStatusAutoChanged.UseVisualStyleBackColor = True
    '
    'tbLogBase
    '
    Me.tbLogBase.Location = New System.Drawing.Point(580, 16)
    Me.tbLogBase.Name = "tbLogBase"
    Me.tbLogBase.Size = New System.Drawing.Size(100, 20)
    Me.tbLogBase.TabIndex = 33
    '
    'Label61
    '
    Me.Label61.AutoSize = True
    Me.Label61.Location = New System.Drawing.Point(414, 20)
    Me.Label61.Name = "Label61"
    Me.Label61.Size = New System.Drawing.Size(164, 13)
    Me.Label61.TabIndex = 32
    Me.Label61.Text = "Log base for coreferential chains:"
    '
    'chbCheckCns
    '
    Me.chbCheckCns.AutoSize = True
    Me.chbCheckCns.Location = New System.Drawing.Point(6, 134)
    Me.chbCheckCns.Name = "chbCheckCns"
    Me.chbCheckCns.Size = New System.Drawing.Size(282, 17)
    Me.chbCheckCns.TabIndex = 31
    Me.chbCheckCns.Text = "Check the constraint ranking in new versions of Cesax"
    Me.chbCheckCns.UseVisualStyleBackColor = True
    '
    'chbCheckCat
    '
    Me.chbCheckCat.AutoSize = True
    Me.chbCheckCat.Location = New System.Drawing.Point(6, 112)
    Me.chbCheckCat.Name = "chbCheckCat"
    Me.chbCheckCat.Size = New System.Drawing.Size(267, 17)
    Me.chbCheckCat.TabIndex = 30
    Me.chbCheckCat.Text = "Check for new categories in new versions of Cesax"
    Me.chbCheckCat.UseVisualStyleBackColor = True
    '
    'tbAbsMaxIPdist
    '
    Me.tbAbsMaxIPdist.Location = New System.Drawing.Point(303, 190)
    Me.tbAbsMaxIPdist.Name = "tbAbsMaxIPdist"
    Me.tbAbsMaxIPdist.Size = New System.Drawing.Size(96, 20)
    Me.tbAbsMaxIPdist.TabIndex = 29
    '
    'Label45
    '
    Me.Label45.AutoSize = True
    Me.Label45.Location = New System.Drawing.Point(6, 193)
    Me.Label45.Name = "Label45"
    Me.Label45.Size = New System.Drawing.Size(150, 13)
    Me.Label45.TabIndex = 28
    Me.Label45.Text = "Absolute maximum IP distance"
    '
    'chbShowAllAnt
    '
    Me.chbShowAllAnt.AutoSize = True
    Me.chbShowAllAnt.Location = New System.Drawing.Point(6, 88)
    Me.chbShowAllAnt.Name = "chbShowAllAnt"
    Me.chbShowAllAnt.Size = New System.Drawing.Size(293, 17)
    Me.chbShowAllAnt.TabIndex = 27
    Me.chbShowAllAnt.Text = "Show the whole antecedent stack -- no matter how large"
    Me.chbShowAllAnt.UseVisualStyleBackColor = True
    '
    'tbProfile
    '
    Me.tbProfile.Location = New System.Drawing.Point(303, 216)
    Me.tbProfile.Name = "tbProfile"
    Me.tbProfile.Size = New System.Drawing.Size(96, 20)
    Me.tbProfile.TabIndex = 26
    '
    'Label44
    '
    Me.Label44.AutoSize = True
    Me.Label44.Location = New System.Drawing.Point(6, 219)
    Me.Label44.Name = "Label44"
    Me.Label44.Size = New System.Drawing.Size(184, 13)
    Me.Label44.TabIndex = 25
    Me.Label44.Text = "Profile depth level (zero for no profile):"
    '
    'tbMaxIPdist
    '
    Me.tbMaxIPdist.Location = New System.Drawing.Point(303, 164)
    Me.tbMaxIPdist.Name = "tbMaxIPdist"
    Me.tbMaxIPdist.Size = New System.Drawing.Size(96, 20)
    Me.tbMaxIPdist.TabIndex = 24
    '
    'Label43
    '
    Me.Label43.AutoSize = True
    Me.Label43.Location = New System.Drawing.Point(6, 167)
    Me.Label43.Name = "Label43"
    Me.Label43.Size = New System.Drawing.Size(256, 13)
    Me.Label43.TabIndex = 23
    Me.Label43.Text = "Maximum IP distance in antecedent stack (e.g. Auto)"
    '
    'tbUserName
    '
    Me.tbUserName.Location = New System.Drawing.Point(223, 242)
    Me.tbUserName.Name = "tbUserName"
    Me.tbUserName.Size = New System.Drawing.Size(176, 20)
    Me.tbUserName.TabIndex = 22
    '
    'Label41
    '
    Me.Label41.AutoSize = True
    Me.Label41.Location = New System.Drawing.Point(6, 245)
    Me.Label41.Name = "Label41"
    Me.Label41.Size = New System.Drawing.Size(61, 13)
    Me.Label41.TabIndex = 21
    Me.Label41.Text = "User name:"
    '
    'tbDriveChange
    '
    Me.tbDriveChange.Location = New System.Drawing.Point(223, 268)
    Me.tbDriveChange.Name = "tbDriveChange"
    Me.tbDriveChange.Size = New System.Drawing.Size(176, 20)
    Me.tbDriveChange.TabIndex = 4
    '
    'Label14
    '
    Me.Label14.AutoSize = True
    Me.Label14.Location = New System.Drawing.Point(6, 271)
    Me.Label14.Name = "Label14"
    Me.Label14.Size = New System.Drawing.Size(133, 13)
    Me.Label14.TabIndex = 3
    Me.Label14.Text = "Automatic drive change(s):"
    '
    'chbUserProPGN
    '
    Me.chbUserProPGN.AutoSize = True
    Me.chbUserProPGN.Location = New System.Drawing.Point(6, 65)
    Me.chbUserProPGN.Name = "chbUserProPGN"
    Me.chbUserProPGN.Size = New System.Drawing.Size(306, 17)
    Me.chbUserProPGN.TabIndex = 20
    Me.chbUserProPGN.Text = "Ask user to clarify the PGN for a pro/dem when it is needed"
    Me.chbUserProPGN.UseVisualStyleBackColor = True
    '
    'chbDebugging
    '
    Me.chbDebugging.AutoSize = True
    Me.chbDebugging.Location = New System.Drawing.Point(6, 19)
    Me.chbDebugging.Name = "chbDebugging"
    Me.chbDebugging.Size = New System.Drawing.Size(78, 17)
    Me.chbDebugging.TabIndex = 19
    Me.chbDebugging.Text = "Debugging"
    Me.chbDebugging.UseVisualStyleBackColor = True
    '
    'chbShowCODE
    '
    Me.chbShowCODE.AutoSize = True
    Me.chbShowCODE.Location = New System.Drawing.Point(6, 42)
    Me.chbShowCODE.Name = "chbShowCODE"
    Me.chbShowCODE.Size = New System.Drawing.Size(180, 17)
    Me.chbShowCODE.TabIndex = 17
    Me.chbShowCODE.Text = "Show CODE nodes in the syntax"
    Me.chbShowCODE.UseVisualStyleBackColor = True
    '
    'cmdWorkDir
    '
    Me.cmdWorkDir.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdWorkDir.Location = New System.Drawing.Point(697, 9)
    Me.cmdWorkDir.Name = "cmdWorkDir"
    Me.cmdWorkDir.Size = New System.Drawing.Size(27, 20)
    Me.cmdWorkDir.TabIndex = 2
    Me.cmdWorkDir.Text = "..."
    Me.cmdWorkDir.UseVisualStyleBackColor = True
    '
    'tbWorkDir
    '
    Me.tbWorkDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbWorkDir.Location = New System.Drawing.Point(112, 9)
    Me.tbWorkDir.Name = "tbWorkDir"
    Me.tbWorkDir.Size = New System.Drawing.Size(579, 20)
    Me.tbWorkDir.TabIndex = 1
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(8, 12)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(92, 13)
    Me.Label1.TabIndex = 0
    Me.Label1.Text = "Working Directory"
    '
    'tpCoref
    '
    Me.tpCoref.Controls.Add(Me.cboColor)
    Me.tpCoref.Controls.Add(Me.Label11)
    Me.tpCoref.Controls.Add(Me.tbCorefName)
    Me.tpCoref.Controls.Add(Me.Label5)
    Me.tpCoref.Controls.Add(Me.tbCorefDescr)
    Me.tpCoref.Controls.Add(Me.Label4)
    Me.tpCoref.Controls.Add(Me.dgvCorefType)
    Me.tpCoref.Controls.Add(Me.Label2)
    Me.tpCoref.Location = New System.Drawing.Point(4, 22)
    Me.tpCoref.Name = "tpCoref"
    Me.tpCoref.Size = New System.Drawing.Size(732, 414)
    Me.tpCoref.TabIndex = 1
    Me.tpCoref.Text = "Coreference Types"
    Me.tpCoref.UseVisualStyleBackColor = True
    '
    'cboColor
    '
    Me.cboColor.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cboColor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboColor.FormattingEnabled = True
    Me.cboColor.Location = New System.Drawing.Point(192, 37)
    Me.cboColor.Name = "cboColor"
    Me.cboColor.Size = New System.Drawing.Size(532, 21)
    Me.cboColor.TabIndex = 7
    '
    'Label11
    '
    Me.Label11.AutoSize = True
    Me.Label11.Location = New System.Drawing.Point(160, 40)
    Me.Label11.Name = "Label11"
    Me.Label11.Size = New System.Drawing.Size(34, 13)
    Me.Label11.TabIndex = 6
    Me.Label11.Text = "Color:"
    '
    'tbCorefName
    '
    Me.tbCorefName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbCorefName.Location = New System.Drawing.Point(192, 7)
    Me.tbCorefName.Name = "tbCorefName"
    Me.tbCorefName.Size = New System.Drawing.Size(532, 20)
    Me.tbCorefName.TabIndex = 5
    '
    'Label5
    '
    Me.Label5.AutoSize = True
    Me.Label5.Location = New System.Drawing.Point(157, 10)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(38, 13)
    Me.Label5.TabIndex = 4
    Me.Label5.Text = "Name:"
    '
    'tbCorefDescr
    '
    Me.tbCorefDescr.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbCorefDescr.Location = New System.Drawing.Point(160, 77)
    Me.tbCorefDescr.Name = "tbCorefDescr"
    Me.tbCorefDescr.Size = New System.Drawing.Size(564, 334)
    Me.tbCorefDescr.TabIndex = 3
    Me.tbCorefDescr.Text = ""
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(157, 61)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(63, 13)
    Me.Label4.TabIndex = 2
    Me.Label4.Text = "Description:"
    '
    'dgvCorefType
    '
    Me.dgvCorefType.AllowUserToAddRows = False
    Me.dgvCorefType.AllowUserToDeleteRows = False
    Me.dgvCorefType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.dgvCorefType.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvCorefType.Location = New System.Drawing.Point(6, 26)
    Me.dgvCorefType.Name = "dgvCorefType"
    Me.dgvCorefType.ReadOnly = True
    Me.dgvCorefType.Size = New System.Drawing.Size(148, 385)
    Me.dgvCorefType.TabIndex = 1
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(3, 10)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(96, 13)
    Me.Label2.TabIndex = 0
    Me.Label2.Text = "Coreference Name"
    '
    'tpPhrase
    '
    Me.tpPhrase.Controls.Add(Me.SplitContainer8)
    Me.tpPhrase.Location = New System.Drawing.Point(4, 22)
    Me.tpPhrase.Name = "tpPhrase"
    Me.tpPhrase.Size = New System.Drawing.Size(732, 414)
    Me.tpPhrase.TabIndex = 2
    Me.tpPhrase.Text = "Phrase Types"
    Me.tpPhrase.UseVisualStyleBackColor = True
    '
    'SplitContainer8
    '
    Me.SplitContainer8.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer8.Location = New System.Drawing.Point(0, 0)
    Me.SplitContainer8.Name = "SplitContainer8"
    Me.SplitContainer8.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'SplitContainer8.Panel1
    '
    Me.SplitContainer8.Panel1.Controls.Add(Me.SplitContainer9)
    '
    'SplitContainer8.Panel2
    '
    Me.SplitContainer8.Panel2.Controls.Add(Me.SplitContainer7)
    Me.SplitContainer8.Size = New System.Drawing.Size(732, 414)
    Me.SplitContainer8.SplitterDistance = 172
    Me.SplitContainer8.TabIndex = 17
    '
    'SplitContainer9
    '
    Me.SplitContainer9.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer9.Location = New System.Drawing.Point(0, 0)
    Me.SplitContainer9.Name = "SplitContainer9"
    '
    'SplitContainer9.Panel1
    '
    Me.SplitContainer9.Panel1.Controls.Add(Me.tbPhraseInclude)
    Me.SplitContainer9.Panel1.Controls.Add(Me.Label71)
    '
    'SplitContainer9.Panel2
    '
    Me.SplitContainer9.Panel2.Controls.Add(Me.tbPhraseExclude)
    Me.SplitContainer9.Panel2.Controls.Add(Me.Label72)
    Me.SplitContainer9.Size = New System.Drawing.Size(732, 172)
    Me.SplitContainer9.SplitterDistance = 357
    Me.SplitContainer9.TabIndex = 0
    '
    'tbPhraseInclude
    '
    Me.tbPhraseInclude.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbPhraseInclude.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbPhraseInclude.ForeColor = System.Drawing.Color.Green
    Me.tbPhraseInclude.Location = New System.Drawing.Point(3, 25)
    Me.tbPhraseInclude.Name = "tbPhraseInclude"
    Me.tbPhraseInclude.Size = New System.Drawing.Size(351, 144)
    Me.tbPhraseInclude.TabIndex = 1
    Me.tbPhraseInclude.Text = ""
    '
    'Label71
    '
    Me.Label71.AutoSize = True
    Me.Label71.Location = New System.Drawing.Point(8, 9)
    Me.Label71.Name = "Label71"
    Me.Label71.Size = New System.Drawing.Size(191, 13)
    Me.Label71.TabIndex = 0
    Me.Label71.Text = "Include phrase types for coreferencing:"
    '
    'tbPhraseExclude
    '
    Me.tbPhraseExclude.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbPhraseExclude.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbPhraseExclude.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
    Me.tbPhraseExclude.Location = New System.Drawing.Point(3, 25)
    Me.tbPhraseExclude.Name = "tbPhraseExclude"
    Me.tbPhraseExclude.Size = New System.Drawing.Size(365, 144)
    Me.tbPhraseExclude.TabIndex = 1
    Me.tbPhraseExclude.Text = ""
    '
    'Label72
    '
    Me.Label72.AutoSize = True
    Me.Label72.Location = New System.Drawing.Point(3, 9)
    Me.Label72.Name = "Label72"
    Me.Label72.Size = New System.Drawing.Size(226, 13)
    Me.Label72.TabIndex = 0
    Me.Label72.Text = "These phrase types cannot be an antecedent:"
    '
    'SplitContainer7
    '
    Me.SplitContainer7.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer7.Location = New System.Drawing.Point(0, 0)
    Me.SplitContainer7.Name = "SplitContainer7"
    '
    'SplitContainer7.Panel1
    '
    Me.SplitContainer7.Panel1.Controls.Add(Me.dgvPhraseType)
    Me.SplitContainer7.Panel1.Controls.Add(Me.Label7)
    '
    'SplitContainer7.Panel2
    '
    Me.SplitContainer7.Panel2.Controls.Add(Me.tbPhraseName)
    Me.SplitContainer7.Panel2.Controls.Add(Me.Label70)
    Me.SplitContainer7.Panel2.Controls.Add(Me.tbPhraseNode)
    Me.SplitContainer7.Panel2.Controls.Add(Me.Label10)
    Me.SplitContainer7.Panel2.Controls.Add(Me.Label6)
    Me.SplitContainer7.Panel2.Controls.Add(Me.cboPhraseTarget)
    Me.SplitContainer7.Panel2.Controls.Add(Me.Label8)
    Me.SplitContainer7.Panel2.Controls.Add(Me.cboPhraseType)
    Me.SplitContainer7.Panel2.Controls.Add(Me.tbPhraseChild)
    Me.SplitContainer7.Panel2.Controls.Add(Me.Label9)
    Me.SplitContainer7.Size = New System.Drawing.Size(732, 238)
    Me.SplitContainer7.SplitterDistance = 166
    Me.SplitContainer7.TabIndex = 16
    '
    'dgvPhraseType
    '
    Me.dgvPhraseType.AllowUserToAddRows = False
    Me.dgvPhraseType.AllowUserToDeleteRows = False
    Me.dgvPhraseType.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dgvPhraseType.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
    Me.dgvPhraseType.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvPhraseType.Location = New System.Drawing.Point(3, 18)
    Me.dgvPhraseType.Name = "dgvPhraseType"
    Me.dgvPhraseType.ReadOnly = True
    Me.dgvPhraseType.Size = New System.Drawing.Size(160, 217)
    Me.dgvPhraseType.TabIndex = 7
    '
    'Label7
    '
    Me.Label7.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Label7.AutoSize = True
    Me.Label7.Location = New System.Drawing.Point(3, 4)
    Me.Label7.Name = "Label7"
    Me.Label7.Size = New System.Drawing.Size(49, 13)
    Me.Label7.TabIndex = 6
    Me.Label7.Text = "Category"
    '
    'tbPhraseName
    '
    Me.tbPhraseName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbPhraseName.Location = New System.Drawing.Point(76, 18)
    Me.tbPhraseName.Name = "tbPhraseName"
    Me.tbPhraseName.Size = New System.Drawing.Size(478, 20)
    Me.tbPhraseName.TabIndex = 17
    '
    'Label70
    '
    Me.Label70.AutoSize = True
    Me.Label70.Location = New System.Drawing.Point(11, 21)
    Me.Label70.Name = "Label70"
    Me.Label70.Size = New System.Drawing.Size(38, 13)
    Me.Label70.TabIndex = 16
    Me.Label70.Text = "Name:"
    '
    'tbPhraseNode
    '
    Me.tbPhraseNode.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbPhraseNode.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbPhraseNode.Location = New System.Drawing.Point(76, 59)
    Me.tbPhraseNode.Name = "tbPhraseNode"
    Me.tbPhraseNode.Size = New System.Drawing.Size(121, 21)
    Me.tbPhraseNode.TabIndex = 9
    '
    'Label10
    '
    Me.Label10.AutoSize = True
    Me.Label10.Location = New System.Drawing.Point(11, 139)
    Me.Label10.Name = "Label10"
    Me.Label10.Size = New System.Drawing.Size(38, 13)
    Me.Label10.TabIndex = 15
    Me.Label10.Text = "Target"
    '
    'Label6
    '
    Me.Label6.AutoSize = True
    Me.Label6.Location = New System.Drawing.Point(11, 60)
    Me.Label6.Name = "Label6"
    Me.Label6.Size = New System.Drawing.Size(68, 13)
    Me.Label6.TabIndex = 8
    Me.Label6.Text = "Phrase label:"
    '
    'cboPhraseTarget
    '
    Me.cboPhraseTarget.FormattingEnabled = True
    Me.cboPhraseTarget.Location = New System.Drawing.Point(76, 136)
    Me.cboPhraseTarget.Name = "cboPhraseTarget"
    Me.cboPhraseTarget.Size = New System.Drawing.Size(121, 21)
    Me.cboPhraseTarget.TabIndex = 14
    '
    'Label8
    '
    Me.Label8.AutoSize = True
    Me.Label8.Location = New System.Drawing.Point(11, 86)
    Me.Label8.Name = "Label8"
    Me.Label8.Size = New System.Drawing.Size(58, 13)
    Me.Label8.TabIndex = 10
    Me.Label8.Text = "Child label:"
    '
    'cboPhraseType
    '
    Me.cboPhraseType.FormattingEnabled = True
    Me.cboPhraseType.Location = New System.Drawing.Point(76, 109)
    Me.cboPhraseType.Name = "cboPhraseType"
    Me.cboPhraseType.Size = New System.Drawing.Size(121, 21)
    Me.cboPhraseType.TabIndex = 13
    '
    'tbPhraseChild
    '
    Me.tbPhraseChild.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbPhraseChild.ForeColor = System.Drawing.Color.Green
    Me.tbPhraseChild.Location = New System.Drawing.Point(76, 83)
    Me.tbPhraseChild.Name = "tbPhraseChild"
    Me.tbPhraseChild.Size = New System.Drawing.Size(121, 21)
    Me.tbPhraseChild.TabIndex = 11
    '
    'Label9
    '
    Me.Label9.AutoSize = True
    Me.Label9.Location = New System.Drawing.Point(11, 112)
    Me.Label9.Name = "Label9"
    Me.Label9.Size = New System.Drawing.Size(31, 13)
    Me.Label9.TabIndex = 12
    Me.Label9.Text = "Type"
    '
    'tpPronoun
    '
    Me.tpPronoun.Controls.Add(Me.Label26)
    Me.tpPronoun.Controls.Add(Me.tbProVarMBE)
    Me.tpPronoun.Controls.Add(Me.Label25)
    Me.tpPronoun.Controls.Add(Me.tbProVarEmodE)
    Me.tpPronoun.Controls.Add(Me.Label24)
    Me.tpPronoun.Controls.Add(Me.tbProVarME)
    Me.tpPronoun.Controls.Add(Me.Label23)
    Me.tpPronoun.Controls.Add(Me.tbProNotes)
    Me.tpPronoun.Controls.Add(Me.Label22)
    Me.tpPronoun.Controls.Add(Me.tbProPGN)
    Me.tpPronoun.Controls.Add(Me.Label21)
    Me.tpPronoun.Controls.Add(Me.tbProVarOE)
    Me.tpPronoun.Controls.Add(Me.Label16)
    Me.tpPronoun.Controls.Add(Me.tbProDescr)
    Me.tpPronoun.Controls.Add(Me.Label15)
    Me.tpPronoun.Controls.Add(Me.tbProName)
    Me.tpPronoun.Controls.Add(Me.Label13)
    Me.tpPronoun.Controls.Add(Me.dgvPronoun)
    Me.tpPronoun.Controls.Add(Me.Label12)
    Me.tpPronoun.Location = New System.Drawing.Point(4, 22)
    Me.tpPronoun.Name = "tpPronoun"
    Me.tpPronoun.Size = New System.Drawing.Size(732, 414)
    Me.tpPronoun.TabIndex = 3
    Me.tpPronoun.Text = "Pronouns"
    Me.tpPronoun.UseVisualStyleBackColor = True
    '
    'Label26
    '
    Me.Label26.AutoSize = True
    Me.Label26.Location = New System.Drawing.Point(166, 233)
    Me.Label26.Name = "Label26"
    Me.Label26.Size = New System.Drawing.Size(77, 13)
    Me.Label26.TabIndex = 18
    Me.Label26.Text = "Modern British:"
    '
    'tbProVarMBE
    '
    Me.tbProVarMBE.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbProVarMBE.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbProVarMBE.Location = New System.Drawing.Point(246, 229)
    Me.tbProVarMBE.Name = "tbProVarMBE"
    Me.tbProVarMBE.Size = New System.Drawing.Size(474, 36)
    Me.tbProVarMBE.TabIndex = 17
    Me.tbProVarMBE.Text = ""
    '
    'Label25
    '
    Me.Label25.AutoSize = True
    Me.Label25.Location = New System.Drawing.Point(166, 191)
    Me.Label25.Name = "Label25"
    Me.Label25.Size = New System.Drawing.Size(72, 13)
    Me.Label25.TabIndex = 16
    Me.Label25.Text = "Early Modern:"
    '
    'tbProVarEmodE
    '
    Me.tbProVarEmodE.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbProVarEmodE.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbProVarEmodE.Location = New System.Drawing.Point(246, 187)
    Me.tbProVarEmodE.Name = "tbProVarEmodE"
    Me.tbProVarEmodE.Size = New System.Drawing.Size(474, 36)
    Me.tbProVarEmodE.TabIndex = 15
    Me.tbProVarEmodE.Text = ""
    '
    'Label24
    '
    Me.Label24.AutoSize = True
    Me.Label24.Location = New System.Drawing.Point(166, 149)
    Me.Label24.Name = "Label24"
    Me.Label24.Size = New System.Drawing.Size(78, 13)
    Me.Label24.TabIndex = 14
    Me.Label24.Text = "Middle English:"
    '
    'tbProVarME
    '
    Me.tbProVarME.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbProVarME.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbProVarME.Location = New System.Drawing.Point(246, 145)
    Me.tbProVarME.Name = "tbProVarME"
    Me.tbProVarME.Size = New System.Drawing.Size(474, 36)
    Me.tbProVarME.TabIndex = 13
    Me.tbProVarME.Text = ""
    '
    'Label23
    '
    Me.Label23.AutoSize = True
    Me.Label23.Location = New System.Drawing.Point(166, 107)
    Me.Label23.Name = "Label23"
    Me.Label23.Size = New System.Drawing.Size(63, 13)
    Me.Label23.TabIndex = 12
    Me.Label23.Text = "Old English:"
    '
    'tbProNotes
    '
    Me.tbProNotes.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbProNotes.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbProNotes.Location = New System.Drawing.Point(169, 271)
    Me.tbProNotes.Name = "tbProNotes"
    Me.tbProNotes.Size = New System.Drawing.Size(551, 140)
    Me.tbProNotes.TabIndex = 11
    Me.tbProNotes.Text = ""
    '
    'Label22
    '
    Me.Label22.AutoSize = True
    Me.Label22.Location = New System.Drawing.Point(166, 255)
    Me.Label22.Name = "Label22"
    Me.Label22.Size = New System.Drawing.Size(38, 13)
    Me.Label22.TabIndex = 10
    Me.Label22.Text = "Notes:"
    '
    'tbProPGN
    '
    Me.tbProPGN.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbProPGN.Location = New System.Drawing.Point(297, 60)
    Me.tbProPGN.Name = "tbProPGN"
    Me.tbProPGN.Size = New System.Drawing.Size(423, 20)
    Me.tbProPGN.TabIndex = 9
    '
    'Label21
    '
    Me.Label21.AutoSize = True
    Me.Label21.Location = New System.Drawing.Point(166, 60)
    Me.Label21.Name = "Label21"
    Me.Label21.Size = New System.Drawing.Size(125, 13)
    Me.Label21.TabIndex = 8
    Me.Label21.Text = "Person/Gender/Number:"
    '
    'tbProVarOE
    '
    Me.tbProVarOE.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbProVarOE.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbProVarOE.Location = New System.Drawing.Point(246, 103)
    Me.tbProVarOE.Name = "tbProVarOE"
    Me.tbProVarOE.Size = New System.Drawing.Size(474, 36)
    Me.tbProVarOE.TabIndex = 7
    Me.tbProVarOE.Text = ""
    '
    'Label16
    '
    Me.Label16.AutoSize = True
    Me.Label16.Location = New System.Drawing.Point(166, 84)
    Me.Label16.Name = "Label16"
    Me.Label16.Size = New System.Drawing.Size(264, 13)
    Me.Label16.TabIndex = 6
    Me.Label16.Text = "Variants (divided by the | sign, possibly with wildcard *):"
    '
    'tbProDescr
    '
    Me.tbProDescr.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbProDescr.Location = New System.Drawing.Point(235, 34)
    Me.tbProDescr.Name = "tbProDescr"
    Me.tbProDescr.Size = New System.Drawing.Size(485, 20)
    Me.tbProDescr.TabIndex = 5
    '
    'Label15
    '
    Me.Label15.AutoSize = True
    Me.Label15.Location = New System.Drawing.Point(166, 34)
    Me.Label15.Name = "Label15"
    Me.Label15.Size = New System.Drawing.Size(63, 13)
    Me.Label15.TabIndex = 4
    Me.Label15.Text = "Description:"
    '
    'tbProName
    '
    Me.tbProName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbProName.Location = New System.Drawing.Point(235, 8)
    Me.tbProName.Name = "tbProName"
    Me.tbProName.Size = New System.Drawing.Size(485, 20)
    Me.tbProName.TabIndex = 3
    '
    'Label13
    '
    Me.Label13.AutoSize = True
    Me.Label13.Location = New System.Drawing.Point(166, 8)
    Me.Label13.Name = "Label13"
    Me.Label13.Size = New System.Drawing.Size(38, 13)
    Me.Label13.TabIndex = 2
    Me.Label13.Text = "Name:"
    '
    'dgvPronoun
    '
    Me.dgvPronoun.AllowUserToAddRows = False
    Me.dgvPronoun.AllowUserToDeleteRows = False
    Me.dgvPronoun.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.dgvPronoun.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvPronoun.Location = New System.Drawing.Point(6, 24)
    Me.dgvPronoun.Name = "dgvPronoun"
    Me.dgvPronoun.ReadOnly = True
    Me.dgvPronoun.Size = New System.Drawing.Size(148, 387)
    Me.dgvPronoun.TabIndex = 1
    '
    'Label12
    '
    Me.Label12.AutoSize = True
    Me.Label12.Location = New System.Drawing.Point(3, 8)
    Me.Label12.Name = "Label12"
    Me.Label12.Size = New System.Drawing.Size(72, 13)
    Me.Label12.TabIndex = 0
    Me.Label12.Text = "PronounClass"
    '
    'tpNPfeat
    '
    Me.tpNPfeat.Controls.Add(Me.SplitContainer1)
    Me.tpNPfeat.Location = New System.Drawing.Point(4, 22)
    Me.tpNPfeat.Name = "tpNPfeat"
    Me.tpNPfeat.Size = New System.Drawing.Size(732, 414)
    Me.tpNPfeat.TabIndex = 4
    Me.tpNPfeat.Text = "NP Features"
    Me.tpNPfeat.UseVisualStyleBackColor = True
    '
    'SplitContainer1
    '
    Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
    Me.SplitContainer1.Name = "SplitContainer1"
    '
    'SplitContainer1.Panel1
    '
    Me.SplitContainer1.Panel1.Controls.Add(Me.dgvNPfeat)
    Me.SplitContainer1.Panel1.Controls.Add(Me.Label28)
    '
    'SplitContainer1.Panel2
    '
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbNPfeatVariants)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label31)
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbNPfeatDescr)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label30)
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbNPfeatName)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label29)
    Me.SplitContainer1.Size = New System.Drawing.Size(732, 414)
    Me.SplitContainer1.SplitterDistance = 163
    Me.SplitContainer1.TabIndex = 0
    '
    'dgvNPfeat
    '
    Me.dgvNPfeat.AllowUserToAddRows = False
    Me.dgvNPfeat.AllowUserToDeleteRows = False
    Me.dgvNPfeat.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dgvNPfeat.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvNPfeat.Location = New System.Drawing.Point(6, 24)
    Me.dgvNPfeat.Name = "dgvNPfeat"
    Me.dgvNPfeat.ReadOnly = True
    Me.dgvNPfeat.Size = New System.Drawing.Size(154, 387)
    Me.dgvNPfeat.TabIndex = 1
    '
    'Label28
    '
    Me.Label28.AutoSize = True
    Me.Label28.Location = New System.Drawing.Point(3, 8)
    Me.Label28.Name = "Label28"
    Me.Label28.Size = New System.Drawing.Size(84, 13)
    Me.Label28.TabIndex = 0
    Me.Label28.Text = "NP feature type:"
    '
    'tbNPfeatVariants
    '
    Me.tbNPfeatVariants.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbNPfeatVariants.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbNPfeatVariants.Location = New System.Drawing.Point(6, 77)
    Me.tbNPfeatVariants.Name = "tbNPfeatVariants"
    Me.tbNPfeatVariants.Size = New System.Drawing.Size(551, 334)
    Me.tbNPfeatVariants.TabIndex = 5
    Me.tbNPfeatVariants.Text = ""
    '
    'Label31
    '
    Me.Label31.AutoSize = True
    Me.Label31.Location = New System.Drawing.Point(3, 61)
    Me.Label31.Name = "Label31"
    Me.Label31.Size = New System.Drawing.Size(163, 13)
    Me.Label31.TabIndex = 4
    Me.Label31.Text = "All possible variants, one per line:"
    '
    'tbNPfeatDescr
    '
    Me.tbNPfeatDescr.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbNPfeatDescr.Location = New System.Drawing.Point(72, 31)
    Me.tbNPfeatDescr.Name = "tbNPfeatDescr"
    Me.tbNPfeatDescr.Size = New System.Drawing.Size(485, 20)
    Me.tbNPfeatDescr.TabIndex = 3
    '
    'Label30
    '
    Me.Label30.AutoSize = True
    Me.Label30.Location = New System.Drawing.Point(3, 34)
    Me.Label30.Name = "Label30"
    Me.Label30.Size = New System.Drawing.Size(63, 13)
    Me.Label30.TabIndex = 2
    Me.Label30.Text = "Description:"
    '
    'tbNPfeatName
    '
    Me.tbNPfeatName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbNPfeatName.Location = New System.Drawing.Point(72, 5)
    Me.tbNPfeatName.Name = "tbNPfeatName"
    Me.tbNPfeatName.Size = New System.Drawing.Size(485, 20)
    Me.tbNPfeatName.TabIndex = 1
    '
    'Label29
    '
    Me.Label29.AutoSize = True
    Me.Label29.Location = New System.Drawing.Point(3, 8)
    Me.Label29.Name = "Label29"
    Me.Label29.Size = New System.Drawing.Size(38, 13)
    Me.Label29.TabIndex = 0
    Me.Label29.Text = "Name:"
    '
    'tpCons
    '
    Me.tpCons.Controls.Add(Me.SplitContainer3)
    Me.tpCons.Location = New System.Drawing.Point(4, 22)
    Me.tpCons.Name = "tpCons"
    Me.tpCons.Size = New System.Drawing.Size(732, 414)
    Me.tpCons.TabIndex = 5
    Me.tpCons.Text = "Constraints"
    Me.tpCons.UseVisualStyleBackColor = True
    '
    'SplitContainer3
    '
    Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
    Me.SplitContainer3.Name = "SplitContainer3"
    '
    'SplitContainer3.Panel1
    '
    Me.SplitContainer3.Panel1.Controls.Add(Me.dgvCons)
    Me.SplitContainer3.Panel1.Controls.Add(Me.Label36)
    '
    'SplitContainer3.Panel2
    '
    Me.SplitContainer3.Panel2.Controls.Add(Me.cmdCnsDown)
    Me.SplitContainer3.Panel2.Controls.Add(Me.cmdCnsUp)
    Me.SplitContainer3.Panel2.Controls.Add(Me.tbConsDescr)
    Me.SplitContainer3.Panel2.Controls.Add(Me.Label40)
    Me.SplitContainer3.Panel2.Controls.Add(Me.tbConsMult)
    Me.SplitContainer3.Panel2.Controls.Add(Me.Label39)
    Me.SplitContainer3.Panel2.Controls.Add(Me.tbConsLevel)
    Me.SplitContainer3.Panel2.Controls.Add(Me.Label38)
    Me.SplitContainer3.Panel2.Controls.Add(Me.tbConsName)
    Me.SplitContainer3.Panel2.Controls.Add(Me.Label37)
    Me.SplitContainer3.Size = New System.Drawing.Size(732, 414)
    Me.SplitContainer3.SplitterDistance = 158
    Me.SplitContainer3.TabIndex = 0
    '
    'dgvCons
    '
    Me.dgvCons.AllowUserToAddRows = False
    Me.dgvCons.AllowUserToDeleteRows = False
    Me.dgvCons.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dgvCons.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvCons.Location = New System.Drawing.Point(6, 24)
    Me.dgvCons.Name = "dgvCons"
    Me.dgvCons.ReadOnly = True
    Me.dgvCons.Size = New System.Drawing.Size(149, 387)
    Me.dgvCons.TabIndex = 2
    '
    'Label36
    '
    Me.Label36.AutoSize = True
    Me.Label36.Location = New System.Drawing.Point(3, 8)
    Me.Label36.Name = "Label36"
    Me.Label36.Size = New System.Drawing.Size(57, 13)
    Me.Label36.TabIndex = 1
    Me.Label36.Text = "Constraint:"
    '
    'cmdCnsDown
    '
    Me.cmdCnsDown.Image = CType(resources.GetObject("cmdCnsDown.Image"), System.Drawing.Image)
    Me.cmdCnsDown.Location = New System.Drawing.Point(196, 34)
    Me.cmdCnsDown.Name = "cmdCnsDown"
    Me.cmdCnsDown.Size = New System.Drawing.Size(32, 46)
    Me.cmdCnsDown.TabIndex = 11
    Me.cmdCnsDown.UseVisualStyleBackColor = True
    '
    'cmdCnsUp
    '
    Me.cmdCnsUp.Image = CType(resources.GetObject("cmdCnsUp.Image"), System.Drawing.Image)
    Me.cmdCnsUp.Location = New System.Drawing.Point(157, 34)
    Me.cmdCnsUp.Name = "cmdCnsUp"
    Me.cmdCnsUp.Size = New System.Drawing.Size(33, 46)
    Me.cmdCnsUp.TabIndex = 10
    Me.cmdCnsUp.UseVisualStyleBackColor = True
    '
    'tbConsDescr
    '
    Me.tbConsDescr.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbConsDescr.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbConsDescr.Location = New System.Drawing.Point(11, 108)
    Me.tbConsDescr.Name = "tbConsDescr"
    Me.tbConsDescr.Size = New System.Drawing.Size(551, 303)
    Me.tbConsDescr.TabIndex = 9
    Me.tbConsDescr.Text = ""
    '
    'Label40
    '
    Me.Label40.AutoSize = True
    Me.Label40.Location = New System.Drawing.Point(8, 92)
    Me.Label40.Name = "Label40"
    Me.Label40.Size = New System.Drawing.Size(122, 13)
    Me.Label40.TabIndex = 8
    Me.Label40.Text = "Description/explanation:"
    '
    'tbConsMult
    '
    Me.tbConsMult.Location = New System.Drawing.Point(77, 60)
    Me.tbConsMult.Name = "tbConsMult"
    Me.tbConsMult.Size = New System.Drawing.Size(61, 20)
    Me.tbConsMult.TabIndex = 7
    '
    'Label39
    '
    Me.Label39.AutoSize = True
    Me.Label39.Location = New System.Drawing.Point(8, 63)
    Me.Label39.Name = "Label39"
    Me.Label39.Size = New System.Drawing.Size(71, 13)
    Me.Label39.TabIndex = 6
    Me.Label39.Text = "Multiplication:"
    '
    'tbConsLevel
    '
    Me.tbConsLevel.Location = New System.Drawing.Point(77, 34)
    Me.tbConsLevel.Name = "tbConsLevel"
    Me.tbConsLevel.Size = New System.Drawing.Size(61, 20)
    Me.tbConsLevel.TabIndex = 5
    '
    'Label38
    '
    Me.Label38.AutoSize = True
    Me.Label38.Location = New System.Drawing.Point(8, 37)
    Me.Label38.Name = "Label38"
    Me.Label38.Size = New System.Drawing.Size(36, 13)
    Me.Label38.TabIndex = 4
    Me.Label38.Text = "Level:"
    '
    'tbConsName
    '
    Me.tbConsName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbConsName.Location = New System.Drawing.Point(77, 8)
    Me.tbConsName.Name = "tbConsName"
    Me.tbConsName.Size = New System.Drawing.Size(485, 20)
    Me.tbConsName.TabIndex = 3
    '
    'Label37
    '
    Me.Label37.AutoSize = True
    Me.Label37.Location = New System.Drawing.Point(8, 11)
    Me.Label37.Name = "Label37"
    Me.Label37.Size = New System.Drawing.Size(38, 13)
    Me.Label37.TabIndex = 2
    Me.Label37.Text = "Name:"
    '
    'tpCat
    '
    Me.tpCat.Controls.Add(Me.Label56)
    Me.tpCat.Controls.Add(Me.tbCatType)
    Me.tpCat.Controls.Add(Me.Label55)
    Me.tpCat.Controls.Add(Me.Label47)
    Me.tpCat.Controls.Add(Me.tbCatMBE)
    Me.tpCat.Controls.Add(Me.Label48)
    Me.tpCat.Controls.Add(Me.tbCatEmodE)
    Me.tpCat.Controls.Add(Me.Label49)
    Me.tpCat.Controls.Add(Me.tbCatME)
    Me.tpCat.Controls.Add(Me.Label50)
    Me.tpCat.Controls.Add(Me.tbCatNotes)
    Me.tpCat.Controls.Add(Me.Label51)
    Me.tpCat.Controls.Add(Me.tbCatOE)
    Me.tpCat.Controls.Add(Me.Label52)
    Me.tpCat.Controls.Add(Me.tbCatDescr)
    Me.tpCat.Controls.Add(Me.Label53)
    Me.tpCat.Controls.Add(Me.tbCatName)
    Me.tpCat.Controls.Add(Me.Label54)
    Me.tpCat.Controls.Add(Me.dgvCategory)
    Me.tpCat.Controls.Add(Me.Label46)
    Me.tpCat.Location = New System.Drawing.Point(4, 22)
    Me.tpCat.Name = "tpCat"
    Me.tpCat.Size = New System.Drawing.Size(732, 414)
    Me.tpCat.TabIndex = 6
    Me.tpCat.Text = "Categories"
    Me.tpCat.UseVisualStyleBackColor = True
    '
    'Label56
    '
    Me.Label56.AutoSize = True
    Me.Label56.ForeColor = System.Drawing.Color.Maroon
    Me.Label56.Location = New System.Drawing.Point(400, 66)
    Me.Label56.Name = "Label56"
    Me.Label56.Size = New System.Drawing.Size(318, 13)
    Me.Label56.TabIndex = 36
    Me.Label56.Text = "(The category type should be the category's part after the hyphen)"
    '
    'tbCatType
    '
    Me.tbCatType.Location = New System.Drawing.Point(223, 63)
    Me.tbCatType.Name = "tbCatType"
    Me.tbCatType.Size = New System.Drawing.Size(171, 20)
    Me.tbCatType.TabIndex = 35
    '
    'Label55
    '
    Me.Label55.AutoSize = True
    Me.Label55.Location = New System.Drawing.Point(154, 66)
    Me.Label55.Name = "Label55"
    Me.Label55.Size = New System.Drawing.Size(31, 13)
    Me.Label55.TabIndex = 34
    Me.Label55.Text = "Type"
    '
    'Label47
    '
    Me.Label47.AutoSize = True
    Me.Label47.Location = New System.Drawing.Point(154, 239)
    Me.Label47.Name = "Label47"
    Me.Label47.Size = New System.Drawing.Size(77, 13)
    Me.Label47.TabIndex = 33
    Me.Label47.Text = "Modern British:"
    '
    'tbCatMBE
    '
    Me.tbCatMBE.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbCatMBE.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbCatMBE.Location = New System.Drawing.Point(234, 235)
    Me.tbCatMBE.Name = "tbCatMBE"
    Me.tbCatMBE.Size = New System.Drawing.Size(486, 36)
    Me.tbCatMBE.TabIndex = 32
    Me.tbCatMBE.Text = ""
    '
    'Label48
    '
    Me.Label48.AutoSize = True
    Me.Label48.Location = New System.Drawing.Point(154, 197)
    Me.Label48.Name = "Label48"
    Me.Label48.Size = New System.Drawing.Size(72, 13)
    Me.Label48.TabIndex = 31
    Me.Label48.Text = "Early Modern:"
    '
    'tbCatEmodE
    '
    Me.tbCatEmodE.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbCatEmodE.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbCatEmodE.Location = New System.Drawing.Point(234, 193)
    Me.tbCatEmodE.Name = "tbCatEmodE"
    Me.tbCatEmodE.Size = New System.Drawing.Size(486, 36)
    Me.tbCatEmodE.TabIndex = 30
    Me.tbCatEmodE.Text = ""
    '
    'Label49
    '
    Me.Label49.AutoSize = True
    Me.Label49.Location = New System.Drawing.Point(154, 155)
    Me.Label49.Name = "Label49"
    Me.Label49.Size = New System.Drawing.Size(78, 13)
    Me.Label49.TabIndex = 29
    Me.Label49.Text = "Middle English:"
    '
    'tbCatME
    '
    Me.tbCatME.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbCatME.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbCatME.Location = New System.Drawing.Point(234, 151)
    Me.tbCatME.Name = "tbCatME"
    Me.tbCatME.Size = New System.Drawing.Size(486, 36)
    Me.tbCatME.TabIndex = 28
    Me.tbCatME.Text = ""
    '
    'Label50
    '
    Me.Label50.AutoSize = True
    Me.Label50.Location = New System.Drawing.Point(154, 113)
    Me.Label50.Name = "Label50"
    Me.Label50.Size = New System.Drawing.Size(63, 13)
    Me.Label50.TabIndex = 27
    Me.Label50.Text = "Old English:"
    '
    'tbCatNotes
    '
    Me.tbCatNotes.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbCatNotes.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbCatNotes.Location = New System.Drawing.Point(157, 274)
    Me.tbCatNotes.Name = "tbCatNotes"
    Me.tbCatNotes.Size = New System.Drawing.Size(563, 86)
    Me.tbCatNotes.TabIndex = 26
    Me.tbCatNotes.Text = ""
    '
    'Label51
    '
    Me.Label51.AutoSize = True
    Me.Label51.Location = New System.Drawing.Point(154, 258)
    Me.Label51.Name = "Label51"
    Me.Label51.Size = New System.Drawing.Size(38, 13)
    Me.Label51.TabIndex = 25
    Me.Label51.Text = "Notes:"
    '
    'tbCatOE
    '
    Me.tbCatOE.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbCatOE.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbCatOE.Location = New System.Drawing.Point(234, 109)
    Me.tbCatOE.Name = "tbCatOE"
    Me.tbCatOE.Size = New System.Drawing.Size(486, 36)
    Me.tbCatOE.TabIndex = 24
    Me.tbCatOE.Text = ""
    '
    'Label52
    '
    Me.Label52.AutoSize = True
    Me.Label52.Location = New System.Drawing.Point(154, 90)
    Me.Label52.Name = "Label52"
    Me.Label52.Size = New System.Drawing.Size(264, 13)
    Me.Label52.TabIndex = 23
    Me.Label52.Text = "Variants (divided by the | sign, possibly with wildcard *):"
    '
    'tbCatDescr
    '
    Me.tbCatDescr.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbCatDescr.Location = New System.Drawing.Point(223, 37)
    Me.tbCatDescr.Name = "tbCatDescr"
    Me.tbCatDescr.Size = New System.Drawing.Size(497, 20)
    Me.tbCatDescr.TabIndex = 22
    '
    'Label53
    '
    Me.Label53.AutoSize = True
    Me.Label53.Location = New System.Drawing.Point(154, 37)
    Me.Label53.Name = "Label53"
    Me.Label53.Size = New System.Drawing.Size(63, 13)
    Me.Label53.TabIndex = 21
    Me.Label53.Text = "Description:"
    '
    'tbCatName
    '
    Me.tbCatName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbCatName.Location = New System.Drawing.Point(223, 7)
    Me.tbCatName.Name = "tbCatName"
    Me.tbCatName.Size = New System.Drawing.Size(497, 20)
    Me.tbCatName.TabIndex = 20
    '
    'Label54
    '
    Me.Label54.AutoSize = True
    Me.Label54.Location = New System.Drawing.Point(154, 10)
    Me.Label54.Name = "Label54"
    Me.Label54.Size = New System.Drawing.Size(38, 13)
    Me.Label54.TabIndex = 19
    Me.Label54.Text = "Name:"
    '
    'dgvCategory
    '
    Me.dgvCategory.AllowUserToAddRows = False
    Me.dgvCategory.AllowUserToDeleteRows = False
    Me.dgvCategory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.dgvCategory.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvCategory.Location = New System.Drawing.Point(3, 26)
    Me.dgvCategory.Name = "dgvCategory"
    Me.dgvCategory.ReadOnly = True
    Me.dgvCategory.Size = New System.Drawing.Size(148, 333)
    Me.dgvCategory.TabIndex = 3
    '
    'Label46
    '
    Me.Label46.AutoSize = True
    Me.Label46.Location = New System.Drawing.Point(0, 10)
    Me.Label46.Name = "Label46"
    Me.Label46.Size = New System.Drawing.Size(49, 13)
    Me.Label46.TabIndex = 2
    Me.Label46.Text = "Category"
    '
    'tpStatus
    '
    Me.tpStatus.Controls.Add(Me.SplitContainer4)
    Me.tpStatus.Location = New System.Drawing.Point(4, 22)
    Me.tpStatus.Name = "tpStatus"
    Me.tpStatus.Size = New System.Drawing.Size(732, 414)
    Me.tpStatus.TabIndex = 7
    Me.tpStatus.Text = "Status"
    Me.tpStatus.UseVisualStyleBackColor = True
    '
    'SplitContainer4
    '
    Me.SplitContainer4.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer4.Location = New System.Drawing.Point(0, 0)
    Me.SplitContainer4.Name = "SplitContainer4"
    '
    'SplitContainer4.Panel1
    '
    Me.SplitContainer4.Panel1.Controls.Add(Me.dgvStatus)
    Me.SplitContainer4.Panel1.Controls.Add(Me.Label58)
    '
    'SplitContainer4.Panel2
    '
    Me.SplitContainer4.Panel2.Controls.Add(Me.tbStatusDesc)
    Me.SplitContainer4.Panel2.Controls.Add(Me.Label59)
    Me.SplitContainer4.Panel2.Controls.Add(Me.tbStatusName)
    Me.SplitContainer4.Panel2.Controls.Add(Me.Label60)
    Me.SplitContainer4.Size = New System.Drawing.Size(732, 414)
    Me.SplitContainer4.SplitterDistance = 173
    Me.SplitContainer4.TabIndex = 0
    '
    'dgvStatus
    '
    Me.dgvStatus.AllowUserToAddRows = False
    Me.dgvStatus.AllowUserToDeleteRows = False
    Me.dgvStatus.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dgvStatus.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
    Me.dgvStatus.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvStatus.Location = New System.Drawing.Point(8, 26)
    Me.dgvStatus.Name = "dgvStatus"
    Me.dgvStatus.ReadOnly = True
    Me.dgvStatus.Size = New System.Drawing.Size(162, 385)
    Me.dgvStatus.TabIndex = 1
    '
    'Label58
    '
    Me.Label58.AutoSize = True
    Me.Label58.Location = New System.Drawing.Point(8, 10)
    Me.Label58.Name = "Label58"
    Me.Label58.Size = New System.Drawing.Size(37, 13)
    Me.Label58.TabIndex = 0
    Me.Label58.Text = "Status"
    '
    'tbStatusDesc
    '
    Me.tbStatusDesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbStatusDesc.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbStatusDesc.Location = New System.Drawing.Point(3, 76)
    Me.tbStatusDesc.Name = "tbStatusDesc"
    Me.tbStatusDesc.Size = New System.Drawing.Size(544, 334)
    Me.tbStatusDesc.TabIndex = 9
    Me.tbStatusDesc.Text = ""
    '
    'Label59
    '
    Me.Label59.AutoSize = True
    Me.Label59.Location = New System.Drawing.Point(0, 60)
    Me.Label59.Name = "Label59"
    Me.Label59.Size = New System.Drawing.Size(125, 13)
    Me.Label59.TabIndex = 8
    Me.Label59.Text = "Description of this status:"
    '
    'tbStatusName
    '
    Me.tbStatusName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbStatusName.Location = New System.Drawing.Point(69, 10)
    Me.tbStatusName.Name = "tbStatusName"
    Me.tbStatusName.Size = New System.Drawing.Size(478, 20)
    Me.tbStatusName.TabIndex = 7
    '
    'Label60
    '
    Me.Label60.AutoSize = True
    Me.Label60.Location = New System.Drawing.Point(3, 11)
    Me.Label60.Name = "Label60"
    Me.Label60.Size = New System.Drawing.Size(38, 13)
    Me.Label60.TabIndex = 6
    Me.Label60.Text = "Name:"
    '
    'tpCharting
    '
    Me.tpCharting.Controls.Add(Me.SplitContainer5)
    Me.tpCharting.Location = New System.Drawing.Point(4, 22)
    Me.tpCharting.Name = "tpCharting"
    Me.tpCharting.Size = New System.Drawing.Size(732, 414)
    Me.tpCharting.TabIndex = 8
    Me.tpCharting.Text = "Charting"
    Me.tpCharting.UseVisualStyleBackColor = True
    '
    'SplitContainer5
    '
    Me.SplitContainer5.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer5.Location = New System.Drawing.Point(0, 0)
    Me.SplitContainer5.Name = "SplitContainer5"
    '
    'SplitContainer5.Panel1
    '
    Me.SplitContainer5.Panel1.Controls.Add(Me.dgvTemplate)
    Me.SplitContainer5.Panel1.Controls.Add(Me.Label62)
    '
    'SplitContainer5.Panel2
    '
    Me.SplitContainer5.Panel2.Controls.Add(Me.tbTemplateDescr)
    Me.SplitContainer5.Panel2.Controls.Add(Me.Label69)
    Me.SplitContainer5.Panel2.Controls.Add(Me.SplitContainer6)
    Me.SplitContainer5.Panel2.Controls.Add(Me.tbTemplateName)
    Me.SplitContainer5.Panel2.Controls.Add(Me.Label63)
    Me.SplitContainer5.Size = New System.Drawing.Size(732, 414)
    Me.SplitContainer5.SplitterDistance = 171
    Me.SplitContainer5.TabIndex = 0
    '
    'dgvTemplate
    '
    Me.dgvTemplate.AllowUserToAddRows = False
    Me.dgvTemplate.AllowUserToDeleteRows = False
    Me.dgvTemplate.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dgvTemplate.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
    Me.dgvTemplate.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvTemplate.Location = New System.Drawing.Point(4, 23)
    Me.dgvTemplate.Name = "dgvTemplate"
    Me.dgvTemplate.ReadOnly = True
    Me.dgvTemplate.Size = New System.Drawing.Size(162, 388)
    Me.dgvTemplate.TabIndex = 3
    '
    'Label62
    '
    Me.Label62.AutoSize = True
    Me.Label62.Location = New System.Drawing.Point(4, 7)
    Me.Label62.Name = "Label62"
    Me.Label62.Size = New System.Drawing.Size(83, 13)
    Me.Label62.TabIndex = 2
    Me.Label62.Text = "Template name:"
    '
    'tbTemplateDescr
    '
    Me.tbTemplateDescr.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbTemplateDescr.Location = New System.Drawing.Point(3, 49)
    Me.tbTemplateDescr.Name = "tbTemplateDescr"
    Me.tbTemplateDescr.Size = New System.Drawing.Size(546, 55)
    Me.tbTemplateDescr.TabIndex = 12
    Me.tbTemplateDescr.Text = ""
    '
    'Label69
    '
    Me.Label69.AutoSize = True
    Me.Label69.Location = New System.Drawing.Point(5, 33)
    Me.Label69.Name = "Label69"
    Me.Label69.Size = New System.Drawing.Size(63, 13)
    Me.Label69.TabIndex = 11
    Me.Label69.Text = "Description:"
    '
    'SplitContainer6
    '
    Me.SplitContainer6.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.SplitContainer6.Location = New System.Drawing.Point(3, 110)
    Me.SplitContainer6.Name = "SplitContainer6"
    '
    'SplitContainer6.Panel1
    '
    Me.SplitContainer6.Panel1.Controls.Add(Me.dgvTcell)
    '
    'SplitContainer6.Panel2
    '
    Me.SplitContainer6.Panel2.Controls.Add(Me.cmdCellDel)
    Me.SplitContainer6.Panel2.Controls.Add(Me.cmdCellNew)
    Me.SplitContainer6.Panel2.Controls.Add(Me.cmdCellDown)
    Me.SplitContainer6.Panel2.Controls.Add(Me.cmdCellUp)
    Me.SplitContainer6.Panel2.Controls.Add(Me.Label68)
    Me.SplitContainer6.Panel2.Controls.Add(Me.tbTcellEnv)
    Me.SplitContainer6.Panel2.Controls.Add(Me.tbTcellDescr)
    Me.SplitContainer6.Panel2.Controls.Add(Me.Label67)
    Me.SplitContainer6.Panel2.Controls.Add(Me.tbTcellContent)
    Me.SplitContainer6.Panel2.Controls.Add(Me.Label66)
    Me.SplitContainer6.Panel2.Controls.Add(Me.tbTcellId)
    Me.SplitContainer6.Panel2.Controls.Add(Me.Label65)
    Me.SplitContainer6.Panel2.Controls.Add(Me.tbTcellName)
    Me.SplitContainer6.Panel2.Controls.Add(Me.Label64)
    Me.SplitContainer6.Size = New System.Drawing.Size(551, 301)
    Me.SplitContainer6.SplitterDistance = 76
    Me.SplitContainer6.TabIndex = 10
    '
    'dgvTcell
    '
    Me.dgvTcell.AllowUserToAddRows = False
    Me.dgvTcell.AllowUserToDeleteRows = False
    Me.dgvTcell.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
    Me.dgvTcell.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvTcell.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgvTcell.Location = New System.Drawing.Point(0, 0)
    Me.dgvTcell.Name = "dgvTcell"
    Me.dgvTcell.ReadOnly = True
    Me.dgvTcell.Size = New System.Drawing.Size(76, 301)
    Me.dgvTcell.TabIndex = 4
    '
    'cmdCellDel
    '
    Me.cmdCellDel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCellDel.Location = New System.Drawing.Point(374, 29)
    Me.cmdCellDel.Name = "cmdCellDel"
    Me.cmdCellDel.Size = New System.Drawing.Size(92, 25)
    Me.cmdCellDel.TabIndex = 23
    Me.cmdCellDel.Text = "&Delete"
    Me.cmdCellDel.UseVisualStyleBackColor = True
    '
    'cmdCellNew
    '
    Me.cmdCellNew.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCellNew.Location = New System.Drawing.Point(276, 29)
    Me.cmdCellNew.Name = "cmdCellNew"
    Me.cmdCellNew.Size = New System.Drawing.Size(92, 25)
    Me.cmdCellNew.TabIndex = 22
    Me.cmdCellNew.Text = "&New"
    Me.cmdCellNew.UseVisualStyleBackColor = True
    '
    'cmdCellDown
    '
    Me.cmdCellDown.Image = CType(resources.GetObject("cmdCellDown.Image"), System.Drawing.Image)
    Me.cmdCellDown.Location = New System.Drawing.Point(172, 29)
    Me.cmdCellDown.Name = "cmdCellDown"
    Me.cmdCellDown.Size = New System.Drawing.Size(32, 37)
    Me.cmdCellDown.TabIndex = 21
    Me.cmdCellDown.UseVisualStyleBackColor = True
    '
    'cmdCellUp
    '
    Me.cmdCellUp.Image = CType(resources.GetObject("cmdCellUp.Image"), System.Drawing.Image)
    Me.cmdCellUp.Location = New System.Drawing.Point(133, 29)
    Me.cmdCellUp.Name = "cmdCellUp"
    Me.cmdCellUp.Size = New System.Drawing.Size(33, 37)
    Me.cmdCellUp.TabIndex = 20
    Me.cmdCellUp.UseVisualStyleBackColor = True
    '
    'Label68
    '
    Me.Label68.AutoSize = True
    Me.Label68.Location = New System.Drawing.Point(9, 204)
    Me.Label68.Name = "Label68"
    Me.Label68.Size = New System.Drawing.Size(52, 13)
    Me.Label68.TabIndex = 19
    Me.Label68.Text = "Remarks:"
    '
    'tbTcellEnv
    '
    Me.tbTcellEnv.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbTcellEnv.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbTcellEnv.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbTcellEnv.Location = New System.Drawing.Point(3, 181)
    Me.tbTcellEnv.Name = "tbTcellEnv"
    Me.tbTcellEnv.Size = New System.Drawing.Size(463, 21)
    Me.tbTcellEnv.TabIndex = 18
    '
    'tbTcellDescr
    '
    Me.tbTcellDescr.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbTcellDescr.Location = New System.Drawing.Point(3, 220)
    Me.tbTcellDescr.Name = "tbTcellDescr"
    Me.tbTcellDescr.Size = New System.Drawing.Size(465, 78)
    Me.tbTcellDescr.TabIndex = 17
    Me.tbTcellDescr.Text = ""
    '
    'Label67
    '
    Me.Label67.AutoSize = True
    Me.Label67.Location = New System.Drawing.Point(9, 165)
    Me.Label67.Name = "Label67"
    Me.Label67.Size = New System.Drawing.Size(69, 13)
    Me.Label67.TabIndex = 16
    Me.Label67.Text = "Environment:"
    '
    'tbTcellContent
    '
    Me.tbTcellContent.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbTcellContent.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbTcellContent.ForeColor = System.Drawing.Color.Maroon
    Me.tbTcellContent.Location = New System.Drawing.Point(71, 72)
    Me.tbTcellContent.Name = "tbTcellContent"
    Me.tbTcellContent.Size = New System.Drawing.Size(395, 79)
    Me.tbTcellContent.TabIndex = 15
    Me.tbTcellContent.Text = ""
    '
    'Label66
    '
    Me.Label66.AutoSize = True
    Me.Label66.Location = New System.Drawing.Point(9, 75)
    Me.Label66.Name = "Label66"
    Me.Label66.Size = New System.Drawing.Size(52, 13)
    Me.Label66.TabIndex = 14
    Me.Label66.Text = "Contents:"
    '
    'tbTcellId
    '
    Me.tbTcellId.Location = New System.Drawing.Point(71, 29)
    Me.tbTcellId.Name = "tbTcellId"
    Me.tbTcellId.ReadOnly = True
    Me.tbTcellId.Size = New System.Drawing.Size(56, 20)
    Me.tbTcellId.TabIndex = 13
    '
    'Label65
    '
    Me.Label65.AutoSize = True
    Me.Label65.Location = New System.Drawing.Point(9, 32)
    Me.Label65.Name = "Label65"
    Me.Label65.Size = New System.Drawing.Size(47, 13)
    Me.Label65.TabIndex = 12
    Me.Label65.Text = "Position:"
    '
    'tbTcellName
    '
    Me.tbTcellName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbTcellName.Location = New System.Drawing.Point(71, 3)
    Me.tbTcellName.Name = "tbTcellName"
    Me.tbTcellName.Size = New System.Drawing.Size(395, 20)
    Me.tbTcellName.TabIndex = 11
    '
    'Label64
    '
    Me.Label64.AutoSize = True
    Me.Label64.Location = New System.Drawing.Point(9, 6)
    Me.Label64.Name = "Label64"
    Me.Label64.Size = New System.Drawing.Size(56, 13)
    Me.Label64.TabIndex = 10
    Me.Label64.Text = "Cell name:"
    '
    'tbTemplateName
    '
    Me.tbTemplateName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbTemplateName.Location = New System.Drawing.Point(95, 7)
    Me.tbTemplateName.Name = "tbTemplateName"
    Me.tbTemplateName.Size = New System.Drawing.Size(454, 20)
    Me.tbTemplateName.TabIndex = 9
    '
    'Label63
    '
    Me.Label63.AutoSize = True
    Me.Label63.Location = New System.Drawing.Point(5, 8)
    Me.Label63.Name = "Label63"
    Me.Label63.Size = New System.Drawing.Size(83, 13)
    Me.Label63.TabIndex = 8
    Me.Label63.Text = "Template name:"
    '
    'tpXrel
    '
    Me.tpXrel.Controls.Add(Me.SplitContainer10)
    Me.tpXrel.Location = New System.Drawing.Point(4, 22)
    Me.tpXrel.Name = "tpXrel"
    Me.tpXrel.Size = New System.Drawing.Size(732, 414)
    Me.tpXrel.TabIndex = 9
    Me.tpXrel.Text = "XRelations"
    Me.tpXrel.UseVisualStyleBackColor = True
    '
    'SplitContainer10
    '
    Me.SplitContainer10.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer10.Location = New System.Drawing.Point(0, 0)
    Me.SplitContainer10.Name = "SplitContainer10"
    '
    'SplitContainer10.Panel1
    '
    Me.SplitContainer10.Panel1.Controls.Add(Me.Label77)
    Me.SplitContainer10.Panel1.Controls.Add(Me.dgvXrel)
    '
    'SplitContainer10.Panel2
    '
    Me.SplitContainer10.Panel2.Controls.Add(Me.Label81)
    Me.SplitContainer10.Panel2.Controls.Add(Me.tbXrelXname)
    Me.SplitContainer10.Panel2.Controls.Add(Me.cboXrelType)
    Me.SplitContainer10.Panel2.Controls.Add(Me.tbXrelDescr)
    Me.SplitContainer10.Panel2.Controls.Add(Me.Label78)
    Me.SplitContainer10.Panel2.Controls.Add(Me.Label79)
    Me.SplitContainer10.Panel2.Controls.Add(Me.tbXrelName)
    Me.SplitContainer10.Panel2.Controls.Add(Me.Label80)
    Me.SplitContainer10.Size = New System.Drawing.Size(732, 414)
    Me.SplitContainer10.SplitterDistance = 138
    Me.SplitContainer10.TabIndex = 0
    '
    'Label77
    '
    Me.Label77.AutoSize = True
    Me.Label77.Location = New System.Drawing.Point(3, 9)
    Me.Label77.Name = "Label77"
    Me.Label77.Size = New System.Drawing.Size(102, 13)
    Me.Label77.TabIndex = 1
    Me.Label77.Text = "Relation or Function"
    '
    'dgvXrel
    '
    Me.dgvXrel.AllowUserToAddRows = False
    Me.dgvXrel.AllowUserToDeleteRows = False
    Me.dgvXrel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dgvXrel.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
    Me.dgvXrel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvXrel.Location = New System.Drawing.Point(0, 25)
    Me.dgvXrel.Name = "dgvXrel"
    Me.dgvXrel.ReadOnly = True
    Me.dgvXrel.Size = New System.Drawing.Size(137, 389)
    Me.dgvXrel.TabIndex = 0
    '
    'Label81
    '
    Me.Label81.AutoSize = True
    Me.Label81.Location = New System.Drawing.Point(5, 131)
    Me.Label81.Name = "Label81"
    Me.Label81.Size = New System.Drawing.Size(63, 13)
    Me.Label81.TabIndex = 19
    Me.Label81.Text = "Description:"
    '
    'tbXrelXname
    '
    Me.tbXrelXname.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbXrelXname.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbXrelXname.ForeColor = System.Drawing.Color.Green
    Me.tbXrelXname.Location = New System.Drawing.Point(123, 59)
    Me.tbXrelXname.Name = "tbXrelXname"
    Me.tbXrelXname.Size = New System.Drawing.Size(467, 85)
    Me.tbXrelXname.TabIndex = 18
    Me.tbXrelXname.Text = ""
    '
    'cboXrelType
    '
    Me.cboXrelType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboXrelType.FormattingEnabled = True
    Me.cboXrelType.Location = New System.Drawing.Point(123, 32)
    Me.cboXrelType.Name = "cboXrelType"
    Me.cboXrelType.Size = New System.Drawing.Size(217, 21)
    Me.cboXrelType.TabIndex = 17
    '
    'tbXrelDescr
    '
    Me.tbXrelDescr.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbXrelDescr.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbXrelDescr.Location = New System.Drawing.Point(3, 150)
    Me.tbXrelDescr.Name = "tbXrelDescr"
    Me.tbXrelDescr.Size = New System.Drawing.Size(587, 264)
    Me.tbXrelDescr.TabIndex = 16
    Me.tbXrelDescr.Text = ""
    '
    'Label78
    '
    Me.Label78.AutoSize = True
    Me.Label78.Location = New System.Drawing.Point(5, 35)
    Me.Label78.Name = "Label78"
    Me.Label78.Size = New System.Drawing.Size(34, 13)
    Me.Label78.TabIndex = 14
    Me.Label78.Text = "Type:"
    '
    'Label79
    '
    Me.Label79.AutoSize = True
    Me.Label79.Location = New System.Drawing.Point(5, 62)
    Me.Label79.Name = "Label79"
    Me.Label79.Size = New System.Drawing.Size(112, 13)
    Me.Label79.TabIndex = 12
    Me.Label79.Text = "Xpath axis or function:"
    '
    'tbXrelName
    '
    Me.tbXrelName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbXrelName.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbXrelName.Location = New System.Drawing.Point(123, 6)
    Me.tbXrelName.Name = "tbXrelName"
    Me.tbXrelName.Size = New System.Drawing.Size(467, 20)
    Me.tbXrelName.TabIndex = 11
    '
    'Label80
    '
    Me.Label80.AutoSize = True
    Me.Label80.Location = New System.Drawing.Point(5, 9)
    Me.Label80.Name = "Label80"
    Me.Label80.Size = New System.Drawing.Size(38, 13)
    Me.Label80.TabIndex = 10
    Me.Label80.Text = "Name:"
    '
    'tpMos
    '
    Me.tpMos.Controls.Add(Me.SplitContainer11)
    Me.tpMos.Location = New System.Drawing.Point(4, 22)
    Me.tpMos.Name = "tpMos"
    Me.tpMos.Size = New System.Drawing.Size(732, 414)
    Me.tpMos.TabIndex = 10
    Me.tpMos.Text = "Mos"
    Me.tpMos.UseVisualStyleBackColor = True
    '
    'SplitContainer11
    '
    Me.SplitContainer11.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer11.Location = New System.Drawing.Point(0, 0)
    Me.SplitContainer11.Name = "SplitContainer11"
    '
    'SplitContainer11.Panel1
    '
    Me.SplitContainer11.Panel1.Controls.Add(Me.cmdMosDown)
    Me.SplitContainer11.Panel1.Controls.Add(Me.cmdMosUp)
    Me.SplitContainer11.Panel1.Controls.Add(Me.dgvMos)
    Me.SplitContainer11.Panel1.Controls.Add(Me.Label82)
    '
    'SplitContainer11.Panel2
    '
    Me.SplitContainer11.Panel2.Controls.Add(Me.SplitContainer12)
    Me.SplitContainer11.Size = New System.Drawing.Size(732, 414)
    Me.SplitContainer11.SplitterDistance = 173
    Me.SplitContainer11.TabIndex = 0
    '
    'cmdMosDown
    '
    Me.cmdMosDown.Image = CType(resources.GetObject("cmdMosDown.Image"), System.Drawing.Image)
    Me.cmdMosDown.Location = New System.Drawing.Point(135, 6)
    Me.cmdMosDown.Name = "cmdMosDown"
    Me.cmdMosDown.Size = New System.Drawing.Size(32, 37)
    Me.cmdMosDown.TabIndex = 37
    Me.cmdMosDown.UseVisualStyleBackColor = True
    '
    'cmdMosUp
    '
    Me.cmdMosUp.Image = CType(resources.GetObject("cmdMosUp.Image"), System.Drawing.Image)
    Me.cmdMosUp.Location = New System.Drawing.Point(96, 6)
    Me.cmdMosUp.Name = "cmdMosUp"
    Me.cmdMosUp.Size = New System.Drawing.Size(33, 37)
    Me.cmdMosUp.TabIndex = 36
    Me.cmdMosUp.UseVisualStyleBackColor = True
    '
    'dgvMos
    '
    Me.dgvMos.AllowUserToAddRows = False
    Me.dgvMos.AllowUserToDeleteRows = False
    Me.dgvMos.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dgvMos.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
    Me.dgvMos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvMos.Location = New System.Drawing.Point(5, 46)
    Me.dgvMos.Name = "dgvMos"
    Me.dgvMos.ReadOnly = True
    Me.dgvMos.Size = New System.Drawing.Size(162, 363)
    Me.dgvMos.TabIndex = 5
    '
    'Label82
    '
    Me.Label82.AutoSize = True
    Me.Label82.Location = New System.Drawing.Point(3, 30)
    Me.Label82.Name = "Label82"
    Me.Label82.Size = New System.Drawing.Size(38, 13)
    Me.Label82.TabIndex = 4
    Me.Label82.Text = "Name:"
    '
    'SplitContainer12
    '
    Me.SplitContainer12.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer12.Location = New System.Drawing.Point(0, 0)
    Me.SplitContainer12.Name = "SplitContainer12"
    Me.SplitContainer12.Orientation = System.Windows.Forms.Orientation.Horizontal
    '
    'SplitContainer12.Panel1
    '
    Me.SplitContainer12.Panel1.Controls.Add(Me.tbMosName)
    Me.SplitContainer12.Panel1.Controls.Add(Me.Label90)
    Me.SplitContainer12.Panel1.Controls.Add(Me.tbMosCond)
    Me.SplitContainer12.Panel1.Controls.Add(Me.Label83)
    Me.SplitContainer12.Panel1.Controls.Add(Me.tbMosTrigger)
    Me.SplitContainer12.Panel1.Controls.Add(Me.Label84)
    '
    'SplitContainer12.Panel2
    '
    Me.SplitContainer12.Panel2.Controls.Add(Me.SplitContainer13)
    Me.SplitContainer12.Size = New System.Drawing.Size(555, 414)
    Me.SplitContainer12.SplitterDistance = 145
    Me.SplitContainer12.TabIndex = 0
    '
    'tbMosName
    '
    Me.tbMosName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbMosName.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbMosName.ForeColor = System.Drawing.Color.Navy
    Me.tbMosName.Location = New System.Drawing.Point(96, 5)
    Me.tbMosName.Name = "tbMosName"
    Me.tbMosName.Size = New System.Drawing.Size(264, 20)
    Me.tbMosName.TabIndex = 18
    '
    'Label90
    '
    Me.Label90.AutoSize = True
    Me.Label90.Location = New System.Drawing.Point(6, 6)
    Me.Label90.Name = "Label90"
    Me.Label90.Size = New System.Drawing.Size(38, 13)
    Me.Label90.TabIndex = 17
    Me.Label90.Text = "Name:"
    '
    'tbMosCond
    '
    Me.tbMosCond.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbMosCond.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbMosCond.ForeColor = System.Drawing.Color.Maroon
    Me.tbMosCond.Location = New System.Drawing.Point(3, 72)
    Me.tbMosCond.Name = "tbMosCond"
    Me.tbMosCond.Size = New System.Drawing.Size(546, 70)
    Me.tbMosCond.TabIndex = 16
    Me.tbMosCond.Text = ""
    '
    'Label83
    '
    Me.Label83.AutoSize = True
    Me.Label83.Location = New System.Drawing.Point(6, 56)
    Me.Label83.Name = "Label83"
    Me.Label83.Size = New System.Drawing.Size(65, 13)
    Me.Label83.TabIndex = 15
    Me.Label83.Text = "Condition(s):"
    '
    'tbMosTrigger
    '
    Me.tbMosTrigger.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbMosTrigger.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbMosTrigger.ForeColor = System.Drawing.Color.Navy
    Me.tbMosTrigger.Location = New System.Drawing.Point(96, 31)
    Me.tbMosTrigger.Name = "tbMosTrigger"
    Me.tbMosTrigger.Size = New System.Drawing.Size(454, 20)
    Me.tbMosTrigger.TabIndex = 14
    '
    'Label84
    '
    Me.Label84.AutoSize = True
    Me.Label84.Location = New System.Drawing.Point(6, 34)
    Me.Label84.Name = "Label84"
    Me.Label84.Size = New System.Drawing.Size(83, 13)
    Me.Label84.TabIndex = 13
    Me.Label84.Text = "Trigger Label(s):"
    '
    'SplitContainer13
    '
    Me.SplitContainer13.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer13.Location = New System.Drawing.Point(0, 0)
    Me.SplitContainer13.Name = "SplitContainer13"
    '
    'SplitContainer13.Panel1
    '
    Me.SplitContainer13.Panel1.Controls.Add(Me.dgvMosAct)
    '
    'SplitContainer13.Panel2
    '
    Me.SplitContainer13.Panel2.Controls.Add(Me.cboMosActDir)
    Me.SplitContainer13.Panel2.Controls.Add(Me.Label86)
    Me.SplitContainer13.Panel2.Controls.Add(Me.cmdMosActDel)
    Me.SplitContainer13.Panel2.Controls.Add(Me.cmdMosActNew)
    Me.SplitContainer13.Panel2.Controls.Add(Me.cmdMosActDown)
    Me.SplitContainer13.Panel2.Controls.Add(Me.cmdMosActUp)
    Me.SplitContainer13.Panel2.Controls.Add(Me.Label85)
    Me.SplitContainer13.Panel2.Controls.Add(Me.tbMosActDescr)
    Me.SplitContainer13.Panel2.Controls.Add(Me.tbMosActArg)
    Me.SplitContainer13.Panel2.Controls.Add(Me.Label87)
    Me.SplitContainer13.Panel2.Controls.Add(Me.tbMosActId)
    Me.SplitContainer13.Panel2.Controls.Add(Me.Label88)
    Me.SplitContainer13.Panel2.Controls.Add(Me.tbMosActOp)
    Me.SplitContainer13.Panel2.Controls.Add(Me.Label89)
    Me.SplitContainer13.Size = New System.Drawing.Size(555, 265)
    Me.SplitContainer13.SplitterDistance = 107
    Me.SplitContainer13.TabIndex = 0
    '
    'dgvMosAct
    '
    Me.dgvMosAct.AllowUserToAddRows = False
    Me.dgvMosAct.AllowUserToDeleteRows = False
    Me.dgvMosAct.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
    Me.dgvMosAct.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvMosAct.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgvMosAct.Location = New System.Drawing.Point(0, 0)
    Me.dgvMosAct.Name = "dgvMosAct"
    Me.dgvMosAct.ReadOnly = True
    Me.dgvMosAct.Size = New System.Drawing.Size(107, 265)
    Me.dgvMosAct.TabIndex = 5
    '
    'cboMosActDir
    '
    Me.cboMosActDir.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboMosActDir.FormattingEnabled = True
    Me.cboMosActDir.Location = New System.Drawing.Point(64, 135)
    Me.cboMosActDir.Name = "cboMosActDir"
    Me.cboMosActDir.Size = New System.Drawing.Size(185, 21)
    Me.cboMosActDir.TabIndex = 39
    '
    'Label86
    '
    Me.Label86.AutoSize = True
    Me.Label86.Location = New System.Drawing.Point(2, 138)
    Me.Label86.Name = "Label86"
    Me.Label86.Size = New System.Drawing.Size(52, 13)
    Me.Label86.TabIndex = 38
    Me.Label86.Text = "Direction:"
    '
    'cmdMosActDel
    '
    Me.cmdMosActDel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdMosActDel.Location = New System.Drawing.Point(345, 30)
    Me.cmdMosActDel.Name = "cmdMosActDel"
    Me.cmdMosActDel.Size = New System.Drawing.Size(92, 25)
    Me.cmdMosActDel.TabIndex = 37
    Me.cmdMosActDel.Text = "&Delete"
    Me.cmdMosActDel.UseVisualStyleBackColor = True
    '
    'cmdMosActNew
    '
    Me.cmdMosActNew.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdMosActNew.Location = New System.Drawing.Point(247, 30)
    Me.cmdMosActNew.Name = "cmdMosActNew"
    Me.cmdMosActNew.Size = New System.Drawing.Size(92, 25)
    Me.cmdMosActNew.TabIndex = 36
    Me.cmdMosActNew.Text = "&New"
    Me.cmdMosActNew.UseVisualStyleBackColor = True
    '
    'cmdMosActDown
    '
    Me.cmdMosActDown.Image = CType(resources.GetObject("cmdMosActDown.Image"), System.Drawing.Image)
    Me.cmdMosActDown.Location = New System.Drawing.Point(174, 30)
    Me.cmdMosActDown.Name = "cmdMosActDown"
    Me.cmdMosActDown.Size = New System.Drawing.Size(32, 37)
    Me.cmdMosActDown.TabIndex = 35
    Me.cmdMosActDown.UseVisualStyleBackColor = True
    '
    'cmdMosActUp
    '
    Me.cmdMosActUp.Image = CType(resources.GetObject("cmdMosActUp.Image"), System.Drawing.Image)
    Me.cmdMosActUp.Location = New System.Drawing.Point(135, 30)
    Me.cmdMosActUp.Name = "cmdMosActUp"
    Me.cmdMosActUp.Size = New System.Drawing.Size(33, 37)
    Me.cmdMosActUp.TabIndex = 34
    Me.cmdMosActUp.UseVisualStyleBackColor = True
    '
    'Label85
    '
    Me.Label85.AutoSize = True
    Me.Label85.Location = New System.Drawing.Point(2, 160)
    Me.Label85.Name = "Label85"
    Me.Label85.Size = New System.Drawing.Size(52, 13)
    Me.Label85.TabIndex = 33
    Me.Label85.Text = "Remarks:"
    '
    'tbMosActDescr
    '
    Me.tbMosActDescr.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbMosActDescr.Location = New System.Drawing.Point(5, 176)
    Me.tbMosActDescr.Name = "tbMosActDescr"
    Me.tbMosActDescr.Size = New System.Drawing.Size(434, 86)
    Me.tbMosActDescr.TabIndex = 31
    Me.tbMosActDescr.Text = ""
    '
    'tbMosActArg
    '
    Me.tbMosActArg.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbMosActArg.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbMosActArg.ForeColor = System.Drawing.Color.Maroon
    Me.tbMosActArg.Location = New System.Drawing.Point(64, 73)
    Me.tbMosActArg.Name = "tbMosActArg"
    Me.tbMosActArg.Size = New System.Drawing.Size(373, 56)
    Me.tbMosActArg.TabIndex = 29
    Me.tbMosActArg.Text = ""
    '
    'Label87
    '
    Me.Label87.AutoSize = True
    Me.Label87.Location = New System.Drawing.Point(2, 77)
    Me.Label87.Name = "Label87"
    Me.Label87.Size = New System.Drawing.Size(66, 13)
    Me.Label87.TabIndex = 28
    Me.Label87.Text = "Argument(s):"
    '
    'tbMosActId
    '
    Me.tbMosActId.Location = New System.Drawing.Point(64, 30)
    Me.tbMosActId.Name = "tbMosActId"
    Me.tbMosActId.ReadOnly = True
    Me.tbMosActId.Size = New System.Drawing.Size(65, 20)
    Me.tbMosActId.TabIndex = 27
    '
    'Label88
    '
    Me.Label88.AutoSize = True
    Me.Label88.Location = New System.Drawing.Point(2, 33)
    Me.Label88.Name = "Label88"
    Me.Label88.Size = New System.Drawing.Size(47, 13)
    Me.Label88.TabIndex = 26
    Me.Label88.Text = "Position:"
    '
    'tbMosActOp
    '
    Me.tbMosActOp.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbMosActOp.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbMosActOp.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
    Me.tbMosActOp.Location = New System.Drawing.Point(64, 4)
    Me.tbMosActOp.Name = "tbMosActOp"
    Me.tbMosActOp.Size = New System.Drawing.Size(373, 20)
    Me.tbMosActOp.TabIndex = 25
    '
    'Label89
    '
    Me.Label89.AutoSize = True
    Me.Label89.Location = New System.Drawing.Point(2, 7)
    Me.Label89.Name = "Label89"
    Me.Label89.Size = New System.Drawing.Size(56, 13)
    Me.Label89.TabIndex = 24
    Me.Label89.Text = "Operation:"
    '
    'OpenFileDialog1
    '
    Me.OpenFileDialog1.FileName = "OpenFileDialog1"
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(636, 486)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(92, 25)
    Me.cmdCancel.TabIndex = 10
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'cmdOk
    '
    Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOk.Enabled = False
    Me.cmdOk.Location = New System.Drawing.Point(538, 486)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(92, 25)
    Me.cmdOk.TabIndex = 9
    Me.cmdOk.Text = "&Apply"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'cmdSetFile
    '
    Me.cmdSetFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdSetFile.Location = New System.Drawing.Point(712, 11)
    Me.cmdSetFile.Name = "cmdSetFile"
    Me.cmdSetFile.Size = New System.Drawing.Size(28, 20)
    Me.cmdSetFile.TabIndex = 13
    Me.cmdSetFile.Text = "..."
    Me.cmdSetFile.UseVisualStyleBackColor = True
    '
    'tbSettingsFile
    '
    Me.tbSettingsFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbSettingsFile.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbSettingsFile.Location = New System.Drawing.Point(78, 12)
    Me.tbSettingsFile.Name = "tbSettingsFile"
    Me.tbSettingsFile.ReadOnly = True
    Me.tbSettingsFile.Size = New System.Drawing.Size(628, 20)
    Me.tbSettingsFile.TabIndex = 12
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(4, 15)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(64, 13)
    Me.Label3.TabIndex = 11
    Me.Label3.Text = "Settings file:"
    '
    'Label17
    '
    Me.Label17.AutoSize = True
    Me.Label17.Location = New System.Drawing.Point(166, 34)
    Me.Label17.Name = "Label17"
    Me.Label17.Size = New System.Drawing.Size(63, 13)
    Me.Label17.TabIndex = 4
    Me.Label17.Text = "Description:"
    '
    'RichTextBox1
    '
    Me.RichTextBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.RichTextBox1.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.RichTextBox1.Location = New System.Drawing.Point(169, 121)
    Me.RichTextBox1.Name = "RichTextBox1"
    Me.RichTextBox1.Size = New System.Drawing.Size(418, 215)
    Me.RichTextBox1.TabIndex = 7
    Me.RichTextBox1.Text = ""
    '
    'Label18
    '
    Me.Label18.AutoSize = True
    Me.Label18.Location = New System.Drawing.Point(166, 105)
    Me.Label18.Name = "Label18"
    Me.Label18.Size = New System.Drawing.Size(264, 13)
    Me.Label18.TabIndex = 6
    Me.Label18.Text = "Variants (divided by the | sign, possibly with wildcard *):"
    '
    'TextBox1
    '
    Me.TextBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.TextBox1.Location = New System.Drawing.Point(235, 34)
    Me.TextBox1.Name = "TextBox1"
    Me.TextBox1.Size = New System.Drawing.Size(352, 20)
    Me.TextBox1.TabIndex = 5
    '
    'TextBox2
    '
    Me.TextBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.TextBox2.Location = New System.Drawing.Point(235, 8)
    Me.TextBox2.Name = "TextBox2"
    Me.TextBox2.Size = New System.Drawing.Size(352, 20)
    Me.TextBox2.TabIndex = 3
    '
    'Label19
    '
    Me.Label19.AutoSize = True
    Me.Label19.Location = New System.Drawing.Point(166, 8)
    Me.Label19.Name = "Label19"
    Me.Label19.Size = New System.Drawing.Size(38, 13)
    Me.Label19.TabIndex = 2
    Me.Label19.Text = "Name:"
    '
    'DataGridView1
    '
    Me.DataGridView1.AllowUserToAddRows = False
    Me.DataGridView1.AllowUserToDeleteRows = False
    Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.DataGridView1.Location = New System.Drawing.Point(6, 24)
    Me.DataGridView1.Name = "DataGridView1"
    Me.DataGridView1.ReadOnly = True
    Me.DataGridView1.Size = New System.Drawing.Size(148, 312)
    Me.DataGridView1.TabIndex = 1
    '
    'Label20
    '
    Me.Label20.AutoSize = True
    Me.Label20.Location = New System.Drawing.Point(3, 8)
    Me.Label20.Name = "Label20"
    Me.Label20.Size = New System.Drawing.Size(72, 13)
    Me.Label20.TabIndex = 0
    Me.Label20.Text = "PronounClass"
    '
    'cmdNew
    '
    Me.cmdNew.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.cmdNew.Location = New System.Drawing.Point(10, 486)
    Me.cmdNew.Name = "cmdNew"
    Me.cmdNew.Size = New System.Drawing.Size(92, 25)
    Me.cmdNew.TabIndex = 14
    Me.cmdNew.Text = "&New"
    Me.cmdNew.UseVisualStyleBackColor = True
    '
    'cmdDel
    '
    Me.cmdDel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.cmdDel.Location = New System.Drawing.Point(108, 486)
    Me.cmdDel.Name = "cmdDel"
    Me.cmdDel.Size = New System.Drawing.Size(92, 25)
    Me.cmdDel.TabIndex = 15
    Me.cmdDel.Text = "&Delete"
    Me.cmdDel.UseVisualStyleBackColor = True
    '
    'SplitContainer2
    '
    Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
    Me.SplitContainer2.Name = "SplitContainer2"
    '
    'SplitContainer2.Panel1
    '
    Me.SplitContainer2.Panel1.Controls.Add(Me.DataGridView2)
    Me.SplitContainer2.Panel1.Controls.Add(Me.Label32)
    '
    'SplitContainer2.Panel2
    '
    Me.SplitContainer2.Panel2.Controls.Add(Me.RichTextBox2)
    Me.SplitContainer2.Panel2.Controls.Add(Me.Label33)
    Me.SplitContainer2.Panel2.Controls.Add(Me.TextBox3)
    Me.SplitContainer2.Panel2.Controls.Add(Me.Label34)
    Me.SplitContainer2.Panel2.Controls.Add(Me.TextBox4)
    Me.SplitContainer2.Panel2.Controls.Add(Me.Label35)
    Me.SplitContainer2.Size = New System.Drawing.Size(728, 362)
    Me.SplitContainer2.SplitterDistance = 163
    Me.SplitContainer2.TabIndex = 0
    '
    'DataGridView2
    '
    Me.DataGridView2.AllowUserToAddRows = False
    Me.DataGridView2.AllowUserToDeleteRows = False
    Me.DataGridView2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.DataGridView2.Location = New System.Drawing.Point(6, 24)
    Me.DataGridView2.Name = "DataGridView2"
    Me.DataGridView2.ReadOnly = True
    Me.DataGridView2.Size = New System.Drawing.Size(154, 335)
    Me.DataGridView2.TabIndex = 1
    '
    'Label32
    '
    Me.Label32.AutoSize = True
    Me.Label32.Location = New System.Drawing.Point(3, 8)
    Me.Label32.Name = "Label32"
    Me.Label32.Size = New System.Drawing.Size(84, 13)
    Me.Label32.TabIndex = 0
    Me.Label32.Text = "NP feature type:"
    '
    'RichTextBox2
    '
    Me.RichTextBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.RichTextBox2.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.RichTextBox2.Location = New System.Drawing.Point(6, 77)
    Me.RichTextBox2.Name = "RichTextBox2"
    Me.RichTextBox2.Size = New System.Drawing.Size(547, 282)
    Me.RichTextBox2.TabIndex = 5
    Me.RichTextBox2.Text = ""
    '
    'Label33
    '
    Me.Label33.AutoSize = True
    Me.Label33.Location = New System.Drawing.Point(3, 61)
    Me.Label33.Name = "Label33"
    Me.Label33.Size = New System.Drawing.Size(163, 13)
    Me.Label33.TabIndex = 4
    Me.Label33.Text = "All possible variants, one per line:"
    '
    'TextBox3
    '
    Me.TextBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.TextBox3.Location = New System.Drawing.Point(72, 31)
    Me.TextBox3.Name = "TextBox3"
    Me.TextBox3.Size = New System.Drawing.Size(481, 20)
    Me.TextBox3.TabIndex = 3
    '
    'Label34
    '
    Me.Label34.AutoSize = True
    Me.Label34.Location = New System.Drawing.Point(3, 34)
    Me.Label34.Name = "Label34"
    Me.Label34.Size = New System.Drawing.Size(63, 13)
    Me.Label34.TabIndex = 2
    Me.Label34.Text = "Description:"
    '
    'TextBox4
    '
    Me.TextBox4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.TextBox4.Location = New System.Drawing.Point(72, 5)
    Me.TextBox4.Name = "TextBox4"
    Me.TextBox4.Size = New System.Drawing.Size(481, 20)
    Me.TextBox4.TabIndex = 1
    '
    'Label35
    '
    Me.Label35.AutoSize = True
    Me.Label35.Location = New System.Drawing.Point(3, 8)
    Me.Label35.Name = "Label35"
    Me.Label35.Size = New System.Drawing.Size(38, 13)
    Me.Label35.TabIndex = 0
    Me.Label35.Text = "Name:"
    '
    'DataGridView3
    '
    Me.DataGridView3.AllowUserToAddRows = False
    Me.DataGridView3.AllowUserToDeleteRows = False
    Me.DataGridView3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.DataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.DataGridView3.Location = New System.Drawing.Point(6, 24)
    Me.DataGridView3.Name = "DataGridView3"
    Me.DataGridView3.ReadOnly = True
    Me.DataGridView3.Size = New System.Drawing.Size(154, 387)
    Me.DataGridView3.TabIndex = 1
    '
    'Label73
    '
    Me.Label73.AutoSize = True
    Me.Label73.Location = New System.Drawing.Point(3, 8)
    Me.Label73.Name = "Label73"
    Me.Label73.Size = New System.Drawing.Size(84, 13)
    Me.Label73.TabIndex = 0
    Me.Label73.Text = "NP feature type:"
    '
    'RichTextBox3
    '
    Me.RichTextBox3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.RichTextBox3.Font = New System.Drawing.Font("Courier New", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.RichTextBox3.Location = New System.Drawing.Point(6, 77)
    Me.RichTextBox3.Name = "RichTextBox3"
    Me.RichTextBox3.Size = New System.Drawing.Size(551, 334)
    Me.RichTextBox3.TabIndex = 5
    Me.RichTextBox3.Text = ""
    '
    'Label74
    '
    Me.Label74.AutoSize = True
    Me.Label74.Location = New System.Drawing.Point(3, 61)
    Me.Label74.Name = "Label74"
    Me.Label74.Size = New System.Drawing.Size(163, 13)
    Me.Label74.TabIndex = 4
    Me.Label74.Text = "All possible variants, one per line:"
    '
    'TextBox5
    '
    Me.TextBox5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.TextBox5.Location = New System.Drawing.Point(72, 31)
    Me.TextBox5.Name = "TextBox5"
    Me.TextBox5.Size = New System.Drawing.Size(485, 20)
    Me.TextBox5.TabIndex = 3
    '
    'Label75
    '
    Me.Label75.AutoSize = True
    Me.Label75.Location = New System.Drawing.Point(3, 34)
    Me.Label75.Name = "Label75"
    Me.Label75.Size = New System.Drawing.Size(63, 13)
    Me.Label75.TabIndex = 2
    Me.Label75.Text = "Description:"
    '
    'TextBox6
    '
    Me.TextBox6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.TextBox6.Location = New System.Drawing.Point(72, 5)
    Me.TextBox6.Name = "TextBox6"
    Me.TextBox6.Size = New System.Drawing.Size(485, 20)
    Me.TextBox6.TabIndex = 1
    '
    'Label76
    '
    Me.Label76.AutoSize = True
    Me.Label76.Location = New System.Drawing.Point(3, 8)
    Me.Label76.Name = "Label76"
    Me.Label76.Size = New System.Drawing.Size(38, 13)
    Me.Label76.TabIndex = 0
    Me.Label76.Text = "Name:"
    '
    'frmSetting
    '
    Me.AcceptButton = Me.cmdOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(740, 536)
    Me.Controls.Add(Me.cmdDel)
    Me.Controls.Add(Me.cmdNew)
    Me.Controls.Add(Me.cmdSetFile)
    Me.Controls.Add(Me.tbSettingsFile)
    Me.Controls.Add(Me.Label3)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.cmdOk)
    Me.Controls.Add(Me.TabControl1)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmSetting"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
    Me.Text = "Adjust Settings for Cesax"
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.TabControl1.ResumeLayout(False)
    Me.tbGeneral.ResumeLayout(False)
    Me.tbGeneral.PerformLayout()
    Me.GroupBox3.ResumeLayout(False)
    Me.GroupBox3.PerformLayout()
    Me.GroupBox1.ResumeLayout(False)
    Me.GroupBox1.PerformLayout()
    Me.tpCoref.ResumeLayout(False)
    Me.tpCoref.PerformLayout()
    CType(Me.dgvCorefType, System.ComponentModel.ISupportInitialize).EndInit()
    Me.tpPhrase.ResumeLayout(False)
    Me.SplitContainer8.Panel1.ResumeLayout(False)
    Me.SplitContainer8.Panel2.ResumeLayout(False)
    CType(Me.SplitContainer8, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer8.ResumeLayout(False)
    Me.SplitContainer9.Panel1.ResumeLayout(False)
    Me.SplitContainer9.Panel1.PerformLayout()
    Me.SplitContainer9.Panel2.ResumeLayout(False)
    Me.SplitContainer9.Panel2.PerformLayout()
    CType(Me.SplitContainer9, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer9.ResumeLayout(False)
    Me.SplitContainer7.Panel1.ResumeLayout(False)
    Me.SplitContainer7.Panel1.PerformLayout()
    Me.SplitContainer7.Panel2.ResumeLayout(False)
    Me.SplitContainer7.Panel2.PerformLayout()
    CType(Me.SplitContainer7, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer7.ResumeLayout(False)
    CType(Me.dgvPhraseType, System.ComponentModel.ISupportInitialize).EndInit()
    Me.tpPronoun.ResumeLayout(False)
    Me.tpPronoun.PerformLayout()
    CType(Me.dgvPronoun, System.ComponentModel.ISupportInitialize).EndInit()
    Me.tpNPfeat.ResumeLayout(False)
    Me.SplitContainer1.Panel1.ResumeLayout(False)
    Me.SplitContainer1.Panel1.PerformLayout()
    Me.SplitContainer1.Panel2.ResumeLayout(False)
    Me.SplitContainer1.Panel2.PerformLayout()
    CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer1.ResumeLayout(False)
    CType(Me.dgvNPfeat, System.ComponentModel.ISupportInitialize).EndInit()
    Me.tpCons.ResumeLayout(False)
    Me.SplitContainer3.Panel1.ResumeLayout(False)
    Me.SplitContainer3.Panel1.PerformLayout()
    Me.SplitContainer3.Panel2.ResumeLayout(False)
    Me.SplitContainer3.Panel2.PerformLayout()
    CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer3.ResumeLayout(False)
    CType(Me.dgvCons, System.ComponentModel.ISupportInitialize).EndInit()
    Me.tpCat.ResumeLayout(False)
    Me.tpCat.PerformLayout()
    CType(Me.dgvCategory, System.ComponentModel.ISupportInitialize).EndInit()
    Me.tpStatus.ResumeLayout(False)
    Me.SplitContainer4.Panel1.ResumeLayout(False)
    Me.SplitContainer4.Panel1.PerformLayout()
    Me.SplitContainer4.Panel2.ResumeLayout(False)
    Me.SplitContainer4.Panel2.PerformLayout()
    CType(Me.SplitContainer4, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer4.ResumeLayout(False)
    CType(Me.dgvStatus, System.ComponentModel.ISupportInitialize).EndInit()
    Me.tpCharting.ResumeLayout(False)
    Me.SplitContainer5.Panel1.ResumeLayout(False)
    Me.SplitContainer5.Panel1.PerformLayout()
    Me.SplitContainer5.Panel2.ResumeLayout(False)
    Me.SplitContainer5.Panel2.PerformLayout()
    CType(Me.SplitContainer5, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer5.ResumeLayout(False)
    CType(Me.dgvTemplate, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer6.Panel1.ResumeLayout(False)
    Me.SplitContainer6.Panel2.ResumeLayout(False)
    Me.SplitContainer6.Panel2.PerformLayout()
    CType(Me.SplitContainer6, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer6.ResumeLayout(False)
    CType(Me.dgvTcell, System.ComponentModel.ISupportInitialize).EndInit()
    Me.tpXrel.ResumeLayout(False)
    Me.SplitContainer10.Panel1.ResumeLayout(False)
    Me.SplitContainer10.Panel1.PerformLayout()
    Me.SplitContainer10.Panel2.ResumeLayout(False)
    Me.SplitContainer10.Panel2.PerformLayout()
    CType(Me.SplitContainer10, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer10.ResumeLayout(False)
    CType(Me.dgvXrel, System.ComponentModel.ISupportInitialize).EndInit()
    Me.tpMos.ResumeLayout(False)
    Me.SplitContainer11.Panel1.ResumeLayout(False)
    Me.SplitContainer11.Panel1.PerformLayout()
    Me.SplitContainer11.Panel2.ResumeLayout(False)
    CType(Me.SplitContainer11, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer11.ResumeLayout(False)
    CType(Me.dgvMos, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer12.Panel1.ResumeLayout(False)
    Me.SplitContainer12.Panel1.PerformLayout()
    Me.SplitContainer12.Panel2.ResumeLayout(False)
    CType(Me.SplitContainer12, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer12.ResumeLayout(False)
    Me.SplitContainer13.Panel1.ResumeLayout(False)
    Me.SplitContainer13.Panel2.ResumeLayout(False)
    Me.SplitContainer13.Panel2.PerformLayout()
    CType(Me.SplitContainer13, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer13.ResumeLayout(False)
    CType(Me.dgvMosAct, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer2.Panel1.ResumeLayout(False)
    Me.SplitContainer2.Panel1.PerformLayout()
    Me.SplitContainer2.Panel2.ResumeLayout(False)
    Me.SplitContainer2.Panel2.PerformLayout()
    CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
    Me.SplitContainer2.ResumeLayout(False)
    CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.bsProjType, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.bsProjEdit, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
  Friend WithEvents tbGeneral As System.Windows.Forms.TabPage
  Friend WithEvents cmdWorkDir As System.Windows.Forms.Button
  Friend WithEvents tbWorkDir As System.Windows.Forms.TextBox
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
  Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
  Friend WithEvents bsProjType As System.Windows.Forms.BindingSource
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents cmdSetFile As System.Windows.Forms.Button
  Friend WithEvents tbSettingsFile As System.Windows.Forms.TextBox
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents bsProjEdit As System.Windows.Forms.BindingSource
  Friend WithEvents tbDriveChange As System.Windows.Forms.TextBox
  Friend WithEvents Label14 As System.Windows.Forms.Label
  Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
  Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
  Friend WithEvents chbUserProPGN As System.Windows.Forms.CheckBox
  Friend WithEvents chbDebugging As System.Windows.Forms.CheckBox
  Friend WithEvents chbShowCODE As System.Windows.Forms.CheckBox
  Friend WithEvents tpCoref As System.Windows.Forms.TabPage
  Friend WithEvents tpPhrase As System.Windows.Forms.TabPage
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents dgvCorefType As System.Windows.Forms.DataGridView
  Friend WithEvents tbCorefDescr As System.Windows.Forms.RichTextBox
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents tbCorefName As System.Windows.Forms.TextBox
  Friend WithEvents Label5 As System.Windows.Forms.Label
  Friend WithEvents tbPhraseNode As System.Windows.Forms.TextBox
  Friend WithEvents Label6 As System.Windows.Forms.Label
  Friend WithEvents dgvPhraseType As System.Windows.Forms.DataGridView
  Friend WithEvents Label7 As System.Windows.Forms.Label
  Friend WithEvents tbPhraseChild As System.Windows.Forms.TextBox
  Friend WithEvents Label8 As System.Windows.Forms.Label
  Friend WithEvents cboPhraseType As System.Windows.Forms.ComboBox
  Friend WithEvents Label9 As System.Windows.Forms.Label
  Friend WithEvents Label10 As System.Windows.Forms.Label
  Friend WithEvents cboPhraseTarget As System.Windows.Forms.ComboBox
  Friend WithEvents cboColor As System.Windows.Forms.ComboBox
  Friend WithEvents Label11 As System.Windows.Forms.Label
  Friend WithEvents tpPronoun As System.Windows.Forms.TabPage
  Friend WithEvents Label16 As System.Windows.Forms.Label
  Friend WithEvents tbProDescr As System.Windows.Forms.TextBox
  Friend WithEvents Label15 As System.Windows.Forms.Label
  Friend WithEvents tbProName As System.Windows.Forms.TextBox
  Friend WithEvents Label13 As System.Windows.Forms.Label
  Friend WithEvents dgvPronoun As System.Windows.Forms.DataGridView
  Friend WithEvents Label12 As System.Windows.Forms.Label
  Friend WithEvents tbProVarOE As System.Windows.Forms.RichTextBox
  Friend WithEvents tbProPGN As System.Windows.Forms.TextBox
  Friend WithEvents Label21 As System.Windows.Forms.Label
  Friend WithEvents Label17 As System.Windows.Forms.Label
  Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
  Friend WithEvents Label18 As System.Windows.Forms.Label
  Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
  Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
  Friend WithEvents Label19 As System.Windows.Forms.Label
  Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
  Friend WithEvents Label20 As System.Windows.Forms.Label
  Friend WithEvents tbProNotes As System.Windows.Forms.RichTextBox
  Friend WithEvents Label22 As System.Windows.Forms.Label
  Friend WithEvents Label26 As System.Windows.Forms.Label
  Friend WithEvents tbProVarMBE As System.Windows.Forms.RichTextBox
  Friend WithEvents Label25 As System.Windows.Forms.Label
  Friend WithEvents tbProVarEmodE As System.Windows.Forms.RichTextBox
  Friend WithEvents Label24 As System.Windows.Forms.Label
  Friend WithEvents tbProVarME As System.Windows.Forms.RichTextBox
  Friend WithEvents Label23 As System.Windows.Forms.Label
  Friend WithEvents cmdPeriodDef As System.Windows.Forms.Button
  Friend WithEvents tbPeriodDef As System.Windows.Forms.TextBox
  Friend WithEvents Label27 As System.Windows.Forms.Label
  Friend WithEvents tpNPfeat As System.Windows.Forms.TabPage
  Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
  Friend WithEvents Label28 As System.Windows.Forms.Label
  Friend WithEvents tbNPfeatVariants As System.Windows.Forms.RichTextBox
  Friend WithEvents Label31 As System.Windows.Forms.Label
  Friend WithEvents tbNPfeatDescr As System.Windows.Forms.TextBox
  Friend WithEvents Label30 As System.Windows.Forms.Label
  Friend WithEvents tbNPfeatName As System.Windows.Forms.TextBox
  Friend WithEvents Label29 As System.Windows.Forms.Label
  Friend WithEvents cmdNew As System.Windows.Forms.Button
  Friend WithEvents cmdDel As System.Windows.Forms.Button
  Friend WithEvents dgvNPfeat As System.Windows.Forms.DataGridView
  Friend WithEvents tpCons As System.Windows.Forms.TabPage
  Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
  Friend WithEvents dgvCons As System.Windows.Forms.DataGridView
  Friend WithEvents Label36 As System.Windows.Forms.Label
  Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
  Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
  Friend WithEvents Label32 As System.Windows.Forms.Label
  Friend WithEvents RichTextBox2 As System.Windows.Forms.RichTextBox
  Friend WithEvents Label33 As System.Windows.Forms.Label
  Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
  Friend WithEvents Label34 As System.Windows.Forms.Label
  Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
  Friend WithEvents Label35 As System.Windows.Forms.Label
  Friend WithEvents tbConsDescr As System.Windows.Forms.RichTextBox
  Friend WithEvents Label40 As System.Windows.Forms.Label
  Friend WithEvents tbConsMult As System.Windows.Forms.TextBox
  Friend WithEvents Label39 As System.Windows.Forms.Label
  Friend WithEvents tbConsLevel As System.Windows.Forms.TextBox
  Friend WithEvents Label38 As System.Windows.Forms.Label
  Friend WithEvents tbConsName As System.Windows.Forms.TextBox
  Friend WithEvents Label37 As System.Windows.Forms.Label
  Friend WithEvents tbUserName As System.Windows.Forms.TextBox
  Friend WithEvents Label41 As System.Windows.Forms.Label
  Friend WithEvents cmdChainDict As System.Windows.Forms.Button
  Friend WithEvents tbChainDict As System.Windows.Forms.TextBox
  Friend WithEvents Label42 As System.Windows.Forms.Label
  Friend WithEvents tbMaxIPdist As System.Windows.Forms.TextBox
  Friend WithEvents Label43 As System.Windows.Forms.Label
  Friend WithEvents tbProfile As System.Windows.Forms.TextBox
  Friend WithEvents Label44 As System.Windows.Forms.Label
  Friend WithEvents chbShowAllAnt As System.Windows.Forms.CheckBox
  Friend WithEvents cmdCnsDown As System.Windows.Forms.Button
  Friend WithEvents cmdCnsUp As System.Windows.Forms.Button
  Friend WithEvents tbAbsMaxIPdist As System.Windows.Forms.TextBox
  Friend WithEvents Label45 As System.Windows.Forms.Label
  Friend WithEvents tpCat As System.Windows.Forms.TabPage
  Friend WithEvents Label47 As System.Windows.Forms.Label
  Friend WithEvents tbCatMBE As System.Windows.Forms.RichTextBox
  Friend WithEvents Label48 As System.Windows.Forms.Label
  Friend WithEvents tbCatEmodE As System.Windows.Forms.RichTextBox
  Friend WithEvents Label49 As System.Windows.Forms.Label
  Friend WithEvents tbCatME As System.Windows.Forms.RichTextBox
  Friend WithEvents Label50 As System.Windows.Forms.Label
  Friend WithEvents tbCatNotes As System.Windows.Forms.RichTextBox
  Friend WithEvents Label51 As System.Windows.Forms.Label
  Friend WithEvents tbCatOE As System.Windows.Forms.RichTextBox
  Friend WithEvents Label52 As System.Windows.Forms.Label
  Friend WithEvents tbCatDescr As System.Windows.Forms.TextBox
  Friend WithEvents Label53 As System.Windows.Forms.Label
  Friend WithEvents tbCatName As System.Windows.Forms.TextBox
  Friend WithEvents Label54 As System.Windows.Forms.Label
  Friend WithEvents dgvCategory As System.Windows.Forms.DataGridView
  Friend WithEvents Label46 As System.Windows.Forms.Label
  Friend WithEvents Label56 As System.Windows.Forms.Label
  Friend WithEvents tbCatType As System.Windows.Forms.TextBox
  Friend WithEvents Label55 As System.Windows.Forms.Label
  Friend WithEvents tbCorpusStudio As System.Windows.Forms.TextBox
  Friend WithEvents Label57 As System.Windows.Forms.Label
  Friend WithEvents cmdCorpusStudio As System.Windows.Forms.Button
  Friend WithEvents chbCheckCns As System.Windows.Forms.CheckBox
  Friend WithEvents chbCheckCat As System.Windows.Forms.CheckBox
  Friend WithEvents tpStatus As System.Windows.Forms.TabPage
  Friend WithEvents SplitContainer4 As System.Windows.Forms.SplitContainer
  Friend WithEvents dgvStatus As System.Windows.Forms.DataGridView
  Friend WithEvents Label58 As System.Windows.Forms.Label
  Friend WithEvents tbStatusDesc As System.Windows.Forms.RichTextBox
  Friend WithEvents Label59 As System.Windows.Forms.Label
  Friend WithEvents tbStatusName As System.Windows.Forms.TextBox
  Friend WithEvents Label60 As System.Windows.Forms.Label
  Friend WithEvents tbLogBase As System.Windows.Forms.TextBox
  Friend WithEvents Label61 As System.Windows.Forms.Label
  Friend WithEvents tpCharting As System.Windows.Forms.TabPage
  Friend WithEvents SplitContainer5 As System.Windows.Forms.SplitContainer
  Friend WithEvents dgvTemplate As System.Windows.Forms.DataGridView
  Friend WithEvents Label62 As System.Windows.Forms.Label
  Friend WithEvents SplitContainer6 As System.Windows.Forms.SplitContainer
  Friend WithEvents dgvTcell As System.Windows.Forms.DataGridView
  Friend WithEvents tbTemplateName As System.Windows.Forms.TextBox
  Friend WithEvents Label63 As System.Windows.Forms.Label
  Friend WithEvents tbTcellDescr As System.Windows.Forms.RichTextBox
  Friend WithEvents Label67 As System.Windows.Forms.Label
  Friend WithEvents tbTcellContent As System.Windows.Forms.RichTextBox
  Friend WithEvents Label66 As System.Windows.Forms.Label
  Friend WithEvents tbTcellId As System.Windows.Forms.TextBox
  Friend WithEvents Label65 As System.Windows.Forms.Label
  Friend WithEvents tbTcellName As System.Windows.Forms.TextBox
  Friend WithEvents Label64 As System.Windows.Forms.Label
  Friend WithEvents Label68 As System.Windows.Forms.Label
  Friend WithEvents tbTcellEnv As System.Windows.Forms.TextBox
  Friend WithEvents cmdCellDown As System.Windows.Forms.Button
  Friend WithEvents cmdCellUp As System.Windows.Forms.Button
  Friend WithEvents tbTemplateDescr As System.Windows.Forms.RichTextBox
  Friend WithEvents Label69 As System.Windows.Forms.Label
  Friend WithEvents cmdCellDel As System.Windows.Forms.Button
  Friend WithEvents cmdCellNew As System.Windows.Forms.Button
  Friend WithEvents SplitContainer7 As System.Windows.Forms.SplitContainer
  Friend WithEvents tbPhraseName As System.Windows.Forms.TextBox
  Friend WithEvents Label70 As System.Windows.Forms.Label
  Friend WithEvents SplitContainer8 As System.Windows.Forms.SplitContainer
  Friend WithEvents SplitContainer9 As System.Windows.Forms.SplitContainer
  Friend WithEvents tbPhraseInclude As System.Windows.Forms.RichTextBox
  Friend WithEvents Label71 As System.Windows.Forms.Label
  Friend WithEvents tbPhraseExclude As System.Windows.Forms.RichTextBox
  Friend WithEvents Label72 As System.Windows.Forms.Label
  Friend WithEvents tpXrel As System.Windows.Forms.TabPage
  Friend WithEvents SplitContainer10 As System.Windows.Forms.SplitContainer
  Friend WithEvents Label77 As System.Windows.Forms.Label
  Friend WithEvents dgvXrel As System.Windows.Forms.DataGridView
  Friend WithEvents tbXrelDescr As System.Windows.Forms.RichTextBox
  Friend WithEvents Label78 As System.Windows.Forms.Label
  Friend WithEvents Label79 As System.Windows.Forms.Label
  Friend WithEvents tbXrelName As System.Windows.Forms.TextBox
  Friend WithEvents Label80 As System.Windows.Forms.Label
  Friend WithEvents DataGridView3 As System.Windows.Forms.DataGridView
  Friend WithEvents Label73 As System.Windows.Forms.Label
  Friend WithEvents RichTextBox3 As System.Windows.Forms.RichTextBox
  Friend WithEvents Label74 As System.Windows.Forms.Label
  Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
  Friend WithEvents Label75 As System.Windows.Forms.Label
  Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
  Friend WithEvents Label76 As System.Windows.Forms.Label
  Friend WithEvents cboXrelType As System.Windows.Forms.ComboBox
  Friend WithEvents tbXrelXname As System.Windows.Forms.RichTextBox
  Friend WithEvents Label81 As System.Windows.Forms.Label
  Friend WithEvents tpMos As System.Windows.Forms.TabPage
  Friend WithEvents SplitContainer11 As System.Windows.Forms.SplitContainer
  Friend WithEvents dgvMos As System.Windows.Forms.DataGridView
  Friend WithEvents Label82 As System.Windows.Forms.Label
  Friend WithEvents SplitContainer12 As System.Windows.Forms.SplitContainer
  Friend WithEvents tbMosCond As System.Windows.Forms.RichTextBox
  Friend WithEvents Label83 As System.Windows.Forms.Label
  Friend WithEvents tbMosTrigger As System.Windows.Forms.TextBox
  Friend WithEvents Label84 As System.Windows.Forms.Label
  Friend WithEvents SplitContainer13 As System.Windows.Forms.SplitContainer
  Friend WithEvents dgvMosAct As System.Windows.Forms.DataGridView
  Friend WithEvents cmdMosActDel As System.Windows.Forms.Button
  Friend WithEvents cmdMosActNew As System.Windows.Forms.Button
  Friend WithEvents cmdMosActDown As System.Windows.Forms.Button
  Friend WithEvents cmdMosActUp As System.Windows.Forms.Button
  Friend WithEvents Label85 As System.Windows.Forms.Label
  Friend WithEvents tbMosActDescr As System.Windows.Forms.RichTextBox
  Friend WithEvents tbMosActArg As System.Windows.Forms.RichTextBox
  Friend WithEvents Label87 As System.Windows.Forms.Label
  Friend WithEvents tbMosActId As System.Windows.Forms.TextBox
  Friend WithEvents Label88 As System.Windows.Forms.Label
  Friend WithEvents tbMosActOp As System.Windows.Forms.TextBox
  Friend WithEvents Label89 As System.Windows.Forms.Label
  Friend WithEvents cboMosActDir As System.Windows.Forms.ComboBox
  Friend WithEvents Label86 As System.Windows.Forms.Label
  Friend WithEvents tbMosName As System.Windows.Forms.TextBox
  Friend WithEvents Label90 As System.Windows.Forms.Label
  Friend WithEvents cmdMosDown As System.Windows.Forms.Button
  Friend WithEvents cmdMosUp As System.Windows.Forms.Button
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents chbStatusAutoNotes As System.Windows.Forms.CheckBox
  Friend WithEvents chbStatusAutoChanged As System.Windows.Forms.CheckBox
  Friend WithEvents chbGenActionInet As System.Windows.Forms.CheckBox
  Friend WithEvents chbUpdStartup As System.Windows.Forms.CheckBox
End Class
