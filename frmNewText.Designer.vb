<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNewText
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNewText))
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
    Me.cmdOk = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.tbGenDirName = New System.Windows.Forms.TextBox()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.tbGenFileName = New System.Windows.Forms.TextBox()
    Me.cmdGenDir = New System.Windows.Forms.Button()
    Me.TabControl1 = New System.Windows.Forms.TabControl()
    Me.tpInfo = New System.Windows.Forms.TabPage()
    Me.GroupBox1 = New System.Windows.Forms.GroupBox()
    Me.tbGenGenre = New System.Windows.Forms.TextBox()
    Me.Label9 = New System.Windows.Forms.Label()
    Me.Label6 = New System.Windows.Forms.Label()
    Me.tbGenSource = New System.Windows.Forms.RichTextBox()
    Me.Label5 = New System.Windows.Forms.Label()
    Me.tbGenTitle = New System.Windows.Forms.TextBox()
    Me.Label3 = New System.Windows.Forms.Label()
    Me.tbGenDistributor = New System.Windows.Forms.TextBox()
    Me.Label4 = New System.Windows.Forms.Label()
    Me.tbGenSubType = New System.Windows.Forms.TextBox()
    Me.Label71 = New System.Windows.Forms.Label()
    Me.tbGenCreaManu = New System.Windows.Forms.TextBox()
    Me.Label69 = New System.Windows.Forms.Label()
    Me.tbGenCreaOrig = New System.Windows.Forms.TextBox()
    Me.Label68 = New System.Windows.Forms.Label()
    Me.GroupBox2 = New System.Windows.Forms.GroupBox()
    Me.Label7 = New System.Windows.Forms.Label()
    Me.tbGenLngName = New System.Windows.Forms.TextBox()
    Me.Label70 = New System.Windows.Forms.Label()
    Me.tbGenLngId = New System.Windows.Forms.TextBox()
    Me.Label67 = New System.Windows.Forms.Label()
    Me.tbGenEditor = New System.Windows.Forms.RichTextBox()
    Me.Label66 = New System.Windows.Forms.Label()
    Me.tbGenAuthor = New System.Windows.Forms.TextBox()
    Me.Label65 = New System.Windows.Forms.Label()
    Me.tpText = New System.Windows.Forms.TabPage()
    Me.Label8 = New System.Windows.Forms.Label()
    Me.tbGenText = New System.Windows.Forms.RichTextBox()
    Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
    Me.Label10 = New System.Windows.Forms.Label()
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.StatusStrip1.SuspendLayout()
    Me.TabControl1.SuspendLayout()
    Me.tpInfo.SuspendLayout()
    Me.GroupBox1.SuspendLayout()
    Me.GroupBox2.SuspendLayout()
    Me.tpText.SuspendLayout()
    Me.SuspendLayout()
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 546)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(803, 22)
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
    'cmdOk
    '
    Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOk.Location = New System.Drawing.Point(721, 520)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(75, 23)
    Me.cmdOk.TabIndex = 1
    Me.cmdOk.Text = "&Ok"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(640, 520)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 2
    Me.cmdCancel.Text = "&Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(12, 35)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(49, 13)
    Me.Label1.TabIndex = 3
    Me.Label1.Text = "Directory"
    '
    'tbGenDirName
    '
    Me.tbGenDirName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbGenDirName.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbGenDirName.Location = New System.Drawing.Point(70, 32)
    Me.tbGenDirName.Name = "tbGenDirName"
    Me.tbGenDirName.ReadOnly = True
    Me.tbGenDirName.Size = New System.Drawing.Size(661, 20)
    Me.tbGenDirName.TabIndex = 4
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(12, 9)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(52, 13)
    Me.Label2.TabIndex = 5
    Me.Label2.Text = "Filename:"
    '
    'tbGenFileName
    '
    Me.tbGenFileName.Location = New System.Drawing.Point(70, 6)
    Me.tbGenFileName.Name = "tbGenFileName"
    Me.tbGenFileName.Size = New System.Drawing.Size(256, 20)
    Me.tbGenFileName.TabIndex = 1
    '
    'cmdGenDir
    '
    Me.cmdGenDir.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdGenDir.Location = New System.Drawing.Point(737, 32)
    Me.cmdGenDir.Name = "cmdGenDir"
    Me.cmdGenDir.Size = New System.Drawing.Size(59, 20)
    Me.cmdGenDir.TabIndex = 2
    Me.cmdGenDir.Text = "..."
    Me.cmdGenDir.UseVisualStyleBackColor = True
    '
    'TabControl1
    '
    Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.TabControl1.Controls.Add(Me.tpInfo)
    Me.TabControl1.Controls.Add(Me.tpText)
    Me.TabControl1.Location = New System.Drawing.Point(15, 58)
    Me.TabControl1.Name = "TabControl1"
    Me.TabControl1.SelectedIndex = 0
    Me.TabControl1.Size = New System.Drawing.Size(781, 456)
    Me.TabControl1.TabIndex = 13
    '
    'tpInfo
    '
    Me.tpInfo.Controls.Add(Me.GroupBox1)
    Me.tpInfo.Location = New System.Drawing.Point(4, 22)
    Me.tpInfo.Name = "tpInfo"
    Me.tpInfo.Padding = New System.Windows.Forms.Padding(3)
    Me.tpInfo.Size = New System.Drawing.Size(773, 430)
    Me.tpInfo.TabIndex = 0
    Me.tpInfo.Text = "Information"
    Me.tpInfo.UseVisualStyleBackColor = True
    '
    'GroupBox1
    '
    Me.GroupBox1.Controls.Add(Me.tbGenGenre)
    Me.GroupBox1.Controls.Add(Me.Label9)
    Me.GroupBox1.Controls.Add(Me.Label6)
    Me.GroupBox1.Controls.Add(Me.tbGenSource)
    Me.GroupBox1.Controls.Add(Me.Label5)
    Me.GroupBox1.Controls.Add(Me.tbGenTitle)
    Me.GroupBox1.Controls.Add(Me.Label3)
    Me.GroupBox1.Controls.Add(Me.tbGenDistributor)
    Me.GroupBox1.Controls.Add(Me.Label4)
    Me.GroupBox1.Controls.Add(Me.tbGenSubType)
    Me.GroupBox1.Controls.Add(Me.Label71)
    Me.GroupBox1.Controls.Add(Me.tbGenCreaManu)
    Me.GroupBox1.Controls.Add(Me.Label69)
    Me.GroupBox1.Controls.Add(Me.tbGenCreaOrig)
    Me.GroupBox1.Controls.Add(Me.Label68)
    Me.GroupBox1.Controls.Add(Me.GroupBox2)
    Me.GroupBox1.Controls.Add(Me.tbGenEditor)
    Me.GroupBox1.Controls.Add(Me.Label66)
    Me.GroupBox1.Controls.Add(Me.tbGenAuthor)
    Me.GroupBox1.Controls.Add(Me.Label65)
    Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(767, 424)
    Me.GroupBox1.TabIndex = 9
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Text information"
    '
    'tbGenGenre
    '
    Me.tbGenGenre.Location = New System.Drawing.Point(120, 211)
    Me.tbGenGenre.Name = "tbGenGenre"
    Me.tbGenGenre.Size = New System.Drawing.Size(138, 20)
    Me.tbGenGenre.TabIndex = 10
    '
    'Label9
    '
    Me.Label9.AutoSize = True
    Me.Label9.Location = New System.Drawing.Point(15, 214)
    Me.Label9.Name = "Label9"
    Me.Label9.Size = New System.Drawing.Size(39, 13)
    Me.Label9.TabIndex = 33
    Me.Label9.Text = "Genre:"
    '
    'Label6
    '
    Me.Label6.AutoSize = True
    Me.Label6.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
    Me.Label6.Location = New System.Drawing.Point(264, 188)
    Me.Label6.Name = "Label6"
    Me.Label6.Size = New System.Drawing.Size(350, 13)
    Me.Label6.TabIndex = 32
    Me.Label6.Text = "(Use a time-period or genre abbreviation. Keep empty if this is not known)"
    '
    'tbGenSource
    '
    Me.tbGenSource.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbGenSource.Location = New System.Drawing.Point(75, 72)
    Me.tbGenSource.Name = "tbGenSource"
    Me.tbGenSource.Size = New System.Drawing.Size(666, 55)
    Me.tbGenSource.TabIndex = 5
    Me.tbGenSource.Text = ""
    '
    'Label5
    '
    Me.Label5.AutoSize = True
    Me.Label5.Location = New System.Drawing.Point(15, 75)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(41, 13)
    Me.Label5.TabIndex = 30
    Me.Label5.Text = "Source"
    '
    'tbGenTitle
    '
    Me.tbGenTitle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbGenTitle.Location = New System.Drawing.Point(75, 20)
    Me.tbGenTitle.Name = "tbGenTitle"
    Me.tbGenTitle.Size = New System.Drawing.Size(666, 20)
    Me.tbGenTitle.TabIndex = 3
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(15, 23)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(27, 13)
    Me.Label3.TabIndex = 26
    Me.Label3.Text = "Title"
    '
    'tbGenDistributor
    '
    Me.tbGenDistributor.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbGenDistributor.Location = New System.Drawing.Point(75, 46)
    Me.tbGenDistributor.Name = "tbGenDistributor"
    Me.tbGenDistributor.Size = New System.Drawing.Size(666, 20)
    Me.tbGenDistributor.TabIndex = 4
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(15, 49)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(54, 13)
    Me.Label4.TabIndex = 28
    Me.Label4.Text = "Distributor"
    '
    'tbGenSubType
    '
    Me.tbGenSubType.Location = New System.Drawing.Point(120, 185)
    Me.tbGenSubType.Name = "tbGenSubType"
    Me.tbGenSubType.Size = New System.Drawing.Size(138, 20)
    Me.tbGenSubType.TabIndex = 9
    '
    'Label71
    '
    Me.Label71.AutoSize = True
    Me.Label71.Location = New System.Drawing.Point(15, 188)
    Me.Label71.Name = "Label71"
    Me.Label71.Size = New System.Drawing.Size(90, 13)
    Me.Label71.TabIndex = 24
    Me.Label71.Text = "Subtype category"
    '
    'tbGenCreaManu
    '
    Me.tbGenCreaManu.Location = New System.Drawing.Point(373, 159)
    Me.tbGenCreaManu.Name = "tbGenCreaManu"
    Me.tbGenCreaManu.Size = New System.Drawing.Size(138, 20)
    Me.tbGenCreaManu.TabIndex = 8
    '
    'Label69
    '
    Me.Label69.AutoSize = True
    Me.Label69.Location = New System.Drawing.Point(268, 162)
    Me.Label69.Name = "Label69"
    Me.Label69.Size = New System.Drawing.Size(99, 13)
    Me.Label69.TabIndex = 22
    Me.Label69.Text = "Date of manuscript:"
    '
    'tbGenCreaOrig
    '
    Me.tbGenCreaOrig.Location = New System.Drawing.Point(120, 159)
    Me.tbGenCreaOrig.Name = "tbGenCreaOrig"
    Me.tbGenCreaOrig.Size = New System.Drawing.Size(138, 20)
    Me.tbGenCreaOrig.TabIndex = 7
    '
    'Label68
    '
    Me.Label68.AutoSize = True
    Me.Label68.Location = New System.Drawing.Point(15, 162)
    Me.Label68.Name = "Label68"
    Me.Label68.Size = New System.Drawing.Size(81, 13)
    Me.Label68.TabIndex = 20
    Me.Label68.Text = "Date of original:"
    '
    'GroupBox2
    '
    Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GroupBox2.Controls.Add(Me.Label7)
    Me.GroupBox2.Controls.Add(Me.tbGenLngName)
    Me.GroupBox2.Controls.Add(Me.Label70)
    Me.GroupBox2.Controls.Add(Me.tbGenLngId)
    Me.GroupBox2.Controls.Add(Me.Label67)
    Me.GroupBox2.Location = New System.Drawing.Point(18, 320)
    Me.GroupBox2.Name = "GroupBox2"
    Me.GroupBox2.Size = New System.Drawing.Size(723, 83)
    Me.GroupBox2.TabIndex = 19
    Me.GroupBox2.TabStop = False
    Me.GroupBox2.Text = "Language information"
    '
    'Label7
    '
    Me.Label7.AutoSize = True
    Me.Label7.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
    Me.Label7.Location = New System.Drawing.Point(246, 22)
    Me.Label7.Name = "Label7"
    Me.Label7.Size = New System.Drawing.Size(265, 13)
    Me.Label7.TabIndex = 33
    Me.Label7.Text = "Three-letter code or extension (e.g: Eng_hist, Che_Lat)"
    '
    'tbGenLngName
    '
    Me.tbGenLngName.Location = New System.Drawing.Point(102, 45)
    Me.tbGenLngName.Name = "tbGenLngName"
    Me.tbGenLngName.Size = New System.Drawing.Size(314, 20)
    Me.tbGenLngName.TabIndex = 13
    '
    'Label70
    '
    Me.Label70.AutoSize = True
    Me.Label70.Location = New System.Drawing.Point(5, 48)
    Me.Label70.Name = "Label70"
    Me.Label70.Size = New System.Drawing.Size(84, 13)
    Me.Label70.TabIndex = 8
    Me.Label70.Text = "Language name"
    '
    'tbGenLngId
    '
    Me.tbGenLngId.Location = New System.Drawing.Point(102, 19)
    Me.tbGenLngId.Name = "tbGenLngId"
    Me.tbGenLngId.Size = New System.Drawing.Size(138, 20)
    Me.tbGenLngId.TabIndex = 12
    '
    'Label67
    '
    Me.Label67.AutoSize = True
    Me.Label67.Location = New System.Drawing.Point(5, 22)
    Me.Label67.Name = "Label67"
    Me.Label67.Size = New System.Drawing.Size(88, 13)
    Me.Label67.TabIndex = 6
    Me.Label67.Text = "Ethnologue code"
    '
    'tbGenEditor
    '
    Me.tbGenEditor.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbGenEditor.Location = New System.Drawing.Point(74, 237)
    Me.tbGenEditor.Name = "tbGenEditor"
    Me.tbGenEditor.Size = New System.Drawing.Size(667, 65)
    Me.tbGenEditor.TabIndex = 11
    Me.tbGenEditor.Text = ""
    '
    'Label66
    '
    Me.Label66.AutoSize = True
    Me.Label66.Location = New System.Drawing.Point(15, 240)
    Me.Label66.Name = "Label66"
    Me.Label66.Size = New System.Drawing.Size(45, 13)
    Me.Label66.TabIndex = 17
    Me.Label66.Text = "Editor(s)"
    '
    'tbGenAuthor
    '
    Me.tbGenAuthor.Location = New System.Drawing.Point(75, 133)
    Me.tbGenAuthor.Name = "tbGenAuthor"
    Me.tbGenAuthor.Size = New System.Drawing.Size(436, 20)
    Me.tbGenAuthor.TabIndex = 6
    '
    'Label65
    '
    Me.Label65.AutoSize = True
    Me.Label65.Location = New System.Drawing.Point(15, 136)
    Me.Label65.Name = "Label65"
    Me.Label65.Size = New System.Drawing.Size(38, 13)
    Me.Label65.TabIndex = 15
    Me.Label65.Text = "Author"
    '
    'tpText
    '
    Me.tpText.Controls.Add(Me.Label8)
    Me.tpText.Controls.Add(Me.tbGenText)
    Me.tpText.Location = New System.Drawing.Point(4, 22)
    Me.tpText.Name = "tpText"
    Me.tpText.Padding = New System.Windows.Forms.Padding(3)
    Me.tpText.Size = New System.Drawing.Size(773, 430)
    Me.tpText.TabIndex = 1
    Me.tpText.Text = "Text"
    Me.tpText.UseVisualStyleBackColor = True
    '
    'Label8
    '
    Me.Label8.AutoSize = True
    Me.Label8.Location = New System.Drawing.Point(6, 8)
    Me.Label8.Name = "Label8"
    Me.Label8.Size = New System.Drawing.Size(104, 13)
    Me.Label8.TabIndex = 1
    Me.Label8.Text = "Paste your text here:"
    '
    'tbGenText
    '
    Me.tbGenText.AcceptsTab = True
    Me.tbGenText.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbGenText.EnableAutoDragDrop = True
    Me.tbGenText.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbGenText.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
    Me.tbGenText.Location = New System.Drawing.Point(6, 24)
    Me.tbGenText.Name = "tbGenText"
    Me.tbGenText.Size = New System.Drawing.Size(761, 400)
    Me.tbGenText.TabIndex = 14
    Me.tbGenText.Text = ""
    '
    'Label10
    '
    Me.Label10.AutoSize = True
    Me.Label10.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
    Me.Label10.Location = New System.Drawing.Point(332, 9)
    Me.Label10.Name = "Label10"
    Me.Label10.Size = New System.Drawing.Size(243, 13)
    Me.Label10.TabIndex = 34
    Me.Label10.Text = "Use no spaces. The extension will be set to: .psdx"
    '
    'Timer1
    '
    '
    'frmNewText
    '
    Me.AcceptButton = Me.cmdOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(803, 568)
    Me.Controls.Add(Me.Label10)
    Me.Controls.Add(Me.TabControl1)
    Me.Controls.Add(Me.cmdGenDir)
    Me.Controls.Add(Me.tbGenFileName)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.tbGenDirName)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.cmdOk)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmNewText"
    Me.Text = "Create a new text"
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.TabControl1.ResumeLayout(False)
    Me.tpInfo.ResumeLayout(False)
    Me.GroupBox1.ResumeLayout(False)
    Me.GroupBox1.PerformLayout()
    Me.GroupBox2.ResumeLayout(False)
    Me.GroupBox2.PerformLayout()
    Me.tpText.ResumeLayout(False)
    Me.tpText.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents tbGenDirName As System.Windows.Forms.TextBox
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents tbGenFileName As System.Windows.Forms.TextBox
  Friend WithEvents cmdGenDir As System.Windows.Forms.Button
  Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
  Friend WithEvents tpInfo As System.Windows.Forms.TabPage
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents tbGenGenre As System.Windows.Forms.TextBox
  Friend WithEvents Label9 As System.Windows.Forms.Label
  Friend WithEvents Label6 As System.Windows.Forms.Label
  Friend WithEvents tbGenSource As System.Windows.Forms.RichTextBox
  Friend WithEvents Label5 As System.Windows.Forms.Label
  Friend WithEvents tbGenTitle As System.Windows.Forms.TextBox
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents tbGenDistributor As System.Windows.Forms.TextBox
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents tbGenSubType As System.Windows.Forms.TextBox
  Friend WithEvents Label71 As System.Windows.Forms.Label
  Friend WithEvents tbGenCreaManu As System.Windows.Forms.TextBox
  Friend WithEvents Label69 As System.Windows.Forms.Label
  Friend WithEvents tbGenCreaOrig As System.Windows.Forms.TextBox
  Friend WithEvents Label68 As System.Windows.Forms.Label
  Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
  Friend WithEvents Label7 As System.Windows.Forms.Label
  Friend WithEvents tbGenLngName As System.Windows.Forms.TextBox
  Friend WithEvents Label70 As System.Windows.Forms.Label
  Friend WithEvents tbGenLngId As System.Windows.Forms.TextBox
  Friend WithEvents Label67 As System.Windows.Forms.Label
  Friend WithEvents tbGenEditor As System.Windows.Forms.RichTextBox
  Friend WithEvents Label66 As System.Windows.Forms.Label
  Friend WithEvents tbGenAuthor As System.Windows.Forms.TextBox
  Friend WithEvents Label65 As System.Windows.Forms.Label
  Friend WithEvents tpText As System.Windows.Forms.TabPage
  Friend WithEvents Label8 As System.Windows.Forms.Label
  Friend WithEvents tbGenText As System.Windows.Forms.RichTextBox
  Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
  Friend WithEvents Label10 As System.Windows.Forms.Label
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
End Class
