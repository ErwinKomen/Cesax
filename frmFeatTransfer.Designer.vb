<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFeatTransfer
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFeatTransfer))
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
    Me.cmdOk = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.tbExportDir = New System.Windows.Forms.TextBox()
    Me.cmdExportDir = New System.Windows.Forms.Button()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.GroupBox2 = New System.Windows.Forms.GroupBox()
    Me.tbFeatBlackList = New System.Windows.Forms.TextBox()
    Me.Label6 = New System.Windows.Forms.Label()
    Me.tbFeatWhiteList = New System.Windows.Forms.TextBox()
    Me.rbRestrOnly = New System.Windows.Forms.RadioButton()
    Me.rbRestrNone = New System.Windows.Forms.RadioButton()
    Me.Label5 = New System.Windows.Forms.Label()
    Me.tbFeatFilter = New System.Windows.Forms.RichTextBox()
    Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
    Me.GroupBox3 = New System.Windows.Forms.GroupBox()
    Me.lboFeatDbase = New System.Windows.Forms.ListBox()
    Me.Label7 = New System.Windows.Forms.Label()
    Me.tbFeatDbase = New System.Windows.Forms.TextBox()
    Me.tbImportFile = New System.Windows.Forms.TextBox()
    Me.cmdImportFile = New System.Windows.Forms.Button()
    Me.GroupBox4 = New System.Windows.Forms.GroupBox()
    Me.chbChildPde = New System.Windows.Forms.CheckBox()
    Me.chbAttrNotes = New System.Windows.Forms.CheckBox()
    Me.chbAttrStatus = New System.Windows.Forms.CheckBox()
    Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
    Me.TabControl1 = New System.Windows.Forms.TabControl()
    Me.tpImport = New System.Windows.Forms.TabPage()
    Me.GroupBox1 = New System.Windows.Forms.GroupBox()
    Me.rbResId = New System.Windows.Forms.RadioButton()
    Me.rbTextForEtree = New System.Windows.Forms.RadioButton()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.tpExport = New System.Windows.Forms.TabPage()
    Me.Label3 = New System.Windows.Forms.Label()
    Me.StatusStrip1.SuspendLayout()
    Me.GroupBox2.SuspendLayout()
    Me.GroupBox3.SuspendLayout()
    Me.GroupBox4.SuspendLayout()
    Me.TabControl1.SuspendLayout()
    Me.tpImport.SuspendLayout()
    Me.GroupBox1.SuspendLayout()
    Me.tpExport.SuspendLayout()
    Me.SuspendLayout()
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 547)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(688, 22)
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
    Me.cmdOk.Location = New System.Drawing.Point(604, 521)
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
    Me.cmdCancel.Location = New System.Drawing.Point(520, 521)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 2
    Me.cmdCancel.Text = "&Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'Timer1
    '
    '
    'tbExportDir
    '
    Me.tbExportDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbExportDir.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbExportDir.Location = New System.Drawing.Point(107, 6)
    Me.tbExportDir.Name = "tbExportDir"
    Me.tbExportDir.ReadOnly = True
    Me.tbExportDir.Size = New System.Drawing.Size(469, 20)
    Me.tbExportDir.TabIndex = 4
    '
    'cmdExportDir
    '
    Me.cmdExportDir.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdExportDir.Location = New System.Drawing.Point(593, 5)
    Me.cmdExportDir.Name = "cmdExportDir"
    Me.cmdExportDir.Size = New System.Drawing.Size(63, 21)
    Me.cmdExportDir.TabIndex = 5
    Me.cmdExportDir.Text = "..."
    Me.cmdExportDir.UseVisualStyleBackColor = True
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(12, 62)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(0, 13)
    Me.Label2.TabIndex = 6
    '
    'GroupBox2
    '
    Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GroupBox2.Controls.Add(Me.tbFeatBlackList)
    Me.GroupBox2.Controls.Add(Me.Label6)
    Me.GroupBox2.Controls.Add(Me.tbFeatWhiteList)
    Me.GroupBox2.Controls.Add(Me.rbRestrOnly)
    Me.GroupBox2.Controls.Add(Me.rbRestrNone)
    Me.GroupBox2.Location = New System.Drawing.Point(221, 182)
    Me.GroupBox2.Name = "GroupBox2"
    Me.GroupBox2.Size = New System.Drawing.Size(458, 134)
    Me.GroupBox2.TabIndex = 10
    Me.GroupBox2.TabStop = False
    Me.GroupBox2.Text = "Restrictions"
    '
    'tbFeatBlackList
    '
    Me.tbFeatBlackList.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFeatBlackList.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbFeatBlackList.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbFeatBlackList.Location = New System.Drawing.Point(24, 103)
    Me.tbFeatBlackList.Name = "tbFeatBlackList"
    Me.tbFeatBlackList.ReadOnly = True
    Me.tbFeatBlackList.Size = New System.Drawing.Size(422, 20)
    Me.tbFeatBlackList.TabIndex = 4
    '
    'Label6
    '
    Me.Label6.AutoSize = True
    Me.Label6.Location = New System.Drawing.Point(21, 87)
    Me.Label6.Name = "Label6"
    Me.Label6.Size = New System.Drawing.Size(207, 13)
    Me.Label6.TabIndex = 3
    Me.Label6.Text = "Exclude features with the following values:"
    '
    'tbFeatWhiteList
    '
    Me.tbFeatWhiteList.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFeatWhiteList.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbFeatWhiteList.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbFeatWhiteList.Location = New System.Drawing.Point(24, 64)
    Me.tbFeatWhiteList.Name = "tbFeatWhiteList"
    Me.tbFeatWhiteList.ReadOnly = True
    Me.tbFeatWhiteList.Size = New System.Drawing.Size(422, 20)
    Me.tbFeatWhiteList.TabIndex = 2
    '
    'rbRestrOnly
    '
    Me.rbRestrOnly.AutoSize = True
    Me.rbRestrOnly.Location = New System.Drawing.Point(6, 41)
    Me.rbRestrOnly.Name = "rbRestrOnly"
    Me.rbRestrOnly.Size = New System.Drawing.Size(234, 17)
    Me.rbRestrOnly.TabIndex = 1
    Me.rbRestrOnly.Text = "Only copy features with the following values:"
    Me.rbRestrOnly.UseVisualStyleBackColor = True
    '
    'rbRestrNone
    '
    Me.rbRestrNone.AutoSize = True
    Me.rbRestrNone.Checked = True
    Me.rbRestrNone.Location = New System.Drawing.Point(6, 19)
    Me.rbRestrNone.Name = "rbRestrNone"
    Me.rbRestrNone.Size = New System.Drawing.Size(92, 17)
    Me.rbRestrNone.TabIndex = 0
    Me.rbRestrNone.TabStop = True
    Me.rbRestrNone.Text = "No restrictions"
    Me.rbRestrNone.UseVisualStyleBackColor = True
    '
    'Label5
    '
    Me.Label5.AutoSize = True
    Me.Label5.Location = New System.Drawing.Point(12, 375)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(401, 13)
    Me.Label5.TabIndex = 11
    Me.Label5.Text = "Currently selected feature (changes should be made through Corpus/FilterFeatures)" & _
        ":"
    '
    'tbFeatFilter
    '
    Me.tbFeatFilter.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFeatFilter.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbFeatFilter.ForeColor = System.Drawing.Color.Maroon
    Me.tbFeatFilter.Location = New System.Drawing.Point(15, 391)
    Me.tbFeatFilter.Name = "tbFeatFilter"
    Me.tbFeatFilter.ReadOnly = True
    Me.tbFeatFilter.Size = New System.Drawing.Size(664, 124)
    Me.tbFeatFilter.TabIndex = 12
    Me.tbFeatFilter.Text = ""
    '
    'GroupBox3
    '
    Me.GroupBox3.Controls.Add(Me.lboFeatDbase)
    Me.GroupBox3.Location = New System.Drawing.Point(15, 182)
    Me.GroupBox3.Name = "GroupBox3"
    Me.GroupBox3.Size = New System.Drawing.Size(200, 186)
    Me.GroupBox3.TabIndex = 13
    Me.GroupBox3.TabStop = False
    Me.GroupBox3.Text = "Select the feature in the database"
    '
    'lboFeatDbase
    '
    Me.lboFeatDbase.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.lboFeatDbase.FormattingEnabled = True
    Me.lboFeatDbase.Location = New System.Drawing.Point(9, 19)
    Me.lboFeatDbase.Name = "lboFeatDbase"
    Me.lboFeatDbase.Size = New System.Drawing.Size(185, 160)
    Me.lboFeatDbase.TabIndex = 0
    '
    'Label7
    '
    Me.Label7.AutoSize = True
    Me.Label7.Location = New System.Drawing.Point(228, 338)
    Me.Label7.Name = "Label7"
    Me.Label7.Size = New System.Drawing.Size(52, 13)
    Me.Label7.TabIndex = 14
    Me.Label7.Text = "Selected:"
    '
    'tbFeatDbase
    '
    Me.tbFeatDbase.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbFeatDbase.Location = New System.Drawing.Point(283, 335)
    Me.tbFeatDbase.Name = "tbFeatDbase"
    Me.tbFeatDbase.ReadOnly = True
    Me.tbFeatDbase.Size = New System.Drawing.Size(100, 20)
    Me.tbFeatDbase.TabIndex = 15
    '
    'tbImportFile
    '
    Me.tbImportFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbImportFile.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbImportFile.Location = New System.Drawing.Point(87, 6)
    Me.tbImportFile.Name = "tbImportFile"
    Me.tbImportFile.ReadOnly = True
    Me.tbImportFile.Size = New System.Drawing.Size(489, 20)
    Me.tbImportFile.TabIndex = 8
    '
    'cmdImportFile
    '
    Me.cmdImportFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdImportFile.Location = New System.Drawing.Point(590, 5)
    Me.cmdImportFile.Name = "cmdImportFile"
    Me.cmdImportFile.Size = New System.Drawing.Size(63, 21)
    Me.cmdImportFile.TabIndex = 9
    Me.cmdImportFile.Text = "..."
    Me.cmdImportFile.UseVisualStyleBackColor = True
    '
    'GroupBox4
    '
    Me.GroupBox4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GroupBox4.Controls.Add(Me.chbChildPde)
    Me.GroupBox4.Controls.Add(Me.chbAttrNotes)
    Me.GroupBox4.Controls.Add(Me.chbAttrStatus)
    Me.GroupBox4.Location = New System.Drawing.Point(410, 319)
    Me.GroupBox4.Name = "GroupBox4"
    Me.GroupBox4.Size = New System.Drawing.Size(269, 53)
    Me.GroupBox4.TabIndex = 17
    Me.GroupBox4.TabStop = False
    Me.GroupBox4.Text = "Also copy attributes ..."
    '
    'chbChildPde
    '
    Me.chbChildPde.AutoSize = True
    Me.chbChildPde.Location = New System.Drawing.Point(179, 19)
    Me.chbChildPde.Name = "chbChildPde"
    Me.chbChildPde.Size = New System.Drawing.Size(78, 17)
    Me.chbChildPde.TabIndex = 2
    Me.chbChildPde.Text = "&Translation"
    Me.chbChildPde.UseVisualStyleBackColor = True
    '
    'chbAttrNotes
    '
    Me.chbAttrNotes.AutoSize = True
    Me.chbAttrNotes.Location = New System.Drawing.Point(98, 19)
    Me.chbAttrNotes.Name = "chbAttrNotes"
    Me.chbAttrNotes.Size = New System.Drawing.Size(54, 17)
    Me.chbAttrNotes.TabIndex = 1
    Me.chbAttrNotes.Text = "&Notes"
    Me.chbAttrNotes.UseVisualStyleBackColor = True
    '
    'chbAttrStatus
    '
    Me.chbAttrStatus.AutoSize = True
    Me.chbAttrStatus.Location = New System.Drawing.Point(6, 19)
    Me.chbAttrStatus.Name = "chbAttrStatus"
    Me.chbAttrStatus.Size = New System.Drawing.Size(56, 17)
    Me.chbAttrStatus.TabIndex = 0
    Me.chbAttrStatus.Text = "&Status"
    Me.chbAttrStatus.UseVisualStyleBackColor = True
    '
    'OpenFileDialog1
    '
    Me.OpenFileDialog1.FileName = "OpenFileDialog1"
    '
    'TabControl1
    '
    Me.TabControl1.Controls.Add(Me.tpImport)
    Me.TabControl1.Controls.Add(Me.tpExport)
    Me.TabControl1.Location = New System.Drawing.Point(15, 12)
    Me.TabControl1.Name = "TabControl1"
    Me.TabControl1.SelectedIndex = 0
    Me.TabControl1.Size = New System.Drawing.Size(664, 164)
    Me.TabControl1.TabIndex = 18
    '
    'tpImport
    '
    Me.tpImport.Controls.Add(Me.GroupBox1)
    Me.tpImport.Controls.Add(Me.Label1)
    Me.tpImport.Controls.Add(Me.tbImportFile)
    Me.tpImport.Controls.Add(Me.cmdImportFile)
    Me.tpImport.Location = New System.Drawing.Point(4, 22)
    Me.tpImport.Name = "tpImport"
    Me.tpImport.Padding = New System.Windows.Forms.Padding(3)
    Me.tpImport.Size = New System.Drawing.Size(656, 138)
    Me.tpImport.TabIndex = 0
    Me.tpImport.Text = "Import"
    Me.tpImport.UseVisualStyleBackColor = True
    '
    'GroupBox1
    '
    Me.GroupBox1.Controls.Add(Me.rbResId)
    Me.GroupBox1.Controls.Add(Me.rbTextForEtree)
    Me.GroupBox1.Location = New System.Drawing.Point(3, 38)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(650, 100)
    Me.GroupBox1.TabIndex = 13
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Import matching method"
    '
    'rbResId
    '
    Me.rbResId.AutoSize = True
    Me.rbResId.Location = New System.Drawing.Point(6, 31)
    Me.rbResId.Name = "rbResId"
    Me.rbResId.Size = New System.Drawing.Size(448, 17)
    Me.rbResId.TabIndex = 11
    Me.rbResId.TabStop = True
    Me.rbResId.Text = "ResId (the databases have the same number of records, covering the same text loca" & _
        "tions)"
    Me.rbResId.UseVisualStyleBackColor = True
    '
    'rbTextForEtree
    '
    Me.rbTextForEtree.AutoSize = True
    Me.rbTextForEtree.Location = New System.Drawing.Point(6, 63)
    Me.rbTextForEtree.Name = "rbTextForEtree"
    Me.rbTextForEtree.Size = New System.Drawing.Size(448, 17)
    Me.rbTextForEtree.TabIndex = 12
    Me.rbTextForEtree.TabStop = True
    Me.rbTextForEtree.Text = "Text/Forest/Etree (the database entries are uniquely discernable per text/line/co" & _
        "nstituent)"
    Me.rbTextForEtree.UseVisualStyleBackColor = True
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(3, 9)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(78, 13)
    Me.Label1.TabIndex = 10
    Me.Label1.Text = "Import from file:"
    '
    'tpExport
    '
    Me.tpExport.Controls.Add(Me.Label3)
    Me.tpExport.Controls.Add(Me.tbExportDir)
    Me.tpExport.Controls.Add(Me.cmdExportDir)
    Me.tpExport.Location = New System.Drawing.Point(4, 22)
    Me.tpExport.Name = "tpExport"
    Me.tpExport.Padding = New System.Windows.Forms.Padding(3)
    Me.tpExport.Size = New System.Drawing.Size(656, 138)
    Me.tpExport.TabIndex = 1
    Me.tpExport.Text = "Export"
    Me.tpExport.UseVisualStyleBackColor = True
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(6, 9)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(95, 13)
    Me.Label3.TabIndex = 6
    Me.Label3.Text = "Export to directory:"
    '
    'frmFeatTransfer
    '
    Me.AcceptButton = Me.cmdOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(688, 569)
    Me.Controls.Add(Me.TabControl1)
    Me.Controls.Add(Me.GroupBox4)
    Me.Controls.Add(Me.tbFeatDbase)
    Me.Controls.Add(Me.Label7)
    Me.Controls.Add(Me.GroupBox3)
    Me.Controls.Add(Me.tbFeatFilter)
    Me.Controls.Add(Me.Label5)
    Me.Controls.Add(Me.GroupBox2)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.cmdOk)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmFeatTransfer"
    Me.Text = "Import and export feature values"
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.GroupBox2.ResumeLayout(False)
    Me.GroupBox2.PerformLayout()
    Me.GroupBox3.ResumeLayout(False)
    Me.GroupBox4.ResumeLayout(False)
    Me.GroupBox4.PerformLayout()
    Me.TabControl1.ResumeLayout(False)
    Me.tpImport.ResumeLayout(False)
    Me.tpImport.PerformLayout()
    Me.GroupBox1.ResumeLayout(False)
    Me.GroupBox1.PerformLayout()
    Me.tpExport.ResumeLayout(False)
    Me.tpExport.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents tbExportDir As System.Windows.Forms.TextBox
  Friend WithEvents cmdExportDir As System.Windows.Forms.Button
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
  Friend WithEvents tbFeatWhiteList As System.Windows.Forms.TextBox
  Friend WithEvents rbRestrOnly As System.Windows.Forms.RadioButton
  Friend WithEvents rbRestrNone As System.Windows.Forms.RadioButton
  Friend WithEvents Label5 As System.Windows.Forms.Label
  Friend WithEvents tbFeatFilter As System.Windows.Forms.RichTextBox
  Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
  Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
  Friend WithEvents lboFeatDbase As System.Windows.Forms.ListBox
  Friend WithEvents tbFeatBlackList As System.Windows.Forms.TextBox
  Friend WithEvents Label6 As System.Windows.Forms.Label
  Friend WithEvents Label7 As System.Windows.Forms.Label
  Friend WithEvents tbFeatDbase As System.Windows.Forms.TextBox
  Friend WithEvents tbImportFile As System.Windows.Forms.TextBox
  Friend WithEvents cmdImportFile As System.Windows.Forms.Button
  Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
  Friend WithEvents chbAttrStatus As System.Windows.Forms.CheckBox
  Friend WithEvents chbAttrNotes As System.Windows.Forms.CheckBox
  Friend WithEvents chbChildPde As System.Windows.Forms.CheckBox
  Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
  Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
  Friend WithEvents tpImport As System.Windows.Forms.TabPage
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents rbResId As System.Windows.Forms.RadioButton
  Friend WithEvents rbTextForEtree As System.Windows.Forms.RadioButton
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents tpExport As System.Windows.Forms.TabPage
  Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
