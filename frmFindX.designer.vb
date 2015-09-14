<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFindX
  Inherits System.Windows.Forms.Form

  'Form overrides dispose to clean up the component list.
  <System.Diagnostics.DebuggerNonUserCode()> _
  Protected Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing AndAlso components IsNot Nothing Then
      components.Dispose()
    End If
    MyBase.Dispose(disposing)
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  <System.Diagnostics.DebuggerStepThrough()> _
  Private Sub InitializeComponent()
    Me.components = New System.ComponentModel.Container
    Me.lbSearch = New System.Windows.Forms.Label
    Me.cmdPrev = New System.Windows.Forms.Button
    Me.cmdNext = New System.Windows.Forms.Button
    Me.chbWhole = New System.Windows.Forms.CheckBox
    Me.tbFind = New System.Windows.Forms.RichTextBox
    Me.Label1 = New System.Windows.Forms.Label
    Me.tbQrySearch = New System.Windows.Forms.TextBox
    Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
    Me.dgvQryRel = New System.Windows.Forms.DataGridView
    Me.Label4 = New System.Windows.Forms.Label
    Me.cboRelArg = New System.Windows.Forms.ComboBox
    Me.Label3 = New System.Windows.Forms.Label
    Me.cboRelName = New System.Windows.Forms.ComboBox
    Me.tbRelCond = New System.Windows.Forms.TextBox
    Me.Label2 = New System.Windows.Forms.Label
    Me.TabControl1 = New System.Windows.Forms.TabControl
    Me.tpXpath = New System.Windows.Forms.TabPage
    Me.tpQuery = New System.Windows.Forms.TabPage
    Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
    Me.dgvQry = New System.Windows.Forms.DataGridView
    Me.tbQryDescr = New System.Windows.Forms.RichTextBox
    Me.Label6 = New System.Windows.Forms.Label
    Me.tbQryName = New System.Windows.Forms.TextBox
    Me.Label5 = New System.Windows.Forms.Label
    Me.tpResults = New System.Windows.Forms.TabPage
    Me.dgvResult = New System.Windows.Forms.DataGridView
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
    Me.mnuQuery = New System.Windows.Forms.ToolStripMenuItem
    Me.mnuQryNew = New System.Windows.Forms.ToolStripMenuItem
    Me.mnuQryDel = New System.Windows.Forms.ToolStripMenuItem
    Me.mnuRelation = New System.Windows.Forms.ToolStripMenuItem
    Me.mnuRelNew = New System.Windows.Forms.ToolStripMenuItem
    Me.mnuRelDel = New System.Windows.Forms.ToolStripMenuItem
    Me.cmdSearch = New System.Windows.Forms.Button
    Me.mnuQryConvert = New System.Windows.Forms.ToolStripMenuItem
    Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
    Me.SplitContainer1.Panel1.SuspendLayout()
    Me.SplitContainer1.Panel2.SuspendLayout()
    Me.SplitContainer1.SuspendLayout()
    CType(Me.dgvQryRel, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.TabControl1.SuspendLayout()
    Me.tpXpath.SuspendLayout()
    Me.tpQuery.SuspendLayout()
    Me.SplitContainer2.Panel1.SuspendLayout()
    Me.SplitContainer2.Panel2.SuspendLayout()
    Me.SplitContainer2.SuspendLayout()
    CType(Me.dgvQry, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.tpResults.SuspendLayout()
    CType(Me.dgvResult, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.MenuStrip1.SuspendLayout()
    Me.SuspendLayout()
    '
    'lbSearch
    '
    Me.lbSearch.AutoSize = True
    Me.lbSearch.Location = New System.Drawing.Point(3, 5)
    Me.lbSearch.Name = "lbSearch"
    Me.lbSearch.Size = New System.Drawing.Size(149, 13)
    Me.lbSearch.TabIndex = 0
    Me.lbSearch.Text = "Search for (Xpath expression):"
    '
    'cmdPrev
    '
    Me.cmdPrev.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdPrev.Location = New System.Drawing.Point(542, 320)
    Me.cmdPrev.Name = "cmdPrev"
    Me.cmdPrev.Size = New System.Drawing.Size(79, 23)
    Me.cmdPrev.TabIndex = 2
    Me.cmdPrev.Text = "&Previous"
    Me.cmdPrev.UseVisualStyleBackColor = True
    '
    'cmdNext
    '
    Me.cmdNext.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdNext.Location = New System.Drawing.Point(627, 320)
    Me.cmdNext.Name = "cmdNext"
    Me.cmdNext.Size = New System.Drawing.Size(79, 23)
    Me.cmdNext.TabIndex = 3
    Me.cmdNext.Text = "&Next"
    Me.cmdNext.UseVisualStyleBackColor = True
    '
    'chbWhole
    '
    Me.chbWhole.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.chbWhole.AutoSize = True
    Me.chbWhole.Location = New System.Drawing.Point(3, 324)
    Me.chbWhole.Name = "chbWhole"
    Me.chbWhole.Size = New System.Drawing.Size(57, 17)
    Me.chbWhole.TabIndex = 4
    Me.chbWhole.Text = "&Option"
    Me.chbWhole.UseVisualStyleBackColor = True
    '
    'tbFind
    '
    Me.tbFind.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFind.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbFind.ForeColor = System.Drawing.Color.Navy
    Me.tbFind.Location = New System.Drawing.Point(6, 21)
    Me.tbFind.Name = "tbFind"
    Me.tbFind.Size = New System.Drawing.Size(697, 250)
    Me.tbFind.TabIndex = 5
    Me.tbFind.Text = ""
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(3, 36)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(81, 13)
    Me.Label1.TabIndex = 6
    Me.Label1.Text = "Search domain:"
    '
    'tbQrySearch
    '
    Me.tbQrySearch.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbQrySearch.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbQrySearch.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
    Me.tbQrySearch.Location = New System.Drawing.Point(96, 33)
    Me.tbQrySearch.Name = "tbQrySearch"
    Me.tbQrySearch.Size = New System.Drawing.Size(469, 21)
    Me.tbQrySearch.TabIndex = 7
    '
    'SplitContainer1
    '
    Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.SplitContainer1.Location = New System.Drawing.Point(3, 148)
    Me.SplitContainer1.Name = "SplitContainer1"
    '
    'SplitContainer1.Panel1
    '
    Me.SplitContainer1.Panel1.Controls.Add(Me.dgvQryRel)
    '
    'SplitContainer1.Panel2
    '
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
    Me.SplitContainer1.Panel2.Controls.Add(Me.cboRelArg)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label3)
    Me.SplitContainer1.Panel2.Controls.Add(Me.cboRelName)
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbRelCond)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label2)
    Me.SplitContainer1.Size = New System.Drawing.Size(565, 105)
    Me.SplitContainer1.SplitterDistance = 221
    Me.SplitContainer1.TabIndex = 8
    '
    'dgvQryRel
    '
    Me.dgvQryRel.AllowUserToAddRows = False
    Me.dgvQryRel.AllowUserToDeleteRows = False
    Me.dgvQryRel.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
    Me.dgvQryRel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvQryRel.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgvQryRel.Location = New System.Drawing.Point(0, 0)
    Me.dgvQryRel.Name = "dgvQryRel"
    Me.dgvQryRel.ReadOnly = True
    Me.dgvQryRel.Size = New System.Drawing.Size(221, 105)
    Me.dgvQryRel.TabIndex = 0
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(3, 33)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(71, 13)
    Me.Label4.TabIndex = 5
    Me.Label4.Text = "Operating on:"
    '
    'cboRelArg
    '
    Me.cboRelArg.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboRelArg.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cboRelArg.FormattingEnabled = True
    Me.cboRelArg.Location = New System.Drawing.Point(111, 30)
    Me.cboRelArg.Name = "cboRelArg"
    Me.cboRelArg.Size = New System.Drawing.Size(121, 23)
    Me.cboRelArg.TabIndex = 4
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(3, 6)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(102, 13)
    Me.Label3.TabIndex = 3
    Me.Label3.Text = "Relation or function:"
    '
    'cboRelName
    '
    Me.cboRelName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cboRelName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboRelName.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cboRelName.FormattingEnabled = True
    Me.cboRelName.Location = New System.Drawing.Point(111, 3)
    Me.cboRelName.Name = "cboRelName"
    Me.cboRelName.Size = New System.Drawing.Size(226, 23)
    Me.cboRelName.TabIndex = 2
    '
    'tbRelCond
    '
    Me.tbRelCond.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbRelCond.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbRelCond.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
    Me.tbRelCond.Location = New System.Drawing.Point(111, 57)
    Me.tbRelCond.Name = "tbRelCond"
    Me.tbRelCond.Size = New System.Drawing.Size(230, 21)
    Me.tbRelCond.TabIndex = 1
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(3, 60)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(96, 13)
    Me.Label2.TabIndex = 0
    Me.Label2.Text = "Outcome matches:"
    '
    'TabControl1
    '
    Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.TabControl1.Controls.Add(Me.tpXpath)
    Me.TabControl1.Controls.Add(Me.tpQuery)
    Me.TabControl1.Controls.Add(Me.tpResults)
    Me.TabControl1.Location = New System.Drawing.Point(3, 27)
    Me.TabControl1.Name = "TabControl1"
    Me.TabControl1.SelectedIndex = 0
    Me.TabControl1.Size = New System.Drawing.Size(703, 288)
    Me.TabControl1.TabIndex = 9
    '
    'tpXpath
    '
    Me.tpXpath.Controls.Add(Me.tbFind)
    Me.tpXpath.Controls.Add(Me.lbSearch)
    Me.tpXpath.Location = New System.Drawing.Point(4, 22)
    Me.tpXpath.Name = "tpXpath"
    Me.tpXpath.Padding = New System.Windows.Forms.Padding(3)
    Me.tpXpath.Size = New System.Drawing.Size(695, 262)
    Me.tpXpath.TabIndex = 0
    Me.tpXpath.Text = "Xpath"
    Me.tpXpath.UseVisualStyleBackColor = True
    '
    'tpQuery
    '
    Me.tpQuery.Controls.Add(Me.SplitContainer2)
    Me.tpQuery.Location = New System.Drawing.Point(4, 22)
    Me.tpQuery.Name = "tpQuery"
    Me.tpQuery.Padding = New System.Windows.Forms.Padding(3)
    Me.tpQuery.Size = New System.Drawing.Size(695, 262)
    Me.tpQuery.TabIndex = 1
    Me.tpQuery.Text = "QueryBuilder"
    Me.tpQuery.UseVisualStyleBackColor = True
    '
    'SplitContainer2
    '
    Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer2.Location = New System.Drawing.Point(3, 3)
    Me.SplitContainer2.Name = "SplitContainer2"
    '
    'SplitContainer2.Panel1
    '
    Me.SplitContainer2.Panel1.Controls.Add(Me.dgvQry)
    '
    'SplitContainer2.Panel2
    '
    Me.SplitContainer2.Panel2.Controls.Add(Me.tbQryDescr)
    Me.SplitContainer2.Panel2.Controls.Add(Me.Label6)
    Me.SplitContainer2.Panel2.Controls.Add(Me.tbQryName)
    Me.SplitContainer2.Panel2.Controls.Add(Me.Label5)
    Me.SplitContainer2.Panel2.Controls.Add(Me.tbQrySearch)
    Me.SplitContainer2.Panel2.Controls.Add(Me.Label1)
    Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer1)
    Me.SplitContainer2.Size = New System.Drawing.Size(689, 256)
    Me.SplitContainer2.SplitterDistance = 120
    Me.SplitContainer2.TabIndex = 9
    '
    'dgvQry
    '
    Me.dgvQry.AllowUserToAddRows = False
    Me.dgvQry.AllowUserToDeleteRows = False
    Me.dgvQry.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
    Me.dgvQry.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvQry.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgvQry.Location = New System.Drawing.Point(0, 0)
    Me.dgvQry.Name = "dgvQry"
    Me.dgvQry.ReadOnly = True
    Me.dgvQry.Size = New System.Drawing.Size(120, 256)
    Me.dgvQry.TabIndex = 0
    '
    'tbQryDescr
    '
    Me.tbQryDescr.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbQryDescr.Location = New System.Drawing.Point(2, 71)
    Me.tbQryDescr.Name = "tbQryDescr"
    Me.tbQryDescr.Size = New System.Drawing.Size(563, 71)
    Me.tbQryDescr.TabIndex = 12
    Me.tbQryDescr.Text = ""
    '
    'Label6
    '
    Me.Label6.AutoSize = True
    Me.Label6.Location = New System.Drawing.Point(3, 55)
    Me.Label6.Name = "Label6"
    Me.Label6.Size = New System.Drawing.Size(59, 13)
    Me.Label6.TabIndex = 11
    Me.Label6.Text = "Comments:"
    '
    'tbQryName
    '
    Me.tbQryName.Location = New System.Drawing.Point(96, 7)
    Me.tbQryName.Name = "tbQryName"
    Me.tbQryName.Size = New System.Drawing.Size(189, 20)
    Me.tbQryName.TabIndex = 10
    '
    'Label5
    '
    Me.Label5.AutoSize = True
    Me.Label5.Location = New System.Drawing.Point(3, 10)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(67, 13)
    Me.Label5.TabIndex = 9
    Me.Label5.Text = "Query name:"
    '
    'tpResults
    '
    Me.tpResults.Controls.Add(Me.dgvResult)
    Me.tpResults.Location = New System.Drawing.Point(4, 22)
    Me.tpResults.Name = "tpResults"
    Me.tpResults.Size = New System.Drawing.Size(695, 262)
    Me.tpResults.TabIndex = 2
    Me.tpResults.Text = "Results"
    Me.tpResults.UseVisualStyleBackColor = True
    '
    'dgvResult
    '
    Me.dgvResult.AllowUserToAddRows = False
    Me.dgvResult.AllowUserToDeleteRows = False
    Me.dgvResult.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
    Me.dgvResult.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvResult.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgvResult.Location = New System.Drawing.Point(0, 0)
    Me.dgvResult.Name = "dgvResult"
    Me.dgvResult.ReadOnly = True
    Me.dgvResult.Size = New System.Drawing.Size(695, 262)
    Me.dgvResult.TabIndex = 0
    '
    'Timer1
    '
    '
    'MenuStrip1
    '
    Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuQuery, Me.mnuRelation})
    Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
    Me.MenuStrip1.Name = "MenuStrip1"
    Me.MenuStrip1.Size = New System.Drawing.Size(711, 24)
    Me.MenuStrip1.TabIndex = 10
    Me.MenuStrip1.Text = "MenuStrip1"
    '
    'mnuQuery
    '
    Me.mnuQuery.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuQryNew, Me.mnuQryDel, Me.ToolStripSeparator1, Me.mnuQryConvert})
    Me.mnuQuery.Name = "mnuQuery"
    Me.mnuQuery.Size = New System.Drawing.Size(49, 20)
    Me.mnuQuery.Text = "&Query"
    '
    'mnuQryNew
    '
    Me.mnuQryNew.Name = "mnuQryNew"
    Me.mnuQryNew.Size = New System.Drawing.Size(152, 22)
    Me.mnuQryNew.Text = "&New"
    '
    'mnuQryDel
    '
    Me.mnuQryDel.Name = "mnuQryDel"
    Me.mnuQryDel.Size = New System.Drawing.Size(152, 22)
    Me.mnuQryDel.Text = "&Delete"
    '
    'mnuRelation
    '
    Me.mnuRelation.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuRelNew, Me.mnuRelDel})
    Me.mnuRelation.Name = "mnuRelation"
    Me.mnuRelation.Size = New System.Drawing.Size(58, 20)
    Me.mnuRelation.Text = "&Relation"
    '
    'mnuRelNew
    '
    Me.mnuRelNew.Name = "mnuRelNew"
    Me.mnuRelNew.Size = New System.Drawing.Size(105, 22)
    Me.mnuRelNew.Text = "&New"
    '
    'mnuRelDel
    '
    Me.mnuRelDel.Name = "mnuRelDel"
    Me.mnuRelDel.Size = New System.Drawing.Size(105, 22)
    Me.mnuRelDel.Text = "&Delete"
    '
    'cmdSearch
    '
    Me.cmdSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdSearch.Location = New System.Drawing.Point(457, 320)
    Me.cmdSearch.Name = "cmdSearch"
    Me.cmdSearch.Size = New System.Drawing.Size(79, 23)
    Me.cmdSearch.TabIndex = 11
    Me.cmdSearch.Text = "&Search"
    Me.cmdSearch.UseVisualStyleBackColor = True
    '
    'mnuQryConvert
    '
    Me.mnuQryConvert.Name = "mnuQryConvert"
    Me.mnuQryConvert.Size = New System.Drawing.Size(152, 22)
    Me.mnuQryConvert.Text = "&Convert"
    '
    'ToolStripSeparator1
    '
    Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
    Me.ToolStripSeparator1.Size = New System.Drawing.Size(149, 6)
    '
    'frmFindX
    '
    Me.AcceptButton = Me.cmdNext
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(711, 347)
    Me.ControlBox = False
    Me.Controls.Add(Me.cmdSearch)
    Me.Controls.Add(Me.TabControl1)
    Me.Controls.Add(Me.chbWhole)
    Me.Controls.Add(Me.cmdNext)
    Me.Controls.Add(Me.cmdPrev)
    Me.Controls.Add(Me.MenuStrip1)
    Me.KeyPreview = True
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmFindX"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Xpath searching"
    Me.TopMost = True
    Me.SplitContainer1.Panel1.ResumeLayout(False)
    Me.SplitContainer1.Panel2.ResumeLayout(False)
    Me.SplitContainer1.Panel2.PerformLayout()
    Me.SplitContainer1.ResumeLayout(False)
    CType(Me.dgvQryRel, System.ComponentModel.ISupportInitialize).EndInit()
    Me.TabControl1.ResumeLayout(False)
    Me.tpXpath.ResumeLayout(False)
    Me.tpXpath.PerformLayout()
    Me.tpQuery.ResumeLayout(False)
    Me.SplitContainer2.Panel1.ResumeLayout(False)
    Me.SplitContainer2.Panel2.ResumeLayout(False)
    Me.SplitContainer2.Panel2.PerformLayout()
    Me.SplitContainer2.ResumeLayout(False)
    CType(Me.dgvQry, System.ComponentModel.ISupportInitialize).EndInit()
    Me.tpResults.ResumeLayout(False)
    CType(Me.dgvResult, System.ComponentModel.ISupportInitialize).EndInit()
    Me.MenuStrip1.ResumeLayout(False)
    Me.MenuStrip1.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents lbSearch As System.Windows.Forms.Label
  Friend WithEvents cmdPrev As System.Windows.Forms.Button
  Friend WithEvents cmdNext As System.Windows.Forms.Button
  Friend WithEvents chbWhole As System.Windows.Forms.CheckBox
  Friend WithEvents tbFind As System.Windows.Forms.RichTextBox
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents tbQrySearch As System.Windows.Forms.TextBox
  Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
  Friend WithEvents dgvQryRel As System.Windows.Forms.DataGridView
  Friend WithEvents tbRelCond As System.Windows.Forms.TextBox
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents cboRelArg As System.Windows.Forms.ComboBox
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents cboRelName As System.Windows.Forms.ComboBox
  Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
  Friend WithEvents tpXpath As System.Windows.Forms.TabPage
  Friend WithEvents tpQuery As System.Windows.Forms.TabPage
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
  Friend WithEvents tbQryName As System.Windows.Forms.TextBox
  Friend WithEvents Label5 As System.Windows.Forms.Label
  Friend WithEvents tbQryDescr As System.Windows.Forms.RichTextBox
  Friend WithEvents Label6 As System.Windows.Forms.Label
  Friend WithEvents dgvQry As System.Windows.Forms.DataGridView
  Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
  Friend WithEvents mnuQuery As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents mnuQryNew As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents mnuQryDel As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents mnuRelation As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents mnuRelNew As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents mnuRelDel As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents tpResults As System.Windows.Forms.TabPage
  Friend WithEvents dgvResult As System.Windows.Forms.DataGridView
  Friend WithEvents cmdSearch As System.Windows.Forms.Button
  Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
  Friend WithEvents mnuQryConvert As System.Windows.Forms.ToolStripMenuItem
End Class
