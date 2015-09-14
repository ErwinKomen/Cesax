<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPosTag
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
    Me.components = New System.ComponentModel.Container
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPosTag))
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar
    Me.Label1 = New System.Windows.Forms.Label
    Me.tbOrg = New System.Windows.Forms.RichTextBox
    Me.tbTrn = New System.Windows.Forms.RichTextBox
    Me.Label2 = New System.Windows.Forms.Label
    Me.GroupBox1 = New System.Windows.Forms.GroupBox
    Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
    Me.dgvPosTag = New System.Windows.Forms.DataGridView
    Me.Label3 = New System.Windows.Forms.Label
    Me.tbVern = New System.Windows.Forms.TextBox
    Me.tbPos = New System.Windows.Forms.TextBox
    Me.Label4 = New System.Windows.Forms.Label
    Me.tbMorph = New System.Windows.Forms.TextBox
    Me.Label5 = New System.Windows.Forms.Label
    Me.lboPos = New System.Windows.Forms.ListBox
    Me.Label6 = New System.Windows.Forms.Label
    Me.dgvPosVern = New System.Windows.Forms.DataGridView
    Me.Label7 = New System.Windows.Forms.Label
    Me.cmdOk = New System.Windows.Forms.Button
    Me.cmdSkip = New System.Windows.Forms.Button
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.StatusStrip1.SuspendLayout()
    Me.GroupBox1.SuspendLayout()
    Me.SplitContainer1.Panel1.SuspendLayout()
    Me.SplitContainer1.Panel2.SuspendLayout()
    Me.SplitContainer1.SuspendLayout()
    CType(Me.dgvPosTag, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.dgvPosVern, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 512)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(629, 22)
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
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(12, 9)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(28, 13)
    Me.Label1.TabIndex = 1
    Me.Label1.Text = "Text"
    '
    'tbOrg
    '
    Me.tbOrg.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbOrg.Location = New System.Drawing.Point(77, 6)
    Me.tbOrg.Name = "tbOrg"
    Me.tbOrg.Size = New System.Drawing.Size(540, 77)
    Me.tbOrg.TabIndex = 2
    Me.tbOrg.Text = ""
    '
    'tbTrn
    '
    Me.tbTrn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbTrn.Location = New System.Drawing.Point(77, 89)
    Me.tbTrn.Name = "tbTrn"
    Me.tbTrn.Size = New System.Drawing.Size(540, 77)
    Me.tbTrn.TabIndex = 4
    Me.tbTrn.Text = ""
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(12, 92)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(59, 13)
    Me.Label2.TabIndex = 3
    Me.Label2.Text = "Translation"
    '
    'GroupBox1
    '
    Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GroupBox1.Controls.Add(Me.SplitContainer1)
    Me.GroupBox1.Location = New System.Drawing.Point(0, 172)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(629, 337)
    Me.GroupBox1.TabIndex = 5
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Part-of-speech specification"
    '
    'SplitContainer1
    '
    Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
    Me.SplitContainer1.Location = New System.Drawing.Point(3, 16)
    Me.SplitContainer1.Name = "SplitContainer1"
    '
    'SplitContainer1.Panel1
    '
    Me.SplitContainer1.Panel1.Controls.Add(Me.dgvPosTag)
    '
    'SplitContainer1.Panel2
    '
    Me.SplitContainer1.Panel2.Controls.Add(Me.cmdCancel)
    Me.SplitContainer1.Panel2.Controls.Add(Me.cmdSkip)
    Me.SplitContainer1.Panel2.Controls.Add(Me.cmdOk)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label7)
    Me.SplitContainer1.Panel2.Controls.Add(Me.dgvPosVern)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
    Me.SplitContainer1.Panel2.Controls.Add(Me.lboPos)
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbMorph)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbPos)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbVern)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label3)
    Me.SplitContainer1.Size = New System.Drawing.Size(623, 318)
    Me.SplitContainer1.SplitterDistance = 207
    Me.SplitContainer1.TabIndex = 0
    '
    'dgvPosTag
    '
    Me.dgvPosTag.AllowUserToAddRows = False
    Me.dgvPosTag.AllowUserToDeleteRows = False
    Me.dgvPosTag.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
    Me.dgvPosTag.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvPosTag.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgvPosTag.Location = New System.Drawing.Point(0, 0)
    Me.dgvPosTag.Name = "dgvPosTag"
    Me.dgvPosTag.ReadOnly = True
    Me.dgvPosTag.Size = New System.Drawing.Size(207, 318)
    Me.dgvPosTag.TabIndex = 0
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(3, 6)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(36, 13)
    Me.Label3.TabIndex = 0
    Me.Label3.Text = "Word:"
    '
    'tbVern
    '
    Me.tbVern.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbVern.Location = New System.Drawing.Point(68, 3)
    Me.tbVern.Name = "tbVern"
    Me.tbVern.Size = New System.Drawing.Size(335, 20)
    Me.tbVern.TabIndex = 1
    '
    'tbPos
    '
    Me.tbPos.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbPos.Location = New System.Drawing.Point(105, 29)
    Me.tbPos.Name = "tbPos"
    Me.tbPos.Size = New System.Drawing.Size(114, 20)
    Me.tbPos.TabIndex = 3
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(3, 32)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(79, 13)
    Me.Label4.TabIndex = 2
    Me.Label4.Text = "Part-of-speech:"
    '
    'tbMorph
    '
    Me.tbMorph.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbMorph.Location = New System.Drawing.Point(105, 55)
    Me.tbMorph.Name = "tbMorph"
    Me.tbMorph.Size = New System.Drawing.Size(298, 20)
    Me.tbMorph.TabIndex = 5
    '
    'Label5
    '
    Me.Label5.AutoSize = True
    Me.Label5.Location = New System.Drawing.Point(3, 58)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(96, 13)
    Me.Label5.TabIndex = 4
    Me.Label5.Text = "Morphology/Gloss:"
    '
    'lboPos
    '
    Me.lboPos.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.lboPos.FormattingEnabled = True
    Me.lboPos.Location = New System.Drawing.Point(6, 127)
    Me.lboPos.Name = "lboPos"
    Me.lboPos.Size = New System.Drawing.Size(120, 186)
    Me.lboPos.TabIndex = 6
    '
    'Label6
    '
    Me.Label6.AutoSize = True
    Me.Label6.Location = New System.Drawing.Point(3, 111)
    Me.Label6.Name = "Label6"
    Me.Label6.Size = New System.Drawing.Size(116, 13)
    Me.Label6.TabIndex = 7
    Me.Label6.Text = "Part-of-speech options:"
    '
    'dgvPosVern
    '
    Me.dgvPosVern.AllowUserToAddRows = False
    Me.dgvPosVern.AllowUserToDeleteRows = False
    Me.dgvPosVern.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dgvPosVern.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
    Me.dgvPosVern.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvPosVern.Location = New System.Drawing.Point(132, 127)
    Me.dgvPosVern.Name = "dgvPosVern"
    Me.dgvPosVern.ReadOnly = True
    Me.dgvPosVern.Size = New System.Drawing.Size(277, 188)
    Me.dgvPosVern.TabIndex = 8
    '
    'Label7
    '
    Me.Label7.AutoSize = True
    Me.Label7.Location = New System.Drawing.Point(132, 111)
    Me.Label7.Name = "Label7"
    Me.Label7.Size = New System.Drawing.Size(130, 13)
    Me.Label7.TabIndex = 9
    Me.Label7.Text = "Lookup (POS/Vernacular)"
    '
    'cmdOk
    '
    Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.cmdOk.Location = New System.Drawing.Point(225, 29)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(58, 23)
    Me.cmdOk.TabIndex = 10
    Me.cmdOk.Text = "&Accept"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'cmdSkip
    '
    Me.cmdSkip.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdSkip.DialogResult = System.Windows.Forms.DialogResult.Ignore
    Me.cmdSkip.Location = New System.Drawing.Point(289, 29)
    Me.cmdSkip.Name = "cmdSkip"
    Me.cmdSkip.Size = New System.Drawing.Size(48, 23)
    Me.cmdSkip.TabIndex = 11
    Me.cmdSkip.Text = "&Skip"
    Me.cmdSkip.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(343, 29)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(60, 23)
    Me.cmdCancel.TabIndex = 13
    Me.cmdCancel.Text = "&Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'frmPosTag
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(629, 534)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.tbTrn)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.tbOrg)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmPosTag"
    Me.Text = "Part-of-speech tagger"
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.GroupBox1.ResumeLayout(False)
    Me.SplitContainer1.Panel1.ResumeLayout(False)
    Me.SplitContainer1.Panel2.ResumeLayout(False)
    Me.SplitContainer1.Panel2.PerformLayout()
    Me.SplitContainer1.ResumeLayout(False)
    CType(Me.dgvPosTag, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.dgvPosVern, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents tbOrg As System.Windows.Forms.RichTextBox
  Friend WithEvents tbTrn As System.Windows.Forms.RichTextBox
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
  Friend WithEvents dgvPosTag As System.Windows.Forms.DataGridView
  Friend WithEvents tbVern As System.Windows.Forms.TextBox
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents Label7 As System.Windows.Forms.Label
  Friend WithEvents dgvPosVern As System.Windows.Forms.DataGridView
  Friend WithEvents Label6 As System.Windows.Forms.Label
  Friend WithEvents lboPos As System.Windows.Forms.ListBox
  Friend WithEvents tbMorph As System.Windows.Forms.TextBox
  Friend WithEvents Label5 As System.Windows.Forms.Label
  Friend WithEvents tbPos As System.Windows.Forms.TextBox
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdSkip As System.Windows.Forms.Button
  Friend WithEvents cmdOk As System.Windows.Forms.Button
End Class
