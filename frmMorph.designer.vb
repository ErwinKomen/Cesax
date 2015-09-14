<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMorph
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMorph))
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.cmdOk = New System.Windows.Forms.Button
    Me.Label1 = New System.Windows.Forms.Label
    Me.wbSentence = New System.Windows.Forms.WebBrowser
    Me.Label2 = New System.Windows.Forms.Label
    Me.lboOptions = New System.Windows.Forms.ListBox
    Me.cmdNone = New System.Windows.Forms.Button
    Me.cmdAdd = New System.Windows.Forms.Button
    Me.tbInfo = New System.Windows.Forms.RichTextBox
    Me.Label3 = New System.Windows.Forms.Label
    Me.Label4 = New System.Windows.Forms.Label
    Me.tbFeat = New System.Windows.Forms.RichTextBox
    Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
    Me.cmdFull = New System.Windows.Forms.Button
    Me.tbFull = New System.Windows.Forms.TextBox
    Me.tbRestricted = New System.Windows.Forms.TextBox
    Me.cmdRestricted = New System.Windows.Forms.Button
    Me.StatusStrip1.SuspendLayout()
    Me.SplitContainer1.Panel1.SuspendLayout()
    Me.SplitContainer1.Panel2.SuspendLayout()
    Me.SplitContainer1.SuspendLayout()
    Me.SuspendLayout()
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 531)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(940, 22)
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
    Me.ToolStripProgressBar1.Size = New System.Drawing.Size(100, 16)
    '
    'Timer1
    '
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(853, 476)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 1
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'cmdOk
    '
    Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOk.Location = New System.Drawing.Point(772, 505)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(75, 23)
    Me.cmdOk.TabIndex = 2
    Me.cmdOk.Text = "&Ok"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(0, 10)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(56, 13)
    Me.Label1.TabIndex = 3
    Me.Label1.Text = "Sentence:"
    '
    'wbSentence
    '
    Me.wbSentence.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.wbSentence.Location = New System.Drawing.Point(3, 26)
    Me.wbSentence.MinimumSize = New System.Drawing.Size(20, 20)
    Me.wbSentence.Name = "wbSentence"
    Me.wbSentence.Size = New System.Drawing.Size(934, 138)
    Me.wbSentence.TabIndex = 4
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(0, 2)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(178, 13)
    Me.Label2.TabIndex = 5
    Me.Label2.Text = "Choose one of the following options:"
    '
    'lboOptions
    '
    Me.lboOptions.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.lboOptions.FormattingEnabled = True
    Me.lboOptions.Location = New System.Drawing.Point(3, 18)
    Me.lboOptions.Name = "lboOptions"
    Me.lboOptions.Size = New System.Drawing.Size(754, 277)
    Me.lboOptions.TabIndex = 6
    '
    'cmdNone
    '
    Me.cmdNone.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdNone.Location = New System.Drawing.Point(853, 505)
    Me.cmdNone.Name = "cmdNone"
    Me.cmdNone.Size = New System.Drawing.Size(75, 23)
    Me.cmdNone.TabIndex = 7
    Me.cmdNone.Text = "&None"
    Me.cmdNone.UseVisualStyleBackColor = True
    '
    'cmdAdd
    '
    Me.cmdAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdAdd.Location = New System.Drawing.Point(772, 476)
    Me.cmdAdd.Name = "cmdAdd"
    Me.cmdAdd.Size = New System.Drawing.Size(75, 23)
    Me.cmdAdd.TabIndex = 8
    Me.cmdAdd.Text = "&Add"
    Me.cmdAdd.UseVisualStyleBackColor = True
    '
    'tbInfo
    '
    Me.tbInfo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbInfo.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbInfo.Location = New System.Drawing.Point(6, 18)
    Me.tbInfo.Name = "tbInfo"
    Me.tbInfo.ReadOnly = True
    Me.tbInfo.Size = New System.Drawing.Size(167, 112)
    Me.tbInfo.TabIndex = 9
    Me.tbInfo.Text = ""
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(3, 2)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(62, 13)
    Me.Label3.TabIndex = 10
    Me.Label3.Text = "Information:"
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(3, 139)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(97, 13)
    Me.Label4.TabIndex = 12
    Me.Label4.Text = "Features (editable):"
    '
    'tbFeat
    '
    Me.tbFeat.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFeat.Location = New System.Drawing.Point(6, 155)
    Me.tbFeat.Name = "tbFeat"
    Me.tbFeat.Size = New System.Drawing.Size(167, 141)
    Me.tbFeat.TabIndex = 11
    Me.tbFeat.Text = ""
    '
    'SplitContainer1
    '
    Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.SplitContainer1.Location = New System.Drawing.Point(0, 170)
    Me.SplitContainer1.Name = "SplitContainer1"
    '
    'SplitContainer1.Panel1
    '
    Me.SplitContainer1.Panel1.Controls.Add(Me.lboOptions)
    Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
    '
    'SplitContainer1.Panel2
    '
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label3)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbInfo)
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbFeat)
    Me.SplitContainer1.Size = New System.Drawing.Size(940, 304)
    Me.SplitContainer1.SplitterDistance = 760
    Me.SplitContainer1.TabIndex = 13
    '
    'cmdFull
    '
    Me.cmdFull.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdFull.Location = New System.Drawing.Point(3, 476)
    Me.cmdFull.Name = "cmdFull"
    Me.cmdFull.Size = New System.Drawing.Size(171, 23)
    Me.cmdFull.TabIndex = 14
    Me.cmdFull.Text = "&Full generalization:"
    Me.cmdFull.UseVisualStyleBackColor = True
    '
    'tbFull
    '
    Me.tbFull.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbFull.Location = New System.Drawing.Point(180, 478)
    Me.tbFull.Name = "tbFull"
    Me.tbFull.ReadOnly = True
    Me.tbFull.Size = New System.Drawing.Size(269, 20)
    Me.tbFull.TabIndex = 15
    '
    'tbRestricted
    '
    Me.tbRestricted.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbRestricted.Location = New System.Drawing.Point(180, 507)
    Me.tbRestricted.Name = "tbRestricted"
    Me.tbRestricted.ReadOnly = True
    Me.tbRestricted.Size = New System.Drawing.Size(269, 20)
    Me.tbRestricted.TabIndex = 17
    '
    'cmdRestricted
    '
    Me.cmdRestricted.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdRestricted.Location = New System.Drawing.Point(3, 505)
    Me.cmdRestricted.Name = "cmdRestricted"
    Me.cmdRestricted.Size = New System.Drawing.Size(171, 23)
    Me.cmdRestricted.TabIndex = 16
    Me.cmdRestricted.Text = "&Restricted generalization"
    Me.cmdRestricted.UseVisualStyleBackColor = True
    '
    'frmMorph
    '
    Me.AcceptButton = Me.cmdOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(940, 553)
    Me.Controls.Add(Me.tbRestricted)
    Me.Controls.Add(Me.cmdRestricted)
    Me.Controls.Add(Me.tbFull)
    Me.Controls.Add(Me.cmdFull)
    Me.Controls.Add(Me.SplitContainer1)
    Me.Controls.Add(Me.cmdAdd)
    Me.Controls.Add(Me.cmdNone)
    Me.Controls.Add(Me.wbSentence)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.cmdOk)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmMorph"
    Me.Text = "Choose the correct morphology features for this word"
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.SplitContainer1.Panel1.ResumeLayout(False)
    Me.SplitContainer1.Panel1.PerformLayout()
    Me.SplitContainer1.Panel2.ResumeLayout(False)
    Me.SplitContainer1.Panel2.PerformLayout()
    Me.SplitContainer1.ResumeLayout(False)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents wbSentence As System.Windows.Forms.WebBrowser
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents lboOptions As System.Windows.Forms.ListBox
  Friend WithEvents cmdNone As System.Windows.Forms.Button
  Friend WithEvents cmdAdd As System.Windows.Forms.Button
  Friend WithEvents tbInfo As System.Windows.Forms.RichTextBox
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents tbFeat As System.Windows.Forms.RichTextBox
  Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
  Friend WithEvents cmdFull As System.Windows.Forms.Button
  Friend WithEvents tbFull As System.Windows.Forms.TextBox
  Friend WithEvents tbRestricted As System.Windows.Forms.TextBox
  Friend WithEvents cmdRestricted As System.Windows.Forms.Button
End Class
