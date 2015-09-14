<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChain
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmChain))
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
    Me.cmdOk = New System.Windows.Forms.Button
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.GroupBox1 = New System.Windows.Forms.GroupBox
    Me.rbRecalc = New System.Windows.Forms.RadioButton
    Me.rbUpdate = New System.Windows.Forms.RadioButton
    Me.GroupBox2 = New System.Windows.Forms.GroupBox
    Me.cmdDir = New System.Windows.Forms.Button
    Me.tbDir = New System.Windows.Forms.TextBox
    Me.rbAll = New System.Windows.Forms.RadioButton
    Me.rbCurrent = New System.Windows.Forms.RadioButton
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
    Me.StatusStrip1.SuspendLayout()
    Me.GroupBox1.SuspendLayout()
    Me.GroupBox2.SuspendLayout()
    Me.SuspendLayout()
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 200)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(475, 22)
    Me.StatusStrip1.TabIndex = 0
    Me.StatusStrip1.Text = "StatusStrip1"
    '
    'ToolStripStatusLabel1
    '
    Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
    Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(19, 17)
    Me.ToolStripStatusLabel1.Text = "..."
    '
    'cmdOk
    '
    Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.cmdOk.Location = New System.Drawing.Point(307, 174)
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
    Me.cmdCancel.Location = New System.Drawing.Point(388, 174)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 2
    Me.cmdCancel.Text = "Ca&ncel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'GroupBox1
    '
    Me.GroupBox1.Controls.Add(Me.rbRecalc)
    Me.GroupBox1.Controls.Add(Me.rbUpdate)
    Me.GroupBox1.Location = New System.Drawing.Point(12, 116)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(155, 73)
    Me.GroupBox1.TabIndex = 3
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Action"
    '
    'rbRecalc
    '
    Me.rbRecalc.AutoSize = True
    Me.rbRecalc.Location = New System.Drawing.Point(6, 42)
    Me.rbRecalc.Name = "rbRecalc"
    Me.rbRecalc.Size = New System.Drawing.Size(116, 17)
    Me.rbRecalc.TabIndex = 1
    Me.rbRecalc.TabStop = True
    Me.rbRecalc.Text = "&Recalculate chains"
    Me.rbRecalc.UseVisualStyleBackColor = True
    '
    'rbUpdate
    '
    Me.rbUpdate.AutoSize = True
    Me.rbUpdate.Location = New System.Drawing.Point(6, 19)
    Me.rbUpdate.Name = "rbUpdate"
    Me.rbUpdate.Size = New System.Drawing.Size(113, 17)
    Me.rbUpdate.TabIndex = 0
    Me.rbUpdate.TabStop = True
    Me.rbUpdate.Text = "&Update if outdated"
    Me.rbUpdate.UseVisualStyleBackColor = True
    '
    'GroupBox2
    '
    Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GroupBox2.Controls.Add(Me.cmdDir)
    Me.GroupBox2.Controls.Add(Me.tbDir)
    Me.GroupBox2.Controls.Add(Me.rbAll)
    Me.GroupBox2.Controls.Add(Me.rbCurrent)
    Me.GroupBox2.Location = New System.Drawing.Point(12, 12)
    Me.GroupBox2.Name = "GroupBox2"
    Me.GroupBox2.Size = New System.Drawing.Size(452, 98)
    Me.GroupBox2.TabIndex = 4
    Me.GroupBox2.TabStop = False
    Me.GroupBox2.Text = "Scope"
    '
    'cmdDir
    '
    Me.cmdDir.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdDir.Location = New System.Drawing.Point(417, 65)
    Me.cmdDir.Name = "cmdDir"
    Me.cmdDir.Size = New System.Drawing.Size(29, 20)
    Me.cmdDir.TabIndex = 3
    Me.cmdDir.Text = "..."
    Me.cmdDir.UseVisualStyleBackColor = True
    '
    'tbDir
    '
    Me.tbDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDir.Location = New System.Drawing.Point(7, 65)
    Me.tbDir.Name = "tbDir"
    Me.tbDir.Size = New System.Drawing.Size(407, 20)
    Me.tbDir.TabIndex = 2
    '
    'rbAll
    '
    Me.rbAll.AutoSize = True
    Me.rbAll.Location = New System.Drawing.Point(6, 42)
    Me.rbAll.Name = "rbAll"
    Me.rbAll.Size = New System.Drawing.Size(134, 17)
    Me.rbAll.TabIndex = 1
    Me.rbAll.TabStop = True
    Me.rbAll.Text = "&All texts in this directory"
    Me.rbAll.UseVisualStyleBackColor = True
    '
    'rbCurrent
    '
    Me.rbCurrent.AutoSize = True
    Me.rbCurrent.Location = New System.Drawing.Point(6, 19)
    Me.rbCurrent.Name = "rbCurrent"
    Me.rbCurrent.Size = New System.Drawing.Size(101, 17)
    Me.rbCurrent.TabIndex = 0
    Me.rbCurrent.TabStop = True
    Me.rbCurrent.Text = "&Current text only"
    Me.rbCurrent.UseVisualStyleBackColor = True
    '
    'Timer1
    '
    '
    'frmChain
    '
    Me.AcceptButton = Me.cmdOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(475, 222)
    Me.Controls.Add(Me.GroupBox2)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.cmdOk)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmChain"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "List coreferential chains ..."
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.GroupBox1.ResumeLayout(False)
    Me.GroupBox1.PerformLayout()
    Me.GroupBox2.ResumeLayout(False)
    Me.GroupBox2.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents rbRecalc As System.Windows.Forms.RadioButton
  Friend WithEvents rbUpdate As System.Windows.Forms.RadioButton
  Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
  Friend WithEvents rbAll As System.Windows.Forms.RadioButton
  Friend WithEvents rbCurrent As System.Windows.Forms.RadioButton
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents cmdDir As System.Windows.Forms.Button
  Friend WithEvents tbDir As System.Windows.Forms.TextBox
  Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
End Class
