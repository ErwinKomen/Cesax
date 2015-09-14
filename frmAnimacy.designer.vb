<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAnimacy
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAnimacy))
    Me.GroupBox1 = New System.Windows.Forms.GroupBox
    Me.wbChain = New System.Windows.Forms.WebBrowser
    Me.GroupBox2 = New System.Windows.Forms.GroupBox
    Me.rbUnknown = New System.Windows.Forms.RadioButton
    Me.rbInanim = New System.Windows.Forms.RadioButton
    Me.rbAnim = New System.Windows.Forms.RadioButton
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.cmdOk = New System.Windows.Forms.Button
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar
    Me.GroupBox1.SuspendLayout()
    Me.GroupBox2.SuspendLayout()
    Me.StatusStrip1.SuspendLayout()
    Me.SuspendLayout()
    '
    'GroupBox1
    '
    Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GroupBox1.Controls.Add(Me.wbChain)
    Me.GroupBox1.Location = New System.Drawing.Point(2, 108)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(762, 253)
    Me.GroupBox1.TabIndex = 0
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "This chain..."
    '
    'wbChain
    '
    Me.wbChain.Dock = System.Windows.Forms.DockStyle.Fill
    Me.wbChain.Location = New System.Drawing.Point(3, 16)
    Me.wbChain.MinimumSize = New System.Drawing.Size(20, 20)
    Me.wbChain.Name = "wbChain"
    Me.wbChain.Size = New System.Drawing.Size(756, 234)
    Me.wbChain.TabIndex = 0
    '
    'GroupBox2
    '
    Me.GroupBox2.Controls.Add(Me.rbUnknown)
    Me.GroupBox2.Controls.Add(Me.rbInanim)
    Me.GroupBox2.Controls.Add(Me.rbAnim)
    Me.GroupBox2.Location = New System.Drawing.Point(2, 12)
    Me.GroupBox2.Name = "GroupBox2"
    Me.GroupBox2.Size = New System.Drawing.Size(200, 90)
    Me.GroupBox2.TabIndex = 1
    Me.GroupBox2.TabStop = False
    Me.GroupBox2.Text = "Select the animacy"
    '
    'rbUnknown
    '
    Me.rbUnknown.AutoSize = True
    Me.rbUnknown.Location = New System.Drawing.Point(10, 65)
    Me.rbUnknown.Name = "rbUnknown"
    Me.rbUnknown.Size = New System.Drawing.Size(71, 17)
    Me.rbUnknown.TabIndex = 2
    Me.rbUnknown.TabStop = True
    Me.rbUnknown.Text = "&Unknown"
    Me.rbUnknown.UseVisualStyleBackColor = True
    '
    'rbInanim
    '
    Me.rbInanim.AutoSize = True
    Me.rbInanim.Location = New System.Drawing.Point(10, 42)
    Me.rbInanim.Name = "rbInanim"
    Me.rbInanim.Size = New System.Drawing.Size(71, 17)
    Me.rbInanim.TabIndex = 1
    Me.rbInanim.TabStop = True
    Me.rbInanim.Text = "&Inanimate"
    Me.rbInanim.UseVisualStyleBackColor = True
    '
    'rbAnim
    '
    Me.rbAnim.AutoSize = True
    Me.rbAnim.Location = New System.Drawing.Point(10, 19)
    Me.rbAnim.Name = "rbAnim"
    Me.rbAnim.Size = New System.Drawing.Size(63, 17)
    Me.rbAnim.TabIndex = 0
    Me.rbAnim.TabStop = True
    Me.rbAnim.Text = "&Animate"
    Me.rbAnim.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(686, 79)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 2
    Me.cmdCancel.Text = "&Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'cmdOk
    '
    Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOk.Location = New System.Drawing.Point(686, 50)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(75, 23)
    Me.cmdOk.TabIndex = 3
    Me.cmdOk.Text = "&Ok"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'Timer1
    '
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 364)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(766, 22)
    Me.StatusStrip1.TabIndex = 4
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
    'frmAnimacy
    '
    Me.AcceptButton = Me.cmdOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(766, 386)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Controls.Add(Me.cmdOk)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.GroupBox2)
    Me.Controls.Add(Me.GroupBox1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmAnimacy"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Animacy..."
    Me.GroupBox1.ResumeLayout(False)
    Me.GroupBox2.ResumeLayout(False)
    Me.GroupBox2.PerformLayout()
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents wbChain As System.Windows.Forms.WebBrowser
  Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
  Friend WithEvents rbUnknown As System.Windows.Forms.RadioButton
  Friend WithEvents rbInanim As System.Windows.Forms.RadioButton
  Friend WithEvents rbAnim As System.Windows.Forms.RadioButton
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
End Class
