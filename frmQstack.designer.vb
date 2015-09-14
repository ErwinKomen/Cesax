<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmQstack
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmQstack))
    Me.dgvQstack = New System.Windows.Forms.DataGridView
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.cmdExport = New System.Windows.Forms.Button
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.GroupBox1 = New System.Windows.Forms.GroupBox
    Me.cmdReset = New System.Windows.Forms.Button
    Me.cmdDelete = New System.Windows.Forms.Button
    CType(Me.dgvQstack, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.StatusStrip1.SuspendLayout()
    Me.GroupBox1.SuspendLayout()
    Me.SuspendLayout()
    '
    'dgvQstack
    '
    Me.dgvQstack.AllowUserToAddRows = False
    Me.dgvQstack.AllowUserToDeleteRows = False
    Me.dgvQstack.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
    Me.dgvQstack.BorderStyle = System.Windows.Forms.BorderStyle.None
    Me.dgvQstack.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvQstack.Dock = System.Windows.Forms.DockStyle.Fill
    Me.dgvQstack.Location = New System.Drawing.Point(3, 16)
    Me.dgvQstack.MultiSelect = False
    Me.dgvQstack.Name = "dgvQstack"
    Me.dgvQstack.ReadOnly = True
    Me.dgvQstack.Size = New System.Drawing.Size(382, 209)
    Me.dgvQstack.TabIndex = 0
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 272)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(412, 22)
    Me.StatusStrip1.TabIndex = 1
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
    'cmdExport
    '
    Me.cmdExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdExport.Location = New System.Drawing.Point(325, 246)
    Me.cmdExport.Name = "cmdExport"
    Me.cmdExport.Size = New System.Drawing.Size(75, 23)
    Me.cmdExport.TabIndex = 2
    Me.cmdExport.Text = "E&xport"
    Me.cmdExport.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(244, 246)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 3
    Me.cmdCancel.Text = "&Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'GroupBox1
    '
    Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GroupBox1.Controls.Add(Me.dgvQstack)
    Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(388, 228)
    Me.GroupBox1.TabIndex = 4
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Relations"
    '
    'cmdReset
    '
    Me.cmdReset.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdReset.DialogResult = System.Windows.Forms.DialogResult.Retry
    Me.cmdReset.Location = New System.Drawing.Point(163, 246)
    Me.cmdReset.Name = "cmdReset"
    Me.cmdReset.Size = New System.Drawing.Size(75, 23)
    Me.cmdReset.TabIndex = 5
    Me.cmdReset.Text = "&Reset"
    Me.cmdReset.UseVisualStyleBackColor = True
    '
    'cmdDelete
    '
    Me.cmdDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.cmdDelete.DialogResult = System.Windows.Forms.DialogResult.Retry
    Me.cmdDelete.Location = New System.Drawing.Point(12, 246)
    Me.cmdDelete.Name = "cmdDelete"
    Me.cmdDelete.Size = New System.Drawing.Size(144, 23)
    Me.cmdDelete.TabIndex = 6
    Me.cmdDelete.Text = "&Delete selected condition"
    Me.cmdDelete.UseVisualStyleBackColor = True
    '
    'frmQstack
    '
    Me.AcceptButton = Me.cmdExport
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(412, 294)
    Me.Controls.Add(Me.cmdDelete)
    Me.Controls.Add(Me.cmdReset)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.cmdExport)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.MinimumSize = New System.Drawing.Size(420, 185)
    Me.Name = "frmQstack"
    Me.Text = "Query building relations"
    CType(Me.dgvQstack, System.ComponentModel.ISupportInitialize).EndInit()
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.GroupBox1.ResumeLayout(False)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents dgvQstack As System.Windows.Forms.DataGridView
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents cmdExport As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents cmdReset As System.Windows.Forms.Button
  Friend WithEvents cmdDelete As System.Windows.Forms.Button
End Class
