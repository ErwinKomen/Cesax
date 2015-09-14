<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmHelp
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmHelp))
    Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
    Me.cmdBack = New System.Windows.Forms.ToolStripButton
    Me.cmdForward = New System.Windows.Forms.ToolStripButton
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar
    Me.wbHelp = New System.Windows.Forms.WebBrowser
    Me.ToolStrip1.SuspendLayout()
    Me.StatusStrip1.SuspendLayout()
    Me.SuspendLayout()
    '
    'ToolStrip1
    '
    Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.cmdBack, Me.cmdForward})
    Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
    Me.ToolStrip1.Name = "ToolStrip1"
    Me.ToolStrip1.Size = New System.Drawing.Size(688, 25)
    Me.ToolStrip1.TabIndex = 0
    Me.ToolStrip1.Text = "ToolStrip1"
    '
    'cmdBack
    '
    Me.cmdBack.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
    Me.cmdBack.Image = CType(resources.GetObject("cmdBack.Image"), System.Drawing.Image)
    Me.cmdBack.ImageTransparentColor = System.Drawing.Color.Magenta
    Me.cmdBack.Name = "cmdBack"
    Me.cmdBack.Size = New System.Drawing.Size(23, 22)
    Me.cmdBack.Text = "<"
    Me.cmdBack.ToolTipText = "ALT + <"
    '
    'cmdForward
    '
    Me.cmdForward.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
    Me.cmdForward.Image = CType(resources.GetObject("cmdForward.Image"), System.Drawing.Image)
    Me.cmdForward.ImageTransparentColor = System.Drawing.Color.Magenta
    Me.cmdForward.Name = "cmdForward"
    Me.cmdForward.Size = New System.Drawing.Size(23, 22)
    Me.cmdForward.Text = ">"
    Me.cmdForward.ToolTipText = "ALT + >"
    '
    'Timer1
    '
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 756)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(688, 22)
    Me.StatusStrip1.TabIndex = 2
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
    Me.ToolStripProgressBar1.Size = New System.Drawing.Size(400, 16)
    '
    'wbHelp
    '
    Me.wbHelp.Dock = System.Windows.Forms.DockStyle.Fill
    Me.wbHelp.Location = New System.Drawing.Point(0, 25)
    Me.wbHelp.MinimumSize = New System.Drawing.Size(20, 20)
    Me.wbHelp.Name = "wbHelp"
    Me.wbHelp.Size = New System.Drawing.Size(688, 731)
    Me.wbHelp.TabIndex = 3
    '
    'frmHelp
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(688, 778)
    Me.Controls.Add(Me.wbHelp)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Controls.Add(Me.ToolStrip1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.KeyPreview = True
    Me.Name = "frmHelp"
    Me.Text = "Support for CESAX"
    Me.ToolStrip1.ResumeLayout(False)
    Me.ToolStrip1.PerformLayout()
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
  Friend WithEvents cmdBack As System.Windows.Forms.ToolStripButton
  Friend WithEvents cmdForward As System.Windows.Forms.ToolStripButton
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents wbHelp As System.Windows.Forms.WebBrowser
End Class
