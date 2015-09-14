<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReport
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmReport))
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.MenuStrip1 = New System.Windows.Forms.MenuStrip
    Me.mnuFile = New System.Windows.Forms.ToolStripMenuItem
    Me.mnuFileClose = New System.Windows.Forms.ToolStripMenuItem
    Me.mnuFileSave = New System.Windows.Forms.ToolStripMenuItem
    Me.wbReport = New System.Windows.Forms.WebBrowser
    Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
    Me.StatusStrip1.SuspendLayout()
    Me.MenuStrip1.SuspendLayout()
    Me.SuspendLayout()
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 540)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(756, 22)
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
    'MenuStrip1
    '
    Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFile})
    Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
    Me.MenuStrip1.Name = "MenuStrip1"
    Me.MenuStrip1.Size = New System.Drawing.Size(756, 24)
    Me.MenuStrip1.TabIndex = 1
    Me.MenuStrip1.Text = "MenuStrip1"
    '
    'mnuFile
    '
    Me.mnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuFileSave, Me.mnuFileClose})
    Me.mnuFile.Name = "mnuFile"
    Me.mnuFile.Size = New System.Drawing.Size(35, 20)
    Me.mnuFile.Text = "&File"
    '
    'mnuFileClose
    '
    Me.mnuFileClose.Name = "mnuFileClose"
    Me.mnuFileClose.Size = New System.Drawing.Size(152, 22)
    Me.mnuFileClose.Text = "&Close"
    '
    'mnuFileSave
    '
    Me.mnuFileSave.Name = "mnuFileSave"
    Me.mnuFileSave.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.S), System.Windows.Forms.Keys)
    Me.mnuFileSave.Size = New System.Drawing.Size(152, 22)
    Me.mnuFileSave.Text = "&Save"
    '
    'wbReport
    '
    Me.wbReport.Dock = System.Windows.Forms.DockStyle.Fill
    Me.wbReport.Location = New System.Drawing.Point(0, 24)
    Me.wbReport.MinimumSize = New System.Drawing.Size(20, 20)
    Me.wbReport.Name = "wbReport"
    Me.wbReport.Size = New System.Drawing.Size(756, 516)
    Me.wbReport.TabIndex = 2
    '
    'frmReport
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(756, 562)
    Me.Controls.Add(Me.wbReport)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Controls.Add(Me.MenuStrip1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.MainMenuStrip = Me.MenuStrip1
    Me.Name = "frmReport"
    Me.Text = "Report on missing pronouns"
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.MenuStrip1.ResumeLayout(False)
    Me.MenuStrip1.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
  Friend WithEvents mnuFile As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents mnuFileSave As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents mnuFileClose As System.Windows.Forms.ToolStripMenuItem
  Friend WithEvents wbReport As System.Windows.Forms.WebBrowser
  Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
End Class
