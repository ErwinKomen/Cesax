<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDirs
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
    Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
    Me.OK_Button = New System.Windows.Forms.Button
    Me.Cancel_Button = New System.Windows.Forms.Button
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar
    Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
    Me.lbSrc = New System.Windows.Forms.Label
    Me.tbSrcDir = New System.Windows.Forms.TextBox
    Me.cmdSrcDir = New System.Windows.Forms.Button
    Me.cmdDstDir = New System.Windows.Forms.Button
    Me.tbDstDir = New System.Windows.Forms.TextBox
    Me.lbDst = New System.Windows.Forms.Label
    Me.TableLayoutPanel1.SuspendLayout()
    Me.StatusStrip1.SuspendLayout()
    Me.SuspendLayout()
    '
    'TableLayoutPanel1
    '
    Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.TableLayoutPanel1.ColumnCount = 2
    Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
    Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
    Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
    Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
    Me.TableLayoutPanel1.Location = New System.Drawing.Point(274, 101)
    Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
    Me.TableLayoutPanel1.RowCount = 1
    Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
    Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
    Me.TableLayoutPanel1.TabIndex = 0
    '
    'OK_Button
    '
    Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
    Me.OK_Button.Location = New System.Drawing.Point(3, 3)
    Me.OK_Button.Name = "OK_Button"
    Me.OK_Button.Size = New System.Drawing.Size(67, 23)
    Me.OK_Button.TabIndex = 0
    Me.OK_Button.Text = "OK"
    '
    'Cancel_Button
    '
    Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
    Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
    Me.Cancel_Button.Name = "Cancel_Button"
    Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
    Me.Cancel_Button.TabIndex = 1
    Me.Cancel_Button.Text = "Cancel"
    '
    'Timer1
    '
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 133)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(435, 22)
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
    'lbSrc
    '
    Me.lbSrc.AutoSize = True
    Me.lbSrc.Location = New System.Drawing.Point(12, 9)
    Me.lbSrc.Name = "lbSrc"
    Me.lbSrc.Size = New System.Drawing.Size(87, 13)
    Me.lbSrc.TabIndex = 2
    Me.lbSrc.Text = "Source directory:"
    '
    'tbSrcDir
    '
    Me.tbSrcDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbSrcDir.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbSrcDir.Location = New System.Drawing.Point(15, 25)
    Me.tbSrcDir.Name = "tbSrcDir"
    Me.tbSrcDir.ReadOnly = True
    Me.tbSrcDir.Size = New System.Drawing.Size(358, 20)
    Me.tbSrcDir.TabIndex = 3
    '
    'cmdSrcDir
    '
    Me.cmdSrcDir.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdSrcDir.Location = New System.Drawing.Point(387, 25)
    Me.cmdSrcDir.Name = "cmdSrcDir"
    Me.cmdSrcDir.Size = New System.Drawing.Size(33, 20)
    Me.cmdSrcDir.TabIndex = 4
    Me.cmdSrcDir.Text = "..."
    Me.cmdSrcDir.UseVisualStyleBackColor = True
    '
    'cmdDstDir
    '
    Me.cmdDstDir.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdDstDir.Location = New System.Drawing.Point(387, 64)
    Me.cmdDstDir.Name = "cmdDstDir"
    Me.cmdDstDir.Size = New System.Drawing.Size(33, 20)
    Me.cmdDstDir.TabIndex = 7
    Me.cmdDstDir.Text = "..."
    Me.cmdDstDir.UseVisualStyleBackColor = True
    '
    'tbDstDir
    '
    Me.tbDstDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDstDir.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbDstDir.Location = New System.Drawing.Point(15, 64)
    Me.tbDstDir.Name = "tbDstDir"
    Me.tbDstDir.ReadOnly = True
    Me.tbDstDir.Size = New System.Drawing.Size(358, 20)
    Me.tbDstDir.TabIndex = 6
    '
    'lbDst
    '
    Me.lbDst.AutoSize = True
    Me.lbDst.Location = New System.Drawing.Point(12, 48)
    Me.lbDst.Name = "lbDst"
    Me.lbDst.Size = New System.Drawing.Size(106, 13)
    Me.lbDst.TabIndex = 5
    Me.lbDst.Text = "Destination directory:"
    '
    'frmDirs
    '
    Me.AcceptButton = Me.OK_Button
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.Cancel_Button
    Me.ClientSize = New System.Drawing.Size(435, 155)
    Me.Controls.Add(Me.cmdDstDir)
    Me.Controls.Add(Me.tbDstDir)
    Me.Controls.Add(Me.lbDst)
    Me.Controls.Add(Me.cmdSrcDir)
    Me.Controls.Add(Me.tbSrcDir)
    Me.Controls.Add(Me.lbSrc)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Controls.Add(Me.TableLayoutPanel1)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmDirs"
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Specify the input and the output directories..."
    Me.TableLayoutPanel1.ResumeLayout(False)
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
  Friend WithEvents OK_Button As System.Windows.Forms.Button
  Friend WithEvents Cancel_Button As System.Windows.Forms.Button
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents lbSrc As System.Windows.Forms.Label
  Friend WithEvents tbSrcDir As System.Windows.Forms.TextBox
  Friend WithEvents cmdSrcDir As System.Windows.Forms.Button
  Friend WithEvents cmdDstDir As System.Windows.Forms.Button
  Friend WithEvents tbDstDir As System.Windows.Forms.TextBox
  Friend WithEvents lbDst As System.Windows.Forms.Label

End Class
