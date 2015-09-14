<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRevision
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
    Me.Label1 = New System.Windows.Forms.Label
    Me.tbWho = New System.Windows.Forms.TextBox
    Me.Label2 = New System.Windows.Forms.Label
    Me.dtWhen = New System.Windows.Forms.DateTimePicker
    Me.Label3 = New System.Windows.Forms.Label
    Me.tbComment = New System.Windows.Forms.RichTextBox
    Me.tbWhen = New System.Windows.Forms.TextBox
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.Label4 = New System.Windows.Forms.Label
    Me.tbFile = New System.Windows.Forms.TextBox
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
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
    Me.TableLayoutPanel1.Location = New System.Drawing.Point(277, 210)
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
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(12, 9)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(33, 13)
    Me.Label1.TabIndex = 1
    Me.Label1.Text = "Who:"
    '
    'tbWho
    '
    Me.tbWho.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbWho.Location = New System.Drawing.Point(51, 6)
    Me.tbWho.Name = "tbWho"
    Me.tbWho.Size = New System.Drawing.Size(369, 20)
    Me.tbWho.TabIndex = 2
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(12, 35)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(39, 13)
    Me.Label2.TabIndex = 3
    Me.Label2.Text = "When:"
    '
    'dtWhen
    '
    Me.dtWhen.Location = New System.Drawing.Point(250, 32)
    Me.dtWhen.Name = "dtWhen"
    Me.dtWhen.Size = New System.Drawing.Size(170, 20)
    Me.dtWhen.TabIndex = 4
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(12, 64)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(132, 13)
    Me.Label3.TabIndex = 5
    Me.Label3.Text = "Comments for this revision:"
    '
    'tbComment
    '
    Me.tbComment.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbComment.Location = New System.Drawing.Point(15, 80)
    Me.tbComment.Name = "tbComment"
    Me.tbComment.Size = New System.Drawing.Size(405, 124)
    Me.tbComment.TabIndex = 6
    Me.tbComment.Text = ""
    '
    'tbWhen
    '
    Me.tbWhen.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbWhen.Location = New System.Drawing.Point(51, 32)
    Me.tbWhen.Name = "tbWhen"
    Me.tbWhen.ReadOnly = True
    Me.tbWhen.Size = New System.Drawing.Size(193, 20)
    Me.tbWhen.TabIndex = 7
    '
    'Timer1
    '
    '
    'Label4
    '
    Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(12, 222)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(26, 13)
    Me.Label4.TabIndex = 8
    Me.Label4.Text = "File:"
    '
    'tbFile
    '
    Me.tbFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.tbFile.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbFile.Location = New System.Drawing.Point(44, 219)
    Me.tbFile.Name = "tbFile"
    Me.tbFile.ReadOnly = True
    Me.tbFile.Size = New System.Drawing.Size(200, 20)
    Me.tbFile.TabIndex = 9
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 242)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(435, 22)
    Me.StatusStrip1.TabIndex = 10
    Me.StatusStrip1.Text = "StatusStrip1"
    '
    'ToolStripStatusLabel1
    '
    Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
    Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(19, 17)
    Me.ToolStripStatusLabel1.Text = "..."
    '
    'frmRevision
    '
    Me.AcceptButton = Me.OK_Button
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.Cancel_Button
    Me.ClientSize = New System.Drawing.Size(435, 264)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Controls.Add(Me.tbFile)
    Me.Controls.Add(Me.Label4)
    Me.Controls.Add(Me.tbWhen)
    Me.Controls.Add(Me.tbComment)
    Me.Controls.Add(Me.Label3)
    Me.Controls.Add(Me.dtWhen)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.tbWho)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.TableLayoutPanel1)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmRevision"
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Supply the information for this revision"
    Me.TableLayoutPanel1.ResumeLayout(False)
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
  Friend WithEvents Cancel_Button As System.Windows.Forms.Button
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents tbWho As System.Windows.Forms.TextBox
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents dtWhen As System.Windows.Forms.DateTimePicker
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents tbComment As System.Windows.Forms.RichTextBox
  Friend WithEvents tbWhen As System.Windows.Forms.TextBox
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents tbFile As System.Windows.Forms.TextBox
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel

End Class
