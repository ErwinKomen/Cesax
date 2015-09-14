<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMetaInfo
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
    Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
    Me.OK_Button = New System.Windows.Forms.Button()
    Me.Cancel_Button = New System.Windows.Forms.Button()
    Me.tbMetaFile = New System.Windows.Forms.TextBox()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.tbMetaDir = New System.Windows.Forms.TextBox()
    Me.cmdMetaFile = New System.Windows.Forms.Button()
    Me.cmdMetaDir = New System.Windows.Forms.Button()
    Me.grpFormat = New System.Windows.Forms.GroupBox()
    Me.rbFieldAlternated = New System.Windows.Forms.RadioButton()
    Me.rbFieldColumns = New System.Windows.Forms.RadioButton()
    Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
    Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
    Me.TableLayoutPanel1.SuspendLayout()
    Me.grpFormat.SuspendLayout()
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
    Me.TableLayoutPanel1.Location = New System.Drawing.Point(464, 186)
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
    'tbMetaFile
    '
    Me.tbMetaFile.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbMetaFile.Location = New System.Drawing.Point(12, 31)
    Me.tbMetaFile.Name = "tbMetaFile"
    Me.tbMetaFile.ReadOnly = True
    Me.tbMetaFile.Size = New System.Drawing.Size(557, 20)
    Me.tbMetaFile.TabIndex = 1
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(12, 15)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(257, 13)
    Me.Label1.TabIndex = 2
    Me.Label1.Text = "Tab-separated (csv/txt) file containing the meta-data:"
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(12, 67)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(207, 13)
    Me.Label2.TabIndex = 4
    Me.Label2.Text = "Directory where the .psdx files are located:"
    '
    'tbMetaDir
    '
    Me.tbMetaDir.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbMetaDir.Location = New System.Drawing.Point(12, 83)
    Me.tbMetaDir.Name = "tbMetaDir"
    Me.tbMetaDir.ReadOnly = True
    Me.tbMetaDir.Size = New System.Drawing.Size(557, 20)
    Me.tbMetaDir.TabIndex = 3
    '
    'cmdMetaFile
    '
    Me.cmdMetaFile.Location = New System.Drawing.Point(575, 31)
    Me.cmdMetaFile.Name = "cmdMetaFile"
    Me.cmdMetaFile.Size = New System.Drawing.Size(32, 20)
    Me.cmdMetaFile.TabIndex = 5
    Me.cmdMetaFile.Text = "..."
    Me.cmdMetaFile.UseVisualStyleBackColor = True
    '
    'cmdMetaDir
    '
    Me.cmdMetaDir.Location = New System.Drawing.Point(575, 83)
    Me.cmdMetaDir.Name = "cmdMetaDir"
    Me.cmdMetaDir.Size = New System.Drawing.Size(32, 20)
    Me.cmdMetaDir.TabIndex = 6
    Me.cmdMetaDir.Text = "..."
    Me.cmdMetaDir.UseVisualStyleBackColor = True
    '
    'grpFormat
    '
    Me.grpFormat.Controls.Add(Me.rbFieldAlternated)
    Me.grpFormat.Controls.Add(Me.rbFieldColumns)
    Me.grpFormat.Location = New System.Drawing.Point(12, 112)
    Me.grpFormat.Name = "grpFormat"
    Me.grpFormat.Size = New System.Drawing.Size(446, 100)
    Me.grpFormat.TabIndex = 7
    Me.grpFormat.TabStop = False
    Me.grpFormat.Text = "Format of the meta-data within the (csv/text) file"
    '
    'rbFieldAlternated
    '
    Me.rbFieldAlternated.AutoSize = True
    Me.rbFieldAlternated.Location = New System.Drawing.Point(6, 63)
    Me.rbFieldAlternated.Name = "rbFieldAlternated"
    Me.rbFieldAlternated.Size = New System.Drawing.Size(197, 17)
    Me.rbFieldAlternated.TabIndex = 1
    Me.rbFieldAlternated.TabStop = True
    Me.rbFieldAlternated.Text = "&Alternated field name/value columns"
    Me.rbFieldAlternated.UseVisualStyleBackColor = True
    '
    'rbFieldColumns
    '
    Me.rbFieldColumns.AutoSize = True
    Me.rbFieldColumns.Location = New System.Drawing.Point(6, 30)
    Me.rbFieldColumns.Name = "rbFieldColumns"
    Me.rbFieldColumns.Size = New System.Drawing.Size(158, 17)
    Me.rbFieldColumns.TabIndex = 0
    Me.rbFieldColumns.TabStop = True
    Me.rbFieldColumns.Text = "&Field names - value columns"
    Me.rbFieldColumns.UseVisualStyleBackColor = True
    '
    'OpenFileDialog1
    '
    Me.OpenFileDialog1.FileName = "OpenFileDialog1"
    '
    'frmMetaInfo
    '
    Me.AcceptButton = Me.OK_Button
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.Cancel_Button
    Me.ClientSize = New System.Drawing.Size(622, 227)
    Me.Controls.Add(Me.grpFormat)
    Me.Controls.Add(Me.cmdMetaDir)
    Me.Controls.Add(Me.cmdMetaFile)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.tbMetaDir)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.tbMetaFile)
    Me.Controls.Add(Me.TableLayoutPanel1)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmMetaInfo"
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Meta data information"
    Me.TableLayoutPanel1.ResumeLayout(False)
    Me.grpFormat.ResumeLayout(False)
    Me.grpFormat.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
  Friend WithEvents OK_Button As System.Windows.Forms.Button
  Friend WithEvents Cancel_Button As System.Windows.Forms.Button
  Friend WithEvents tbMetaFile As System.Windows.Forms.TextBox
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents tbMetaDir As System.Windows.Forms.TextBox
  Friend WithEvents cmdMetaFile As System.Windows.Forms.Button
  Friend WithEvents cmdMetaDir As System.Windows.Forms.Button
  Friend WithEvents grpFormat As System.Windows.Forms.GroupBox
  Friend WithEvents rbFieldAlternated As System.Windows.Forms.RadioButton
  Friend WithEvents rbFieldColumns As System.Windows.Forms.RadioButton
  Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
  Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog

End Class
