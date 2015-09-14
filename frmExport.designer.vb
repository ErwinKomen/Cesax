<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmExport
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
    Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
    Me.OK_Button = New System.Windows.Forms.Button
    Me.Cancel_Button = New System.Windows.Forms.Button
    Me.rbPsd = New System.Windows.Forms.RadioButton
    Me.GroupBox1 = New System.Windows.Forms.GroupBox
    Me.rbPsdSimple = New System.Windows.Forms.RadioButton
    Me.rbPde = New System.Windows.Forms.RadioButton
    Me.TableLayoutPanel1.SuspendLayout()
    Me.GroupBox1.SuspendLayout()
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
    Me.TableLayoutPanel1.Location = New System.Drawing.Point(211, 146)
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
    'rbPsd
    '
    Me.rbPsd.AutoSize = True
    Me.rbPsd.Location = New System.Drawing.Point(6, 19)
    Me.rbPsd.Name = "rbPsd"
    Me.rbPsd.Size = New System.Drawing.Size(205, 17)
    Me.rbPsd.TabIndex = 3
    Me.rbPsd.TabStop = True
    Me.rbPsd.Text = "psd: &Penn-Treebank 2 text format (full)"
    Me.rbPsd.UseVisualStyleBackColor = True
    '
    'GroupBox1
    '
    Me.GroupBox1.Controls.Add(Me.rbPsdSimple)
    Me.GroupBox1.Controls.Add(Me.rbPde)
    Me.GroupBox1.Controls.Add(Me.rbPsd)
    Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(345, 128)
    Me.GroupBox1.TabIndex = 5
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Available formats"
    '
    'rbPsdSimple
    '
    Me.rbPsdSimple.AutoSize = True
    Me.rbPsdSimple.Location = New System.Drawing.Point(6, 59)
    Me.rbPsdSimple.Name = "rbPsdSimple"
    Me.rbPsdSimple.Size = New System.Drawing.Size(234, 17)
    Me.rbPsdSimple.TabIndex = 5
    Me.rbPsdSimple.TabStop = True
    Me.rbPsdSimple.Text = "psd: &Penn-Treebank 2 text format (simplified)"
    Me.rbPsdSimple.UseVisualStyleBackColor = True
    '
    'rbPde
    '
    Me.rbPde.AutoSize = True
    Me.rbPde.Location = New System.Drawing.Point(6, 105)
    Me.rbPde.Name = "rbPde"
    Me.rbPde.Size = New System.Drawing.Size(195, 17)
    Me.rbPde.TabIndex = 4
    Me.rbPde.TabStop = True
    Me.rbPde.Text = "psdy: &English translation XML format"
    Me.rbPde.UseVisualStyleBackColor = True
    '
    'frmExport
    '
    Me.AcceptButton = Me.OK_Button
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.Cancel_Button
    Me.ClientSize = New System.Drawing.Size(377, 187)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.TableLayoutPanel1)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmExport"
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Choose the export format"
    Me.TableLayoutPanel1.ResumeLayout(False)
    Me.GroupBox1.ResumeLayout(False)
    Me.GroupBox1.PerformLayout()
    Me.ResumeLayout(False)

  End Sub
  Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
  Friend WithEvents OK_Button As System.Windows.Forms.Button
  Friend WithEvents Cancel_Button As System.Windows.Forms.Button
  Friend WithEvents rbPsd As System.Windows.Forms.RadioButton
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents rbPde As System.Windows.Forms.RadioButton
  Friend WithEvents rbPsdSimple As System.Windows.Forms.RadioButton

End Class
