<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSpss
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSpss))
    Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
    Me.OK_Button = New System.Windows.Forms.Button
    Me.Cancel_Button = New System.Windows.Forms.Button
    Me.lboExclude = New System.Windows.Forms.ListBox
    Me.Label1 = New System.Windows.Forms.Label
    Me.Label2 = New System.Windows.Forms.Label
    Me.Label3 = New System.Windows.Forms.Label
    Me.lboIndep = New System.Windows.Forms.ListBox
    Me.lboDep = New System.Windows.Forms.ListBox
    Me.lboFeatures = New System.Windows.Forms.ListBox
    Me.GroupBox1 = New System.Windows.Forms.GroupBox
    Me.rbDep = New System.Windows.Forms.RadioButton
    Me.rbIndep = New System.Windows.Forms.RadioButton
    Me.rbNoStat = New System.Windows.Forms.RadioButton
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar
    Me.GroupBox2 = New System.Windows.Forms.GroupBox
    Me.rbUseCat = New System.Windows.Forms.RadioButton
    Me.rbUseFeature = New System.Windows.Forms.RadioButton
    Me.TableLayoutPanel1.SuspendLayout()
    Me.GroupBox1.SuspendLayout()
    Me.StatusStrip1.SuspendLayout()
    Me.GroupBox2.SuspendLayout()
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
    Me.TableLayoutPanel1.Location = New System.Drawing.Point(348, 309)
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
    'lboExclude
    '
    Me.lboExclude.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.lboExclude.FormattingEnabled = True
    Me.lboExclude.Location = New System.Drawing.Point(348, 25)
    Me.lboExclude.Name = "lboExclude"
    Me.lboExclude.Size = New System.Drawing.Size(143, 69)
    Me.lboExclude.TabIndex = 1
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(12, 9)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(51, 13)
    Me.Label1.TabIndex = 2
    Me.Label1.Text = "Features:"
    '
    'Label2
    '
    Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(345, 9)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(130, 13)
    Me.Label2.TabIndex = 3
    Me.Label2.Text = "Do not include in statistics"
    '
    'Label3
    '
    Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(345, 105)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(112, 13)
    Me.Label3.TabIndex = 5
    Me.Label3.Text = "Independant variables"
    '
    'lboIndep
    '
    Me.lboIndep.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.lboIndep.FormattingEnabled = True
    Me.lboIndep.Location = New System.Drawing.Point(348, 121)
    Me.lboIndep.Name = "lboIndep"
    Me.lboIndep.Size = New System.Drawing.Size(143, 173)
    Me.lboIndep.TabIndex = 4
    '
    'lboDep
    '
    Me.lboDep.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.lboDep.FormattingEnabled = True
    Me.lboDep.Location = New System.Drawing.Point(31, 65)
    Me.lboDep.Name = "lboDep"
    Me.lboDep.Size = New System.Drawing.Size(143, 56)
    Me.lboDep.TabIndex = 6
    Me.lboDep.Visible = False
    '
    'lboFeatures
    '
    Me.lboFeatures.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.lboFeatures.FormattingEnabled = True
    Me.lboFeatures.Location = New System.Drawing.Point(15, 25)
    Me.lboFeatures.Name = "lboFeatures"
    Me.lboFeatures.Size = New System.Drawing.Size(121, 277)
    Me.lboFeatures.TabIndex = 8
    '
    'GroupBox1
    '
    Me.GroupBox1.Controls.Add(Me.rbDep)
    Me.GroupBox1.Controls.Add(Me.rbIndep)
    Me.GroupBox1.Controls.Add(Me.rbNoStat)
    Me.GroupBox1.Location = New System.Drawing.Point(149, 24)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(190, 94)
    Me.GroupBox1.TabIndex = 9
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Characteristics"
    '
    'rbDep
    '
    Me.rbDep.AutoSize = True
    Me.rbDep.Location = New System.Drawing.Point(6, 65)
    Me.rbDep.Name = "rbDep"
    Me.rbDep.Size = New System.Drawing.Size(78, 17)
    Me.rbDep.TabIndex = 2
    Me.rbDep.TabStop = True
    Me.rbDep.Text = "&Dependant"
    Me.rbDep.UseVisualStyleBackColor = True
    '
    'rbIndep
    '
    Me.rbIndep.AutoSize = True
    Me.rbIndep.Location = New System.Drawing.Point(6, 42)
    Me.rbIndep.Name = "rbIndep"
    Me.rbIndep.Size = New System.Drawing.Size(85, 17)
    Me.rbIndep.TabIndex = 1
    Me.rbIndep.TabStop = True
    Me.rbIndep.Text = "&Independant"
    Me.rbIndep.UseVisualStyleBackColor = True
    '
    'rbNoStat
    '
    Me.rbNoStat.AutoSize = True
    Me.rbNoStat.Location = New System.Drawing.Point(6, 19)
    Me.rbNoStat.Name = "rbNoStat"
    Me.rbNoStat.Size = New System.Drawing.Size(82, 17)
    Me.rbNoStat.TabIndex = 0
    Me.rbNoStat.TabStop = True
    Me.rbNoStat.Text = "&No statistics"
    Me.rbNoStat.UseVisualStyleBackColor = True
    '
    'Timer1
    '
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 341)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(506, 22)
    Me.StatusStrip1.TabIndex = 10
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
    'GroupBox2
    '
    Me.GroupBox2.Controls.Add(Me.rbUseFeature)
    Me.GroupBox2.Controls.Add(Me.rbUseCat)
    Me.GroupBox2.Controls.Add(Me.lboDep)
    Me.GroupBox2.Location = New System.Drawing.Point(149, 174)
    Me.GroupBox2.Name = "GroupBox2"
    Me.GroupBox2.Size = New System.Drawing.Size(190, 128)
    Me.GroupBox2.TabIndex = 11
    Me.GroupBox2.TabStop = False
    Me.GroupBox2.Text = "Dependant variable (outcome)"
    '
    'rbUseCat
    '
    Me.rbUseCat.AutoSize = True
    Me.rbUseCat.Location = New System.Drawing.Point(6, 19)
    Me.rbUseCat.Name = "rbUseCat"
    Me.rbUseCat.Size = New System.Drawing.Size(168, 17)
    Me.rbUseCat.TabIndex = 8
    Me.rbUseCat.TabStop = True
    Me.rbUseCat.Text = "Use [&Cat] (sSubcategorization)"
    Me.rbUseCat.UseVisualStyleBackColor = True
    '
    'rbUseFeature
    '
    Me.rbUseFeature.AutoSize = True
    Me.rbUseFeature.Location = New System.Drawing.Point(6, 42)
    Me.rbUseFeature.Name = "rbUseFeature"
    Me.rbUseFeature.Size = New System.Drawing.Size(129, 17)
    Me.rbUseFeature.TabIndex = 9
    Me.rbUseFeature.TabStop = True
    Me.rbUseFeature.Text = "Use &Feature selected:"
    Me.rbUseFeature.UseVisualStyleBackColor = True
    '
    'frmSpss
    '
    Me.AcceptButton = Me.OK_Button
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.Cancel_Button
    Me.ClientSize = New System.Drawing.Size(506, 363)
    Me.Controls.Add(Me.GroupBox2)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.lboFeatures)
    Me.Controls.Add(Me.Label3)
    Me.Controls.Add(Me.lboIndep)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.lboExclude)
    Me.Controls.Add(Me.TableLayoutPanel1)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmSpss"
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Indicate the character of the features"
    Me.TableLayoutPanel1.ResumeLayout(False)
    Me.GroupBox1.ResumeLayout(False)
    Me.GroupBox1.PerformLayout()
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.GroupBox2.ResumeLayout(False)
    Me.GroupBox2.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
  Friend WithEvents OK_Button As System.Windows.Forms.Button
  Friend WithEvents Cancel_Button As System.Windows.Forms.Button
  Friend WithEvents lboExclude As System.Windows.Forms.ListBox
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents lboIndep As System.Windows.Forms.ListBox
  Friend WithEvents lboDep As System.Windows.Forms.ListBox
  Friend WithEvents lboFeatures As System.Windows.Forms.ListBox
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents rbDep As System.Windows.Forms.RadioButton
  Friend WithEvents rbIndep As System.Windows.Forms.RadioButton
  Friend WithEvents rbNoStat As System.Windows.Forms.RadioButton
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
  Friend WithEvents rbUseFeature As System.Windows.Forms.RadioButton
  Friend WithEvents rbUseCat As System.Windows.Forms.RadioButton

End Class
