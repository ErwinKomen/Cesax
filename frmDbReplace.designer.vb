<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDbReplace
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDbReplace))
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.cboFindFeatureName = New System.Windows.Forms.ComboBox
    Me.cmdAddF = New System.Windows.Forms.Button
    Me.cmdDelF = New System.Windows.Forms.Button
    Me.tbFindFeatureValue = New System.Windows.Forms.TextBox
    Me.cmdResetF = New System.Windows.Forms.Button
    Me.GroupBox1 = New System.Windows.Forms.GroupBox
    Me.tbFindResultRepl = New System.Windows.Forms.TextBox
    Me.Label5 = New System.Windows.Forms.Label
    Me.tbFindResultValue = New System.Windows.Forms.TextBox
    Me.cmdDelR = New System.Windows.Forms.Button
    Me.cmdAddR = New System.Windows.Forms.Button
    Me.cboFindResultName = New System.Windows.Forms.ComboBox
    Me.GroupBox2 = New System.Windows.Forms.GroupBox
    Me.tbFindFeatureRepl = New System.Windows.Forms.TextBox
    Me.Label4 = New System.Windows.Forms.Label
    Me.Label2 = New System.Windows.Forms.Label
    Me.dgvFindRepl = New System.Windows.Forms.DataGridView
    Me.cmdAll = New System.Windows.Forms.Button
    Me.cmdStep = New System.Windows.Forms.Button
    Me.Label1 = New System.Windows.Forms.Label
    Me.tbFRname = New System.Windows.Forms.TextBox
    Me.Label3 = New System.Windows.Forms.Label
    Me.Label6 = New System.Windows.Forms.Label
    Me.cmdFill = New System.Windows.Forms.Button
    Me.GroupBox3 = New System.Windows.Forms.GroupBox
    Me.cmdFRdel = New System.Windows.Forms.Button
    Me.cmdFRkeep = New System.Windows.Forms.Button
    Me.tbFRfind = New System.Windows.Forms.RichTextBox
    Me.tbFRrepl = New System.Windows.Forms.RichTextBox
    Me.StatusStrip1.SuspendLayout()
    Me.GroupBox1.SuspendLayout()
    Me.GroupBox2.SuspendLayout()
    CType(Me.dgvFindRepl, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.GroupBox3.SuspendLayout()
    Me.SuspendLayout()
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 480)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(675, 22)
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
    Me.ToolStripProgressBar1.Size = New System.Drawing.Size(100, 16)
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(426, 454)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 3
    Me.cmdCancel.Text = "&Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'Timer1
    '
    '
    'cboFindFeatureName
    '
    Me.cboFindFeatureName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cboFindFeatureName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboFindFeatureName.FormattingEnabled = True
    Me.cboFindFeatureName.Location = New System.Drawing.Point(75, 19)
    Me.cboFindFeatureName.Name = "cboFindFeatureName"
    Me.cboFindFeatureName.Size = New System.Drawing.Size(195, 21)
    Me.cboFindFeatureName.TabIndex = 5
    '
    'cmdAddF
    '
    Me.cmdAddF.Location = New System.Drawing.Point(13, 45)
    Me.cmdAddF.Name = "cmdAddF"
    Me.cmdAddF.Size = New System.Drawing.Size(56, 21)
    Me.cmdAddF.TabIndex = 6
    Me.cmdAddF.Text = "&Add"
    Me.cmdAddF.UseVisualStyleBackColor = True
    '
    'cmdDelF
    '
    Me.cmdDelF.Location = New System.Drawing.Point(13, 19)
    Me.cmdDelF.Name = "cmdDelF"
    Me.cmdDelF.Size = New System.Drawing.Size(56, 21)
    Me.cmdDelF.TabIndex = 7
    Me.cmdDelF.Text = "Re&move"
    Me.cmdDelF.UseVisualStyleBackColor = True
    '
    'tbFindFeatureValue
    '
    Me.tbFindFeatureValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFindFeatureValue.Location = New System.Drawing.Point(75, 46)
    Me.tbFindFeatureValue.Name = "tbFindFeatureValue"
    Me.tbFindFeatureValue.Size = New System.Drawing.Size(195, 20)
    Me.tbFindFeatureValue.TabIndex = 8
    '
    'cmdResetF
    '
    Me.cmdResetF.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdResetF.Location = New System.Drawing.Point(575, 17)
    Me.cmdResetF.Name = "cmdResetF"
    Me.cmdResetF.Size = New System.Drawing.Size(88, 61)
    Me.cmdResetF.TabIndex = 9
    Me.cmdResetF.Text = "&Reset"
    Me.cmdResetF.UseVisualStyleBackColor = True
    '
    'GroupBox1
    '
    Me.GroupBox1.Controls.Add(Me.tbFindResultRepl)
    Me.GroupBox1.Controls.Add(Me.Label5)
    Me.GroupBox1.Controls.Add(Me.tbFindResultValue)
    Me.GroupBox1.Controls.Add(Me.cmdDelR)
    Me.GroupBox1.Controls.Add(Me.cmdAddR)
    Me.GroupBox1.Controls.Add(Me.cboFindResultName)
    Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(275, 128)
    Me.GroupBox1.TabIndex = 10
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Find //Result[@...]"
    '
    'tbFindResultRepl
    '
    Me.tbFindResultRepl.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFindResultRepl.Location = New System.Drawing.Point(13, 92)
    Me.tbFindResultRepl.Name = "tbFindResultRepl"
    Me.tbFindResultRepl.Size = New System.Drawing.Size(256, 20)
    Me.tbFindResultRepl.TabIndex = 15
    '
    'Label5
    '
    Me.Label5.AutoSize = True
    Me.Label5.Location = New System.Drawing.Point(10, 76)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(73, 13)
    Me.Label5.TabIndex = 14
    Me.Label5.Text = "Replacement:"
    '
    'tbFindResultValue
    '
    Me.tbFindResultValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFindResultValue.Location = New System.Drawing.Point(75, 46)
    Me.tbFindResultValue.Name = "tbFindResultValue"
    Me.tbFindResultValue.Size = New System.Drawing.Size(194, 20)
    Me.tbFindResultValue.TabIndex = 13
    '
    'cmdDelR
    '
    Me.cmdDelR.Location = New System.Drawing.Point(13, 19)
    Me.cmdDelR.Name = "cmdDelR"
    Me.cmdDelR.Size = New System.Drawing.Size(56, 21)
    Me.cmdDelR.TabIndex = 12
    Me.cmdDelR.Text = "Remo&ve"
    Me.cmdDelR.UseVisualStyleBackColor = True
    '
    'cmdAddR
    '
    Me.cmdAddR.Location = New System.Drawing.Point(13, 45)
    Me.cmdAddR.Name = "cmdAddR"
    Me.cmdAddR.Size = New System.Drawing.Size(56, 21)
    Me.cmdAddR.TabIndex = 11
    Me.cmdAddR.Text = "A&dd"
    Me.cmdAddR.UseVisualStyleBackColor = True
    '
    'cboFindResultName
    '
    Me.cboFindResultName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cboFindResultName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboFindResultName.FormattingEnabled = True
    Me.cboFindResultName.Location = New System.Drawing.Point(75, 19)
    Me.cboFindResultName.Name = "cboFindResultName"
    Me.cboFindResultName.Size = New System.Drawing.Size(194, 21)
    Me.cboFindResultName.TabIndex = 10
    '
    'GroupBox2
    '
    Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GroupBox2.Controls.Add(Me.tbFindFeatureRepl)
    Me.GroupBox2.Controls.Add(Me.Label4)
    Me.GroupBox2.Controls.Add(Me.cboFindFeatureName)
    Me.GroupBox2.Controls.Add(Me.cmdAddF)
    Me.GroupBox2.Controls.Add(Me.cmdDelF)
    Me.GroupBox2.Controls.Add(Me.tbFindFeatureValue)
    Me.GroupBox2.Location = New System.Drawing.Point(293, 12)
    Me.GroupBox2.Name = "GroupBox2"
    Me.GroupBox2.Size = New System.Drawing.Size(276, 128)
    Me.GroupBox2.TabIndex = 11
    Me.GroupBox2.TabStop = False
    Me.GroupBox2.Text = "Find //Result[Feature[@Name...]]"
    '
    'tbFindFeatureRepl
    '
    Me.tbFindFeatureRepl.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFindFeatureRepl.Location = New System.Drawing.Point(13, 92)
    Me.tbFindFeatureRepl.Name = "tbFindFeatureRepl"
    Me.tbFindFeatureRepl.Size = New System.Drawing.Size(257, 20)
    Me.tbFindFeatureRepl.TabIndex = 10
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(10, 76)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(73, 13)
    Me.Label4.TabIndex = 9
    Me.Label4.Text = "Replacement:"
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(9, 143)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(195, 13)
    Me.Label2.TabIndex = 13
    Me.Label2.Text = "Summary of find and replacment criteria:"
    '
    'dgvFindRepl
    '
    Me.dgvFindRepl.AllowUserToAddRows = False
    Me.dgvFindRepl.AllowUserToDeleteRows = False
    Me.dgvFindRepl.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dgvFindRepl.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
    Me.dgvFindRepl.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvFindRepl.Location = New System.Drawing.Point(12, 159)
    Me.dgvFindRepl.Name = "dgvFindRepl"
    Me.dgvFindRepl.ReadOnly = True
    Me.dgvFindRepl.Size = New System.Drawing.Size(388, 318)
    Me.dgvFindRepl.TabIndex = 14
    '
    'cmdAll
    '
    Me.cmdAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdAll.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdAll.Location = New System.Drawing.Point(588, 454)
    Me.cmdAll.Name = "cmdAll"
    Me.cmdAll.Size = New System.Drawing.Size(75, 23)
    Me.cmdAll.TabIndex = 26
    Me.cmdAll.Text = "&All"
    Me.cmdAll.UseVisualStyleBackColor = True
    '
    'cmdStep
    '
    Me.cmdStep.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdStep.Location = New System.Drawing.Point(507, 454)
    Me.cmdStep.Name = "cmdStep"
    Me.cmdStep.Size = New System.Drawing.Size(75, 23)
    Me.cmdStep.TabIndex = 25
    Me.cmdStep.Text = "&Step"
    Me.cmdStep.UseVisualStyleBackColor = True
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(6, 16)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(38, 13)
    Me.Label1.TabIndex = 27
    Me.Label1.Text = "Name:"
    '
    'tbFRname
    '
    Me.tbFRname.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFRname.Location = New System.Drawing.Point(9, 32)
    Me.tbFRname.Name = "tbFRname"
    Me.tbFRname.Size = New System.Drawing.Size(242, 20)
    Me.tbFRname.TabIndex = 28
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(6, 63)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(59, 13)
    Me.Label3.TabIndex = 29
    Me.Label3.Text = "Find value:"
    '
    'Label6
    '
    Me.Label6.AutoSize = True
    Me.Label6.Location = New System.Drawing.Point(6, 154)
    Me.Label6.Name = "Label6"
    Me.Label6.Size = New System.Drawing.Size(102, 13)
    Me.Label6.TabIndex = 31
    Me.Label6.Text = "Replacement value:"
    '
    'cmdFill
    '
    Me.cmdFill.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdFill.Location = New System.Drawing.Point(575, 83)
    Me.cmdFill.Name = "cmdFill"
    Me.cmdFill.Size = New System.Drawing.Size(88, 57)
    Me.cmdFill.TabIndex = 33
    Me.cmdFill.Text = "&Fill"
    Me.cmdFill.UseVisualStyleBackColor = True
    '
    'GroupBox3
    '
    Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GroupBox3.Controls.Add(Me.tbFRrepl)
    Me.GroupBox3.Controls.Add(Me.tbFRfind)
    Me.GroupBox3.Controls.Add(Me.cmdFRkeep)
    Me.GroupBox3.Controls.Add(Me.cmdFRdel)
    Me.GroupBox3.Controls.Add(Me.Label1)
    Me.GroupBox3.Controls.Add(Me.tbFRname)
    Me.GroupBox3.Controls.Add(Me.Label3)
    Me.GroupBox3.Controls.Add(Me.Label6)
    Me.GroupBox3.Location = New System.Drawing.Point(406, 146)
    Me.GroupBox3.Name = "GroupBox3"
    Me.GroupBox3.Size = New System.Drawing.Size(257, 302)
    Me.GroupBox3.TabIndex = 34
    Me.GroupBox3.TabStop = False
    Me.GroupBox3.Text = "Selected criterion"
    '
    'cmdFRdel
    '
    Me.cmdFRdel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdFRdel.Location = New System.Drawing.Point(9, 273)
    Me.cmdFRdel.Name = "cmdFRdel"
    Me.cmdFRdel.Size = New System.Drawing.Size(242, 23)
    Me.cmdFRdel.TabIndex = 33
    Me.cmdFRdel.Text = "&Delete this criterion"
    Me.cmdFRdel.UseVisualStyleBackColor = True
    '
    'cmdFRkeep
    '
    Me.cmdFRkeep.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdFRkeep.Location = New System.Drawing.Point(9, 244)
    Me.cmdFRkeep.Name = "cmdFRkeep"
    Me.cmdFRkeep.Size = New System.Drawing.Size(242, 23)
    Me.cmdFRkeep.TabIndex = 34
    Me.cmdFRkeep.Text = "&Keep only this criterion"
    Me.cmdFRkeep.UseVisualStyleBackColor = True
    '
    'tbFRfind
    '
    Me.tbFRfind.Location = New System.Drawing.Point(9, 79)
    Me.tbFRfind.Name = "tbFRfind"
    Me.tbFRfind.Size = New System.Drawing.Size(242, 72)
    Me.tbFRfind.TabIndex = 35
    Me.tbFRfind.Text = ""
    '
    'tbFRrepl
    '
    Me.tbFRrepl.Location = New System.Drawing.Point(9, 170)
    Me.tbFRrepl.Name = "tbFRrepl"
    Me.tbFRrepl.Size = New System.Drawing.Size(242, 68)
    Me.tbFRrepl.TabIndex = 36
    Me.tbFRrepl.Text = ""
    '
    'frmDbReplace
    '
    Me.AcceptButton = Me.cmdStep
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(675, 502)
    Me.Controls.Add(Me.GroupBox3)
    Me.Controls.Add(Me.cmdFill)
    Me.Controls.Add(Me.cmdAll)
    Me.Controls.Add(Me.cmdStep)
    Me.Controls.Add(Me.dgvFindRepl)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.GroupBox2)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.cmdResetF)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmDbReplace"
    Me.Text = "Corpus database Global Find and Replace"
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.GroupBox1.ResumeLayout(False)
    Me.GroupBox1.PerformLayout()
    Me.GroupBox2.ResumeLayout(False)
    Me.GroupBox2.PerformLayout()
    CType(Me.dgvFindRepl, System.ComponentModel.ISupportInitialize).EndInit()
    Me.GroupBox3.ResumeLayout(False)
    Me.GroupBox3.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents cboFindFeatureName As System.Windows.Forms.ComboBox
  Friend WithEvents cmdAddF As System.Windows.Forms.Button
  Friend WithEvents cmdDelF As System.Windows.Forms.Button
  Friend WithEvents tbFindFeatureValue As System.Windows.Forms.TextBox
  Friend WithEvents cmdResetF As System.Windows.Forms.Button
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents tbFindResultValue As System.Windows.Forms.TextBox
  Friend WithEvents cmdDelR As System.Windows.Forms.Button
  Friend WithEvents cmdAddR As System.Windows.Forms.Button
  Friend WithEvents cboFindResultName As System.Windows.Forms.ComboBox
  Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents tbFindResultRepl As System.Windows.Forms.TextBox
  Friend WithEvents Label5 As System.Windows.Forms.Label
  Friend WithEvents tbFindFeatureRepl As System.Windows.Forms.TextBox
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents dgvFindRepl As System.Windows.Forms.DataGridView
  Friend WithEvents cmdAll As System.Windows.Forms.Button
  Friend WithEvents cmdStep As System.Windows.Forms.Button
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents tbFRname As System.Windows.Forms.TextBox
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents Label6 As System.Windows.Forms.Label
  Friend WithEvents cmdFill As System.Windows.Forms.Button
  Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
  Friend WithEvents cmdFRdel As System.Windows.Forms.Button
  Friend WithEvents cmdFRkeep As System.Windows.Forms.Button
  Friend WithEvents tbFRfind As System.Windows.Forms.RichTextBox
  Friend WithEvents tbFRrepl As System.Windows.Forms.RichTextBox
End Class
