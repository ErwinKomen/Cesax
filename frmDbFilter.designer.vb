<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDbFilter
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDbFilter))
    Me.Label1 = New System.Windows.Forms.Label
    Me.cmdOk = New System.Windows.Forms.Button
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.tbFilterXpath = New System.Windows.Forms.RichTextBox
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.cboFeatureName = New System.Windows.Forms.ComboBox
    Me.cmdAddF = New System.Windows.Forms.Button
    Me.cmdDelF = New System.Windows.Forms.Button
    Me.tbFeatureValue = New System.Windows.Forms.TextBox
    Me.cmdResetF = New System.Windows.Forms.Button
    Me.GroupBox1 = New System.Windows.Forms.GroupBox
    Me.tbResultValue = New System.Windows.Forms.TextBox
    Me.cmdDelR = New System.Windows.Forms.Button
    Me.cmdAddR = New System.Windows.Forms.Button
    Me.cboResultName = New System.Windows.Forms.ComboBox
    Me.GroupBox2 = New System.Windows.Forms.GroupBox
    Me.StatusStrip1.SuspendLayout()
    Me.GroupBox1.SuspendLayout()
    Me.GroupBox2.SuspendLayout()
    Me.SuspendLayout()
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(12, 105)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(236, 13)
    Me.Label1.TabIndex = 0
    Me.Label1.Text = "Provide your filter expression using Xpath syntax:"
    '
    'cmdOk
    '
    Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOk.Location = New System.Drawing.Point(587, 421)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(75, 23)
    Me.cmdOk.TabIndex = 1
    Me.cmdOk.Text = "&Ok"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 447)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(674, 22)
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
    Me.cmdCancel.Location = New System.Drawing.Point(506, 421)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 3
    Me.cmdCancel.Text = "&Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'tbFilterXpath
    '
    Me.tbFilterXpath.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFilterXpath.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbFilterXpath.Location = New System.Drawing.Point(12, 121)
    Me.tbFilterXpath.Name = "tbFilterXpath"
    Me.tbFilterXpath.Size = New System.Drawing.Size(650, 294)
    Me.tbFilterXpath.TabIndex = 4
    Me.tbFilterXpath.Text = ""
    '
    'Timer1
    '
    '
    'cboFeatureName
    '
    Me.cboFeatureName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cboFeatureName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboFeatureName.FormattingEnabled = True
    Me.cboFeatureName.Location = New System.Drawing.Point(68, 19)
    Me.cboFeatureName.Name = "cboFeatureName"
    Me.cboFeatureName.Size = New System.Drawing.Size(184, 21)
    Me.cboFeatureName.TabIndex = 5
    '
    'cmdAddF
    '
    Me.cmdAddF.Location = New System.Drawing.Point(6, 45)
    Me.cmdAddF.Name = "cmdAddF"
    Me.cmdAddF.Size = New System.Drawing.Size(56, 21)
    Me.cmdAddF.TabIndex = 6
    Me.cmdAddF.Text = "&Add"
    Me.cmdAddF.UseVisualStyleBackColor = True
    '
    'cmdDelF
    '
    Me.cmdDelF.Location = New System.Drawing.Point(6, 19)
    Me.cmdDelF.Name = "cmdDelF"
    Me.cmdDelF.Size = New System.Drawing.Size(56, 21)
    Me.cmdDelF.TabIndex = 7
    Me.cmdDelF.Text = "Re&move"
    Me.cmdDelF.UseVisualStyleBackColor = True
    '
    'tbFeatureValue
    '
    Me.tbFeatureValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFeatureValue.Location = New System.Drawing.Point(68, 46)
    Me.tbFeatureValue.Name = "tbFeatureValue"
    Me.tbFeatureValue.Size = New System.Drawing.Size(184, 20)
    Me.tbFeatureValue.TabIndex = 8
    '
    'cmdResetF
    '
    Me.cmdResetF.Location = New System.Drawing.Point(557, 17)
    Me.cmdResetF.Name = "cmdResetF"
    Me.cmdResetF.Size = New System.Drawing.Size(105, 85)
    Me.cmdResetF.TabIndex = 9
    Me.cmdResetF.Text = "&Reset"
    Me.cmdResetF.UseVisualStyleBackColor = True
    '
    'GroupBox1
    '
    Me.GroupBox1.Controls.Add(Me.tbResultValue)
    Me.GroupBox1.Controls.Add(Me.cmdDelR)
    Me.GroupBox1.Controls.Add(Me.cmdAddR)
    Me.GroupBox1.Controls.Add(Me.cboResultName)
    Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(275, 90)
    Me.GroupBox1.TabIndex = 10
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "//Result[@...]"
    '
    'tbResultValue
    '
    Me.tbResultValue.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbResultValue.Location = New System.Drawing.Point(75, 46)
    Me.tbResultValue.Name = "tbResultValue"
    Me.tbResultValue.Size = New System.Drawing.Size(194, 20)
    Me.tbResultValue.TabIndex = 13
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
    'cboResultName
    '
    Me.cboResultName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cboResultName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboResultName.FormattingEnabled = True
    Me.cboResultName.Location = New System.Drawing.Point(75, 19)
    Me.cboResultName.Name = "cboResultName"
    Me.cboResultName.Size = New System.Drawing.Size(194, 21)
    Me.cboResultName.TabIndex = 10
    '
    'GroupBox2
    '
    Me.GroupBox2.Controls.Add(Me.cboFeatureName)
    Me.GroupBox2.Controls.Add(Me.cmdAddF)
    Me.GroupBox2.Controls.Add(Me.cmdDelF)
    Me.GroupBox2.Controls.Add(Me.tbFeatureValue)
    Me.GroupBox2.Location = New System.Drawing.Point(293, 12)
    Me.GroupBox2.Name = "GroupBox2"
    Me.GroupBox2.Size = New System.Drawing.Size(258, 90)
    Me.GroupBox2.TabIndex = 11
    Me.GroupBox2.TabStop = False
    Me.GroupBox2.Text = "//Result[Feature[@Name...]]"
    '
    'frmDbFilter
    '
    Me.AcceptButton = Me.cmdOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(674, 469)
    Me.Controls.Add(Me.GroupBox2)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.cmdResetF)
    Me.Controls.Add(Me.tbFilterXpath)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Controls.Add(Me.cmdOk)
    Me.Controls.Add(Me.Label1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmDbFilter"
    Me.Text = "Corpus database filter"
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.GroupBox1.ResumeLayout(False)
    Me.GroupBox1.PerformLayout()
    Me.GroupBox2.ResumeLayout(False)
    Me.GroupBox2.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents tbFilterXpath As System.Windows.Forms.RichTextBox
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents cboFeatureName As System.Windows.Forms.ComboBox
  Friend WithEvents cmdAddF As System.Windows.Forms.Button
  Friend WithEvents cmdDelF As System.Windows.Forms.Button
  Friend WithEvents tbFeatureValue As System.Windows.Forms.TextBox
  Friend WithEvents cmdResetF As System.Windows.Forms.Button
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents tbResultValue As System.Windows.Forms.TextBox
  Friend WithEvents cmdDelR As System.Windows.Forms.Button
  Friend WithEvents cmdAddR As System.Windows.Forms.Button
  Friend WithEvents cboResultName As System.Windows.Forms.ComboBox
  Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
End Class
