<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFeatToText
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
    Me.components = New System.ComponentModel.Container()
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFeatToText))
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
    Me.cmdOk = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.Label1 = New System.Windows.Forms.Label()
    Me.tbTextDir = New System.Windows.Forms.TextBox()
    Me.cmdTextDir = New System.Windows.Forms.Button()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.tbFeatCat = New System.Windows.Forms.TextBox()
    Me.tbFeatname = New System.Windows.Forms.TextBox()
    Me.GroupBox1 = New System.Windows.Forms.GroupBox()
    Me.Label4 = New System.Windows.Forms.Label()
    Me.Label3 = New System.Windows.Forms.Label()
    Me.GroupBox2 = New System.Windows.Forms.GroupBox()
    Me.tbFeatBlackList = New System.Windows.Forms.TextBox()
    Me.Label6 = New System.Windows.Forms.Label()
    Me.tbFeatWhiteList = New System.Windows.Forms.TextBox()
    Me.rbRestrOnly = New System.Windows.Forms.RadioButton()
    Me.rbRestrNone = New System.Windows.Forms.RadioButton()
    Me.Label5 = New System.Windows.Forms.Label()
    Me.tbFeatFilter = New System.Windows.Forms.RichTextBox()
    Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
    Me.GroupBox3 = New System.Windows.Forms.GroupBox()
    Me.lboFeatDbase = New System.Windows.Forms.ListBox()
    Me.Label7 = New System.Windows.Forms.Label()
    Me.tbFeatDbase = New System.Windows.Forms.TextBox()
    Me.StatusStrip1.SuspendLayout()
    Me.GroupBox1.SuspendLayout()
    Me.GroupBox2.SuspendLayout()
    Me.GroupBox3.SuspendLayout()
    Me.SuspendLayout()
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 477)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(689, 22)
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
    Me.ToolStripProgressBar1.Size = New System.Drawing.Size(300, 16)
    '
    'cmdOk
    '
    Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOk.Location = New System.Drawing.Point(602, 451)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(75, 23)
    Me.cmdOk.TabIndex = 1
    Me.cmdOk.Text = "&Ok"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(521, 451)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 2
    Me.cmdCancel.Text = "&Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'Timer1
    '
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(12, 9)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(288, 13)
    Me.Label1.TabIndex = 3
    Me.Label1.Text = "Directory that contains the texts you want to add features to"
    '
    'tbTextDir
    '
    Me.tbTextDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbTextDir.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbTextDir.Location = New System.Drawing.Point(15, 25)
    Me.tbTextDir.Name = "tbTextDir"
    Me.tbTextDir.ReadOnly = True
    Me.tbTextDir.Size = New System.Drawing.Size(593, 20)
    Me.tbTextDir.TabIndex = 4
    '
    'cmdTextDir
    '
    Me.cmdTextDir.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdTextDir.Location = New System.Drawing.Point(614, 25)
    Me.cmdTextDir.Name = "cmdTextDir"
    Me.cmdTextDir.Size = New System.Drawing.Size(63, 21)
    Me.cmdTextDir.TabIndex = 5
    Me.cmdTextDir.Text = "..."
    Me.cmdTextDir.UseVisualStyleBackColor = True
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(12, 62)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(0, 13)
    Me.Label2.TabIndex = 6
    '
    'tbFeatCat
    '
    Me.tbFeatCat.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbFeatCat.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbFeatCat.Location = New System.Drawing.Point(61, 16)
    Me.tbFeatCat.Name = "tbFeatCat"
    Me.tbFeatCat.Size = New System.Drawing.Size(100, 20)
    Me.tbFeatCat.TabIndex = 7
    '
    'tbFeatname
    '
    Me.tbFeatname.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbFeatname.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbFeatname.Location = New System.Drawing.Point(61, 42)
    Me.tbFeatname.Name = "tbFeatname"
    Me.tbFeatname.Size = New System.Drawing.Size(100, 20)
    Me.tbFeatname.TabIndex = 8
    '
    'GroupBox1
    '
    Me.GroupBox1.Controls.Add(Me.Label4)
    Me.GroupBox1.Controls.Add(Me.Label3)
    Me.GroupBox1.Controls.Add(Me.tbFeatname)
    Me.GroupBox1.Controls.Add(Me.tbFeatCat)
    Me.GroupBox1.Location = New System.Drawing.Point(221, 231)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(200, 73)
    Me.GroupBox1.TabIndex = 9
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Properties of the feature in the texts"
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(6, 45)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(35, 13)
    Me.Label4.TabIndex = 9
    Me.Label4.Text = "Name"
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(6, 19)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(49, 13)
    Me.Label3.TabIndex = 0
    Me.Label3.Text = "Category"
    '
    'GroupBox2
    '
    Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GroupBox2.Controls.Add(Me.tbFeatBlackList)
    Me.GroupBox2.Controls.Add(Me.Label6)
    Me.GroupBox2.Controls.Add(Me.tbFeatWhiteList)
    Me.GroupBox2.Controls.Add(Me.rbRestrOnly)
    Me.GroupBox2.Controls.Add(Me.rbRestrNone)
    Me.GroupBox2.Location = New System.Drawing.Point(221, 62)
    Me.GroupBox2.Name = "GroupBox2"
    Me.GroupBox2.Size = New System.Drawing.Size(456, 134)
    Me.GroupBox2.TabIndex = 10
    Me.GroupBox2.TabStop = False
    Me.GroupBox2.Text = "Restrictions"
    '
    'tbFeatBlackList
    '
    Me.tbFeatBlackList.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFeatBlackList.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbFeatBlackList.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbFeatBlackList.Location = New System.Drawing.Point(24, 103)
    Me.tbFeatBlackList.Name = "tbFeatBlackList"
    Me.tbFeatBlackList.ReadOnly = True
    Me.tbFeatBlackList.Size = New System.Drawing.Size(420, 20)
    Me.tbFeatBlackList.TabIndex = 4
    '
    'Label6
    '
    Me.Label6.AutoSize = True
    Me.Label6.Location = New System.Drawing.Point(21, 87)
    Me.Label6.Name = "Label6"
    Me.Label6.Size = New System.Drawing.Size(207, 13)
    Me.Label6.TabIndex = 3
    Me.Label6.Text = "Exclude features with the following values:"
    '
    'tbFeatWhiteList
    '
    Me.tbFeatWhiteList.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFeatWhiteList.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbFeatWhiteList.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbFeatWhiteList.Location = New System.Drawing.Point(24, 64)
    Me.tbFeatWhiteList.Name = "tbFeatWhiteList"
    Me.tbFeatWhiteList.ReadOnly = True
    Me.tbFeatWhiteList.Size = New System.Drawing.Size(420, 20)
    Me.tbFeatWhiteList.TabIndex = 2
    '
    'rbRestrOnly
    '
    Me.rbRestrOnly.AutoSize = True
    Me.rbRestrOnly.Location = New System.Drawing.Point(6, 41)
    Me.rbRestrOnly.Name = "rbRestrOnly"
    Me.rbRestrOnly.Size = New System.Drawing.Size(234, 17)
    Me.rbRestrOnly.TabIndex = 1
    Me.rbRestrOnly.Text = "Only copy features with the following values:"
    Me.rbRestrOnly.UseVisualStyleBackColor = True
    '
    'rbRestrNone
    '
    Me.rbRestrNone.AutoSize = True
    Me.rbRestrNone.Checked = True
    Me.rbRestrNone.Location = New System.Drawing.Point(6, 19)
    Me.rbRestrNone.Name = "rbRestrNone"
    Me.rbRestrNone.Size = New System.Drawing.Size(92, 17)
    Me.rbRestrNone.TabIndex = 0
    Me.rbRestrNone.TabStop = True
    Me.rbRestrNone.Text = "No restrictions"
    Me.rbRestrNone.UseVisualStyleBackColor = True
    '
    'Label5
    '
    Me.Label5.AutoSize = True
    Me.Label5.Location = New System.Drawing.Point(12, 307)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(401, 13)
    Me.Label5.TabIndex = 11
    Me.Label5.Text = "Currently selected feature (changes should be made through Corpus/FilterFeatures)" & _
    ":"
    '
    'tbFeatFilter
    '
    Me.tbFeatFilter.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFeatFilter.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbFeatFilter.ForeColor = System.Drawing.Color.Maroon
    Me.tbFeatFilter.Location = New System.Drawing.Point(15, 323)
    Me.tbFeatFilter.Name = "tbFeatFilter"
    Me.tbFeatFilter.ReadOnly = True
    Me.tbFeatFilter.Size = New System.Drawing.Size(662, 122)
    Me.tbFeatFilter.TabIndex = 12
    Me.tbFeatFilter.Text = ""
    '
    'GroupBox3
    '
    Me.GroupBox3.Controls.Add(Me.lboFeatDbase)
    Me.GroupBox3.Location = New System.Drawing.Point(15, 62)
    Me.GroupBox3.Name = "GroupBox3"
    Me.GroupBox3.Size = New System.Drawing.Size(200, 242)
    Me.GroupBox3.TabIndex = 13
    Me.GroupBox3.TabStop = False
    Me.GroupBox3.Text = "Select the feature in the database"
    '
    'lboFeatDbase
    '
    Me.lboFeatDbase.FormattingEnabled = True
    Me.lboFeatDbase.Location = New System.Drawing.Point(9, 19)
    Me.lboFeatDbase.Name = "lboFeatDbase"
    Me.lboFeatDbase.Size = New System.Drawing.Size(185, 212)
    Me.lboFeatDbase.TabIndex = 0
    '
    'Label7
    '
    Me.Label7.AutoSize = True
    Me.Label7.Location = New System.Drawing.Point(227, 211)
    Me.Label7.Name = "Label7"
    Me.Label7.Size = New System.Drawing.Size(52, 13)
    Me.Label7.TabIndex = 14
    Me.Label7.Text = "Selected:"
    '
    'tbFeatDbase
    '
    Me.tbFeatDbase.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbFeatDbase.Location = New System.Drawing.Point(282, 208)
    Me.tbFeatDbase.Name = "tbFeatDbase"
    Me.tbFeatDbase.ReadOnly = True
    Me.tbFeatDbase.Size = New System.Drawing.Size(100, 20)
    Me.tbFeatDbase.TabIndex = 15
    '
    'frmFeatToText
    '
    Me.AcceptButton = Me.cmdOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(689, 499)
    Me.Controls.Add(Me.tbFeatDbase)
    Me.Controls.Add(Me.Label7)
    Me.Controls.Add(Me.GroupBox3)
    Me.Controls.Add(Me.tbFeatFilter)
    Me.Controls.Add(Me.Label5)
    Me.Controls.Add(Me.GroupBox2)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.cmdTextDir)
    Me.Controls.Add(Me.tbTextDir)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.cmdOk)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmFeatToText"
    Me.Text = "Transfer features from database to texts"
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.GroupBox1.ResumeLayout(False)
    Me.GroupBox1.PerformLayout()
    Me.GroupBox2.ResumeLayout(False)
    Me.GroupBox2.PerformLayout()
    Me.GroupBox3.ResumeLayout(False)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents tbTextDir As System.Windows.Forms.TextBox
  Friend WithEvents cmdTextDir As System.Windows.Forms.Button
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents tbFeatCat As System.Windows.Forms.TextBox
  Friend WithEvents tbFeatname As System.Windows.Forms.TextBox
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
  Friend WithEvents tbFeatWhiteList As System.Windows.Forms.TextBox
  Friend WithEvents rbRestrOnly As System.Windows.Forms.RadioButton
  Friend WithEvents rbRestrNone As System.Windows.Forms.RadioButton
  Friend WithEvents Label5 As System.Windows.Forms.Label
  Friend WithEvents tbFeatFilter As System.Windows.Forms.RichTextBox
  Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
  Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
  Friend WithEvents lboFeatDbase As System.Windows.Forms.ListBox
  Friend WithEvents tbFeatBlackList As System.Windows.Forms.TextBox
  Friend WithEvents Label6 As System.Windows.Forms.Label
  Friend WithEvents Label7 As System.Windows.Forms.Label
  Friend WithEvents tbFeatDbase As System.Windows.Forms.TextBox
End Class
