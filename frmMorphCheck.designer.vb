<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMorphCheck
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMorphCheck))
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.cmdOk = New System.Windows.Forms.Button
    Me.Label1 = New System.Windows.Forms.Label
    Me.wbSentence = New System.Windows.Forms.WebBrowser
    Me.Label3 = New System.Windows.Forms.Label
    Me.tbPos = New System.Windows.Forms.TextBox
    Me.Label4 = New System.Windows.Forms.Label
    Me.tbFeat = New System.Windows.Forms.RichTextBox
    Me.GroupBox1 = New System.Windows.Forms.GroupBox
    Me.rbDeleteLemmaVern = New System.Windows.Forms.RadioButton
    Me.rbDeleteLemmaVernPOS = New System.Windows.Forms.RadioButton
    Me.rbChangeToLemmaDictPOS = New System.Windows.Forms.RadioButton
    Me.rbAllowPOS = New System.Windows.Forms.RadioButton
    Me.tbLemma = New System.Windows.Forms.TextBox
    Me.Label2 = New System.Windows.Forms.Label
    Me.Label5 = New System.Windows.Forms.Label
    Me.tbMWcat = New System.Windows.Forms.TextBox
    Me.tbCheck = New System.Windows.Forms.TextBox
    Me.Label6 = New System.Windows.Forms.Label
    Me.lboMorphDict = New System.Windows.Forms.ListBox
    Me.StatusStrip1.SuspendLayout()
    Me.GroupBox1.SuspendLayout()
    Me.SuspendLayout()
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 454)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(720, 22)
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
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(633, 428)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 1
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'cmdOk
    '
    Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOk.Location = New System.Drawing.Point(552, 428)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(75, 23)
    Me.cmdOk.TabIndex = 2
    Me.cmdOk.Text = "&Ok"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(12, 9)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(56, 13)
    Me.Label1.TabIndex = 3
    Me.Label1.Text = "Sentence:"
    '
    'wbSentence
    '
    Me.wbSentence.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.wbSentence.Location = New System.Drawing.Point(15, 26)
    Me.wbSentence.MinimumSize = New System.Drawing.Size(20, 20)
    Me.wbSentence.Name = "wbSentence"
    Me.wbSentence.Size = New System.Drawing.Size(693, 138)
    Me.wbSentence.TabIndex = 4
    '
    'Label3
    '
    Me.Label3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(347, 191)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(75, 13)
    Me.Label3.TabIndex = 7
    Me.Label3.Text = "POS in YCOE:"
    '
    'tbPos
    '
    Me.tbPos.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbPos.Location = New System.Drawing.Point(428, 188)
    Me.tbPos.Name = "tbPos"
    Me.tbPos.Size = New System.Drawing.Size(118, 20)
    Me.tbPos.TabIndex = 8
    '
    'Label4
    '
    Me.Label4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(552, 170)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(48, 13)
    Me.Label4.TabIndex = 9
    Me.Label4.Text = "Features"
    '
    'tbFeat
    '
    Me.tbFeat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFeat.Location = New System.Drawing.Point(555, 188)
    Me.tbFeat.Name = "tbFeat"
    Me.tbFeat.Size = New System.Drawing.Size(153, 94)
    Me.tbFeat.TabIndex = 10
    Me.tbFeat.Text = ""
    '
    'GroupBox1
    '
    Me.GroupBox1.Controls.Add(Me.rbDeleteLemmaVern)
    Me.GroupBox1.Controls.Add(Me.rbDeleteLemmaVernPOS)
    Me.GroupBox1.Controls.Add(Me.rbChangeToLemmaDictPOS)
    Me.GroupBox1.Controls.Add(Me.rbAllowPOS)
    Me.GroupBox1.Location = New System.Drawing.Point(15, 170)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(281, 131)
    Me.GroupBox1.TabIndex = 11
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Choose action to take"
    '
    'rbDeleteLemmaVern
    '
    Me.rbDeleteLemmaVern.AutoSize = True
    Me.rbDeleteLemmaVern.Location = New System.Drawing.Point(6, 88)
    Me.rbDeleteLemmaVern.Name = "rbDeleteLemmaVern"
    Me.rbDeleteLemmaVern.Size = New System.Drawing.Size(214, 17)
    Me.rbDeleteLemmaVern.TabIndex = 3
    Me.rbDeleteLemmaVern.TabStop = True
    Me.rbDeleteLemmaVern.Text = "Delete Lemma's for occurrences of &Vern"
    Me.rbDeleteLemmaVern.UseVisualStyleBackColor = True
    '
    'rbDeleteLemmaVernPOS
    '
    Me.rbDeleteLemmaVernPOS.AutoSize = True
    Me.rbDeleteLemmaVernPOS.Location = New System.Drawing.Point(6, 65)
    Me.rbDeleteLemmaVernPOS.Name = "rbDeleteLemmaVernPOS"
    Me.rbDeleteLemmaVernPOS.Size = New System.Drawing.Size(244, 17)
    Me.rbDeleteLemmaVernPOS.TabIndex = 2
    Me.rbDeleteLemmaVernPOS.TabStop = True
    Me.rbDeleteLemmaVernPOS.Text = "Delete &Lemma's for combinations of Vern/POS"
    Me.rbDeleteLemmaVernPOS.UseVisualStyleBackColor = True
    '
    'rbChangeToLemmaDictPOS
    '
    Me.rbChangeToLemmaDictPOS.AutoSize = True
    Me.rbChangeToLemmaDictPOS.Location = New System.Drawing.Point(6, 42)
    Me.rbChangeToLemmaDictPOS.Name = "rbChangeToLemmaDictPOS"
    Me.rbChangeToLemmaDictPOS.Size = New System.Drawing.Size(260, 17)
    Me.rbChangeToLemmaDictPOS.TabIndex = 1
    Me.rbChangeToLemmaDictPOS.TabStop = True
    Me.rbChangeToLemmaDictPOS.Text = "&Change POS of Vern/Lemma --> Lemma dict POS"
    Me.rbChangeToLemmaDictPOS.UseVisualStyleBackColor = True
    Me.rbChangeToLemmaDictPOS.Visible = False
    '
    'rbAllowPOS
    '
    Me.rbAllowPOS.AutoSize = True
    Me.rbAllowPOS.Location = New System.Drawing.Point(6, 19)
    Me.rbAllowPOS.Name = "rbAllowPOS"
    Me.rbAllowPOS.Size = New System.Drawing.Size(213, 17)
    Me.rbAllowPOS.TabIndex = 0
    Me.rbAllowPOS.TabStop = True
    Me.rbAllowPOS.Text = "&Allow combination of Vern/Lemma/POS"
    Me.rbAllowPOS.UseVisualStyleBackColor = True
    '
    'tbLemma
    '
    Me.tbLemma.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbLemma.Location = New System.Drawing.Point(428, 214)
    Me.tbLemma.Name = "tbLemma"
    Me.tbLemma.Size = New System.Drawing.Size(118, 20)
    Me.tbLemma.TabIndex = 13
    '
    'Label2
    '
    Me.Label2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(302, 217)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(121, 13)
    Me.Label2.TabIndex = 12
    Me.Label2.Text = "Lemma from OE-tagged:"
    '
    'Label5
    '
    Me.Label5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Label5.AutoSize = True
    Me.Label5.Location = New System.Drawing.Point(312, 243)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(110, 13)
    Me.Label5.TabIndex = 14
    Me.Label5.Text = "Lemma dictionary cat:"
    '
    'tbMWcat
    '
    Me.tbMWcat.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbMWcat.Location = New System.Drawing.Point(428, 240)
    Me.tbMWcat.Name = "tbMWcat"
    Me.tbMWcat.Size = New System.Drawing.Size(118, 20)
    Me.tbMWcat.TabIndex = 15
    '
    'tbCheck
    '
    Me.tbCheck.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbCheck.Location = New System.Drawing.Point(428, 266)
    Me.tbCheck.Name = "tbCheck"
    Me.tbCheck.Size = New System.Drawing.Size(118, 20)
    Me.tbCheck.TabIndex = 17
    '
    'Label6
    '
    Me.Label6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Label6.AutoSize = True
    Me.Label6.Location = New System.Drawing.Point(305, 269)
    Me.Label6.Name = "Label6"
    Me.Label6.Size = New System.Drawing.Size(117, 13)
    Me.Label6.TabIndex = 16
    Me.Label6.Text = "Lemma dictionary POS:"
    '
    'lboMorphDict
    '
    Me.lboMorphDict.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.lboMorphDict.FormattingEnabled = True
    Me.lboMorphDict.Location = New System.Drawing.Point(15, 307)
    Me.lboMorphDict.Name = "lboMorphDict"
    Me.lboMorphDict.Size = New System.Drawing.Size(693, 108)
    Me.lboMorphDict.TabIndex = 18
    '
    'frmMorphCheck
    '
    Me.AcceptButton = Me.cmdOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(720, 476)
    Me.Controls.Add(Me.lboMorphDict)
    Me.Controls.Add(Me.tbCheck)
    Me.Controls.Add(Me.Label6)
    Me.Controls.Add(Me.tbMWcat)
    Me.Controls.Add(Me.Label5)
    Me.Controls.Add(Me.tbLemma)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.tbFeat)
    Me.Controls.Add(Me.Label4)
    Me.Controls.Add(Me.tbPos)
    Me.Controls.Add(Me.Label3)
    Me.Controls.Add(Me.wbSentence)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.cmdOk)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmMorphCheck"
    Me.Text = "Choose the correct morphology features for this word"
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.GroupBox1.ResumeLayout(False)
    Me.GroupBox1.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents wbSentence As System.Windows.Forms.WebBrowser
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents tbPos As System.Windows.Forms.TextBox
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents tbFeat As System.Windows.Forms.RichTextBox
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents rbDeleteLemmaVern As System.Windows.Forms.RadioButton
  Friend WithEvents rbDeleteLemmaVernPOS As System.Windows.Forms.RadioButton
  Friend WithEvents rbChangeToLemmaDictPOS As System.Windows.Forms.RadioButton
  Friend WithEvents rbAllowPOS As System.Windows.Forms.RadioButton
  Friend WithEvents tbLemma As System.Windows.Forms.TextBox
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents Label5 As System.Windows.Forms.Label
  Friend WithEvents tbMWcat As System.Windows.Forms.TextBox
  Friend WithEvents tbCheck As System.Windows.Forms.TextBox
  Friend WithEvents Label6 As System.Windows.Forms.Label
  Friend WithEvents lboMorphDict As System.Windows.Forms.ListBox
End Class
