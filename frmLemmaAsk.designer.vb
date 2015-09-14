<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLemmaAsk
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLemmaAsk))
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar
    Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
    Me.Label4 = New System.Windows.Forms.Label
    Me.lboMorphLemma = New System.Windows.Forms.ListBox
    Me.tbLemma = New System.Windows.Forms.TextBox
    Me.Label6 = New System.Windows.Forms.Label
    Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
    Me.dgvOEdict = New System.Windows.Forms.DataGridView
    Me.tbOEdictKey = New System.Windows.Forms.TextBox
    Me.tbOElemma = New System.Windows.Forms.TextBox
    Me.Label9 = New System.Windows.Forms.Label
    Me.Label8 = New System.Windows.Forms.Label
    Me.tbOEpos = New System.Windows.Forms.TextBox
    Me.tbOEfeat = New System.Windows.Forms.TextBox
    Me.Label7 = New System.Windows.Forms.Label
    Me.tbOEdictEntry = New System.Windows.Forms.RichTextBox
    Me.Label5 = New System.Windows.Forms.Label
    Me.tbWord = New System.Windows.Forms.TextBox
    Me.Label1 = New System.Windows.Forms.Label
    Me.tbPos = New System.Windows.Forms.TextBox
    Me.Label2 = New System.Windows.Forms.Label
    Me.Label3 = New System.Windows.Forms.Label
    Me.wbContext = New System.Windows.Forms.WebBrowser
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.cmdOk = New System.Windows.Forms.Button
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
    Me.cmdSaveStop = New System.Windows.Forms.Button
    Me.cmdSkip = New System.Windows.Forms.Button
    Me.StatusStrip1.SuspendLayout()
    Me.SplitContainer1.Panel1.SuspendLayout()
    Me.SplitContainer1.Panel2.SuspendLayout()
    Me.SplitContainer1.SuspendLayout()
    Me.SplitContainer2.Panel1.SuspendLayout()
    Me.SplitContainer2.Panel2.SuspendLayout()
    Me.SplitContainer2.SuspendLayout()
    CType(Me.dgvOEdict, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 531)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(859, 22)
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
    'SplitContainer1
    '
    Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.SplitContainer1.Location = New System.Drawing.Point(0, 134)
    Me.SplitContainer1.Name = "SplitContainer1"
    '
    'SplitContainer1.Panel1
    '
    Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
    Me.SplitContainer1.Panel1.Controls.Add(Me.lboMorphLemma)
    '
    'SplitContainer1.Panel2
    '
    Me.SplitContainer1.Panel2.Controls.Add(Me.tbLemma)
    Me.SplitContainer1.Panel2.Controls.Add(Me.Label6)
    Me.SplitContainer1.Panel2.Controls.Add(Me.SplitContainer2)
    Me.SplitContainer1.Size = New System.Drawing.Size(859, 394)
    Me.SplitContainer1.SplitterDistance = 282
    Me.SplitContainer1.TabIndex = 1
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(3, 17)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(141, 13)
    Me.Label4.TabIndex = 1
    Me.Label4.Text = "Is your lemma among these?"
    '
    'lboMorphLemma
    '
    Me.lboMorphLemma.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.lboMorphLemma.FormattingEnabled = True
    Me.lboMorphLemma.Location = New System.Drawing.Point(3, 33)
    Me.lboMorphLemma.Name = "lboMorphLemma"
    Me.lboMorphLemma.Size = New System.Drawing.Size(271, 342)
    Me.lboMorphLemma.TabIndex = 0
    '
    'tbLemma
    '
    Me.tbLemma.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbLemma.Location = New System.Drawing.Point(256, 14)
    Me.tbLemma.Name = "tbLemma"
    Me.tbLemma.Size = New System.Drawing.Size(303, 20)
    Me.tbLemma.TabIndex = 4
    '
    'Label6
    '
    Me.Label6.AutoSize = True
    Me.Label6.Location = New System.Drawing.Point(3, 17)
    Me.Label6.Name = "Label6"
    Me.Label6.Size = New System.Drawing.Size(247, 13)
    Me.Label6.TabIndex = 3
    Me.Label6.Text = "The lemma that should be recorded for this word is:"
    '
    'SplitContainer2
    '
    Me.SplitContainer2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.SplitContainer2.Location = New System.Drawing.Point(-1, 55)
    Me.SplitContainer2.Name = "SplitContainer2"
    '
    'SplitContainer2.Panel1
    '
    Me.SplitContainer2.Panel1.Controls.Add(Me.dgvOEdict)
    Me.SplitContainer2.Panel1.Controls.Add(Me.tbOEdictKey)
    '
    'SplitContainer2.Panel2
    '
    Me.SplitContainer2.Panel2.Controls.Add(Me.tbOElemma)
    Me.SplitContainer2.Panel2.Controls.Add(Me.Label9)
    Me.SplitContainer2.Panel2.Controls.Add(Me.Label8)
    Me.SplitContainer2.Panel2.Controls.Add(Me.tbOEpos)
    Me.SplitContainer2.Panel2.Controls.Add(Me.tbOEfeat)
    Me.SplitContainer2.Panel2.Controls.Add(Me.Label7)
    Me.SplitContainer2.Panel2.Controls.Add(Me.tbOEdictEntry)
    Me.SplitContainer2.Panel2.Controls.Add(Me.Label5)
    Me.SplitContainer2.Size = New System.Drawing.Size(563, 333)
    Me.SplitContainer2.SplitterDistance = 183
    Me.SplitContainer2.TabIndex = 2
    '
    'dgvOEdict
    '
    Me.dgvOEdict.AllowUserToAddRows = False
    Me.dgvOEdict.AllowUserToDeleteRows = False
    Me.dgvOEdict.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dgvOEdict.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
    Me.dgvOEdict.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
    Me.dgvOEdict.Location = New System.Drawing.Point(0, 32)
    Me.dgvOEdict.Name = "dgvOEdict"
    Me.dgvOEdict.ReadOnly = True
    Me.dgvOEdict.Size = New System.Drawing.Size(180, 298)
    Me.dgvOEdict.TabIndex = 1
    '
    'tbOEdictKey
    '
    Me.tbOEdictKey.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbOEdictKey.Location = New System.Drawing.Point(0, 6)
    Me.tbOEdictKey.Name = "tbOEdictKey"
    Me.tbOEdictKey.Size = New System.Drawing.Size(180, 20)
    Me.tbOEdictKey.TabIndex = 0
    '
    'tbOElemma
    '
    Me.tbOElemma.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbOElemma.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbOElemma.Location = New System.Drawing.Point(92, 6)
    Me.tbOElemma.Name = "tbOElemma"
    Me.tbOElemma.Size = New System.Drawing.Size(281, 20)
    Me.tbOElemma.TabIndex = 9
    '
    'Label9
    '
    Me.Label9.AutoSize = True
    Me.Label9.Location = New System.Drawing.Point(42, 9)
    Me.Label9.Name = "Label9"
    Me.Label9.Size = New System.Drawing.Size(44, 13)
    Me.Label9.TabIndex = 8
    Me.Label9.Text = "Lemma:"
    '
    'Label8
    '
    Me.Label8.AutoSize = True
    Me.Label8.Location = New System.Drawing.Point(5, 61)
    Me.Label8.Name = "Label8"
    Me.Label8.Size = New System.Drawing.Size(81, 13)
    Me.Label8.TabIndex = 7
    Me.Label8.Text = "Part-of-Speech:"
    '
    'tbOEpos
    '
    Me.tbOEpos.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbOEpos.Location = New System.Drawing.Point(92, 58)
    Me.tbOEpos.Name = "tbOEpos"
    Me.tbOEpos.Size = New System.Drawing.Size(75, 20)
    Me.tbOEpos.TabIndex = 6
    '
    'tbOEfeat
    '
    Me.tbOEfeat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbOEfeat.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbOEfeat.Location = New System.Drawing.Point(92, 32)
    Me.tbOEfeat.Name = "tbOEfeat"
    Me.tbOEfeat.Size = New System.Drawing.Size(281, 20)
    Me.tbOEfeat.TabIndex = 3
    '
    'Label7
    '
    Me.Label7.AutoSize = True
    Me.Label7.Location = New System.Drawing.Point(5, 84)
    Me.Label7.Name = "Label7"
    Me.Label7.Size = New System.Drawing.Size(65, 13)
    Me.Label7.TabIndex = 2
    Me.Label7.Text = "Definition(s):"
    '
    'tbOEdictEntry
    '
    Me.tbOEdictEntry.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbOEdictEntry.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbOEdictEntry.Location = New System.Drawing.Point(3, 100)
    Me.tbOEdictEntry.Name = "tbOEdictEntry"
    Me.tbOEdictEntry.Size = New System.Drawing.Size(370, 230)
    Me.tbOEdictEntry.TabIndex = 0
    Me.tbOEdictEntry.Text = ""
    '
    'Label5
    '
    Me.Label5.AutoSize = True
    Me.Label5.Location = New System.Drawing.Point(29, 35)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(57, 13)
    Me.Label5.TabIndex = 1
    Me.Label5.Text = "Feature(s):"
    '
    'tbWord
    '
    Me.tbWord.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbWord.Location = New System.Drawing.Point(99, 12)
    Me.tbWord.Name = "tbWord"
    Me.tbWord.Size = New System.Drawing.Size(257, 20)
    Me.tbWord.TabIndex = 2
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(12, 15)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(36, 13)
    Me.Label1.TabIndex = 3
    Me.Label1.Text = "Word:"
    '
    'tbPos
    '
    Me.tbPos.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbPos.Location = New System.Drawing.Point(99, 38)
    Me.tbPos.Name = "tbPos"
    Me.tbPos.Size = New System.Drawing.Size(75, 20)
    Me.tbPos.TabIndex = 4
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(12, 41)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(81, 13)
    Me.Label2.TabIndex = 5
    Me.Label2.Text = "Part-of-Speech:"
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(193, 41)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(46, 13)
    Me.Label3.TabIndex = 6
    Me.Label3.Text = "Context:"
    '
    'wbContext
    '
    Me.wbContext.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.wbContext.Location = New System.Drawing.Point(245, 38)
    Me.wbContext.MinimumSize = New System.Drawing.Size(20, 20)
    Me.wbContext.Name = "wbContext"
    Me.wbContext.Size = New System.Drawing.Size(600, 90)
    Me.wbContext.TabIndex = 7
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(770, 9)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 8
    Me.cmdCancel.Text = "&Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'cmdOk
    '
    Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOk.Location = New System.Drawing.Point(689, 10)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(75, 23)
    Me.cmdOk.TabIndex = 9
    Me.cmdOk.Text = "&Ok"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'Timer1
    '
    '
    'Timer2
    '
    Me.Timer2.Interval = 500
    '
    'cmdSaveStop
    '
    Me.cmdSaveStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdSaveStop.Location = New System.Drawing.Point(574, 9)
    Me.cmdSaveStop.Name = "cmdSaveStop"
    Me.cmdSaveStop.Size = New System.Drawing.Size(109, 23)
    Me.cmdSaveStop.TabIndex = 10
    Me.cmdSaveStop.Text = "&Save and stop"
    Me.cmdSaveStop.UseVisualStyleBackColor = True
    '
    'cmdSkip
    '
    Me.cmdSkip.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdSkip.Location = New System.Drawing.Point(459, 9)
    Me.cmdSkip.Name = "cmdSkip"
    Me.cmdSkip.Size = New System.Drawing.Size(109, 23)
    Me.cmdSkip.TabIndex = 11
    Me.cmdSkip.Text = "S&kip this one"
    Me.cmdSkip.UseVisualStyleBackColor = True
    '
    'frmLemmaAsk
    '
    Me.AcceptButton = Me.cmdOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(859, 553)
    Me.Controls.Add(Me.cmdSkip)
    Me.Controls.Add(Me.cmdSaveStop)
    Me.Controls.Add(Me.cmdOk)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.wbContext)
    Me.Controls.Add(Me.Label3)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.tbPos)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.tbWord)
    Me.Controls.Add(Me.SplitContainer1)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmLemmaAsk"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
    Me.Text = "Please provide the lemma for the word that lacks it"
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.SplitContainer1.Panel1.ResumeLayout(False)
    Me.SplitContainer1.Panel1.PerformLayout()
    Me.SplitContainer1.Panel2.ResumeLayout(False)
    Me.SplitContainer1.Panel2.PerformLayout()
    Me.SplitContainer1.ResumeLayout(False)
    Me.SplitContainer2.Panel1.ResumeLayout(False)
    Me.SplitContainer2.Panel1.PerformLayout()
    Me.SplitContainer2.Panel2.ResumeLayout(False)
    Me.SplitContainer2.Panel2.PerformLayout()
    Me.SplitContainer2.ResumeLayout(False)
    CType(Me.dgvOEdict, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
  Friend WithEvents tbWord As System.Windows.Forms.TextBox
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents tbPos As System.Windows.Forms.TextBox
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents wbContext As System.Windows.Forms.WebBrowser
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents lboMorphLemma As System.Windows.Forms.ListBox
  Friend WithEvents Label5 As System.Windows.Forms.Label
  Friend WithEvents tbOEdictEntry As System.Windows.Forms.RichTextBox
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents tbLemma As System.Windows.Forms.TextBox
  Friend WithEvents Label6 As System.Windows.Forms.Label
  Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
  Friend WithEvents dgvOEdict As System.Windows.Forms.DataGridView
  Friend WithEvents tbOEdictKey As System.Windows.Forms.TextBox
  Friend WithEvents Label7 As System.Windows.Forms.Label
  Friend WithEvents Label8 As System.Windows.Forms.Label
  Friend WithEvents tbOEpos As System.Windows.Forms.TextBox
  Friend WithEvents tbOEfeat As System.Windows.Forms.TextBox
  Friend WithEvents Timer2 As System.Windows.Forms.Timer
  Friend WithEvents tbOElemma As System.Windows.Forms.TextBox
  Friend WithEvents Label9 As System.Windows.Forms.Label
  Friend WithEvents cmdSaveStop As System.Windows.Forms.Button
  Friend WithEvents cmdSkip As System.Windows.Forms.Button
End Class
