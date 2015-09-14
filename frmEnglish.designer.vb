<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEnglish
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEnglish))
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
    Me.GroupBox1 = New System.Windows.Forms.GroupBox()
    Me.rbMiddleEnglishMED = New System.Windows.Forms.RadioButton()
    Me.rbLateModEng = New System.Windows.Forms.RadioButton()
    Me.rbEarlyModEng = New System.Windows.Forms.RadioButton()
    Me.rbMiddleEnglish = New System.Windows.Forms.RadioButton()
    Me.rbOldEnglish = New System.Windows.Forms.RadioButton()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.cmdOk = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.rbOldEnglishBT = New System.Windows.Forms.RadioButton()
    Me.StatusStrip1.SuspendLayout()
    Me.GroupBox1.SuspendLayout()
    Me.SuspendLayout()
    '
    'Timer1
    '
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 229)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(341, 22)
    Me.StatusStrip1.TabIndex = 0
    Me.StatusStrip1.Text = "StatusStrip1"
    '
    'ToolStripStatusLabel1
    '
    Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
    Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(19, 17)
    Me.ToolStripStatusLabel1.Text = "..."
    '
    'GroupBox1
    '
    Me.GroupBox1.Controls.Add(Me.rbOldEnglishBT)
    Me.GroupBox1.Controls.Add(Me.rbMiddleEnglishMED)
    Me.GroupBox1.Controls.Add(Me.rbLateModEng)
    Me.GroupBox1.Controls.Add(Me.rbEarlyModEng)
    Me.GroupBox1.Controls.Add(Me.rbMiddleEnglish)
    Me.GroupBox1.Controls.Add(Me.rbOldEnglish)
    Me.GroupBox1.Location = New System.Drawing.Point(23, 43)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(215, 171)
    Me.GroupBox1.TabIndex = 1
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Language"
    '
    'rbMiddleEnglishMED
    '
    Me.rbMiddleEnglishMED.AutoSize = True
    Me.rbMiddleEnglishMED.Location = New System.Drawing.Point(6, 88)
    Me.rbMiddleEnglishMED.Name = "rbMiddleEnglishMED"
    Me.rbMiddleEnglishMED.Size = New System.Drawing.Size(126, 17)
    Me.rbMiddleEnglishMED.TabIndex = 4
    Me.rbMiddleEnglishMED.TabStop = True
    Me.rbMiddleEnglishMED.Text = "Middle English - ME&D"
    Me.rbMiddleEnglishMED.UseVisualStyleBackColor = True
    '
    'rbLateModEng
    '
    Me.rbLateModEng.AutoSize = True
    Me.rbLateModEng.Location = New System.Drawing.Point(6, 134)
    Me.rbLateModEng.Name = "rbLateModEng"
    Me.rbLateModEng.Size = New System.Drawing.Size(164, 17)
    Me.rbLateModEng.TabIndex = 3
    Me.rbLateModEng.TabStop = True
    Me.rbLateModEng.Text = "&Late Modern English (LmodE)"
    Me.rbLateModEng.UseVisualStyleBackColor = True
    '
    'rbEarlyModEng
    '
    Me.rbEarlyModEng.AutoSize = True
    Me.rbEarlyModEng.Location = New System.Drawing.Point(6, 111)
    Me.rbEarlyModEng.Name = "rbEarlyModEng"
    Me.rbEarlyModEng.Size = New System.Drawing.Size(167, 17)
    Me.rbEarlyModEng.TabIndex = 2
    Me.rbEarlyModEng.TabStop = True
    Me.rbEarlyModEng.Text = "&Early Modern English (eModE)"
    Me.rbEarlyModEng.UseVisualStyleBackColor = True
    '
    'rbMiddleEnglish
    '
    Me.rbMiddleEnglish.AutoSize = True
    Me.rbMiddleEnglish.Location = New System.Drawing.Point(6, 65)
    Me.rbMiddleEnglish.Name = "rbMiddleEnglish"
    Me.rbMiddleEnglish.Size = New System.Drawing.Size(118, 17)
    Me.rbMiddleEnglish.TabIndex = 1
    Me.rbMiddleEnglish.TabStop = True
    Me.rbMiddleEnglish.Text = "&Middle English (ME)"
    Me.rbMiddleEnglish.UseVisualStyleBackColor = True
    '
    'rbOldEnglish
    '
    Me.rbOldEnglish.AutoSize = True
    Me.rbOldEnglish.Location = New System.Drawing.Point(6, 19)
    Me.rbOldEnglish.Name = "rbOldEnglish"
    Me.rbOldEnglish.Size = New System.Drawing.Size(102, 17)
    Me.rbOldEnglish.TabIndex = 0
    Me.rbOldEnglish.TabStop = True
    Me.rbOldEnglish.Text = "&Old English (OE)"
    Me.rbOldEnglish.UseVisualStyleBackColor = True
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.ForeColor = System.Drawing.Color.Maroon
    Me.Label1.Location = New System.Drawing.Point(20, 14)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(296, 13)
    Me.Label1.TabIndex = 2
    Me.Label1.Text = "Please indicate the English period you would like to work with"
    '
    'cmdOk
    '
    Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOk.Location = New System.Drawing.Point(254, 156)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(75, 23)
    Me.cmdOk.TabIndex = 4
    Me.cmdOk.Text = "&Ok"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(254, 191)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 5
    Me.cmdCancel.Text = "E&xit"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'rbOldEnglishBT
    '
    Me.rbOldEnglishBT.AutoSize = True
    Me.rbOldEnglishBT.Location = New System.Drawing.Point(6, 42)
    Me.rbOldEnglishBT.Name = "rbOldEnglishBT"
    Me.rbOldEnglishBT.Size = New System.Drawing.Size(192, 17)
    Me.rbOldEnglishBT.TabIndex = 5
    Me.rbOldEnglishBT.TabStop = True
    Me.rbOldEnglishBT.Text = "Old English &Bossworth/Toller (OEB)"
    Me.rbOldEnglishBT.UseVisualStyleBackColor = True
    '
    'frmEnglish
    '
    Me.AcceptButton = Me.cmdOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(341, 251)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.cmdOk)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.StatusStrip1)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmEnglish"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "English language period selection"
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.GroupBox1.ResumeLayout(False)
    Me.GroupBox1.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents rbLateModEng As System.Windows.Forms.RadioButton
  Friend WithEvents rbEarlyModEng As System.Windows.Forms.RadioButton
  Friend WithEvents rbMiddleEnglish As System.Windows.Forms.RadioButton
  Friend WithEvents rbOldEnglish As System.Windows.Forms.RadioButton
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents rbMiddleEnglishMED As System.Windows.Forms.RadioButton
  Friend WithEvents rbOldEnglishBT As System.Windows.Forms.RadioButton
End Class
