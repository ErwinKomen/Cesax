<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConvert
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmConvert))
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.cmdOk = New System.Windows.Forms.Button()
    Me.GroupBox1 = New System.Windows.Forms.GroupBox()
    Me.rbOther = New System.Windows.Forms.RadioButton()
    Me.tbEthno = New System.Windows.Forms.TextBox()
    Me.rbLak = New System.Windows.Forms.RadioButton()
    Me.rbLezgi = New System.Windows.Forms.RadioButton()
    Me.rbWelsh = New System.Windows.Forms.RadioButton()
    Me.rbFrench = New System.Windows.Forms.RadioButton()
    Me.rbSpanish = New System.Windows.Forms.RadioButton()
    Me.rbDutch = New System.Windows.Forms.RadioButton()
    Me.rbGerman = New System.Windows.Forms.RadioButton()
    Me.rbEnglishTranscribed = New System.Windows.Forms.RadioButton()
    Me.rbEnglishWritten = New System.Windows.Forms.RadioButton()
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.rbChechen = New System.Windows.Forms.RadioButton()
    Me.GroupBox1.SuspendLayout()
    Me.StatusStrip1.SuspendLayout()
    Me.SuspendLayout()
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(278, 323)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 9
    Me.cmdCancel.Text = "E&xit"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'cmdOk
    '
    Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOk.Location = New System.Drawing.Point(278, 288)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(75, 23)
    Me.cmdOk.TabIndex = 8
    Me.cmdOk.Text = "&Ok"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'GroupBox1
    '
    Me.GroupBox1.Controls.Add(Me.rbChechen)
    Me.GroupBox1.Controls.Add(Me.rbOther)
    Me.GroupBox1.Controls.Add(Me.tbEthno)
    Me.GroupBox1.Controls.Add(Me.rbLak)
    Me.GroupBox1.Controls.Add(Me.rbLezgi)
    Me.GroupBox1.Controls.Add(Me.rbWelsh)
    Me.GroupBox1.Controls.Add(Me.rbFrench)
    Me.GroupBox1.Controls.Add(Me.rbSpanish)
    Me.GroupBox1.Controls.Add(Me.rbDutch)
    Me.GroupBox1.Controls.Add(Me.rbGerman)
    Me.GroupBox1.Controls.Add(Me.rbEnglishTranscribed)
    Me.GroupBox1.Controls.Add(Me.rbEnglishWritten)
    Me.GroupBox1.Location = New System.Drawing.Point(23, 21)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(234, 325)
    Me.GroupBox1.TabIndex = 7
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Language"
    '
    'rbOther
    '
    Me.rbOther.AutoSize = True
    Me.rbOther.Location = New System.Drawing.Point(6, 273)
    Me.rbOther.Name = "rbOther"
    Me.rbOther.Size = New System.Drawing.Size(225, 17)
    Me.rbOther.TabIndex = 11
    Me.rbOther.TabStop = True
    Me.rbOther.Text = "Other language (3-letter ethnologue code):"
    Me.rbOther.UseVisualStyleBackColor = True
    '
    'tbEthno
    '
    Me.tbEthno.Location = New System.Drawing.Point(25, 296)
    Me.tbEthno.Name = "tbEthno"
    Me.tbEthno.Size = New System.Drawing.Size(55, 20)
    Me.tbEthno.TabIndex = 10
    '
    'rbLak
    '
    Me.rbLak.AutoSize = True
    Me.rbLak.Location = New System.Drawing.Point(6, 205)
    Me.rbLak.Name = "rbLak"
    Me.rbLak.Size = New System.Drawing.Size(43, 17)
    Me.rbLak.TabIndex = 8
    Me.rbLak.TabStop = True
    Me.rbLak.Text = "La&k"
    Me.rbLak.UseVisualStyleBackColor = True
    '
    'rbLezgi
    '
    Me.rbLezgi.AutoSize = True
    Me.rbLezgi.Location = New System.Drawing.Point(6, 182)
    Me.rbLezgi.Name = "rbLezgi"
    Me.rbLezgi.Size = New System.Drawing.Size(50, 17)
    Me.rbLezgi.TabIndex = 7
    Me.rbLezgi.TabStop = True
    Me.rbLezgi.Text = "&Lezgi"
    Me.rbLezgi.UseVisualStyleBackColor = True
    '
    'rbWelsh
    '
    Me.rbWelsh.AutoSize = True
    Me.rbWelsh.Location = New System.Drawing.Point(6, 159)
    Me.rbWelsh.Name = "rbWelsh"
    Me.rbWelsh.Size = New System.Drawing.Size(55, 17)
    Me.rbWelsh.TabIndex = 6
    Me.rbWelsh.TabStop = True
    Me.rbWelsh.Text = "&Welsh"
    Me.rbWelsh.UseVisualStyleBackColor = True
    '
    'rbFrench
    '
    Me.rbFrench.AutoSize = True
    Me.rbFrench.Location = New System.Drawing.Point(6, 136)
    Me.rbFrench.Name = "rbFrench"
    Me.rbFrench.Size = New System.Drawing.Size(58, 17)
    Me.rbFrench.TabIndex = 5
    Me.rbFrench.TabStop = True
    Me.rbFrench.Text = "&French"
    Me.rbFrench.UseVisualStyleBackColor = True
    '
    'rbSpanish
    '
    Me.rbSpanish.AutoSize = True
    Me.rbSpanish.Location = New System.Drawing.Point(6, 113)
    Me.rbSpanish.Name = "rbSpanish"
    Me.rbSpanish.Size = New System.Drawing.Size(63, 17)
    Me.rbSpanish.TabIndex = 4
    Me.rbSpanish.TabStop = True
    Me.rbSpanish.Text = "&Spanish"
    Me.rbSpanish.UseVisualStyleBackColor = True
    '
    'rbDutch
    '
    Me.rbDutch.AutoSize = True
    Me.rbDutch.Location = New System.Drawing.Point(6, 90)
    Me.rbDutch.Name = "rbDutch"
    Me.rbDutch.Size = New System.Drawing.Size(54, 17)
    Me.rbDutch.TabIndex = 3
    Me.rbDutch.TabStop = True
    Me.rbDutch.Text = "&Dutch"
    Me.rbDutch.UseVisualStyleBackColor = True
    '
    'rbGerman
    '
    Me.rbGerman.AutoSize = True
    Me.rbGerman.Location = New System.Drawing.Point(6, 67)
    Me.rbGerman.Name = "rbGerman"
    Me.rbGerman.Size = New System.Drawing.Size(62, 17)
    Me.rbGerman.TabIndex = 2
    Me.rbGerman.TabStop = True
    Me.rbGerman.Text = "&German"
    Me.rbGerman.UseVisualStyleBackColor = True
    '
    'rbEnglishTranscribed
    '
    Me.rbEnglishTranscribed.AutoSize = True
    Me.rbEnglishTranscribed.Location = New System.Drawing.Point(6, 43)
    Me.rbEnglishTranscribed.Name = "rbEnglishTranscribed"
    Me.rbEnglishTranscribed.Size = New System.Drawing.Size(168, 17)
    Me.rbEnglishTranscribed.TabIndex = 1
    Me.rbEnglishTranscribed.TabStop = True
    Me.rbEnglishTranscribed.Text = "English - &transcribed dialogues"
    Me.rbEnglishTranscribed.UseVisualStyleBackColor = True
    '
    'rbEnglishWritten
    '
    Me.rbEnglishWritten.AutoSize = True
    Me.rbEnglishWritten.Location = New System.Drawing.Point(6, 19)
    Me.rbEnglishWritten.Name = "rbEnglishWritten"
    Me.rbEnglishWritten.Size = New System.Drawing.Size(99, 17)
    Me.rbEnglishWritten.TabIndex = 0
    Me.rbEnglishWritten.TabStop = True
    Me.rbEnglishWritten.Text = "&English - written"
    Me.rbEnglishWritten.UseVisualStyleBackColor = True
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 357)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(365, 22)
    Me.StatusStrip1.TabIndex = 6
    Me.StatusStrip1.Text = "StatusStrip1"
    '
    'ToolStripStatusLabel1
    '
    Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
    Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(19, 17)
    Me.ToolStripStatusLabel1.Text = "..."
    '
    'Timer1
    '
    '
    'rbChechen
    '
    Me.rbChechen.AutoSize = True
    Me.rbChechen.Location = New System.Drawing.Point(6, 228)
    Me.rbChechen.Name = "rbChechen"
    Me.rbChechen.Size = New System.Drawing.Size(68, 17)
    Me.rbChechen.TabIndex = 12
    Me.rbChechen.TabStop = True
    Me.rbChechen.Text = "&Chechen"
    Me.rbChechen.UseVisualStyleBackColor = True
    '
    'frmConvert
    '
    Me.AcceptButton = Me.cmdOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(365, 379)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.cmdOk)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmConvert"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Select the language (and type) of your texts"
    Me.GroupBox1.ResumeLayout(False)
    Me.GroupBox1.PerformLayout()
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents rbGerman As System.Windows.Forms.RadioButton
  Friend WithEvents rbEnglishTranscribed As System.Windows.Forms.RadioButton
  Friend WithEvents rbEnglishWritten As System.Windows.Forms.RadioButton
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents rbWelsh As System.Windows.Forms.RadioButton
  Friend WithEvents rbFrench As System.Windows.Forms.RadioButton
  Friend WithEvents rbSpanish As System.Windows.Forms.RadioButton
  Friend WithEvents rbDutch As System.Windows.Forms.RadioButton
  Friend WithEvents rbOther As System.Windows.Forms.RadioButton
  Friend WithEvents tbEthno As System.Windows.Forms.TextBox
  Friend WithEvents rbLak As System.Windows.Forms.RadioButton
  Friend WithEvents rbLezgi As System.Windows.Forms.RadioButton
  Friend WithEvents rbChechen As System.Windows.Forms.RadioButton
End Class
