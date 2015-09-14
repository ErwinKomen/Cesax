<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLanguage
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLanguage))
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
    Me.GroupBox1 = New System.Windows.Forms.GroupBox()
    Me.rbSelf = New System.Windows.Forms.RadioButton()
    Me.rbChechen = New System.Windows.Forms.RadioButton()
    Me.rbDutch = New System.Windows.Forms.RadioButton()
    Me.rbEnglish = New System.Windows.Forms.RadioButton()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.Label2 = New System.Windows.Forms.Label()
    Me.cmdOk = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.rbWelsh = New System.Windows.Forms.RadioButton()
    Me.rbGerman = New System.Windows.Forms.RadioButton()
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
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 277)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(456, 22)
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
    Me.GroupBox1.Controls.Add(Me.rbGerman)
    Me.GroupBox1.Controls.Add(Me.rbWelsh)
    Me.GroupBox1.Controls.Add(Me.rbSelf)
    Me.GroupBox1.Controls.Add(Me.rbChechen)
    Me.GroupBox1.Controls.Add(Me.rbDutch)
    Me.GroupBox1.Controls.Add(Me.rbEnglish)
    Me.GroupBox1.Location = New System.Drawing.Point(23, 43)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(215, 222)
    Me.GroupBox1.TabIndex = 1
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Language"
    '
    'rbSelf
    '
    Me.rbSelf.AutoSize = True
    Me.rbSelf.Location = New System.Drawing.Point(6, 198)
    Me.rbSelf.Name = "rbSelf"
    Me.rbSelf.Size = New System.Drawing.Size(192, 17)
    Me.rbSelf.TabIndex = 3
    Me.rbSelf.TabStop = True
    Me.rbSelf.Text = "I locate the period definitions myself"
    Me.rbSelf.UseVisualStyleBackColor = True
    '
    'rbChechen
    '
    Me.rbChechen.AutoSize = True
    Me.rbChechen.Location = New System.Drawing.Point(6, 67)
    Me.rbChechen.Name = "rbChechen"
    Me.rbChechen.Size = New System.Drawing.Size(68, 17)
    Me.rbChechen.TabIndex = 2
    Me.rbChechen.TabStop = True
    Me.rbChechen.Text = "&Chechen"
    Me.rbChechen.UseVisualStyleBackColor = True
    '
    'rbDutch
    '
    Me.rbDutch.AutoSize = True
    Me.rbDutch.Location = New System.Drawing.Point(6, 43)
    Me.rbDutch.Name = "rbDutch"
    Me.rbDutch.Size = New System.Drawing.Size(86, 17)
    Me.rbDutch.TabIndex = 1
    Me.rbDutch.TabStop = True
    Me.rbDutch.Text = "&Dutch (CGN)"
    Me.rbDutch.UseVisualStyleBackColor = True
    '
    'rbEnglish
    '
    Me.rbEnglish.AutoSize = True
    Me.rbEnglish.Location = New System.Drawing.Point(6, 19)
    Me.rbEnglish.Name = "rbEnglish"
    Me.rbEnglish.Size = New System.Drawing.Size(59, 17)
    Me.rbEnglish.TabIndex = 0
    Me.rbEnglish.TabStop = True
    Me.rbEnglish.Text = "&English"
    Me.rbEnglish.UseVisualStyleBackColor = True
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.ForeColor = System.Drawing.Color.Maroon
    Me.Label1.Location = New System.Drawing.Point(20, 14)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(392, 13)
    Me.Label1.TabIndex = 2
    Me.Label1.Text = "Please indicate from what language you would like to download period information"
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.ForeColor = System.Drawing.Color.Navy
    Me.Label2.Location = New System.Drawing.Point(20, 27)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(432, 13)
    Me.Label2.TabIndex = 3
    Me.Label2.Text = "(Or if you already have a period definition file, select ""I locate the period def" & _
        "initions myself"")"
    '
    'cmdOk
    '
    Me.cmdOk.Location = New System.Drawing.Point(377, 206)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(75, 23)
    Me.cmdOk.TabIndex = 4
    Me.cmdOk.Text = "&Ok"
    Me.cmdOk.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(377, 241)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 5
    Me.cmdCancel.Text = "E&xit"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'rbWelsh
    '
    Me.rbWelsh.AutoSize = True
    Me.rbWelsh.Location = New System.Drawing.Point(6, 90)
    Me.rbWelsh.Name = "rbWelsh"
    Me.rbWelsh.Size = New System.Drawing.Size(55, 17)
    Me.rbWelsh.TabIndex = 4
    Me.rbWelsh.TabStop = True
    Me.rbWelsh.Text = "&Welsh"
    Me.rbWelsh.UseVisualStyleBackColor = True
    '
    'rbGerman
    '
    Me.rbGerman.AutoSize = True
    Me.rbGerman.Location = New System.Drawing.Point(6, 113)
    Me.rbGerman.Name = "rbGerman"
    Me.rbGerman.Size = New System.Drawing.Size(62, 17)
    Me.rbGerman.TabIndex = 5
    Me.rbGerman.TabStop = True
    Me.rbGerman.Text = "&German"
    Me.rbGerman.UseVisualStyleBackColor = True
    '
    'frmLanguage
    '
    Me.AcceptButton = Me.cmdOk
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(456, 299)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.cmdOk)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.StatusStrip1)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmLanguage"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Download period definitions"
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
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents rbSelf As System.Windows.Forms.RadioButton
  Friend WithEvents rbChechen As System.Windows.Forms.RadioButton
  Friend WithEvents rbDutch As System.Windows.Forms.RadioButton
  Friend WithEvents rbEnglish As System.Windows.Forms.RadioButton
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents rbGerman As System.Windows.Forms.RadioButton
  Friend WithEvents rbWelsh As System.Windows.Forms.RadioButton
End Class
