<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCategory
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
    Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
    Me.OK_Button = New System.Windows.Forms.Button()
    Me.Cancel_Button = New System.Windows.Forms.Button()
    Me.rbNP = New System.Windows.Forms.RadioButton()
    Me.rbAdv = New System.Windows.Forms.RadioButton()
    Me.rbVb = New System.Windows.Forms.RadioButton()
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.chbCheck = New System.Windows.Forms.CheckBox()
    Me.GroupBox1 = New System.Windows.Forms.GroupBox()
    Me.rbVbUnacc = New System.Windows.Forms.RadioButton()
    Me.rbPrepNorm = New System.Windows.Forms.RadioButton()
    Me.rbCogn = New System.Windows.Forms.RadioButton()
    Me.rbAnim = New System.Windows.Forms.RadioButton()
    Me.rbWh = New System.Windows.Forms.RadioButton()
    Me.rbGr = New System.Windows.Forms.RadioButton()
    Me.grpOptions = New System.Windows.Forms.GroupBox()
    Me.Label1 = New System.Windows.Forms.Label()
    Me.tbFeatDefFile = New System.Windows.Forms.TextBox()
    Me.cmdFeatDefFile = New System.Windows.Forms.Button()
    Me.tbVbType = New System.Windows.Forms.RadioButton()
    Me.TableLayoutPanel1.SuspendLayout()
    Me.GroupBox1.SuspendLayout()
    Me.grpOptions.SuspendLayout()
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
    Me.TableLayoutPanel1.Location = New System.Drawing.Point(180, 334)
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
    Me.OK_Button.TabIndex = 8
    Me.OK_Button.Text = "OK"
    '
    'Cancel_Button
    '
    Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
    Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
    Me.Cancel_Button.Name = "Cancel_Button"
    Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
    Me.Cancel_Button.TabIndex = 9
    Me.Cancel_Button.Text = "Cancel"
    '
    'rbNP
    '
    Me.rbNP.AutoSize = True
    Me.rbNP.Location = New System.Drawing.Point(6, 65)
    Me.rbNP.Name = "rbNP"
    Me.rbNP.Size = New System.Drawing.Size(303, 17)
    Me.rbNP.TabIndex = 1
    Me.rbNP.TabStop = True
    Me.rbNP.Text = "&NP - all noun phrase features: NPtype, GrRole, PGN, IPdist"
    Me.rbNP.UseVisualStyleBackColor = True
    '
    'rbAdv
    '
    Me.rbAdv.AutoSize = True
    Me.rbAdv.Location = New System.Drawing.Point(6, 19)
    Me.rbAdv.Name = "rbAdv"
    Me.rbAdv.Size = New System.Drawing.Size(200, 17)
    Me.rbAdv.TabIndex = 5
    Me.rbAdv.TabStop = True
    Me.rbAdv.Text = "&Adv - adverbs modifying an NP or PP"
    Me.rbAdv.UseVisualStyleBackColor = True
    '
    'rbVb
    '
    Me.rbVb.AutoSize = True
    Me.rbVb.Location = New System.Drawing.Point(6, 180)
    Me.rbVb.Name = "rbVb"
    Me.rbVb.Size = New System.Drawing.Size(73, 17)
    Me.rbVb.TabIndex = 6
    Me.rbVb.TabStop = True
    Me.rbVb.Text = "&Vb - verbs"
    Me.rbVb.UseVisualStyleBackColor = True
    '
    'Timer1
    '
    '
    'chbCheck
    '
    Me.chbCheck.AutoSize = True
    Me.chbCheck.Location = New System.Drawing.Point(6, 19)
    Me.chbCheck.Name = "chbCheck"
    Me.chbCheck.Size = New System.Drawing.Size(114, 17)
    Me.chbCheck.TabIndex = 7
    Me.chbCheck.Text = "&Check Differences"
    Me.chbCheck.UseVisualStyleBackColor = True
    '
    'GroupBox1
    '
    Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GroupBox1.Controls.Add(Me.tbVbType)
    Me.GroupBox1.Controls.Add(Me.rbVbUnacc)
    Me.GroupBox1.Controls.Add(Me.rbPrepNorm)
    Me.GroupBox1.Controls.Add(Me.rbCogn)
    Me.GroupBox1.Controls.Add(Me.rbAnim)
    Me.GroupBox1.Controls.Add(Me.rbWh)
    Me.GroupBox1.Controls.Add(Me.rbGr)
    Me.GroupBox1.Controls.Add(Me.rbNP)
    Me.GroupBox1.Controls.Add(Me.rbAdv)
    Me.GroupBox1.Controls.Add(Me.rbVb)
    Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(311, 257)
    Me.GroupBox1.TabIndex = 5
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Feature Category"
    '
    'rbVbUnacc
    '
    Me.rbVbUnacc.AutoSize = True
    Me.rbVbUnacc.Location = New System.Drawing.Point(6, 203)
    Me.rbVbUnacc.Name = "rbVbUnacc"
    Me.rbVbUnacc.Size = New System.Drawing.Size(188, 17)
    Me.rbVbUnacc.TabIndex = 9
    Me.rbVbUnacc.TabStop = True
    Me.rbVbUnacc.Text = "Vb - verbs &unaccusative from psdx"
    Me.rbVbUnacc.UseVisualStyleBackColor = True
    '
    'rbPrepNorm
    '
    Me.rbPrepNorm.AutoSize = True
    Me.rbPrepNorm.Location = New System.Drawing.Point(6, 157)
    Me.rbPrepNorm.Name = "rbPrepNorm"
    Me.rbPrepNorm.Size = New System.Drawing.Size(141, 17)
    Me.rbPrepNorm.TabIndex = 8
    Me.rbPrepNorm.TabStop = True
    Me.rbPrepNorm.Text = "&Preposition normalisation"
    Me.rbPrepNorm.UseVisualStyleBackColor = True
    '
    'rbCogn
    '
    Me.rbCogn.AutoSize = True
    Me.rbCogn.Location = New System.Drawing.Point(6, 42)
    Me.rbCogn.Name = "rbCogn"
    Me.rbCogn.Size = New System.Drawing.Size(179, 17)
    Me.rbCogn.TabIndex = 7
    Me.rbCogn.TabStop = True
    Me.rbCogn.Text = "&Cognitive status - NPs in clauses"
    Me.rbCogn.UseVisualStyleBackColor = True
    '
    'rbAnim
    '
    Me.rbAnim.AutoSize = True
    Me.rbAnim.Location = New System.Drawing.Point(6, 134)
    Me.rbAnim.Name = "rbAnim"
    Me.rbAnim.Size = New System.Drawing.Size(88, 17)
    Me.rbAnim.TabIndex = 4
    Me.rbAnim.TabStop = True
    Me.rbAnim.Text = "NP - ani&macy"
    Me.rbAnim.UseVisualStyleBackColor = True
    '
    'rbWh
    '
    Me.rbWh.AutoSize = True
    Me.rbWh.Location = New System.Drawing.Point(6, 111)
    Me.rbWh.Name = "rbWh"
    Me.rbWh.Size = New System.Drawing.Size(108, 17)
    Me.rbWh.TabIndex = 3
    Me.rbWh.TabStop = True
    Me.rbWh.Text = "NP - &wh elements"
    Me.rbWh.UseVisualStyleBackColor = True
    '
    'rbGr
    '
    Me.rbGr.AutoSize = True
    Me.rbGr.Location = New System.Drawing.Point(6, 88)
    Me.rbGr.Name = "rbGr"
    Me.rbGr.Size = New System.Drawing.Size(147, 17)
    Me.rbGr.TabIndex = 2
    Me.rbGr.TabStop = True
    Me.rbGr.Text = "NP - &grammatical role only"
    Me.rbGr.UseVisualStyleBackColor = True
    '
    'grpOptions
    '
    Me.grpOptions.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.grpOptions.Controls.Add(Me.chbCheck)
    Me.grpOptions.Location = New System.Drawing.Point(12, 312)
    Me.grpOptions.Name = "grpOptions"
    Me.grpOptions.Size = New System.Drawing.Size(162, 48)
    Me.grpOptions.TabIndex = 6
    Me.grpOptions.TabStop = False
    Me.grpOptions.Text = "Options"
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(15, 272)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(107, 13)
    Me.Label1.TabIndex = 7
    Me.Label1.Text = "Feature definition file:"
    '
    'tbFeatDefFile
    '
    Me.tbFeatDefFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbFeatDefFile.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbFeatDefFile.Location = New System.Drawing.Point(12, 288)
    Me.tbFeatDefFile.Name = "tbFeatDefFile"
    Me.tbFeatDefFile.ReadOnly = True
    Me.tbFeatDefFile.Size = New System.Drawing.Size(257, 20)
    Me.tbFeatDefFile.TabIndex = 8
    '
    'cmdFeatDefFile
    '
    Me.cmdFeatDefFile.Location = New System.Drawing.Point(275, 288)
    Me.cmdFeatDefFile.Name = "cmdFeatDefFile"
    Me.cmdFeatDefFile.Size = New System.Drawing.Size(48, 20)
    Me.cmdFeatDefFile.TabIndex = 9
    Me.cmdFeatDefFile.Text = "..."
    Me.cmdFeatDefFile.UseVisualStyleBackColor = True
    '
    'tbVbType
    '
    Me.tbVbType.AutoSize = True
    Me.tbVbType.Location = New System.Drawing.Point(6, 226)
    Me.tbVbType.Name = "tbVbType"
    Me.tbVbType.Size = New System.Drawing.Size(145, 17)
    Me.tbVbType.TabIndex = 10
    Me.tbVbType.TabStop = True
    Me.tbVbType.Text = "Vb - verb types in general"
    Me.tbVbType.UseVisualStyleBackColor = True
    '
    'frmCategory
    '
    Me.AcceptButton = Me.OK_Button
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.Cancel_Button
    Me.ClientSize = New System.Drawing.Size(338, 375)
    Me.Controls.Add(Me.cmdFeatDefFile)
    Me.Controls.Add(Me.tbFeatDefFile)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.grpOptions)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.TableLayoutPanel1)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmCategory"
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Select the feature category"
    Me.TableLayoutPanel1.ResumeLayout(False)
    Me.GroupBox1.ResumeLayout(False)
    Me.GroupBox1.PerformLayout()
    Me.grpOptions.ResumeLayout(False)
    Me.grpOptions.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
  Friend WithEvents OK_Button As System.Windows.Forms.Button
  Friend WithEvents Cancel_Button As System.Windows.Forms.Button
  Friend WithEvents rbNP As System.Windows.Forms.RadioButton
  Friend WithEvents rbAdv As System.Windows.Forms.RadioButton
  Friend WithEvents rbVb As System.Windows.Forms.RadioButton
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents chbCheck As System.Windows.Forms.CheckBox
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents grpOptions As System.Windows.Forms.GroupBox
  Friend WithEvents rbGr As System.Windows.Forms.RadioButton
  Friend WithEvents rbWh As System.Windows.Forms.RadioButton
  Friend WithEvents rbAnim As System.Windows.Forms.RadioButton
  Friend WithEvents rbCogn As System.Windows.Forms.RadioButton
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents tbFeatDefFile As System.Windows.Forms.TextBox
  Friend WithEvents cmdFeatDefFile As System.Windows.Forms.Button
  Friend WithEvents rbPrepNorm As System.Windows.Forms.RadioButton
  Friend WithEvents rbVbUnacc As System.Windows.Forms.RadioButton
  Friend WithEvents tbVbType As System.Windows.Forms.RadioButton

End Class
