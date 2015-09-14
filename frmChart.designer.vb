<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChart
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
    Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
    Me.OK_Button = New System.Windows.Forms.Button
    Me.Cancel_Button = New System.Windows.Forms.Button
    Me.chbSubCl = New System.Windows.Forms.CheckBox
    Me.grpOutput = New System.Windows.Forms.GroupBox
    Me.rbChart = New System.Windows.Forms.RadioButton
    Me.rbWordOrder = New System.Windows.Forms.RadioButton
    Me.GroupBox1 = New System.Windows.Forms.GroupBox
    Me.numSptc = New System.Windows.Forms.NumericUpDown
    Me.numV2ptc = New System.Windows.Forms.NumericUpDown
    Me.numV1ptc = New System.Windows.Forms.NumericUpDown
    Me.Label3 = New System.Windows.Forms.Label
    Me.Label2 = New System.Windows.Forms.Label
    Me.Label1 = New System.Windows.Forms.Label
    Me.chbAllSections = New System.Windows.Forms.CheckBox
    Me.TableLayoutPanel1.SuspendLayout()
    Me.grpOutput.SuspendLayout()
    Me.GroupBox1.SuspendLayout()
    CType(Me.numSptc, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.numV2ptc, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.numV1ptc, System.ComponentModel.ISupportInitialize).BeginInit()
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
    Me.TableLayoutPanel1.Location = New System.Drawing.Point(217, 142)
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
    'chbSubCl
    '
    Me.chbSubCl.AutoSize = True
    Me.chbSubCl.Location = New System.Drawing.Point(12, 94)
    Me.chbSubCl.Name = "chbSubCl"
    Me.chbSubCl.Size = New System.Drawing.Size(120, 17)
    Me.chbSubCl.TabIndex = 1
    Me.chbSubCl.Text = "Expand &Subclauses"
    Me.chbSubCl.UseVisualStyleBackColor = True
    '
    'grpOutput
    '
    Me.grpOutput.Controls.Add(Me.rbChart)
    Me.grpOutput.Controls.Add(Me.rbWordOrder)
    Me.grpOutput.Location = New System.Drawing.Point(12, 6)
    Me.grpOutput.Name = "grpOutput"
    Me.grpOutput.Size = New System.Drawing.Size(145, 82)
    Me.grpOutput.TabIndex = 2
    Me.grpOutput.TabStop = False
    Me.grpOutput.Text = "Output type"
    '
    'rbChart
    '
    Me.rbChart.AutoSize = True
    Me.rbChart.Location = New System.Drawing.Point(7, 44)
    Me.rbChart.Name = "rbChart"
    Me.rbChart.Size = New System.Drawing.Size(96, 17)
    Me.rbChart.TabIndex = 1
    Me.rbChart.TabStop = True
    Me.rbChart.Text = "&Chart in a table"
    Me.rbChart.UseVisualStyleBackColor = True
    '
    'rbWordOrder
    '
    Me.rbWordOrder.AutoSize = True
    Me.rbWordOrder.Location = New System.Drawing.Point(7, 20)
    Me.rbWordOrder.Name = "rbWordOrder"
    Me.rbWordOrder.Size = New System.Drawing.Size(124, 17)
    Me.rbWordOrder.TabIndex = 0
    Me.rbWordOrder.TabStop = True
    Me.rbWordOrder.Text = "&Word order overview"
    Me.rbWordOrder.UseVisualStyleBackColor = True
    '
    'GroupBox1
    '
    Me.GroupBox1.Controls.Add(Me.numSptc)
    Me.GroupBox1.Controls.Add(Me.numV2ptc)
    Me.GroupBox1.Controls.Add(Me.numV1ptc)
    Me.GroupBox1.Controls.Add(Me.Label3)
    Me.GroupBox1.Controls.Add(Me.Label2)
    Me.GroupBox1.Controls.Add(Me.Label1)
    Me.GroupBox1.Location = New System.Drawing.Point(163, 6)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(200, 128)
    Me.GroupBox1.TabIndex = 3
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Cut-off percentages"
    '
    'numSptc
    '
    Me.numSptc.Location = New System.Drawing.Point(89, 70)
    Me.numSptc.Minimum = New Decimal(New Integer() {50, 0, 0, 0})
    Me.numSptc.Name = "numSptc"
    Me.numSptc.Size = New System.Drawing.Size(78, 20)
    Me.numSptc.TabIndex = 8
    Me.numSptc.Value = New Decimal(New Integer() {50, 0, 0, 0})
    '
    'numV2ptc
    '
    Me.numV2ptc.Location = New System.Drawing.Point(89, 44)
    Me.numV2ptc.Minimum = New Decimal(New Integer() {50, 0, 0, 0})
    Me.numV2ptc.Name = "numV2ptc"
    Me.numV2ptc.Size = New System.Drawing.Size(78, 20)
    Me.numV2ptc.TabIndex = 7
    Me.numV2ptc.Value = New Decimal(New Integer() {50, 0, 0, 0})
    '
    'numV1ptc
    '
    Me.numV1ptc.Location = New System.Drawing.Point(89, 20)
    Me.numV1ptc.Minimum = New Decimal(New Integer() {50, 0, 0, 0})
    Me.numV1ptc.Name = "numV1ptc"
    Me.numV1ptc.Size = New System.Drawing.Size(78, 20)
    Me.numV1ptc.TabIndex = 6
    Me.numV1ptc.Value = New Decimal(New Integer() {50, 0, 0, 0})
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(6, 72)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(79, 13)
    Me.Label3.TabIndex = 4
    Me.Label3.Text = "Sbj percentage"
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(6, 46)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(77, 13)
    Me.Label2.TabIndex = 2
    Me.Label2.Text = "V2 percentage"
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(6, 22)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(77, 13)
    Me.Label1.TabIndex = 0
    Me.Label1.Text = "V1 percentage"
    '
    'chbAllSections
    '
    Me.chbAllSections.AutoSize = True
    Me.chbAllSections.Location = New System.Drawing.Point(12, 117)
    Me.chbAllSections.Name = "chbAllSections"
    Me.chbAllSections.Size = New System.Drawing.Size(116, 17)
    Me.chbAllSections.TabIndex = 4
    Me.chbAllSections.Text = "Include &all sections"
    Me.chbAllSections.UseVisualStyleBackColor = True
    '
    'frmChart
    '
    Me.AcceptButton = Me.OK_Button
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.Cancel_Button
    Me.ClientSize = New System.Drawing.Size(376, 183)
    Me.Controls.Add(Me.chbAllSections)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.grpOutput)
    Me.Controls.Add(Me.chbSubCl)
    Me.Controls.Add(Me.TableLayoutPanel1)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmChart"
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Options for this chart"
    Me.TableLayoutPanel1.ResumeLayout(False)
    Me.grpOutput.ResumeLayout(False)
    Me.grpOutput.PerformLayout()
    Me.GroupBox1.ResumeLayout(False)
    Me.GroupBox1.PerformLayout()
    CType(Me.numSptc, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.numV2ptc, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.numV1ptc, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
  Friend WithEvents OK_Button As System.Windows.Forms.Button
  Friend WithEvents Cancel_Button As System.Windows.Forms.Button
  Friend WithEvents chbSubCl As System.Windows.Forms.CheckBox
  Friend WithEvents grpOutput As System.Windows.Forms.GroupBox
  Friend WithEvents rbChart As System.Windows.Forms.RadioButton
  Friend WithEvents rbWordOrder As System.Windows.Forms.RadioButton
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents numV1ptc As System.Windows.Forms.NumericUpDown
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents numSptc As System.Windows.Forms.NumericUpDown
  Friend WithEvents numV2ptc As System.Windows.Forms.NumericUpDown
  Friend WithEvents chbAllSections As System.Windows.Forms.CheckBox

End Class
