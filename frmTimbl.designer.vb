<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTimbl
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
    Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
    Me.OK_Button = New System.Windows.Forms.Button
    Me.Cancel_Button = New System.Windows.Forms.Button
    Me.Label1 = New System.Windows.Forms.Label
    Me.lboPeriod = New System.Windows.Forms.ListBox
    Me.Label2 = New System.Windows.Forms.Label
    Me.tbPtcTrain = New System.Windows.Forms.TextBox
    Me.tbPtcTest = New System.Windows.Forms.TextBox
    Me.Label3 = New System.Windows.Forms.Label
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.TableLayoutPanel1.SuspendLayout()
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
    Me.TableLayoutPanel1.Location = New System.Drawing.Point(277, 150)
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
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(12, 9)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(208, 13)
    Me.Label1.TabIndex = 1
    Me.Label1.Text = "Select the periods you want to be included"
    '
    'lboPeriod
    '
    Me.lboPeriod.FormattingEnabled = True
    Me.lboPeriod.HorizontalScrollbar = True
    Me.lboPeriod.Location = New System.Drawing.Point(15, 25)
    Me.lboPeriod.MultiColumn = True
    Me.lboPeriod.Name = "lboPeriod"
    Me.lboPeriod.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
    Me.lboPeriod.Size = New System.Drawing.Size(405, 95)
    Me.lboPeriod.TabIndex = 2
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(12, 135)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(116, 13)
    Me.Label2.TabIndex = 3
    Me.Label2.Text = "Percentage trainingset:"
    '
    'tbPtcTrain
    '
    Me.tbPtcTrain.Location = New System.Drawing.Point(134, 132)
    Me.tbPtcTrain.Name = "tbPtcTrain"
    Me.tbPtcTrain.Size = New System.Drawing.Size(43, 20)
    Me.tbPtcTrain.TabIndex = 4
    '
    'tbPtcTest
    '
    Me.tbPtcTest.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbPtcTest.Location = New System.Drawing.Point(134, 158)
    Me.tbPtcTest.Name = "tbPtcTest"
    Me.tbPtcTest.ReadOnly = True
    Me.tbPtcTest.Size = New System.Drawing.Size(43, 20)
    Me.tbPtcTest.TabIndex = 6
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(12, 161)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(99, 13)
    Me.Label3.TabIndex = 5
    Me.Label3.Text = "Percentage testset:"
    '
    'Timer1
    '
    '
    'frmTimbl
    '
    Me.AcceptButton = Me.OK_Button
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.Cancel_Button
    Me.ClientSize = New System.Drawing.Size(435, 191)
    Me.Controls.Add(Me.tbPtcTest)
    Me.Controls.Add(Me.Label3)
    Me.Controls.Add(Me.tbPtcTrain)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.lboPeriod)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.TableLayoutPanel1)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmTimbl"
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Adjust parameters for TIMBL preparation"
    Me.TableLayoutPanel1.ResumeLayout(False)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
  Friend WithEvents Cancel_Button As System.Windows.Forms.Button
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents lboPeriod As System.Windows.Forms.ListBox
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents tbPtcTrain As System.Windows.Forms.TextBox
  Friend WithEvents tbPtcTest As System.Windows.Forms.TextBox
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents Timer1 As System.Windows.Forms.Timer

End Class
