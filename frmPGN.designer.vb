<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPGN
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPGN))
    Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
    Me.OK_Button = New System.Windows.Forms.Button
    Me.Cancel_Button = New System.Windows.Forms.Button
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.Label1 = New System.Windows.Forms.Label
    Me.Label3 = New System.Windows.Forms.Label
    Me.tbPGN = New System.Windows.Forms.TextBox
    Me.Label2 = New System.Windows.Forms.Label
    Me.lboPGN = New System.Windows.Forms.ListBox
    Me.Label4 = New System.Windows.Forms.Label
    Me.wbClause = New System.Windows.Forms.WebBrowser
    Me.Label5 = New System.Windows.Forms.Label
    Me.cboNPtype = New System.Windows.Forms.ComboBox
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
    Me.TableLayoutPanel1.Location = New System.Drawing.Point(430, 377)
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
    'Timer1
    '
    '
    'Label1
    '
    Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(127, 230)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(159, 13)
    Me.Label1.TabIndex = 1
    Me.Label1.Text = "Correct PGN for this constituent:"
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(12, 9)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(42, 13)
    Me.Label3.TabIndex = 4
    Me.Label3.Text = "Clause:"
    '
    'tbPGN
    '
    Me.tbPGN.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.tbPGN.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.tbPGN.ForeColor = System.Drawing.Color.Blue
    Me.tbPGN.Location = New System.Drawing.Point(130, 252)
    Me.tbPGN.Name = "tbPGN"
    Me.tbPGN.Size = New System.Drawing.Size(100, 21)
    Me.tbPGN.TabIndex = 6
    '
    'Label2
    '
    Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(12, 230)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(109, 13)
    Me.Label2.TabIndex = 7
    Me.Label2.Text = "Possible PGN values:"
    '
    'lboPGN
    '
    Me.lboPGN.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.lboPGN.FormattingEnabled = True
    Me.lboPGN.Location = New System.Drawing.Point(15, 246)
    Me.lboPGN.Name = "lboPGN"
    Me.lboPGN.Size = New System.Drawing.Size(106, 160)
    Me.lboPGN.TabIndex = 8
    '
    'Label4
    '
    Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.Label4.AutoSize = True
    Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
    Me.Label4.Location = New System.Drawing.Point(236, 255)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(85, 13)
    Me.Label4.TabIndex = 9
    Me.Label4.Text = "E.g: 3, 3s, or 2fp"
    '
    'wbClause
    '
    Me.wbClause.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.wbClause.Location = New System.Drawing.Point(15, 26)
    Me.wbClause.MinimumSize = New System.Drawing.Size(20, 20)
    Me.wbClause.Name = "wbClause"
    Me.wbClause.Size = New System.Drawing.Size(558, 198)
    Me.wbClause.TabIndex = 10
    '
    'Label5
    '
    Me.Label5.AutoSize = True
    Me.Label5.Location = New System.Drawing.Point(127, 284)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(271, 13)
    Me.Label5.TabIndex = 11
    Me.Label5.Text = "Is this a Personal Pronoun or a Demonstrative Pronoun?"
    '
    'cboNPtype
    '
    Me.cboNPtype.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboNPtype.FormattingEnabled = True
    Me.cboNPtype.Location = New System.Drawing.Point(130, 300)
    Me.cboNPtype.Name = "cboNPtype"
    Me.cboNPtype.Size = New System.Drawing.Size(121, 21)
    Me.cboNPtype.TabIndex = 12
    '
    'frmPGN
    '
    Me.AcceptButton = Me.OK_Button
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.Cancel_Button
    Me.ClientSize = New System.Drawing.Size(588, 418)
    Me.Controls.Add(Me.cboNPtype)
    Me.Controls.Add(Me.Label5)
    Me.Controls.Add(Me.wbClause)
    Me.Controls.Add(Me.Label4)
    Me.Controls.Add(Me.lboPGN)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.tbPGN)
    Me.Controls.Add(Me.Label3)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.TableLayoutPanel1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmPGN"
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Please supply the correct Person/Gender/Number for the indicated constituent"
    Me.TableLayoutPanel1.ResumeLayout(False)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
  Friend WithEvents OK_Button As System.Windows.Forms.Button
  Friend WithEvents Cancel_Button As System.Windows.Forms.Button
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents tbPGN As System.Windows.Forms.TextBox
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents lboPGN As System.Windows.Forms.ListBox
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents wbClause As System.Windows.Forms.WebBrowser
  Friend WithEvents Label5 As System.Windows.Forms.Label
  Friend WithEvents cboNPtype As System.Windows.Forms.ComboBox

End Class
