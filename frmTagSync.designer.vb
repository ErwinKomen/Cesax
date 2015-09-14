<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTagSync
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTagSync))
    Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
    Me.cmdNo = New System.Windows.Forms.Button
    Me.cmdYes = New System.Windows.Forms.Button
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.Label1 = New System.Windows.Forms.Label
    Me.Label2 = New System.Windows.Forms.Label
    Me.wbOEtagged = New System.Windows.Forms.WebBrowser
    Me.wbYCOE = New System.Windows.Forms.WebBrowser
    Me.Label3 = New System.Windows.Forms.Label
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.cmdOEtagLeft = New System.Windows.Forms.Button
    Me.cmdYcoeLeft = New System.Windows.Forms.Button
    Me.cmdOEtagRight = New System.Windows.Forms.Button
    Me.cmdYcoeRight = New System.Windows.Forms.Button
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar
    Me.cmdYcoeNextLine = New System.Windows.Forms.Button
    Me.cmdYcoeSearch = New System.Windows.Forms.Button
    Me.cmdOEtagSearch = New System.Windows.Forms.Button
    Me.Label4 = New System.Windows.Forms.Label
    Me.tbYcoeLine = New System.Windows.Forms.TextBox
    Me.cmdYcoeGoto = New System.Windows.Forms.Button
    Me.tbOEfile = New System.Windows.Forms.TextBox
    Me.TableLayoutPanel1.SuspendLayout()
    Me.StatusStrip1.SuspendLayout()
    Me.SuspendLayout()
    '
    'TableLayoutPanel1
    '
    Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.TableLayoutPanel1.ColumnCount = 3
    Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 52.17391!))
    Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 47.82609!))
    Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 91.0!))
    Me.TableLayoutPanel1.Controls.Add(Me.cmdNo, 0, 0)
    Me.TableLayoutPanel1.Controls.Add(Me.cmdYes, 0, 0)
    Me.TableLayoutPanel1.Controls.Add(Me.cmdCancel, 0, 0)
    Me.TableLayoutPanel1.Location = New System.Drawing.Point(641, 438)
    Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
    Me.TableLayoutPanel1.RowCount = 1
    Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
    Me.TableLayoutPanel1.Size = New System.Drawing.Size(242, 29)
    Me.TableLayoutPanel1.TabIndex = 0
    '
    'cmdNo
    '
    Me.cmdNo.Anchor = System.Windows.Forms.AnchorStyles.None
    Me.cmdNo.DialogResult = System.Windows.Forms.DialogResult.No
    Me.cmdNo.Location = New System.Drawing.Point(159, 3)
    Me.cmdNo.Name = "cmdNo"
    Me.cmdNo.Size = New System.Drawing.Size(74, 23)
    Me.cmdNo.TabIndex = 2
    Me.cmdNo.Text = "&No"
    '
    'cmdYes
    '
    Me.cmdYes.Anchor = System.Windows.Forms.AnchorStyles.None
    Me.cmdYes.DialogResult = System.Windows.Forms.DialogResult.Yes
    Me.cmdYes.Location = New System.Drawing.Point(81, 3)
    Me.cmdYes.Name = "cmdYes"
    Me.cmdYes.Size = New System.Drawing.Size(66, 23)
    Me.cmdYes.TabIndex = 0
    Me.cmdYes.Text = "&Yes"
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = System.Windows.Forms.AnchorStyles.None
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(3, 3)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(72, 23)
    Me.cmdCancel.TabIndex = 1
    Me.cmdCancel.Text = "Cancel"
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.ForeColor = System.Drawing.Color.Maroon
    Me.Label1.Location = New System.Drawing.Point(12, 9)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(347, 26)
    Me.Label1.TabIndex = 1
    Me.Label1.Text = "Is the bolded red word in the two versions to be regarded as equivalent?" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Provide" & _
        " re-synchronization through the < and > buttons."
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(12, 50)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(168, 13)
    Me.Label2.TabIndex = 2
    Me.Label2.Text = "Text from the ""OE-tagged"" corpus"
    '
    'wbOEtagged
    '
    Me.wbOEtagged.Location = New System.Drawing.Point(15, 66)
    Me.wbOEtagged.MinimumSize = New System.Drawing.Size(20, 20)
    Me.wbOEtagged.Name = "wbOEtagged"
    Me.wbOEtagged.Size = New System.Drawing.Size(865, 142)
    Me.wbOEtagged.TabIndex = 3
    '
    'wbYCOE
    '
    Me.wbYCOE.Location = New System.Drawing.Point(15, 251)
    Me.wbYCOE.MinimumSize = New System.Drawing.Size(20, 20)
    Me.wbYCOE.Name = "wbYCOE"
    Me.wbYCOE.Size = New System.Drawing.Size(865, 181)
    Me.wbYCOE.TabIndex = 5
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(12, 235)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(101, 13)
    Me.Label3.TabIndex = 4
    Me.Label3.Text = "Text from the YCOE"
    '
    'Timer1
    '
    '
    'cmdOEtagLeft
    '
    Me.cmdOEtagLeft.Location = New System.Drawing.Point(186, 40)
    Me.cmdOEtagLeft.Name = "cmdOEtagLeft"
    Me.cmdOEtagLeft.Size = New System.Drawing.Size(31, 23)
    Me.cmdOEtagLeft.TabIndex = 6
    Me.cmdOEtagLeft.Text = "<"
    Me.cmdOEtagLeft.UseVisualStyleBackColor = True
    '
    'cmdYcoeLeft
    '
    Me.cmdYcoeLeft.Location = New System.Drawing.Point(186, 225)
    Me.cmdYcoeLeft.Name = "cmdYcoeLeft"
    Me.cmdYcoeLeft.Size = New System.Drawing.Size(31, 23)
    Me.cmdYcoeLeft.TabIndex = 7
    Me.cmdYcoeLeft.Text = "<"
    Me.cmdYcoeLeft.UseVisualStyleBackColor = True
    '
    'cmdOEtagRight
    '
    Me.cmdOEtagRight.Location = New System.Drawing.Point(223, 40)
    Me.cmdOEtagRight.Name = "cmdOEtagRight"
    Me.cmdOEtagRight.Size = New System.Drawing.Size(31, 23)
    Me.cmdOEtagRight.TabIndex = 8
    Me.cmdOEtagRight.Text = ">"
    Me.cmdOEtagRight.UseVisualStyleBackColor = True
    '
    'cmdYcoeRight
    '
    Me.cmdYcoeRight.Location = New System.Drawing.Point(223, 225)
    Me.cmdYcoeRight.Name = "cmdYcoeRight"
    Me.cmdYcoeRight.Size = New System.Drawing.Size(31, 23)
    Me.cmdYcoeRight.TabIndex = 9
    Me.cmdYcoeRight.Text = ">"
    Me.cmdYcoeRight.UseVisualStyleBackColor = True
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 470)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(895, 22)
    Me.StatusStrip1.TabIndex = 10
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
    Me.ToolStripProgressBar1.Size = New System.Drawing.Size(400, 16)
    '
    'cmdYcoeNextLine
    '
    Me.cmdYcoeNextLine.Location = New System.Drawing.Point(284, 225)
    Me.cmdYcoeNextLine.Name = "cmdYcoeNextLine"
    Me.cmdYcoeNextLine.Size = New System.Drawing.Size(75, 23)
    Me.cmdYcoeNextLine.TabIndex = 11
    Me.cmdYcoeNextLine.Text = "Next &Line"
    Me.cmdYcoeNextLine.UseVisualStyleBackColor = True
    '
    'cmdYcoeSearch
    '
    Me.cmdYcoeSearch.Location = New System.Drawing.Point(379, 225)
    Me.cmdYcoeSearch.Name = "cmdYcoeSearch"
    Me.cmdYcoeSearch.Size = New System.Drawing.Size(75, 23)
    Me.cmdYcoeSearch.TabIndex = 12
    Me.cmdYcoeSearch.Text = "&Search..."
    Me.cmdYcoeSearch.UseVisualStyleBackColor = True
    '
    'cmdOEtagSearch
    '
    Me.cmdOEtagSearch.Location = New System.Drawing.Point(379, 40)
    Me.cmdOEtagSearch.Name = "cmdOEtagSearch"
    Me.cmdOEtagSearch.Size = New System.Drawing.Size(75, 23)
    Me.cmdOEtagSearch.TabIndex = 13
    Me.cmdOEtagSearch.Text = "S&earch..."
    Me.cmdOEtagSearch.UseVisualStyleBackColor = True
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(472, 235)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(166, 13)
    Me.Label4.TabIndex = 14
    Me.Label4.Text = "Goto the following line (last digits):"
    '
    'tbYcoeLine
    '
    Me.tbYcoeLine.Location = New System.Drawing.Point(641, 228)
    Me.tbYcoeLine.Name = "tbYcoeLine"
    Me.tbYcoeLine.Size = New System.Drawing.Size(100, 20)
    Me.tbYcoeLine.TabIndex = 15
    '
    'cmdYcoeGoto
    '
    Me.cmdYcoeGoto.Location = New System.Drawing.Point(747, 225)
    Me.cmdYcoeGoto.Name = "cmdYcoeGoto"
    Me.cmdYcoeGoto.Size = New System.Drawing.Size(75, 23)
    Me.cmdYcoeGoto.TabIndex = 16
    Me.cmdYcoeGoto.Text = "&Go"
    Me.cmdYcoeGoto.UseVisualStyleBackColor = True
    '
    'tbOEfile
    '
    Me.tbOEfile.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.tbOEfile.Location = New System.Drawing.Point(460, 42)
    Me.tbOEfile.Name = "tbOEfile"
    Me.tbOEfile.ReadOnly = True
    Me.tbOEfile.Size = New System.Drawing.Size(420, 20)
    Me.tbOEfile.TabIndex = 17
    '
    'frmTagSync
    '
    Me.AcceptButton = Me.cmdYes
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(895, 492)
    Me.Controls.Add(Me.tbOEfile)
    Me.Controls.Add(Me.cmdYcoeGoto)
    Me.Controls.Add(Me.tbYcoeLine)
    Me.Controls.Add(Me.Label4)
    Me.Controls.Add(Me.cmdOEtagSearch)
    Me.Controls.Add(Me.cmdYcoeSearch)
    Me.Controls.Add(Me.cmdYcoeNextLine)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Controls.Add(Me.cmdYcoeRight)
    Me.Controls.Add(Me.cmdOEtagRight)
    Me.Controls.Add(Me.cmdYcoeLeft)
    Me.Controls.Add(Me.cmdOEtagLeft)
    Me.Controls.Add(Me.wbYCOE)
    Me.Controls.Add(Me.Label3)
    Me.Controls.Add(Me.wbOEtagged)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.TableLayoutPanel1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmTagSync"
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
    Me.Text = "Tagger synchronisation"
    Me.TableLayoutPanel1.ResumeLayout(False)
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
  Friend WithEvents cmdYes As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents wbOEtagged As System.Windows.Forms.WebBrowser
  Friend WithEvents wbYCOE As System.Windows.Forms.WebBrowser
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents cmdNo As System.Windows.Forms.Button
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents cmdOEtagLeft As System.Windows.Forms.Button
  Friend WithEvents cmdYcoeLeft As System.Windows.Forms.Button
  Friend WithEvents cmdOEtagRight As System.Windows.Forms.Button
  Friend WithEvents cmdYcoeRight As System.Windows.Forms.Button
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents cmdYcoeNextLine As System.Windows.Forms.Button
  Friend WithEvents cmdYcoeSearch As System.Windows.Forms.Button
  Friend WithEvents cmdOEtagSearch As System.Windows.Forms.Button
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents tbYcoeLine As System.Windows.Forms.TextBox
  Friend WithEvents cmdYcoeGoto As System.Windows.Forms.Button
  Friend WithEvents tbOEfile As System.Windows.Forms.TextBox

End Class
