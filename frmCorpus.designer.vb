<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCorpus
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
    Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCorpus))
    Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
    Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
    Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar
    Me.gbFeat = New System.Windows.Forms.GroupBox
    Me.tbDbFeat12 = New System.Windows.Forms.TextBox
    Me.lbDbFeat12 = New System.Windows.Forms.Label
    Me.tbDbFeat11 = New System.Windows.Forms.TextBox
    Me.lbDbFeat11 = New System.Windows.Forms.Label
    Me.tbDbFeat10 = New System.Windows.Forms.TextBox
    Me.lbDbFeat10 = New System.Windows.Forms.Label
    Me.tbDbFeat9 = New System.Windows.Forms.TextBox
    Me.lbDbFeat9 = New System.Windows.Forms.Label
    Me.tbDbFeat8 = New System.Windows.Forms.TextBox
    Me.lbDbFeat8 = New System.Windows.Forms.Label
    Me.cboStatus = New System.Windows.Forms.ComboBox
    Me.Label2 = New System.Windows.Forms.Label
    Me.tbCat = New System.Windows.Forms.TextBox
    Me.Label1 = New System.Windows.Forms.Label
    Me.tbDbFeat7 = New System.Windows.Forms.TextBox
    Me.lbDbFeat7 = New System.Windows.Forms.Label
    Me.tbDbFeat6 = New System.Windows.Forms.TextBox
    Me.lbDbFeat6 = New System.Windows.Forms.Label
    Me.tbDbFeat5 = New System.Windows.Forms.TextBox
    Me.lbDbFeat5 = New System.Windows.Forms.Label
    Me.tbDbFeat4 = New System.Windows.Forms.TextBox
    Me.lbDbFeat4 = New System.Windows.Forms.Label
    Me.tbDbFeat3 = New System.Windows.Forms.TextBox
    Me.lbDbFeat3 = New System.Windows.Forms.Label
    Me.tbDbFeat2 = New System.Windows.Forms.TextBox
    Me.lbDbFeat2 = New System.Windows.Forms.Label
    Me.tbDbFeat1 = New System.Windows.Forms.TextBox
    Me.lbDbFeat1 = New System.Windows.Forms.Label
    Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
    Me.GroupBox3 = New System.Windows.Forms.GroupBox
    Me.tbDbNotes = New System.Windows.Forms.RichTextBox
    Me.GroupBox2 = New System.Windows.Forms.GroupBox
    Me.tbDbToNotes = New System.Windows.Forms.RichTextBox
    Me.gbFeatTo = New System.Windows.Forms.GroupBox
    Me.tbDbToFeat12 = New System.Windows.Forms.TextBox
    Me.lbDbToFeat12 = New System.Windows.Forms.Label
    Me.tbDbToFeat11 = New System.Windows.Forms.TextBox
    Me.lbDbToFeat11 = New System.Windows.Forms.Label
    Me.tbDbToFeat10 = New System.Windows.Forms.TextBox
    Me.lbDbToFeat10 = New System.Windows.Forms.Label
    Me.tbDbToFeat9 = New System.Windows.Forms.TextBox
    Me.lbDbToFeat9 = New System.Windows.Forms.Label
    Me.tbDbToFeat8 = New System.Windows.Forms.TextBox
    Me.lbDbToFeat8 = New System.Windows.Forms.Label
    Me.cboStatusTo = New System.Windows.Forms.ComboBox
    Me.Label3 = New System.Windows.Forms.Label
    Me.tbCatTo = New System.Windows.Forms.TextBox
    Me.Label4 = New System.Windows.Forms.Label
    Me.tbDbToFeat7 = New System.Windows.Forms.TextBox
    Me.lbDbToFeat7 = New System.Windows.Forms.Label
    Me.tbDbToFeat6 = New System.Windows.Forms.TextBox
    Me.lbDbToFeat6 = New System.Windows.Forms.Label
    Me.tbDbToFeat5 = New System.Windows.Forms.TextBox
    Me.lbDbToFeat5 = New System.Windows.Forms.Label
    Me.tbDbToFeat4 = New System.Windows.Forms.TextBox
    Me.lbDbToFeat4 = New System.Windows.Forms.Label
    Me.tbDbToFeat3 = New System.Windows.Forms.TextBox
    Me.lbDbToFeat3 = New System.Windows.Forms.Label
    Me.tbDbToFeat2 = New System.Windows.Forms.TextBox
    Me.lbDbToFeat2 = New System.Windows.Forms.Label
    Me.tbDbToFeat1 = New System.Windows.Forms.TextBox
    Me.lbDbToFeat1 = New System.Windows.Forms.Label
    Me.cmdStep = New System.Windows.Forms.Button
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.tbInstructions = New System.Windows.Forms.RichTextBox
    Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
    Me.tmeFill = New System.Windows.Forms.Timer(Me.components)
    Me.cmdAll = New System.Windows.Forms.Button
    Me.cmdClear = New System.Windows.Forms.Button
    Me.Label5 = New System.Windows.Forms.Label
    Me.StatusStrip1.SuspendLayout()
    Me.gbFeat.SuspendLayout()
    Me.SplitContainer1.Panel1.SuspendLayout()
    Me.SplitContainer1.Panel2.SuspendLayout()
    Me.SplitContainer1.SuspendLayout()
    Me.GroupBox3.SuspendLayout()
    Me.GroupBox2.SuspendLayout()
    Me.gbFeatTo.SuspendLayout()
    Me.SuspendLayout()
    '
    'StatusStrip1
    '
    Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripProgressBar1})
    Me.StatusStrip1.Location = New System.Drawing.Point(0, 667)
    Me.StatusStrip1.Name = "StatusStrip1"
    Me.StatusStrip1.Size = New System.Drawing.Size(838, 22)
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
    Me.ToolStripProgressBar1.Size = New System.Drawing.Size(300, 16)
    '
    'gbFeat
    '
    Me.gbFeat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.gbFeat.Controls.Add(Me.tbDbFeat12)
    Me.gbFeat.Controls.Add(Me.lbDbFeat12)
    Me.gbFeat.Controls.Add(Me.tbDbFeat11)
    Me.gbFeat.Controls.Add(Me.lbDbFeat11)
    Me.gbFeat.Controls.Add(Me.tbDbFeat10)
    Me.gbFeat.Controls.Add(Me.lbDbFeat10)
    Me.gbFeat.Controls.Add(Me.tbDbFeat9)
    Me.gbFeat.Controls.Add(Me.lbDbFeat9)
    Me.gbFeat.Controls.Add(Me.tbDbFeat8)
    Me.gbFeat.Controls.Add(Me.lbDbFeat8)
    Me.gbFeat.Controls.Add(Me.cboStatus)
    Me.gbFeat.Controls.Add(Me.Label2)
    Me.gbFeat.Controls.Add(Me.tbCat)
    Me.gbFeat.Controls.Add(Me.Label1)
    Me.gbFeat.Controls.Add(Me.tbDbFeat7)
    Me.gbFeat.Controls.Add(Me.lbDbFeat7)
    Me.gbFeat.Controls.Add(Me.tbDbFeat6)
    Me.gbFeat.Controls.Add(Me.lbDbFeat6)
    Me.gbFeat.Controls.Add(Me.tbDbFeat5)
    Me.gbFeat.Controls.Add(Me.lbDbFeat5)
    Me.gbFeat.Controls.Add(Me.tbDbFeat4)
    Me.gbFeat.Controls.Add(Me.lbDbFeat4)
    Me.gbFeat.Controls.Add(Me.tbDbFeat3)
    Me.gbFeat.Controls.Add(Me.lbDbFeat3)
    Me.gbFeat.Controls.Add(Me.tbDbFeat2)
    Me.gbFeat.Controls.Add(Me.lbDbFeat2)
    Me.gbFeat.Controls.Add(Me.tbDbFeat1)
    Me.gbFeat.Controls.Add(Me.lbDbFeat1)
    Me.gbFeat.Location = New System.Drawing.Point(3, 3)
    Me.gbFeat.Name = "gbFeat"
    Me.gbFeat.Size = New System.Drawing.Size(394, 406)
    Me.gbFeat.TabIndex = 19
    Me.gbFeat.TabStop = False
    Me.gbFeat.Text = "Selection criteria"
    '
    'tbDbFeat12
    '
    Me.tbDbFeat12.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbFeat12.Location = New System.Drawing.Point(97, 305)
    Me.tbDbFeat12.Name = "tbDbFeat12"
    Me.tbDbFeat12.Size = New System.Drawing.Size(291, 20)
    Me.tbDbFeat12.TabIndex = 26
    '
    'lbDbFeat12
    '
    Me.lbDbFeat12.AutoSize = True
    Me.lbDbFeat12.Location = New System.Drawing.Point(8, 308)
    Me.lbDbFeat12.Name = "lbDbFeat12"
    Me.lbDbFeat12.Size = New System.Drawing.Size(83, 13)
    Me.lbDbFeat12.TabIndex = 27
    Me.lbDbFeat12.Text = "Feature1 Name:"
    '
    'tbDbFeat11
    '
    Me.tbDbFeat11.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbFeat11.Location = New System.Drawing.Point(97, 279)
    Me.tbDbFeat11.Name = "tbDbFeat11"
    Me.tbDbFeat11.Size = New System.Drawing.Size(291, 20)
    Me.tbDbFeat11.TabIndex = 24
    '
    'lbDbFeat11
    '
    Me.lbDbFeat11.AutoSize = True
    Me.lbDbFeat11.Location = New System.Drawing.Point(8, 282)
    Me.lbDbFeat11.Name = "lbDbFeat11"
    Me.lbDbFeat11.Size = New System.Drawing.Size(83, 13)
    Me.lbDbFeat11.TabIndex = 25
    Me.lbDbFeat11.Text = "Feature1 Name:"
    '
    'tbDbFeat10
    '
    Me.tbDbFeat10.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbFeat10.Location = New System.Drawing.Point(97, 253)
    Me.tbDbFeat10.Name = "tbDbFeat10"
    Me.tbDbFeat10.Size = New System.Drawing.Size(291, 20)
    Me.tbDbFeat10.TabIndex = 22
    '
    'lbDbFeat10
    '
    Me.lbDbFeat10.AutoSize = True
    Me.lbDbFeat10.Location = New System.Drawing.Point(8, 256)
    Me.lbDbFeat10.Name = "lbDbFeat10"
    Me.lbDbFeat10.Size = New System.Drawing.Size(83, 13)
    Me.lbDbFeat10.TabIndex = 23
    Me.lbDbFeat10.Text = "Feature1 Name:"
    '
    'tbDbFeat9
    '
    Me.tbDbFeat9.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbFeat9.Location = New System.Drawing.Point(97, 227)
    Me.tbDbFeat9.Name = "tbDbFeat9"
    Me.tbDbFeat9.Size = New System.Drawing.Size(291, 20)
    Me.tbDbFeat9.TabIndex = 20
    '
    'lbDbFeat9
    '
    Me.lbDbFeat9.AutoSize = True
    Me.lbDbFeat9.Location = New System.Drawing.Point(8, 230)
    Me.lbDbFeat9.Name = "lbDbFeat9"
    Me.lbDbFeat9.Size = New System.Drawing.Size(83, 13)
    Me.lbDbFeat9.TabIndex = 21
    Me.lbDbFeat9.Text = "Feature1 Name:"
    '
    'tbDbFeat8
    '
    Me.tbDbFeat8.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbFeat8.Location = New System.Drawing.Point(97, 201)
    Me.tbDbFeat8.Name = "tbDbFeat8"
    Me.tbDbFeat8.Size = New System.Drawing.Size(291, 20)
    Me.tbDbFeat8.TabIndex = 18
    '
    'lbDbFeat8
    '
    Me.lbDbFeat8.AutoSize = True
    Me.lbDbFeat8.Location = New System.Drawing.Point(8, 204)
    Me.lbDbFeat8.Name = "lbDbFeat8"
    Me.lbDbFeat8.Size = New System.Drawing.Size(83, 13)
    Me.lbDbFeat8.TabIndex = 19
    Me.lbDbFeat8.Text = "Feature1 Name:"
    '
    'cboStatus
    '
    Me.cboStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cboStatus.FormattingEnabled = True
    Me.cboStatus.Location = New System.Drawing.Point(97, 379)
    Me.cboStatus.Name = "cboStatus"
    Me.cboStatus.Size = New System.Drawing.Size(291, 21)
    Me.cboStatus.TabIndex = 17
    '
    'Label2
    '
    Me.Label2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(8, 382)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(40, 13)
    Me.Label2.TabIndex = 16
    Me.Label2.Text = "Status:"
    '
    'tbCat
    '
    Me.tbCat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbCat.Location = New System.Drawing.Point(97, 353)
    Me.tbCat.Name = "tbCat"
    Me.tbCat.Size = New System.Drawing.Size(291, 20)
    Me.tbCat.TabIndex = 13
    '
    'Label1
    '
    Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(8, 356)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(52, 13)
    Me.Label1.TabIndex = 14
    Me.Label1.Text = "Category:"
    '
    'tbDbFeat7
    '
    Me.tbDbFeat7.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbFeat7.Location = New System.Drawing.Point(97, 175)
    Me.tbDbFeat7.Name = "tbDbFeat7"
    Me.tbDbFeat7.Size = New System.Drawing.Size(291, 20)
    Me.tbDbFeat7.TabIndex = 9
    '
    'lbDbFeat7
    '
    Me.lbDbFeat7.AutoSize = True
    Me.lbDbFeat7.Location = New System.Drawing.Point(8, 178)
    Me.lbDbFeat7.Name = "lbDbFeat7"
    Me.lbDbFeat7.Size = New System.Drawing.Size(83, 13)
    Me.lbDbFeat7.TabIndex = 12
    Me.lbDbFeat7.Text = "Feature1 Name:"
    '
    'tbDbFeat6
    '
    Me.tbDbFeat6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbFeat6.Location = New System.Drawing.Point(97, 149)
    Me.tbDbFeat6.Name = "tbDbFeat6"
    Me.tbDbFeat6.Size = New System.Drawing.Size(291, 20)
    Me.tbDbFeat6.TabIndex = 8
    '
    'lbDbFeat6
    '
    Me.lbDbFeat6.AutoSize = True
    Me.lbDbFeat6.Location = New System.Drawing.Point(8, 152)
    Me.lbDbFeat6.Name = "lbDbFeat6"
    Me.lbDbFeat6.Size = New System.Drawing.Size(83, 13)
    Me.lbDbFeat6.TabIndex = 10
    Me.lbDbFeat6.Text = "Feature1 Name:"
    '
    'tbDbFeat5
    '
    Me.tbDbFeat5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbFeat5.Location = New System.Drawing.Point(97, 123)
    Me.tbDbFeat5.Name = "tbDbFeat5"
    Me.tbDbFeat5.Size = New System.Drawing.Size(291, 20)
    Me.tbDbFeat5.TabIndex = 7
    '
    'lbDbFeat5
    '
    Me.lbDbFeat5.AutoSize = True
    Me.lbDbFeat5.Location = New System.Drawing.Point(8, 126)
    Me.lbDbFeat5.Name = "lbDbFeat5"
    Me.lbDbFeat5.Size = New System.Drawing.Size(83, 13)
    Me.lbDbFeat5.TabIndex = 8
    Me.lbDbFeat5.Text = "Feature1 Name:"
    '
    'tbDbFeat4
    '
    Me.tbDbFeat4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbFeat4.Location = New System.Drawing.Point(97, 97)
    Me.tbDbFeat4.Name = "tbDbFeat4"
    Me.tbDbFeat4.Size = New System.Drawing.Size(291, 20)
    Me.tbDbFeat4.TabIndex = 6
    '
    'lbDbFeat4
    '
    Me.lbDbFeat4.AutoSize = True
    Me.lbDbFeat4.Location = New System.Drawing.Point(8, 100)
    Me.lbDbFeat4.Name = "lbDbFeat4"
    Me.lbDbFeat4.Size = New System.Drawing.Size(83, 13)
    Me.lbDbFeat4.TabIndex = 6
    Me.lbDbFeat4.Text = "Feature1 Name:"
    '
    'tbDbFeat3
    '
    Me.tbDbFeat3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbFeat3.Location = New System.Drawing.Point(97, 71)
    Me.tbDbFeat3.Name = "tbDbFeat3"
    Me.tbDbFeat3.Size = New System.Drawing.Size(291, 20)
    Me.tbDbFeat3.TabIndex = 5
    '
    'lbDbFeat3
    '
    Me.lbDbFeat3.AutoSize = True
    Me.lbDbFeat3.Location = New System.Drawing.Point(8, 74)
    Me.lbDbFeat3.Name = "lbDbFeat3"
    Me.lbDbFeat3.Size = New System.Drawing.Size(83, 13)
    Me.lbDbFeat3.TabIndex = 4
    Me.lbDbFeat3.Text = "Feature1 Name:"
    '
    'tbDbFeat2
    '
    Me.tbDbFeat2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbFeat2.Location = New System.Drawing.Point(97, 45)
    Me.tbDbFeat2.Name = "tbDbFeat2"
    Me.tbDbFeat2.Size = New System.Drawing.Size(291, 20)
    Me.tbDbFeat2.TabIndex = 4
    '
    'lbDbFeat2
    '
    Me.lbDbFeat2.AutoSize = True
    Me.lbDbFeat2.Location = New System.Drawing.Point(8, 48)
    Me.lbDbFeat2.Name = "lbDbFeat2"
    Me.lbDbFeat2.Size = New System.Drawing.Size(83, 13)
    Me.lbDbFeat2.TabIndex = 2
    Me.lbDbFeat2.Text = "Feature1 Name:"
    '
    'tbDbFeat1
    '
    Me.tbDbFeat1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbFeat1.Location = New System.Drawing.Point(97, 20)
    Me.tbDbFeat1.Name = "tbDbFeat1"
    Me.tbDbFeat1.Size = New System.Drawing.Size(291, 20)
    Me.tbDbFeat1.TabIndex = 3
    '
    'lbDbFeat1
    '
    Me.lbDbFeat1.AutoSize = True
    Me.lbDbFeat1.Location = New System.Drawing.Point(8, 22)
    Me.lbDbFeat1.Name = "lbDbFeat1"
    Me.lbDbFeat1.Size = New System.Drawing.Size(83, 13)
    Me.lbDbFeat1.TabIndex = 0
    Me.lbDbFeat1.Text = "Feature1 Name:"
    '
    'SplitContainer1
    '
    Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.SplitContainer1.Location = New System.Drawing.Point(0, 72)
    Me.SplitContainer1.Name = "SplitContainer1"
    '
    'SplitContainer1.Panel1
    '
    Me.SplitContainer1.Panel1.Controls.Add(Me.GroupBox3)
    Me.SplitContainer1.Panel1.Controls.Add(Me.gbFeat)
    '
    'SplitContainer1.Panel2
    '
    Me.SplitContainer1.Panel2.Controls.Add(Me.GroupBox2)
    Me.SplitContainer1.Panel2.Controls.Add(Me.gbFeatTo)
    Me.SplitContainer1.Size = New System.Drawing.Size(838, 563)
    Me.SplitContainer1.SplitterDistance = 399
    Me.SplitContainer1.TabIndex = 20
    '
    'GroupBox3
    '
    Me.GroupBox3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GroupBox3.Controls.Add(Me.tbDbNotes)
    Me.GroupBox3.Location = New System.Drawing.Point(3, 415)
    Me.GroupBox3.Name = "GroupBox3"
    Me.GroupBox3.Size = New System.Drawing.Size(393, 145)
    Me.GroupBox3.TabIndex = 20
    Me.GroupBox3.TabStop = False
    Me.GroupBox3.Text = "Notes to look for"
    '
    'tbDbNotes
    '
    Me.tbDbNotes.Dock = System.Windows.Forms.DockStyle.Fill
    Me.tbDbNotes.Location = New System.Drawing.Point(3, 16)
    Me.tbDbNotes.Name = "tbDbNotes"
    Me.tbDbNotes.Size = New System.Drawing.Size(387, 126)
    Me.tbDbNotes.TabIndex = 10
    Me.tbDbNotes.Text = ""
    '
    'GroupBox2
    '
    Me.GroupBox2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GroupBox2.Controls.Add(Me.tbDbToNotes)
    Me.GroupBox2.Location = New System.Drawing.Point(3, 415)
    Me.GroupBox2.Name = "GroupBox2"
    Me.GroupBox2.Size = New System.Drawing.Size(429, 145)
    Me.GroupBox2.TabIndex = 21
    Me.GroupBox2.TabStop = False
    Me.GroupBox2.Text = "Fixed text to REPLACE the notes in each record"
    '
    'tbDbToNotes
    '
    Me.tbDbToNotes.Dock = System.Windows.Forms.DockStyle.Fill
    Me.tbDbToNotes.Location = New System.Drawing.Point(3, 16)
    Me.tbDbToNotes.Name = "tbDbToNotes"
    Me.tbDbToNotes.Size = New System.Drawing.Size(423, 126)
    Me.tbDbToNotes.TabIndex = 10
    Me.tbDbToNotes.Text = ""
    '
    'gbFeatTo
    '
    Me.gbFeatTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.gbFeatTo.Controls.Add(Me.tbDbToFeat12)
    Me.gbFeatTo.Controls.Add(Me.lbDbToFeat12)
    Me.gbFeatTo.Controls.Add(Me.tbDbToFeat11)
    Me.gbFeatTo.Controls.Add(Me.lbDbToFeat11)
    Me.gbFeatTo.Controls.Add(Me.tbDbToFeat10)
    Me.gbFeatTo.Controls.Add(Me.lbDbToFeat10)
    Me.gbFeatTo.Controls.Add(Me.tbDbToFeat9)
    Me.gbFeatTo.Controls.Add(Me.lbDbToFeat9)
    Me.gbFeatTo.Controls.Add(Me.tbDbToFeat8)
    Me.gbFeatTo.Controls.Add(Me.lbDbToFeat8)
    Me.gbFeatTo.Controls.Add(Me.cboStatusTo)
    Me.gbFeatTo.Controls.Add(Me.Label3)
    Me.gbFeatTo.Controls.Add(Me.tbCatTo)
    Me.gbFeatTo.Controls.Add(Me.Label4)
    Me.gbFeatTo.Controls.Add(Me.tbDbToFeat7)
    Me.gbFeatTo.Controls.Add(Me.lbDbToFeat7)
    Me.gbFeatTo.Controls.Add(Me.tbDbToFeat6)
    Me.gbFeatTo.Controls.Add(Me.lbDbToFeat6)
    Me.gbFeatTo.Controls.Add(Me.tbDbToFeat5)
    Me.gbFeatTo.Controls.Add(Me.lbDbToFeat5)
    Me.gbFeatTo.Controls.Add(Me.tbDbToFeat4)
    Me.gbFeatTo.Controls.Add(Me.lbDbToFeat4)
    Me.gbFeatTo.Controls.Add(Me.tbDbToFeat3)
    Me.gbFeatTo.Controls.Add(Me.lbDbToFeat3)
    Me.gbFeatTo.Controls.Add(Me.tbDbToFeat2)
    Me.gbFeatTo.Controls.Add(Me.lbDbToFeat2)
    Me.gbFeatTo.Controls.Add(Me.tbDbToFeat1)
    Me.gbFeatTo.Controls.Add(Me.lbDbToFeat1)
    Me.gbFeatTo.Location = New System.Drawing.Point(3, 4)
    Me.gbFeatTo.Name = "gbFeatTo"
    Me.gbFeatTo.Size = New System.Drawing.Size(429, 405)
    Me.gbFeatTo.TabIndex = 20
    Me.gbFeatTo.TabStop = False
    Me.gbFeatTo.Text = "Copy the following features to each record"
    '
    'tbDbToFeat12
    '
    Me.tbDbToFeat12.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbToFeat12.Location = New System.Drawing.Point(97, 304)
    Me.tbDbToFeat12.Name = "tbDbToFeat12"
    Me.tbDbToFeat12.Size = New System.Drawing.Size(326, 20)
    Me.tbDbToFeat12.TabIndex = 30
    '
    'lbDbToFeat12
    '
    Me.lbDbToFeat12.AutoSize = True
    Me.lbDbToFeat12.Location = New System.Drawing.Point(8, 307)
    Me.lbDbToFeat12.Name = "lbDbToFeat12"
    Me.lbDbToFeat12.Size = New System.Drawing.Size(83, 13)
    Me.lbDbToFeat12.TabIndex = 31
    Me.lbDbToFeat12.Text = "Feature1 Name:"
    '
    'tbDbToFeat11
    '
    Me.tbDbToFeat11.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbToFeat11.Location = New System.Drawing.Point(97, 278)
    Me.tbDbToFeat11.Name = "tbDbToFeat11"
    Me.tbDbToFeat11.Size = New System.Drawing.Size(326, 20)
    Me.tbDbToFeat11.TabIndex = 28
    '
    'lbDbToFeat11
    '
    Me.lbDbToFeat11.AutoSize = True
    Me.lbDbToFeat11.Location = New System.Drawing.Point(8, 281)
    Me.lbDbToFeat11.Name = "lbDbToFeat11"
    Me.lbDbToFeat11.Size = New System.Drawing.Size(83, 13)
    Me.lbDbToFeat11.TabIndex = 29
    Me.lbDbToFeat11.Text = "Feature1 Name:"
    '
    'tbDbToFeat10
    '
    Me.tbDbToFeat10.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbToFeat10.Location = New System.Drawing.Point(97, 252)
    Me.tbDbToFeat10.Name = "tbDbToFeat10"
    Me.tbDbToFeat10.Size = New System.Drawing.Size(326, 20)
    Me.tbDbToFeat10.TabIndex = 26
    '
    'lbDbToFeat10
    '
    Me.lbDbToFeat10.AutoSize = True
    Me.lbDbToFeat10.Location = New System.Drawing.Point(8, 255)
    Me.lbDbToFeat10.Name = "lbDbToFeat10"
    Me.lbDbToFeat10.Size = New System.Drawing.Size(83, 13)
    Me.lbDbToFeat10.TabIndex = 27
    Me.lbDbToFeat10.Text = "Feature1 Name:"
    '
    'tbDbToFeat9
    '
    Me.tbDbToFeat9.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbToFeat9.Location = New System.Drawing.Point(97, 226)
    Me.tbDbToFeat9.Name = "tbDbToFeat9"
    Me.tbDbToFeat9.Size = New System.Drawing.Size(326, 20)
    Me.tbDbToFeat9.TabIndex = 24
    '
    'lbDbToFeat9
    '
    Me.lbDbToFeat9.AutoSize = True
    Me.lbDbToFeat9.Location = New System.Drawing.Point(8, 229)
    Me.lbDbToFeat9.Name = "lbDbToFeat9"
    Me.lbDbToFeat9.Size = New System.Drawing.Size(83, 13)
    Me.lbDbToFeat9.TabIndex = 25
    Me.lbDbToFeat9.Text = "Feature1 Name:"
    '
    'tbDbToFeat8
    '
    Me.tbDbToFeat8.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbToFeat8.Location = New System.Drawing.Point(97, 201)
    Me.tbDbToFeat8.Name = "tbDbToFeat8"
    Me.tbDbToFeat8.Size = New System.Drawing.Size(326, 20)
    Me.tbDbToFeat8.TabIndex = 22
    '
    'lbDbToFeat8
    '
    Me.lbDbToFeat8.AutoSize = True
    Me.lbDbToFeat8.Location = New System.Drawing.Point(8, 204)
    Me.lbDbToFeat8.Name = "lbDbToFeat8"
    Me.lbDbToFeat8.Size = New System.Drawing.Size(83, 13)
    Me.lbDbToFeat8.TabIndex = 23
    Me.lbDbToFeat8.Text = "Feature1 Name:"
    '
    'cboStatusTo
    '
    Me.cboStatusTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cboStatusTo.FormattingEnabled = True
    Me.cboStatusTo.Location = New System.Drawing.Point(97, 378)
    Me.cboStatusTo.Name = "cboStatusTo"
    Me.cboStatusTo.Size = New System.Drawing.Size(326, 21)
    Me.cboStatusTo.TabIndex = 21
    '
    'Label3
    '
    Me.Label3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(8, 381)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(40, 13)
    Me.Label3.TabIndex = 20
    Me.Label3.Text = "Status:"
    '
    'tbCatTo
    '
    Me.tbCatTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbCatTo.Location = New System.Drawing.Point(97, 352)
    Me.tbCatTo.Name = "tbCatTo"
    Me.tbCatTo.Size = New System.Drawing.Size(326, 20)
    Me.tbCatTo.TabIndex = 18
    '
    'Label4
    '
    Me.Label4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(8, 355)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(52, 13)
    Me.Label4.TabIndex = 19
    Me.Label4.Text = "Category:"
    '
    'tbDbToFeat7
    '
    Me.tbDbToFeat7.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbToFeat7.Location = New System.Drawing.Point(97, 175)
    Me.tbDbToFeat7.Name = "tbDbToFeat7"
    Me.tbDbToFeat7.Size = New System.Drawing.Size(326, 20)
    Me.tbDbToFeat7.TabIndex = 9
    '
    'lbDbToFeat7
    '
    Me.lbDbToFeat7.AutoSize = True
    Me.lbDbToFeat7.Location = New System.Drawing.Point(8, 178)
    Me.lbDbToFeat7.Name = "lbDbToFeat7"
    Me.lbDbToFeat7.Size = New System.Drawing.Size(83, 13)
    Me.lbDbToFeat7.TabIndex = 12
    Me.lbDbToFeat7.Text = "Feature1 Name:"
    '
    'tbDbToFeat6
    '
    Me.tbDbToFeat6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbToFeat6.Location = New System.Drawing.Point(97, 149)
    Me.tbDbToFeat6.Name = "tbDbToFeat6"
    Me.tbDbToFeat6.Size = New System.Drawing.Size(326, 20)
    Me.tbDbToFeat6.TabIndex = 8
    '
    'lbDbToFeat6
    '
    Me.lbDbToFeat6.AutoSize = True
    Me.lbDbToFeat6.Location = New System.Drawing.Point(8, 152)
    Me.lbDbToFeat6.Name = "lbDbToFeat6"
    Me.lbDbToFeat6.Size = New System.Drawing.Size(83, 13)
    Me.lbDbToFeat6.TabIndex = 10
    Me.lbDbToFeat6.Text = "Feature1 Name:"
    '
    'tbDbToFeat5
    '
    Me.tbDbToFeat5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbToFeat5.Location = New System.Drawing.Point(97, 123)
    Me.tbDbToFeat5.Name = "tbDbToFeat5"
    Me.tbDbToFeat5.Size = New System.Drawing.Size(326, 20)
    Me.tbDbToFeat5.TabIndex = 7
    '
    'lbDbToFeat5
    '
    Me.lbDbToFeat5.AutoSize = True
    Me.lbDbToFeat5.Location = New System.Drawing.Point(8, 126)
    Me.lbDbToFeat5.Name = "lbDbToFeat5"
    Me.lbDbToFeat5.Size = New System.Drawing.Size(83, 13)
    Me.lbDbToFeat5.TabIndex = 8
    Me.lbDbToFeat5.Text = "Feature1 Name:"
    '
    'tbDbToFeat4
    '
    Me.tbDbToFeat4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbToFeat4.Location = New System.Drawing.Point(97, 97)
    Me.tbDbToFeat4.Name = "tbDbToFeat4"
    Me.tbDbToFeat4.Size = New System.Drawing.Size(326, 20)
    Me.tbDbToFeat4.TabIndex = 6
    '
    'lbDbToFeat4
    '
    Me.lbDbToFeat4.AutoSize = True
    Me.lbDbToFeat4.Location = New System.Drawing.Point(8, 100)
    Me.lbDbToFeat4.Name = "lbDbToFeat4"
    Me.lbDbToFeat4.Size = New System.Drawing.Size(83, 13)
    Me.lbDbToFeat4.TabIndex = 6
    Me.lbDbToFeat4.Text = "Feature1 Name:"
    '
    'tbDbToFeat3
    '
    Me.tbDbToFeat3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbToFeat3.Location = New System.Drawing.Point(97, 71)
    Me.tbDbToFeat3.Name = "tbDbToFeat3"
    Me.tbDbToFeat3.Size = New System.Drawing.Size(326, 20)
    Me.tbDbToFeat3.TabIndex = 5
    '
    'lbDbToFeat3
    '
    Me.lbDbToFeat3.AutoSize = True
    Me.lbDbToFeat3.Location = New System.Drawing.Point(8, 74)
    Me.lbDbToFeat3.Name = "lbDbToFeat3"
    Me.lbDbToFeat3.Size = New System.Drawing.Size(83, 13)
    Me.lbDbToFeat3.TabIndex = 4
    Me.lbDbToFeat3.Text = "Feature1 Name:"
    '
    'tbDbToFeat2
    '
    Me.tbDbToFeat2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbToFeat2.Location = New System.Drawing.Point(97, 45)
    Me.tbDbToFeat2.Name = "tbDbToFeat2"
    Me.tbDbToFeat2.Size = New System.Drawing.Size(326, 20)
    Me.tbDbToFeat2.TabIndex = 4
    '
    'lbDbToFeat2
    '
    Me.lbDbToFeat2.AutoSize = True
    Me.lbDbToFeat2.Location = New System.Drawing.Point(8, 48)
    Me.lbDbToFeat2.Name = "lbDbToFeat2"
    Me.lbDbToFeat2.Size = New System.Drawing.Size(83, 13)
    Me.lbDbToFeat2.TabIndex = 2
    Me.lbDbToFeat2.Text = "Feature1 Name:"
    '
    'tbDbToFeat1
    '
    Me.tbDbToFeat1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.tbDbToFeat1.Location = New System.Drawing.Point(97, 21)
    Me.tbDbToFeat1.Name = "tbDbToFeat1"
    Me.tbDbToFeat1.Size = New System.Drawing.Size(326, 20)
    Me.tbDbToFeat1.TabIndex = 3
    '
    'lbDbToFeat1
    '
    Me.lbDbToFeat1.AutoSize = True
    Me.lbDbToFeat1.Location = New System.Drawing.Point(8, 22)
    Me.lbDbToFeat1.Name = "lbDbToFeat1"
    Me.lbDbToFeat1.Size = New System.Drawing.Size(83, 13)
    Me.lbDbToFeat1.TabIndex = 0
    Me.lbDbToFeat1.Text = "Feature1 Name:"
    '
    'cmdStep
    '
    Me.cmdStep.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdStep.Location = New System.Drawing.Point(679, 641)
    Me.cmdStep.Name = "cmdStep"
    Me.cmdStep.Size = New System.Drawing.Size(75, 23)
    Me.cmdStep.TabIndex = 21
    Me.cmdStep.Text = "&Step"
    Me.cmdStep.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancel.Location = New System.Drawing.Point(598, 641)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 22
    Me.cmdCancel.Text = "&Cancel"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'tbInstructions
    '
    Me.tbInstructions.BackColor = System.Drawing.SystemColors.InactiveBorder
    Me.tbInstructions.BorderStyle = System.Windows.Forms.BorderStyle.None
    Me.tbInstructions.ForeColor = System.Drawing.Color.Maroon
    Me.tbInstructions.Location = New System.Drawing.Point(3, 12)
    Me.tbInstructions.Name = "tbInstructions"
    Me.tbInstructions.Size = New System.Drawing.Size(413, 54)
    Me.tbInstructions.TabIndex = 23
    Me.tbInstructions.Text = "Instructions:" & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(10) & "1. Enter (or delete) the selection criteria on the left (or leave " & _
        "blank)" & Global.Microsoft.VisualBasic.ChrW(10) & "2. Enter the feature values to be added to each record satisfying the sel" & _
        "ection criteria" & Global.Microsoft.VisualBasic.ChrW(10)
    '
    'Timer1
    '
    '
    'tmeFill
    '
    '
    'cmdAll
    '
    Me.cmdAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdAll.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdAll.Location = New System.Drawing.Point(760, 641)
    Me.cmdAll.Name = "cmdAll"
    Me.cmdAll.Size = New System.Drawing.Size(75, 23)
    Me.cmdAll.TabIndex = 24
    Me.cmdAll.Text = "&All"
    Me.cmdAll.UseVisualStyleBackColor = True
    '
    'cmdClear
    '
    Me.cmdClear.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.cmdClear.Location = New System.Drawing.Point(6, 641)
    Me.cmdClear.Name = "cmdClear"
    Me.cmdClear.Size = New System.Drawing.Size(75, 23)
    Me.cmdClear.TabIndex = 25
    Me.cmdClear.Text = "C&lear"
    Me.cmdClear.UseVisualStyleBackColor = True
    '
    'Label5
    '
    Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Label5.AutoSize = True
    Me.Label5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
    Me.Label5.Location = New System.Drawing.Point(618, 12)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(220, 13)
    Me.Label5.TabIndex = 26
    Me.Label5.Text = "(only the first 10 features are accessible here)"
    Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
    '
    'frmCorpus
    '
    Me.AcceptButton = Me.cmdStep
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.cmdCancel
    Me.ClientSize = New System.Drawing.Size(838, 689)
    Me.Controls.Add(Me.Label5)
    Me.Controls.Add(Me.cmdClear)
    Me.Controls.Add(Me.cmdAll)
    Me.Controls.Add(Me.tbInstructions)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.cmdStep)
    Me.Controls.Add(Me.SplitContainer1)
    Me.Controls.Add(Me.StatusStrip1)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmCorpus"
    Me.Text = "Copy annotation..."
    Me.StatusStrip1.ResumeLayout(False)
    Me.StatusStrip1.PerformLayout()
    Me.gbFeat.ResumeLayout(False)
    Me.gbFeat.PerformLayout()
    Me.SplitContainer1.Panel1.ResumeLayout(False)
    Me.SplitContainer1.Panel2.ResumeLayout(False)
    Me.SplitContainer1.ResumeLayout(False)
    Me.GroupBox3.ResumeLayout(False)
    Me.GroupBox2.ResumeLayout(False)
    Me.gbFeatTo.ResumeLayout(False)
    Me.gbFeatTo.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
  Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
  Friend WithEvents ToolStripProgressBar1 As System.Windows.Forms.ToolStripProgressBar
  Friend WithEvents gbFeat As System.Windows.Forms.GroupBox
  Friend WithEvents tbDbFeat7 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbFeat7 As System.Windows.Forms.Label
  Friend WithEvents tbDbFeat6 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbFeat6 As System.Windows.Forms.Label
  Friend WithEvents tbDbFeat5 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbFeat5 As System.Windows.Forms.Label
  Friend WithEvents tbDbFeat4 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbFeat4 As System.Windows.Forms.Label
  Friend WithEvents tbDbFeat3 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbFeat3 As System.Windows.Forms.Label
  Friend WithEvents tbDbFeat2 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbFeat2 As System.Windows.Forms.Label
  Friend WithEvents tbDbFeat1 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbFeat1 As System.Windows.Forms.Label
  Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
  Friend WithEvents gbFeatTo As System.Windows.Forms.GroupBox
  Friend WithEvents tbDbToFeat7 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbToFeat7 As System.Windows.Forms.Label
  Friend WithEvents tbDbToFeat6 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbToFeat6 As System.Windows.Forms.Label
  Friend WithEvents tbDbToFeat5 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbToFeat5 As System.Windows.Forms.Label
  Friend WithEvents tbDbToFeat4 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbToFeat4 As System.Windows.Forms.Label
  Friend WithEvents tbDbToFeat3 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbToFeat3 As System.Windows.Forms.Label
  Friend WithEvents tbDbToFeat2 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbToFeat2 As System.Windows.Forms.Label
  Friend WithEvents tbDbToFeat1 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbToFeat1 As System.Windows.Forms.Label
  Friend WithEvents cmdStep As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents tbInstructions As System.Windows.Forms.RichTextBox
  Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
  Friend WithEvents tbDbNotes As System.Windows.Forms.RichTextBox
  Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
  Friend WithEvents tbDbToNotes As System.Windows.Forms.RichTextBox
  Friend WithEvents Timer1 As System.Windows.Forms.Timer
  Friend WithEvents tmeFill As System.Windows.Forms.Timer
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents tbCat As System.Windows.Forms.TextBox
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents cboStatus As System.Windows.Forms.ComboBox
  Friend WithEvents cboStatusTo As System.Windows.Forms.ComboBox
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents tbCatTo As System.Windows.Forms.TextBox
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents cmdAll As System.Windows.Forms.Button
  Friend WithEvents tbDbFeat10 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbFeat10 As System.Windows.Forms.Label
  Friend WithEvents tbDbFeat9 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbFeat9 As System.Windows.Forms.Label
  Friend WithEvents tbDbFeat8 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbFeat8 As System.Windows.Forms.Label
  Friend WithEvents tbDbToFeat10 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbToFeat10 As System.Windows.Forms.Label
  Friend WithEvents tbDbToFeat9 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbToFeat9 As System.Windows.Forms.Label
  Friend WithEvents tbDbToFeat8 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbToFeat8 As System.Windows.Forms.Label
  Friend WithEvents cmdClear As System.Windows.Forms.Button
  Friend WithEvents tbDbFeat12 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbFeat12 As System.Windows.Forms.Label
  Friend WithEvents tbDbFeat11 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbFeat11 As System.Windows.Forms.Label
  Friend WithEvents tbDbToFeat12 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbToFeat12 As System.Windows.Forms.Label
  Friend WithEvents tbDbToFeat11 As System.Windows.Forms.TextBox
  Friend WithEvents lbDbToFeat11 As System.Windows.Forms.Label
  Friend WithEvents Label5 As System.Windows.Forms.Label
End Class
