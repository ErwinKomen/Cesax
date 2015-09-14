Public Class DgvHandle
  ' =============================== LOCAL TYPES ==========================================
  Private Structure CtrInfo
    Dim Ctl As Control  ' A pointer to the actual control
    Dim Field As String ' The field associated with it in the datatable
    Dim Type As String  ' The type of the control
  End Structure
  ' =============================== LOCAL VARIABLES ======================================
  Private WithEvents dgvLocal As DataGridView = Nothing   ' The datagridview we are servicing
  Private WithEvents tdlLocal As DataSet = Nothing        ' The dataset we are working with
  Private dlgSave As New SaveFileDialog
  Private strParentTable As String = ""   ' THe parent table's name (if any)
  Private bsLocal As New BindingSource    ' The bindingsource we are to be using...
  ' Private dvLocal As New DataView         ' The dataview source we are to be using...
  Private strTable As String = ""         ' The name of the datatable
  Private strId As String = ""            ' Name of the ID field
  Private Shared strSort As String = ""          ' Sort order (if defined)
  Private strChanged As String = ""       ' Possible name of a "Changed" item
  Private strCreated As String = ""       ' Possible name of a "Created" item
  Private strTbChanged As String = ""     ' Possible name of a textbox where the "Changed" time should be shown
  Private arCtrInfo() As CtrInfo          ' Array with names and types of controls "bound" to me
  Private intSelectedId As Integer = -1   ' The ID of the selected item
  Private intControls As Integer          ' The number of controls so far
  Private bInit As Boolean = False        ' Initialisation flag
  Private bDirty As Boolean = False       ' Dirty flag
  Private bEdit As Boolean = False        ' Whether editing is enabled
  Private Shared bIsSelecting As Boolean = False ' Whether selection event takes place
  ' Private bIsSelecting As Boolean = False ' Whether selection event takes place
  Private bIsDgv As Boolean = True        ' Whether this concerns handling an actual DGV or not
  Private bEdtOk As Boolean = True        ' Whether the editor may be used (read by "ReadIntoEditor"
  Private Shared bRefresh As Boolean = False      ' Whether we are refreshing
  Private objForm As Form = Nothing               ' A reference to the calling form
  Private dtrCurrent As DataRow                   ' The currently selected datarow
  Private Shared bMouseClick As Boolean = False   ' Whether a mouse has actually clicked
  Private loc_strColumns As String = ""           ' columns to be shown
  Public Delegate Sub BeforeSelChanged(ByVal intSelId As Integer)
  Private opBeforeSelChanged As BeforeSelChanged = Nothing
  ' ------------------------------------------------------------------------------------
  ' Name:   EdtOk
  ' Goal:   Set or reset EditorOk flag for this editor
  ' History:
  ' 27-11-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Property EdtOk() As Boolean
    Get
      EdtOk = bEdtOk
    End Get
    Set(ByVal value As Boolean)
      bEdtOk = value
    End Set
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   IsSelecting
  ' Goal:   Show to the caller, that we are busy with selection
  ' History:
  ' 27-11-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public ReadOnly Property IsSelecting() As Boolean
    Get
      IsSelecting = bIsSelecting
    End Get
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   ParentTable
  ' Goal:   Allow user to set a parent row 
  ' History:
  ' 28-11-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public WriteOnly Property ParentTable() As String
    Set(ByVal value As String)
      ' Get the first row ...
      strParentTable = value
    End Set
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   SelRow
  ' Goal:   Return the currently selected datarow 
  ' History:
  ' 21-12-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public ReadOnly Property SelRow() As DataRow
    Get
      SelRow = dtrCurrent
    End Get
  End Property
  ' ------------------------------------------------------------------------------------
  ' Name:   Filter
  ' Goal:   Set the filter in the BindingSource
  ' History:
  ' 01-12-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Property Filter(Optional ByVal bNoComplaint As Boolean = False) As String
    Set(ByVal value As String)
      Try
        ' Do we have a [bsLocal]?
        If (Not bsLocal Is Nothing) Then
          ' Set the filter by which the datatable is divided
          bsLocal.Filter = ""
          bsLocal.Filter = value
        End If
      Catch ex As Exception
        ' Show the error
        If (Not bNoComplaint) Then HandleErr("DgvHandle/Filter error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      End Try
    End Set
    Get
      Try
        ' Do we actually have a [bsLocal]?
        If (Not bsLocal Is Nothing) Then
          If (bsLocal.Filter Is Nothing) Then
            Return ""
          Else
            ' Try to return it immediately
            Return bsLocal.Filter.ToString
          End If
        Else
          ' There is no filter specified
          Return ""
        End If
      Catch ex As Exception
        ' Show the error
        If (Not bNoComplaint) Then HandleErr("DgvHandle/Filter error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
        ' Return empty-handed
        Return ""
      End Try
    End Get
  End Property
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  Init
  ' Goal :  Get the datagridview and the bindingsource, set them up etc.
  ' Parameters:
  '         tdlThis         - Dataset where the table resides
  '         strTableName    - Name of the table in the dataset
  '         strIdName       - Name of the numerical ID field
  '         strSortOrder    - Optional sort order
  '         strColumns      - columns to be shown, separated by ;
  '         strChangedItem  - Name of the field where "Changed" time should be placed (optional)
  '         strCreatedItem  - Name of the field where "Created" time should be placed (optional)
  '         dgvThis         - the datagridview to be used
  ' History:
  ' 27-11-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function Init(ByRef frmThis As Form, ByRef tdlThis As DataSet, _
      ByVal strTableName As String, ByVal strIdName As String, ByVal strSortOrder As String, _
      ByVal strColumns As String, ByVal strChangedTb As String, ByVal strChangedItem As String, _
      ByVal strCreatedItem As String, ByRef dgvThis As DataGridView, _
      ByRef opBefSelChgSub As BeforeSelChanged) As Boolean
    Dim intControls As Integer = 0  ' Number of controls defined

    Try
      ' Trim the table name
      strTableName = Trim(strTableName)
      ' Are we getting real information?
      If (frmThis Is Nothing) OrElse (tdlThis Is Nothing) OrElse (strTableName = "") Then
        ' Unable to initialize
        Return False
      End If
      ' Are we getting real information?
      If (dgvThis Is Nothing) AndAlso (strIdName = "") Then
        ' This is not a DGV proper type
        bIsDgv = False
      ElseIf (dgvThis Is Nothing) OrElse (strIdName = "") OrElse _
        (strColumns = "") Then
        ' Unable to initialize
        Return False
      Else
        ' This is a proper DGV type
        bIsDgv = True
      End If
      ' Set up the control array
      ReDim arCtrInfo(0) : intControls = 0
      ' Set the local copies
      objForm = frmThis
      dgvLocal = dgvThis
      tdlLocal = tdlThis
      strTable = strTableName
      strSort = Trim(strSortOrder)
      strChanged = Trim(strChangedItem)
      strCreated = Trim(strCreatedItem)
      strTbChanged = Trim(strChangedTb)
      strId = Trim(strIdName)
      opBeforeSelChanged = opBefSelChgSub
      ' Is there any content in the periodinformation already?
      If (tdlLocal.Tables(strTableName) Is Nothing) Then
        ' Create this table - it is needed!!
        tdlLocal.Tables.Add(strTableName)
      End If
      ' Is this of type DGV?
      If (bIsDgv) Then
        ' Set the bindingsource for the dgv
        With bsLocal
          .DataSource = tdlLocal
          .DataMember = strTableName
          .Sort = strSort
        End With
        '' Alternative bindingsource: dataview
        'With dvLocal
        '  .Table = tdlLocal.Tables(strTableName)
        '  .Sort = strSort
        'End With
        ' Set the output for the Period Editor DataGridView
        SetupDgv(dgvLocal, bsLocal)
        ' SetupDgv(dgvLocal, dvLocal)
        ' Columns should have been supplied
        '  Debug.Print(tdlLocal.Tables("meta").Rows.Count)
        ' Make sure only the correct columns are visible
        SetDgvColumns(dgvLocal, False, strColumns)
        loc_strColumns = strColumns
      End If
      ' Reset the dirty flag
      DoEditDirty(False)
      ' Show we are initialised
      bInit = True
      ' Return success
      Return True
    Catch ex As Exception
      ' Warn the user
      HandleErr("InitPerEditor error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failiure
      Return False
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ShowValues
  ' Goal:   Show the values associated with datarow 0 of a non DGV instance
  ' History:
  ' 27-11-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub ShowValues()
    Dim bIsSelBkUp As Boolean     ' Backup value of 'IsSelecting'

    Try
      ' Is this really a non-DGV instance?
      If (Not bIsDgv) Then
        ' Make sure this does not count as "dirty"
        bIsSelBkUp = bIsSelecting
        bIsSelecting = True
        If (tdlLocal.Tables(strTable).Rows.Count > 0) Then
          ' Set the values of row "0" of the non-dgv instance
          SetValues(tdlLocal.Tables(strTable).Rows(0))
        End If
        ' Allow dirty setting again
        bIsSelecting = bIsSelBkUp
      End If
    Catch ex As Exception
      ' Warn the user
      HandleErr("ShowValues error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   BindControl
  ' Goal:   Add one control to the table
  ' History:
  ' 27-11-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub BindControl(ByRef ctlThis As Control, ByVal strField As String, ByVal strType As String)

    Try
      ' Make room
      ReDim Preserve arCtrInfo(0 To intControls)
      ' Put the information here
      With arCtrInfo(intControls)
        .Ctl = ctlThis
        .Field = strField
        .Type = strType
      End With
      ' Increment number of controls
      intControls += 1
    Catch ex As Exception
      ' Warn user
      HandleErr("DgvHandle/BindControl error: " & ex.Message)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   DoEditDirty
  ' Goal:   Set or reset dirty flag for this editor
  ' History:
  ' 27-11-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function DoEditDirty(ByVal bSet As Boolean) As Boolean
    Dim dtrFound() As DataRow   ' Result of SELECT function
    Dim tbThis As TextBox       ' Textbox for changed time

    Try
      ' Don't let query selection fool "dirty"
      If (bIsSelecting) Then Return bSet
      ' Deal with the flag
      bDirty = bSet
      ' Have we changed?
      If (bDirty) AndAlso (strChanged <> "") Then
        ' Establish the exact datarow that needs date-adapting
        ' Get the correct datarow
        If (bIsDgv) Then
          dtrFound = tdlLocal.Tables(strTable).Select(strId & "=" & intSelectedId)
        Else
          dtrFound = tdlLocal.Tables(strTable).Select("")
        End If
        ' Found anything?
        If (dtrFound.Length > 0) Then
          ' Set the "Changed" date to now
          dtrFound(0).Item(strChanged) = Now
          ' Make sure it is shown 
          If (strTbChanged <> "") Then
            tbThis = objForm.Controls.Find(strTbChanged, True)(0)
            ' Additional safeguard!!
            If (Not tbThis Is Nothing) Then
              tbThis.Text = Format(dtrFound(0).Item(strChanged), "G")
            End If
          End If
        End If
      End If
      ' Return the value
      Return bSet
    Catch ex As Exception
      ' Warn user
      HandleErr("DgvHandle/DoEditDirty: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------
  ' Name:       SetDgvColumns()
  ' Goal:       Select which columns are shown in the datagrid view 
  ' History:
  ' 23-11-2009  ERK Created
  '---------------------------------------------------------------
  Public Sub SetDgvColumns(ByRef dgvThis As DataGridView, ByVal bRowHeadersVisible As Boolean, _
                           ByVal strColumns As String)
    Dim intI As Integer       ' Counter
    Dim arColumn() As String  ' Array of columns

    Try
      ' Access the DGV
      With dgvThis
        ' First disable all columns
        For intI = 0 To .ColumnCount - 1
          ' Disable this column
          .Columns(intI).Visible = False
          .Columns(intI).SortMode = DataGridViewColumnSortMode.Programmatic
        Next intI
        ' Transform the [strColumns] into an array
        arColumn = Split(strColumns, ";")
        ' Now find and enable the columns defined in [arColumn]
        For intI = 0 To UBound(arColumn)
          ' Enable this column
          With .Columns(arColumn(intI))
            .Visible = True
            ' .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
          End With
        Next intI
        .RowHeadersVisible = bRowHeadersVisible
        '' ======= DEBUGGING ===========
        'For intI = 1 To .Columns.Count
        '  With .Columns(intI - 1)
        '    Debug.Print("SetDgvColumns [" & intI - 1 & ":" & .Name & "] is " & IIf(.Visible, "visible", "hidden"))
        '  End With
        'Next intI
        '' ============================
      End With
    Catch ex As Exception
      ' Give the user an error message
      HandleErr("SetDgvColumns error: " & ex.Message)
    End Try
  End Sub
  '---------------------------------------------------------------
  ' Name:       SetupDgv()
  ' Goal:       Set up a datagrid view 
  ' History:
  ' 16-06-2009  ERK Created
  '---------------------------------------------------------------
  Public Sub SetupDgv(ByRef dgvThis As DataGridView, ByRef objData As Object)
    Dim intI As Integer     ' Counter

    Try
      ' Access the DGV
      With dgvThis
        .AutoGenerateColumns = True
        .DataSource = objData
        .RowHeadersVisible = True
        .SelectionMode = DataGridViewSelectionMode.FullRowSelect
        .MultiSelect = True
        .ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        ' Set column widths
        For intI = 1 To .Columns.Count
          'With .Columns(intI - 1)
          '  Debug.Print("SetupDgv Column [" & intI - 1 & ":" & .Name & "] is " & IIf(.Visible, "visible", "hidden"))
          'End With
          ' Only change if visible
          If (.Columns(intI - 1).Visible) Then
            .Columns(intI - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
          End If
        Next intI
        ' Select the correct row
        If (.Rows.Count > 0) Then
          .Rows(0).Selected = True
          .FirstDisplayedScrollingRowIndex = 0
        End If
        ' Scroll to the correct place
        .Refresh()
      End With
    Catch ex As Exception
      ' Give the user an error message
      HandleErr("SetupDgv error: " & ex.Message & vbCrLf & ex.StackTrace)
    End Try
  End Sub
  '---------------------------------------------------------------
  ' Name:       SetItem()
  ' Goal:       Copy the changed item to the dataset
  ' History:
  ' 27-11-2009  ERK Created
  ' 11-12-2012  ERK Return DGV to selectedId
  '---------------------------------------------------------------
  Public Sub SetItem(ByRef ctlThis As Control)
    Dim dtrFound() As DataRow     ' Result of SELECT function
    Dim intCtrIdx As Integer      ' The index of the control in [arCtrInfo]
    Dim intBackupId As Integer    ' We backup the selected ID
    Dim strType As String         ' The type of the control
    Dim strField As String        ' The field for this control
    Dim dtThis As DateTimePicker  ' DateTime picker control
    Dim chbThis As CheckBox       ' Checkbox control
    Dim cboThis As ComboBox       ' Combobox control
    Dim webThis As WebBrowser     ' Webbrowser control

    Try
      ' Can we adapt changes?
      If (bInit) AndAlso (Not bIsSelecting) AndAlso _
        ((Not bIsDgv) OrElse (intSelectedId >= 0)) Then
        ' Backup the selected ID
        intBackupId = intSelectedId
        ' Get the correct datarow
        If (bIsDgv) Then
          dtrFound = tdlLocal.Tables(strTable).Select(strId & "=" & intSelectedId)
        Else
          dtrFound = tdlLocal.Tables(strTable).Select("")
        End If
        ' Anything found?
        If (dtrFound.Length > 0) Then
          ' Get the index of the control
          intCtrIdx = GetControlIdx(ctlThis.Name)
          If (intCtrIdx < 0) Then
            ' Can't handle this control
            Status("DgvHandle/SetItem: control is not set: " & ctlThis.Name)
          End If
          ' Get the type of the control
          strType = arCtrInfo(intCtrIdx).Type
          strField = arCtrInfo(intCtrIdx).Field
          If (strType <> "") Then
            Select Case strType
              Case "textbox", "richtext"
                dtrFound(0).Item(strField) = ctlThis.Text
              Case "browser"
                webThis = ctlThis
                dtrFound(0).Item(strField) = webThis.DocumentText
              Case "combo"
                cboThis = ctlThis
                If (cboThis.ValueMember = "") Then
                  ' Use the "Item" property
                  dtrFound(0).Item(strField) = cboThis.SelectedItem
                Else
                  ' Use the "Value" property
                  dtrFound(0).Item(strField) = cboThis.SelectedValue
                End If
              Case "number"
                dtrFound(0).Item(strField) = GetNumber(ctlThis.Text)
              Case "datetime"
                dtThis = ctlThis
                dtrFound(0).Item(strField) = dtThis.Value
              Case "checkbox"
                chbThis = ctlThis
                dtrFound(0).Item(strField) = IIf(chbThis.Checked, "True", "False")
            End Select
            ' Restore the selected id
            intSelectedId = intBackupId
            ' Make sure THIS entry in the dgv is the selected one
            With dgvLocal
              For intI = 0 To .RowCount - 1
                ' Is this the row that should be selected?
                If (.Rows(intI).Cells(strId).Value = intSelectedId) Then
                  ' Check if it actually IS selected
                  If (Not .Rows(intI).Selected) Then
                    ' Select this one
                    .Rows(intI).Selected = True
                    ' How do we make sure the DGV actually SCROLLS to get
                    ' this item into view?
                    .FirstDisplayedScrollingRowIndex = intI
                  End If
                  ' Exit the for loop
                  Exit For
                End If
              Next intI
            End With
            ' Make sure changes are processed?
            DoEditDirty(True)
          End If
        End If
      End If
    Catch ex As Exception
      ' Warn user
      HandleErr("DgvHandle/SetItem: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
    End Try
  End Sub
  '---------------------------------------------------------------
  ' Name:       SetValues()
  ' Goal:       Set all the values of the controls associated with the DgvHandle
  '               to the values that are in the datarow
  ' History:
  ' 27-11-2009  ERK Created
  '---------------------------------------------------------------
  Private Sub SetValues(ByRef dtrThis As DataRow)
    Dim dtThis As DateTimePicker  ' DateTime picker control
    Dim tbThis As TextBox         ' Textbox control
    Dim rtThis As RichTextBox     ' RichTextBox control
    Dim chbThis As CheckBox       ' Checkbox control
    Dim cboThis As ComboBox       ' Combobox control
    Dim webThis As WebBrowser     ' Webbrowser control
    Dim intI As Integer           ' Counter

    Try
      ' Set the current datarow
      dtrCurrent = dtrThis
      ' Is there content?
      If (dtrThis Is Nothing) Then
        ' Leave without setting anything
        Exit Sub
      End If
      ' Make sure values can actually be changed!
      AllowEdit(True)
      ' Access all associated controls
      For intI = 0 To UBound(arCtrInfo)
        ' Access this element
        With arCtrInfo(intI)
          ' Action depends on the type of control
          Select Case .Type
            Case "textbox", "number"
              tbThis = .Ctl
              ' Is this the "Changed" one?
              If (.Field = strChanged) Then
                ' Does "Changed" already have a value?
                If (dtrThis.Item(.Field).ToString = "") Then
                  ' Do give it a value: NOW
                  dtrThis.Item(.Field) = Now
                End If
                tbThis.Text = Format(dtrThis.Item(.Field), "G")
              ElseIf (dtrThis.Item(.Field) & "" <> tbThis.Text) Then
                tbThis.Text = dtrThis.Item(.Field) & ""
              End If
            Case "combo"
              cboThis = .Ctl
              If (cboThis.ValueMember = "") Then
                ' Use the "Item" property
                cboThis.SelectedItem = dtrThis.Item(.Field)
              Else
                ' Use the "Value" property
                cboThis.SelectedValue = dtrThis.Item(.Field)
              End If
            Case "checkbox"
              chbThis = .Ctl
              chbThis.Checked = (dtrThis.Item(.Field) & "" = "True")
            Case "richtext"
              rtThis = .Ctl
              ' Only make changes when needed!
              If (dtrThis.Item(.Field) & "" <> rtThis.Text) Then
                rtThis.Text = dtrThis.Item(.Field) & ""
              End If
            Case "browser"
              webThis = .Ctl
              ' Only make changes when needed!
              If (dtrThis.Item(.Field).ToString <> webThis.DocumentText) Then
                webThis.DocumentText = dtrThis.Item(.Field).ToString
              End If
            Case "datetime"
              dtThis = .Ctl
              If (dtrThis.Item(.Field).ToString <> "") Then
                dtThis.Value = dtrThis.Item(.Field)
              End If
          End Select
        End With
      Next intI
    Catch ex As Exception
      ' Warn user
      HandleErr("DgvHandle/SetValues Error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '---------------------------------------------------------------
  ' Name:       IsVisible()
  ' Goal:       Find out if I am visible or not
  ' History:
  ' 04-10-2010  ERK Created
  '---------------------------------------------------------------
  Private Function IsVisible() As Boolean
    Try
      ' See if there are controls bound to me
      If (arCtrInfo.Length > 0) AndAlso (arCtrInfo(0).Ctl IsNot Nothing) Then
        ' Take the visibility of the first control bound to me
        Return arCtrInfo(0).Ctl.Visible
      ElseIf (objForm IsNot Nothing) AndAlso (objForm.Controls(0) IsNot Nothing) Then
        ' There is no bound controle. Check it by the form's controls
        Return objForm.Controls(0).Visible
      Else
        Return False
      End If
    Catch ex As Exception
      ' Warn user
      HandleErr("DgvHandle/IsVisible Error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------
  ' Name:       GetControlIdx()
  ' Goal:       Find the index in the array for this item
  ' History:
  ' 27-11-2009  ERK Created
  '---------------------------------------------------------------
  Private Function GetControlIdx(ByVal strName As String) As Integer
    Dim intI As Integer  ' Counter

    ' Find the entry in [arCtrInfo]
    For intI = 0 To arCtrInfo.Length - 1
      ' Is this the one?
      If (arCtrInfo(intI).Ctl.Name = strName) Then
        ' Return the type
        Return intI
      End If
    Next intI
    ' Found nothing - return failure
    Return -1
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  AllowEdit
  ' Goal :  Allow or disallow the editing of associated controls
  ' History:
  ' 27-11-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Sub AllowEdit(ByVal bSet As Boolean)
    Dim intI As Integer           ' Counter
    Dim tbThis As TextBox         ' Textbox control
    Dim rtThis As RichTextBox     ' RichTextBox control
    Dim dtThis As DateTimePicker  ' Date/time picker control
    Dim chbThis As CheckBox       ' Checkbox control
    Dim cboThis As ComboBox       ' Combobox control
    Dim bIsSelBkUp As Boolean     ' Backup value of 'IsSelecting'
    Dim colDisable As Color = Color.LightBlue
    Dim colEnable As Color = Color.White
    Dim colAllow As Color = Color.FromArgb(255, 255, 255, 192)
    Dim colWhite As Color = Color.FromArgb(255, 255, 255, 255)

    Try
      ' We are selecting
      bIsSelBkUp = bIsSelecting
      bIsSelecting = True
      ' Walk all controls
      For intI = 0 To UBound(arCtrInfo)
        ' Is there anything at all?
        If (UBound(arCtrInfo) = 0) AndAlso (arCtrInfo(0).Ctl Is Nothing) Then
          ' There is nothing, so quit
          Exit Sub
        End If
        ' Access this one
        With arCtrInfo(intI)
          ' Only do something when the backcolor is okay
          If (.Ctl.BackColor = colEnable) OrElse (.Ctl.BackColor = colAllow) OrElse (.Ctl.BackColor = colWhite) _
            OrElse (.Ctl.BackColor = colDisable) OrElse (.Ctl.BackColor.Name = "Window") Then
            ' Look at the type
            Select Case .Type
              Case "textbox", "number"
                tbThis = .Ctl
                ' tbThis.ReadOnly = (Not bSet)
                ' Are we disallowing editing?
                If (Not bSet) Then
                  ' Clear the text
                  tbThis.Text = ""
                  ' Set background color
                  tbThis.BackColor = colDisable
                Else
                  ' Set background color back
                  tbThis.BackColor = colEnable
                End If
              Case "richtext"
                rtThis = .Ctl
                ' rtThis.ReadOnly = (Not bSet)
                ' Are we disallowing editing?
                If (Not bSet) Then
                  ' Clear the text
                  rtThis.Text = ""
                  ' Set background color
                  rtThis.BackColor = colDisable
                Else
                  ' Set background color back
                  rtThis.BackColor = colEnable
                End If
              Case "datetime"
                dtThis = .Ctl
                dtThis.Enabled = bSet
              Case "checkbox"
                chbThis = .Ctl
                chbThis.Enabled = bSet
              Case "combo"
                cboThis = .Ctl
                cboThis.Enabled = bSet
            End Select
          End If
        End With
      Next intI
      ' We are no longer selecting
      bIsSelecting = bIsSelBkUp
    Catch ex As Exception
      ' Warn the user
      HandleErr("DgvHandle/AllowEdit error: " & ex.Message)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  Refresh
  ' Goal :  Make sure [SelectionChanged] is called or something like that!
  ' History:
  ' 26-01-2010  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Sub Refresh()
    Dim intRow As Integer = -1    ' The row selected
    Dim dtrRow As DataRow         ' New datarow

    Try
      ' Does a form already exist?
      If (objForm Is Nothing) OrElse (bRefresh) Then Exit Sub
      ' Are we visible already?
      If (objForm.Visible) AndAlso (bInit) AndAlso (bEdtOk) Then
        ' Is this the correct tab, which is visible?
        If (Not IsVisible()) Then Exit Sub
        ' Indicate we are busy
        bRefresh = True
        ' ======== DEBUG ===========
        ' intSelectedId = dgvLocal.Rows(intRow).Cells(strId).Value
        '         intSelectedId = SelectedId()
        ' ==========================
        '' Make sure we can be written --> unfortunately does not work...
        'AllowEdit(True)
        ' MsgBox("Hod " & strTable)
        ' Have a look at the correct table
        With tdlLocal.Tables(strTable)
          ' Are there any rows to be displayed (filtered)?
          If (bsLocal.Count = 0) Then
            ' Is this a DGV one?
            If (bIsDgv) Then
              ' Disallow editing
              AllowEdit(False)
              ' Indicate we are disturbable again
              bRefresh = False
              ' Exit 
              Exit Sub
            Else
              ' Make a new row!
              dtrRow = tdlLocal.Tables(strTable).NewRow
              tdlLocal.Tables(strTable).Rows.Add(dtrRow)
            End If
          End If
        End With
        ' Allow editing
        AllowEdit(True)
        ' Is there a DGV?
        If (Not dgvLocal Is Nothing) Then
          With dgvLocal
            ' Check whether there are any rows to be selected in the dgv
            If (.RowCount = 0) Then
              ' Indicate we are disturbable again
              bRefresh = False
              ' Just exit for the moment...
              Exit Sub
            End If
            ' Is anything selected?
            If (.SelectedCells.Count > 0) Then
              ' If there are selected cells, then we have to take THEM into account first!
              intRow = .SelectedCells(0).RowIndex
            ElseIf (.CurrentCell Is Nothing) Then
              ' Indicate we are disturbable again
              bRefresh = False
              ' If there is no current cell, then we won't show anything either
              Exit Sub
            ElseIf (tdlLocal.Tables(strTable).Rows.Count = 0) Then
              ' Indicate we are disturbable again
              bRefresh = False
              ' Leave - there is nothing to select!
              Exit Sub
            Else
              intRow = .CurrentCell.RowIndex
            End If
            ' Determine the QueryId
            'If (dgvLocal.Columns(strId) Is Nothing) Then Stop
            'If (dgvLocal.Rows(intRow).Cells(strId) Is Nothing) Then Stop
            intSelectedId = .Rows(intRow).Cells(strId).Value
            ' Is there any Delegate function to be called?
            If (Not opBeforeSelChanged Is Nothing) Then
              ' Call this procedure
              opBeforeSelChanged(intSelectedId)
            End If
            ' Show selection
            SelectId(intSelectedId)
            ' Reset the editor's dirty flag
            DoEditDirty(False)
          End With
        End If
      End If
      ' Indicate we are disturbable again
      bRefresh = False
    Catch ex As Exception
      ' Warn user
      HandleErr("DgvHandle/Refresh error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub

  Private Sub dgvLocal_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvLocal.ColumnHeaderMouseClick
    Try
      Dim newColumn As DataGridViewColumn = dgvLocal.Columns(e.ColumnIndex)
      Dim oldColumn As DataGridViewColumn = dgvLocal.SortedColumn
      Dim direction As System.ComponentModel.ListSortDirection
      Dim strOrder As String = strSort  ' Sort order
      Dim strCol As String = ""         ' Column

      ' Check if a mouse has actually been clicked
      If (Not bMouseClick) Then Exit Sub
      ' If oldColumn is null, then the DataGridView is not currently sorted.
      If oldColumn IsNot Nothing Then
        Status("Sorting...")
        ' Sort the same column again, reversing the SortOrder.
        If oldColumn Is newColumn AndAlso dgvLocal.SortOrder = _
            SortOrder.Ascending Then
          direction = System.ComponentModel.ListSortDirection.Descending
        Else

          ' Sort a new column and remove the old SortGlyph.
          direction = System.ComponentModel.ListSortDirection.Ascending
          Status("Clearing old sort...")
          oldColumn.HeaderCell.SortGlyphDirection = SortOrder.None
        End If
      Else
        direction = System.ComponentModel.ListSortDirection.Ascending
      End If

      ' Tell the user that we are going to sort by this feature
      Status("Sorting on [" & dgvLocal.Columns(e.ColumnIndex).Name & "] ... (please be patient)")
      ' Sort the selected column.
      dgvLocal.Sort(newColumn, direction)
      ' Indicate the direction of the Glyph
      If direction = System.ComponentModel.ListSortDirection.Ascending Then
        newColumn.HeaderCell.SortGlyphDirection = SortOrder.Ascending
      Else
        newColumn.HeaderCell.SortGlyphDirection = SortOrder.Descending
      End If
      ' Indicate that we have dealt with the mouseclick event
      '  Debug.Print(e.Clicks & e.Delta)

      '' Get the column
      'strCol = dgvLocal.Columns(e.ColumnIndex).Name
      '' Check if it is in here already
      'If (InStr(bsLocal.Sort, strCol) > 0) Then
      '  ' It is in here -- but are there more values
      '  If (InStr(bsLocal.Sort, ",") > 0) Then
      '    ' There are more values, so replace
      '    strOrder = strCol & " ASC"
      '  Else
      '    ' Get the direction
      '    If (InStr(bsLocal.Sort, " ASC") > 0) Then
      '      ' Make descending
      '      strOrder = strCol & " DESC"
      '    Else
      '      strOrder = strCol & " ASC"
      '    End If
      '  End If
      'Else
      '  ' it is not yet in here, so replace
      '  strOrder = strCol & " ASC"
      'End If
      '' Signal what we are going to sort
      'Status("Sorting [" & strCol & "]")
      '' Do the sorting
      'bsLocal.Sort = strOrder

      ' Indicate that the mouseclick has been handled
      bMouseClick = False
      Status("Ready (and thanks!)")
    Catch ex As Exception
      ' Warn user
      HandleErr("DgvHandle/Refresh error: " & ex.Message)
    End Try
  End Sub

  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  dgvLocal_DataBindingComplete
  ' Goal :  Make sure only the correct columns are beign shown
  ' History:
  ' 12-08-2014  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub dgvLocal_DataBindingComplete(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewBindingCompleteEventArgs) Handles dgvLocal.DataBindingComplete
    Dim intI As Integer ' Counter

    Try
      If (Not bInit) OrElse (dgvLocal Is Nothing) Then Exit Sub
      ' Make sure only the correct columns are visible
      ' Access the DGV
      With dgvLocal
        ' First disable all columns
        For intI = 0 To .ColumnCount - 1
          ' Checkk this column
          If (Not IsInSemiStack(loc_strColumns, .Columns(intI).Name)) Then
            .Columns(intI).Visible = False
          End If
        Next intI
      End With
    Catch ex As Exception
      ' Warn user
      HandleErr("DgvHandle/dgvLocalDataBinding error: " & ex.Message)
    End Try
  End Sub

  Private Sub dgvLocal_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgvLocal.MouseClick
    ' Tell that a mouse has clicked
    bMouseClick = True
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  dgvLocal_SelectionChanged
  ' Goal :  Show the indicated value in the appropriate controls
  ' History:
  ' 27-11-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Sub dgvLocal_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvLocal.SelectionChanged
    Try
      ' Call the more general Refresh subroutine
      Refresh()
      '   Debug.Print(Me.tdlLocal.DataSetName)
    Catch ex As Exception
      ' Warn user
      HandleErr("DgvHandle/SelectionChanged error: " & ex.Message)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   NonDgvChanged
  ' Goal:   When the data of a non-dgv has changed, show it
  ' History:
  ' 03-12-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub NonDgvChanged()
    ' Is this a non-dgv?
    If (Not bIsDgv) Then
      ' Okay, check whether there is a row zerio
      If (tdlLocal.Tables(strTable).Rows.Count > 0) Then
        ' Show it
        SetValues(tdlLocal.Tables(strTable).Rows(0))
      End If
    End If
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   SelectId
  ' Goal:   Select the indicated ID
  ' History:
  ' 27-11-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Private Sub SelectId(ByVal intItemId As Integer)
    Dim dtrFound() As DataRow   ' Result of the query

    Try
      ' Make sure this does not count as "Dirty"
      bIsSelecting = True
      ' Is this a DGV or not?
      If (bIsDgv) Then
        ' Find the data belonging to this periodId
        dtrFound = tdlLocal.Tables(strTable).Select(strId & "=" & intItemId)
        ' Found anything?
        If (dtrFound.Length > 0) Then
          SetValues(dtrFound(0))
        Else
          ' What if nothing is found?
          MsgBox("DgvHandle/SelectId: cannot find ID=" & intItemId)
        End If
      Else
        ' Get row number 0
        SetValues(tdlLocal.Tables(strTable).Rows(0))
      End If
      ' Switch off the query dirty stuff
      DoEditDirty(False)
      ' Switch off "dirty" protection
      bIsSelecting = False
    Catch ex As Exception
      ' Warn user
      HandleErr("DgvHandle/SelectId error: " & ex.Message)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   SelectName
  ' Goal:   Select the indicated "Name" vlaue
  ' History:
  ' 30-11-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub SelectName(ByVal strName As String)
    Dim dtrFound() As DataRow   ' Result of the query

    Try
      ' Make sure this does not count as "Dirty"
      bIsSelecting = True
      ' Is this a DGV or not?
      If (bIsDgv) Then
        ' Find the data belonging to this periodId
        dtrFound = tdlLocal.Tables(strTable).Select("Name='" & strName & "'")
        ' Found anything?
        If (dtrFound.Length > 0) Then
          ' Set the selected ID in the DGV
          SelectDgvId(dtrFound(0).Item(strId))
        Else
          ' The indicated row is not found...
        End If
      Else
        ' Get row number 0
        SetValues(tdlLocal.Tables(strTable).Rows(0))
      End If
      ' Switch off the query dirty stuff
      DoEditDirty(False)
      ' Switch off "dirty" protection
      bIsSelecting = False
    Catch ex As Exception
      ' Warn user
      HandleErr("DgvHandle/SelectName error: " & ex.Message)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  SelectedId
  ' Goal :  Return the Id value of the currently selected one
  ' History:
  ' 30-11-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function SelectedId() As Integer
    Dim intIdx As Integer = -1    ' Index of selected one
    Dim intRowIdx As Integer      ' Index of row

    ' Is any line selected?
    If (dgvLocal.SelectedCells.Count > 0) Then
      ' Determine which line is currently selected
      intRowIdx = dgvLocal.SelectedCells(0).RowIndex
      intIdx = dgvLocal.Rows(intRowIdx).Cells(strId).Value
    ElseIf (Not dgvLocal.CurrentCell Is Nothing) Then
      intRowIdx = dgvLocal.CurrentCell.RowIndex
      intIdx = dgvLocal.Rows(intRowIdx).Cells(strId).Value
    End If
    ' Return value found
    SelectedId = intIdx
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  SelectDgvId
  ' Goal :  Select a new period ID in the dgv
  ' History:
  ' 27-11-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Sub SelectDgvId(ByVal intItemId As Integer)
    Dim intI As Integer           ' Counter

    Try
      ' Make sure this new period is selected
      With dgvLocal
        For intI = 0 To .RowCount - 1
          ' Is this the correct row?
          If (.Rows(intI).Cells(strId).Value = intItemId) Then
            ' Delete all previous selections
            .CurrentCell = Nothing
            ' Select this one
            .Rows(intI).Selected = True
            ' If not visible, then scroll
            If (.FirstDisplayedScrollingRowIndex > intI) OrElse (.FirstDisplayedScrollingRowIndex + .DisplayedRowCount(False) < intI) Then
              .FirstDisplayedScrollingRowIndex = intI
            End If
            ' Exit the for loop
            Exit For
          End If
        Next intI
      End With
    Catch ex As Exception
      ' Warn user
      HandleErr("DgvHandle/SelectDgvId error: " & ex.Message)
    End Try
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  Exists
  ' Goal :  Find out whether a row fulfilling the criteria in [strSelect] already exists
  ' History:
  ' 28-11-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function Exists(ByVal strSelect As String) As Boolean
    Dim dtrFound() As DataRow ' Result of SELECT

    Try
      ' Find out...
      dtrFound = tdlLocal.Tables(strTable).Select(strSelect)
      ' Return result
      Exists = (dtrFound.Length > 0)
    Catch ex As Exception
      ' Warn user
      HandleErr("DgvHandle/Exists error: " & ex.Message)
      ' Return failure
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  AddNew
  ' Goal :  Add a new element to the table associated with this DGV
  '         The [arFldVal()] is an array containing pairs of Field, Value etc.
  ' History:
  ' 28-11-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function AddNew(ByVal ParamArray arFldVal() As String) As Boolean
    Dim dtrNew As DataRow   ' The new datarow added to the table
    Dim strField As String  ' The name of the field
    Dim strValue As String  ' The value represented as string
    Dim intId As Integer    ' The ID of the new element
    Dim intStart As Integer ' index in [arFldVal] to start from
    Dim intI As Integer     ' Counter

    Try
      ' Make a new datarow for the table
      dtrNew = tdlLocal.Tables(strTable).NewRow
      ' Is the ID (as first element) passed on?
      If (UBound(arFldVal) > 0) AndAlso (arFldVal(0) Like strId) Then
        ' Retrieve the ID from [arFldVal]
        intId = arFldVal(1)
        ' Set start value to "2"
        intStart = 2
      Else
        ' Add a new ID value for this table
        intId = tdlLocal.Tables(strTable).Rows.Count + 1
        ' Set start value to "0"
        intStart = 0
      End If
      dtrNew.Item(strId) = intId
      ' Process all elements of [arFldVal]
      For intI = intStart To UBound(arFldVal) Step 2
        ' Get the [Field] and [Value]
        strField = arFldVal(intI)
        strValue = arFldVal(intI + 1)
        ' Store these values in the datatable
        dtrNew.Item(strField) = strValue
      Next intI
      ' If a "Changed" field is defined, set it
      If (strChanged <> "") Then dtrNew.Item(strChanged) = Now
      ' If a "Created" field is defined, set it
      If (strCreated <> "") Then dtrNew.Item(strCreated) = Now
      ' Set the correct parent
      If (strParentTable <> "") Then
        dtrNew.SetParentRow(GetParentRow)
      End If
      ' Add the row to the table
      tdlLocal.Tables(strTable).Rows.Add(dtrNew)
      ' Make sure this new table is now being shown
      SelectDgvId(intId)
      ' Make sure the value of [SelectedId] is adapted
      intSelectedId = intId
      ' Return success
      Return True
    Catch ex As Exception
      ' Show user the error
      HandleErr("DgvHandle/AddNew error: " & ex.Message)
      ' Return failure
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  GetParentRow
  ' Goal :  Try to get or make and get the parent table's row
  ' History:
  ' 28-11-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Function GetParentRow() As DataRow
    Dim dtrParent As DataRow = Nothing  ' Try to get or make a parent row

    With tdlLocal.Tables(strParentTable)
      If (.Rows.Count = 0) Then
        ' Add a new row
        dtrParent = .NewRow
        .Rows.Add(dtrParent)
      Else
        dtrParent = .Rows(0)
      End If
      ' REturn the correct row
      GetParentRow = dtrParent
    End With
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  Remove
  ' Goal :  Delete one row from the dataset
  '         Return the ID of the row now selected
  ' History:
  ' 28-11-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function Remove() As Boolean
    Dim intSelId As Integer       ' The period ID as we see is selected
    Dim intRow As Integer         ' The selected row
    Dim intI As Integer           ' Counter
    Dim dtrFound() As DataRow     ' Selected rows

    Try
      ' Are any rows selected?
      If (dgvLocal.SelectedCells.Count = 0) Then
        ' Warn user
        MsgBox("First select a row to be deleted")
        Return False
      End If
      ' Get the selected row index
      intRow = dgvLocal.SelectedCells(0).RowIndex
      ' Get the PeriodId value
      intSelId = dgvLocal.Rows(intRow).Cells(strId).Value
      ' Find the right row
      dtrFound = tdlLocal.Tables(strTable).Select(strId & "=" & intSelId)
      If (dtrFound.Length > 0) Then
        ' Show we are selecting
        bIsSelecting = True
        ' Delete this row
        dtrFound(0).Delete()
      End If
      ' Find rows that now need adaptation
      dtrFound = tdlLocal.Tables(strTable).Select(strId & ">" & intSelId)
      ' Visit all rows
      For intI = 0 To dtrFound.Length - 1
        'Adapt their ID
        dtrFound(intI).Item(strId) -= 1
      Next intI
      ' Okay, now we are finished selecting
      bIsSelecting = False
      ' Now check if something is still selected...
      If (dgvLocal.SelectedCells.Count = 0) AndAlso (dgvLocal.Rows.Count > 0) Then
        ' Try to select the last row
        intSelId = dgvLocal.Rows(dgvLocal.Rows.Count - 1).Cells(strId).Value
        ' Just make sure changes are accepted
        tdlLocal.AcceptChanges()
        ' Select this one
        SelectDgvId(intSelId)
      ElseIf (dgvLocal.Rows.Count = 0) Then
        ' Just make sure changes are accepted
        tdlLocal.AcceptChanges()
        ' Disallow editing
        AllowEdit(False)
      Else
        ' Just make sure changes are accepted
        tdlLocal.AcceptChanges()
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Show user the error
      HandleErr("DgvHandle/Remove error: " & ex.Message)
      ' Return failure
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :    ChangeLoc
  ' Goal :    Change the location where this table's file comes from/is saved to
  ' Return :  True if a new location has been found
  ' History:
  ' 23-11-2009  ERK Created
  ' 28-11-2009  ERK Adapted for [DgvHandle]
  ' ---------------------------------------------------------------------------------------------------------
  Public Function ChangeLoc(ByRef tbFileLoc As TextBox, ByVal strDir As String, _
                            ByVal strExtension As String) As Boolean
    Dim strNewFile As String   ' New location where file is to be saved
    Dim strOrgFile As String  ' Original location of the file

    Try
      ' Get the original location of the file
      strOrgFile = tbFileLoc.Text
      ' Ask user for a new file location for saving
      With dlgSave
        ' Set the directory by default to the Query directory
        .InitialDirectory = strDir
        ' Set the file name to the one without path
        .FileName = IO.Path.GetFileName(strOrgFile)
        ' Set filters
        .Filter = strTable & " files|*" & strExtension
        ' Set extension
        .DefaultExt = IO.Path.GetExtension(strOrgFile)
        ' Ask for overwrite permission
        .OverwritePrompt = True
        ' Set text
        .Title = "Specify the new location for the backup of this file"
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.No, Windows.Forms.DialogResult.Cancel
            ' Leave!
            Return False
        End Select
        ' Get the file name
        strNewFile = .FileName
      End With
      ' Check whether the new file name is the same as the old one
      If (strNewFile = strOrgFile) Then
        ' Just leave
        Status("New file name is the same as the existing one")
        Return False
      End If
      ' Check whether the names as such are the same
      If (IO.Path.GetFileNameWithoutExtension(strNewFile) <> IO.Path.GetFileNameWithoutExtension(strOrgFile)) Then
        ' Warn user
        MsgBox("The name of the file should be identical - if you want to change that, then change the name of the query")
        Return False
      End If
      ' Try to copy..
      Try
        ' Check whether the source file's drive exists
        If (DriveExists(strOrgFile)) Then
          ' Is the date of the source NEWER than that of the destination?
          If (IO.File.GetLastWriteTime(strOrgFile) > IO.File.GetLastWriteTime(strNewFile)) Then
            ' Copy the file, overwriting it
            IO.File.Copy(strOrgFile, strNewFile, True)
          ElseIf (IO.File.Exists(strNewFile)) Then
            ' Ask whether user really wants to overwrite...
            MsgBox("The file " & strNewFile & " already exists." & vbCrLf & _
                   "You should select the current " & strTable & ", do " & strTable & _
                   "/Remove, and then " & strTable & "/Add the already existing file")
            ' Leave without changing anything
            Return False
          End If
        End If
      Catch ex As Exception
        ' If copying failed, then still the location can be changed!
      End Try
      ' Set the changed file name
      tbFileLoc.Text = strNewFile
      ' Return success
      Return True
    Catch ex As Exception
      ' Wann user
      HandleErr("DgvHandle/Changeloc error: " & ex.Message)
      ' Return failure
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :    DriveExists
  ' Goal :    Does the drive indicated by [strFile] exist?
  ' History:
  ' 02-12-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Function DriveExists(ByVal strFile As String) As Boolean
    Dim arDrive() As String   ' Array of drives
    Dim intI As Integer       ' Counter

    ' Get the drives of this machine
    arDrive = IO.Directory.GetLogicalDrives
    For intI = 0 To UBound(arDrive)
      ' Is it this drive?
      If (InStr(strFile, arDrive(intI), CompareMethod.Text) = 1) Then
        ' GOt it!
        Return True
      End If
    Next intI
    ' Drive does not exist on this machine
    Return False
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  Kill
  ' Goal :  Make this instance of a DgvHandle inactive/inaccessable
  ' History:
  ' 21-12-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Sub Kill()
    ' Set the most important pointers to nothing
    bsLocal = Nothing
    ' dvLocal = Nothing
    dgvLocal = Nothing
    ' Reset the "Init" flag, which prevents e.g. SelectionChanged to act
    bInit = False
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  GetNumber
  ' Goal :  Transform a string into a number, if possible
  ' History:
  ' 27-11-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Private Function GetNumber(ByVal strIn As String) As Integer
    ' Trim it
    strIn = Trim(strIn)
    ' Is there anything?
    If (strIn = "") Then
      ' Return zero
      Return 0
    End If
    ' Is it numeric?
    If (Not IsNumeric(strIn)) Then
      ' Return zero
      Return 0
    End If
    ' Return the number
    GetNumber = CInt(strIn)
  End Function
End Class
