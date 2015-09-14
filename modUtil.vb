Option Strict On
Option Explicit On
Option Compare Binary
Imports System.Net
Module modUtil
  ' ==========================PUBLIC==============================
  Public strSetFile As String = ""   ' Actual settings file
  ' ==========================PRIVATE=============================
  Private Const def_SetName As String = "CesaxSettings.xml"
  Private Const def_SetPsd As String = "Cesax.xsd"
  Private Const def_Applic As String = "Cesax"
  Private SaveDial As SaveFileDialog = New SaveFileDialog
  ' ------------------------------------------------------------------------------------
  ' Name:   ControlPressed
  ' Goal:   Determine whether the user pressed Control or not
  ' History:
  ' 28-04-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function ControlPressed() As Boolean
    ' See MS: ms-help://MS.VSCC.v90/MS.msdnexpress.v90.en/dv_fxmancli/html/1e184048-0ae3-4067-a200-d4ba31dbc2cb.htm
    Return ((Control.ModifierKeys And Keys.Control) = Keys.Control)
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   ShiftPressed
  ' Goal:   Determine whether the user pressed SHIFT or not
  ' History:
  ' 30-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function ShiftPressed() As Boolean
    ' Check whether SHIFT was pressed
    Return ((Control.ModifierKeys And Keys.Shift) = Keys.Shift)
  End Function
  Public Function AltPressed() As Boolean
    Return ((Control.ModifierKeys And Keys.Alt) = Keys.Alt)
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   HandleErr
  ' Goal:   Deal with an error in a uniform way
  ' History:
  ' 10-08-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub HandleErr(ByVal strMsg As String)
    ' Show the form with the correct string
    With frmError
      .ErrText = strMsg
      Select Case .ShowDialog
        Case DialogResult.Abort
          ' Quit the program without saving
          End
        Case DialogResult.OK
          ' First try to save the data
          Try
            ' Add the save function here (but what do we need to save?)
            If (Not pdxCurrentFile Is Nothing) Then
              If (strCurrentFile <> "") Then
                ' Save the CRP information to the selected filename as XML
                pdxCurrentFile.Save(strCurrentFile)
              End If
            End If
          Catch ex As Exception
            ' If an exception occurs, that is too bad
          End Try
          ' Now stop the program
        Case DialogResult.Cancel
          ' Signal that we need to stop the current action
          bInterrupt = True
        Case DialogResult.Ignore
          ' Continue as if all is well
      End Select
    End With
    '' Ask user what to do
    'Select Case MsgBox(strMsg & vbCrLf & vbCrLf & "=============================" & _
    '                   vbCrLf & "Take a deep breath...!" & _
    '                   vbCrLf & "What would you like to do?" & _
    '                   vbCrLf & "  Continue  (Yes)" & _
    '                   vbCrLf & "  Interrupt (No)" & _
    '                   vbCrLf & "  Exit this program without saving (Cancel)", MsgBoxStyle.YesNoCancel)
    '  Case MsgBoxResult.Yes
    '    ' Continue!
    '  Case MsgBoxResult.No
    '    ' Signal that we need to stop
    '    bInterrupt = True
    '  Case MsgBoxResult.Cancel
    '    ' The program should finish
    '    End
    'End Select
  End Sub
  '---------------------------------------------------------------
  ' Name:       CopyTableRow()
  ' Goal:       Copy a row of data from one table to the other
  ' History:
  ' 01-03-2011  ERK Created
  '---------------------------------------------------------------
  Public Function CopyTableRow(ByRef dtrSrc As DataRow, ByRef tblDst As DataTable, ByRef dtrParent As DataRow) As Boolean
    Dim dtrNew As DataRow ' New datarow to be made
    Dim intI As Integer   ' Counter

    Try
      ' Validate
      If (dtrSrc Is Nothing) OrElse (tblDst Is Nothing) Then Return False
      ' Make a new datarow in the destination set
      dtrNew = tblDst.NewRow
      ' Copy the elements from source to destination
      For intI = 0 To dtrSrc.Table.Columns.Count - 1
        ' Copy this column
        dtrNew.Item(intI) = dtrSrc.Item(intI)
      Next intI
      ' Set the parentrow of the new table
      dtrNew.SetParentRow(dtrParent)
      ' Add this row to the destination table
      tblDst.Rows.Add(dtrNew)
      ' Return success
      Return True
    Catch ex As Exception
      ' Give the user an error message
      HandleErr("modUtil/CopyTableRow error: " & ex.Message)
      ' Return empty...
      Return False
    End Try
  End Function
  '---------------------------------------------------------------
  ' Name:       GetFileName()
  ' Goal:       Ask for a file name from the user
  ' History:
  ' 04-12-2008  ERK Created
  '---------------------------------------------------------------
  Public Function GetFileName(ByRef FileDial As OpenFileDialog, ByVal strInitDir As String, _
    ByRef strFileName As String, ByVal strFilter As String) As Boolean
    Try
      ' Ask user to identify a settings file
      With FileDial
        ' Set the initial directory to start looking from
        .InitialDirectory = strInitDir
        ' Only allow one selection
        .Multiselect = False
        ' Do check whether given file exists
        .CheckFileExists = True
        ' Provide default extension
        .DefaultExt = "xml"
        ' Initial file name
        .FileName = strFileName
        ' Set filter to be looking for
        If (strFilter = "") Then
          .Filter = ""
          ' TODO: deal with this - set directory if filter is empty
        Else
          .Filter = strFilter & "|All files |*.*"
        End If
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK, Windows.Forms.DialogResult.Yes
            ' Continue with the program and return success
            GetFileName = True
          Case Else
            ' Close the form and return to main program
            Return False
        End Select
        ' Retrieve the file name
        strFileName = .FileName
      End With
      ' Return success, if there is a filename
      Return (strFileName <> "")
    Catch ex As Exception
      ' Give the user an error message
      HandleErr("GetFileName error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return empty...
      Return False
    End Try
  End Function
  '---------------------------------------------------------------
  ' Name:       GetDirName()
  ' Goal:       Ask for a directory name from the user
  ' History:
  ' 06-01-2009  ERK Created
  '---------------------------------------------------------------
  Public Function GetDirName(ByRef DirDial As FolderBrowserDialog, ByRef strDir As String, _
    ByVal strDescr As String, Optional ByVal strInitialDir As String = "") As Boolean
    Try
      ' Ask user to identify a settings file
      With DirDial
        ' Is there an initial directory?
        If (strDir <> "") Then
          .SelectedPath = strDir
        ElseIf (strInitialDir <> "") Then
          .SelectedPath = strInitialDir
        Else
          ' Set the initial directory to start looking from
          .RootFolder = Environment.SpecialFolder.MyComputer
        End If
        ' Set the description
        .Description = strDescr
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK, Windows.Forms.DialogResult.Yes
            ' Continue with the program and return success
            ' GetDirName = True
          Case Else
            ' Close the form and return to main program
            Return False
        End Select
        ' Retrieve the file name
        strDir = .SelectedPath
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Give the user an error message
      HandleErr("GetDirName error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return empty...
      Return False
    End Try
  End Function
  '---------------------------------------------------------------
  ' Name:       FileSaveAs()
  ' Goal:       Ask user where to save the file
  ' History:
  ' 10-12-2008  ERK Created
  '---------------------------------------------------------------
  Public Function FileSaveAs(ByVal strInitDir As String, ByRef strFileName As String, _
     ByVal strFilter As String) As Boolean
    Try
      With SaveDial
        ' Set the initial directory
        .InitialDirectory = strInitDir
        ' Provide the default extension
        .DefaultExt = "xml"
        ' Initial filename
        .FileName = strFileName
        ' Set filter
        If (strFilter = "") Then
          .Filter = ""
          ' TODO: deal with this - set directory if filter is empty
        Else
          .Filter = strFilter & "|All files |*.*"
        End If
        Select Case .ShowDialog
          Case DialogResult.OK, DialogResult.Yes
            ' Continue with the dialog and return success
            FileSaveAs = True
          Case Else
            ' Close the form and return to main program
            FileSaveAs = False
        End Select
        ' Retrieve the filename
        strFileName = .FileName
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Give the user an error message
      HandleErr("modUtil/FileSaveAs error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return empty...
      Return False
    End Try
  End Function
  '---------------------------------------------------------------
  ' Name:       GetSetDir()
  ' Goal:       Get the settings directory that belongs to this program
  ' History:
  ' 15-11-2011  ERK Created
  '---------------------------------------------------------------
  Public Function GetSetDir() As String
    Dim strDir As String  ' Directory

    Try
      strDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & _
            "\" & def_Applic
      ' Check if it exists
      If (Not IO.Directory.Exists(strDir)) Then IO.Directory.CreateDirectory(strDir)
      ' Return the result
      Return strDir
    Catch ex As Exception
      ' Give the user an error message
      HandleErr("modUtil/GetSetDir error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return empty...
      Return ""
    End Try
  End Function
  '---------------------------------------------------------------
  ' Name:       GetDocDir()
  ' Goal:       Get the documents directory that belongs to this program
  ' History:
  ' 07-02-2013  ERK Created
  '---------------------------------------------------------------
  Public Function GetDocDir() As String
    Dim strDir As String  ' Directory

    Try
      strDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & _
            "\ru\" & def_Applic
      ' Check if it exists
      If (Not IO.Directory.Exists(strDir)) Then IO.Directory.CreateDirectory(strDir)
      ' Return the result
      Return strDir
    Catch ex As Exception
      ' Give the user an error message
      HandleErr("modUtil/GetDocDir error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return empty...
      Return ""
    End Try
  End Function
  '---------------------------------------------------------------
  ' Name:       GetLocalDir()
  ' Goal:       Get the documents directory that belongs to this program
  '             This directory is not in the "roaming" section...
  ' History:
  ' 07-02-2013  ERK Created
  '---------------------------------------------------------------
  Public Function GetLocalDir(Optional ByVal strAppl As String = def_Applic) As String
    Dim strMyDoc As String  ' Directory of my documents
    Dim strDir As String  ' Directory

    Try
      ' Get my documents
      strMyDoc = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
      strDir = strMyDoc & "\ru\" & strAppl
      ' Check if it exists
      If (Not IO.Directory.Exists(strDir)) Then IO.Directory.CreateDirectory(strDir)

      '' Old
      'strDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) & _
      '      "\ru\" & def_Applic
      '' Check if it exists
      'If (Not IO.Directory.Exists(strDir)) Then IO.Directory.CreateDirectory(strDir)
      ' Return the result
      Return strDir
    Catch ex As Exception
      ' Give the user an error message
      HandleErr("modUtil/GetLocalDir error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return empty...
      Return ""
    End Try
  End Function

  '---------------------------------------------------------------
  ' Name:       GetSettingsFile()
  ' Goal:       Try to find the correct XML settings file for this program
  ' History:
  ' 04-12-2008  ERK Created
  '---------------------------------------------------------------
  Public Function GetSettingsFile() As String
    Dim strSetDefault As String = ""    ' Default location where settings file should be
    Dim strSetFileLoc As String = Application.StartupPath
    Dim ofdThis As New OpenFileDialog
    Dim sfdThis As New SaveFileDialog

    Try
      ' Fill in the default settings file location of this application
      strSetDefault = strSetFileLoc & "\" & def_SetName & ".txt"
      ' First try to retrieve the settings file from the registry
      strSetFile = GetSetting(def_Applic, "Settings", "Location", strSetDefault)
      ' If this is the first time usage, then SetDefault equals SetFile
      If (strSetDefault = strSetFile) Then
        ' Check existence of settings file first
        If (IO.File.Exists(strSetDefault)) Then
          ' Make a foldername within ApplicationData
          strSetFile = GetSetDir()
          ' Create folder if needed
          If (Not IO.Directory.Exists(strSetFile)) Then IO.Directory.CreateDirectory(strSetFile)
          ' Maket he full settings file 
          strSetFile &= "\" & def_SetName
          ' Copy the default settings file there
          IO.File.Copy(strSetDefault, strSetFile, True)
        End If
      End If
      ' Check again
      If (strSetDefault = strSetFile) Then
        ' Allow running, if the file exists...
        If (IO.File.Exists(strSetFile)) Then
          ' But do complain!!!
          Select Case MsgBox("You are using the settings file stored with the program." & vbCrLf & _
                 "This is bad practice - each user should be able to change his/her own settings." & vbCrLf & _
                 "The program will run, but make sure you change your life!", MsgBoxStyle.OkCancel)
            Case MsgBoxResult.Ok
              ' Good, do continue!
            Case MsgBoxResult.Cancel
              ' Opt out anyway
              End
          End Select
        Else
          ' Something is terribly wrong - closing
          MsgBox("Cannot run " & def_Applic & "." & vbCrLf & _
                 "Cannot find settings file: " & strSetFile & vbCrLf & _
                 "Closing the program.")
          ' Completely exit program
          End
        End If
      End If
      ' See if this file exists
      If (Not IO.File.Exists(strSetFile)) Then
        ' Ask user to identify a settings file
        If (Not GetFileName(ofdThis, strSetFileLoc, strSetFile, def_Applic & " Settings|*.xml")) Then
          ' Return empty file
          strSetFile = ""
        End If
      End If
      ' See if this file exists
      If (strSetFile <> "") Then
        ' Save the settings file location in the registry
        Call SaveSetting(def_Applic, "Settings", "Location", strSetFile)
        ' Read the settings file
        If (ReadSettingsXML(strSetFile)) Then
          ' Read the appropriate variables
          Call ReadSettings()
        Else
          MsgBox("Getsettingsfile" & vbCrLf & "Unable to properly read settings from " & strSetFile)
          ' Make sure we don't return anything!!!
          strSetFile = ""
        End If
      End If
      ' Return to caller
      GetSettingsFile = strSetFile
    Catch ex As Exception
      ' Give the user an error message
      HandleErr("GetSettingsFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return empty...
      GetSettingsFile = ""
    End Try
  End Function
  '---------------------------------------------------------------
  ' Name:       AddComboBoxColumn()
  ' Goal:       Add one DataGridViewComboBoxColumn
  ' History:
  ' 16-06-2009  ERK Created
  '---------------------------------------------------------------
  Public Sub AddComboBoxColumn(ByRef dgvThis As DataGridView, ByVal strName As String, ByVal strItems As String)
    Dim comboboxColumn As New DataGridViewComboBoxColumn()  ' The new cbo Column
    Dim intI As Integer                                     ' Counter
    Dim arItem() As String                                  ' Array of items

    Try
      ' Create the cbo Column
      comboboxColumn = CreateComboBoxColumn(strName)
      ' Get the items
      arItem = Split(strItems, ";")
      ' Fill the cbo Column
      For intI = 1 To arItem.Length
        comboboxColumn.Items.Add(arItem(intI - 1))
      Next intI
      ' Set the header name of this column
      comboboxColumn.HeaderText = strName
      ' Tack this example column onto the end.
      dgvThis.Columns.Add(comboboxColumn)
    Catch ex As Exception
      ' Give the user an error message
      HandleErr("AddComboBoxColumn error: " & ex.Message)
    End Try
  End Sub
  '---------------------------------------------------------------
  ' Name:       SetDgvColumns()
  ' Goal:       Select which columns are shown in the datagrid view 
  ' History:
  ' 23-11-2009  ERK Created
  '---------------------------------------------------------------
  Public Sub SetDgvColumns(ByRef dgvThis As DataGridView, ByVal bRowHeadersVisible As Boolean, _
                           ByVal ParamArray arColumn() As String)
    Dim intI As Integer   ' Counter

    Try
      ' Access the DGV
      With dgvThis
        ' First disable all columns
        For intI = 0 To .ColumnCount - 1
          ' Disable this column
          .Columns(intI).Visible = False
        Next intI
        ' Now find and enable the columns defined in [arColumn]
        For intI = 0 To UBound(arColumn)
          ' Enable this column
          With .Columns(arColumn(intI))
            .Visible = True
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader
          End With
        Next intI
        .RowHeadersVisible = bRowHeadersVisible
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
  Public Sub SetupDgv(ByRef dgvThis As DataGridView, ByRef objData As Object, _
                       ByVal ParamArray arColumn() As String)
    Dim intI As Integer     ' Counter
    Dim strColumn As String ' Name of column
    Dim strValues As String ' The values (separated by ;)

    Try
      ' Access the DGV
      With dgvThis
        .AutoGenerateColumns = True
        .DataSource = objData
        .RowHeadersVisible = True
        .SelectionMode = DataGridViewSelectionMode.RowHeaderSelect
        .MultiSelect = True
        .ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        ' Check all columns that need attention
        For intI = 0 To arColumn.Length - 1 Step 2
          ' Get column name
          strColumn = arColumn(intI)
          strValues = arColumn(intI + 1)
          ' Is there such a column?
          If (.Columns.Contains(strColumn)) Then
            ' Now remove the automatically generated "Cmp" column
            .Columns.Remove(strColumn)
            ' Add the column again with the correct values, displayed using a combobox for selection
            AddComboBoxColumn(dgvThis, strColumn, strValues)
          End If
        Next intI
        ' Set column widths
        For intI = 1 To .Columns.Count
          .Columns(intI - 1).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
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
      HandleErr("SetupDgv error: " & ex.Message)
    End Try
  End Sub
  '---------------------------------------------------------------
  ' Name:       CreateComboBoxColumn()
  ' Goal:       Create a DataGridViewComboBoxColumn
  ' History:
  ' 16-06-2009  ERK Created
  '---------------------------------------------------------------
  Private Function CreateComboBoxColumn(ByVal strName As String) As DataGridViewComboBoxColumn
    Dim column As New DataGridViewComboBoxColumn()

    Try
      With column
        .DataPropertyName = strName
        .HeaderText = strName
        .DropDownWidth = 20
        .Width = 80
        .MaxDropDownItems = 3
        .FlatStyle = FlatStyle.Standard
      End With
      Return column
    Catch ex As Exception
      ' Give the user an error message
      HandleErr("CreateComboBoxColumn error: " & ex.Message)
      ' Return nothing
      Return Nothing
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  GetSettingValue
  ' Goal :  Return the VALUE of a general setting named [strName]
  ' History:
  ' 19-11-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function GetSettingValue(ByVal strName As String) As String
    ' Call the more general function
    GetSettingValue = GetTableSetting(tdlSettings, strName)
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  GetTableSetting
  ' Goal :  Return the VALUE of a general setting named [strName]
  ' History:
  ' 19-11-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function GetTableSetting(ByRef tdlThis As DataSet, ByVal strName As String, _
                                  Optional ByVal strDefault As String = "") As String
    Dim dtrSettingName() As DataRow   ' Datarow of the setting with the particular name

    ' Get the datarow of the setting with the indicated name
    dtrSettingName = tdlThis.Tables("Setting").Select("Name='" & strName & "'")
    ' Did we find anything?
    If (dtrSettingName.Length = 0) Then
      ' Nothing found = return default
      GetTableSetting = strDefault
    Else
      ' Return the appropriate value
      GetTableSetting = dtrSettingName(0).Item("Value").ToString
    End If
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  SetTableSetting
  ' Goal :  Set the value of a general setting named [strName]
  ' History:
  ' 14-12-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function SetTableSetting(ByRef tdlThis As DataSet, ByVal strName As String, ByVal strValue As String) As Boolean
    Dim dtrSettingName() As DataRow   ' Datarow of the setting with the particular name
    Dim dtrNew As DataRow             ' New row for [Setting] table
    Dim dtrParent As DataRow          ' The parent datarow (GENERAL section)

    Try
      ' Get the datarow of the setting with the indicated name
      dtrSettingName = tdlThis.Tables("Setting").Select("Name='" & strName & "'")
      ' Did we find anything?
      If (dtrSettingName.Length = 0) Then
        ' First determine the parent row
        If (tdlThis.Tables("General") Is Nothing) OrElse (tdlThis.Tables("General").Rows.Count = 0) Then
          ' Create this row
          dtrParent = tdlThis.Tables("General").NewRow
          tdlThis.Tables("General").Rows.Add(dtrParent)
        Else
          ' Take the first row
          dtrParent = tdlThis.Tables("General").Rows(0)
        End If
        ' This particular setting does not exist yet. We are now going to add it
        With tdlThis.Tables("Setting")
          dtrNew = .NewRow
          dtrNew.Item("Name") = strName
          dtrNew.Item("Value") = strValue
          ' Set the parent row
          dtrNew.SetParentRow(dtrParent)
          ' Add this row to the table
          .Rows.Add(dtrNew)
        End With
      Else
        ' Set the appropriate value
        dtrSettingName(0).Item("Value") = strValue
      End If
      ' Return success
      SetTableSetting = True
    Catch ex As Exception
      ' Return failure
      SetTableSetting = False
    End Try
  End Function
  '----------------------------------------------------------------------------------------------------------
  ' Name:       ClearTable()
  ' Goal:       Clear this table
  ' History:
  ' 02-06-2010  ERK Created
  '----------------------------------------------------------------------------------------------------------
  Public Sub ClearTable(ByRef tblThis As DataTable)
    Dim intI As Integer   ' Counter needed to delete all rows
    ' Dim intPtc As Integer ' Percentage

    Try
      ' Clear the table's rows
      With tblThis
        For intI = .Rows.Count - 1 To 0 Step -1
          '' Show where we are
          'intPtc = (.Rows.Count - intI) * 100 \ .Rows.Count
          'Status("Clearing " & intPtc & "%", intPtc)
          ' Delete this row
          .Rows(intI).Delete()
          Application.DoEvents()
        Next intI
      End With
      ' Make sure changes are processed
      'tdlSettings.AcceptChanges()
    Catch ex As Exception
      ' Warn the user
      HandleErr("modUtil/ClearTable error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  '----------------------------------------------------------------------------------------------------------
  ' Name:       GetUniqueId()
  ' Goal:       Get a unique ID number from table [tblThis]
  ' History:
  ' 20-07-2010  ERK Created
  '----------------------------------------------------------------------------------------------------------
  Public Function GetUniqueId(ByRef tblThis As DataTable, ByVal strIdName As String) As Integer
    Dim intI As Integer   ' Counter
    Dim intMax As Integer ' Maximum found so far
    Dim intId As Integer  ' Current Id value

    Try
      ' Validate
      If (tblThis Is Nothing) OrElse (strIdName = "") Then Return -1
      ' Check if the field actually exists in this table
      If (tblThis.Columns(strIdName) Is Nothing) Then Return -1
      ' Are there any rows?
      If (tblThis.Rows.Count = 0) Then Return 1
      ' Start with the first row's number
      intMax = CInt(tblThis.Rows(0).Item(strIdName))
      ' Try find highest value so far
      For intI = 1 To tblThis.Rows.Count - 1
        ' Get this value
        intId = CInt(tblThis.Rows(intI).Item(strIdName))
        ' Is this value larger?
        If (intId > intMax) Then intMax = intId
      Next intI
      ' Return a value that is 1 larger than the maximum
      Return intMax + 1
    Catch ex As Exception
      ' Warn the user
      HandleErr("modUtil/GetUniqueId error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   AddSemiStack
  ' Goal:   Add string [strAdd] to the semicolon separated stack [strOrg]
  ' History:
  ' 23-06-2010  ERK Created
  ' 01-09-2010  ERK Assume [strAdd] is also a semicolon separated collection
  ' ------------------------------------------------------------------------------------
  Public Sub AddSemiStack(ByRef strOrg As String, ByVal strAdd As String, _
                          Optional ByVal bUnique As Boolean = False, Optional ByVal strDivider As String = ";")
    Dim arAdd() As String ' Collection [strAdd] broken up
    Dim intI As Integer   ' Counter

    Try
      ' Validate
      If (strAdd = "") Then Exit Sub
      ' Check if there already is something in [strOrg]
      If (strOrg = "") Then
        ' Straight copy
        strOrg = strAdd
      Else
        ' Break up the [strAdd]
        arAdd = Split(strAdd, strDivider)
        For intI = 0 To UBound(arAdd)
          ' Do we need to check unicity?
          If (Not bUnique) OrElse (Not IsInSemiStack(strOrg, arAdd(intI))) Then
            ' Copy with semicolon
            strOrg &= strDivider & arAdd(intI)
          End If
        Next intI
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modAuto/AddSemiStack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ------------------------------------------------------------------------------------
  ' Name:   IsInSemiStack
  ' Goal:   Check whether the item [strItem] is part of the semicolon separated stack [strOrg]
  ' History:
  ' 30-06-2010  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function IsInSemiStack(ByRef strOrg As String, ByVal strItem As String, _
                                Optional ByVal strDivider As String = ";") As Boolean
    Dim arThis() As String  ' Where we unpack the stack
    Dim intI As Integer     ' Counter

    Try
      ' Validate
      If (strOrg = "") OrElse (strItem = "") Then Return False
      ' Try a possibly faster method
      If (InStr(strOrg, strDivider) = 0) Then
        ' There is just one item
        Return (strOrg = strItem)
      Else
        ' There is a list of items --> Unpack them
        arThis = Split(strOrg, strDivider)
        ' Check all elements
        For intI = 0 To UBound(arThis)
          ' Check this item
          If (arThis(intI) Like strItem) Then Return True
        Next intI
        ' Return failure
        Return False
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("modUtil/IsInSemiStack error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function
  '---------------------------------------------------------------------------------------------------------
  ' Name:       GetFileInDirectory()
  ' Goal:       Locate file [strFile] in directory [strDir] or one of its subdirectories
  '             Return the resulting full filename in [strFullFile] if it is found
  ' History:
  ' 08-07-2013  ERK Created
  '---------------------------------------------------------------------------------------------------------
  Public Function GetFileInDirectory(ByVal strDir As String, ByVal strFile As String, ByRef strFullFile As String) As Boolean
    Dim intI As Integer       ' Counter
    Dim strExtn As String     ' Extension
    Dim arFile() As String    ' Array of possibilities
    Dim strFileOnly As String ' Only the filename of [strFile]

    Try
      ' Validate
      If (strDir = "") Then Return False
      If (strFile = "") Then Return False
      ' Check existence of directory
      If (Not IO.Directory.Exists(strDir)) Then Return False
      ' Get the extension
      strExtn = IO.Path.GetExtension(strFile)
      strFileOnly = IO.Path.GetFileName(strFile)
      ' Look for files
      arFile = IO.Directory.GetFiles(strDir, "*" & strExtn, IO.SearchOption.AllDirectories)
      For intI = 0 To arFile.Length - 1
        ' Access this one
        strFullFile = arFile(intI)
        ' Check if this is the one
        If (IO.Path.GetFileName(strFullFile) = strFileOnly) Then
          ' This is the one!
          Return True
        End If
      Next intI
      ' We did not find it...
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("modUtil/GetFileInDirectory error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    End Try
  End Function
  '----------------------------------------------------------------------------------------------------------
  ' Name:       TryAppendLog()
  ' Goal:       Log user activity
  '             Pass on user/computer/projtype/CRPname
  ' History:
  ' 21-10-2013  ERK Created
  '----------------------------------------------------------------------------------------------------------
  Public Function TryAppendLog(ByVal strPrjType As String, ByVal strCrpName As String) As Boolean
    Dim strParams As String     ' Parameters to pass on
    Dim strFile As String = ""  ' Log file name
    Dim urlFile As String = "http://erwinkomen.ruhosting.nl/software/Cesax/"
    Dim strUser As String = My.User.Name
    Dim strComp As String = My.Computer.Name
    Dim reqThis As System.Net.WebRequest
    Dim resThis As System.Net.WebResponse

    Try
      ' Validate
      If (My.Settings.GenActionInet <> "True") Then Return True
      If (Not My.Computer.Network.IsAvailable) Then Return True
      ' Repair project type
      If (strPrjType = "") Then strPrjType = "(undetermined)"
      ' Construct a set of parameters
      strParams = "act.php?user=" & strUser & "&comp=" & strComp & "&type=" & strPrjType & "&crp=" & strCrpName
      ' Navigate to the right page
      reqThis = System.Net.WebRequest.Create(urlFile & strParams)
      reqThis.Method = "GET"
      reqThis.Timeout = 1000
      ' Send the request
      Try
        resThis = reqThis.GetResponse()
      Catch ex As Exception
        ' There is no need to add on any exception!
        Status("request timed out")
      End Try
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modUtil/TryAppendLog error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Delete thefile, if it exists
      If (strFile <> "") AndAlso (IO.File.Exists(strFile)) Then IO.File.Delete(strFile)
      ' return failure
      Return False
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  DownloadHtml
  ' Goal :  Download an HTML page from the internet
  ' History:
  ' 14-01-2014  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function DownloadHtml(ByVal strSrc As String, ByVal strDst As String, _
           ByRef wbThis As WebBrowser, Optional ByVal bVerbose As Boolean = True) As Boolean
    Dim smRemote As IO.Stream = Nothing
    Dim smLocal As IO.StreamWriter = Nothing
    Dim rdThis As IO.StreamReader = Nothing
    Dim strBack As String
    Dim buffer As Byte() = New Byte(1023) {}
    Dim intBytesRead As Integer
    Dim intBytesProcessed As Integer = 0

    Try
      ' Validate
      If (strSrc = "") Then Return False
      ' Navigate browser to the page
      With wbThis
        ' Navigate to it
        .Navigate(New Uri(strSrc))
        ' Wait until we get there
        While (.ReadyState <> WebBrowserReadyState.Complete)
          ' Wait until we get there
          Application.DoEvents()
          ' Debug.Print(wbThis.StatusText)
        End While
        ' Read the content with a streamreader
        smRemote = .DocumentStream
        rdThis = New IO.StreamReader(smRemote, System.Text.Encoding.GetEncoding("iso-8859-1"))
        smRemote.Position = 0
        strBack = rdThis.ReadToEnd
        rdThis.Close()
        ' Write the text away as UTF8
        IO.File.WriteAllText(strDst, strBack, System.Text.Encoding.UTF8)
      End With
      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modUtil/DownloadHtml error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    Finally
      If (smRemote IsNot Nothing) Then smRemote.Close()
      If (smLocal IsNot Nothing) Then smLocal.Close()
    End Try
  End Function
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  DownloadFile
  ' Goal :  Download a file from the internet
  ' History:
  ' 14-01-2014  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function DownloadFile(ByVal strSrc As String, ByVal strDst As String, _
                               Optional ByVal bVerbose As Boolean = True) As Boolean
    Dim wclThis As WebClient = New WebClient
    Dim strUri As New System.Uri(strSrc)
    Dim reqThis As System.Net.WebRequest = Nothing
    Dim resThis As System.Net.WebResponse = Nothing
    Dim smRemote As IO.Stream = Nothing
    Dim smLocal As IO.Stream = Nothing
    Dim buffer As Byte() = New Byte(1023) {}
    Dim intBytesRead As Integer
    Dim intBytesProcessed As Integer = 0
    Dim intWait As Integer = 20000    ' Millisecons waiting time

    Try
      ' Validate
      If (strSrc = "") Then Return False
      ' Check if we are connected to the internet
      If (Not My.Computer.Network.IsAvailable) Then Return False

      ' Check if the file exists on the server
      Try
        reqThis = System.Net.WebRequest.Create(strSrc)
      Catch e As WebException
        Stop
      Catch ex As Exception
        ' If there is an error, we will not complain
        If (bVerbose) Then Logging("Downloadfile Webrequest returns: " & ex.Message)
        Return False
      End Try
      ' Check what we have
      If (reqThis Is Nothing) Then
        If (bVerbose) Then Logging("Downloadfile: webrequest returns zero. File=" & strSrc)
        Return False
      End If
      ' Set the timeout
      reqThis.Timeout = 15000
      reqThis.ImpersonationLevel = System.Security.Principal.TokenImpersonationLevel.Anonymous
      ' Check the request - HOW??
      'Try
      '  resThis = reqThis.GetResponse
      'Catch ex As Exception
      '  ' If there is an error, we will not complain
      '  If (bVerbose) Then Logging("Downloadfile: problem getting a respons" & vbCrLf & _
      '    "   There was a problem downloading: " & strSrc)
      '  Return False
      'End Try
      ' Loop until we have a positive response
      Do
        Try
          resThis = reqThis.GetResponse
        Catch e As WebException
          Dim webResp As HttpWebResponse
          webResp = CType(e.Response, HttpWebResponse)
          If (webResp Is Nothing) Then
            ' If there is an error, we will not complain
            If (bVerbose) Then Logging("Downloadfile: problem getting a respons" & vbCrLf & _
              "   There was a problem downloading: " & strSrc & vbCrLf & e.Message)
            ' No need to take further action
            Return False
          Else
            If (webResp.StatusCode = 503) Then
              ' If this is [503] then wait... for some seconds...
              Status("Waiting ...")
              System.Threading.Thread.Sleep(intWait)
            Else
              If (bVerbose) Then Logging("DownloadFile problem:" & vbCrLf & _
                          " Status code = " & webResp.StatusCode & vbCrLf & _
                          " Status description: " & webResp.StatusDescription)
            End If
          End If
        Catch ex As Exception
          ' If there is an error, we will not complain
          If (bVerbose) Then Logging("Downloadfile: problem getting a respons" & vbCrLf & _
            "   There was a problem downloading: " & strSrc & vbCrLf & ex.Message)
        End Try
        ' Do we need to wait?
        If (resThis Is Nothing) Then
          ' Give way for others
          Application.DoEvents()
          ' Issue the request afresh
          reqThis = System.Net.WebRequest.Create(strSrc)
          ' Show what we have done
          Logging("Issued new request at: " & strSrc)
        End If
      Loop While (resThis Is Nothing)
      If (resThis IsNot Nothing) Then
        ' Get stream object for remote file
        Try
          smRemote = resThis.GetResponseStream
        Catch ex As Exception
          If (bVerbose) Then Logging("DownloadFile: could not get stream")
          Return False
        End Try
        ' Start creating local file
        smLocal = IO.File.Create(strDst)
        ' Loop 
        Do
          ' Read block from buffer
          Try
            intBytesRead = smRemote.Read(buffer, 0, buffer.Length)
          Catch ex As Exception
            If (bVerbose) Then Logging("DownloadFile: could not read buffer")
            Return False
          End Try
          ' Write block to buffer
          smLocal.Write(buffer, 0, intBytesRead)
          ' Keep track of total
          intBytesProcessed += intBytesRead
        Loop While (intBytesRead > 0)

      End If

      ' Return success
      Return True
    Catch ex As Exception
      ' Show error
      HandleErr("modUtil/DownloadFile error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' return failure
      Return False
    Finally
      If (resThis IsNot Nothing) Then resThis.Close()
      If (smRemote IsNot Nothing) Then smRemote.Close()
      If (smLocal IsNot Nothing) Then smLocal.Close()
    End Try
  End Function
End Module
