Imports System.Windows.Forms

Public Class frmMetaInfo
  ' =================================== LOCAL VARIABLES =================================================
  Private loc_strMethod As String = "columns"   ' Method that has been chosen
  Private loc_strFile As String = ""            ' File with txt/csv
  Private loc_strDir As String = ""             ' Directory with the psdx files
  ' =====================================================================================================
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   FileName, PsdxDir, Method
  ' Goal:   Return the correct properties to the user
  ' History:
  ' 03-10-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Public ReadOnly Property FileName() As String
    Get
      Return loc_strFile
    End Get
  End Property
  Public ReadOnly Property PsdxDir() As String
    Get
      Return loc_strDir
    End Get
  End Property
  Public ReadOnly Property Method() As String
    Get
      Return loc_strMethod
    End Get
  End Property
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   OK_Button_Click, Cancel_Button_Click
  ' Goal:   Close the form and return OK or Cancel
  ' History:
  ' 03-10-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub
  Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdMetaFile_Click
  ' Goal:   Elicit the txt/csv file
  ' History:
  ' 03-10-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdMetaFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMetaFile.Click
    Try
      With Me.OpenFileDialog1
        ' Set initial directory to the working directory of Cesax
        .InitialDirectory = strWorkDir
        .Title = "Locate the txt/csv file containing the meta data"
        ' Note the file type to open
        .Filter = "Tab-separated text file (*.txt;*.csv)|*.txt;*.csv"
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK, Windows.Forms.DialogResult.Yes
            loc_strFile = .FileName
            Me.tbMetaFile.Text = loc_strFile
        End Select
      End With
    Catch ex As Exception

    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   cmdMetaDir_Click
  ' Goal:   Elicit the directory
  ' History:
  ' 03-10-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub cmdMetaDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMetaDir.Click
    Try
      With Me.FolderBrowserDialog1
        .SelectedPath = strWorkDir
        .Description = "Provide the directory containing the .psdx files"
        Select Case .ShowDialog
          Case Windows.Forms.DialogResult.OK, Windows.Forms.DialogResult.Yes
            loc_strDir = .SelectedPath
            Me.tbMetaDir.Text = loc_strDir
        End Select
      End With
    Catch ex As Exception

    End Try
  End Sub
  ' ------------------------------------------------------------------------------------------------------------
  ' Name:   rbFieldColumns_CheckedChanged, rbFieldAlternated_CheckedChanged
  ' Goal:   Note changes in the method chosen
  ' History:
  ' 03-10-2014  ERK Created
  ' ------------------------------------------------------------------------------------------------------------
  Private Sub rbFieldColumns_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbFieldColumns.CheckedChanged
    loc_strMethod = "columns"
  End Sub
  Private Sub rbFieldAlternated_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbFieldAlternated.CheckedChanged
    loc_strMethod = "alternated"
  End Sub
End Class
