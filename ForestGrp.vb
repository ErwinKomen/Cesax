Option Explicit On
Public Class ForestGrp
  ' ===================================== LOCAL VARIABLES ====================================================
  Private loc_strInFile As String = ""              ' Name of the file to read from
  Private loc_strOutFile As String = ""             ' Name of the temporary file written to
  Private loc_filRead As IO.StreamReader = Nothing  ' Reader for this file
  Private loc_filOut As IO.StreamWriter = Nothing   ' Writer for temporary output
  Private loc_intSize As Integer = 1                ' Size of the file
  Private loc_intPos As Integer = 0                 ' Where we are reading the file
  Private loc_strTag As String = ""                 ' The tag to be used
  Private Const DEFAULT_TAG As String = "forestGrp"
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Load
  ' Goal :  Store the name of the file to be used
  ' History:
  ' 25-03-2010  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public WriteOnly Property Load() As String
    Set(ByVal value As String)
      Dim infThis As IO.FileInfo = Nothing

      ' Get the name of the input file
      loc_strInFile = value
      ' Try to open it for reading
      loc_filRead = New IO.StreamReader(loc_strInFile)
      ' Reset the reading
      loc_intPos = 0
      ' Determine the size of the file
      infThis = New IO.FileInfo(loc_strInFile)
      loc_intSize = infThis.Length
    End Set
  End Property
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Close
  ' Goal :  Close the input file, releasing it
  ' History:
  ' 20-04-2010  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Sub Close()
    loc_filRead.Close()
  End Sub
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Tag
  ' Goal :  The tag to be used to recognize the start and the end of the sections to be read
  ' History:
  ' 26-03-2010  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Property Tag() As String
    Get
      Tag = loc_strTag
    End Get
    Set(ByVal value As String)
      loc_strTag = Trim(value)
    End Set
  End Property
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  File
  ' Goal :  Give the name of the temporary output file that is used
  ' History:
  ' 25-03-2010  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public ReadOnly Property File() As String
    Get
      ' Do we need to make a temporary output file?
      If (loc_strOutFile = "") Then CreateTempOut()
      ' Return the name of the temporary output file
      File = loc_strOutFile
    End Get
  End Property
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Ptc
  ' Goal :  Give the percentage as to where we are reading the input
  ' History:
  ' 25-03-2010  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public ReadOnly Property Ptc() As Integer
    Get
      Try
        Ptc = (100 * (loc_intPos / loc_intSize))
        ' Make sure the maximum is not reached
        If (Ptc > 100) Then Ptc = 100
      Catch ex As Exception
        ' Show the exception
        HandleErr("ForestGrp/Ptc error: " & ex.Message & vbCrLf)
        ' Return something anyway
        Ptc = 0
      End Try
    End Get
  End Property
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  GetNext
  ' Goal :  Read the next <forestGrp> section
  ' History:
  ' 25-03-2010  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function GetNext() As Boolean
    Dim bGrpStart As Boolean = False  ' Detected the start of the forestgroup
    Dim bGrpEnd As Boolean = False    ' Detected the end of the forestgroup section
    Dim strLine As String = ""        ' One line that has been read from the input

    Try
      ' Check if we have a tag
      If (loc_strTag = "") Then
        ' Take the default tag
        loc_strTag = DEFAULT_TAG
      End If
      ' Do we need to make a temporary output file?
      If (loc_strOutFile = "") Then
        CreateTempOut()
      Else
        ' Make sure the temporary output file is made empty
        loc_filOut = New IO.StreamWriter(loc_strOutFile)
      End If
      ' Is there still something to be read?
      If (loc_filRead.EndOfStream) Then
        ' Return failure
        Return False
      Else
        ' Read until we have <forestGrp>
        Do
          ' Read a line
          strLine = loc_filRead.ReadLine
          ' Increment the size
          loc_intPos += strLine.Length + 1
          ' Note the start of the group
          bGrpStart = (InStr(strLine, "<" & loc_strTag & ">") > 0) OrElse _
                      (InStr(strLine, "<" & loc_strTag & " ") > 0)
        Loop While (Not bGrpStart) AndAlso (Not loc_filRead.EndOfStream)
        ' Did we get it?
        If (Not bGrpStart) Then
          ' Finalize the output
          loc_filOut.Flush()
          loc_filOut.Close()
          ' We didn't get a start, so return false
          Return False
        End If
        ' Copy input to output until the end of the <forestGrp>
        While Not (bGrpEnd) AndAlso (Not loc_filRead.EndOfStream)
          ' Copy line to the output
          loc_filOut.WriteLine(strLine)
          ' Read a line
          strLine = loc_filRead.ReadLine
          ' Increment the size
          loc_intPos += strLine.Length + 1
          ' Determine whether we got to the end
          bGrpEnd = (InStr(strLine, "</" & loc_strTag & ">") > 0)
        End While
        ' Did we get the end of the forestgrp?
        If (bGrpEnd) Then
          ' Flush it
          loc_filOut.WriteLine(strLine)
        Else
          ' Make one ourselves
          loc_filOut.WriteLine("</" & loc_strTag & ">")
        End If
        ' Flush all the output and close it
        loc_filOut.Flush()
        loc_filOut.Close()
      End If
      ' Return success
      Return True
    Catch ex As Exception
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Clear
  ' Goal :  Delete temporary file
  ' History:
  ' 25-03-2010  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Sub Clear()
    ' Is there a filename?
    If (loc_strOutFile <> "") Then
      ' Try delete it
      IO.File.Delete(loc_strOutFile)
    End If
  End Sub
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  CreateTempOut
  ' Goal :  Create a temporary output file
  ' History:
  ' 25-03-2010  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Private Sub CreateTempOut()
    ' Create a temporary output file name
    loc_strOutFile = IO.Path.GetTempFileName
    ' Open it for writing
    loc_filOut = New IO.StreamWriter(loc_strOutFile)
  End Sub

  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Finalize
  ' Goal :  Close the class object by removing temporary stuff
  ' History:
  ' 25-03-2010  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Protected Overrides Sub Finalize()
    ' Delete the temporary output file
    If (loc_strOutFile <> "") Then
      IO.File.Delete(loc_strOutFile)
    End If
    ' Perform the standard finalizations
    MyBase.Finalize()
  End Sub
End Class
