Option Strict On
Option Explicit On
Option Compare Binary
Public Class StringColl
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  StringColl
  ' Goal :  A collection of strings 
  ' History:
  ' 20-02-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  ' =================================LOCAL CONSTANTS==========================================================
  Private Const CHUNK As Integer = 512            ' Size of one chunk to be added with a ReDim operation
  Private Const BUF_SIZE As Integer = 2047
  ' =================================LOCAL VARIABLES==========================================================
  Private arList(0 To CHUNK - 1) As String         ' List of strings
  Private arFreq(0 To CHUNK - 1) As Integer       ' Frequency of occurrance
  Private arExmp(0 To CHUNK - 1) As String        ' List of examples
  Private loc_intNum As Integer = 0               ' Number of elements in [arList]
  Private arBuffer(0 To BUF_SIZE) As String ' String buffer
  Private bBusy As Boolean = False                ' A process is busy with me that cannot be interruped
  ' ==========================================================================================================
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Clear
  ' Goal :  Reset the list of strings
  ' History:
  ' 20-02-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Sub Clear()
    ReDim arList(0 To CHUNK - 1)
    ReDim arFreq(0 To CHUNK - 1)
    ReDim arExmp(0 To CHUNK - 1)
    loc_intNum = 0
  End Sub
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Add
  ' Goal :  Add one string to the collection
  ' History:
  ' 20-02-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Sub Add(ByVal strEntry As String, Optional ByVal strExmp As String = "")
    Try
      ' Wait until we have exclusive access
      While (bBusy)
        ' Give room for others
        Application.DoEvents()
        ' Check on interrupt
        If (bInterrupt) Then Exit Sub
      End While
      ' Gain exclusive access
      bBusy = True
      ' Access it
      arList(loc_intNum) = strEntry
      ' Reset the frequency
      arFreq(loc_intNum) = 1
      ' Set the example
      arExmp(loc_intNum) = strExmp
      ' Increment counter
      loc_intNum += 1
      ' Check whether more space is needed
      If (loc_intNum > UBound(arList)) Then
        ' Add another chunk
        ReDim Preserve arList(0 To UBound(arList) + CHUNK)
        ReDim Preserve arFreq(0 To UBound(arFreq) + CHUNK)
        ReDim Preserve arExmp(0 To UBound(arExmp) + CHUNK)
      End If
      ' Free exclusive access
      bBusy = False
    Catch ex As Exception
      ' Free exclusive access
      bBusy = False
      ' Warn user
      HandleErr("StringColl/Add error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  AddSemi
  ' Goal :  Add strings separated by semicolons to the collection
  ' History:
  ' 20-02-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Sub AddSemi(ByVal strEntry As String, Optional ByVal strExmp As String = "")
    Dim intI As Integer     ' Counter
    Dim arEntry() As String ' String of entries

    Try
      ' Convert input into array
      arEntry = Split(strEntry, ";")
      ' Process them one by one
      For intI = 0 To arEntry.Length - 1
        Add(arEntry(intI))
      Next intI
    Catch ex As Exception
      ' Show error
      HandleErr("StringColl/AddSemi error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  AddUnique
  ' Goal :  Add one string to the collection if it is not on there already.
  '         If it is on there, increase its frequency
  ' History:
  ' 20-02-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Sub AddUnique(ByVal strEntry As String, Optional ByVal strExmp As String = "")
    Dim intId As Integer  ' Position where the entry might be

    ' Try find it...
    intId = Find(strEntry)
    ' Was it found?
    If (intId >= 0) Then
      ' Add to its frequency
      arFreq(intId) += 1
    Else
      ' Just add it
      Add(strEntry, strExmp)
    End If
  End Sub
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Flush
  ' Goal :  Append the text we have to the file given
  ' History:
  ' 25-09-2013  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function Flush(ByVal strFile As String) As Boolean
    Dim intI As Integer       ' Counter
    Dim wriThis As IO.StreamWriter

    Try
      ' Open a streamwriter for appending
      wriThis = New IO.StreamWriter(strFile, True)
      ' Walk the current buffer
      For intI = 0 To loc_intNum - 1
        ' Append string
        wriThis.WriteLine(arList(intI))
      Next intI
      ' Close streamwriter
      wriThis.Close()
      ' Return success
      Return True
    Catch ex As Exception
      ' Show the error
      HandleErr("StringColl/Save error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  MostFreqIdx
  ' Goal :  Get the index of the value with the highest frequency
  ' History:
  ' 20-02-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public ReadOnly Property MostFreqIdx() As Integer
    Get
      Dim intI As Integer         ' counter
      Dim intFreq As Integer = 0  ' Start frequency
      Dim intIdx As Integer = 0   ' Index of the best match

      ' Visit all
      For intI = 0 To UBound(arFreq)
        If (arFreq(intI) > intFreq) Then
          ' Found a higher one
          intFreq = arFreq(intI)
          intIdx = intI
        End If
      Next intI
      ' Return the index of the most frequent one
      MostFreqIdx = intIdx
    End Get
  End Property
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Find
  ' Goal :  Find out if this element is in the collection and if so, return its index
  ' History:
  ' 20-02-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function Find(ByVal strEntry As String) As Integer
    Dim intI As Integer           ' Counter

    ' Go through the [arList] array
    For intI = 1 To loc_intNum
      ' Is this the correct string?
      If (arList(intI - 1) = strEntry) Then
        ' Return success
        Return intI - 1
      End If
    Next intI
    ' If we have come here, it means there is failure
    Find = -1
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  FindInSemiList
  ' Goal :  Look in all elements, and return the index of the entry, where the given string 
  '           is found between ; signs
  ' History:
  ' 09-07-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function FindInSemiList(ByVal strEntry As String) As Integer
    Dim intI As Integer           ' Counter

    ' Adapt the entry
    strEntry = ";" & strEntry & ";"
    ' Go through the [arList] array
    For intI = 1 To loc_intNum
      ' Is this the correct string?
      If (InStr(arList(intI - 1), strEntry) > 0) Then
        ' Return success
        Return intI - 1
      End If
    Next intI
    ' If we have come here, it means there is failure
    FindInSemiList = -1
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Exists
  ' Goal :  Check if the element already is in the collection
  ' History:
  ' 20-02-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function Exists(ByVal strEntry As String) As Boolean
    Dim intI As Integer           ' Counter

    ' Go through the [arList] array
    For intI = 1 To loc_intNum
      ' Is this the correct string?
      If (arList(intI - 1) = strEntry) Then
        ' Return success
        Return True
      End If
    Next intI
    ' If we have come here, it means there is failure
    Exists = False
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Count
  ' Goal :  Get the total amount of strings in the collection
  ' History:
  ' 20-02-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function Count() As Integer
    Count = loc_intNum
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Item
  ' Goal :  The content of each item can be retrieved or changed even
  ' History:
  ' 20-02-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Property Item(ByVal intId As Integer) As String
    Get
      ' Check validity of [intId]
      If (intId >= 0) AndAlso (intId < loc_intNum) Then
        ' Retrieve entry
        Item = arList(intId)
      Else
        Item = ""
      End If
    End Get
    Set(ByVal strItem As String)
      ' Check validity of [intId]
      If (intId >= 0) AndAlso (intId < loc_intNum) Then
        ' Set entry
        arList(intId) = strItem
      End If
    End Set
  End Property
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Exmp
  ' Goal :  The content of each Example can be retrieved or changed even
  ' History:
  ' 20-02-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Property Exmp(ByVal intId As Integer) As String
    Get
      ' Check validity of [intId]
      If (intId >= 0) AndAlso (intId < loc_intNum) Then
        ' Retrieve entry
        Exmp = arExmp(intId)
      Else
        Exmp = ""
      End If
    End Get
    Set(ByVal strExmp As String)
      ' Check validity of [intId]
      If (intId >= 0) AndAlso (intId < loc_intNum) Then
        ' Set entry
        arExmp(intId) = strExmp
      End If
    End Set

  End Property
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Text
  ' Goal :  The complete text (of arList)
  ' History:
  ' 02-03-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public ReadOnly Property Text() As String
    Get
      Dim strOut As String = "" ' Output collector
      Dim arOut() As String     ' Better array

      Try
        ' Wait until we have exclusive access
        While (bBusy)
          ' Give room for others
          Application.DoEvents()
          ' Check on interrupt
          If (bInterrupt) Then Return ""
        End While
        ' Gain exclusive access
        bBusy = True
        ' Copy the output to a better array
        ReDim arOut(0 To loc_intNum)
        Array.Copy(arList, arOut, loc_intNum + 1)
        ' Combine the output
        Text = Join(arOut, vbCrLf)
        '' Add all to output
        'For intI = 1 To loc_intNum
        '  strOut &= arList(intI - 1) & vbCrLf
        'Next intI
        'Text = strOut
        ' Give up exclusive access
        bBusy = False
      Catch ex As Exception
        ' Warn user
        HandleErr("StringColl/Text error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
        ' Give up exclusive access
        bBusy = False
        ' Return failure
        Return ""
      End Try
    End Get
  End Property
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Semi
  ' Goal :  The complete text (of arList), but separated by semicolons
  ' History:
  ' 14-03-2013  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public ReadOnly Property Semi() As String
    Get
      Dim strOut As String = "" ' Output collector
      Dim arOut() As String     ' Better array

      ' Copy the output to a better array
      ReDim arOut(0 To loc_intNum - 1)
      Array.Copy(arList, arOut, loc_intNum)
      ' Combine the output
      strOut = Join(arOut, ";")
      ' Return the result
      Return strOut
    End Get
  End Property
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Html
  ' Goal :  The complete text (of arList)
  ' History:
  ' 02-03-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public ReadOnly Property Html(Optional ByVal strDelim As String = "/") As String
    Get
      Dim strOut As String = "" ' Output collector
      Dim arOne() As String     ' Array with the [arList] item broken up in parts divided by /
      Dim intI As Integer       ' Counter
      Dim intPtc As Integer     ' Counter

      Try
        ' Start HTML
        strOut = "<html><body><table>" & vbCrLf
        ' Add all to output
        For intI = 1 To loc_intNum
          ' Where are we?
          intPtc = (intI * 100) \ loc_intNum
          Status("Html " & intPtc & "%", intPtc)
          ' Continue...
          arOne = Split(arList(intI - 1), strDelim)
          strOut &= "<tr><td>" & arOne(0) & "</td>" & _
            "<td>" & arFreq(intI - 1) & "</td>"
          ' Possibly add example
          If (arExmp(intI - 1) <> "") Then
            strOut &= "<td>" & arExmp(intI - 1) & "</td>"
          End If
          ' Possibly add category
          If (arOne.Length > 1) Then
            strOut &= "<td>" & arOne(1) & "</td>"
          End If
          ' Finish this row
          strOut &= "</tr>" & vbCrLf
        Next intI
        ' Finish HTML
        strOut &= "</table></body></html>" & vbCrLf
        ' Return result
        Html = strOut
      Catch ex As Exception
        ' Show there is an error
        HandleErr("StringColl/Html error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
        ' Return failure
        Return ""
      End Try
    End Get
  End Property
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  FreqTable
  ' Goal :  A table with the item and its frequency in HTML form
  ' History:
  ' 02-03-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public ReadOnly Property FreqTable() As String
    Get
      Dim strOut As String = "" ' Output collector
      Dim arOne() As String     ' Array with the [arList] item broken up in parts divided by /
      Dim intI As Integer       ' Counter

      Try
        ' Start HTML
        strOut = "<table>" & vbCrLf
        ' Add all to output
        For intI = 1 To loc_intNum
          arOne = Split(arList(intI - 1), "/")
          strOut &= "<tr><td>" & arOne(0) & "</td>" & _
            "<td>" & arFreq(intI - 1) & "</td>"
          ' Possibly add example
          If (arExmp(intI - 1) <> "") Then
            strOut &= "<td>" & arExmp(intI - 1) & "</td>"
          End If
          ' Possibly add category
          If (arOne.Length > 1) Then
            strOut &= "<td>" & arOne(1) & "</td>"
          End If
          ' Finish this row
          strOut &= "</tr>" & vbCrLf
        Next intI
        ' Finish HTML
        strOut &= "</table>" & vbCrLf
        ' Return result
        Return strOut
      Catch ex As Exception
        ' Show there is an error
        HandleErr("StringColl/FreqTable error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
        ' Return failure
        Return ""
      End Try
    End Get
  End Property
  '' ----------------------------------------------------------------------------------------------------------
  '' Name :  SortedText
  '' Goal :  The complete text (of arList)
  '' History:
  '' 02-03-2009  ERK Created
  '' ----------------------------------------------------------------------------------------------------------
  'Public ReadOnly Property SortedText() As String
  '  Get
  '    Dim strOut As String = "" ' Output collector
  '    Dim intI As Integer       ' Counter

  '    ' Sort the array
  '    Array.Sort(arList, 0, loc_intNum)
  '    ' Add all to output
  '    For intI = 1 To loc_intNum
  '      strOut &= arList(intI - 1) & vbCrLf
  '    Next intI
  '    SortedText = strOut
  '  End Get
  'End Property
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Freq
  ' Goal :  The frequency of an item can be retrieved
  ' History:
  ' 20-02-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public ReadOnly Property Freq(ByVal intId As Integer) As Integer
    Get
      ' Check validity of [intId]
      If (intId >= 0) AndAlso (intId < loc_intNum) Then
        ' Retrieve entry
        Freq = arFreq(intId)
      Else
        Freq = 0
      End If
    End Get
    '  Set(ByVal value As Integer)
    '    ' Check validity of [intId]
    '    If (intId >= 0) AndAlso (intId < loc_intNum) Then
    '      ' Set the frequency
    '      arFreq(intId) = value
    '    End If
    '  End Set
  End Property
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  DelString
  ' Goal :  Remove the [intId] from the [arList] and adapt the list and the total number [loc_intNum]
  ' History:
  ' 20-02-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Sub DelString(ByVal strEntry As String)
    Dim intI As Integer           ' Counter

    ' Try and find it
    intI = Find(strEntry)
    ' Try delete it
    DelItem(intI)
  End Sub
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  DelItem
  ' Goal :  Remove the [intId] from the [arList] and adapt the list and the total number [loc_intNum]
  ' History:
  ' 20-02-2009  ERK Created
  ' 16-09-2014  ERK Bugfix inside [for inti] loop
  ' ----------------------------------------------------------------------------------------------------------
  Public Sub DelItem(ByVal intId As Integer)
    Dim intI As Integer           ' Counter

    Try
      ' Check if this item is in the correct range
      If (intId < 0) OrElse (intId > loc_intNum) Then
        ' No need to do anything
        Exit Sub
      End If
      ' Go through the [arList] array
      For intI = intId To loc_intNum - 1
        ' Move next one down
        arList(intI) = arList(intI + 1)
        ' Also for frequency
        arFreq(intI) = arFreq(intI + 1)
        ' Also for example
        arExmp(intI) = arExmp(intI + 1)
      Next intI
      ' Reset the top one to zero - just in case
      arList(loc_intNum - 1) = ""
      arFreq(loc_intNum - 1) = 0
      arExmp(loc_intNum - 1) = ""
      ' Adjust the maximum number of [arList]
      loc_intNum -= 1
    Catch ex As Exception
      ' Show the error
      HandleErr("StringColl/DelItem error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Class
