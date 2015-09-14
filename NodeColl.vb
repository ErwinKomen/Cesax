Option Strict On
Option Explicit On
Option Compare Binary
Imports System.Xml
Public Class NodeColl
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Stack
  ' Goal :  A collection of XML <eTree> nodes
  ' History:
  ' 20-02-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  ' =================================LOCAL CONSTANTS==========================================================
  Private Const CHUNK As Integer = 512            ' Size of one chunk to be added with a ReDim operation
  ' =================================LOCAL VARIABLES==========================================================
  Private arList(0 To CHUNK - 1) As XmlNode ' List of XML nodes
  Private arName(0 To CHUNK - 1) As String  ' List of names
  Private loc_intNum As Integer = 0         ' Number of elements in [arList]
  ' ==========================================================================================================
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Clear
  ' Goal :  Reset the list of strings
  ' History:
  ' 20-02-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Sub Clear()
    ReDim arList(0 To CHUNK - 1) : ReDim arName(0 To CHUNK - 1)
    loc_intNum = 0
  End Sub
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Add
  ' Goal :  Add one string to the collection
  ' History:
  ' 20-02-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Sub Add(ByRef xndThis As XmlNode, Optional ByVal strName As String = "")
    Try
      ' Access it
      arList(loc_intNum) = xndThis : arName(loc_intNum) = strName
      ' Increment counter
      loc_intNum += 1
      ' Check whether more space is needed
      If (loc_intNum > UBound(arList)) Then
        ' Add another chunk
        ReDim Preserve arList(0 To UBound(arList) + CHUNK)
        ReDim Preserve arName(0 To UBound(arList) + CHUNK)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("NodeColl/Add error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  AddUnique
  ' Goal :  Add one string to the collection if it is not on there already.
  '         If it is on there, increase its frequency
  ' History:
  ' 20-02-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Sub AddUnique(ByRef xndThis As XmlNode)
    Dim intId As Integer  ' Position where the entry might be

    Try
      ' Try find it...
      intId = Find(xndThis)
      ' Was it found?
      If (intId >= 0) Then
        ' Don't add it
      Else
        ' Just add it
        Add(xndThis)
      End If
    Catch ex As Exception
      ' Show error
      HandleErr("NodeColl/AddUnique error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Find
  ' Goal :  Find out if this element is in the collection and if so, return its index
  ' History:
  ' 20-02-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function Find(ByRef xndThis As XmlNode) As Integer
    Dim intI As Integer     ' Counter
    Dim intId As Integer    ' The ID of the node we are looking for

    Try
      ' validate
      If (xndThis Is Nothing) Then Return -1
      If (xndThis.Name = "eTree") Then
        ' Determine the ID of the node we are looking for
        intId = CInt(xndThis.Attributes("Id").Value)
        ' Go through the [arList] array
        For intI = 1 To loc_intNum
          ' Double check
          If (arList(intI - 1) Is Nothing) Then Return -1
          If (arList(intI - 1).Name = "eTree") Then
            ' Is this the correct string?
            If (CInt(arList(intI - 1).Attributes("Id").Value) = intId) Then
              ' Return success
              Return intI - 1
            End If
          End If
        Next intI
      Else
        ' Go through the [arList] array
        For intI = 1 To loc_intNum
          ' New method of checking
          If (arList(intI - 1) Is xndThis) Then
            ' Return success
            Return intI - 1
          End If
        Next intI
      End If
      ' If we have come here, it means there is failure
      Return -1
    Catch ex As Exception
      ' Show error
      HandleErr("NodeColl/Find error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return -1
    End Try
  End Function
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  Exists
  ' Goal :  Check if the element already is in the collection
  ' History:
  ' 20-02-2009  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Function Exists(ByRef xndThis As XmlNode) As Boolean
    Dim intI As Integer           ' Counter
    Dim intId As Integer    ' The ID of the node we are looking for

    Try
      ' Validate
      If (xndThis Is Nothing) Then Return False
      ' What do we have?
      If (xndThis.Name = "eTree") Then
        ' Determine the ID of the node we are looking for
        intId = CInt(xndThis.Attributes("Id").Value)
        ' Go through the [arList] array
        For intI = 1 To loc_intNum
          ' Double check
          If (arList(intI - 1) Is Nothing) Then
            ' Return failure
            Return False
          End If
          ' Is this the correct string?
          If (CInt(arList(intI - 1).Attributes("Id").Value) = intId) Then
            ' Return success
            Return True
          End If
        Next intI
      Else
        ' Go through the [arList] array
        For intI = 1 To loc_intNum
          ' New method of checking
          If (arList(intI - 1) Is xndThis) Then
            ' Return success
            Return True
          End If
        Next intI
      End If
      ' If we have come here, it means there is failure
      Return False
    Catch ex As Exception
      ' Show error
      HandleErr("NodeColl/Exists error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return False
    End Try
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
  Public Property Item(ByVal intId As Integer) As XmlNode
    Get
      Try
        ' Check validity of [intId]
        If (intId >= 0) AndAlso (intId < loc_intNum) Then
          ' Retrieve entry
          Item = arList(intId)
        Else
          Item = Nothing
        End If
      Catch ex As Exception
        ' Show error
        MsgBox("NodeColl/Item error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
        ' Return failure
        Return Nothing
      End Try
    End Get
    Set(ByVal xndThis As XmlNode)
      Try
        ' Check validity of [intId]
        If (intId >= 0) AndAlso (intId < loc_intNum) Then
          ' Set entry
          arList(intId) = xndThis
        End If
      Catch ex As Exception
        ' Show error
        HandleErr("NodeColl/Item error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      End Try
    End Set
  End Property
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  ItName
  ' Goal :  Retrieve or change the [Name] of the item at location [intId]
  ' History:
  ' 07-11-2013  ERK Created
  ' ----------------------------------------------------------------------------------------------------------
  Public Property ItName(ByVal intId As Integer) As String
    Get
      Try
        ' Check validity of [intId]
        If (intId >= 0) AndAlso (intId < loc_intNum) Then
          ' Retrieve entry
          Return arName(intId)
        Else
          Return ""
        End If
      Catch ex As Exception
        ' Show error
        MsgBox("NodeColl/ItName error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
        ' Return failure
        Return Nothing
      End Try
    End Get
    Set(ByVal value As String)
      Try
        ' Check validity of [intId]
        If (intId >= 0) AndAlso (intId < loc_intNum) Then
          ' Set entry
          arName(intId) = value
        End If
      Catch ex As Exception
        ' Show error
        HandleErr("NodeColl/ItName error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      End Try
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
        ' Copy the output to a better array
        ReDim arOut(0 To loc_intNum)
        Array.Copy(arList, arOut, loc_intNum + 1)
        ' Combine the output
        Text = Join(arOut, vbCrLf)
      Catch ex As Exception
        ' Show error
        HandleErr("NodeColl/Text error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
        ' Return failure
        Return ""
      End Try
    End Get
  End Property
  ' ----------------------------------------------------------------------------------------------------------
  ' Name :  DelItem
  ' Goal :  Remove the [intId] from the [arList] and adapt the list and the total number [loc_intNum]
  ' History:
  ' 20-02-2009  ERK Created
  ' 16-09-2014  ERK Bugfix in [For intI] loop
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
      Next intI
      ' Reset the top one to zero - just in case
      arList(loc_intNum - 1) = Nothing
      ' Adjust the maximum number of [arList]
      loc_intNum -= 1
    Catch ex As Exception
      ' Show error
      HandleErr("NodeColl/DelItem error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Class
