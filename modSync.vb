Module modSync
  Private Declare Auto Function SendScrollPostMessage Lib "user32.dll" Alias "SendMessage" ( _
    ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByRef lParam As Point) As Integer
  Private Const WM_USER_area = &H400
  Private Const EM_GETSCROLLPOS = WM_USER_area + 221
  Private Const EM_SETSCROLLPOS = WM_USER_area + 222
  ' ------------------------------------------------------------------------------------
  ' Name:   GetVpos
  ' Goal:   Get the vertical scroll position from [rtbThis]
  ' History:
  ' 25-02-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Function GetVpos(ByRef rtbThis As RichTextBox) As Integer
    Dim pntThis As Point      ' One point

    Try
      ' Validate
      If (rtbThis Is Nothing) Then Return 0
      ' Post the request
      SendScrollPostMessage(rtbThis.Handle, EM_GETSCROLLPOS, New IntPtr(0), pntThis)
      ' Return the result
      Return pntThis.Y
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSync/GetVpos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
      ' Return failure
      Return 0
    End Try
  End Function
  ' ------------------------------------------------------------------------------------
  ' Name:   SetVpos
  ' Goal:   Set the vertical scroll position in [rtbThis]
  ' History:
  ' 25-02-2011  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub SetVpos(ByRef rtbThis As RichTextBox, ByVal intY As Integer)
    Dim pntThis As Point      ' One point

    Try
      ' Validate
      If (rtbThis Is Nothing) Then Exit Sub
      ' Get the position as it is right now
      SendScrollPostMessage(rtbThis.Handle, EM_GETSCROLLPOS, New IntPtr(0), pntThis)
      ' Change the y-value
      pntThis.Y = intY
      ' Post the request
      SendScrollPostMessage(rtbThis.Handle, EM_SETSCROLLPOS, New IntPtr(0), pntThis)
    Catch ex As Exception
      ' Warn the user
      HandleErr("modSync/setVpos error: " & ex.Message & vbCrLf & ex.StackTrace & vbCrLf)
    End Try
  End Sub
End Module
