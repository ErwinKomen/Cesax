Option Strict On
Option Explicit On
Option Compare Binary
Imports System.Text.RegularExpressions
Module modAux
  '---------------------------------------------------------------
  ' Name:       Status()
  ' Goal:       Show status text in appropriate place (statusbar)
  ' History:
  ' ERK   07-07-2004    Created for VB6
  ' ERK   24-06-2006    Adapted for VB2005
  '---------------------------------------------------------------
  Public Sub Status(ByRef varText As String, Optional ByVal intPercent As Integer = -1)
    Dim frmThis As System.Windows.Forms.Form
    Dim sbThis As StatusStrip
    Dim tbProgr As ToolStripProgressBar

    ' Zoek het dichtsbijzijnde formulier op met een statusstrip!
    frmThis = GetClosestForm()
    ' Kijk of dit formulier een statusstrip heeft
    sbThis = GetStatusStrip(frmThis)
    While (sbThis Is Nothing) AndAlso (Not frmThis Is Nothing)
      ' Probeer een parent formulier te vinden
      frmThis = frmThis.ParentForm
      ' Kijk of deze een statusstripg heeft
      sbThis = GetStatusStrip(frmThis)
    End While
    ' Was het resultaat positief?
    If (Not frmThis Is Nothing) Then
      ' plaats de statusboodschap op de statusstrip
      sbThis.Items(0).Text = varText
      ' Get the progressbar, if it exists
      If (sbThis.Items.Count > 1) Then
        tbProgr = CType(sbThis.Items(1), ToolStripProgressBar)
        ' Is there any percentage to show?
        If (intPercent < 0) Then
          ' Hide percentage indicator
          tbProgr.Visible = False
        Else
          ' Show percentage indicator
          tbProgr.Visible = True
          ' Set percentage
          tbProgr.Maximum = 100
          tbProgr.Minimum = 0
          tbProgr.Value = intPercent
        End If
      End If
    End If
    System.Windows.Forms.Application.DoEvents()
PROC_EXIT:
    Exit Sub
ERR_EXIT:
    If (Err.Number = 438) Then GoTo PROC_EXIT
    MsgBox("Status Error: " & Err.Description)
  End Sub
  '---------------------------------------------------------------
  ' Name:       GetClosestForm()
  ' Goal:       Probeer handle voor formulier te krijgen waar we in zitten
  ' History:
  ' 27-06-2006  ERK Created
  '---------------------------------------------------------------
  Public Function GetClosestForm() As Form
    Dim frmThis As System.Windows.Forms.Form
    Dim frmOpen As System.Windows.Forms.Form

    ' Probeer eerst te kijken of er een activeform is
    frmThis = Form.ActiveForm
    If (frmThis Is Nothing) Then
      ' Kijk of er wel een formulier open is
      frmOpen = System.Windows.Forms.Application.OpenForms.Item(0)
      If (Not frmOpen Is Nothing) Then
        ' Ga naar boven totaan het parent formulier
        While (Not frmOpen.ParentForm Is Nothing)
          frmOpen = frmOpen.ParentForm
        End While
        frmThis = frmOpen
      End If
    End If
    ' Geef evt. gevonden formulier terug
    GetClosestForm = frmThis
  End Function
  '---------------------------------------------------------------
  ' Name:       GetStatusStrip()
  ' Goal:       Kijk of het gegeven formulier een StatusStrip bevat
  ' History:
  ' 27-06-2006  ERK Created
  '---------------------------------------------------------------
  Private Function GetStatusStrip(ByVal frmThis As System.Windows.Forms.Form) As StatusStrip
    Dim ctlThis As Control

    If (frmThis Is Nothing) Then
      GetStatusStrip = Nothing
    Else
      For Each ctlThis In frmThis.Controls
        If (TypeOf ctlThis Is System.Windows.Forms.StatusStrip) Then
          GetStatusStrip = CType(ctlThis, StatusStrip)
          Exit Function
        End If
      Next
    End If
    GetStatusStrip = Nothing
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  GetDelim
  ' Goal :  Get an appropriate delimiter
  ' History:
  ' 13-07-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Public Function GetDelim(ByRef strText As String, ByVal ParamArray arDelim() As String) As String
    Dim intI As Integer   'Counter

    ' Is this nonsense?
    If (arDelim.Length = 0) Then
      ' Return without anything
      GetDelim = ""
    End If
    ' Try all elements
    For intI = 0 To UBound(arDelim)
      If (InStr(strText, arDelim(intI)) > 0) Then
        ' This is the one
        Return arDelim(intI)
      End If
    Next intI
    ' Return the first one
    GetDelim = arDelim(0)
  End Function
  '------------------------------------------------------------------------------------------------------------------
  ' Name :  GetDelimDeep
  ' Goal :  Get an appropriate delimiter
  ' History:
  ' 13-07-2009  ERK Created
  ' -----------------------------------------------------------------------------------------------------------------
  Public Function GetDelimDeep(ByRef strText As String) As String
    Dim intCr As Integer  ' Number of CRs
    Dim intLf As Integer  ' Number of LFs

    Try
      ' Validate
      If (Trim(strText) = "") Then Return ""
      ' Get the number of CRs and LFs
      intCr = Regex.Matches(strText, "\r").Count
      intLf = Regex.Matches(strText, "\n").Count
      ' Are they the same?
      If (intCr = 0) Then
        ' It is a LF
        Return vbLf
      ElseIf (intLf = 0) Then
        ' It is CR
        Return vbCr
      Else
        ' Compare the two
        If (intCr = intLf) Then
          ' It is crlf
          Return vbCrLf
        Else
          ' Check which is larger
          If (intCr > intLf) Then
            ' Convert CRLF into 
            strText = strText.Replace(vbCrLf, vbCr)
            Return vbCr
          Else
            ' Convert CRLF into 
            strText = strText.Replace(vbCrLf, vbLf)
            Return vbLf
          End If
        End If
      End If
    Catch ex As Exception
      ' Give error
      HandleErr("modAux/GetDelimDeep error: " & ex.Message & vbCrLf & "Stack:" & vbCrLf & ex.StackTrace)
      ' Return failure
      Return ""
    End Try
  End Function
End Module
