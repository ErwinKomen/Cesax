Module modDgv
  ' ------------------------------------------------------------------------------------
  ' Name:   DgvClear
  ' Goal:   If an instance of a DGV handle exists, then clear it
  ' History:
  ' 21-12-2009  ERK Created
  ' ------------------------------------------------------------------------------------
  Public Sub DgvClear(ByRef objDgv As DgvHandle)
    ' Does the handle still exist?
    If (Not objDgv Is Nothing) Then
      ' Clear it
      objDgv.Kill()
    End If
  End Sub
  ' ---------------------------------------------------------------------------------------------------------
  ' Name :  CtlChanged
  ' Goal :  (1) Process the change immediately in [tdlCrp]
  '         (2) Set dirty flag...
  ' Return: True if the dirty flag needs to be set
  ' History:
  ' 30-11-2009  ERK Created
  ' ---------------------------------------------------------------------------------------------------------
  Public Function CtlChanged(ByRef objEd As DgvHandle, ByRef ctlThis As Control, ByVal bInit As Boolean, _
                         Optional ByRef objChgEd As DgvHandle = Nothing) As Boolean
    ' Are we initialized?
    If (bInit) AndAlso (Not objEd Is Nothing) Then
      ' Make sure changes are processed
      With objEd
        If (Not .IsSelecting) Then
          ' Set the item in the DgvHandle
          .SetItem(ctlThis)
          ' Is a separate [objChgEd] specified?
          If (Not objChgEd Is Nothing) Then
            ' Set this object dirty, so that "Changed" changes
            objChgEd.DoEditDirty(True)
          End If
          ' Set dirty bit
          Return True
        End If
      End With
    End If
    ' Return false for the dirty flag
    Return False
  End Function
End Module
