Attribute VB_Name = "mdlGeneral"
Option Explicit

Public Function GetRelColumnHit(InThisListview As ListView, X As Single) As Integer
    Dim RunningSpan As Single
    Dim Count As Integer
    
    GetRelColumnHit = 0
    
    For Count = 1 To InThisListview.ColumnHeaders.Count
        RunningSpan = RunningSpan + InThisListview.ColumnHeaders(Count).Width
        If (X < RunningSpan) Then
            GetRelColumnHit = Count
            Exit Function
        End If
    Next

End Function

'This function uses the absolute screen location of the pointer to figure out column hits
Public Function GetAbsColumnHit(InThisForm As Form, ThisListview As ListView) As Integer
    Dim RunningSpan As Single
    Dim Count As Integer
    Dim RelativeX As Single
    Dim RelativeY As Single
    Dim FormThinBorder As Single
    Dim FormTitleBar As Single
    
    GetAbsColumnHit = 0
    
    'Mousecondition returns coordinates in pixels (you need to do a conversion to/from twips for the listview and the form)
    FormThinBorder = (InThisForm.Width - InThisForm.ScaleWidth) / 2 'in twips
    FormTitleBar = (InThisForm.Height - InThisForm.ScaleHeight - FormThinBorder) 'in twips
    RelativeX = MouseCondition.X * Screen.TwipsPerPixelX - (InThisForm.Left + FormThinBorder + ThisListview.Left)
    RelativeY = MouseCondition.Y * Screen.TwipsPerPixelY - (InThisForm.Top + FormTitleBar + FormThinBorder + ThisListview.Top)
    
    'Check to see if it is within the height of the listview
    If (RelativeY < 0 Or RelativeY > ThisListview.Height) Then Exit Function
    
    'Find which column contains the cursor
    For Count = 1 To ThisListview.ColumnHeaders.Count
        RunningSpan = RunningSpan + ThisListview.ColumnHeaders(Count).Width
        If (RelativeX < RunningSpan) Then
            GetAbsColumnHit = Count
            Exit For
        End If
    Next
    
End Function
