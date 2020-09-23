Attribute VB_Name = "ModListBox"
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - October 2002
'If you got it elsewhere - they stole it from PSC.

'Written by MrBobo - enjoy
'Please visit our website - www.psst.com.au

'Easily modified for use with ComboBoxes

Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Private Const LB_FINDSTRINGEXACT = &H1A2
Public Function IsInList(nItem As String, lBox As ListBox) As Boolean
    'Rather than removing duplicates it can be
    'far more efficient to ensure an item does
    'not exist in the list before adding it!
    'Not used in this demo - but this is what
    'I use more often than RemoveDupesFromList
    Dim z As Long
    z = SendMessageByString(lBox.hwnd, LB_FINDSTRINGEXACT, 0, nItem)
    IsInList = CBool(z + 1)
End Function
Public Sub RemoveDupesFromList(lBox As ListBox)
    'API makes the task so easy for us.
    'It is also consistantly faster than VB
    Dim z As Long, q As Long
    For z = 0 To lBox.ListCount - 1
        Do Until q = -1
            'Returned Index = APICommand(ListBox window handle, Message, Start Index, Search String)
            q = SendMessageByString(lBox.hwnd, LB_FINDSTRINGEXACT, z + 1, lBox.List(z))
            If q <> -1 Then
                If q = z Then Exit Do 'Dont remove original
                lBox.RemoveItem q
            End If
        Loop
        q = 0 'reset found flag
    Next
End Sub
Public Sub RemoveDupes(lst As ListBox)
    'Very neat VB routine
    'Submitted on: 9/26/2000 9:09:37 AM
    'By: Fredrik Schultz
    Dim iPos As Integer, Tcnt As Long
    iPos = 0
    If lst.ListCount < 1 Then Exit Sub
    Do While iPos < lst.ListCount
        lst.Text = lst.List(iPos)
        If lst.ListIndex <> iPos Then
            lst.RemoveItem iPos
        Else
            iPos = iPos + 1
        End If
    Loop
    lst.Text = "~~~^^~~~"
End Sub



