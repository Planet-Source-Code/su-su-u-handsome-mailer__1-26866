Attribute VB_Name = "ini"
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long
Declare Function WSACleanup Lib "ws2_32.dll" () As Long

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE
Global r%
Global entry$
Global iniPath$
Global Pull$




Function GetFromINI(AppName$, KeyName$, filename$) As String
Dim RetStr As String
    RetStr = String(255, Chr(0))
        GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), filename$))
End Function
Public Sub Ontop(FormName As Form)
    Call SetWindowPos(FormName.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, flags)
End Sub
Public Sub Notontop(FormName As Form)
    Call SetWindowPos(FormName.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, flags)
End Sub
Public Sub ReadList(list As ListBox, filename As String, Optional ClearList As Boolean)
    On Error GoTo Err
    Open filename For Input As #1
    Do While Not EOF(1)
        Input #1, lstinpuT
        list.AddItem lstinpuT
    Loop
    Close #1
    Exit Sub
Err:
    Exit Sub
End Sub
Public Sub WriteList(list As ListBox, filename As String)

    If list.ListCount <= 0 Then
        Exit Sub
        End
    End If
    On Error GoTo Err
    Open filename For Output As #1
    For i = 0 To list.ListCount - 1
        Print #1, list.list(i)
    Next
    Close #1
    Exit Sub
Err:
    MsgBox "Error In WriteList" & Chr(13) & Chr(13) & Err.Number _
    & " - " & Err.Description, vbCritical, "Error"
    Exit Sub
End Sub


