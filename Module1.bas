Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szEXEFile As String * 260
End Type

Dim Pid As Long
Dim WindowHandle As Long

Private Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim Temp&, tLng&
    Dim WndName$
    GetWindowThreadProcessId hWnd, Temp
    
    EnumWindowsProc = 1
    If Temp = Pid Then
        tLng = GetWindowTextLength(hWnd)
        WndName = Space(tLng + 1)
        tLng = GetWindowText(hWnd, WndName, tLng + 1)
        WndName = Left$(WndName, tLng)
        If InStr(WndName, "MSCTFIME UI") Then Exit Function
        If InStr(WndName, "Default IME") Then Exit Function
        If WndName = vbNullString Then Exit Function
        Form1.List2.AddItem hWnd
    End If
End Function

Public Sub GetHandleFromPid(ProcessName As String)
    Dim uProcess As PROCESSENTRY32, hSnapShot As Long, Result As Long
    
    uProcess.dwSize = Len(uProcess)
    hSnapShot = CreateToolhelpSnapshot(2&, 0&)
    Result = ProcessFirst(hSnapShot, uProcess)
    
    Do
        If Split(uProcess.szEXEFile, vbNullChar)(0) = ProcessName Then
            Pid = uProcess.th32ProcessID
            Call EnumWindows(AddressOf EnumWindowsProc, 0)
        End If
        
        Result = ProcessNext(hSnapShot, uProcess)
    Loop While Result
End Sub
