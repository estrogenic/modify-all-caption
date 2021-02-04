VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2475
   ClientLeft      =   12135
   ClientTop       =   1980
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   3975
   Begin VB.Timer Timer3 
      Interval        =   1500
      Left            =   1200
      Top             =   1920
   End
   Begin VB.ListBox List2 
      Height          =   1680
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   720
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   240
      Top             =   1920
   End
   Begin VB.ListBox List1 
      Height          =   1680
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

Private Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal flgs As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpmii As MENUITEMINFO) As Long

Private Type MENUITEMINFO
cbSize As Long
fMask As Long
fType As Long
fState As Long
wID As Long
hSubMenu As Long
hbmpChecked As Long
hbmpUnchecked As Long
dwItemData As Long
dwTypeData As String
cch As Long
End Type

Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2

Dim now$, now2$

Private Function MakeLayeredWnd(hWnd As Long) As Long
     Dim WndStyle As Long
     WndStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
     WndStyle = WndStyle Or WS_EX_LAYERED
     MakeLayeredWnd = SetWindowLong(hWnd, GWL_EXSTYLE, WndStyle)
End Function

Private Function GetName(lWindow As Long) As String
    Dim wndTitle As String, Temp As String
    Dim tLen&
    wndTitle = Space(256)
    GetClassName lWindow, wndTitle, 256
    If InStr(1, wndTitle, vbNullChar) Then
        Temp = Left$(wndTitle, InStr(1, wndTitle, vbNullChar) - 1)
    End If
    
    tLen = GetWindowTextLength(lWindow)
    GetWindowText lWindow, wndTitle, 256
    GetName = Left$(wndTitle, tLen) & "/" & Temp & "/"
End Function

Function Random_String(ByVal Length As Long) As String
    Randomize
    
    Dim StrA    As String
    Dim i         As Long
    
    Random_String = ""
    StrA = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    
    Do
        If i >= Length Then: Exit Do
        Random_String = Random_String & Mid(StrA, Int((Len(StrA)) * Rnd) + 1, 1)
        i = i + 1
    Loop
End Function

Function CC(TOP_hWnd&, ToChange$)
Dim lWindow As Long
Dim lChild As Long, lChild2 As Long

List1.Clear

lWindow = TOP_hWnd '//FindWindow(vbNullString, Caption)
If lWindow = 0 Then Exit Function
List1.AddItem lWindow

Dim Except&
Except = lWindow
Do
    Call SetWindowText(lWindow, ToChange)
    If lWindow <> Except Then EnableWindow lWindow, 0
    EnableWindow lWindow, 1
    MakeLayeredWnd lWindow
    SetLayeredWindowAttributes lWindow, 0, 255 * (0.85), LWA_ALPHA

    lChild = GetWindow(lWindow, GW_CHILD)
    If lChild Then
        List1.AddItem lChild
        Do
            lChild2 = GetWindow(lChild, GW_HWNDNEXT)
            If lChild2 Then List1.AddItem lChild2
            lChild = lChild2
        Loop Until lChild2 = 0
    End If
    lWindow = List1.List(List1.ListCount - 1)
    List1.RemoveItem (List1.ListCount - 1)
Loop Until List1.ListCount = 0
End Function

Private Function CM(TOP_hWnd&, ToChange$)
Dim lWindow&, hMenu&, hSubMenu&, lngID&
    Dim MII As MENUITEMINFO
    Dim i%, m%
    
    lWindow = TOP_hWnd '//FindWindow(vbNullString, "Memory Viewer")
    hMenu = GetMenu(lWindow)
    
    For i = 0 To GetMenuItemCount(hMenu) - 1
        MII.cbSize = Len(MII)
        MII.fMask = &H10 Or 1
        MII.fState = &H80
        MII.dwTypeData = ToChange '//Random_String(7)
        SetMenuItemInfo hMenu, i, True, MII '//change main menus caption
        
        hSubMenu = GetSubMenu(hMenu, i) '//0
        
        For m = 0 To GetMenuItemCount(hSubMenu) - 1
            lngID = GetMenuItemID(hSubMenu, m) '//-1
            lWindow = GetMenuString(hMenu, lngID, sztext, 40, 0) '//0
            If lWindow > 1 Then Call ModifyMenu(hMenu, lngID, 0, lngID, ToChange) '//change sub menus caption
        Next m
        
        DrawMenuBar hMenu
    Next i
End Function

Private Sub Form_Load()
GetHandleFromPid "exeinfope.exe"
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Timer1_Timer()
Dim i%
now = Random_String(10)

For i = 0 To List2.ListCount - 1
    CC List2.List(i), now
Next i
End Sub

Private Sub Timer2_Timer()
    Dim i%
    now2 = Random_String(5)

    For i = 0 To List2.ListCount - 1
        CM List2.List(i), now2
    Next i
End Sub

Private Sub Timer3_Timer()
Timer1.Enabled = False
Timer2.Enabled = False

List2.Clear

GetHandleFromPid "exeinfope.exe"

Do While List2.ListCount = 0
DoEvents
Loop

Timer1.Enabled = True
Timer2.Enabled = True
End Sub
