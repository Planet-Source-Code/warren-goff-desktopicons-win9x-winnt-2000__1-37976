VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "DesktopIcons.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1770
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton RestoreIcons 
      Caption         =   "RestoreIcons"
      Height          =   450
      Left            =   2565
      TabIndex        =   1
      Top             =   555
      Width           =   1545
   End
   Begin VB.CommandButton SaveIcons 
      Caption         =   "SaveIcons"
      Height          =   450
      Left            =   285
      TabIndex        =   0
      Top             =   525
      Width           =   1545
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========Desktop SysListView staff=============

Private Declare Function FindWindow Lib "user32" Alias _
"FindWindowA" (ByVal lpClassName As String, ByVal _
lpWindowName As String) As Long

Private Declare Function FindWindowEx& Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName As String, ByVal lpWindowName As String)
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Const GWL_STYLE = (-16)
Private Const LVS_AUTOARRANGE = &H100

Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, lParam As Any)
Private Const LVM_FIRST = &H1000
Private Const LVM_GETTITEMCOUNT& = (LVM_FIRST + 4)
Private Const LVM_SETITEMPOSITION& = (LVM_FIRST + 15)
Private Const LVM_GETITEMPOSITION& = (LVM_FIRST + 16)
Private Const WM_COMMAND = &H111
Private Const IDM_TOGGLEAUTOARRANGE = &H7041

'=============================================================
Dim ptOriginal() As POINTAPI
Dim ptCurrent() As POINTAPI
Dim xScreen As Long, yScreen As Long
Dim bAutoArrange As Boolean
Private Sub RestoreIcons_Click()
    Dim lWidth, lHeight As Long
    Dim Resolution As String
    Dim h As Long, nCount, z(2000), i As Long
    lWidth = Screen.Width \ Screen.TwipsPerPixelX
    lHeight = Screen.Height \ Screen.TwipsPerPixelY
    Resolution = "\" & Str(lWidth) & "-" & Str(lHeight)
    nCount = 0
    Open App.Path & Resolution For Input As #1
    Do While Not EOF(1)
    Input #1, z(nCount)
    nCount = nCount + 1
    Loop
    Close #1
   h = GetSysLVHwnd
   'nCount = SendMessage(h, LVM_GETTITEMCOUNT, 0, 0&)
   For i = 0 To nCount - 1
       Call SendMessage(h, LVM_SETITEMPOSITION, i, ByVal CLng(z(i)))     'ByVal CLng(ptOriginal(i).X + ptOriginal(i).Y * &H10000))
   Next
   If bAutoArrange Then
      Call SendMessage(GetParent(h), WM_COMMAND, IDM_TOGGLEAUTOARRANGE, ByVal 0&)
   End If


End Sub

Private Sub SaveIcons_Click()
    Dim lWidth, lHeight As Long
    Dim Resolution As String
    Dim pid As Long, tid As Long, lStyle As Long
    Dim hProcess As Long, lpSysShared As Long, dwSize As Long
    Dim nCount As Long, lWritten As Long, hFileMapping As Long
    Dim h As Long, i, z As Long
    lWidth = Screen.Width \ Screen.TwipsPerPixelX
    lHeight = Screen.Height \ Screen.TwipsPerPixelY
    Resolution = "\" & Str(lWidth) & "-" & Str(lHeight)
   h = GetSysLVHwnd
   If h = 0 Then Exit Sub
   If (GetWindowLong(h, GWL_STYLE) And LVS_AUTOARRANGE) = LVS_AUTOARRANGE Then
      bAutoArrange = True
      Call SendMessage(GetParent(h), WM_COMMAND, IDM_TOGGLEAUTOARRANGE, ByVal 0&)
   End If
   tid = GetWindowThreadProcessId(h, pid)
   nCount = SendMessage(h, LVM_GETTITEMCOUNT, 0, 0&)
   If nCount = 0 Then Exit Sub
   xScreen = Screen.Width \ Screen.TwipsPerPixelX
   yScreen = Screen.Height \ Screen.TwipsPerPixelY
   ReDim ptOriginal(nCount - 1)
   ReDim ptCurrent(nCount - 1)
   dwSize = Len(ptOriginal(0))
   If IsWindowsNT Then
      lpSysShared = GetMemSharedNT(pid, dwSize, hProcess)
      WriteProcessMemory hProcess, ByVal lpSysShared, ptOriginal(0), dwSize, lWritten
      For i = 0 To nCount - 1
          SendMessage h, LVM_GETITEMPOSITION, i, ByVal lpSysShared
          ReadProcessMemory hProcess, ByVal lpSysShared, ptOriginal(i), dwSize, lWritten
      Next i
      FreeMemSharedNT hProcess, lpSysShared, dwSize
   Else
      lpSysShared = GetMemShared95(dwSize, hFileMapping)
      CopyMemory ByVal lpSysShared, ptOriginal(0), dwSize
      For i = 0 To nCount - 1
          SendMessage h, LVM_GETITEMPOSITION, i, ByVal lpSysShared
          CopyMemory ptOriginal(i), ByVal lpSysShared, dwSize
          ptCurrent(i).X = xScreen / 2
          ptCurrent(i).Y = yScreen / 2
      Next i
      FreeMemShared95 hFileMapping, lpSysShared
   End If
      Open App.Path & Resolution For Output As #1
       For i = 0 To nCount - 1
         z = ptOriginal(i).X + ptOriginal(i).Y * &H10000
         Print #1, z
      Next i
      Close #1
End Sub
Private Function GetSysLVHwnd() As Long
   Dim h As Long
   h = FindWindow("Progman", vbNullString)
   h = FindWindowEx(h, 0, "SHELLDLL_defVIEW", vbNullString)
   GetSysLVHwnd = FindWindowEx(h, 0, "SysListView32", vbNullString)
End Function

