You will note that this works for both Win 9X and WinNT/2000!!!   I take no credit for this code as I found it and altered it only slightly to save the icon positions for a given resolution.   Frankly, I am a novice at API and much of this is over my head.   Here is where I found the code and the author and I've reproduced his post:

http://www.codeguru.com/vb/comments/991.shtml

Ben: bczulowski@hotmail.com

Attribute VB_Name = "mSharedMemory"
Option Explicit
'Some API (SendMessage for example) use pointers to structures to be filled
'with some data. If you're sending such message to window belong to your
'process - no problem. But if you try to send this message to different
'process GPF can occure, because structure address belong to calling process
'memory space and target process can not achive this address. Here is
'work around.
'For Win95/98/ME we can use File Mapping, because OS place mapped files
'into shareable memory space. But we can't use this trick for NT - NT map
'files into calling process memory area. In this case, we can use
'VirtualAllocEx function to reserve memory inside target process.

'=========Checking OS staff=============
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long

'========= Win95/98/ME Shared memory staff===============
Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, ByVal lpFileMappigAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const SECTION_QUERY = &H1
Const SECTION_MAP_WRITE = &H2
Const SECTION_MAP_READ = &H4
Const SECTION_MAP_EXECUTE = &H8
Const SECTION_EXTEND_SIZE = &H10
Const SECTION_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE
Const FILE_MAP_ALL_ACCESS = SECTION_ALL_ACCESS

'============NT Shared memory staff======================
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Const PROCESS_VM_OPERATION = &H8
Const PROCESS_VM_READ = &H10
Const PROCESS_VM_WRITE = &H20
Const PROCESS_ALL_ACCESS = 0
Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Const MEM_COMMIT = &H1000
Const MEM_RESERVE = &H2000
Const MEM_DECOMMIT = &H4000
Const MEM_RELEASE = &H8000
Const MEM_FREE = &H10000
Const MEM_PRIVATE = &H20000
Const MEM_MAPPED = &H40000
Const MEM_TOP_DOWN = &H100000

'==========Memory access constants===========
Private Const PAGE_NOACCESS = &H1&
Private Const PAGE_READONLY = &H2&
Private Const PAGE_READWRITE = &H4&
Private Const PAGE_WRITECOPY = &H8&
Private Const PAGE_EXECUTE = &H10&
Private Const PAGE_EXECUTE_READ = &H20&
Private Const PAGE_EXECUTE_READWRITE = &H40&
Private Const PAGE_EXECUTE_WRITECOPY = &H80&
Private Const PAGE_GUARD = &H100&
Private Const PAGE_NOCACHE = &H200&

Public Function GetMemShared95(ByVal memSize As Long, hFile As Long) As Long
    hFile = CreateFileMapping(&HFFFFFFFF, 0, PAGE_READWRITE, 0, memSize, vbNullString)
    GetMemShared95 = MapViewOfFile(hFile, FILE_MAP_ALL_ACCESS, 0, 0, 0)
End Function

Public Sub FreeMemShared95(ByVal hFile As Long, ByVal lpMem As Long)
    UnmapViewOfFile lpMem
    CloseHandle hFile
End Sub

Public Function GetMemSharedNT(ByVal pid As Long, ByVal memSize As Long, hProcess As Long) As Long
    hProcess = OpenProcess(PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE, False, pid)
    GetMemSharedNT = VirtualAllocEx(ByVal hProcess, ByVal 0&, ByVal memSize, MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
End Function

Public Sub FreeMemSharedNT(ByVal hProcess As Long, ByVal MemAddress As Long, ByVal memSize As Long)
   Call VirtualFreeEx(hProcess, ByVal MemAddress, memSize, MEM_RELEASE)
   CloseHandle hProcess
End Sub

Public Function IsWindowsNT() As Boolean
   Dim verinfo As OSVERSIONINFO
   verinfo.dwOSVersionInfoSize = Len(verinfo)
   If (GetVersionEx(verinfo)) = 0 Then Exit Function
   If verinfo.dwPlatformId = 2 Then IsWindowsNT = True
End Function

''''''''''''''''''''''''''

Attribute VB_Name = "Module1"
Option Explicit

Public Enum SHUFFLE_TYPE
    RANDOM
    SINE
    CIRCLES
'More depend on your fantasy and geomethry knowledge :)
End Enum
'=========Desktop SysListView staff=============
Type POINTAPI
     x As Long
     y As Long
End Type

Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)
Private Declare Function FindWindowEx& Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName As String, ByVal lpWindowName As String)
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Const GWL_STYLE = (-16)
Private Const LVS_AUTOARRANGE = &H100

Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any)
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

Public Sub ShuffleDesktopIcons(ByVal ShuffleType As SHUFFLE_TYPE)
   Dim h As Long, nCount As Long, i As Long
   Dim FactorX As Single, FactorY As Single, Radius As Single
   Dim x As Long, y As Long, cx As Long, cy As Long
   FactorX = Rnd
   FactorY = Rnd
   Radius = xScreen * Rnd / 2
   h = GetSysLVHwnd
   nCount = SendMessage(h, LVM_GETTITEMCOUNT, 0, 0&)
   For i = 0 To nCount - 1
       Select Case ShuffleType
              Case RANDOM
                   x = Int(Rnd * xScreen)
                   y = Int(Rnd * yScreen)
              Case SINE
                   x = FactorX * xScreen * i \ nCount
                   y = FactorY * yScreen * (1 - Sin(i * 6.28 / nCount)) \ 2
              Case CIRCLES
                   x = xScreen / 2 - Radius * Cos(i * 6.28 / nCount)
                   y = yScreen / 2 - Radius * Sin(i * 6.28 / nCount)
       End Select
       Call SendMessage(h, LVM_SETITEMPOSITION, i, ByVal CLng(x + y * &H10000))
   Next
End Sub

Public Sub RestoreDesktopIcons()
   Dim h As Long, nCount As Long, i As Long
   h = GetSysLVHwnd
   nCount = SendMessage(h, LVM_GETTITEMCOUNT, 0, 0&)
   For i = 0 To nCount - 1
       Call SendMessage(h, LVM_SETITEMPOSITION, i, ByVal CLng(ptOriginal(i).x + ptOriginal(i).y * &H10000))
   Next
   If bAutoArrange Then
      Call SendMessage(GetParent(h), WM_COMMAND, IDM_TOGGLEAUTOARRANGE, ByVal 0&)
   End If
End Sub

Public Function StoreDeskTopInfo() As Boolean
   Dim pid As Long, tid As Long, lStyle As Long
   Dim hProcess As Long, lpSysShared As Long, dwSize As Long
   Dim nCount As Long, lWritten As Long, hFileMapping As Long
   Dim h As Long, i As Long
   h = GetSysLVHwnd
   If h = 0 Then Exit Function
   If (GetWindowLong(h, GWL_STYLE) And LVS_AUTOARRANGE) = LVS_AUTOARRANGE Then
      bAutoArrange = True
      Call SendMessage(GetParent(h), WM_COMMAND, IDM_TOGGLEAUTOARRANGE, ByVal 0&)
   End If
   tid = GetWindowThreadProcessId(h, pid)
   nCount = SendMessage(h, LVM_GETTITEMCOUNT, 0, 0&)
   If nCount = 0 Then Exit Function
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
          ptCurrent(i).x = xScreen / 2
          ptCurrent(i).y = yScreen / 2
      Next i
      FreeMemShared95 hFileMapping, lpSysShared
   End If
   StoreDeskTopInfo = True
End Function

Private Function GetSysLVHwnd() As Long
   Dim h As Long
   h = FindWindow("Progman", vbNullString)
   h = FindWindowEx(h, 0, "SHELLDLL_defVIEW", vbNullString)
   GetSysLVHwnd = FindWindowEx(h, 0, "SysListView32", vbNullString)
End Function

'''''''''''''''''''''''''''

VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   1875
   End
   Begin VB.Timer Timer1 
      Left            =   4260
      Top             =   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2580
      TabIndex        =   1
      Top             =   780
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   780
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1875
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bRunning As Boolean

Private Sub Command1_Click()
   If Not bRunning Then
      If StoreDeskTopInfo Then Timer1.Enabled = True
      Caption = "See your desktop dancing!"
      bRunning = True
      Command1.Enabled = False
      Command2.Enabled = True
   End If
End Sub

Private Sub Command2_Click()
   If bRunning Then
      Timer1.Enabled = False
      RestoreDesktopIcons
      bRunning = False
      Caption = "Desktop shuffle sample"
      Command1.Enabled = True
      Command2.Enabled = False
   End If
End Sub

Private Sub Form_Load()
   Timer1.Interval = 200
   Timer1.Enabled = False
   Caption = "Desktop shuffle demo"
   Label1 = "Desktop shuffling type"
   With Combo1
        .AddItem "RANDOM"
        .AddItem "SINE"
        .AddItem "CIRCLES"
        .ListIndex = 0
   End With
   Command1.Caption = "&Start"
   Command2.Caption = "&Stop"
   Command1.Enabled = True
   Command2.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If bRunning Then RestoreDesktopIcons
End Sub

Private Sub Timer1_Timer()
    ShuffleDesktopIcons Combo1.ListIndex
End Sub

''''''''''''''''''''''''''''

enjoy !!!
