Attribute VB_Name = "modMain"
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function ShowWindow& Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long)
Public Declare Function SetForegroundWindow& Lib "user32" (ByVal hwnd As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
                                                                               ByVal lpFile As String, ByVal lpParameters As String, _
                                                                               ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2

Public Const WM_MOUSEMOVE = &H200
Public Const MAX_TIP_LENGTH As Long = 64

Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10

Public Const MF_BYCOMMAND = &H0&
Public Const SC_CLOSE = &HF060&
Public Const WS_MINIMIZEBOX = &H20000
Public Const MF_SYSMENU = &H2000&

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * MAX_TIP_LENGTH
End Type

Public Function AddIcon(nidTrayIcon As NOTIFYICONDATA) As Boolean
  
'Show the icon specified in the systray.
AddIcon = Shell_NotifyIcon(NIM_ADD, nidTrayIcon)
  
End Function

Public Function DeleteIcon(nidTrayIcon As NOTIFYICONDATA) As Boolean
  
'Remove the icon from the systray.
DeleteIcon = Shell_NotifyIcon(NIM_DELETE, nidTrayIcon)

End Function
