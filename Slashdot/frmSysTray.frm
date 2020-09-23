VERSION 5.00
Begin VB.Form frmSysTray 
   BorderStyle     =   0  'None
   ClientHeight    =   1365
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4485
   Icon            =   "frmSysTray.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSysTrayTip 
      Height          =   315
      Left            =   150
      TabIndex        =   1
      Text            =   "Slashdot News "
      Top             =   660
      Width           =   4125
   End
   Begin VB.PictureBox picSysTray 
      Height          =   555
      Left            =   90
      Picture         =   "frmSysTray.frx":000C
      ScaleHeight     =   495
      ScaleWidth      =   465
      TabIndex        =   0
      Top             =   30
      Width           =   525
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private IconData As NOTIFYICONDATA

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209

Private Sub Form_Load()

With IconData
   .cbSize = Len(IconData)
   .hIcon = picSysTray.Picture
   .hwnd = hwnd
   .szTip = Left(txtSysTrayTip.Text, MAX_TIP_LENGTH - 1) & vbNullChar
   .uCallbackMessage = WM_MOUSEMOVE
   .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
   .uId = vbNull
End With
  
retval = AddIcon(IconData)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Dim msgCallBackMessage As Long

msgCallBackMessage = x / Screen.TwipsPerPixelX

Select Case msgCallBackMessage
    Case WM_MOUSEMOVE
      Debug.Print "Mouse is moving"
    Case WM_LBUTTONDOWN
      Debug.Print "Left button went down"
    Case WM_LBUTTONUP
      Debug.Print "Left button came up"
    Case WM_LBUTTONDBLCLK
      Debug.Print "Double click catched from left button"
    Case WM_RBUTTONDOWN
      Debug.Print "Right button went down"
    Case WM_RBUTTONUP
      Debug.Print "Right button came up"
      PopupMenu mnuPopUp
    Case WM_RBUTTONDBLCLK
      Debug.Print "Double click catched from right button"
    Case WM_MBUTTONDOWN
      Debug.Print "Middle button went down"
    Case WM_MBUTTONUP
      Debug.Print "Middle button came up"
    Case WM_MBUTTONDBLCLK
      Debug.Print "Double click catched from middle button"
  End Select

End Sub

Private Sub mnuExit_Click()

retval = DeleteIcon(IconData)
End

End Sub


Private Sub mnuHide_Click()

ShowWindow frmMain.hwnd, SW_HIDE

End Sub

Private Sub mnuShow_Click()

ShowWindow frmMain.hwnd, SW_RESTORE
SetForegroundWindow frmMain.hwnd

End Sub


