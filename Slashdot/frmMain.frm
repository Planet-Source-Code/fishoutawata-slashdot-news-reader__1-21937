VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Slashdot "
   ClientHeight    =   2970
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerInterval 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   990
      Top             =   2130
   End
   Begin VB.PictureBox picArticle 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   0
      ScaleHeight     =   2985
      ScaleWidth      =   5745
      TabIndex        =   0
      Top             =   0
      Width           =   5745
      Begin VB.Label lblNavNext 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5310
         MouseIcon       =   "frmMain.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   15
         ToolTipText     =   "Move one Article Forward"
         Top             =   2160
         Width           =   345
      End
      Begin VB.Label lblNavBack 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4890
         MouseIcon       =   "frmMain.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   14
         ToolTipText     =   "Move one Article Back"
         Top             =   2160
         Width           =   345
      End
      Begin VB.Line Line2 
         X1              =   60
         X2              =   60
         Y1              =   2580
         Y2              =   2910
      End
      Begin VB.Line Line1 
         X1              =   60
         X2              =   5550
         Y1              =   2565
         Y2              =   2565
      End
      Begin VB.Label lblCountdown 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4530
         TabIndex        =   13
         ToolTipText     =   $"frmMain.frx":0614
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblCountdownStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Next Update:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3510
         TabIndex        =   12
         Top             =   2640
         Width           =   1005
      End
      Begin VB.Label lblUrlText 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1170
         MouseIcon       =   "frmMain.frx":06C6
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   1470
         Width           =   4485
      End
      Begin VB.Label lblTimeText 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1170
         TabIndex        =   10
         Top             =   1110
         Width           =   4485
      End
      Begin VB.Label lblAuthorText 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1170
         TabIndex        =   9
         Top             =   690
         Width           =   4485
      End
      Begin VB.Label lblTitleText 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1170
         TabIndex        =   8
         Top             =   300
         Width           =   4485
      End
      Begin VB.Label lblTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   600
         TabIndex        =   7
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label lblOf 
         BackStyle       =   0  'Transparent
         Caption         =   "of"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   330
         TabIndex        =   6
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label lblCurrent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   5
         Top             =   2640
         Width           =   225
      End
      Begin VB.Label lblUrl 
         BackStyle       =   0  'Transparent
         Caption         =   "Url:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   4
         Top             =   1530
         Width           =   1185
      End
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Submitted:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   3
         Top             =   1125
         Width           =   1185
      End
      Begin VB.Label lblAuthor 
         BackStyle       =   0  'Transparent
         Caption         =   "Author:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   2
         Top             =   705
         Width           =   1185
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   1
         Top             =   300
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objSlashDot     As cSlashdot
Private iCountDown      As Integer
Private iCurrentArticle As Integer


Private Sub Form_Load()

Dim hMenu  As Long
Dim retval As Long

Load frmSysTray

hMenu = GetSystemMenu(hwnd, False)
retval = DeleteMenu(hMenu, SC_CLOSE, MF_BYCOMMAND)
InitForm

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

MoveForm

End Sub

Private Sub lblClose_Click()

ShowWindow Me.hwnd, SW_HIDE

End Sub

Private Sub lblNavBack_Click()

If iCurrentArticle > 1 Then
    UpdateForm (iCurrentArticle - 1)
End If

End Sub

Private Sub lblNavNext_Click()

If iCurrentArticle < 10 Then
    UpdateForm (iCurrentArticle + 1)
End If

End Sub


Private Sub lblUrlText_Click()

Dim lRetVal As Long

lRetVal = ShellExecute(Me.hwnd, "", lblUrlText.Caption, "", "", SW_SHOWDEFAULT)
    
End Sub

Private Sub pbTitleBar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

MoveForm

End Sub

Public Sub MoveForm()

Dim Ret&
ReleaseCapture
Ret& = SendMessage(Me.hwnd, &H112, &HF012, 0)

End Sub

Private Sub picArticle_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

'Does not work in Windows 2000
'MoveForm

PopupMenu frmSysTray.mnuPopUp

End Sub

Private Sub TimerInterval_Timer()


If iCountDown > 0 Then
    iCountDown = iCountDown - 1
    lblCountdown.Caption = iCountDown & " minutes"
Else
    InitForm
End If

End Sub

Private Sub TimerRetrieve_Timer()

objSlashDot.RetieveArticles

TimerInterval.Enabled = True

End Sub

Public Sub InitForm()

Dim iIndex As Integer

If Not IsObject(objSlashDot) Then
    Set objSlashDot = New cSlashdot
Else
    Set objSlashDot = Nothing
    Set objSlashDot = New cSlashdot
End If

iIndex = 1

iCountDown = 30

objSlashDot.RetieveArticles

lblTitleText.Caption = objSlashDot.GetArticle(iIndex).Title
lblAuthorText.Caption = objSlashDot.GetArticle(iIndex).Author
lblTimeText.Caption = objSlashDot.GetArticle(iIndex).Submitted
lblUrlText.Caption = objSlashDot.GetArticle(iIndex).Url

lblCountdown.Caption = "30 minutes"

TimerInterval.Enabled = True

lblCurrent.Caption = CStr(iIndex)
lblTotal.Caption = CStr(objSlashDot.ArticleCount)

iCurrentArticle = iIndex

End Sub

Public Sub UpdateForm(iIndex As Integer)

lblTitleText.Caption = objSlashDot.GetArticle(iIndex).Title
lblAuthorText.Caption = objSlashDot.GetArticle(iIndex).Author
lblTimeText.Caption = objSlashDot.GetArticle(iIndex).Submitted
lblUrlText.Caption = objSlashDot.GetArticle(iIndex).Url

lblCurrent.Caption = CStr(iIndex)
lblTotal.Caption = CStr(objSlashDot.ArticleCount)

iCurrentArticle = iIndex

End Sub
