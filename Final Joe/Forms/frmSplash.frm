VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   195
      Top             =   120
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00808080&
      Height          =   3735
      Left            =   15
      Top             =   15
      Width           =   5835
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   -3060
      Picture         =   "frmSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   15360
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "ystem"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   4710
      TabIndex        =   8
      Top             =   1980
      Width           =   840
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   540
      Left            =   4365
      TabIndex        =   7
      Top             =   1740
      Width           =   465
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ayroll"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   450
      Left            =   3150
      TabIndex        =   5
      Top             =   1830
      Width           =   1155
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   885
      Left            =   2655
      TabIndex        =   6
      Top             =   1515
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Joehel V. del Rosario    joehelcute@yahoo.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C25418&
      Height          =   165
      Left            =   2805
      TabIndex        =   4
      Top             =   3495
      Width           =   2925
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developer:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C25418&
      Height          =   165
      Left            =   2040
      TabIndex        =   3
      Top             =   3495
      Width           =   690
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   5325
      Picture         =   "frmSplash.frx":00B6
      Top             =   1905
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COPYRIGHT © Jhels 2007"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C25418&
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   3255
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registered to:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C25418&
      Height          =   195
      Left            =   3900
      TabIndex        =   1
      Top             =   105
      Width           =   1035
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1305
      Left            =   0
      Top             =   3120
      Width           =   7695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "822 MOVERS (PAILO)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C25418&
      Height          =   615
      Left            =   3900
      TabIndex        =   0
      Top             =   360
      Width           =   1830
   End
   Begin VB.Image Image5 
      Height          =   960
      Left            =   -2760
      Picture         =   "frmSplash.frx":01D8
      Stretch         =   -1  'True
      Top             =   -60
      Width           =   15360
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   75
      Picture         =   "frmSplash.frx":035A
      Stretch         =   -1  'True
      Top             =   795
      Width           =   2625
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><>'
'Programmed By: Joehel V. del Rosario    '
'822 MOVERS(PAILO) Trucking Corporation  '
'Date: June 1, 2007                      '
'<><><><><><><><><><><><><><><><><><><><>'

Option Explicit

'for form fading function
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const g = (-20)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000

Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Private Const HWND_NOTOPMOST = -2


Dim tl As Integer

Public Function ShowSplash()
    
    'show form
    SetWindowPos Me.hwnd, HWND_TOPMOST, _
    0, 0, 0, 0, Flags
    Me.Show
    
    DoEvents
    DoEvents
    DoEvents
    
    'continue loading...
    'Call modMain.Main_AfterSD
    MDIMainForm.Show
End Function


Public Function ShowForm()
    
    'show form
    Me.Show
End Function

Public Sub UnloadSplash()
    Me.Enabled = False
    Timer1.Enabled = True
End Sub


Private Sub Form_Deactivate()
    UnloadSplash
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        UnloadSplash
    End If
End Sub
Private Sub Form_Activate()
    SetWindowPos Me.hwnd, HWND_TOPMOST, _
    0, 0, 0, 0, Flags
End Sub



Private Sub Trans(Level As Integer)
        Dim Msg As Long

        On Error Resume Next
        
        Msg = GetWindowLong(Me.hwnd, g)
        Msg = Msg Or WS_EX_LAYERED
        SetWindowLong Me.hwnd, g, Msg
        SetLayeredWindowAttributes Me.hwnd, 0, Level, LWA_ALPHA
        
End Sub

Private Sub Form_Load()
    tl = 256
End Sub
Private Sub Timer1_Timer()
    
    Trans tl
    On Error Resume Next
    tl = tl - 50
    'Me.Width = Me.Width - 1475

    If tl < 25 Then
        Timer1.Enabled = False
        tl = 50
        
        Unload Me
        MDIMainForm.Show
    End If
End Sub


