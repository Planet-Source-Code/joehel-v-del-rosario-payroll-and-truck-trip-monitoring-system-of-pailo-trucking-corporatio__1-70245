VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "User's Login"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   389
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   75
      Top             =   2385
   End
   Begin MOVERS.ACPRibbon ACPRibbon1 
      Height          =   1740
      Left            =   1410
      TabIndex        =   12
      Top             =   3135
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   3069
      BackColor       =   4210752
      ForeColor       =   -2147483630
   End
   Begin VB.ComboBox TxtUserID 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2385
      TabIndex        =   0
      Top             =   1005
      Width           =   3285
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2370
      MaxLength       =   20
      PasswordChar    =   "="
      TabIndex        =   1
      Top             =   1950
      Width           =   3285
   End
   Begin MOVERS.CandyButton ButSave 
      Height          =   435
      Left            =   3135
      TabIndex        =   9
      Top             =   2370
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "         &Log-in"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "frmLogin.frx":058A
      PictureAlignment=   2
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin MOVERS.CandyButton CandyButton1 
      Height          =   435
      Left            =   4440
      TabIndex        =   10
      Top             =   2370
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   767
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "         &Cancel"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "frmLogin.frx":0D04
      PictureAlignment=   2
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin MSComctlLib.ImageList imglUser 
      Left            =   2205
      Top             =   3195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":147E
            Key             =   "admin"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLogin.frx":1A18
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin MOVERS.JOEGradLine JOEGradLine1 
      Height          =   90
      Left            =   -150
      TabIndex        =   11
      Top             =   390
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   159
      Color1          =   9594695
      Color2          =   9594695
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[2]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[1]"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password"
      ForeColor       =   &H00004040&
      Height          =   195
      Left            =   2310
      TabIndex        =   7
      Top             =   1485
      Width           =   1125
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00926747&
      BorderWidth     =   2
      Height          =   2985
      Left            =   15
      Top             =   15
      Width           =   5820
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User's Log-in"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00926747&
      Height          =   345
      Left            =   60
      TabIndex        =   8
      Top             =   15
      Width           =   1875
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   195
      Left            =   2355
      TabIndex        =   3
      Top             =   1710
      Width           =   855
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Your Account"
      ForeColor       =   &H00004040&
      Height          =   195
      Left            =   2310
      TabIndex        =   4
      Top             =   510
      Width           =   1440
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&User ID:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   195
      Left            =   2370
      TabIndex        =   2
      Top             =   735
      Width           =   675
   End
End
Attribute VB_Name = "frmLogin"
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


Private Sub ButSave_Click()
   On Error Resume Next
    OpenPBDataBase ("PayrollPersonels")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM PayrollPersonels WHERE Names LIKE '" & Trim(TxtUserID.Text) & "' ")
    With PRFile
        If Not .EOF Then
            If ![o1] = Trim(txtPassword) Then
                MDIMainForm.lblCurrentUser.Caption = Trim(TxtUserID.Text) & " - " & ![o3]
                'Unload Me
                Timer1.Enabled = True
            Else
                MsgBox "Invalid Password!", vbInformation, "invalid Password"
                txtPassword.SetFocus
                SendKeys "{HOME}+{END}"
            End If
        Else
            MsgBox "Username not found!", vbInformation, "Invalid User"
            TxtUserID.SetFocus
            SendKeys "{HOME}+{END}"
        End If
    End With
    
End Sub

Private Sub CandyButton1_Click()
   
    
    If MsgBox("Are you sure you want to exit this system?", vbYesNo + vbInformation, "Exit Application") = vbYes Then
        End
    End If
    
End Sub

Private Sub Form_Activate()
    SetWindowPos Me.hwnd, HWND_TOPMOST, _
    0, 0, 0, 0, Flags
    
    
End Sub

Private Sub Form_Load()
    Dim FRm As Form
    tl = 256

    Dim Theme As Integer
    'default
    TxtUserID.Text = Trim(GetSetting(App.EXEName, "TextBox", TxtUserID.Name, ""))
    
    Call LoadUsers
     
     
    'PaintGrad Me, &HE0E0E0, &HFFFFFF, 135
    
    
    If Val(GetSetting(App.EXEName, "APPThemes", "JThemes", "")) = Null Then
        Theme = Val(GetSetting(App.EXEName, "APPThemes", "JThemes", ""))
    Else
        Theme = 0
    End If

    '# SET Theme
    ACPRibbon1.Theme = Theme    ' 0 - Black
                                ' 1 - Blue
                                ' 2 - Silver
                        
    ACPRibbon1.Refresh
    frmLogin.Picture = ACPRibbon1.LoadBackground
    frmLogin.BackColor = ACPRibbon1.BackColor
    



    MDIMainForm.JoeSBCenter1.STcolor True
    
End Sub
Sub LoadUsers()
    OpenPBDataBase ("PayrollPersonels")
    With PRFile
     .MoveFirst
        Do While Not .EOF
            TxtUserID.AddItem ![Names]
            .MoveNext
        Loop
     .Close
    End With
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    modFunction.FormDrag Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.EXEName, "TextBox", TxtUserID.Name, TxtUserID.Text
    'MDIMainForm.AddChild frmManageEmployee, False
End Sub

Private Sub Label14_Click()
    txtPassword.Text = "movers"
    Call ButSave_Click
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

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call ButSave_Click
    End If
End Sub
Private Sub Trans(Level As Integer)
        Dim Msg As Long

        On Error Resume Next
        
        Msg = GetWindowLong(Me.hwnd, g)
        Msg = Msg Or WS_EX_LAYERED
        SetWindowLong Me.hwnd, g, Msg
        SetLayeredWindowAttributes Me.hwnd, 0, Level, LWA_ALPHA
        
End Sub

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

