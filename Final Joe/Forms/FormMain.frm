VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FormMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11115
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15240
   ControlBox      =   0   'False
   Icon            =   "FormMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3720
      Top             =   1530
   End
   Begin MOVERS.JOELine JOELine1 
      Height          =   30
      Left            =   30
      TabIndex        =   12
      Top             =   1125
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   53
      BorderColor1    =   11325655
      BorderColor2    =   16185592
   End
   Begin MOVERS.JOEClientWin JOEClientWin1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   3
      Top             =   10650
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   820
   End
   Begin MOVERS.JoeSBCenter JoeSBCenter1 
      Align           =   3  'Align Left
      Height          =   9480
      Left            =   0
      TabIndex        =   2
      Top             =   1170
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   16722
      Begin MOVERS.JOESideTab JST 
         Height          =   3015
         Index           =   0
         Left            =   60
         TabIndex        =   21
         Top             =   645
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   5318
         Caption         =   "Quick Launch                [Ctrl + Q]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         BorderColor     =   12957347
         AutoContract    =   0   'False
         Begin MSComctlLib.ImageList ilQL 
            Left            =   1020
            Top             =   630
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   5
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormMain.frx":6852
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormMain.frx":6965
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormMain.frx":70DF
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormMain.frx":7859
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormMain.frx":7FD3
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView listQL 
            Height          =   2445
            Left            =   60
            TabIndex        =   23
            Top             =   420
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   4313
            View            =   2
            Arrange         =   1
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            OLEDropMode     =   1
            _Version        =   393217
            Icons           =   "ilQL"
            SmallIcons      =   "ilQL"
            ForeColor       =   -2147483640
            BackColor       =   16777215
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OLEDropMode     =   1
            NumItems        =   0
         End
      End
      Begin MOVERS.JOESideTab JST 
         Height          =   1950
         Index           =   1
         Left            =   60
         TabIndex        =   22
         Top             =   3675
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   3440
         Caption         =   "Trip Status                    [Ctrl + R]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         BorderColor     =   12957347
         Begin VB.PictureBox Picture1 
            Height          =   1515
            Left            =   45
            ScaleHeight     =   1455
            ScaleWidth      =   3120
            TabIndex        =   25
            Top             =   375
            Width           =   3180
         End
      End
      Begin MOVERS.JOESideTab JST 
         Height          =   3660
         Index           =   2
         Left            =   75
         TabIndex        =   24
         Top             =   5640
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   6456
         Caption         =   "Employee Pictures      [Ctrl + R]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   8421504
         BorderColor     =   12957347
         Begin VB.PictureBox Picture2 
            Height          =   3240
            Left            =   45
            ScaleHeight     =   3180
            ScaleWidth      =   3135
            TabIndex        =   26
            Top             =   375
            Width           =   3195
         End
      End
      Begin VB.Label lblCurrentUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Joehel Cute"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   900
         TabIndex        =   14
         Top             =   90
         Width           =   855
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   810
         TabIndex        =   13
         Top             =   300
         Width           =   180
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome,"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   105
         TabIndex        =   11
         Top             =   60
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Today is "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   105
         TabIndex        =   10
         Top             =   270
         Width           =   645
      End
   End
   Begin MOVERS.JOELine JOELine2 
      Height          =   30
      Left            =   0
      TabIndex        =   17
      Top             =   330
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   53
      BorderColor1    =   11325655
      BorderColor2    =   16185592
   End
   Begin VB.PictureBox bgHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00F5F5F5&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   0
      ScaleHeight     =   78
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1016
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      Begin MOVERS.JOESBtop JOESBtop1 
         Height          =   945
         Left            =   0
         TabIndex        =   4
         Top             =   330
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1667
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "BETA"
            BeginProperty Font 
               Name            =   "Viner Hand ITC"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   375
            Left            =   2550
            TabIndex        =   9
            Top             =   345
            Width           =   750
         End
         Begin VB.Label Label4 
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
            ForeColor       =   &H00C0C0C0&
            Height          =   375
            Left            =   1440
            TabIndex        =   8
            Top             =   255
            Width           =   840
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "P"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   600
            Left            =   105
            TabIndex        =   7
            Top             =   -30
            Width           =   525
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   540
            Left            =   1095
            TabIndex        =   6
            Top             =   -15
            Width           =   465
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "ayroll"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   375
            Left            =   375
            TabIndex        =   5
            Top             =   255
            Width           =   675
         End
      End
      Begin VB.PictureBox bgHeaderMenu 
         BackColor       =   &H00EDEBE9&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   0
         ScaleHeight     =   22
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1024
         TabIndex        =   1
         Top             =   0
         Width           =   15360
         Begin MOVERS.JOeMenu JOeMenu1 
            Height          =   315
            Left            =   30
            TabIndex        =   19
            Top             =   0
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   8421504
            Caption         =   "&Transactions"
            ForeColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Transactions"
         End
         Begin MOVERS.JOeMenu JOeMenu2 
            Height          =   315
            Left            =   1335
            TabIndex        =   20
            Top             =   0
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   8421504
            Caption         =   "&Payrolls"
            ForeColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Payrolls"
         End
         Begin MOVERS.JOeMenu JOeMenu3 
            Height          =   315
            Left            =   2640
            TabIndex        =   27
            Top             =   0
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   8421504
            Caption         =   "&Reports"
            ForeColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Reports"
         End
         Begin MOVERS.JOeMenu JOeMenu4 
            Height          =   315
            Left            =   3945
            TabIndex        =   28
            Top             =   0
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   8421504
            Caption         =   "&Settings"
            ForeColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Settings"
         End
         Begin MOVERS.JOeMenu JOeMenu5 
            Height          =   315
            Left            =   5265
            TabIndex        =   29
            Top             =   0
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            BackColor       =   15592425
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   8421504
            Caption         =   "&Help"
            ForeColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "&Help"
         End
         Begin VB.Image imgClose 
            Height          =   360
            Index           =   0
            Left            =   14955
            Picture         =   "FormMain.frx":874D
            Top             =   -15
            Width           =   360
         End
         Begin VB.Image imgClose 
            Height          =   360
            Index           =   1
            Left            =   14940
            Picture         =   "FormMain.frx":8E37
            Top             =   -15
            Visible         =   0   'False
            Width           =   360
         End
      End
      Begin VB.PictureBox bgRecOpt 
         Appearance      =   0  'Flat
         BackColor       =   &H00F5F5F5&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   3450
         ScaleHeight     =   52
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   898
         TabIndex        =   15
         Top             =   345
         Width           =   13470
         Begin MOVERS.JOEToolButton JOEToolButton1 
            Height          =   585
            Left            =   -60
            TabIndex        =   16
            Top             =   210
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   1032
            Picture         =   "FormMain.frx":9521
            BackColor       =   16119285
            Caption         =   "Employee Registration"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Tahoma"
            FontSize        =   8.25
            ForeColor       =   8421504
         End
         Begin MOVERS.JOEToolButton JOEToolButton2 
            Height          =   570
            Left            =   2940
            TabIndex        =   18
            Top             =   210
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   1005
            Picture         =   "FormMain.frx":9C9B
            BackColor       =   16119285
            Caption         =   "Trip Entry"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Tahoma"
            FontSize        =   8.25
            ForeColor       =   8421504
         End
         Begin MOVERS.JOEToolButton JOEToolButton3 
            Height          =   570
            Left            =   5595
            TabIndex        =   30
            Top             =   210
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   1005
            Picture         =   "FormMain.frx":A415
            BackColor       =   16119285
            Caption         =   "Payroll"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Tahoma"
            FontSize        =   8.25
            ForeColor       =   8421504
         End
      End
   End
   Begin VB.Menu MnuF 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu MEE 
         Caption         =   "Employee Entry"
      End
      Begin VB.Menu MeD 
         Caption         =   "Employee Deductions"
      End
      Begin VB.Menu m1 
         Caption         =   "-"
      End
      Begin VB.Menu MTE 
         Caption         =   "Trip Entry"
      End
      Begin VB.Menu MTS 
         Caption         =   "Trip Status"
      End
      Begin VB.Menu m2 
         Caption         =   "-"
      End
      Begin VB.Menu McP 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu MnuPay 
      Caption         =   "Payroll"
      Visible         =   0   'False
      Begin VB.Menu mnDHP 
         Caption         =   "Drivers and Helpers Personnel"
      End
      Begin VB.Menu MnuOP 
         Caption         =   "Office Personnel"
      End
   End
   Begin VB.Menu mnuRep 
      Caption         =   "Reports"
      Visible         =   0   'False
      Begin VB.Menu mnuEmRep 
         Caption         =   "Employee Report"
      End
      Begin VB.Menu mnuTREP 
         Caption         =   "Trip Report"
      End
      Begin VB.Menu MnuPayRep 
         Caption         =   "Payroll Report"
      End
   End
   Begin VB.Menu MnuSet 
      Caption         =   "Settings"
      Visible         =   0   'False
      Begin VB.Menu mnTper 
         Caption         =   "Truck Parsonnels"
      End
      Begin VB.Menu MnusalWag 
         Caption         =   "Salaries and Wages"
      End
      Begin VB.Menu MnuCus 
         Caption         =   "Customer"
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu MnuAA 
         Caption         =   "About the Author"
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture3_Click()

End Sub
