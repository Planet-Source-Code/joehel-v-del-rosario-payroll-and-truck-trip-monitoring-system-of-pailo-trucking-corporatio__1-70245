VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIMainForm 
   BackColor       =   &H8000000C&
   Caption         =   "822 MOVERS(PAILO) - PAYROLL &  BILLING SYSTEM"
   ClientHeight    =   10980
   ClientLeft      =   3270
   ClientTop       =   30
   ClientWidth     =   15030
   Icon            =   "822MainForm.frx":0000
   LinkTopic       =   "MDIMainForm"
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3705
      Top             =   2430
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":6FCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":7746
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":7EC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":863A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":8DB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":952E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":9CA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":A422
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":AB9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":B316
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":BA90
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":C20A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":C984
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":D0FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":D878
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":DFF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":E76C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":EEE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":F660
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":FDDA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3090
      Top             =   1800
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   3090
      Top             =   2250
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3075
      Top             =   2745
   End
   Begin MOVERS.ACPRibbon ACPRibbon1 
      Align           =   1  'Align Top
      Height          =   1740
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   3069
      BackColor       =   4210752
      ForeColor       =   -2147483630
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ystem"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   8310
         TabIndex        =   2
         Top             =   30
         Width           =   6915
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3690
      Top             =   1785
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   23
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":10554
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":10B09
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":11283
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":119FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":12177
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":1279B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":12D61
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":13314
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":138AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "822MainForm.frx":14029
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MOVERS.JOEClientWin JOEClientWin1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   10605
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   661
   End
   Begin MOVERS.JoeSBCenter JoeSBCenter1 
      Align           =   3  'Align Left
      Height          =   8865
      Left            =   0
      TabIndex        =   3
      Top             =   1740
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   15637
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1515
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   10395
         Width           =   1140
      End
      Begin VB.PictureBox Picture3 
         Height          =   360
         Left            =   165
         ScaleHeight     =   300
         ScaleWidth      =   2430
         TabIndex        =   4
         Top             =   10555
         Width           =   2490
         Begin VB.ListBox List1 
            Height          =   255
            Left            =   30
            TabIndex        =   8
            Top             =   60
            Width           =   1275
         End
         Begin VB.ListBox List2 
            Height          =   255
            Left            =   30
            TabIndex        =   7
            Top             =   360
            Width           =   1275
         End
         Begin VB.ListBox List3 
            Height          =   255
            Left            =   30
            TabIndex        =   6
            Top             =   615
            Width           =   1275
         End
         Begin VB.ListBox List4 
            Height          =   255
            Left            =   30
            TabIndex        =   5
            Top             =   900
            Width           =   1275
         End
      End
      Begin MOVERS.JOESideTab JST 
         Height          =   3015
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   645
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   5318
         Caption         =   "Quick Launch"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   8388608
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
               NumListImages   =   4
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "822MainForm.frx":147A3
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "822MainForm.frx":148B6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "822MainForm.frx":15030
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "822MainForm.frx":157AA
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView listQL 
            Height          =   2670
            Left            =   0
            TabIndex        =   10
            Top             =   345
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   4710
            View            =   2
            Arrange         =   1
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            OLEDropMode     =   1
            _Version        =   393217
            Icons           =   "ilQL"
            SmallIcons      =   "ilQL"
            ForeColor       =   12735512
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
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
         TabIndex        =   11
         Top             =   3675
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   3440
         Caption         =   "Office Personnels"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   8388608
         Begin VB.Label animates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   $"822MainForm.frx":15F24
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   3810
            Index           =   1
            Left            =   180
            TabIndex        =   13
            Top             =   1770
            Width           =   2520
         End
         Begin VB.Label animates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   $"822MainForm.frx":160C4
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   3810
            Index           =   0
            Left            =   180
            TabIndex        =   12
            Top             =   -2025
            Width           =   2520
         End
      End
      Begin MOVERS.JOESideTab JST 
         Height          =   3105
         Index           =   2
         Left            =   60
         TabIndex        =   14
         Top             =   5640
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   5477
         Caption         =   "Employee Picures"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         ForeColor       =   8388608
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            ForeColor       =   &H80000008&
            Height          =   2760
            Left            =   -30
            ScaleHeight     =   2730
            ScaleWidth      =   2940
            TabIndex        =   15
            Top             =   345
            Width           =   2970
            Begin VB.PictureBox Picture1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   2130
               Left            =   540
               ScaleHeight     =   2100
               ScaleWidth      =   1755
               TabIndex        =   16
               Top             =   300
               Width           =   1785
               Begin VB.Image Image1 
                  Height          =   2070
                  Left            =   30
                  Stretch         =   -1  'True
                  Top             =   15
                  Width           =   1710
               End
            End
         End
      End
      Begin VB.Label lblCurrentUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Not Log-in"
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
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   105
         Width           =   2775
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
         ForeColor       =   &H00C25418&
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   180
      End
   End
End
Attribute VB_Name = "MDIMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><>'
'Programmed By: Joehel V. del Rosario    '
'822 MOVERS(PAILO) Trucking Corporation  '
'Date: June 1, 2007                      '
'<><><><><><><><><><><><><><><><><><><><>'

Private Const m_TabShowQuickLaunch = 0
Private Const m_TabSearch = 1
Private Const m_TabFilterDate = 2

Dim Theme As Integer
'Dim fchild As md ChildMDI


Private Sub ACPRibbon1_ButtonClick(ByVal ID As String, ByVal Caption As String)
If ID = 0 Then
    MDIMainForm.AddChild frmTrips, False
    frmTrips.Picture = ACPRibbon1.LoadBackground
    frmTrips.BackColor = ACPRibbon1.BackColor
    'frmTrips.Height = 819
    'frmTrips.Width = 12240
End If

If ID = 1 Then
    MDIMainForm.AddChild frmTripStatus, False
    frmTripStatus.Picture = ACPRibbon1.LoadBackground
    frmTripStatus.BackColor = ACPRibbon1.BackColor
    frmTripStatus.Option1.BackColor = ACPRibbon1.BackColor
    frmTripStatus.Option2.BackColor = ACPRibbon1.BackColor
    frmTripStatus.Picture1.BackColor = ACPRibbon1.BackColor
    frmTripStatus.Check1.BackColor = ACPRibbon1.BackColor
    frmTripStatus.Check2.BackColor = ACPRibbon1.BackColor
    'frmTrips.Height = 819
    'frmTrips.Width = 12240
End If


If ID = 3 Then
    MDIMainForm.AddChild frmEmployeeEntry, False
    frmEmployeeEntry.Picture = ACPRibbon1.LoadBackground
    frmEmployeeEntry.BackColor = ACPRibbon1.BackColor
    'frmTrips.Height = 819
    'frmTrips.Width = 12240
End If

If ID = 4 Then
    MDIMainForm.AddChild frmEmployeeDeductions, False
    frmEmployeeDeductions.Picture = ACPRibbon1.LoadBackground
    frmEmployeeDeductions.BackColor = ACPRibbon1.BackColor
    frmEmployeeDeductions.Option1.BackColor = ACPRibbon1.BackColor
    frmEmployeeDeductions.Option3.BackColor = ACPRibbon1.BackColor
    frmEmployeeDeductions.Option4.BackColor = ACPRibbon1.BackColor
    frmEmployeeDeductions.Option5.BackColor = ACPRibbon1.BackColor
    frmEmployeeDeductions.Option6.BackColor = ACPRibbon1.BackColor
    frmEmployeeDeductions.Option7.BackColor = ACPRibbon1.BackColor
    frmEmployeeDeductions.Option8.BackColor = ACPRibbon1.BackColor
    frmEmployeeDeductions.Option9.BackColor = ACPRibbon1.BackColor
    'frmTrips.Height = 819
    'frmTrips.Width = 12240
End If

If ID = 5 Then
    MDIMainForm.AddChild frmManageEmployee, False
    Call JST_CaptionClick(m_TabFilterDate)
    frmManageEmployee.Picture = ACPRibbon1.LoadBackground
    frmManageEmployee.BackColor = ACPRibbon1.BackColor
    
End If

If ID = 6 Then
    Me.lblCurrentUser.Caption = "Log-Out"
    frmLogin.Show 1
    frmLogin.Picture = ACPRibbon1.LoadBackground
    frmLogin.BackColor = ACPRibbon1.BackColor

End If

If ID = 7 Then
    If MsgBox("Are you sure you want to exit this system?", vbYesNo + vbInformation, "Exit Application") = vbYes Then
        End
    End If
End If

If ID = 8 Then
    MDIMainForm.AddChild frmEmployeePatroll, False
    Call JST_CaptionClick(m_TabFilterDate)
    frmEmployeePatroll.Picture = ACPRibbon1.LoadBackground
    frmEmployeePatroll.BackColor = ACPRibbon1.BackColor

End If

If ID = 9 Then
    MDIMainForm.AddChild FormPAY, False
    Call JST_CaptionClick(m_TabFilterDate)
    FormPAY.Picture = ACPRibbon1.LoadBackground
    FormPAY.BackColor = ACPRibbon1.BackColor
End If


If ID = 11 Then
    MDIMainForm.AddChild frmTripReport, False
    frmTripReport.Picture = ACPRibbon1.LoadBackground
    frmTripReport.BackColor = ACPRibbon1.BackColor
    frmTripReport.Picture1.BackColor = ACPRibbon1.BackColor
    frmTripReport.Option1.BackColor = ACPRibbon1.BackColor
    frmTripReport.Option2.BackColor = ACPRibbon1.BackColor
    frmTripReport.Check1.BackColor = ACPRibbon1.BackColor
    frmTripReport.Check2.BackColor = ACPRibbon1.BackColor
    frmTripReport.Frame1.BackColor = ACPRibbon1.BackColor
    frmTripReport.Frame2.BackColor = ACPRibbon1.BackColor
    frmTripReport.Frame3.BackColor = ACPRibbon1.BackColor
    frmTripReport.Frame4.BackColor = ACPRibbon1.BackColor
End If


If ID = 12 Then
    MDIMainForm.AddChild frmPayrollReport, False
    Call JST_CaptionClick(m_TabFilterDate)
    frmPayrollReport.Picture = ACPRibbon1.LoadBackground
    frmPayrollReport.BackColor = ACPRibbon1.BackColor
    frmPayrollReport.Option1.BackColor = ACPRibbon1.BackColor
    frmPayrollReport.Option2.BackColor = ACPRibbon1.BackColor
    frmPayrollReport.Option3.BackColor = ACPRibbon1.BackColor
    frmPayrollReport.Frame1.BackColor = ACPRibbon1.BackColor
    frmPayrollReport.Frame2.BackColor = ACPRibbon1.BackColor
    frmPayrollReport.Frame3.BackColor = ACPRibbon1.BackColor
End If


If ID = 13 Then
    MDIMainForm.AddChild FormSettings, False
    FormSettings.Picture = ACPRibbon1.LoadBackground
    FormSettings.BackColor = ACPRibbon1.BackColor
    FormSettings.SSTab1.Tab = 0
End If


If ID = 14 Then
    MDIMainForm.AddChild FormSettings, False
    FormSettings.Picture = ACPRibbon1.LoadBackground
    FormSettings.BackColor = ACPRibbon1.BackColor
    FormSettings.SSTab1.Tab = 1
End If

If ID = 15 Then
    MDIMainForm.AddChild FormSettings, False
    FormSettings.Picture = ACPRibbon1.LoadBackground
    FormSettings.BackColor = ACPRibbon1.BackColor
    FormSettings.SSTab1.Tab = 2
End If

If ID = 16 Then
    ChangeThemes 0
    SaveSetting App.EXEName, "APPThemes", "JThemes", 0
End If

If ID = 17 Then
    ChangeThemes 1
    SaveSetting App.EXEName, "APPThemes", "JThemes", 1
End If

If ID = 18 Then
    ChangeThemes 2
    SaveSetting App.EXEName, "APPThemes", "JThemes", 2
    'Trim(GetSetting(App.EXEName, "TextBox", TxtUserID.Name, ""))
End If



End Sub
Sub ChangeThemes(ThemesA As Integer)
Dim i As Integer
Dim FRm As Form
    '# Set Theme
    ACPRibbon1.Theme = Val(ThemesA)
    '# Refresh control
    ACPRibbon1.Refresh
    
        '# OPTIONAL - Load Background for Form.
        MDIMainForm.Picture = ACPRibbon1.LoadBackground
        
        '# OPTIONAL - Load Background for Form
        MDIMainForm.BackColor = ACPRibbon1.BackColor
        MDIMainForm.JoeSBCenter1.STcolor True
        
        '# Set ImageList to use for icons
        ACPRibbon1.ImageList = ImageList2
        
        '# Set Buttons on Center verticaly    (True = Center, False(Default) = Align on Top)
        ACPRibbon1.ButtonCenter = False
        
        
            For Each ctl In MDIMainForm
                'If TypeOf ctl Is Label Then ctl.ForeColor = ACPRibbon1.ForeColor
                'If TypeOf ctl Is Frame Then ctl.BackColor = ACPRibbon1.BackColor
                If TypeOf ctl Is PictureBox Then: ctl.BackColor = ACPRibbon1.BackColor
                'If TypeOf ctl Is CheckBox Then ctl.BackColor = ACPRibbon1.BackColor
                'If TypeOf ctl Is OptionButton Then ctl.BackColor = ACPRibbon1.BackColor
                If TypeOf ctl Is Label Then: ctl.ForeColor = ACPRibbon1.ForeColor
                If TypeOf ctl Is TextBox Then: ctl.ForeColor = ACPRibbon1.ForeColor
                If TypeOf ctl Is ListView Then: ctl.ForeColor = ACPRibbon1.ForeColor
                If TypeOf ctl Is ListView Then: ctl.BackColor = ACPRibbon1.BackColor
                If TypeOf ctl Is JOESideTab Then: ctl.BackColor = ACPRibbon1.BackColor
                If TypeOf ctl Is JOESideTab Then: ctl.ForeColor = ACPRibbon1.ForeColor
            Next
    
'# Search for all MDIChild loaded
    For Each FRm In Forms
        If FRm.Name = "ChildMDI" Then
            '# Change Theme from MDIChild Forms
            FRm.Picture = ACPRibbon1.LoadBackground
            FRm.BackColor = ACPRibbon1.BackColor
            '# Change Forecolor from all Labels on MDIChild forms
            For Each ctl In FRm
                'If TypeOf ctl Is Label Then ctl.ForeColor = ACPRibbon1.ForeColor
                If TypeOf ctl Is Frame Then: ctl.BackColor = ACPRibbon1.BackColor
                If TypeOf ctl Is PictureBox Then: ctl.BackColor = ACPRibbon1.BackColor
                If TypeOf ctl Is CheckBox Then: ctl.BackColor = ACPRibbon1.BackColor
                If TypeOf ctl Is OptionButton Then: ctl.BackColor = ACPRibbon1.BackColor
                If TypeOf ctl Is Label Then: ctl.ForeColor = ACPRibbon1.ForeColor
                If TypeOf ctl Is TextBox Then: ctl.ForeColor = ACPRibbon1.ForeColor
            Next
        End If
        
        'Frm.
    Next
    
End Sub

Private Sub MDIForm_Load()


'If Val(GetSetting(App.EXEName, "APPThemes", "JThemes", "")) < 0 Then
    Theme = Val(GetSetting(App.EXEName, "APPThemes", "JThemes", ""))
'Else
'    Theme = 1
'End If


'SaveSetting App.EXEName, "APPThemes", "JThemes", 3
    'Trim(GetSetting(App.EXEName, "TextBox", TxtUserID.Name, ""))

'# SET Theme
ChangeThemes Theme           ' 0 - Black
                             ' 1 - Blue
                             ' 2 - Silver
                        

'# OPTIONAL - Load Background for Form.
MDIMainForm.Picture = ACPRibbon1.LoadBackground

'# OPTIONAL - Load Background for Form
MDIMainForm.BackColor = ACPRibbon1.BackColor

'# Set ImageList to use for icons
ACPRibbon1.ImageList = ImageList2

'# Set Buttons on Center verticaly    (True = Center, False(Default) = Align on Top)
ACPRibbon1.ButtonCenter = False

'# Add Tabs ---   ID - Caption
ACPRibbon1.AddTab "1", "Transactions"
ACPRibbon1.AddTab "2", "Payroll"
ACPRibbon1.AddTab "3", "Reports"
ACPRibbon1.AddTab "4", "Settings"
ACPRibbon1.AddTab "5", "Help"

'# Add Cats ---   ID - Tab - Caption - ShowDialogButton     'CAT
ACPRibbon1.AddCat "1", "1", "Trip Ticket", False            '1
ACPRibbon1.AddCat "2", "1", "Employee Records", False       '2
ACPRibbon1.AddCat "3", "1", "Exit Application", False       '3
ACPRibbon1.AddCat "4", "2", "Payrolls and Salaries", False  '4
ACPRibbon1.AddCat "5", "3", "Trip Ticket", False            '5
ACPRibbon1.AddCat "6", "4", "Configurations", False         '6
ACPRibbon1.AddCat "7", "4", "Themes", False                 '7
ACPRibbon1.AddCat "8", "5", "Help", False                   '8

'# Add Button ---    ID - Cat - Capt. - Icons -   More Arrow   - ToolTip
'Transactions Tab
ACPRibbon1.AddButton "0", "1", "ENTRY", 3, False, "Entry of all trips for every trucks"
ACPRibbon1.AddButton "1", "1", "STATUS", 2, False, "View all trip status of every trukcs"
ACPRibbon1.AddButton "2", "1", "MANAGE", 1, False, "Manage trip records by trucks"

ACPRibbon1.AddButton "3", "2", "REGISTRATION", 4, False, "Entry of new employee records"
ACPRibbon1.AddButton "4", "2", "DEDUCTIONS", 8, False, "Entry of employee deductions"
ACPRibbon1.AddButton "5", "2", "MANAGE", 5, False, "Manage employee records"

ACPRibbon1.AddButton "6", "3", "LOG OFF", 6, False, "Log off user"
ACPRibbon1.AddButton "7", "3", "CLOSE", 7, False, "Close this application"

'Payroll Tab
ACPRibbon1.AddButton "8", "4", "EMPLOYEE", 9, False, "View and Print employee payroll"
ACPRibbon1.AddButton "9", "4", "OFFICE PERSONNEL", 10, False, "View and Print employee payroll"
ACPRibbon1.AddButton "10", "4", "ADJUSTMENTS", 11, False, "Edit and adjust employee payroll"

'Report Tab
ACPRibbon1.AddButton "11", "5", "TRIP EXPENSE", 12, False, "View the expense report for every trucks"
ACPRibbon1.AddButton "12", "5", "PAYROLL", 13, False, "View the employee payroll amount"



'Settings Tab
ACPRibbon1.AddButton "13", "6", "TRUCK PERSONNELS", 14, False, "View every truck personnels"
ACPRibbon1.AddButton "14", "6", "SALARIES", 15, False, "View salaries and wages"
ACPRibbon1.AddButton "15", "6", "CUSTOMERS", 16, False, "View customers and wages"
ACPRibbon1.AddButton "16", "7", "BLACK", 17, False, "Change themes to color black"
ACPRibbon1.AddButton "17", "7", "BLUE", 18, False, "Change themes to color blue"
ACPRibbon1.AddButton "18", "7", "SELVER", 19, False, "Change themes to color selver"



'HELP Tab
ACPRibbon1.AddButton "19", "8", "HELP", 20, False, "Show the help options about this application"
ACPRibbon1.AddButton "20", "8", "ABOUT", 21, False, "Show the author and the credits who make this program successful"

'ACPRibbon1.AddButton "12", "5", "PAYROLL", 13, False, "View the employee payroll amount"


'# Repaint Ribbon
ACPRibbon1.Refresh

''''===================
    
    'add Quick Launch Items
    With listQL.ListItems
        .Add , "EmR", "Employee Registration", 2, 2
        .Add , "EmD", "Employee Deductions", 2, 2
        .Add , "TrE", "Trip Entry", 2, 2
        .Add , "TrERep", "Trip Expense Reports", 2, 2
        .Add , "EmPay", "Payrolls", 2, 2
        .Add , "PReP", "Payroll Reports", 2, 2
        .Add , "Rank", "Manage Employee Records", 2, 2
        .Add , "Sett", "System Settings", 2, 2
        .Add , "About", "About the Author", 2, 2
    End With
    
    MDIMainForm.AddChild frmWelcome, False
    frmWelcome.Picture = ACPRibbon1.LoadBackground
    frmWelcome.BackColor = ACPRibbon1.BackColor
    
    
    JOEClientWin1.SBWidth = JoeSBCenter1.Width / Screen.TwipsPerPixelX
    Call LoadLISTname
    Call OpenDATECOver
    
    
    
    frmLogin.Show 1
    frmLogin.Picture = ACPRibbon1.LoadBackground
    frmLogin.BackColor = ACPRibbon1.BackColor

    
End Sub
''''''''----------------------------------------------

Public Function ShowForm()
    
 

End Function

Private Sub JOESBtop1_Resize()
    'JoeSBCenter1.Width = JOESBtop1.Width * Screen.TwipsPerPixelX
    'frmWelcome.JOEres
End Sub

Private Sub JOESBtop1_SizeChange(ByVal newSizeState As eSizeState)
    If newSizeState = ssContracted Then
        JOEClientWin1.SBWidth = JoeSBCenter1.Width / Screen.TwipsPerPixelX
        JoeSBCenter1.Visible = True
        frmWelcome.JOEres 3280
    Else
        JOEClientWin1.SBWidth = 0
        JoeSBCenter1.Visible = False
        frmWelcome.JOEres 180
        
    End If
    'call mdi resize to resize all opened child forms
End Sub

Private Sub jst_BeforeExpand(Index As Integer)
    'resize contained controlsbeofre expanding
    Select Case Index
        Case m_TabShowQuickLaunch
            listQL.Move 90, listQL.Top, JST(Index).Width - 150

        Case m_TabSearch 'search
            'resize
            'Picture1.Move 90, Picture1.Top, JST(Index).Width - 150
        Case m_TabFilterDate 'filter date
            'If Form_CanManageEmployee = False Then
                'MsgBox "There is no Employee's Picture to display.", vbInformation
                'JST(Index).Expanded = False
            'Else
            Picture2.Move 90, Picture2.Top, JST(Index).Width - 150
            'End If
    End Select

End Sub

Private Sub JST_CaptionClick(Index As Integer)
    Select Case Index
        Case m_TabShowQuickLaunch
            JST(m_TabShowQuickLaunch).Height = 3015

        Case m_TabSearch 'search
            JST(m_TabSearch).Height = 1950
        Case m_TabFilterDate 'filter date
            JST(m_TabFilterDate).Height = 3660
    End Select
    
End Sub

Private Sub jst_CompleteExpand(Index As Integer)
    Dim i As Integer
    
    For i = 0 To 2
        If Index <> i Then
            If JST(i).AutoContract = True Then
                JST(i).Expanded = False
            End If
        End If
    Next

End Sub

Private Sub jst_Resize(Index As Integer)
    
    Dim i As Integer
    
    For i = 1 To 2
        JST(i).Move JST(i).Left, (JST(i - 1).Top + JST(i - 1).Height) '- 200
    Next
    
    If JST(Index).Expanded = True Then
        Select Case Index
            Case m_TabShowQuickLaunch
                listQL.Move 90, listQL.Top, JST(Index).Width - 150

            Case m_TabSearch 'search
                'resize
                'Picture1.Move 90, Picture1.Top, JST(Index).Width - 150
            Case m_TabFilterDate 'filter date
                Picture2.Move 90, Picture2.Top, JST(Index).Width - 150
        End Select
    End If

End Sub
Private Sub JOEClientWin1_CloseClick(ByVal sFormName As String, ByVal Index As Integer)
    'close form
    Dim FRm As Form
    
    On Error GoTo RAE
    
    For Each FRm In Forms
        If LCase(Trim(FRm.Name)) = LCase(Trim(sFormName)) Then
            Unload FRm
            Exit For
        End If
    Next
    
RAE:
    Set FRm = Nothing

End Sub

Private Sub JOEClientWin1_FormTabClick(ByVal sFormName As String, ByVal Index As Integer)
    modFuncChild.ActivateMDIChildForm sFormName
End Sub

Public Function Form_ShowQuickLaunch()

    'expand side bar
    'If JOESBtop1.SizeState <> ssContracted Then
    '    JOESBtop1.SizeState = ssContracted
    'End If

    'expand search tab
    If JST(m_TabShowQuickLaunch).Expanded = False Then
        JST(m_TabShowQuickLaunch).Expanded = True
    End If
    
    On Error Resume Next
    JST(m_TabShowQuickLaunch).SetFocus
    'HLTxt txtSearchWhat
    Err.Clear
    
End Function

Public Function Form_ShowSearch()

    'expand side bar
    'If JOESBtop1.SizeState <> ssContracted Then
    '    JOESBtop1.SizeState = ssContracted
    'End If

    'expand search tab
    If JST(m_TabSearch).Expanded = False Then
        JST(m_TabSearch).Expanded = True
    End If
    
    On Error Resume Next
    JST(m_TabSearch).SetFocus
    'HLTxt txtSearchWhat
    Err.Clear
    
End Function


Public Function Form_ShowDateFilter()

    'expand side bar
    'If JOESBtop1.SizeState <> ssContracted Then
    '    JOESBtop1.SizeState = ssContracted
    'End If

    'expand search tab
    If JST(m_TabFilterDate).Expanded = False Then
        JST(m_TabFilterDate).Expanded = True
    End If
    
    On Error Resume Next
    JST(m_TabFilterDate).SetFocus
    'b8DateP.SetFocus
    Err.Clear
    
End Function


Private Sub Label5_Click()
    FormCoveredDate.Show
End Sub

Private Sub listQL_Click()
    Dim selItemKey As String
    
    On Error GoTo RAE
    
    selItemKey = listQL.SelectedItem.Key
    
    Select Case selItemKey
        Case "EmR" 'Employee Registration"
            ACPRibbon1_ButtonClick 3, ""
        Case "EmD" 'Employee Deductions"
            ACPRibbon1_ButtonClick 4, ""
        Case "TrE" 'Trip Entry"
            ACPRibbon1_ButtonClick 0, ""
        Case "EmPay" 'Payrolls"
            ACPRibbon1_ButtonClick 8, ""
        Case "PReP" 'Payroll Reports"
            ACPRibbon1_ButtonClick 12, ""
        Case "Rank" 'Manage Employee List"
            ACPRibbon1_ButtonClick 5, ""
        Case "TrERep" ' trip Exp. Report
            ACPRibbon1_ButtonClick 11, ""
        Case "Sett"
            ACPRibbon1_ButtonClick 13, ""
            'MDIMainForm.AddChild FormSettings, False
        Case "About"
            MDIMainForm.AddChild FormPAY, False
    End Select
  
RAE:

End Sub
Sub ShowSideTab()
    Call JOESBtop1_SizeChange(ssContracted)
End Sub
Private Sub McP_Click()
    End
End Sub

Private Sub listQL_KeyPress(KeyAscii As Integer)
    'Call listQL_DblClick
End Sub

'Private Sub MDIForm_Load()
'End Sub
Sub OpenDATECOver()
On Error Resume Next
If Label5.Caption <> "" Then
    OpenPBDataBase ("DateCover")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM DateCover WHERE Status Like '" & "1" & "' ")
    With PRFile
        If Not .EOF Then
            Label5.Caption = "Covered date for Salary : " & ![CoveredDate]
        Else
            Label5.Caption = "Date NOT SET"
        End If
   End With
 End If
End Sub
Sub LoadLISTname()
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    OpenPBDataBase ("EmployeeInfo")
    With PRFile
        If .RecordCount > 1 Then
            Do Until .EOF
                List1.AddItem ![Ename]
                List2.AddItem ![ECOde]
                If ![EOccupation] = "Helper" Then
                    List3.AddItem ![Ename]
                Else
                    List4.AddItem ![Ename]
                End If
            .MoveNext
            Loop
        End If
        .Close
    End With
End Sub
Private Sub mnuC_Click()
    End
End Sub
Private Sub Timer1_Timer()
    lblDate.Caption = FormatDateTime(Now, vbLongDate)
End Sub

' MDI Form procedures
'-----------------------------------------------------------
Private Sub MDIForm_Resize()
        
    Dim FRm As Form
    
    
    On Error Resume Next
    
    
    'resize childs
    If GetActiveChildCount > 0 Then
        For Each FRm In Forms
        If FRm.Name <> Me.Name Then
            If FRm.MDIChild = True Then
                If FRm.Name = Me.ActiveForm.Name Then
                    ResizeMdiChildForm FRm
                Else
                    FRm.Visible = False
                End If
            End If
        End If
        
        Next
        
    End If
    
    Set FRm = Nothing
End Sub

'Get Opened MDI Child Forms Count
Public Function GetActiveChildCount() As Integer
    
    Dim FRm As Form
    Dim iCount As Integer
    
    iCount = 0
    
    For Each FRm In Forms
        If FRm.Name <> Me.Name Then
            If FRm.MDIChild = True Then
                iCount = iCount + 1
            End If
        End If
    Next
    
    GetActiveChildCount = iCount
    Set FRm = Nothing
    
End Function
'-----------------------------------------------------------
' >> End MDI Form procedures
'------------------------------------------------------------

'------------------------------------------------------------
' Parent To Child procedures
'------------------------------------------------------------

Public Sub AddChild(ByRef CFrm As Form, Optional CloseButton As Boolean = True)

    'load form
    modFuncChild.LoadForm CFrm, CloseButton
    
End Sub



Public Sub ActivateChild(ByRef CFrm As Form)
    'activate form
    Me.JOEClientWin1.SetActiveWindow CFrm.Name
    Form_CanManageEmployee
End Sub
Public Function Form_CanManageEmployee() As Boolean
    
    Dim bReturn As Boolean
    
    On Error Resume Next
    
    bReturn = False
    bReturn = Me.ActiveForm.Form_CanManageEmployee
            
    If bReturn = False Then
        JST(m_TabFilterDate).Expanded = False
    End If
    
    Form_CanManageEmployee = bReturn
    
    Err.Clear
    
End Function

Public Sub RemoveChild(ByVal sFormName As String)
    'remove form
     Me.JOEClientWin1.RemoveChildWindow sFormName
End Sub
Private Sub Timer2_Timer()
    If animates(0).Top = -2055 Then
        animates(1).Top = 1710
    End If
    
    If animates(1).Top = -2055 Then
        animates(0).Top = 1710
    End If
    
    animates(0).Top = animates(0).Top - 5
    animates(1).Top = animates(1).Top - 5
End Sub

Private Sub Timer3_Timer()
    Label5.Visible = Not Label5.Visible
End Sub

Public Sub AFForm_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 83 And Shift = 4 Then
        
    ElseIf KeyCode = 82 And Shift = 4 Then
        
    ElseIf KeyCode = 77 And Shift = 4 Then
        
    ElseIf KeyCode = 84 And Shift = 4 Then
        
    ElseIf KeyCode = 72 And Shift = 4 Then
        
    ElseIf KeyCode = 81 And Shift = 2 Then

    ElseIf KeyCode = 68 And Shift = 2 Then

    End If
    
    'MsgBox KeyCode & " - " & Shift
End Sub


