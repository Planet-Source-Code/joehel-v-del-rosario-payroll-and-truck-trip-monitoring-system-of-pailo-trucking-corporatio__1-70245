VERSION 5.00
Begin VB.Form frmTrips 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Trip Entry"
   ClientHeight    =   8385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   559
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   816
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Gate Pass"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4395
      TabIndex        =   91
      Top             =   1095
      Width           =   1155
   End
   Begin VB.PictureBox GPass 
      BackColor       =   &H0080FFFF&
      Height          =   1830
      Left            =   4305
      ScaleHeight     =   1770
      ScaleWidth      =   3735
      TabIndex        =   88
      Top             =   4065
      Visible         =   0   'False
      Width           =   3795
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "Billings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   2145
         TabIndex        =   93
         Top             =   1305
         Value           =   -1  'True
         Width           =   1110
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "Trip Ticket"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   645
         TabIndex        =   92
         Top             =   1290
         Width           =   1305
      End
      Begin VB.TextBox txtGpass 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   315
         TabIndex        =   90
         Text            =   "000000"
         Top             =   420
         Width           =   3285
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gate Pass Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   0
         TabIndex        =   89
         Top             =   0
         Width           =   3690
      End
   End
   Begin VB.PictureBox sts 
      BackColor       =   &H0080FFFF&
      Height          =   2700
      Left            =   3165
      ScaleHeight     =   2640
      ScaleWidth      =   5775
      TabIndex        =   56
      Top             =   4545
      Visible         =   0   'False
      Width           =   5835
      Begin MOVERS.JOELine JOELine1 
         Height          =   30
         Left            =   30
         TabIndex        =   57
         Top             =   375
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   53
         BorderColor1    =   11325655
         BorderColor2    =   16185592
      End
      Begin MOVERS.CandyButton CandyButton5 
         Height          =   405
         Left            =   2295
         TabIndex        =   65
         Top             =   2160
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "        Save"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "frmTrips.frx":0000
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
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   0
         TabIndex        =   64
         Top             =   420
         Width           =   5670
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   30
         TabIndex        =   63
         Top             =   765
         Width           =   5670
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   0
         TabIndex        =   62
         Top             =   1035
         Width           =   5670
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   0
         TabIndex        =   61
         Top             =   1290
         Width           =   5670
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   0
         TabIndex        =   60
         Top             =   1560
         Width           =   5670
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   0
         TabIndex        =   59
         Top             =   1830
         Width           =   5670
      End
      Begin VB.Image imgClose 
         Height          =   360
         Index           =   1
         Left            =   5385
         Picture         =   "frmTrips.frx":077A
         Top             =   30
         Width           =   360
      End
      Begin VB.Image imgClose 
         Height          =   360
         Index           =   0
         Left            =   5385
         Picture         =   "frmTrips.frx":0E64
         Top             =   30
         Width           =   360
      End
      Begin VB.Label Labelme 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Trip Salary"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   375
         TabIndex        =   58
         Top             =   90
         Width           =   3255
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   2610
         Left            =   0
         Top             =   15
         Width           =   5760
      End
      Begin VB.Image Image1 
         Height          =   285
         Left            =   75
         Picture         =   "frmTrips.frx":154E
         Stretch         =   -1  'True
         Top             =   45
         Width           =   255
      End
   End
   Begin VB.PictureBox CompanyEX 
      BackColor       =   &H0080FFFF&
      Height          =   2340
      Left            =   1605
      ScaleHeight     =   2280
      ScaleWidth      =   9045
      TabIndex        =   66
      Top             =   4920
      Visible         =   0   'False
      Width           =   9105
      Begin VB.TextBox Text20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5250
         TabIndex        =   76
         Top             =   1275
         Width           =   1560
      End
      Begin VB.TextBox Text19 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5250
         TabIndex        =   75
         Top             =   915
         Width           =   1560
      End
      Begin VB.TextBox Text18 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5250
         TabIndex        =   74
         Top             =   540
         Width           =   1560
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1530
         TabIndex        =   73
         Top             =   1245
         Width           =   1560
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1530
         TabIndex        =   72
         Top             =   885
         Width           =   1560
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1530
         TabIndex        =   71
         Top             =   525
         Width           =   1560
      End
      Begin VB.TextBox Text28 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   7020
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   70
         Top             =   795
         Width           =   1965
      End
      Begin MOVERS.JOELine JOELine2 
         Height          =   30
         Left            =   15
         TabIndex        =   67
         Top             =   375
         Width           =   9720
         _ExtentX        =   17145
         _ExtentY        =   53
         BorderColor1    =   11325655
         BorderColor2    =   16185592
      End
      Begin MOVERS.CandyButton CandyButton6 
         Height          =   405
         Left            =   5250
         TabIndex        =   68
         Top             =   1710
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   714
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "        Save"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "frmTrips.frx":1AD8
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
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   2265
         Left            =   0
         Top             =   0
         Width           =   9030
      End
      Begin VB.Image imgClose 
         Height          =   360
         Index           =   2
         Left            =   8655
         Picture         =   "frmTrips.frx":2252
         Top             =   15
         Width           =   360
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Other Charges :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3360
         TabIndex        =   85
         Top             =   1335
         Width           =   1980
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "LTO/TMG  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4260
         TabIndex        =   84
         Top             =   945
         Width           =   1575
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "LOAD  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4545
         TabIndex        =   83
         Top             =   570
         Width           =   1575
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "TOLL FEE  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   540
         TabIndex        =   82
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Meal Allowance :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   75
         TabIndex        =   81
         Top             =   945
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "GAS and OIL :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   80
         Top             =   510
         Width           =   1575
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1665
         TabIndex        =   79
         Top             =   1845
         Width           =   1710
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL EXPENSE :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   78
         Top             =   1830
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Other Charges Info  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7020
         TabIndex        =   77
         Top             =   495
         Width           =   1815
      End
      Begin VB.Image Image3 
         Height          =   285
         Left            =   75
         Picture         =   "frmTrips.frx":293C
         Stretch         =   -1  'True
         Top             =   45
         Width           =   255
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Trip Expenses"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   375
         TabIndex        =   69
         Top             =   90
         Width           =   3255
      End
      Begin VB.Image imgClose 
         Height          =   360
         Index           =   3
         Left            =   8655
         Picture         =   "frmTrips.frx":2EC6
         Top             =   15
         Width           =   360
      End
   End
   Begin MOVERS.CandyButton ButEmpty 
      Height          =   435
      Left            =   10590
      TabIndex        =   50
      Top             =   7515
      Width           =   1485
      _ExtentX        =   2619
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
      Caption         =   "        Empty Trip"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "frmTrips.frx":35B0
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
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   2640
      MultiSelect     =   2  'Extended
      TabIndex        =   7
      Top             =   5295
      Visible         =   0   'False
      Width           =   9345
   End
   Begin VB.ComboBox Combo8 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmTrips.frx":3D2A
      Left            =   8730
      List            =   "frmTrips.frx":3D2C
      TabIndex        =   49
      Top             =   2700
      Width           =   3300
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6390
      TabIndex        =   29
      Top             =   5520
      Width           =   5595
   End
   Begin VB.TextBox Text24 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2655
      TabIndex        =   28
      Top             =   6510
      Width           =   1965
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2655
      TabIndex        =   27
      Top             =   6945
      Width           =   1965
   End
   Begin VB.ComboBox Combo10 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmTrips.frx":3D2E
      Left            =   8565
      List            =   "frmTrips.frx":3D3E
      TabIndex        =   26
      Top             =   4500
      Width           =   3435
   End
   Begin VB.ComboBox Combo9 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmTrips.frx":3D6F
      Left            =   8355
      List            =   "frmTrips.frx":3D71
      TabIndex        =   25
      Top             =   510
      Width           =   3630
   End
   Begin VB.ComboBox Combo7 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmTrips.frx":3D73
      Left            =   8730
      List            =   "frmTrips.frx":3D75
      TabIndex        =   24
      Top             =   2130
      Width           =   3315
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2655
      TabIndex        =   23
      Top             =   6015
      Width           =   1965
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2655
      TabIndex        =   22
      Top             =   5505
      Width           =   1980
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2655
      TabIndex        =   21
      Top             =   4995
      Width           =   9330
   End
   Begin VB.ComboBox Combo5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmTrips.frx":3D77
      Left            =   2655
      List            =   "frmTrips.frx":3D79
      TabIndex        =   20
      Top             =   4485
      Width           =   3465
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      ItemData        =   "frmTrips.frx":3D7B
      Left            =   1650
      List            =   "frmTrips.frx":3D7D
      TabIndex        =   19
      Top             =   1080
      Width           =   2355
   End
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmTrips.frx":3D7F
      Left            =   1650
      List            =   "frmTrips.frx":3D81
      TabIndex        =   18
      Text            =   "822 MOVERS (Pailo)"
      Top             =   525
      Width           =   4785
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmTrips.frx":3D83
      Left            =   8355
      List            =   "frmTrips.frx":3D85
      TabIndex        =   17
      Top             =   1380
      Width           =   3630
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmTrips.frx":3D87
      Left            =   8355
      List            =   "frmTrips.frx":3D89
      TabIndex        =   16
      Top             =   960
      Width           =   3630
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1650
      TabIndex        =   15
      Top             =   3765
      Width           =   3660
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1650
      TabIndex        =   14
      Top             =   3390
      Width           =   3660
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1650
      TabIndex        =   13
      Top             =   3015
      Width           =   3660
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1650
      TabIndex        =   12
      Top             =   2640
      Width           =   3660
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1650
      TabIndex        =   11
      Top             =   2250
      Width           =   3660
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1650
      TabIndex        =   10
      Top             =   1845
      Width           =   3660
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10365
      TabIndex        =   9
      Top             =   3255
      Width           =   1650
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10380
      TabIndex        =   8
      Top             =   3690
      Width           =   1635
   End
   Begin VB.TextBox Text21 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5325
      TabIndex        =   6
      Top             =   1845
      Width           =   1000
   End
   Begin VB.TextBox Text22 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5325
      TabIndex        =   5
      Top             =   2250
      Width           =   1000
   End
   Begin VB.TextBox Text23 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5325
      TabIndex        =   4
      Top             =   2640
      Width           =   1000
   End
   Begin VB.TextBox Text25 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5325
      TabIndex        =   3
      Top             =   3015
      Width           =   1000
   End
   Begin VB.TextBox Text26 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5325
      TabIndex        =   2
      Top             =   3405
      Width           =   1000
   End
   Begin VB.TextBox Text27 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5325
      TabIndex        =   1
      Top             =   3765
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "....."
      Height          =   435
      Left            =   3990
      TabIndex        =   0
      Top             =   1095
      Width           =   390
   End
   Begin MOVERS.CandyButton ButPreview 
      Height          =   435
      Left            =   9030
      TabIndex        =   51
      Top             =   7515
      Width           =   1485
      _ExtentX        =   2619
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
      Caption         =   "        Preview"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "frmTrips.frx":3D8B
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
      Left            =   7455
      TabIndex        =   52
      Top             =   7515
      Width           =   1485
      _ExtentX        =   2619
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
      Caption         =   "        Search Trip"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "frmTrips.frx":4505
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
   Begin MOVERS.CandyButton CandyButton2 
      Height          =   435
      Left            =   5865
      TabIndex        =   53
      Top             =   7515
      Width           =   1485
      _ExtentX        =   2619
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
      Caption         =   "       New Trip"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "frmTrips.frx":4C7F
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
   Begin MOVERS.CandyButton CandyButton3 
      Height          =   435
      Left            =   4290
      TabIndex        =   54
      Top             =   7515
      Width           =   1485
      _ExtentX        =   2619
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
      Caption         =   "        Save Trip"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "frmTrips.frx":53F9
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
   Begin MOVERS.CandyButton CandyButton4 
      Height          =   435
      Left            =   135
      TabIndex        =   55
      Top             =   7500
      Width           =   2010
      _ExtentX        =   3545
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
      Caption         =   "       Trip Expenses"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "frmTrips.frx":5B73
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
   Begin MOVERS.JOETitleBar JOETitleBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   87
      Top             =   0
      Width           =   12240
      _ExtentX        =   21590
      _ExtentY        =   661
      Caption         =   "Trip Entry"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   8388608
      ShadowColor     =   49344
      BorderColor     =   49344
      BackColor       =   12648447
   End
   Begin VB.Label labelGpass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7425
      TabIndex        =   94
      Top             =   3525
      Width           =   1995
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00808080&
      Height          =   750
      Left            =   7380
      Top             =   3345
      Width           =   2115
   End
   Begin VB.Label TRStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "======="
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   4755
      TabIndex        =   86
      Top             =   6330
      Visible         =   0   'False
      Width           =   7245
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Source:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5535
      TabIndex        =   48
      Top             =   5535
      Width           =   1980
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "Expenses :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1500
      TabIndex        =   47
      Top             =   6525
      Width           =   1875
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1650
      TabIndex        =   46
      Top             =   6930
      Width           =   1005
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6915
      TabIndex        =   45
      Top             =   4470
      Width           =   1725
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Point Of Origin :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   990
      TabIndex        =   44
      Top             =   4485
      Width           =   1725
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Covered Date:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6900
      TabIndex        =   43
      Top             =   540
      Width           =   2250
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Truck Type :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7380
      TabIndex        =   42
      Top             =   2760
      Width           =   1725
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Helper  :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9495
      TabIndex        =   41
      Top             =   3720
      Width           =   1050
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Wage Trip :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7485
      TabIndex        =   40
      Top             =   2190
      Width           =   1725
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Driver  :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9510
      TabIndex        =   39
      Top             =   3270
      Width           =   855
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00808080&
      Height          =   2190
      Left            =   7380
      Top             =   1905
      Width           =   4830
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Trip Allowance : "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   975
      TabIndex        =   38
      Top             =   6015
      Width           =   1755
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Number Cases :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1020
      TabIndex        =   37
      Top             =   5430
      Width           =   1785
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer and Address :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   195
      TabIndex        =   36
      Top             =   4995
      Width           =   2445
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   3090
      Left            =   45
      Top             =   4290
      Width           =   12165
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Personels of :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   255
      TabIndex        =   35
      Top             =   540
      Width           =   1470
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Trip Number:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6855
      TabIndex        =   34
      Top             =   1380
      Width           =   1485
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Trip Date :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6990
      TabIndex        =   33
      Top             =   945
      Width           =   1350
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Helpers :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   720
      TabIndex        =   32
      Top             =   2265
      Width           =   1035
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Driver :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   885
      TabIndex        =   31
      Top             =   1845
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Plate Number :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   -405
      TabIndex        =   30
      Top             =   1140
      Width           =   2025
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00808080&
      Height          =   7665
      Left            =   45
      Top             =   435
      Width           =   12165
   End
End
Attribute VB_Name = "frmTrips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><>'
'Programmed By: Joehel V. del Rosario    '
'822 MOVERS(PAILO) Trucking Corporation  '
'Date: June 1, 2007                      '
'<><><><><><><><><><><><><><><><><><><><>'

Dim HelperT As Integer
Dim tmpOTHER As Double
Dim tmpOtherT As Double
Sub SaveCharges(Ename As String, TP As Integer)
Dim aaa As Double
    OpenPBDataBase ("DeductionsInfo")
        Set PRFile = PDbase.OpenRecordset("SELECT * FROM DeductionsInfo WHERE DName Like '" & Trim(Ename) & "' and  DDate Like '" & Trim(Combo3.Text) & _
                                        "' and  DType Like '" & Trim(Combo6.Text) & "' ")
        With PRFile
            If Not .EOF Then
                .Edit
                    ![DAmount] = ![DAmount] + Val(Text14.Text)
                .Update
            Else
              If Ename <> "" Then
                .AddNew
                    ![DName] = Trim(Ename)
                    ![ddate] = Trim(Combo3.Text)
                    ![DType] = Trim(Combo6.Text)
                    If Option2.Value = False Then
                        ![DAmount] = Round(Val(Text14.Text) / Val(TP), 2)
                    Else
                        ![DAmount] = Round(Val(Text14.Text), 2)
                        MsgBox "Cute"
                    End If
                    
                    ![Status] = "1"
                .Update
              End If
            End If
            .Close
        End With
End Sub

Private Sub ButEmpty_Click()
         Combo5.Text = "EMPTY"
         Text12.Text = "EMPTY"
         Text9.Text = "EMPTY"
         Combo10.Text = "EMPTY"
         Combo7.Text = "EMPTY"
         Text6.Text = "EMPTY"
         Text13.SetFocus
End Sub

Private Sub ButPreview_Click()
    Call OPENTrucExpense
    FormTRIPprint.Show 1
End Sub

Private Sub CandyButton1_Click()
        Call SearchTRIPexist
        Call Combo7_Click
End Sub

Private Sub CandyButton2_Click()
         Text13.Text = ""
         Combo5.Text = ""
         Text12.Text = ""
         Text9.Text = ""
         Combo10.Text = ""
         Combo7.Text = ""
         Combo8.Text = ""
         Text7.Text = ""
         Text8.Text = ""
         Text11.Text = ""
         Text4.Text = ""
         Text10.Text = ""
         Text2.Text = ""
         Text1.Text = ""
         Text3.Text = ""
         Text24.Text = ""
         Text6.Text = ""
         Text5.Text = ""
         Combo4.Text = ""
         Text15.Text = ""
         Text16.Text = ""
         Text17.Text = ""
         Text18.Text = ""
         Text19.Text = ""
         Text20.Text = ""
         Text28.Text = ""
         Text21.Text = ""
         Text22.Text = ""
         Text23.Text = ""
         Text25.Text = ""
         Text26.Text = ""
         Text27.Text = ""
                Text3.BackColor = &HE0E0E0
                Text1.BackColor = &HE0E0E0
                Text2.BackColor = &HE0E0E0
                Text10.BackColor = &HE0E0E0
                Text4.BackColor = &HE0E0E0
                Text11.BackColor = &HE0E0E0

         Combo3.Text = Format(Date, "MM/DD/YYYY")
    Combo1.Clear
    For a = 1 To 5
        Combo1.AddItem a
    Next a
    
    TRStatus.Visible = False
    Call LOadCOVERDATE
    Call LOADOrigin
    Combo1.Text = "1"
End Sub

Private Sub CandyButton3_Click()
If Combo1.Text <> "" And Combo4.Text <> "" And Combo9.Text <> "" And Combo3.Text <> "" Then
        If Combo7.Text <> "EMPTY" Then
            Label16.Caption = Text3.Text & "   =   " & Text21.Text
            Label17.Caption = Text1.Text & "   =   " & Text22.Text
            Label30.Caption = Text2.Text & "   =   " & Text23.Text
            Label31.Caption = Text10.Text & "   =   " & Text25.Text
            Label35.Caption = Text4.Text & "   =   " & Text26.Text
            Label36.Caption = Text11.Text & "   =   " & Text27.Text
            
         If TRStatus.Visible = False Then
            sts.Visible = True
         End If
            
            
            
        End If
            Call saveTripsInfo
Else
    MsgBox "Invalid Null.. Please check the date..", vbCritical, "Error"
End If

End Sub

Private Sub CandyButton4_Click()
    CompanyEX.Visible = True
    Call OPENTrucExpense
    Text15.SetFocus
End Sub

Private Sub CandyButton5_Click()
    Call SAVEPayrolls
    sts.Visible = False
    Call Command1_Click
End Sub

Private Sub CandyButton6_Click()
    
    Text24.Text = Val(Text15.Text) + Val(Text16.Text) + Val(Text17.Text) + Val(Text18.Text) + Val(Text19.Text) + Val(Text20.Text)
    Text5.Text = Val(Text13.Text) - Val(Text24)
    Call SAVETrucExpense
    CompanyEX.Visible = False
    
    'Call saveTripsInfo
    
    Call CandyButton3_Click
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            Call CandyButton1_Click
            Combo5.SetFocus
    End If
End Sub
Private Sub Combo10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text9.SetFocus
    End If
End Sub
Private Sub Combo3_Click()
    Call CandyButton1_Click
End Sub
Private Sub Combo3_KeyPress(KeyAscii As Integer)
  Dim txtL As Integer
   If KeyAscii = 13 Then
            txtL = Combo3.SelLength
            txtL = Combo3.SelLength
            txtL = Combo3.SelLength
            
            If txtL >= 11 Or txtL <= 9 Then
                MsgBox "Invalid Date Format", vbCritical, "Invalid"
                Combo3.SetFocus
                SendKeys "{HOME}+{END}"
            Else
                Combo1.SetFocus
            End If
            
            
   End If
End Sub
Private Sub Combo4_Click()
If Combo1.Text <> "" And Combo4.Text <> "" And Combo9.Text <> "" And Combo3.Text <> "" Then
    Text1.Text = ""
    Text2.Text = ""
    Text10.Text = ""
    Text4.Text = ""
    Text11.Text = ""
    
                Text3.BackColor = &HE0E0E0
                Text1.BackColor = &HE0E0E0
                Text2.BackColor = &HE0E0E0
                Text10.BackColor = &HE0E0E0
                Text4.BackColor = &HE0E0E0
                Text11.BackColor = &HE0E0E0
    
    OpenPBDataBase ("TruckPersonel")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM TruckPersonel WHERE PlateNumber Like '" & Combo4.Text & "' ")
    With PRFile
        If Not .EOF Then
            Text3.Text = ![Driver]
            Combo8.Text = ![Tructype]
            '1
            If ![Helper1] <> "" Then
                Text1.Text = ![Helper1]
                HelperT = 1
            End If
            '2
            If ![Helper2] <> "" Then
                Text2.Text = ![Helper2]
                HelperT = 2
            End If
            '3
            If ![Helper3] <> "" Then
                Text10.Text = ![Helper3]
                HelperT = 3
            End If
            '4
            If ![Helper4] <> "" Then
                Text4.Text = ![Helper4]
                HelperT = 4
            End If
            '5
            If ![Helper5] <> "" Then
                Text11.Text = ![Helper5]
                HelperT = 5
            End If
        Else
            MsgBox "Plate Number " & Text5.Text & " Not Exist in the Truck personel", vbInformation, "Not Exist"
            Exit Sub
        End If
      .Close
    End With
 On Error GoTo errt
 Dim traper As Integer
 Combo3.Clear
 Combo1.Text = "1"
    OpenPBDataBase ("TripInfo")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM TripInfo WHERE TruckNumber Like '" & Combo4.Text & "'  and CoverDate Like '" & Combo9.Text & "' ")

    With PRFile
       .MoveFirst
        Do While Not .EOF
         If Not .EOF Then
           Combo3.AddItem ![Tripdate]
           traper = 1
         End If
           .MoveNext
        Loop
        .Close
    End With
    Combo3.Text = Format(Now, "MM/DD/YYYY")
errt:
    If traper = 0 Then
        Combo3.Text = Format(Now, "MM/DD/YYYY")
    End If
    Combo3.SetFocus
    'SendKeys "{HOME}+{END}+"
 Else
    MsgBox "Invalid Null.. Please check the date..", vbCritical, "Error"
 End If
End Sub
Private Sub Combo4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Combo4_Click
    End If
End Sub
Sub OPENWAGES()
    'Open Wages
    Text21.Text = ""
    Text22.Text = ""
    Text23.Text = ""
    Text25.Text = ""
    Text26.Text = ""
    Text27.Text = ""
    
    
    OpenPBDataBase ("SalariesWages")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM SalariesWages WHERE POriginCode Like '" & Trim(Combo7.Text) & "' and TruckTypes Like '" & Combo8.Text & "' ")
    With PRFile
        If Not .EOF Then
            Text7.Text = ![DriverSalary]
            Text21.Text = ![DriverSalary]
            If HelperT = 1 Then
                Text8.Text = ![Helper1]
                Text22.Text = Trim(Text8.Text)
                
            ElseIf HelperT = 2 Then
                Text8.Text = ![Helper2]
                Text22.Text = Trim(Text8.Text)
                Text23.Text = Trim(Text8.Text)
                
            ElseIf HelperT = 3 Then
                Text8.Text = ![Helper3]
                Text22.Text = Trim(Text8.Text)
                Text23.Text = Trim(Text8.Text)
                Text25.Text = Trim(Text8.Text)
                
            ElseIf HelperT = 4 Then
                Text8.Text = ![Helper4]
                Text22.Text = Trim(Text8.Text)
                Text23.Text = Trim(Text8.Text)
                Text25.Text = Trim(Text8.Text)
                Text26.Text = Trim(Text8.Text)
                
            ElseIf HelperT = 5 Then
                Text8.Text = ![Helper5]
                Text22.Text = Trim(Text8.Text)
                Text23.Text = Trim(Text8.Text)
                Text25.Text = Trim(Text8.Text)
                Text26.Text = Trim(Text8.Text)
                Text27.Text = Trim(Text8.Text)
            End If
        Else
            Text7.Text = ""
            Text8.Text = ""
        End If
       .Close
    End With
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo10.SetFocus
    End If
End Sub

Private Sub Combo7_Click()
            '1
            If Text1.Text <> "" Then
                HelperT = 1
            End If
            '2
            If Text2.Text <> "" Then
                HelperT = 2
            End If
            '3
            If Text10.Text <> "" Then
                HelperT = 3
            End If
            '4
            If Text4.Text <> "" Then
                HelperT = 4
            End If
            '5
            If Text11.Text <> "" Then
                HelperT = 5
            End If
    
    Call OPENWAGES
    Combo5.SetFocus
End Sub

Private Sub Combo7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Combo7_Click
    End If
End Sub

Private Sub Combo9_Click()
    Call Combo4_Click
    Call SearchTRIPexist
End Sub

Private Sub Command1_Click()
    Text3.BackColor = &HE0E0E0
    Text1.BackColor = &HE0E0E0
    Text2.BackColor = &HE0E0E0
    Text10.BackColor = &HE0E0E0
    Text4.BackColor = &HE0E0E0
    Text11.BackColor = &HE0E0E0
    
    PayrollExist Text3.Text, Text3, Text21
    PayrollExist Text1.Text, Text1, Text22
    PayrollExist Text2.Text, Text2, Text23
    PayrollExist Text10.Text, Text10, Text25
    PayrollExist Text4.Text, Text4, Text26
    PayrollExist Text11.Text, Text11, Text27
    
End Sub
Sub PayrollExist(Ename As String, TXT As TextBox, TXTS As TextBox)
    OpenPBDataBase ("Payrolls")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM Payrolls WHERE dateTrip Like '" & Trim(Combo3.Text) & "' and Ptime Like '" & _
                                      Trim(Combo1.Text) & "' and Particulars Like '" & Trim(Text9.Text) & _
                                      "' and Ecode Like '" & Trim(Ename) & "' and TruckNumber Like '" & Trim(Combo4.Text) & "' ")
    With PRFile
        If Not .EOF Then
            TXT.BackColor = vbYellow
            TXTS.Text = ![Amount]
        Else
            TXT.BackColor = vbWhite
        End If
     .Close
   End With
End Sub

Private Sub Command2_Click()
    GPass.Visible = True
    txtGpass.SetFocus
    'SendKeys "{HOME}+{END}"
End Sub

Private Sub CompanyEX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose(2).Visible = True
End Sub

Private Sub Form_Load()
    Call LOADPersonelsTypes
    Call LOADOrigin
    Call LOADComboOrigin
    Call LOADTRuckWHeel
    Call LOADdateTime
    Call LOadCOVERDATE
    Call loadCustomers
    Call LOADPlateNumbers
    Combo3.Text = Format(Date, "MM/DD/YYYY")
    Combo1.Text = ""
    Call CandyButton2_Click
End Sub
Private Sub Form_Activate()
    MDIMainForm.ActivateChild Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MDIMainForm.RemoveChild Me.Name
End Sub

Sub loadCustomers()
List1.Clear
    OpenPBDataBase ("CWages")
    With PRFile
        If .RecordCount > 1 Then
            Do Until .EOF
                List1.AddItem ![Customer]
            .MoveNext
            Loop
        End If
        .Close
    End With
End Sub
Sub LOadCOVERDATE()
    On Error Resume Next
    Combo9.Clear
    OpenPBDataBase ("DateCover")
    With PRFile
      .MoveFirst
        Do While Not .EOF
           Combo9.AddItem ![CoveredDate]
           If ![Status] = 1 Then
                Combo9.Text = ![CoveredDate]
           End If
           .MoveNext
        Loop
   End With
End Sub
Sub LOADdateTime()
On Error GoTo errt
 Combo3.Clear
 Combo1.Text = ""
    OpenPBDataBase ("TripInfo")
    With PRFile
       .MoveFirst
        Do While Not .EOF
           Combo3.AddItem ![Tripdate]
           .MoveNext
        Loop
        .Close
    End With
errt:
End Sub
Sub LOADTRuckWHeel()
    OpenPBDataBase ("TruckWheel")
    With PRFile
       .MoveFirst
        Do While Not .EOF
           Combo8.AddItem ![TruckTypes]
           .MoveNext
        Loop
        .Close
    End With
End Sub
Sub LOADComboOrigin()
    OpenPBDataBase ("PointOrigin")
    With PRFile
       .MoveFirst
        Do While Not .EOF
           Combo7.AddItem ![POriginName]
           .MoveNext
        Loop
        .Close
    End With
End Sub
Sub LOADPersonelsTypes()
 On Error Resume Next
    OpenPBDataBase ("Personels")
    With PRFile
      .MoveFirst
        Do While Not .EOF
            Combo2.AddItem ![PersonelsType]
            .MoveNext
        Loop
      .Close
    End With
End Sub
Sub LOADOrigin()
On Error Resume Next
Combo5.Clear
    OpenPBDataBase ("Origin")
    With PRFile
      .MoveFirst
        Do While Not .EOF
            Combo5.AddItem ![Originname]
            .MoveNext
        Loop
      .Close
    End With
End Sub
Sub LOADPlateNumbers()
On Error Resume Next
Combo4.Clear
    OpenPBDataBase ("TruckPersonel")
    With PRFile
      .MoveFirst
        Do While Not .EOF
            Combo4.AddItem ![PlateNumber]
            .MoveNext
        Loop
      .Close
    End With
End Sub

Sub SAVEPayroll(TXTName As String, TXTSalary As String)
    'open for payroll personnels

   'Save to payroll
   OpenPBDataBase ("Payrolls")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM Payrolls WHERE dateTrip Like '" & Trim(Combo3.Text) & "' and Ptime Like '" & _
                                      Trim(Combo1.Text) & "' and CoverDate Like '" & Trim(Combo9.Text) & _
                                      "' and Ecode Like '" & Trim(TXTName) & "' ")
    With PRFile
        If Not .EOF Then
            'Exit Sub
            'MsgBox "Meron"
        Else
            .AddNew
                ![ECOde] = Trim(TXTName)
                ![Amount] = Trim(TXTSalary)
                ![DateTrip] = Trim(Combo3.Text)
                ![truckNumber] = Trim(Combo4.Text)
                ![coverdate] = Trim(Combo9.Text)
                ![TPO] = Trim(Combo5.Text)
                ![Particulars] = Trim(Text9.Text)
                ![Cases] = Trim(Text12.Text)
                ![Ptime] = Trim(Combo1.Text)
                ![Status] = "1"
            .Update
            'MsgBox "wala"
        End If
        .Close
    End With
End Sub
Sub SAVEPayrolls()
Dim HelpersName As String
Dim a As Integer
        
        Call SAVEPayroll(Trim(Text3.Text), Trim(Text21.Text))
    
    If Text1.Text <> "" Then
        Call SAVEPayroll(Trim(Text1.Text), Trim(Text22.Text))
    End If
    
    If Text2.Text <> "" Then
        Call SAVEPayroll(Trim(Text2.Text), Trim(Text23.Text))
    End If
    
    If Text10.Text <> "" Then
        Call SAVEPayroll(Trim(Text10.Text), Trim(Text25.Text))
    End If
    
    If Text4.Text <> "" Then
        Call SAVEPayroll(Trim(Text4.Text), Trim(Text26.Text))
    End If
    
    If Text11.Text <> "" Then
        Call SAVEPayroll(Trim(Text11.Text), Trim(Text27.Text))
    End If
End Sub
Sub SAVETrucExpense()
    OpenPBDataBase ("TruckTripExpense")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM TruckTripExpense WHERE tdate Like '" & Trim(Combo3.Text) & _
                                      "' and PlateNumber Like '" & Trim(Combo4.Text) & "' and ttime Like '" & _
                                      Trim(Combo1.Text) & "' ")
    With PRFile
        If Not .EOF Then
            .Edit
                ![GasOil] = Trim(Text15.Text)
                ![NealAllow] = Trim(Text16.Text)
                ![ToolFee] = Trim(Text17.Text)
                ![Xerox] = Trim(Text18.Text)
                ![Parking] = Trim(Text19.Text)
                ![Charges] = Trim(Text20.Text)
                ![CInfo] = Trim(Text28.Text)
                ![TA] = Trim(Text13.Text)
                ![TC] = Val(Text5.Text)
            .Update
        Else
            .AddNew
                ![Tdate] = Trim(Combo3.Text)
                ![PlateNumber] = Trim(Combo4.Text)
                ![TTime] = Trim(Combo1.Text)
                ![GasOil] = Trim(Text15.Text)
                ![NealAllow] = Trim(Text16.Text)
                ![ToolFee] = Trim(Text17.Text)
                ![Xerox] = Trim(Text18.Text)
                ![Parking] = Trim(Text19.Text)
                ![Charges] = Trim(Text20.Text)
                ![CInfo] = Trim(Text28.Text)
                ![TA] = Trim(Text13.Text)
                ![TC] = Val(Text5.Text)
            .Update
        End If
        .Close
    End With
End Sub
Sub OPENTrucExpense()
On Error Resume Next
    OpenPBDataBase ("TruckTripExpense")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM TruckTripExpense WHERE  tdate  Like '" & Trim(Combo3.Text) & _
                                      "' and PlateNumber Like '" & Trim(Combo4.Text) & "' and ttime Like '" & _
                                      Trim(Combo1.Text) & "' ")
    With PRFile
        If Not .EOF Then
                Text15.Text = ""
                Text16.Text = ""
                Text17.Text = ""
                Text18.Text = ""
                Text19.Text = ""
                Text20.Text = ""
                Text28.Text = ""
                
                Text15.Text = ![GasOil]
                Text16.Text = ![NealAllow]
                Text17.Text = ![ToolFee]
                Text18.Text = ![Xerox]
                Text19.Text = ![Parking]
                Text20.Text = ![Charges]
                Text28.Text = ![CInfo]
                '![TA] = Trim(Text13.Text)
                '![TC] = Trim(Text5.Text)
                Label12.Caption = Format(Val(![GasOil]) + Val(![NealAllow]) + Val(![ToolFee]) + Val(![Xerox]) + Val(![Parking]) + Val(![Charges]))
        Else
                Text15.Text = ""
                Text16.Text = ""
                Text17.Text = ""
                Text18.Text = ""
                Text19.Text = ""
                Text20.Text = ""
                Text28.Text = ""
        End If
        .Close
    End With
End Sub
Sub saveTripsInfo()
    OpenPBDataBase ("TripInfo")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM TripInfo WHERE TripDate Like '" & Combo3.Text & "' and TruckNumber Like '" & Combo4.Text & "' and Timetime Like '" & Combo1.Text & "' and PersonelTypes Like '" & Combo2.Text & "' ")
    With PRFile
        If Not .EOF Then
            .Edit
                ![tripamount] = Trim(Text13.Text)
                ![coverdate] = Trim(Combo9.Text)
                ![TPO] = Trim(Combo5.Text)
                ![Cases] = Trim(Text12.Text)
                ![Particulars] = Trim(Text9.Text)
                ![Status] = Trim(Combo10.Text)
                ![WT] = Trim(Combo7.Text)
                ![TType] = Trim(Combo8.Text)
                ![DS] = Trim(Text7.Text)
                ![HS] = Val(Val(Text8.Text) * HelperT)
                ![tripconsume] = Trim(Text24.Text)
                ![ECOde] = Trim(labelGpass.Caption)
                ![Return] = "0"
                ![Driver] = Trim(Text3.Text)
                ![H1] = Trim(Text1.Text)
                ![H2] = Trim(Text2.Text)
                ![H3] = Trim(Text10.Text)
                ![H4] = Trim(Text4.Text)
                ![H5] = Trim(Text11.Text)
                TRStatus.Caption = "Trip Returned"
            .Update
        Else
            .AddNew
                ![PersonelTypes] = Trim(Combo2.Text)
                ![Tripdate] = Trim(Combo3.Text)
                ![truckNumber] = Trim(Combo4.Text)
                ![TimeTime] = Trim(Combo1.Text)
                
                ![tripamount] = Trim(Text13.Text)
                ![coverdate] = Trim(Combo9.Text)
                ![TPO] = Trim(Combo5.Text)
                ![Cases] = Trim(Text12.Text)
                ![Particulars] = Trim(Text9.Text)
                ![Status] = Trim(Combo10.Text)
                ![WT] = Trim(Combo7.Text)
                ![TType] = Trim(Combo8.Text)
                ![DS] = Trim(Text7.Text)
                ![HS] = Val(Val(Text8.Text) * HelperT)
                ![tripconsume] = Trim(Text24.Text)
                ![ECOde] = Trim(labelGpass.Caption)
                ![Return] = "0"
                ![Driver] = Trim(Text3.Text)
                ![H1] = Trim(Text1.Text)
                ![H2] = Trim(Text2.Text)
                ![H3] = Trim(Text10.Text)
                ![H4] = Trim(Text4.Text)
                ![H5] = Trim(Text11.Text)
                
                TRStatus.Visible = True
                TRStatus.Caption = "Trip Returned"
            .Update
        End If
        .Close
    End With
End Sub

Sub SearchTRIPexist()
On Error Resume Next
If Combo1.Text <> "" And Combo4.Text <> "" And Combo9.Text <> "" And Combo3.Text <> "" Then
    OpenPBDataBase ("TripInfo")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM TripInfo WHERE TripDate Like '" & Trim(Combo3.Text) & _
                                      "' and TruckNumber Like '" & Trim(Combo4.Text) & "' and Timetime Like '" & _
                                      Trim(Combo1.Text) & "' and PersonelTypes Like '" & Trim(Combo2.Text) & "' ")
    With PRFile
        If Not .EOF Then
                 
                    Text13.Text = ![tripamount]
                    txtGpass.Text = ![ECOde]
                    Combo5.Text = ![TPO]
                    Text12.Text = ![Cases]
                    Text9.Text = ![Particulars]
                    Combo10.Text = ![Status]
                    Combo7.Text = ![WT]
                    Combo8.Text = ![TType]
                    Text7.Text = ![DS]
                    Text8.Text = ![HS]
                    Text3.Text = ![Driver]
                    Text1.Text = ![H1]
                    Text2.Text = ![H2]
                    Text10.Text = ![H3]
                    Text4.Text = ![H4]
                    Text11.Text = ![H5]
                    Call Command1_Click
                 If ![Return] = 1 Then
                    'CandyButton1.Enabled = True
                    'CandyButton2.Enabled = True
                    TRStatus.Visible = True
                    TRStatus.Caption = "Return Trip"
                    'Call CandyButton1_Click
                 Else
                    Text24.Text = ![tripconsume]
                    Text5.Text = Round(Val(Text13.Text) - Val(Text24.Text), 2)
                    'CandyButton1.Enabled = True
                    'CandyButton2.Enabled = True
                    TRStatus.Caption = "Trip Returned"
                    TRStatus.Visible = True
                    
                 End If
        Else
                 Text21.Text = ""
                 Text22.Text = ""
                 Text23.Text = ""
                 Text25.Text = ""
                 Text26.Text = ""
                 Text13.Text = ""
                 Text6.Text = ""
                 Combo5.Text = ""
                 Text12.Text = ""
                 Text9.Text = ""
                 Combo10.Text = ""
                 Combo7.Text = ""
                 Text7.Text = ""
                 Text8.Text = ""
                 Text24.Text = ""
                 Text5.Text = ""
                 'Combo7.SetFocus
                 Text28.Text = ""
                 Text3.BackColor = &HE0E0E0
                Text1.BackColor = &HE0E0E0
                Text2.BackColor = &HE0E0E0
                Text10.BackColor = &HE0E0E0
                Text4.BackColor = &HE0E0E0
                Text11.BackColor = &HE0E0E0
                 TRStatus.Visible = False
                 labelGpass.Caption = ""
        End If
        .Close
    End With
Else
    MsgBox "Invalid Null.. Please check the date..", vbCritical, "Error"
End If
End Sub
Sub SearchTRIPexistGatePass()
'On Error Resume Next
        
   If Option2.Value = True Then
            OpenPBDataBase ("TripInfo")
            Set PRFile = PDbase.OpenRecordset("SELECT * FROM TripInfo WHERE Ecode Like '" & Trim(txtGpass.Text) & "' ")
            With PRFile
                If Not .EOF Then
                            Text13.Text = ![tripamount]
                            txtGpass.Text = ![ECOde]
                            Combo5.Text = ![TPO]
                            Combo4.Text = ![truckNumber]
                            Text12.Text = ![Cases]
                            Text9.Text = ![Particulars]
                            Combo10.Text = ![Status]
                            Combo7.Text = ![WT]
                            Combo8.Text = ![TType]
                            Text7.Text = ![DS]
                            Text8.Text = ![HS]
                            Text3.Text = ![Driver]
                            Text1.Text = ![H1]
                            Text2.Text = ![H2]
                            Text10.Text = ![H3]
                            Text4.Text = ![H4]
                            Text11.Text = ![H5]
                            Call Command1_Click
                         If ![Return] = 1 Then
                            'CandyButton1.Enabled = True
                            'CandyButton2.Enabled = True
                            TRStatus.Visible = True
                            TRStatus.Caption = "Return Trip"
                            'Call CandyButton1_Click
                         Else
                            Text24.Text = ![tripconsume]
                            Text5.Text = Round(Val(Text13.Text) - Val(Text24.Text), 2)
                            'CandyButton1.Enabled = True
                            'CandyButton2.Enabled = True
                            TRStatus.Caption = "Trip Returned"
                            TRStatus.Visible = True
                            
                         End If
                            labelGpass.Caption = Trim(txtGpass.Text)
                Else
                         Text21.Text = ""
                         Text22.Text = ""
                         Text23.Text = ""
                         Text25.Text = ""
                         Text26.Text = ""
                         Text13.Text = ""
                         Text6.Text = ""
                         Combo5.Text = ""
                         Text12.Text = ""
                         Text9.Text = ""
                         Combo10.Text = ""
                         Combo7.Text = ""
                         Text7.Text = ""
                         Text8.Text = ""
                         Text24.Text = ""
                         Text5.Text = ""
                         'Combo7.SetFocus
                         Text28.Text = ""
                        Text3.BackColor = &HE0E0E0
                        Text1.BackColor = &HE0E0E0
                        Text2.BackColor = &HE0E0E0
                        Text10.BackColor = &HE0E0E0
                        Text4.BackColor = &HE0E0E0
                        Text11.BackColor = &HE0E0E0
                        TRStatus.Visible = False
                        labelGpass.Caption = ""
                End If
                .Close
            End With
        
    'if OptionButton2 = True
    Else
          Dim ddate As String
            OpenPBDataBase ("Billings")
            Set PRFile = PDbase.OpenRecordset("SELECT * FROM Billings WHERE Phealth Like '" & Trim(txtGpass.Text) & "' ")
            With PRFile
                If Not .EOF Then
                            Text21.Text = ""
                            Text22.Text = ""
                            Text23.Text = ""
                            Text25.Text = ""
                            Text26.Text = ""
                            Text13.Text = ""
                            Text6.Text = ""
                            Combo5.Text = ""
                            Text12.Text = ""
                            Text9.Text = ""
                            Combo10.Text = ""
                            Combo7.Text = ""
                            Text7.Text = ""
                            Text8.Text = ""
                            Text24.Text = ""
                            Text5.Text = ""
                            'Combo7.SetFocus
                            Text28.Text = ""
                            TRStatus.Visible = False
                            ddate = ![Loans]
                            Text12.Text = ![ECOde]
                            Text9.Text = ![Advances]
                            Combo4.Text = ![SSS]
                            
                            Call Combo4_Click
                            Call Text9_KeyPress(13)
                            
                            labelGpass.Caption = Trim(txtGpass.Text)
                            Combo3.Text = ddate
                            labelGpass.Caption = Trim(txtGpass.Text)
                Else
                         Text21.Text = ""
                         Text22.Text = ""
                         Text23.Text = ""
                         Text25.Text = ""
                         Text26.Text = ""
                         Text13.Text = ""
                         Text6.Text = ""
                         Combo5.Text = ""
                         Text12.Text = ""
                         Text9.Text = ""
                         Combo10.Text = ""
                         Combo7.Text = ""
                         Text7.Text = ""
                         Text8.Text = ""
                         Text24.Text = ""
                         Text5.Text = ""
                         'Combo7.SetFocus
                         Text28.Text = ""
                        Text3.BackColor = &HE0E0E0
                        Text1.BackColor = &HE0E0E0
                        Text2.BackColor = &HE0E0E0
                        Text10.BackColor = &HE0E0E0
                        Text4.BackColor = &HE0E0E0
                        Text11.BackColor = &HE0E0E0
                        TRStatus.Visible = False
                        labelGpass.Caption = ""
                End If
                .Close
            End With
    End If
    
    GPass.Visible = False
    
End Sub

Private Sub GPass_Click()
    GPass.Visible = False
    'txtGpass.Text = "000000"
End Sub

Private Sub Label1_Click()
    Text3.BackColor = &HE0E0E0
    Text1.BackColor = &HE0E0E0
    Text2.BackColor = &HE0E0E0
    Text10.BackColor = &HE0E0E0
    Text4.BackColor = &HE0E0E0
    Text11.BackColor = &HE0E0E0
End Sub

Private Sub Label12_Click()
    'cal1.Visible = True
End Sub

Private Sub Label18_Click()
    GPass.Visible = False
    'txtGpass.Text = "000000"
End Sub

Private Sub List1_Click()
    If bNoClick Then Exit Sub
    Text9.Text = List1.Text
End Sub

Private Sub List1_GotFocus()
    SendKeys "{DOWN}"
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call List1_Click
    If KeyCode = vbKeyEscape Then
        List1.Visible = False
        Text9.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        Text9.Text = List1.List(List1.ListIndex)
        List1.Visible = False
        Call Text9_KeyPress(13)
   End If
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If butoon = 1 Then
        Text9.Text = List1.List(List1.ListIndex)
    End If
End Sub



Private Sub Option1_Click()
    txtGpass.SetFocus
End Sub

Private Sub Option2_Click()
    txtGpass.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call OpenEXIST(Trim(Text1.Text), Text1, Text2)
    End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text13.SetFocus
    End If
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CandyButton4_Click
    End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text16.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub Text16_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text17.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub Text17_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text18.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub Text18_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text19.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub Text19_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text20.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub Text20_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text28.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call OpenEXIST(Trim(Text2.Text), Text2, Text10)
    End If
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call OpenEXIST(Trim(Text10.Text), Text10, Text4)
    End If
End Sub

Private Sub Text28_KeyPress(KeyAscii As Integer)
    If KeyAscii = 16 Then
        tmpOTHER = tmpOTHER & KeyAscii
    End If
End Sub

Private Sub Text3_Click()
    'SendKeys "{HOME}+{END}"
End Sub
Private Sub Text1_Click()
    'SendKeys "{HOME}+{END}"
End Sub
Private Sub Text2_Click()
    'SendKeys "{HOME}+{END}"
End Sub
Private Sub Text10_Click()
    'SendKeys "{HOME}+{END}"
End Sub
Private Sub Text4_Click()
    'SendKeys "{HOME}+{END}"
End Sub
Private Sub Text11_Click()
    'SendKeys "{HOME}+{END}"
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call OpenEXIST(Trim(Text4.Text), Text4, Text11)
    End If
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Call OpenEXIST(Trim(Text11.Text), Text11, Combo3)
    End If
End Sub
Private Sub Text15_Change()
    Label12.Caption = Format(Val(Text15.Text) + Val(Text16.Text) + Val(Text17.Text) + Val(Text18.Text) + Val(Text19.Text) + Val(Text20.Text), "###,###.00")
End Sub

Private Sub Text16_Change()
    Label12.Caption = Format(Val(Text15.Text) + Val(Text16.Text) + Val(Text17.Text) + Val(Text18.Text) + Val(Text19.Text) + Val(Text20.Text), "###,###.00")
End Sub

Private Sub Text17_Change()
    Label12.Caption = Format(Val(Text15.Text) + Val(Text16.Text) + Val(Text17.Text) + Val(Text18.Text) + Val(Text19.Text) + Val(Text20.Text), "###,###.00")
End Sub

Private Sub Text18_Change()
    Label12.Caption = Format(Val(Text15.Text) + Val(Text16.Text) + Val(Text17.Text) + Val(Text18.Text) + Val(Text19.Text) + Val(Text20.Text), "###,###.00")
End Sub

Private Sub Text19_Change()
    Label12.Caption = Format(Val(Text15.Text) + Val(Text16.Text) + Val(Text17.Text) + Val(Text18.Text) + Val(Text19.Text) + Val(Text20.Text), "###,###.00")
End Sub

Private Sub Text20_Change()
    Label12.Caption = Format(Val(Text15.Text) + Val(Text16.Text) + Val(Text17.Text) + Val(Text18.Text) + Val(Text19.Text) + Val(Text20.Text), "###,###.00")
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call OpenEXIST(Trim(Text3.Text), Text3, Text1)
    End If
End Sub
Sub OpenEXIST(Names As String, TXT As TextBox, txtSet As TextBox)
  If TXT.Text <> "" Then
        OpenPBDataBase ("EmployeeInfo")
        Set PRFile = PDbase.OpenRecordset("SELECT * FROM EmployeeInfo WHERE EName Like '" & Trim(Names) & "' ")
        With PRFile
            If Not .EOF Then
                txtSet.SetFocus
                SendKeys "{HOME}+{END}"
            Else
                MsgBox "Employee Not Exist..", vbInformation, "Not Found"
                TXT.SetFocus
                SendKeys "{HOME}+{END}"
            End If
        End With
  Else
    Combo3.SetFocus
    SendKeys "{HOME}+{END}"
  End If


End Sub

Private Sub Text5_Change()
    Text24.Text = Val(Text13.Text) - Val(Text5.Text)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If KeyAscii = 13 Then
        Text9.SetFocus
        'List1.Visible = True
    End If
End If
End Sub

Private Sub Text9_Change()
     Call AutoTXTcomplete(List1, Text9)
End Sub

Private Sub Text9_Click()
    Text9.SelStart = 0
    Text9.SelLength = Len(Text9)
    List1.Visible = True
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        List1.Visible = False
    End If
    If KeyCode = vbKeyDown Then
        List1.Visible = True
        Text9.Text = List1.Text
        List1.SetFocus
    End If
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
        DelKey = True
        Exit Sub
    End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
'1
            If Text1.Text <> "" Then
                HelperT = 1
            End If
            '2
            If Text2.Text <> "" Then
                HelperT = 2
            End If
            '3
            If Text10.Text <> "" Then
                HelperT = 3
            End If
            '4
            If Text4.Text <> "" Then
                HelperT = 4
            End If
            '5
            If Text11.Text <> "" Then
                HelperT = 5
            End If
        
        
        
        
        
        OpenPBDataBase ("CWages")
        Set PRFile = PDbase.OpenRecordset("SELECT * FROM CWages WHERE Customer Like '" & Trim(Text9.Text) & "' ")
        With PRFile
            If Not .EOF Then
                Combo7.Text = Trim(![Wages])
                Text6.Text = Trim(![Source])
                Text9.Text = Text9.Text & "   (" & HelperT & ")"
                Call Combo7_Click
                Text12.SetFocus
            Else
                Text9.Text = Text9.Text & "   (" & HelperT & ")"
                Combo7.SetFocus
            End If
            .Close
        End With
        List1.Visible = False
    End If
End Sub
Private Sub text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
        DelKey = True
        Exit Sub
    End If
    
    If KeyCode = vbKeyF12 Then
        Call SAVEPayroll(Trim(Text3.Text), Trim(Text21.Text))
    End If
End Sub
Private Sub Text3_Change()
    Call AutoTXTcomplete(MDIMainForm.List4, Text3)
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
        DelKey = True
        Exit Sub
    End If
    If KeyCode = vbKeyF12 Then
        Call SAVEPayroll(Trim(Text1.Text), Trim(Text22.Text))
    End If
End Sub
Private Sub Text1_Change()
    Call AutoTXTcomplete(MDIMainForm.List3, Text1)
End Sub
Private Sub text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
        DelKey = True
        Exit Sub
    End If
    If KeyCode = vbKeyF12 Then
        Call SAVEPayroll(Trim(Text2.Text), Trim(Text23.Text))
    End If
End Sub
Private Sub Text2_Change()
    Call AutoTXTcomplete(MDIMainForm.List3, Text2)
End Sub
Private Sub text10_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
        DelKey = True
        Exit Sub
    End If
    If KeyCode = vbKeyF12 Then
        Call SAVEPayroll(Trim(Text10.Text), Trim(Text25.Text))
    End If
End Sub
Private Sub text10_Change()
    Call AutoTXTcomplete(MDIMainForm.List3, Text10)
End Sub
Private Sub text4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
        DelKey = True
        Exit Sub
    End If
    If KeyCode = vbKeyF12 Then
        Call SAVEPayroll(Trim(Text4.Text), Trim(Text26.Text))
    End If
End Sub
Private Sub Text4_Change()
    Call AutoTXTcomplete(MDIMainForm.List3, Text4)
End Sub
Private Sub text11_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
        DelKey = True
        Exit Sub
    End If
    If KeyCode = vbKeyF12 Then
        Call SAVEPayroll(Trim(Text11.Text), Trim(Text27.Text))
    End If
End Sub
Private Sub Text11_Change()
    Call AutoTXTcomplete(MDIMainForm.List3, Text11)
End Sub
''==================================================================================================
Private Sub Text1_GotFocus()
    Text1.BackColor = &HC0FFFF
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &HE0E0E0
End Sub
Private Sub Text2_GotFocus()
    Text2.BackColor = &HC0FFFF
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &HE0E0E0
End Sub
Private Sub Text3_GotFocus()
    Text3.BackColor = &HC0FFFF
End Sub
Private Sub Text3_LostFocus()
    Text3.BackColor = &HE0E0E0
End Sub
Private Sub Text4_GotFocus()
    Text4.BackColor = &HC0FFFF
End Sub
Private Sub Text4_LostFocus()
    Text4.BackColor = &HE0E0E0
End Sub
Private Sub Text7_GotFocus()
    Text7.BackColor = &HC0FFFF
End Sub
Private Sub Text7_LostFocus()
    Text7.BackColor = &HE0E0E0
End Sub
Private Sub Text8_GotFocus()
    Text8.BackColor = &HC0FFFF
End Sub
Private Sub Text8_LostFocus()
    Text8.BackColor = &HE0E0E0
End Sub
Private Sub Text9_GotFocus()
    Text9.BackColor = &HC0FFFF
    List1.Visible = True
End Sub
Private Sub Text9_LostFocus()
    Text9.BackColor = &HE0E0E0
    'List1.Visible = False
End Sub
Private Sub Text5_GotFocus()
    Text5.BackColor = &HC0FFFF
End Sub
Private Sub Text5_LostFocus()
    Text5.BackColor = &HE0E0E0
End Sub
Private Sub Text6_GotFocus()
    Text6.BackColor = &HC0FFFF
End Sub
Private Sub Text6_LostFocus()
    Text6.BackColor = &HE0E0E0
End Sub
Private Sub Text12_GotFocus()
    Text12.BackColor = &HC0FFFF
End Sub
Private Sub Text12_LostFocus()
    Text12.BackColor = &HE0E0E0
End Sub
Private Sub Text13_GotFocus()
    Text13.BackColor = &HC0FFFF
End Sub
Private Sub Text13_LostFocus()
    Text13.BackColor = &HE0E0E0
End Sub
Private Sub Text15_GotFocus()
    Text15.BackColor = &HC0FFFF
End Sub
Private Sub Text15_LostFocus()
    Text15.BackColor = &HE0E0E0
End Sub
Private Sub Text16_GotFocus()
    Text16.BackColor = &HC0FFFF
End Sub
Private Sub Text16_LostFocus()
    Text16.BackColor = &HE0E0E0
End Sub
Private Sub Text17_GotFocus()
    Text17.BackColor = &HC0FFFF
End Sub
Private Sub Text17_LostFocus()
    Text17.BackColor = &HE0E0E0
End Sub
Private Sub Text18_GotFocus()
    Text18.BackColor = &HC0FFFF
End Sub
Private Sub Text18_LostFocus()
    Text18.BackColor = &HE0E0E0
End Sub
Private Sub Text19_GotFocus()
    Text19.BackColor = &HC0FFFF
End Sub
Private Sub Text19_LostFocus()
    Text19.BackColor = &HE0E0E0
End Sub
Private Sub Text20_GotFocus()
    Text20.BackColor = &HC0FFFF
End Sub
Private Sub Text20_LostFocus()
    Text20.BackColor = &HE0E0E0
End Sub
Private Sub Text10_GotFocus()
    Text10.BackColor = &HC0FFFF
End Sub
Private Sub Text10_LostFocus()
    Text10.BackColor = &HE0E0E0
End Sub
Private Sub Text11_GotFocus()
    Text11.BackColor = &HC0FFFF
End Sub
Private Sub Text11_LostFocus()
    Text11.BackColor = &HE0E0E0
End Sub
Private Sub Text24_GotFocus()
    Text24.BackColor = &HC0FFFF
End Sub
Private Sub Text24_LostFocus()
    Text24.BackColor = &HE0E0E0
End Sub
Private Sub Combo1_GotFocus()
    Combo1.BackColor = &HC0FFFF
End Sub
Private Sub Combo1_LostFocus()
    Combo1.BackColor = &HE0E0E0
End Sub
Private Sub Combo2_GotFocus()
    Combo2.BackColor = &HC0FFFF
End Sub
Private Sub Combo2_LostFocus()
    Combo2.BackColor = &HE0E0E0
End Sub
Private Sub Combo3_GotFocus()
    Combo3.BackColor = &HC0FFFF
End Sub
Private Sub Combo3_LostFocus()
    Combo3.BackColor = &HE0E0E0
End Sub
Private Sub Combo4_GotFocus()
    Combo4.BackColor = &HC0FFFF
End Sub
Private Sub Combo4_LostFocus()
    Combo4.BackColor = &HE0E0E0
End Sub
Private Sub Combo5_GotFocus()
    Combo5.BackColor = &HC0FFFF
End Sub
Private Sub Combo5_LostFocus()
    Combo5.BackColor = &HE0E0E0
End Sub
Private Sub combo7_GotFocus()
    Combo7.BackColor = &HC0FFFF
End Sub
Private Sub combo7_LostFocus()
    Combo7.BackColor = &HE0E0E0
End Sub
Private Sub combo8_GotFocus()
    Combo8.BackColor = &HC0FFFF
End Sub
Private Sub combo8_LostFocus()
    Combo8.BackColor = &HE0E0E0
End Sub
Private Sub combo9_GotFocus()
    Combo9.BackColor = &HC0FFFF
End Sub
Private Sub combo9_LostFocus()
    Combo9.BackColor = &HE0E0E0
End Sub
Private Sub combo10_GotFocus()
    Combo10.BackColor = &HC0FFFF
End Sub
Private Sub combo10_LostFocus()
    Combo10.BackColor = &HE0E0E0
End Sub
Private Sub imgClose_Click(Index As Integer)
    Select Case Index
            Case 0
                sts.Visible = False
            Case 1
                sts.Visible = False
            Case 2
                CompanyEX.Visible = False
            Case 3
                CompanyEX.Visible = False
    End Select
End Sub

Private Sub imgClose_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
            Case 0
                imgClose(0).Visible = True
            Case 1
                imgClose(1).Visible = False
            Case 2
                imgClose(2).Visible = False
            Case 3
                imgClose(3).Visible = True
    End Select
End Sub

Private Sub sts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose(1).Visible = True
End Sub

Private Sub txtGpass_Change()
    labelGpass.Caption = Trim(txtGpass.Text)
End Sub

Private Sub txtGpass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call SearchTRIPexistGatePass
    End If
End Sub
