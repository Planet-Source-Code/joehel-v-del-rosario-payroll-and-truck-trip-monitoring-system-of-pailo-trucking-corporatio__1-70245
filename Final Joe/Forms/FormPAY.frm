VERSION 5.00
Begin VB.Form FormPAY 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Office Payroll"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   12375
   Begin VB.ComboBox Combo6 
      Appearance      =   0  'Flat
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
      ItemData        =   "FormPAY.frx":0000
      Left            =   1845
      List            =   "FormPAY.frx":0002
      TabIndex        =   53
      Top             =   10410
      Width           =   2985
   End
   Begin VB.TextBox Text17 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9975
      TabIndex        =   4
      Top             =   1995
      Width           =   1680
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9975
      TabIndex        =   3
      Top             =   1620
      Width           =   1680
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9975
      TabIndex        =   2
      Top             =   1230
      Width           =   1680
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2400
      TabIndex        =   48
      Top             =   4590
      Width           =   1680
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2370
      TabIndex        =   17
      Top             =   6750
      Width           =   1680
   End
   Begin VB.TextBox Text14 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2370
      TabIndex        =   18
      Top             =   7140
      Width           =   1680
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2370
      TabIndex        =   16
      Top             =   6390
      Width           =   1680
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2370
      TabIndex        =   15
      Top             =   6030
      Width           =   1680
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2370
      TabIndex        =   14
      Top             =   5655
      Width           =   1680
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9840
      TabIndex        =   34
      Top             =   4125
      Width           =   1680
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9825
      TabIndex        =   12
      Top             =   3720
      Width           =   1680
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6135
      TabIndex        =   31
      Top             =   4125
      Width           =   1680
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2370
      TabIndex        =   13
      Top             =   5250
      Width           =   1680
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6120
      TabIndex        =   11
      Top             =   3720
      Width           =   1680
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2400
      TabIndex        =   27
      Top             =   4140
      Width           =   1680
   End
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
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
      ItemData        =   "FormPAY.frx":0004
      Left            =   6975
      List            =   "FormPAY.frx":0029
      TabIndex        =   8
      Top             =   3105
      Width           =   1140
   End
   Begin VB.ComboBox Combo5 
      Appearance      =   0  'Flat
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
      ItemData        =   "FormPAY.frx":0065
      Left            =   4905
      List            =   "FormPAY.frx":00C6
      TabIndex        =   7
      Top             =   3105
      Width           =   795
   End
   Begin VB.ComboBox Combo3 
      Appearance      =   0  'Flat
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
      ItemData        =   "FormPAY.frx":013D
      Left            =   3315
      List            =   "FormPAY.frx":019E
      TabIndex        =   6
      Top             =   3105
      Width           =   795
   End
   Begin VB.ComboBox Combo4 
      Appearance      =   0  'Flat
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
      ItemData        =   "FormPAY.frx":0215
      Left            =   9345
      List            =   "FormPAY.frx":0217
      TabIndex        =   9
      Top             =   3120
      Width           =   1380
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2400
      TabIndex        =   10
      Top             =   3720
      Width           =   1680
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
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
      Left            =   1500
      TabIndex        =   5
      Top             =   2610
      Width           =   6255
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
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
      ItemData        =   "FormPAY.frx":0219
      Left            =   735
      List            =   "FormPAY.frx":021B
      TabIndex        =   1
      Text            =   "822 MOVERS (Pailo)"
      Top             =   1095
      Width           =   5700
   End
   Begin MOVERS.CandyButton ButPrev 
      Height          =   435
      Left            =   870
      TabIndex        =   55
      Top             =   7740
      Width           =   1425
      _ExtentX        =   2514
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
      Caption         =   "    Preview"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "FormPAY.frx":021D
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
      Height          =   420
      Left            =   2550
      TabIndex        =   56
      Top             =   7755
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "   &Save"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "FormPAY.frx":0997
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
      TabIndex        =   57
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   661
      Caption         =   "Office Personnels Payroll"
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
   Begin VB.Image Image2 
      Height          =   435
      Left            =   -120
      Picture         =   "FormPAY.frx":1111
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   12360
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Covered :"
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
      Left            =   390
      TabIndex        =   54
      Top             =   9885
      Width           =   1950
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Holiday Per Day:"
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
      Left            =   6975
      TabIndex        =   52
      Top             =   2025
      Width           =   2940
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rate OT Per Hour:"
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
      Left            =   6900
      TabIndex        =   51
      Top             =   1635
      Width           =   3015
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Per 15th Day:"
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
      Left            =   7020
      TabIndex        =   50
      Top             =   1260
      Width           =   2895
   End
   Begin VB.Shape Shape5 
      Height          =   1260
      Left            =   7755
      Top             =   1125
      Width           =   4065
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Cola :"
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
      Left            =   885
      TabIndex        =   49
      Top             =   4590
      Width           =   1365
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Advances ::"
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
      Left            =   855
      TabIndex        =   47
      Top             =   6735
      Width           =   1425
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MM/DD/YYY"
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
      Left            =   8370
      TabIndex        =   46
      Top             =   2685
      Width           =   3330
   End
   Begin VB.Label nt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7875
      TabIndex        =   45
      Top             =   6720
      Width           =   2880
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Net Pay :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5280
      TabIndex        =   44
      Top             =   6675
      Width           =   2370
   End
   Begin VB.Label td 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7875
      TabIndex        =   43
      Top             =   5985
      Width           =   2880
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Deductions:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5265
      TabIndex        =   42
      Top             =   5940
      Width           =   2670
   End
   Begin VB.Label ts 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7875
      TabIndex        =   41
      Top             =   5490
      Width           =   2880
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Salary :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5280
      TabIndex        =   40
      Top             =   5445
      Width           =   1890
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00808080&
      Height          =   2550
      Left            =   4800
      Top             =   5070
      Width           =   6990
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Others :"
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
      Left            =   855
      TabIndex        =   39
      Top             =   7140
      Width           =   1425
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Loans :"
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
      Left            =   855
      TabIndex        =   38
      Top             =   6375
      Width           =   1425
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Medicare :"
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
      Left            =   855
      TabIndex        =   37
      Top             =   6015
      Width           =   1425
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "SSS Premium:"
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
      Left            =   840
      TabIndex        =   36
      Top             =   5640
      Width           =   1425
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      Height          =   2550
      Left            =   705
      Top             =   5070
      Width           =   11085
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   540
      Left            =   705
      Top             =   3015
      Width           =   11085
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   5325
      Left            =   705
      Top             =   3030
      Width           =   11085
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Holidays Pay:"
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
      Left            =   8325
      TabIndex        =   35
      Top             =   4125
      Width           =   1365
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Holidays :"
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
      Left            =   8685
      TabIndex        =   33
      Top             =   3735
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Overtime Pay:"
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
      Left            =   4605
      TabIndex        =   32
      Top             =   4125
      Width           =   1365
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Witholding Tax :"
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
      Left            =   825
      TabIndex        =   30
      Top             =   5235
      Width           =   1425
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Overtime Hours:"
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
      Left            =   4545
      TabIndex        =   29
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Days Pay"
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
      Left            =   885
      TabIndex        =   28
      Top             =   4155
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Days Work:"
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
      Left            =   870
      TabIndex        =   26
      Top             =   3720
      Width           =   1200
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Month:"
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
      Left            =   6225
      TabIndex        =   25
      Top             =   3165
      Width           =   1155
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
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
      Left            =   8610
      TabIndex        =   24
      Top             =   3165
      Width           =   1155
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
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
      Left            =   4470
      TabIndex        =   23
      Top             =   3150
      Width           =   1155
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
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
      Left            =   2715
      TabIndex        =   22
      Top             =   3135
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Period:"
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
      Left            =   870
      TabIndex        =   21
      Top             =   3120
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Left            =   735
      TabIndex        =   20
      Top             =   2625
      Width           =   945
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "San Ildefonso Aluminos, Laguna"
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
      Left            =   735
      TabIndex        =   19
      Top             =   1500
      Width           =   5685
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "SALARY OF EMPLOYEE"
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
      Left            =   720
      TabIndex        =   0
      Top             =   1755
      Width           =   2370
   End
End
Attribute VB_Name = "FormPAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><>'
'Programmed By: Joehel V. del Rosario    '
'822 MOVERS(PAILO) Trucking Corporation  '
'Date: June 1, 2007                      '
'<><><><><><><><><><><><><><><><><><><><>'



Private Sub ButPrev_Click()
    FormPayPrint.po.Caption = Combo1.Text
    FormPayPrint.n.Caption = Text1.Text
    FormPayPrint.d.Caption = Label26.Caption
    FormPayPrint.cd.Caption = "Coverage date of Salary from      " & Combo3.Text & " - " & Combo2.Text & "    to   " & Combo5.Text & " - " & Combo2.Text & " - - - - - " & Combo4.Text
    FormPayPrint.nd.Caption = Text2.Text
    FormPayPrint.no.Caption = Text6.Text
    FormPayPrint.nh.Caption = Text9.Text
    FormPayPrint.dp.Caption = Format(Text5.Text, "###,###.00")
    FormPayPrint.op.Caption = Format(Text8.Text, "###,###.00")
    FormPayPrint.hp.Caption = Format(Text10.Text, "###,###.00")
    FormPayPrint.ec.Caption = Format(Text4.Text, "###,###.00")
    FormPayPrint.wh.Caption = Format(Text7.Text, "###,###.00")
    FormPayPrint.ss.Caption = Format(Text11.Text, "###,###.00")
    FormPayPrint.md.Caption = Format(Text12.Text, "###,###.00")
    FormPayPrint.lo.Caption = Format(Text13.Text, "###,###.00")
    FormPayPrint.ad.Caption = Format(Text3.Text, "###,###.00")
    FormPayPrint.ors.Caption = Format(Text14.Text, "###,###.00")
    
    FormPayPrint.ts.Caption = ts.Caption
    FormPayPrint.td.Caption = td.Caption
    FormPayPrint.nt.Caption = nt.Caption

    FormPayPrint.Show 1
End Sub

Sub SAVEPayroll(TXTName As String)
    'open for payroll personnels
    On Error Resume Next
    OpenPBDataBase ("PayrollPersonels")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM PayrollPersonels WHERE Names Like '" & Trim(TXTName) & "' ")
    With PRFile
        If Not .EOF Then
            'do nothing
        Else
            .AddNew
                ![Names] = Trim(TXTName)
            .Update
        End If
        .Close
    End With



OpenPBDataBase ("Payrolls")
   'Save to payroll
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM Payrolls WHERE CoverDate Like '" & Trim(Combo6.Text) & "' and Ecode Like '" & Trim(TXTName) & "' ")
    With PRFile
        If Not .EOF Then
            Exit Sub
        Else
            .AddNew
                ![ECOde] = Trim(TXTName)
                ![Amount] = Val(Val(Text5.Text) + Val(Text8.Text) + Val(Text10.Text)) - Val(Val(Text7.Text) + Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text) + Val(Text14.Text) + Val(Text3.Text))
                ![coverdate] = Trim(Combo6.Text)
                ![Status] = "1"
            .Update
        End If
        .Close
    End With
    MsgBox TXTName & "---" & TXTSalary
    
    'Save Deductions
        OpenPBDataBase ("DeductionsInfo")
        Set PRFile = PDbase.OpenRecordset("SELECT * FROM DeductionsInfo WHERE DName Like '" & Trim(TXTName) & "' and  DateCover Like '" & Trim(Combo6.Text) & "' ")
        With PRFile
            If Not .EOF Then
                .Edit
                    ![DAmount] = ![DAmount] + Val(Text8.Text)
                .Update
            Else
                .AddNew
                    ![DName] = Trim(TXTName)
                    ![DAmount] = Val(Text7.Text) + Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text) + Val(Text14.Text) + Val(Text3.Text)
                    ![Status] = "1"
                    ![DateCover] = Trim(Combo6.Text)
                .Update
            End If
            .Close
        End With
End Sub

Private Sub CandyButton1_Click()
If Text1.Text <> "" Then
    SAVEPayroll Text1.Text
End If

End Sub

Private Sub Form_Activate()
    MDIMainForm.ActivateChild Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MDIMainForm.RemoveChild Me.Name
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Call LOadcombo3
    Call LOADPersonelsTypes

    For a = 2006 To 2050
        Combo4.AddItem a
    Next a
    Label26.Caption = Format(Date, "MMMM DD, YYYY")
End Sub
Sub LOadcombo3()
    On Error Resume Next
    OpenPBDataBase ("DateCover")
    With PRFile
      .MoveFirst
        Do While Not .EOF
            Combo6.AddItem ![CoveredDate]
            
            If ![Status] = "1" Then
                Combo6.Text = ![CoveredDate]
            End If
            .MoveNext
        Loop
   End With

End Sub

Sub LOADPersonelsTypes()
 On Error Resume Next
    OpenPBDataBase ("Personels")
    With PRFile
      .MoveFirst
        Do While Not .EOF
            Combo1.AddItem ![PersonelsType]
            .MoveNext
        Loop
      .Close
    End With
End Sub

Private Sub Text1_Change()
    Call AutoTXTcomplete(MDIMainForm.List1, Text1)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
        DelKey = True
        Exit Sub
    End If
End Sub

Private Sub Text11_Change()
Call TotalSALDEC
End Sub

Private Sub Text12_Change()
Call TotalSALDEC
End Sub

Private Sub Text13_Change()
Call TotalSALDEC
End Sub

Private Sub Text14_Change()
    Call TotalSALDEC
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text6.SetFocus
        Call TotalSALDEC
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub Text3_Change()
Call TotalSALDEC
End Sub

Private Sub Text4_Change()
    Call TotalSALDEC
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text9.SetFocus
        Call TotalSALDEC
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub Text7_Change()
    Call TotalSALDEC
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text4.SetFocus
        Call TotalSALDEC
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text7.SetFocus
        Call TotalSALDEC
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text11.SetFocus
        Call TotalSALDEC
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text12.SetFocus
        Call TotalSALDEC
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text13.SetFocus
        Call TotalSALDEC
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub Text13_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text3.SetFocus
        Call TotalSALDEC
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text14.SetFocus
        Call TotalSALDEC
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub Text6_Change()
    Text8.Text = Round(Val(Val(Text6.Text) * Val(Text16.Text)), 0)
    Call TotalSALDEC
End Sub
Private Sub Text9_Change()
    Text10.Text = Val(Val(Text9.Text) * Val(Text17.Text)) * 2
    Call TotalSALDEC
End Sub
Private Sub Text2_Change()
    Text5.Text = Round(Val(Val(Text2.Text) * Val(Val(Text15.Text) / 15)), 0)
    Call TotalSALDEC
End Sub
Sub TotalSALDEC()
    ts = Format(Round(Val(Text5.Text) + Val(Text8.Text) + Val(Text10.Text) + Val(Text4.Text), 2), "###,###.00")
    td = Format(Round(Val(Text7.Text) + Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text) + Val(Text14.Text) + Val(Text3.Text), 2), "###,###.00")
    nt = Format(Round(Val(Val(Text5.Text) + Val(Text8.Text) + Val(Text10.Text) + Val(Text4.Text)) - Val(Val(Text7.Text) + Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text) + Val(Text14.Text) + Val(Text3.Text)), 2), "###,###.00")
End Sub
Sub hkdf()
OpenPBDataBase ("Payrolls")
   'Save to payroll
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM Payrolls WHERE dateTrip Like '" & Trim(Combo3.Text) & "' and Ptime Like '" & _
                                      Trim(Combo1.Text) & "' and CoverDate Like '" & Trim(Combo9.Text) & _
                                      "' and Ecode Like '" & Trim(TXTName) & "' ")
    With PRFile
        If Not .EOF Then
            Exit Sub
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
        End If
        .Close
    End With

End Sub
''''=============================================
Private Sub Text11_LostFocus()
    Text11.BackColor = &HFFFFFF
End Sub
Private Sub Text11_GotFocus()
    Text11.BackColor = &HC0FFFF
End Sub
Private Sub Text12_LostFocus()
    Text12.BackColor = &HFFFFFF
End Sub
Private Sub Text12_GotFocus()
    Text12.BackColor = &HC0FFFF
End Sub
Private Sub Text13_LostFocus()
    Text13.BackColor = &HFFFFFF
End Sub
Private Sub Text13_GotFocus()
    Text13.BackColor = &HC0FFFF
End Sub
Private Sub Text14_LostFocus()
    Text14.BackColor = &HFFFFFF
End Sub
Private Sub Text14_GotFocus()
    Text14.BackColor = &HC0FFFF
End Sub
Private Sub Text15_LostFocus()
    Text15.BackColor = &HFFFFFF
End Sub
Private Sub Text15_GotFocus()
    Text15.BackColor = &HC0FFFF
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &HFFFFFF
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HC0FFFF
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &HFFFFFF
End Sub
Private Sub Text2_GotFocus()
    Text2.BackColor = &HC0FFFF
End Sub
Private Sub Text3_LostFocus()
    Text3.BackColor = &HFFFFFF
End Sub
Private Sub Text3_GotFocus()
    Text3.BackColor = &HC0FFFF
End Sub
Private Sub Text6_LostFocus()
    Text6.BackColor = &HFFFFFF
End Sub
Private Sub Text6_GotFocus()
    Text6.BackColor = &HC0FFFF
End Sub
Private Sub Text7_LostFocus()
    Text7.BackColor = &HFFFFFF
End Sub
Private Sub Text7_GotFocus()
    Text7.BackColor = &HC0FFFF
End Sub
Private Sub Text9_LostFocus()
    Text9.BackColor = &HFFFFFF
End Sub
Private Sub Text9_GotFocus()
    Text9.BackColor = &HC0FFFF
End Sub
Private Sub Text16_LostFocus()
    Text16.BackColor = &HFFFFFF
End Sub
Private Sub Text16_GotFocus()
    Text16.BackColor = &HC0FFFF
End Sub
Private Sub Text17_LostFocus()
    Text17.BackColor = &HFFFFFF
End Sub
Private Sub Text17_GotFocus()
    Text17.BackColor = &HC0FFFF
End Sub
Private Sub Combo1_LostFocus()
    Combo1.BackColor = &HFFFFFF
End Sub
Private Sub Combo1_GotFocus()
    Combo1.BackColor = &HC0FFFF
End Sub
Private Sub Combo2_LostFocus()
    Combo2.BackColor = &HFFFFFF
End Sub
Private Sub Combo2_GotFocus()
    Combo2.BackColor = &HC0FFFF
End Sub
Private Sub Combo3_LostFocus()
    Combo3.BackColor = &HFFFFFF
End Sub
Private Sub Combo3_GotFocus()
    Combo3.BackColor = &HC0FFFF
End Sub
Private Sub Combo4_LostFocus()
    Combo4.BackColor = &HFFFFFF
End Sub
Private Sub Combo4_GotFocus()
    Combo4.BackColor = &HC0FFFF
End Sub
Private Sub Combo5_LostFocus()
    Combo5.BackColor = &HFFFFFF
End Sub
Private Sub Combo5_GotFocus()
    Combo5.BackColor = &HC0FFFF
End Sub
