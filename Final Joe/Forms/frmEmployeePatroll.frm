VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmEmployeePatroll 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Employee Payroll"
   ClientHeight    =   9255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   617
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   824
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   90
      Top             =   8730
   End
   Begin MOVERS.CandyButton CandyButton2 
      Height          =   390
      Left            =   90
      TabIndex        =   44
      Top             =   8115
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "l<"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   2700
      Left            =   3000
      ScaleHeight     =   2640
      ScaleWidth      =   6795
      TabIndex        =   37
      Top             =   5325
      Visible         =   0   'False
      Width           =   6855
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2085
         Left            =   150
         ScaleHeight     =   2055
         ScaleWidth      =   6540
         TabIndex        =   42
         Top             =   480
         Width           =   6570
         Begin MOVERS.LynxGrid3 listEntrieS1 
            Height          =   2010
            Left            =   15
            TabIndex        =   43
            Top             =   15
            Width           =   6510
            _ExtentX        =   11483
            _ExtentY        =   3545
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorBkg    =   16777215
            BackColorSel    =   15849673
            ForeColorSel    =   10248507
            GridColor       =   15461355
            FocusRectColor  =   10248507
            ThemeColor      =   2
            ColumnSort      =   -1  'True
            Striped         =   -1  'True
            SBackColor2     =   16777215
         End
      End
      Begin MOVERS.JOELine JOELine1 
         Height          =   30
         Left            =   45
         TabIndex        =   39
         Top             =   375
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   53
         BorderColor1    =   11325655
         BorderColor2    =   16185592
      End
      Begin VB.Image Image1 
         Height          =   285
         Left            =   75
         Picture         =   "frmEmployeePatroll.frx":0000
         Stretch         =   -1  'True
         Top             =   60
         Width           =   255
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   2610
         Left            =   -15
         Top             =   60
         Width           =   6780
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Deductions Info"
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
         TabIndex        =   38
         Top             =   105
         Width           =   3255
      End
      Begin VB.Image imgClose 
         Height          =   360
         Index           =   0
         Left            =   6420
         Picture         =   "frmEmployeePatroll.frx":058A
         Top             =   15
         Width           =   360
      End
      Begin VB.Image imgClose 
         Height          =   360
         Index           =   1
         Left            =   6420
         Picture         =   "frmEmployeePatroll.frx":0C74
         Top             =   0
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2820
      Left            =   660
      ScaleHeight     =   2790
      ScaleWidth      =   10920
      TabIndex        =   40
      Top             =   2490
      Width           =   10950
      Begin MOVERS.LynxGrid3 listEntries 
         Height          =   2745
         Left            =   15
         TabIndex        =   41
         Top             =   15
         Width           =   10890
         _ExtentX        =   19209
         _ExtentY        =   4842
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorBkg    =   16777215
         BackColorSel    =   15849673
         ForeColorSel    =   10248507
         GridColor       =   15461355
         FocusRectColor  =   10248507
         ThemeColor      =   2
         ColumnSort      =   -1  'True
         Striped         =   -1  'True
         SBackColor2     =   16777215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid gt 
      Height          =   1170
      Left            =   6030
      TabIndex        =   35
      Top             =   15000
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   2064
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin MOVERS.CandyButton BUTnEW 
      Height          =   465
      Left            =   10725
      TabIndex        =   30
      Top             =   7275
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "        Clear Form"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "frmEmployeePatroll.frx":135E
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
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9930
      TabIndex        =   14
      Top             =   5625
      Width           =   2220
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9930
      TabIndex        =   13
      Top             =   6030
      Width           =   2220
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   9930
      TabIndex        =   12
      Top             =   6465
      Width           =   2220
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1920
      MaxLength       =   13
      TabIndex        =   11
      Text            =   "0000000000000"
      Top             =   1380
      Width           =   3150
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      Top             =   1935
      Width           =   5745
   End
   Begin VB.ComboBox Combo3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      ItemData        =   "frmEmployeePatroll.frx":1AD8
      Left            =   8295
      List            =   "frmEmployeePatroll.frx":1ADA
      TabIndex        =   5
      Top             =   1380
      Width           =   3840
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   9600
      TabIndex        =   4
      Top             =   1860
      Width           =   2520
   End
   Begin MOVERS.JOETitleBar JOETitleBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   661
      Caption         =   "Employee Payroll [Drivers and Helpers]"
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
      ShadowColor     =   65535
      BorderColor     =   49344
      BackColor       =   12648447
   End
   Begin MOVERS.CandyButton ButPrev 
      Height          =   465
      Left            =   9045
      TabIndex        =   31
      Top             =   7290
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   820
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
      Picture         =   "frmEmployeePatroll.frx":1ADC
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
   Begin MOVERS.CandyButton ButSearch 
      Height          =   465
      Left            =   7365
      TabIndex        =   32
      Top             =   7275
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "    Search"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "frmEmployeePatroll.frx":2256
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
   Begin MOVERS.CandyButton ButSave 
      Height          =   465
      Left            =   5685
      TabIndex        =   33
      Top             =   7275
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "    Save"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "frmEmployeePatroll.frx":29D0
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
      Height          =   465
      Left            =   75
      TabIndex        =   34
      Top             =   7275
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   820
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "        View Deductions Info"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "frmEmployeePatroll.frx":314A
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
      Height          =   390
      Left            =   510
      TabIndex        =   45
      Top             =   8130
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "<<"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
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
      Height          =   390
      Left            =   945
      TabIndex        =   46
      Top             =   8130
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ">>"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin MOVERS.CandyButton CandyButton5 
      Height          =   390
      Left            =   1380
      TabIndex        =   47
      Top             =   8130
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ">l"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00808080&
      Height          =   1140
      Left            =   30
      Top             =   6915
      Width           =   12195
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "--------------------"
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
      Height          =   300
      Left            =   1845
      TabIndex        =   49
      Top             =   8205
      Width           =   4890
   End
   Begin VB.Label Ntrips 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   8055
      TabIndex        =   48
      Top             =   8175
      Width           =   4185
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   20
      Height          =   2820
      Left            =   495
      Top             =   2520
      Width           =   11280
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   885
      Left            =   4590
      TabIndex        =   36
      Top             =   375
      Visible         =   0   'False
      Width           =   7515
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "000.00"
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
      Left            =   4785
      TabIndex        =   29
      Top             =   6465
      Width           =   960
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "OTHERS :"
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
      Left            =   3885
      TabIndex        =   28
      Top             =   6465
      Width           =   990
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "000.00"
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
      Left            =   4785
      TabIndex        =   27
      Top             =   6090
      Width           =   960
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "SSS :"
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
      Left            =   4335
      TabIndex        =   26
      Top             =   6060
      Width           =   600
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "000.00"
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
      Left            =   4785
      TabIndex        =   25
      Top             =   5730
      Width           =   960
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "LOANS :"
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
      Left            =   4065
      TabIndex        =   24
      Top             =   5685
      Width           =   840
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "000.00"
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
      Left            =   1425
      TabIndex        =   23
      Top             =   6525
      Width           =   960
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "P-HEALTH :"
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
      Left            =   360
      TabIndex        =   22
      Top             =   6495
      Width           =   1170
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "000.00"
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
      Left            =   1425
      TabIndex        =   21
      Top             =   6120
      Width           =   960
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "SHORTAGES :"
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
      Left            =   90
      TabIndex        =   20
      Top             =   6090
      Width           =   1440
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "000.00"
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
      Left            =   1425
      TabIndex        =   19
      Top             =   5730
      Width           =   960
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "ADVANCES :"
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
      Left            =   255
      TabIndex        =   18
      Top             =   5700
      Width           =   1275
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      Height          =   1365
      Left            =   30
      Top             =   5565
      Width           =   7830
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total  :"
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
      Left            =   8865
      TabIndex        =   17
      Top             =   5700
      Width           =   1110
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Less: Deductions :"
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
      Left            =   8235
      TabIndex        =   16
      Top             =   6075
      Width           =   1800
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Salary :"
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
      Left            =   8850
      TabIndex        =   15
      Top             =   6495
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   4290
      Left            =   30
      Top             =   1290
      Width           =   12195
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Code:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   270
      TabIndex        =   10
      Top             =   1455
      Width           =   1950
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Covered :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   6705
      TabIndex        =   9
      Top             =   1425
      Width           =   1620
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   105
      TabIndex        =   8
      Top             =   1920
      Width           =   1755
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   8250
      TabIndex        =   7
      Top             =   1890
      Width           =   1350
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   6495
      Left            =   30
      Top             =   435
      Width           =   12195
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "822 MOVERS (PAILO)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   165
      TabIndex        =   3
      Top             =   465
      Width           =   3945
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "No. 322 San Ildefonso Alaminos, Laguna"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   180
      TabIndex        =   2
      Top             =   765
      Width           =   4545
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "SALARY OF EMPLOYEE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   180
      TabIndex        =   1
      Top             =   1020
      Width           =   2505
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   15
      Picture         =   "frmEmployeePatroll.frx":38C4
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   12360
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuEditsal 
         Caption         =   "Edit trip amount salary"
      End
      Begin VB.Menu mnuDelSal 
         Caption         =   "Delete trip salary"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRef 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "frmEmployeePatroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><>'
'Programmed By: Joehel V. del Rosario    '
'822 MOVERS(PAILO) Trucking Corporation  '
'Date: June 1, 2007                      '
'<><><><><><><><><><><><><><><><><><><><>'

Dim DS As Integer
Dim FRec As Boolean
Dim LRec As Boolean

Public Function Form_CanManageEmployee() As Boolean
    'If GetTxtVal(b8DPStudent.BoundData) > 0 Then
        Form_CanManageEmployee = True
    'End If
End Function

Sub LoadTruckType()
    OpenPBDataBase ("TruckPersonel")
    With PRFile
     .MoveFirst
     Do While Not .EOF
        If Not .EOF Then
            gt.Col = 0
            gt.Text = ![PlateNumber]
            gt.Col = 1
            gt.Text = ![Tructype]
            gt.Row = gt.Row + 1
            gt.Rows = gt.Rows + 1
        End If
     .MoveNext
     Loop
     .Close
    End With
End Sub

Private Sub ButNew_Click()
    Combo3.Clear
    'Text2.Text = ""
    listEntries.Redraw = False
    listEntries.Clear
    listEntries.Redraw = True
    listEntries.Refresh
    
    listEntrieS1.Redraw = False
    listEntrieS1.Clear
    listEntrieS1.Redraw = True
    listEntrieS1.Refresh
    
    Text1.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text9.Text = ""
    Label20.Caption = "000.00"
    Label15.Caption = "000.00"
    Label11.Caption = "000.00"
    Label13.Caption = "000.00"
    Label22.Caption = "000.00"
    Label23.Visible = False
    Call LOadcombo3
    Text2.SetFocus
    SendKeys "{HOME}+{END}"
End Sub

Private Sub ButPrev_Click()
    Call Text1_KeyPress(13)

'On Error Resume Next
    Call OpenDeducTionINFO
    FormPreviewPayroll.Label1.Caption = Trim(Label1.Caption) & "  "
    FormPreviewPayroll.Label2.Caption = Trim(Label4.Caption) & "  "
    FormPreviewPayroll.Label7.Caption = Trim(Combo3.Text) & "  "
    FormPreviewPayroll.Label9.Caption = Trim(Text4.Text) & "  "
    FormPreviewPayroll.Label11.Caption = Trim(Text5.Text) & "  "
    FormPreviewPayroll.Label13.Caption = Trim(Text9.Text) & "  "
    FormPreviewPayroll.Label5.Caption = Trim(UCase(Text1.Text)) & " - [" & Trim(Text3.Text) & "]"
    FormPreviewPayroll.Label16.Caption = Format(Label17.Caption, "###.00")
    FormPreviewPayroll.Label18.Caption = Format(Label20.Caption, "###.00")
    FormPreviewPayroll.Label20.Caption = Format(Label15.Caption, "###.00")
    FormPreviewPayroll.Label22.Caption = Format(Label11.Caption, "###.00")
    FormPreviewPayroll.Label24.Caption = Format(Label13.Caption, "###.00")
    FormPreviewPayroll.Label26.Caption = Format(Label22.Caption, "###.00")
    
    
    'MDIMainForm.AddChild FormPreviewPayroll, True
    FormPreviewPayroll.Show 1
    'SetParent FormPreviewPayroll.hWnd, FormMain.Picture1.hWnd
    
    'FormPreviewPayroll.Show
   
End Sub

Private Sub ButSave_Click()
    If Text1.Text <> "" And Combo3.Text <> "" Then
    'open for payroll personnels
    On Error Resume Next
    OpenPBDataBase ("PayrollPersonels")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM PayrollPersonels WHERE Names Like '" & Trim(Text1.Text) & "' ")
    With PRFile
        If Not .EOF Then
            'do nothing
        Else
            .AddNew
                ![Names] = Trim(Text1.Text)
            .Update
        End If
        .Close
    End With
    'Paid Payroll status
    OpenPBDataBase ("Payrolls")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM Payrolls WHERE ecode Like '" & Trim(Text1.Text) & "' and Coverdate Like '" & Trim(Combo3.Text) & "' ")
    With PRFile
      .MoveFirst
       Do While Not .EOF
        If Not .EOF Then
            .Edit
                ![Status] = "0"
            .Update
        End If
        .MoveNext
       Loop
        .Close
    End With
    'Paid deduction Status
    OpenPBDataBase ("DeductionsInfo")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM DeductionsInfo WHERE Dname Like '" & Trim(Text1.Text) & "' and Status Like '" & "1" & "' ")
                                     
    With PRFile
       .MoveFirst
           Do While Not .EOF
            If Not .EOF Then
                .Edit
                    ![Status] = "0"
                .Update
            End If
             .MoveNext
           Loop
        .Close
    End With

    End If
End Sub

Private Sub ButSearch_Click()
    If Text2.Text = "" Then
        Call Text1_KeyPress(13)
    ElseIf Text1.Text = "" Then
        Call Text2_KeyPress(13)
    ElseIf Text1.Text <> "" And Text2.Text <> "" Then
        Call Text2_KeyPress(13)
    Else
        Text1.SetFocus
    End If
    Text2.SetFocus
'    SendKeys "{HOME}+{END}"
End Sub

Private Sub CandyButton1_Click()
    Picture1.Visible = True
    Call OpenDeducTionINFO
End Sub

Private Sub CandyButton2_Click()
 EPRepR = 0
 Text1.Text = Trim(frmPayrollReport.listEntries.CellText(EPRepR, 1))
 Call Text1_KeyPress(13)
 CandyButton2.Enabled = False
 CandyButton3.Enabled = False
 CandyButton4.Enabled = True
 CandyButton5.Enabled = True
End Sub
Private Sub CandyButton3_Click()
  EPRepR = EPRepR - 1
  If EPRepR >= 0 Then
    Text1.Text = Trim(frmPayrollReport.listEntries.CellText(EPRepR, 1))
    Call Text1_KeyPress(13)
    CandyButton4.Enabled = True
    CandyButton5.Enabled = True
  Else
    CandyButton2.Enabled = False
    CandyButton3.Enabled = False
  End If
  
End Sub
Private Sub CandyButton4_Click()
  EPRepR = EPRepR + 1
  If EPRepR <= frmPayrollReport.listEntries.RowCount - 1 Then
    Text1.Text = Trim(frmPayrollReport.listEntries.CellText(EPRepR, 1))
    Call Text1_KeyPress(13)
    CandyButton2.Enabled = True
    CandyButton3.Enabled = True
  Else
    CandyButton4.Enabled = False
    CandyButton5.Enabled = False
  End If
End Sub
Private Sub CandyButton5_Click()
 EPRepR = frmPayrollReport.listEntries.RowCount - 1
 Text1.Text = Trim(frmPayrollReport.listEntries.CellText(EPRepR, 1))
 Call Text1_KeyPress(13)
 CandyButton2.Enabled = True
 CandyButton3.Enabled = True
 CandyButton4.Enabled = False
 CandyButton5.Enabled = False
End Sub
Private Sub Combo3_Click()
    If Text1.Text <> "" And Text2.Text <> "" Then
        ButSearch_Click
    End If
End Sub

Private Sub Form_Load()
    'set list columns for the list grid
    With listEntries
        .Redraw = False
        .AddColumn "     Date Trip", 80   '0
        .AddColumn "  Plate No.", 60   '1
        .AddColumn "  Truck Type", 70   '2
        .AddColumn "      Origin", 70   '3
        .AddColumn "                                       Custumers", 300   '4
        .AddColumn "       Cases", 70   '5
        .AddColumn "       Amount", 70   '6
        .Redraw = True
        .Refresh
    End With
    'set list columns for the list grid
    With listEntrieS1
        .Redraw = False
        .AddColumn "           Date ", 100   '0
        .AddColumn "           Amount", 100   '1
        .AddColumn "             Type of Deductions", 230   '2
        .Redraw = True
        .Refresh
    End With
        
        'Call Load functions
         Call LOadcombo3
         Call LoadTruckType
End Sub
Private Sub Form_Activate()
    MDIMainForm.JST(2).Expanded = True
    MDIMainForm.ActivateChild Me
    
    If EPReports = False Then
        CandyButton2.Enabled = False
        CandyButton3.Enabled = False
        CandyButton4.Enabled = False
        CandyButton5.Enabled = False
    Else
        CandyButton2.Enabled = True
        CandyButton3.Enabled = True
        CandyButton4.Enabled = True
        CandyButton5.Enabled = True
        
    End If
    
    If EPRepR = 0 Then
        CandyButton2.Enabled = False
        CandyButton3.Enabled = False
    ElseIf EPRepR = frmPayrollReport.listEntries.RowCount - 1 Then
        CandyButton4.Enabled = False
        CandyButton5.Enabled = False
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MDIMainForm.Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & "NULL.pic")
    MDIMainForm.JST(2).Expanded = False
    MDIMainForm.RemoveChild Me.Name
End Sub
Sub LOadcombo3()
    On Error Resume Next
    OpenPBDataBase ("DateCover")
    With PRFile
      .MoveFirst
        Do While Not .EOF
            Combo3.AddItem ![CoveredDate]
            
            If ![Status] = "1" Then
                Combo3.Text = ![CoveredDate]
            End If
            .MoveNext
        Loop
   End With
End Sub

Private Sub imgClose_Click(Index As Integer)
    Select Case Index
            Case 0
                Picture1.Visible = False
            Case 1
                Picture1.Visible = False
    End Select
End Sub

Private Sub imgClose_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
            Case 0
                imgClose(0).Visible = False
            Case 1
                imgClose(1).Visible = True
    End Select
End Sub

Private Sub listEntries_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Or KeyCode = 114 Then
        Call mnuEditsal_Click
    End If
    
    KeyCode = 0
    listEntries.Refresh
End Sub

Private Sub listEntries_KeyPress(KeyAscii As Integer)
    If KeyAscii = 114 Then
        Call mnuEditsal_Click
    End If
    
    KeyAscii = 0
    listEntries.Refresh
End Sub

Private Sub listEntries_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Or KeyCode = 114 Then
        Call mnuEditsal_Click
    End If
    
    KeyCode = 0
    listEntries.Refresh
End Sub

Private Sub listEntries_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu Me.MnuFile
    End If
End Sub

Private Sub mnuDelSal_Click()
Dim t1 As String
Dim t2 As String
Dim t3 As String
Dim t4 As String
Dim t5 As String
Dim t6 As String
 
 
         With listEntries
                t1 = .CellText(.Row, 0)
                t2 = .CellText(.Row, 1)
                t3 = .CellText(.Row, 3)
                t4 = .CellText(.Row, 4)
                t5 = Format(.CellText(.Row, 5), "###")
                t6 = Format(.CellText(.Row, 6), "###")
          End With
          
          
 'On Error Resume Next
    OpenPBDataBase ("Payrolls")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM Payrolls WHERE ecode Like '" & Trim(Text1.Text) & "' and Coverdate Like '" & Trim(Combo3.Text) & _
                                      "' and DateTrip Like '" & Trim(t1) & _
                                      "' and truckNumber Like '" & Trim(t2) & _
                                      "' and TPO Like '" & Trim(t3) & _
                                      "' and Particulars Like '" & Trim(t4) & _
                                      "' and Cases Like '" & Trim(t5) & _
                                      "' and Amount Like '" & Trim(t6) & "' ")
    With PRFile
        If Not .EOF Then
            .Delete
        End If
        .Close
    End With
DD:
    listEntries.RemoveItem (listEntries.Row)
End Sub

Private Sub mnuEditsal_Click()
Dim t1 As String
Dim t2 As String
Dim t3 As String
Dim t4 As String
Dim t5 As String
Dim t6 As String
Dim IpB As String
 
 
         With listEntries
                t1 = .CellText(.Row, 0)
                t2 = .CellText(.Row, 1)
                t3 = .CellText(.Row, 3)
                t4 = .CellText(.Row, 4)
                t5 = Format(.CellText(.Row, 5), "###")
                t6 = Format(.CellText(.Row, 6), "###")
          End With
          
          
        IpB = InputBox("Enter Salary amount..", "Edit Salary")
          
 'On Error Resume Next
    OpenPBDataBase ("Payrolls")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM Payrolls WHERE ecode Like '" & Trim(Text1.Text) & "' and Coverdate Like '" & Trim(Combo3.Text) & _
                                      "' and DateTrip Like '" & Trim(t1) & _
                                      "' and truckNumber Like '" & Trim(t2) & _
                                      "' and TPO Like '" & Trim(t3) & _
                                      "' and Particulars Like '" & Trim(t4) & _
                                      "' and Cases Like '" & Trim(t5) & _
                                      "' and Amount Like '" & Trim(t6) & "' ")
    'MsgBox t1 & t2 & t3 & t4 & t5 & t6
    With PRFile
         If Not .EOF Then
            .Edit
                ![Amount] = Trim(IpB)
            .Update
        End If
        .Close
    End With
    
    Call Text1_KeyPress(13)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose(0).Visible = True
End Sub

Private Sub Text1_Change()
    Call AutoTXTcomplete(MDIMainForm.List1, Text1)
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &HC0FFFF
    Text1.FontBold = True
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &HFFFFFF
    Text1.FontBold = False
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
        DelKey = True
        Exit Sub
    End If
    If KeyCode = vbKeyF2 Then
        'Call ButPrev_Click
    End If
    
End Sub

Sub OpenMe()
    Call Text1_KeyPress(13)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
    Call RefreshRecs
    
    Label23.Visible = False
    
    MDIMainForm.Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & "NULL.pic")
    
    OpenPBDataBase ("EmployeeInfo")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM EmployeeInfo WHERE Ename Like '" & Trim(Text1.Text) & "' ")
    With PRFile
        If Not .EOF Then
                Text2.Text = ![ECOde]
                Text3.Text = ![EOccupation]
        Else
            Call ButNew_Click
            .Close
            'GoTo FFF
        End If
        .Close
    End With
    Call OpenSALARYDeduction
    
    MDIMainForm.Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & Trim(Text2.Text) & ".pic")
    
    End If
    Exit Sub
'FFF:
 If Error = True Then
    
 End If
End Sub
Sub OpenSALARYDeduction()
Dim r As Integer
Dim C As Integer
Dim iL As Long
On Error Resume Next
Dim PD As String
Dim PP As String
Dim PT As String
Dim po As String
Dim PC As String
Dim PCS As String
Dim PA As String
    
    listEntries.Redraw = False
    listEntries.Clear

  Dim trap1 As Integer
  Dim Trap2 As Integer
  Dim TSALAry As Double
   'open for salary
    OpenPBDataBase ("Payrolls")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM Payrolls WHERE ecode Like '" & Trim(Text1.Text) & "' and Coverdate Like '" & Trim(Combo3.Text) & "' ")
                                     
    With PRFile
      .MoveFirst
       Do While Not .EOF
        If Not .EOF Then
            trap1 = 1
                PD = Trim(![DateTrip])
                PP = Trim(![truckNumber])
                
                For r = 0 To gt.Rows - 1
                    gt.Row = r
                    gt.Col = 0
                    If Trim(PP) = Trim(gt.Text) Then
                        gt.Col = 1
                        PT = Trim(gt.Text)
                    Exit For
                    End If
                Next r
                
                po = Trim(![TPO])
                PC = Trim(![Particulars])
                PCS = Trim(![Cases])
                PA = Trim(![Amount])
                        
                        With listEntries
                            .AddItem (Trim(PD)), 0
                            .CellAlignment(iL, 0) = lgAlignCenterCenter
                            .CellAlignment(iL, 2) = lgAlignCenterCenter
                            .CellAlignment(iL, 3) = lgAlignCenterCenter
                            .CellAlignment(iL, 5) = lgAlignRightCenter
                            .CellAlignment(iL, 6) = lgAlignRightCenter
                            .CellText(iL, 1) = Trim(PP)
                            .CellText(iL, 2) = Trim(PT)
                            .CellText(iL, 3) = Trim(po)
                            .CellText(iL, 4) = Trim(PC)
                            .CellText(iL, 5) = Trim(Format(PCS, "###,###"))
                            .CellText(iL, 6) = Trim(Format(PA, "###,###.00"))
                            iL = iL + 1
                        End With
                
                TSALAry = Round(Val(TSALAry) + Val(![Amount]), 2)
                            
                If ![Status] = "0" Then
                     Trap2 = 1
                End If
        End If
        .MoveNext
       Loop
        .Close
    End With
    
    listEntries.Redraw = True
    listEntries.Refresh

    
    'open for deductions
    Dim s As Double
    Dim p As Double
    Dim a As Double
    Dim d As Double
    Dim O As Double
    Dim L As Double
    Dim aa As Double
    OpenPBDataBase ("DeductionsInfo")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM DeductionsInfo WHERE Dname Like '" & Trim(Text1.Text) & "' and DateCover LIKE '" & Trim(Combo3.Text) & "' ")
    With PRFile
       .MoveFirst
           Do While Not .EOF
            If Not .EOF Then
                If ![DType] = "SSS" Then
                    s = s + Val(![DAmount])
                ElseIf ![DType] = "P-Healt" Then
                    p = p + Val(![DAmount])
                ElseIf ![DType] = "Advance" Then
                    a = a + Val(![DAmount])
                ElseIf ![DType] = "Shortage" Then
                    d = d + Val(![DAmount])
                ElseIf ![DType] = "Loans" Then
                    L = L + Val(![DAmount])
                Else 'If ![DType] = "Others" Then
                    O = O + Val(![DAmount])
                End If
            End If
             .MoveNext
           Loop
        .Close
    End With
    
                            Label20.Caption = Format(Val(s), "###,###.00")
                            Label15.Caption = Format(Val(p), "###,###.00")
                            Label11.Caption = Format(Val(a), "###,###.00")
                            Label13.Caption = Format(Val(d), "###,###.00")
                            Label17.Caption = Format(Val(L), "###,###.00")
                            Label22.Caption = Format(Val(O), "###,###.00")
                            
                            Text4.Text = Format(Round(Val(TSALAry), 2), "###,###,###.00")
                            Text5.Text = Format(Round(Val(Val(s) + Val(p)) + Val(Val(a) + Val(d) + Val(O) + Val(L)), 2), "###,###,###.00")
                            Text9.Text = Format(Round(Val(TSALAry), 2) - Round(Val(Val(s) + Val(p)) + Val(Val(a) + Val(d) + Val(O) + Val(L)), 2), "###,###,###.00")
                            
    If Trap2 = 1 Then
        Label23.Visible = True
        Label23.Caption = "PAID"
    Else
        'Label23.Visible = False
    End If
        
    If trap1 = 0 Then
        Label23.Visible = True
        Label23.Caption = "No Account"
    End If
    Ntrips.Caption = "Total Trips: " & listEntries.RowCount
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    Call RefreshRecs
    Label23.Visible = False
    
    MDIMainForm.Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & "NULL.pic")
    
    OpenPBDataBase ("EmployeeInfo")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM EmployeeInfo WHERE Ecode Like '" & Text2.Text & "' ")
    With PRFile
        If Not .EOF Then
                Text1.Text = ![Ename]
                Text3.Text = ![EOccupation]
        Else
            Call ButNew_Click
            .Close
        End If
       .Close
    End With
    Call OpenSALARYDeduction
    MDIMainForm.Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & Trim(Text2.Text) & ".pic")
End If
End Sub
Private Sub text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
        DelKey = True
        Exit Sub
    End If
    
    If KeyCode = vbKeyF2 Then
        'Call ButPrev_Click
    End If

End Sub

Private Sub Text2_Change()
    Call AutoTXTcomplete(MDIMainForm.List2, Text2)
End Sub

Sub OpenDeducTionINFO()
Dim sts As Integer
Dim DD As String
Dim DA As String
Dim DT As String
Dim iL As Long
On Error Resume Next
    'clear list
    listEntrieS1.Redraw = False
    listEntrieS1.Clear
            OpenPBDataBase ("DeductionsInfo")
            Set PRFile = PDbase.OpenRecordset("SELECT * FROM DeductionsInfo WHERE DName Like '" & Trim(Text1.Text) & "' and DateCover Like '" & Trim(Combo3.Text) & "' ")
            With PRFile
              .MoveFirst
               Do While Not .EOF
                    If Not .EOF Then
                        DD = Trim(![dDate])
                        DA = Trim(![DAmount])
                        DT = Trim(![DType])
                        
                        With listEntrieS1
                            .AddItem (Trim(DD)), 0
                            .CellAlignment(iL, 0) = lgAlignCenterCenter
                            .CellAlignment(iL, 1) = lgAlignRightCenter
                            .CellAlignment(iL, 2) = lgAlignLeftCenter
                            .CellText(iL, 1) = Trim(Format(DA, "###,###.00"))
                            .CellText(iL, 2) = Trim(DT)
                            iL = iL + 1
                        End With
                    
                    End If
               .MoveNext
               Loop
                .Close
            End With
    listEntrieS1.Redraw = True
    listEntrieS1.Refresh
End Sub
Sub RefreshRecs()
    If EPReports = True Then
        Label25.Caption = "Record " & EPRepR + 1 & " of " & frmPayrollReport.listEntries.RowCount
    End If
End Sub

Private Sub Timer1_Timer()
   If Text1.Text <> "" Then
        Text9.Visible = Not Text9.Visible
   Else
        Text9.Visible = True
   End If
End Sub
