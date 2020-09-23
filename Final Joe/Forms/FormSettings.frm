VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FormSettings 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   10065
   ClientLeft      =   495
   ClientTop       =   555
   ClientWidth     =   12405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   12405
   WindowState     =   2  'Maximized
   Begin MOVERS.JOETitleBar JOETitleBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   661
      Caption         =   "Settings"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   8115
      Left            =   -15
      TabIndex        =   1
      Top             =   345
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   14314
      _Version        =   393216
      TabOrientation  =   1
      Tab             =   2
      TabHeight       =   706
      BackColor       =   16777215
      TabCaption(0)   =   "Truck Personnels"
      TabPicture(0)   =   "FormSettings.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Salaries and Wages"
      TabPicture(1)   =   "FormSettings.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Shape10"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label12"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Shape9"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Shape8"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Shape4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Shape6"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Shape5"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label3"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Shape1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label4"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label6"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label7"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label8"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label9"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label10"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label19"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label20"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label21"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "CandyButton4"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "CandyButton8"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "CandyButton6"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Text30"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Text9"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Text8"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Text7"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Text6"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Text5"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Text12"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Text13"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Text14"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Text15"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Text16"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Text17"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Text18"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Text19"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Text20"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Text21"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Text22"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Text23"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Text24"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Text25"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "Text26"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "Text27"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "Text28"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "Text29"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "Combo3"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).ControlCount=   46
      TabCaption(2)   =   "Customers"
      TabPicture(2)   =   "FormSettings.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         Caption         =   "Customers and wages Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Left            =   150
         TabIndex        =   61
         Top             =   120
         Width           =   11955
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
            Height          =   900
            Left            =   1470
            MultiSelect     =   2  'Extended
            TabIndex        =   62
            Top             =   705
            Visible         =   0   'False
            Width           =   5040
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
            Height          =   360
            Left            =   1470
            TabIndex        =   65
            Top             =   360
            Width           =   5040
         End
         Begin VB.ComboBox Combo9 
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
            Height          =   360
            ItemData        =   "FormSettings.frx":0054
            Left            =   1455
            List            =   "FormSettings.frx":005E
            TabIndex        =   64
            Top             =   765
            Width           =   5070
         End
         Begin VB.ComboBox Combo8 
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
            Height          =   360
            ItemData        =   "FormSettings.frx":0076
            Left            =   1455
            List            =   "FormSettings.frx":0083
            TabIndex        =   63
            Top             =   1230
            Width           =   5085
         End
         Begin MOVERS.CandyButton CandyButton1 
            Height          =   465
            Left            =   6645
            TabIndex        =   66
            Top             =   1035
            Width           =   1350
            _ExtentX        =   2381
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
            Caption         =   "   &Save"
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            Picture         =   "FormSettings.frx":00A3
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
            Height          =   465
            Left            =   8250
            TabIndex        =   67
            Top             =   1050
            Width           =   1350
            _ExtentX        =   2381
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
            Caption         =   "      &Delete"
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            Picture         =   "FormSettings.frx":081D
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
         Begin MOVERS.CandyButton CandyButton9 
            Height          =   465
            Left            =   9900
            TabIndex        =   68
            Top             =   1050
            Width           =   1350
            _ExtentX        =   2381
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
            Caption         =   "      &New"
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            Picture         =   "FormSettings.frx":0F97
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
         Begin MOVERS.LynxGrid3 listEntries 
            Height          =   5565
            Left            =   675
            TabIndex        =   69
            Top             =   1725
            Width           =   10590
            _ExtentX        =   18680
            _ExtentY        =   9816
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColorBkg    =   16056319
            BackColorSel    =   15849673
            ForeColorSel    =   10248507
            GridColor       =   15461355
            FocusRectColor  =   10248507
            ThemeColor      =   2
            ThemeStyle      =   3
            ColumnSort      =   -1  'True
            Striped         =   -1  'True
            SBackColor1     =   16056319
            SBackColor2     =   14940667
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Wages :"
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
            Left            =   135
            TabIndex        =   72
            Top             =   795
            Width           =   1755
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Customer :"
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
            TabIndex        =   71
            Top             =   360
            Width           =   1755
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Source :"
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
            Left            =   120
            TabIndex        =   70
            Top             =   1230
            Width           =   1755
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Truck Personnels Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6855
         Left            =   -74685
         TabIndex        =   41
         Top             =   360
         Width           =   11835
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2565
            TabIndex        =   51
            Top             =   3075
            Width           =   7245
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2550
            TabIndex        =   50
            Top             =   3570
            Width           =   7245
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2550
            TabIndex        =   49
            Top             =   4140
            Width           =   7245
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2565
            TabIndex        =   48
            Top             =   4650
            Width           =   7245
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2550
            TabIndex        =   47
            Top             =   5160
            Width           =   7245
         End
         Begin VB.TextBox Combo4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2565
            TabIndex        =   46
            Top             =   2115
            Width           =   7245
         End
         Begin VB.ComboBox Combo6 
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
            Height          =   360
            ItemData        =   "FormSettings.frx":1711
            Left            =   2550
            List            =   "FormSettings.frx":1713
            TabIndex        =   45
            Top             =   1545
            Width           =   1755
         End
         Begin VB.ComboBox Combo5 
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
            Height          =   360
            ItemData        =   "FormSettings.frx":1715
            Left            =   2535
            List            =   "FormSettings.frx":1717
            TabIndex        =   44
            Top             =   1020
            Width           =   1755
         End
         Begin VB.ComboBox Combo1 
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
            Height          =   360
            ItemData        =   "FormSettings.frx":1719
            Left            =   2535
            List            =   "FormSettings.frx":171B
            TabIndex        =   43
            Top             =   450
            Width           =   7290
         End
         Begin VB.ComboBox Combo2 
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
            Height          =   360
            ItemData        =   "FormSettings.frx":171D
            Left            =   2565
            List            =   "FormSettings.frx":1730
            TabIndex        =   42
            Text            =   "0"
            Top             =   2595
            Width           =   750
         End
         Begin MOVERS.CandyButton ButSave 
            Height          =   465
            Left            =   4845
            TabIndex        =   52
            Top             =   6180
            Width           =   1350
            _ExtentX        =   2381
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
            Caption         =   "   &Save"
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            Picture         =   "FormSettings.frx":1743
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
         Begin MOVERS.CandyButton ButDelete 
            Height          =   465
            Left            =   6780
            TabIndex        =   53
            Top             =   6180
            Width           =   1350
            _ExtentX        =   2381
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
            Caption         =   "      &Delete"
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            Picture         =   "FormSettings.frx":1EBD
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
            Left            =   8610
            TabIndex        =   54
            Top             =   6180
            Width           =   1350
            _ExtentX        =   2381
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
            Caption         =   "      &Search"
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            Picture         =   "FormSettings.frx":2637
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
         Begin MOVERS.CandyButton ButNew 
            Height          =   465
            Left            =   10320
            TabIndex        =   55
            Top             =   6180
            Width           =   1350
            _ExtentX        =   2381
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
            Caption         =   "      &New"
            IconHighLiteColor=   0
            CaptionHighLiteColor=   0
            Picture         =   "FormSettings.frx":2DB1
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
         Begin VB.Label Label11 
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
            Left            =   1125
            TabIndex        =   60
            Top             =   1560
            Width           =   2205
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00808080&
            Height          =   855
            Left            =   0
            Top             =   5985
            Width           =   11820
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Personel of :"
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
            Left            =   555
            TabIndex        =   59
            Top             =   495
            Width           =   1845
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Dirvers Name:"
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
            Left            =   1035
            TabIndex        =   58
            Top             =   2100
            Width           =   1500
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Helpers :"
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
            Left            =   960
            TabIndex        =   57
            Top             =   2610
            Width           =   1680
         End
         Begin VB.Label Label5 
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
            Left            =   945
            TabIndex        =   56
            Top             =   1080
            Width           =   2205
         End
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
         ItemData        =   "FormSettings.frx":352B
         Left            =   -72735
         List            =   "FormSettings.frx":352D
         TabIndex        =   26
         Top             =   1095
         Width           =   7125
      End
      Begin VB.TextBox Text29 
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
         Height          =   360
         Left            =   -65040
         TabIndex        =   25
         Top             =   5130
         Width           =   1545
      End
      Begin VB.TextBox Text28 
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
         Height          =   360
         Left            =   -66885
         TabIndex        =   24
         Top             =   5145
         Width           =   1545
      End
      Begin VB.TextBox Text27 
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
         Height          =   360
         Left            =   -68715
         TabIndex        =   23
         Top             =   5145
         Width           =   1545
      End
      Begin VB.TextBox Text26 
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
         Height          =   360
         Left            =   -70530
         TabIndex        =   22
         Top             =   5160
         Width           =   1545
      End
      Begin VB.TextBox Text25 
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
         Height          =   360
         Left            =   -72390
         TabIndex        =   21
         Top             =   5145
         Width           =   1545
      End
      Begin VB.TextBox Text24 
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
         Height          =   360
         Left            =   -74250
         TabIndex        =   20
         Top             =   5145
         Width           =   1545
      End
      Begin VB.TextBox Text23 
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
         Height          =   360
         Left            =   -65040
         TabIndex        =   19
         Top             =   3960
         Width           =   1545
      End
      Begin VB.TextBox Text22 
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
         Height          =   360
         Left            =   -66885
         TabIndex        =   18
         Top             =   3975
         Width           =   1545
      End
      Begin VB.TextBox Text21 
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
         Height          =   360
         Left            =   -68730
         TabIndex        =   17
         Top             =   3975
         Width           =   1545
      End
      Begin VB.TextBox Text20 
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
         Height          =   360
         Left            =   -70530
         TabIndex        =   16
         Top             =   3975
         Width           =   1545
      End
      Begin VB.TextBox Text19 
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
         Height          =   360
         Left            =   -72405
         TabIndex        =   15
         Top             =   3975
         Width           =   1545
      End
      Begin VB.TextBox Text18 
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
         Height          =   360
         Left            =   -74265
         TabIndex        =   14
         Top             =   3975
         Width           =   1545
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
         Height          =   360
         Left            =   -65040
         TabIndex        =   13
         Top             =   2850
         Width           =   1545
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
         Height          =   360
         Left            =   -66885
         TabIndex        =   12
         Top             =   2850
         Width           =   1545
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
         Height          =   360
         Left            =   -68730
         TabIndex        =   11
         Top             =   2850
         Width           =   1545
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
         Height          =   360
         Left            =   -70530
         TabIndex        =   10
         Top             =   2850
         Width           =   1560
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
         Height          =   360
         Left            =   -72405
         TabIndex        =   9
         Top             =   2850
         Width           =   1545
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
         Height          =   360
         Left            =   -74265
         TabIndex        =   8
         Top             =   2865
         Width           =   1545
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
         Height          =   360
         Left            =   -74220
         TabIndex        =   7
         Top             =   6225
         Width           =   1545
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
         Height          =   360
         Left            =   -72360
         TabIndex        =   6
         Top             =   6225
         Width           =   1545
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
         Height          =   360
         Left            =   -70500
         TabIndex        =   5
         Top             =   6240
         Width           =   1545
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
         Height          =   360
         Left            =   -68685
         TabIndex        =   4
         Top             =   6225
         Width           =   1545
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
         Height          =   360
         Left            =   -66855
         TabIndex        =   3
         Top             =   6225
         Width           =   1545
      End
      Begin VB.TextBox Text30 
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
         Height          =   360
         Left            =   -65010
         TabIndex        =   2
         Top             =   6210
         Width           =   1545
      End
      Begin MOVERS.CandyButton CandyButton6 
         Height          =   465
         Left            =   -66285
         TabIndex        =   27
         Top             =   7080
         Width           =   1350
         _ExtentX        =   2381
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
         Caption         =   "   &Save"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "FormSettings.frx":352F
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
      Begin MOVERS.CandyButton CandyButton8 
         Height          =   465
         Left            =   -64740
         TabIndex        =   28
         Top             =   7080
         Width           =   1350
         _ExtentX        =   2381
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
         Caption         =   "      &New"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "FormSettings.frx":3CA9
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
         Height          =   465
         Left            =   -65490
         TabIndex        =   29
         Top             =   1035
         Width           =   1350
         _ExtentX        =   2381
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
         Caption         =   "      &Search"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "FormSettings.frx":4423
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
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "ELF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -74415
         TabIndex        =   40
         Top             =   5550
         Width           =   1755
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Forward"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -74415
         TabIndex        =   39
         Top             =   4395
         Width           =   1755
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "10 Wheels"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -74430
         TabIndex        =   38
         Top             =   3285
         Width           =   1755
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "5 Helper"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -65115
         TabIndex        =   37
         Top             =   2175
         Width           =   1755
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "4 Helper"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -67005
         TabIndex        =   36
         Top             =   2190
         Width           =   1755
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "3 Helper"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -69195
         TabIndex        =   35
         Top             =   2175
         Width           =   1755
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2  Helper"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -70710
         TabIndex        =   34
         Top             =   2190
         Width           =   1755
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1 Helper"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -72555
         TabIndex        =   33
         Top             =   2190
         Width           =   1755
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Driver"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -74355
         TabIndex        =   32
         Top             =   2190
         Width           =   1755
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   3975
         Left            =   -74460
         Top             =   2085
         Width           =   11100
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tracking to :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -74520
         TabIndex        =   31
         Top             =   1140
         Width           =   1755
      End
      Begin VB.Shape Shape5 
         Height          =   4905
         Left            =   -65205
         Top             =   2085
         Width           =   1845
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00808080&
         Height          =   4905
         Left            =   -68865
         Top             =   2085
         Width           =   1845
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00808080&
         Height          =   4905
         Left            =   -72540
         Top             =   2085
         Width           =   1845
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00808080&
         Height          =   450
         Left            =   -74460
         Top             =   2085
         Width           =   11100
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   11235
         Top             =   -195
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   1
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "FormSettings.frx":4B9D
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H00808080&
         Height          =   1200
         Left            =   -74460
         Top             =   3690
         Width           =   11100
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "4 WHEELS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -74385
         TabIndex        =   30
         Top             =   6630
         Width           =   1755
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H00808080&
         Height          =   945
         Left            =   -74460
         Top             =   6045
         Width           =   11100
      End
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   -2205
      Picture         =   "FormSettings.frx":4EEF
      Stretch         =   -1  'True
      Top             =   9000
      Width           =   15795
   End
End
Attribute VB_Name = "FormSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><>'
'Programmed By: Joehel V. del Rosario    '
'822 MOVERS(PAILO) Trucking Corporation  '
'Date: June 1, 2007                      '
'<><><><><><><><><><><><><><><><><><><><>'

Private Sub ButDelete_Click()
 If Combo1.Text <> "" And Combo5.Text <> "" And Combo4.Text <> "" And Text2.Text <> "" Then
    OpenPBDataBase ("TruckPersonel")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM TruckPersonel WHERE PlateNumber Like '" & Trim(Combo5.Text) & "' ")
    With PRFile
        If Not .EOF Then
            .Delete
           MsgBox "Data Deleted", vbInformation, "Deleted"
        Else
           MsgBox "No Data to Delete", vbCritical, "Not Found"
        End If
    End With
 End If

End Sub

Private Sub ButNew_Click()
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text11.Text = ""
    Text10.Text = ""
    Combo4.Text = ""
    Combo5.Text = ""
    Combo1.Text = ""
    Combo6.Text = ""

End Sub

Private Sub ButSave_Click()
 If Combo1.Text <> "" And Combo5.Text <> "" And Combo4.Text <> "" And Text2.Text <> "" Then
    OpenPBDataBase ("TruckPersonel")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM TruckPersonel WHERE PlateNumber Like '" & Trim(Combo5.Text) & "' ")
    With PRFile
        If Not .EOF Then
            .Edit
                ![Driver] = Trim(Combo4.Text)
                ![Tructype] = Trim(Combo6.Text)
                If Text2.Visible = True Then
                    ![Helper1] = Trim(Text2.Text)
                End If
                If Text10.Visible = True Then
                    ![Helper2] = Trim(Text10.Text)
                End If
                If Text11.Visible = True Then
                    ![Helper3] = Trim(Text11.Text)
                End If
                If Text3.Visible = True Then
                    ![Helper4] = Trim(Text3.Text)
                End If
                If Text4.Visible = True Then
                    ![Helper5] = Trim(Text4.Text)
                End If
                ![TotalPerson] = Trim(Combo2.Text) + 1
                ![Perdonels] = Trim(Combo1.Text)
            .Update
        Else
            .AddNew
                ![PlateNumber] = Trim(Combo5.Text)
                ![Tructype] = Trim(Combo6.Text)
                ![Driver] = Trim(Combo4.Text)
                If Text2.Visible = True Then
                    ![Helper1] = Trim(Text2.Text)
                End If
                If Text10.Visible = True Then
                    ![Helper2] = Trim(Text10.Text)
                End If
                If Text11.Visible = True Then
                    ![Helper3] = Trim(Text11.Text)
                End If
                If Text3.Visible = True Then
                    ![Helper4] = Trim(Text3.Text)
                End If
                If Text4.Visible = True Then
                    ![Helper5] = Trim(Text4.Text)
                End If
                ![TotalPerson] = Trim(Combo2.Text) + 1
                ![Perdonels] = Trim(Combo1.Text)
            .Update
        End If
    End With
    
    MsgBox "Data save..", vbInformation, "Saving"
    Combo5.SetFocus
    SendKeys "{HOME}+{END}"
End If

End Sub

Private Sub ButSearch_Click()
    Combo5_Click

End Sub


Private Sub CandyButton1_Click()
If Text1.Text <> "" Then
    OpenPBDataBase ("Cwages")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM Cwages WHERE Customer Like '" & Trim(Text1.Text) & "' ")
    With PRFile
        If Not .EOF Then
            .Edit
                ![Wages] = Trim(Combo9.Text)
                ![Source] = Trim(Combo8.Text)
            .Update
            Text1.SetFocus
            SendKeys "{HOME}+{END}"
        Else
            .AddNew
                ![Customer] = Trim(Text1.Text)
                ![Wages] = Trim(Combo9.Text)
                ![Source] = Trim(Combo8.Text)
            .Update
            Text1.SetFocus
            SendKeys "{HOME}+{END}"
        
        End If
    End With
    
    Call LoadEntries
'    Call PBLOAD(ProgressBar1)
 End If

End Sub


Private Sub CandyButton3_Click()
    If Text1.Text <> "" Then
    OpenPBDataBase ("Cwages")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM Cwages WHERE Customer Like '" & Trim(Text1.Text) & "' and Wages Like '" & Trim(Combo9.Text) & "' ")
    With PRFile
        If Not .EOF Then
                .Delete
                Text1.Text = ""
                Combo9.Text = ""
                Combo8.Text = ""
                Text1.SetFocus
        Else
            MsgBox "Record Not Exist...", vbCritical, "822 MOVERS"
        End If
    End With

    
    End If


    Call LoadEntries
    
End Sub

Sub LOADPlateNumbers()
'On Error Resume Next
    Combo5.Clear
    OpenPBDataBase ("TruckPersonel")
    'Set PRFile = PDbase.OpenRecordset("SELECT * FROM TruckPersonel WHERE Perdonels Like '" & Combo1.Text & "' ")
    With PRFile
     'If Not .EOF Then
      .MoveFirst
        Do While Not .EOF
            Combo5.AddItem ![PlateNumber]
            .MoveNext
        Loop
     'Else
        'Combo4.Clear
     'End If
      .Close
    End With
End Sub

Private Sub CandyButton4_Click()
Call SearchORIGIN
End Sub

Private Sub CandyButton6_Click()
    Call SaveOrigin
    Call SaveTenWheels
    Call SaveForward
    Call SaveELF
    Call Save4W
    MsgBox "Data Save....", vbInformation, "822 MOVERS"
    Call LOADComboOrigin

End Sub

Private Sub CandyButton8_Click()
                Text12.Text = ""
            Text13.Text = ""
            Text14.Text = ""
            Text15.Text = ""
            Text16.Text = ""
            Text17.Text = ""
            Text18.Text = ""
            Text19.Text = ""
            Text20.Text = ""
            Text21.Text = ""
            Text22.Text = ""
            Text23.Text = ""
            Text24.Text = ""
            Text25.Text = ""
            Text26.Text = ""
            Text27.Text = ""
            Text28.Text = ""
            Text29.Text = ""
            
            Text5.Text = ""
            Text6.Text = ""
            Text7.Text = ""
            Text8.Text = ""
            Text9.Text = ""
            Text30.Text = ""
            
            Combo3.SetFocus
            SendKeys "{HOME}+{END}"


End Sub

Private Sub CandyButton9_Click()
    Text1.Text = ""
         Combo9.Text = ""
             Combo8.Text = ""
                Text1.SetFocus
End Sub

Private Sub Combo1_Change()
    Call LOADPlateNumbers
End Sub

Private Sub Combo2_Click()
    If Combo2.Text = 1 Then
        Text2.Visible = True
        Text10.Visible = False
        Text11.Visible = False
        Text3.Visible = False
        Text4.Visible = False
    ElseIf Combo2.Text = 2 Then
        Text2.Visible = True
        Text10.Visible = True
        Text11.Visible = False
        Text3.Visible = False
        Text4.Visible = False
    ElseIf Combo2.Text = 3 Then
        Text2.Visible = True
        Text10.Visible = True
        Text11.Visible = True
        Text3.Visible = False
        Text4.Visible = False
    ElseIf Combo2.Text = 4 Then
        Text2.Visible = True
        Text10.Visible = True
        Text11.Visible = True
        Text3.Visible = True
        Text4.Visible = False
    ElseIf Combo2.Text = 5 Then
        Text2.Visible = True
        Text10.Visible = True
        Text11.Visible = True
        Text3.Visible = True
        Text4.Visible = True
    End If
End Sub

Private Sub Combo3_Click()
    Call SearchORIGIN
End Sub

Private Sub Combo4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
        DelKey = True
        Exit Sub
    End If
End Sub

Private Sub Combo4_Change()
    Call AutoTXTcomplete(MDIMainForm.List4, Combo4)
End Sub

Private Sub Combo9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo8.SetFocus
    End If
End Sub
Private Sub Combo8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CandyButton4_Click
    End If
End Sub

Private Sub ImageList1_Click()

End Sub

Private Sub List1_Click()
        If bNoClick Then Exit Sub
        Text1.Text = List1.Text
End Sub

Private Sub List1_GotFocus()
    SendKeys "{DOWN}"
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call List1_Click
    If KeyCode = vbKeyEscape Then
        List1.Visible = False
        Text1.SetFocus
        SendKeys "{HOME}+{END}"
    End If

End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        Text1.Text = List1.List(List1.ListIndex)
        List1.Visible = False
        Call Text1_KeyPress(13)
   End If
End Sub

Private Sub lvData_Click()
    Text1.Text = lvData.ListItems(lvData.SelectedItem.Index).Text
    Combo9.Text = lvData.ListItems(lvData.SelectedItem.Index).SubItems(1)
    Combo8.Text = lvData.ListItems(lvData.SelectedItem.Index).SubItems(2)
    List1.Visible = False
End Sub

Private Sub lvData_KeyDown(KeyCode As Integer, Shift As Integer)
    Call lvData_Click
End Sub

Private Sub listEntries_Click()
    Text1.Text = listEntries.CellText(listEntries.Row, 0)
    Combo9.Text = listEntries.CellText(listEntries.Row, 1)
    Combo8.Text = listEntries.CellText(listEntries.Row, 2)
End Sub

Private Sub Text1_Change()
    Call AutoTXTcomplete(List1, Text1)
End Sub

Private Sub Text1_Click()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)
    List1.Visible = True

End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        List1.Visible = False
    End If
    If KeyCode = vbKeyDown Then
        List1.Visible = True
        Text1.Text = List1.Text
        List1.SetFocus
    End If
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
        DelKey = True
        Exit Sub
    End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim Xx As Long
    If KeyAscii = 13 Then
    'MsgBox lvData.ListItems.Count
       For Xx = 0 To listEntries.RowCount
            If LCase(Trim(Text1.Text)) = LCase(Trim(listEntries.CellText(Xx, 0))) Then
                 Combo9.Text = listEntries.CellText(Xx, 1)
                 Combo8.Text = listEntries.CellText(Xx, 2)
             List1.Visible = False
             Text1.SetFocus
             SendKeys "{HOME}+{END}"
             Exit Sub
            Else
                Combo9.Text = ""
                Combo8.Text = ""
                Combo9.SetFocus
            End If
       Next Xx
    
        List1.Visible = False
    End If
End Sub
Private Sub text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
        DelKey = True
        Exit Sub
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
End Sub

Private Sub text10_Change()
    Call AutoTXTcomplete(MDIMainForm.List3, Text10)
End Sub
Private Sub text11_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
        DelKey = True
        Exit Sub
    End If
End Sub

Private Sub Text11_Change()
    Call AutoTXTcomplete(MDIMainForm.List3, Text11)
End Sub

Private Sub text3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
        DelKey = True
        Exit Sub
    End If
End Sub

Private Sub Text3_Change()
    Call AutoTXTcomplete(MDIMainForm.List3, Text3)
End Sub
Private Sub text4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
        DelKey = True
        Exit Sub
    End If
End Sub

Private Sub Text4_Change()
    Call AutoTXTcomplete(MDIMainForm.List3, Text4)
End Sub
Private Sub Combo5_Click()
    On Error Resume Next
    OpenPBDataBase ("TruckPersonel")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM TruckPersonel WHERE PlateNumber Like '" & Combo5.Text & "' ")
    With PRFile
        If Not .EOF Then
                Combo4.Text = ![Driver]
                Combo2.Text = Trim(![TotalPerson]) - 1
                Call Combo2_Click
                If Text2.Visible = True Then
                    Text2.Text = ![Helper1]
                End If
                If Text10.Visible = True Then
                    Text10.Text = ![Helper2]
                End If
                If Text11.Visible = True Then
                    Text11.Text = ![Helper3]
                End If
                If Text3.Visible = True Then
                    Text3.Text = ![Helper4]
                End If
                If Text4.Visible = True Then
                    Text4.Text = ![Helper5]
                End If
                
                Combo6.Text = ![Tructype]
                Combo1.Text = ![Perdonels]
        Else
            .Close
        End If
    End With
End Sub
Sub LoadDH()

End Sub
Private Sub Form_Activate()
    'MDIMainForm.JST(2).Expanded = True
    MDIMainForm.ActivateChild Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'MDIMainForm.Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & "NULL.pic")
    'MDIMainForm.JST(2).Expanded = False
    MDIMainForm.RemoveChild Me.Name
End Sub

Private Sub Form_Load()
    With listEntries
        .Redraw = False
        .AddColumn "         Customer", 350   '0
        .AddColumn "         Wages", 230   '1
        .AddColumn "      Source", 80   '2
        
        .Redraw = True
        .Refresh
    End With
    
    Call LoadEntries


    Me.Top = 0
    Me.Left = 0
    Call LOADComboOrigin
    Call LOADPersonelsTypes
    Call LOADTRuckWHeel
    Call LoadEntries
    Call LOADPlateNumbers
    Call loadCustomers
End Sub
Sub loadCustomers()
    'On Error Resume Next
    OpenPBDataBase ("Cwages")
    With PRFile
      .MoveFirst
         Do While Not .EOF
            List1.AddItem ![Customer]
            .MoveNext
         Loop
        .Close
    End With
End Sub
Sub LoadEntries()
    Dim iL As Long
    Dim ec As String
    Dim EN As String
    Dim EO As String

On Error Resume Next

    listEntries.Redraw = False
    listEntries.Clear
    
    OpenPBDataBase ("Cwages")
    With PRFile
      .MoveFirst
        Do While Not .EOF
        If Not .EOF Then
            ec = Trim(![Customer])
            EN = Trim(![Wages])
            EO = Trim(![Source])
            
        With listEntries
            .AddItem (Trim(ec)), 0
            .CellFontBold(iL, 0) = True
            .CellText(iL, 1) = Trim(EN)
            .CellText(iL, 2) = Trim(EO)
            iL = iL + 1
        End With
        .MoveNext
        
        End If
      Loop
      .Close
    End With
    
    'Set vRS = Nothing
    listEntries.Redraw = True
    listEntries.Refresh

End Sub
Sub LOADTRuckWHeel()
    OpenPBDataBase ("TruckWheel")
    With PRFile
       .MoveFirst
        Do While Not .EOF
           Combo6.AddItem ![TruckTypes]
           .MoveNext
        Loop
        .Close
    End With
End Sub

Sub LOADComboOrigin()
    Combo3.Clear
    Combo9.Clear
    OpenPBDataBase ("PointOrigin")
    With PRFile
       .MoveFirst
        Do While Not .EOF
           Combo3.AddItem ![POriginName]
           Combo9.AddItem ![POriginName]
           .MoveNext
        Loop
        .Close
    End With
End Sub

Private Sub KDCButton1_Click()
    Unload Me
    FormMain.Picture1.Visible = False
    FormMain.Show
End Sub

Sub SaveOrigin()
    OpenPBDataBase ("PointOrigin")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM PointOrigin WHERE PoriginName Like '" & Combo3.Text & "' ")
    With PRFile
        If Not .EOF Then
            Exit Sub
        Else
            .AddNew
                ![POriginName] = Trim(Combo3.Text)
            .Update
        End If
        .Close
    End With
End Sub
Sub SaveTenWheels()
'save 10 wheels
    OpenPBDataBase ("SalariesWages")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM SalariesWages WHERE POriginCode Like '" & Combo3.Text & "' and TruckTypes Like '" & Label19.Caption & "' ")
    With PRFile
       If Not .EOF Then
         .Edit
            ![DriverSalary] = Trim(Text12.Text)
            ![Helper1] = Trim(Text13.Text)
            ![Helper2] = Trim(Text14.Text)
            ![Helper3] = Trim(Text15.Text)
            ![Helper4] = Trim(Text16.Text)
            ![Helper5] = Trim(Text17.Text)
        .Update
       Else
        .AddNew
            ![PoriginCode] = Trim(Combo3.Text)
            ![TruckTypes] = Trim(Label19.Caption)
            ![DriverSalary] = Trim(Text12.Text)
            ![Helper1] = Trim(Text13.Text)
            ![Helper2] = Trim(Text14.Text)
            ![Helper3] = Trim(Text15.Text)
            ![Helper4] = Trim(Text16.Text)
            ![Helper5] = Trim(Text17.Text)
        .Update
       End If
      .Close
    End With
End Sub
Sub SaveForward()
    'Save Forward
    OpenPBDataBase ("SalariesWages")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM SalariesWages WHERE POriginCode Like '" & Combo3.Text & "' and TruckTypes Like '" & Label20.Caption & "' ")
    With PRFile
       If Not .EOF Then
         .Edit
            ![DriverSalary] = Trim(Text18.Text)
            ![Helper1] = Trim(Text19.Text)
            ![Helper2] = Trim(Text20.Text)
            ![Helper3] = Trim(Text21.Text)
            ![Helper4] = Trim(Text22.Text)
            ![Helper5] = Trim(Text23.Text)
        .Update
       Else
        .AddNew
            ![PoriginCode] = Trim(Combo3.Text)
            ![TruckTypes] = Trim(Label20.Caption)
            ![DriverSalary] = Trim(Text18.Text)
            ![Helper1] = Trim(Text19.Text)
            ![Helper2] = Trim(Text20.Text)
            ![Helper3] = Trim(Text21.Text)
            ![Helper4] = Trim(Text22.Text)
            ![Helper5] = Trim(Text23.Text)
        .Update
      End If
      .Close
    End With
End Sub
Sub SaveELF()
    'Save to ELF
    OpenPBDataBase ("SalariesWages")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM SalariesWages WHERE POriginCode Like '" & Combo3.Text & "' and TruckTypes Like '" & Label21.Caption & "' ")
    With PRFile
       If Not .EOF Then
         .Edit
            ![DriverSalary] = Trim(Text24.Text)
            ![Helper1] = Trim(Text25.Text)
            ![Helper2] = Trim(Text26.Text)
            ![Helper3] = Trim(Text27.Text)
            ![Helper4] = Trim(Text28.Text)
            ![Helper5] = Trim(Text29.Text)
        .Update
       Else
        .AddNew
            ![PoriginCode] = Trim(Combo3.Text)
            ![TruckTypes] = Trim(Label21.Caption)
            ![DriverSalary] = Trim(Text24.Text)
            ![Helper1] = Trim(Text25.Text)
            ![Helper2] = Trim(Text26.Text)
            ![Helper3] = Trim(Text27.Text)
            ![Helper4] = Trim(Text28.Text)
            ![Helper5] = Trim(Text29.Text)
        .Update
       End If
      .Close
    End With
End Sub
Sub Save4W()
    'Save to ELF
    OpenPBDataBase ("SalariesWages")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM SalariesWages WHERE POriginCode Like '" & Combo3.Text & "' and TruckTypes Like '" & Label12.Caption & "' ")
    With PRFile
       If Not .EOF Then
         .Edit
            ![DriverSalary] = Trim(Text5.Text)
            ![Helper1] = Trim(Text6.Text)
            ![Helper2] = Trim(Text7.Text)
            ![Helper3] = Trim(Text8.Text)
            ![Helper4] = Trim(Text9.Text)
            ![Helper5] = Trim(Text30.Text)
        .Update
       Else
        .AddNew
            ![PoriginCode] = Trim(Combo3.Text)
            ![TruckTypes] = Trim(Label12.Caption)
            ![DriverSalary] = Trim(Text5.Text)
            ![Helper1] = Trim(Text6.Text)
            ![Helper2] = Trim(Text7.Text)
            ![Helper3] = Trim(Text8.Text)
            ![Helper4] = Trim(Text9.Text)
            ![Helper5] = Trim(Text30.Text)
        .Update
       End If
      .Close
    End With
End Sub

Private Sub KDCButton9_Click()
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        List1.AddItem Trim(Text9.Text)
    End If
End Sub
Sub SearchORIGIN()
On Error Resume Next
'Clear All Text
            Text12.Text = ""
            Text13.Text = ""
            Text14.Text = ""
            Text15.Text = ""
            Text16.Text = ""
            Text17.Text = ""
            Text18.Text = ""
            Text19.Text = ""
            Text20.Text = ""
            Text21.Text = ""
            Text22.Text = ""
            Text23.Text = ""
            Text24.Text = ""
            Text25.Text = ""
            Text26.Text = ""
            Text27.Text = ""
            Text28.Text = ""
            Text29.Text = ""
            Text5.Text = ""
            Text6.Text = ""
            Text7.Text = ""
            Text8.Text = ""
            Text9.Text = ""
            Text30.Text = ""

'Search 10 wheels
    OpenPBDataBase ("SalariesWages")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM SalariesWages WHERE POriginCode Like '" & Combo3.Text & "' and TruckTypes Like '" & Label19.Caption & "' ")
    With PRFile
       If Not .EOF Then
            Text12.Text = ![DriverSalary]
            Text13.Text = ![Helper1]
            Text14.Text = ![Helper2]
            Text15.Text = ![Helper3]
            Text16.Text = ![Helper4]
            Text17.Text = ![Helper5]
       End If
      .Close
    End With
    'Search Forward
    OpenPBDataBase ("SalariesWages")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM SalariesWages WHERE POriginCode Like '" & Combo3.Text & "' and TruckTypes Like '" & Label20.Caption & "' ")
    With PRFile
       If Not .EOF Then
            Text18.Text = ![DriverSalary]
            Text19.Text = ![Helper1]
            Text20.Text = ![Helper2]
            Text21.Text = ![Helper3]
            Text22.Text = ![Helper4]
            Text23.Text = ![Helper5]
      End If
      .Close
    End With
    'Search to ELF
    OpenPBDataBase ("SalariesWages")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM SalariesWages WHERE POriginCode Like '" & Combo3.Text & "' and TruckTypes Like '" & Label21.Caption & "' ")
    With PRFile
       If Not .EOF Then
            Text24.Text = ![DriverSalary]
            Text25.Text = ![Helper1]
            Text26.Text = ![Helper2]
            Text27.Text = ![Helper3]
            Text28.Text = ![Helper4]
            Text29.Text = ![Helper5]
       End If
      .Close
    End With
    'Search to 4W
    OpenPBDataBase ("SalariesWages")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM SalariesWages WHERE POriginCode Like '" & Combo3.Text & "' and TruckTypes Like '" & Label12.Caption & "' ")
    With PRFile
       If Not .EOF Then
            Text5.Text = ![DriverSalary]
            Text6.Text = ![Helper1]
            Text7.Text = ![Helper2]
            Text8.Text = ![Helper3]
            Text9.Text = ![Helper4]
            Text30.Text = ![Helper5]
       End If
      .Close
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

