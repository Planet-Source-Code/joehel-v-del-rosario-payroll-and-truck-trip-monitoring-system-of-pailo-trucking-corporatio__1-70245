VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDailyRecord 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Daily Record Entry"
   ClientHeight    =   9645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15315
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   15315
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox ECA 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   285
      ScaleHeight     =   2670
      ScaleWidth      =   5805
      TabIndex        =   64
      Top             =   3270
      Visible         =   0   'False
      Width           =   5835
      Begin VB.TextBox Text12 
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
         Left            =   1305
         TabIndex        =   67
         Top             =   1170
         Width           =   4275
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
         Left            =   1305
         TabIndex        =   66
         Top             =   1665
         Width           =   1845
      End
      Begin VB.TextBox Text10 
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
         Left            =   1305
         TabIndex        =   65
         Top             =   675
         Width           =   3195
      End
      Begin MOVERS.CandyButton CandyButton6 
         Height          =   405
         Left            =   4725
         TabIndex        =   68
         Top             =   2235
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
         Picture         =   "frmDailyRecord.frx":0000
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
      Begin MOVERS.JOELine JOELine11 
         Height          =   30
         Left            =   0
         TabIndex        =   69
         Top             =   360
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   53
         BorderColor1    =   11325655
         BorderColor2    =   16185592
      End
      Begin MOVERS.JOELine JOELine12 
         Height          =   30
         Left            =   15
         TabIndex        =   70
         Top             =   2175
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   53
         BorderColor1    =   11325655
         BorderColor2    =   16185592
      End
      Begin VB.Image imgClose 
         Height          =   360
         Index           =   10
         Left            =   5445
         Picture         =   "frmDailyRecord.frx":077A
         Top             =   -15
         Width           =   360
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   615
         TabIndex        =   74
         Top             =   1200
         Width           =   915
      End
      Begin VB.Image Image13 
         Height          =   285
         Left            =   15
         Picture         =   "frmDailyRecord.frx":0E64
         Stretch         =   -1  'True
         Top             =   30
         Width           =   255
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Cash Advances Entry"
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
         Left            =   315
         TabIndex        =   73
         Top             =   45
         Width           =   3630
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   450
         TabIndex        =   72
         Top             =   1680
         Width           =   1650
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Code:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   270
         TabIndex        =   71
         Top             =   690
         Width           =   1650
      End
      Begin VB.Image imgClose 
         Height          =   360
         Index           =   11
         Left            =   5445
         Picture         =   "frmDailyRecord.frx":15CE
         Top             =   -15
         Width           =   360
      End
   End
   Begin VB.PictureBox TTAE 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   6165
      ScaleHeight     =   2670
      ScaleWidth      =   5805
      TabIndex        =   51
      Top             =   5040
      Visible         =   0   'False
      Width           =   5835
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
         Height          =   1110
         Left            =   1305
         MultiSelect     =   2  'Extended
         TabIndex        =   75
         Top             =   1470
         Visible         =   0   'False
         Width           =   4305
      End
      Begin VB.TextBox Text9 
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
         Left            =   360
         TabIndex        =   62
         Top             =   2295
         Visible         =   0   'False
         Width           =   4110
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
         Left            =   1305
         TabIndex        =   60
         Top             =   1635
         Width           =   1845
      End
      Begin VB.TextBox Text8 
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
         Left            =   1320
         TabIndex        =   53
         Top             =   1170
         Width           =   4275
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "frmDailyRecord.frx":1CB8
         Left            =   1320
         List            =   "frmDailyRecord.frx":1CBA
         TabIndex        =   52
         Top             =   645
         Width           =   1935
      End
      Begin MOVERS.CandyButton CandyButton4 
         Height          =   405
         Left            =   4725
         TabIndex        =   54
         Top             =   2235
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
         Picture         =   "frmDailyRecord.frx":1CBC
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
      Begin MOVERS.JOELine JOELine9 
         Height          =   30
         Left            =   0
         TabIndex        =   55
         Top             =   360
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   53
         BorderColor1    =   11325655
         BorderColor2    =   16185592
      End
      Begin MOVERS.JOELine JOELine10 
         Height          =   30
         Left            =   15
         TabIndex        =   56
         Top             =   2175
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   53
         BorderColor1    =   11325655
         BorderColor2    =   16185592
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "By:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   30
         TabIndex        =   63
         Top             =   2310
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   450
         TabIndex        =   61
         Top             =   1650
         Width           =   1650
      End
      Begin VB.Image imgClose 
         Height          =   360
         Index           =   8
         Left            =   5445
         Picture         =   "frmDailyRecord.frx":2436
         Top             =   -15
         Width           =   360
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Truck Trip Allowance Entry"
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
         Left            =   315
         TabIndex        =   59
         Top             =   45
         Width           =   3030
      End
      Begin VB.Image Image12 
         Height          =   285
         Left            =   15
         Picture         =   "frmDailyRecord.frx":2B20
         Stretch         =   -1  'True
         Top             =   30
         Width           =   255
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   240
         TabIndex        =   58
         Top             =   1200
         Width           =   1650
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Plate #:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   450
         TabIndex        =   57
         Top             =   675
         Width           =   1650
      End
      Begin VB.Image imgClose 
         Height          =   360
         Index           =   9
         Left            =   5445
         Picture         =   "frmDailyRecord.frx":328A
         Top             =   -15
         Width           =   360
      End
   End
   Begin VB.PictureBox TME 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   2460
      ScaleHeight     =   2670
      ScaleWidth      =   5805
      TabIndex        =   28
      Top             =   1920
      Visible         =   0   'False
      Width           =   5835
      Begin VB.ComboBox Combo4 
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
         ItemData        =   "frmDailyRecord.frx":3974
         Left            =   1440
         List            =   "frmDailyRecord.frx":3976
         TabIndex        =   50
         Top             =   555
         Width           =   1935
      End
      Begin VB.TextBox Text6 
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
         Left            =   1455
         TabIndex        =   46
         Top             =   1185
         Width           =   3630
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
         Left            =   1470
         TabIndex        =   45
         Top             =   1770
         Width           =   1845
      End
      Begin MOVERS.CandyButton CandyButton3 
         Height          =   405
         Left            =   4725
         TabIndex        =   29
         Top             =   2235
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
         Picture         =   "frmDailyRecord.frx":3978
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
      Begin MOVERS.JOELine JOELine7 
         Height          =   30
         Left            =   0
         TabIndex        =   30
         Top             =   360
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   53
         BorderColor1    =   11325655
         BorderColor2    =   16185592
      End
      Begin MOVERS.JOELine JOELine8 
         Height          =   30
         Left            =   15
         TabIndex        =   31
         Top             =   2175
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   53
         BorderColor1    =   11325655
         BorderColor2    =   16185592
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Plate #:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   630
         TabIndex        =   49
         Top             =   615
         Width           =   1650
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Items Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   165
         TabIndex        =   48
         Top             =   1200
         Width           =   1650
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   615
         TabIndex        =   47
         Top             =   1785
         Width           =   1650
      End
      Begin VB.Image imgClose 
         Height          =   360
         Index           =   6
         Left            =   5445
         Picture         =   "frmDailyRecord.frx":40F2
         Top             =   0
         Width           =   360
      End
      Begin VB.Image Image11 
         Height          =   285
         Left            =   15
         Picture         =   "frmDailyRecord.frx":47DC
         Stretch         =   -1  'True
         Top             =   30
         Width           =   255
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Truck Maintenance Entry"
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
         Left            =   315
         TabIndex        =   32
         Top             =   45
         Width           =   3030
      End
      Begin VB.Image imgClose 
         Height          =   360
         Index           =   7
         Left            =   5445
         Picture         =   "frmDailyRecord.frx":4F46
         Top             =   0
         Width           =   360
      End
   End
   Begin VB.PictureBox AE 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   3570
      ScaleHeight     =   2670
      ScaleWidth      =   5805
      TabIndex        =   23
      Top             =   1425
      Visible         =   0   'False
      Width           =   5835
      Begin VB.TextBox Text4 
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
         Left            =   1635
         TabIndex        =   42
         Top             =   840
         Width           =   3630
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
         Left            =   1650
         TabIndex        =   41
         Top             =   1440
         Width           =   1845
      End
      Begin MOVERS.CandyButton CandyButton2 
         Height          =   405
         Left            =   4725
         TabIndex        =   24
         Top             =   2235
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
         Picture         =   "frmDailyRecord.frx":5630
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
      Begin MOVERS.JOELine JOELine5 
         Height          =   30
         Left            =   0
         TabIndex        =   25
         Top             =   360
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   53
         BorderColor1    =   11325655
         BorderColor2    =   16185592
      End
      Begin MOVERS.JOELine JOELine6 
         Height          =   30
         Left            =   15
         TabIndex        =   26
         Top             =   2175
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   53
         BorderColor1    =   11325655
         BorderColor2    =   16185592
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Expense Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   105
         TabIndex        =   44
         Top             =   855
         Width           =   1650
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   825
         TabIndex        =   43
         Top             =   1455
         Width           =   1650
      End
      Begin VB.Image Image10 
         Height          =   285
         Left            =   30
         Picture         =   "frmDailyRecord.frx":5DAA
         Stretch         =   -1  'True
         Top             =   15
         Width           =   255
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Administrative Expenses Entry"
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
         Left            =   315
         TabIndex        =   27
         Top             =   45
         Width           =   3510
      End
      Begin VB.Image imgClose 
         Height          =   360
         Index           =   4
         Left            =   5445
         Picture         =   "frmDailyRecord.frx":6514
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgClose 
         Height          =   360
         Index           =   5
         Left            =   5445
         Picture         =   "frmDailyRecord.frx":6BFE
         Top             =   0
         Width           =   360
      End
   End
   Begin VB.PictureBox PE 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   4665
      ScaleHeight     =   2670
      ScaleWidth      =   5805
      TabIndex        =   17
      Top             =   915
      Visible         =   0   'False
      Width           =   5835
      Begin VB.TextBox Text2 
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
         Left            =   1620
         TabIndex        =   38
         Top             =   810
         Width           =   3630
      End
      Begin VB.TextBox Text1 
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
         Left            =   1635
         TabIndex        =   37
         Top             =   1410
         Width           =   1845
      End
      Begin MOVERS.CandyButton CandyButton1 
         Height          =   405
         Left            =   4725
         TabIndex        =   18
         Top             =   2235
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
         Picture         =   "frmDailyRecord.frx":72E8
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
      Begin MOVERS.JOELine JOELine2 
         Height          =   30
         Left            =   0
         TabIndex        =   19
         Top             =   360
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   53
         BorderColor1    =   11325655
         BorderColor2    =   16185592
      End
      Begin MOVERS.JOELine JOELine4 
         Height          =   30
         Left            =   15
         TabIndex        =   22
         Top             =   2175
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   53
         BorderColor1    =   11325655
         BorderColor2    =   16185592
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll For:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   405
         TabIndex        =   40
         Top             =   825
         Width           =   1650
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   765
         TabIndex        =   39
         Top             =   1455
         Width           =   1650
      End
      Begin VB.Image Image9 
         Height          =   285
         Left            =   45
         Picture         =   "frmDailyRecord.frx":7A62
         Stretch         =   -1  'True
         Top             =   45
         Width           =   255
      End
      Begin VB.Image imgClose 
         Height          =   360
         Index           =   3
         Left            =   5445
         Picture         =   "frmDailyRecord.frx":81CC
         Top             =   0
         Width           =   360
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Payroll Entry"
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
         Left            =   315
         TabIndex        =   20
         Top             =   45
         Width           =   1530
      End
      Begin VB.Image imgClose 
         Height          =   360
         Index           =   2
         Left            =   5445
         Picture         =   "frmDailyRecord.frx":88B6
         Top             =   0
         Width           =   360
      End
   End
   Begin VB.PictureBox ARE 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   5250
      ScaleHeight     =   2670
      ScaleWidth      =   5805
      TabIndex        =   13
      Top             =   495
      Visible         =   0   'False
      Width           =   5835
      Begin VB.TextBox TxTAmount 
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
         Left            =   1785
         TabIndex        =   36
         Top             =   1350
         Width           =   1845
      End
      Begin VB.TextBox TxTAccountfrom 
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
         Left            =   1770
         TabIndex        =   35
         Top             =   735
         Width           =   3630
      End
      Begin MOVERS.CandyButton CandyButton5 
         Height          =   405
         Left            =   4725
         TabIndex        =   15
         Top             =   2235
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
         Picture         =   "frmDailyRecord.frx":8FA0
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
      Begin MOVERS.JOELine JOELine1 
         Height          =   30
         Left            =   0
         TabIndex        =   14
         Top             =   360
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   53
         BorderColor1    =   11325655
         BorderColor2    =   16185592
      End
      Begin MOVERS.JOELine JOELine3 
         Height          =   30
         Left            =   0
         TabIndex        =   21
         Top             =   2175
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   53
         BorderColor1    =   11325655
         BorderColor2    =   16185592
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   915
         TabIndex        =   34
         Top             =   1395
         Width           =   1650
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Account from:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   360
         TabIndex        =   33
         Top             =   765
         Width           =   1650
      End
      Begin VB.Image Image8 
         Height          =   300
         Left            =   30
         Picture         =   "frmDailyRecord.frx":971A
         Stretch         =   -1  'True
         Top             =   45
         Width           =   285
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Allowance Entry"
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
         Left            =   315
         TabIndex        =   16
         Top             =   45
         Width           =   1815
      End
      Begin VB.Image imgClose 
         Height          =   360
         Index           =   0
         Left            =   5445
         Picture         =   "frmDailyRecord.frx":FF6C
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgClose 
         Height          =   360
         Index           =   1
         Left            =   5445
         Picture         =   "frmDailyRecord.frx":10656
         Top             =   0
         Width           =   360
      End
   End
   Begin MOVERS.JOETitleBar JOETitleBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   661
      Caption         =   "Daily Record Entry"
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
   Begin MOVERS.LynxGrid3 listAllowance 
      Height          =   2055
      Left            =   60
      TabIndex        =   1
      Top             =   720
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   3625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      BackColorBkg    =   16777215
      BackColorSel    =   8438015
      GridColor       =   11136767
      FocusRectColor  =   33023
      ThemeStyle      =   3
      ColumnSort      =   -1  'True
      Striped         =   -1  'True
      SBackColor2     =   16777215
   End
   Begin MOVERS.LynxGrid3 listPayroll 
      Height          =   1695
      Left            =   60
      TabIndex        =   2
      Top             =   3120
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   2990
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      BackColorBkg    =   16777215
      BackColorSel    =   8438015
      GridColor       =   11136767
      FocusRectColor  =   33023
      ThemeStyle      =   3
      ColumnSort      =   -1  'True
      Striped         =   -1  'True
      SBackColor2     =   16777215
   End
   Begin MOVERS.LynxGrid3 ListAdministrativeEx 
      Height          =   1695
      Left            =   60
      TabIndex        =   3
      Top             =   5160
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   2990
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      BackColorBkg    =   16777215
      BackColorSel    =   8438015
      GridColor       =   11136767
      FocusRectColor  =   33023
      ThemeStyle      =   3
      ColumnSort      =   -1  'True
      Striped         =   -1  'True
      SBackColor2     =   16777215
   End
   Begin MOVERS.LynxGrid3 ListTruckAllowance 
      Height          =   8175
      Left            =   3450
      TabIndex        =   4
      Top             =   705
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   14420
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      BackColorBkg    =   16777215
      BackColorSel    =   8438015
      GridColor       =   11136767
      FocusRectColor  =   33023
      ThemeStyle      =   3
      ColumnSort      =   -1  'True
      Striped         =   -1  'True
      SBackColor2     =   16777215
   End
   Begin MOVERS.LynxGrid3 ListTruckMaintenance 
      Height          =   1695
      Left            =   60
      TabIndex        =   5
      Top             =   7200
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   2990
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      BackColorBkg    =   16777215
      BackColorSel    =   8438015
      GridColor       =   11136767
      FocusRectColor  =   33023
      ThemeStyle      =   3
      ColumnSort      =   -1  'True
      Striped         =   -1  'True
      SBackColor2     =   16777215
   End
   Begin MOVERS.LynxGrid3 ListCashAdvances 
      Height          =   8175
      Left            =   8475
      TabIndex        =   6
      Top             =   735
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   14420
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      BackColorBkg    =   16777215
      BackColorSel    =   8438015
      GridColor       =   11136767
      FocusRectColor  =   33023
      ThemeStyle      =   3
      ColumnSort      =   -1  'True
      Striped         =   -1  'True
      SBackColor2     =   16777215
   End
   Begin MSComctlLib.ImageList imglUser 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmDailyRecord.frx":10D40
            Key             =   "admin"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDailyRecord.frx":112DA
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image7 
      Height          =   285
      Left            =   8475
      Picture         =   "frmDailyRecord.frx":11874
      Stretch         =   -1  'True
      Top             =   435
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CASH Advances"
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
      Left            =   8775
      TabIndex        =   12
      Top             =   480
      Width           =   3645
   End
   Begin VB.Image Image6 
      Height          =   285
      Left            =   3420
      Picture         =   "frmDailyRecord.frx":11FDE
      Stretch         =   -1  'True
      Top             =   420
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Truck Trip Allowance"
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
      Left            =   3720
      TabIndex        =   11
      Top             =   465
      Width           =   4545
   End
   Begin VB.Image Image5 
      Height          =   285
      Left            =   60
      Picture         =   "frmDailyRecord.frx":12748
      Stretch         =   -1  'True
      Top             =   6900
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Truck Maintenance"
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
      Left            =   345
      TabIndex        =   10
      Top             =   6945
      Width           =   2895
   End
   Begin VB.Image Image4 
      Height          =   285
      Left            =   60
      Picture         =   "frmDailyRecord.frx":12EB2
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Administrative Expenses"
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
      Left            =   360
      TabIndex        =   9
      Top             =   4905
      Width           =   2895
   End
   Begin VB.Image Image3 
      Height          =   285
      Left            =   60
      Picture         =   "frmDailyRecord.frx":1361C
      Stretch         =   -1  'True
      Top             =   2820
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Payroll"
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
      Left            =   360
      TabIndex        =   8
      Top             =   2865
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   45
      Picture         =   "frmDailyRecord.frx":13D86
      Stretch         =   -1  'True
      Top             =   420
      Width           =   255
   End
   Begin VB.Label Labelme 
      BackStyle       =   0  'Transparent
      Caption         =   "Allowance Report"
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
      Left            =   345
      TabIndex        =   7
      Top             =   465
      Width           =   2895
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      Height          =   8580
      Left            =   8400
      Top             =   420
      Width           =   4095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   8580
      Left            =   3360
      Top             =   420
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   8580
      Left            =   0
      Top             =   420
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   -255
      Picture         =   "frmDailyRecord.frx":1A5D8
      Stretch         =   -1  'True
      Top             =   9105
      Width           =   15795
   End
End
Attribute VB_Name = "frmDailyRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><>'
'Programmed By: Joehel V. del Rosario    '
'822 MOVERS(PAILO) Trucking Corporation  '
'Date: June 1, 2007                      '
'<><><><><><><><><><><><><><><><><><><><>'

Private Sub CandyButton1_Click()
  If Text1.Text <> "" Then
    
    With listPayroll
        .AddItem (Trim(Text2.Text)), 0
        .CellText(.RowCount - 1, 1) = Format(Val(Text1.Text), "###,###.00")
    End With
    Text2.SetFocus
    SendKeys "{HOME}+{END}"
  Else
    MsgBox "Invalid value..", vbCritical, "Error"
    Text1.SetFocus
    SendKeys "{HOME}+{END}"
  End If
End Sub

Private Sub CandyButton2_Click()
  If Text3.Text <> "" Then
    
    With ListAdministrativeEx
        .AddItem (Trim(Text4.Text)), 0
        .CellText(.RowCount - 1, 1) = Format(Val(Text3.Text), "###,###.00")
    End With
    Text4.SetFocus
    SendKeys "{HOME}+{END}"
  Else
    MsgBox "Invalid value..", vbCritical, "Error"
    Text3.SetFocus
    SendKeys "{HOME}+{END}"
  End If
End Sub

Private Sub CandyButton3_Click()
  If Text5.Text <> "" Then
    
    With ListTruckMaintenance
        .AddItem (Trim(Combo4.Text)), 0
        .CellText(.RowCount - 1, 1) = Trim(Text6.Text)
        .CellText(.RowCount - 1, 2) = Format(Val(Text5.Text), "###,###.00")
    End With
    Combo4.SetFocus
    SendKeys "{HOME}+{END}"
  Else
    MsgBox "Invalid value..", vbCritical, "Error"
    Text5.SetFocus
    SendKeys "{HOME}+{END}"
  End If
End Sub

Private Sub CandyButton4_Click()
  If Text7.Text <> "" Then
    
    With ListTruckAllowance
        '.CheckBoxes = True
        .ImageList = imglUser
        .AddItem (Trim(Combo1.Text)), 0
        .CellText(.RowCount - 1, 1) = Trim(Text8.Text)
        .CellText(.RowCount - 1, 2) = Format(Val(Text7.Text), "###,###.00")
        .CheckBoxes = True
    End With
    Combo1.SetFocus
    SendKeys "{HOME}+{END}"
  Else
    MsgBox "Invalid value..", vbCritical, "Error"
    Text7.SetFocus
    SendKeys "{HOME}+{END}"
  End If
End Sub

Private Sub CandyButton5_Click()
  If TxTAccountfrom.Text <> "" Then
    
    With listAllowance
        .AddItem (Trim(TxTAccountfrom.Text)), 0
        .CellText(.RowCount - 1, 1) = Format(Val(TxTAmount.Text), "###,###.00")
    End With
    TxTAccountfrom.SetFocus
    SendKeys "{HOME}+{END}"
  Else
    MsgBox "Invalid value..", vbCritical, "Error"
    TxTAccountfrom.SetFocus
    SendKeys "{HOME}+{END}"
  End If
End Sub

Private Sub CandyButton6_Click()
  If Text11.Text <> "" Then
    
    With ListCashAdvances
        .AddItem (Trim(Text12.Text)), 0
        .CellText(.RowCount - 1, 1) = Format(Val(Text11.Text), "###,###.00")
    End With
    Text10.SetFocus
    SendKeys "{HOME}+{END}"
  Else
    MsgBox "Invalid value..", vbCritical, "Error"
    Text11.SetFocus
    SendKeys "{HOME}+{END}"
  End If
End Sub


Private Sub Text9_KeyPress(KeyAscii As Integer)
If Text9.Text <> "" Then
    If KeyAscii = 13 Then
        'Open Form Info
            OpenPBDataBase ("EmployeeInfo")
            Set PRFile = PDbase.OpenRecordset("SELECT * FROM EmployeeInfo WHERE ECode Like '" & Text9.Text & "' ")
            With PRFile
                If Not .EOF Then
                    Text9.Text = ![Ename]
                    'ListTruckAllowance.CheckBoxes = True
                    ListTruckAllowance.CellChecked(ListTruckAllowance.Row, 0) = True
               End If
               .Close
            End With
    
        Text9.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End If
End Sub

Private Sub TxTAccountfrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxTAmount.SetFocus
    SendKeys "{HOME}+{END}"
End If
End Sub
Private Sub TxTAmount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call CandyButton5_Click
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1.SetFocus
    SendKeys "{HOME}+{END}"
End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call CandyButton1_Click
End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text3.SetFocus
    SendKeys "{HOME}+{END}"
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call CandyButton2_Click
End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text5.SetFocus
    SendKeys "{HOME}+{END}"
End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call CandyButton3_Click
End If
End Sub
Private Sub Combo4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text6.SetFocus
    SendKeys "{HOME}+{END}"
End If
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text7.SetFocus
    SendKeys "{HOME}+{END}"
End If
End Sub
Private Sub Text8_Change()
     Call AutoTXTcomplete(List1, Text8)
End Sub

Private Sub Text8_Click()
    Text8.SelStart = 0
    Text8.SelLength = Len(Text8)
    List1.Visible = True
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        List1.Visible = False
    End If
    If KeyCode = vbKeyDown Then
        List1.Visible = True
        Text8.Text = List1.Text
        List1.SetFocus
    End If
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
        DelKey = True
        Exit Sub
    End If
End Sub
Private Sub List1_Click()
    If bNoClick Then Exit Sub
    Text8.Text = List1.Text
End Sub

Private Sub List1_GotFocus()
    SendKeys "{DOWN}"
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call List1_Click
    If KeyCode = vbKeyEscape Then
        List1.Visible = False
        Text8.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        Text8.Text = List1.List(List1.ListIndex)
        List1.Visible = False
        Call Text8_KeyPress(13)
   End If
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Text8.Text = List1.List(List1.ListIndex)
    End If
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call CandyButton4_Click
End If
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text8.SetFocus
    SendKeys "{HOME}+{END}"
End If
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        'Open Form Info
            OpenPBDataBase ("EmployeeInfo")
            Set PRFile = PDbase.OpenRecordset("SELECT * FROM EmployeeInfo WHERE ECode Like '" & Text10.Text & "' ")
            With PRFile
                If Not .EOF Then
                    Text12.Text = ![Ename]
               End If
               .Close
            End With
    
    Text11.SetFocus
    SendKeys "{HOME}+{END}"
End If
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call CandyButton6_Click
End If
End Sub
Private Sub Form_Activate()
    'MDIMainForm.JST(2).Expanded = True
    MDIMainForm.ActivateChild Me
    Me.Width = MDIMainForm.Width - 200
End Sub
Private Sub Form_Load()
    Call ListDesignGrid
    
    'Move all picure Entry to Center
    PicAlignment ARE
    PicAlignment PE
    PicAlignment AE
    PicAlignment TME
    PicAlignment TTAE
    PicAlignment ECA
    
    'Load the Plate Numbers
    Call LOADPlateNumbers
    Call loadCustomers
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'MDIMainForm.Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & "NULL.pic")
    'MDIMainForm.JST(2).Expanded = False
    MDIMainForm.RemoveChild Me.Name
End Sub
Private Sub Image1_Click()
    Call listAllowance_DblClick
End Sub
Private Sub Image3_Click()
    Call listPayroll_DblClick
End Sub
Private Sub Image4_Click()
    Call ListAdministrativeEx_DblClick
End Sub

Private Sub Image5_Click()
    Call ListTruckMaintenance_DblClick
End Sub

Private Sub Image6_Click()
        Combo1.Text = ""
        Text8.Text = ""
        Text7.Text = ""
        
        VPics 5
        Label23.Visible = False
        Text9.Visible = False
        Combo1.SetFocus
        SendKeys "{HOME}+{END}"
End Sub

Private Sub Image7_Click()
    Call ListCashAdvances_DblClick
End Sub

Private Sub imgClose_Click(Index As Integer)
    Select Case Index
            Case 1
                ARE.Visible = False
            Case 2
                PE.Visible = False
            Case 5
                AE.Visible = False
            Case 7
                TME.Visible = False
            Case 9
                TTAE.Visible = False
            Case 11
                ECA.Visible = False
    End Select
End Sub
Private Sub imgClose_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
            Case 0
                imgClose(0).Visible = False
            Case 3
                imgClose(3).Visible = False
            Case 4
                imgClose(4).Visible = False
            Case 6
                imgClose(6).Visible = False
            Case 8
                imgClose(8).Visible = False
            Case 10
                imgClose(10).Visible = False
    End Select
End Sub

Private Sub ARE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose(0).Visible = True
End Sub

Private Sub PE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose(3).Visible = True
End Sub
Private Sub AE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose(4).Visible = True
End Sub
Private Sub TME_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose(6).Visible = True
End Sub
Private Sub TTAE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose(8).Visible = True
End Sub
Private Sub ECA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClose(10).Visible = True
End Sub
Private Sub listAllowance_DblClick()
    VPics 1
    TxTAccountfrom.SetFocus
    SendKeys "{HOME}+{END}"
End Sub
Private Sub listPayroll_DblClick()
    VPics 2
    Text2.SetFocus
    SendKeys "{HOME}+{END}"
End Sub
Private Sub ListAdministrativeEx_DblClick()
    VPics 3
    Text4.SetFocus
    SendKeys "{HOME}+{END}"
End Sub
Private Sub ListTruckMaintenance_DblClick()
    VPics 4
    Combo4.SetFocus
    SendKeys "{HOME}+{END}"
End Sub
Private Sub ListTruckAllowance_DblClick()
    If ListTruckAllowance.RowCount <> 0 Then
        Combo1.Text = ""
        Text8.Text = ""
        Text7.Text = ""
        VPics 5
        Combo1.Text = Trim(ListTruckAllowance.CellText(ListTruckAllowance.Row, 0))
        Text8.Text = Trim(ListTruckAllowance.CellText(ListTruckAllowance.Row, 1))
        Text7.Text = Trim(ListTruckAllowance.CellText(ListTruckAllowance.Row, 2))
        Label23.Visible = True
        Text9.Visible = True
        Text9.SetFocus
        SendKeys "{HOME}+{END}"
    Else
        VPics 5
        Label23.Visible = False
        Text9.Visible = False
        Combo1.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub ListTruckAllowance_Click()
    If TTAE.Visible = True Then
        Call ListTruckAllowance_DblClick
    End If
End Sub
Private Sub ListCashAdvances_DblClick()
    VPics 6
    Text10.SetFocus
    SendKeys "{HOME}+{END}"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'------------------------Procedures Call Function-------'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Sub VPics(PicNo As Integer)
    ARE.Visible = False
    PE.Visible = False
    AE.Visible = False
    TME.Visible = False
    TTAE.Visible = False
    ECA.Visible = False
    
    If PicNo = 1 Then: ARE.Visible = True
    If PicNo = 2 Then: PE.Visible = True
    If PicNo = 3 Then: AE.Visible = True
    If PicNo = 4 Then: TME.Visible = True
    If PicNo = 5 Then: TTAE.Visible = True
    If PicNo = 6 Then: ECA.Visible = True
    
    
    
    
End Sub

Sub PicAlignment(PicS As PictureBox)
    PicS.Top = (Me.Height / 2) - (PicS.Height / 2)
    PicS.Left = (Me.Width / 2) - (PicS.Width / 2)
End Sub
Sub LOADPlateNumbers()
On Error Resume Next
Combo4.Clear
    OpenPBDataBase ("TruckPersonel")
    With PRFile
      .MoveFirst
        Do While Not .EOF
            Combo4.AddItem ![PlateNumber]
            Combo1.AddItem ![PlateNumber]
            .MoveNext
        Loop
      .Close
    End With
End Sub
Sub ListDesignGrid()
    'List Allowance
    With listAllowance
        .Redraw = False
            .AddColumn "            Account from", 130   '0
            .AddColumn "     Amount", 78   '1
        .Redraw = True
        .Refresh
    End With
    'List Payroll
    With listPayroll
        .Redraw = False
            .AddColumn "                  Name", 130   '0
            .AddColumn "      Amount", 78   '1
        .Redraw = True
        .Refresh
    End With
    'List Administrative Expenses
    With ListAdministrativeEx
        .Redraw = False
            .AddColumn "          Expense Name", 130   '0
            .AddColumn "      Amount", 78   '1
        .Redraw = True
        .Refresh
    End With
    'List Truck Maintenence
    With ListTruckMaintenance
        .Redraw = False
            '.ImageList = imglUser.ListImages
            .AddColumn " Plate #", 50   '0
            .AddColumn "       Items", 83   '1
            .AddColumn "   Amount", 75   '2
        .Redraw = True
        .Refresh
    End With
    'List Truck Allowance
    With ListTruckAllowance
        .Redraw = False
            '.CheckBoxes = True
            .AddColumn "    Plate #", 60   '0
            .AddColumn "                 Customer", 183   '1
            .AddColumn "       Amount", 78   '2
        .Redraw = True
        .Refresh
    End With
    'List Cash Advances
    With ListCashAdvances
        .Redraw = False
            .AddColumn "                    Name", 180   '0
            .AddColumn "       Amount", 78   '1
        .Redraw = True
        .Refresh
    End With
End Sub
'''''''''''''''''''''''''''''''''''''''''
''Save LIST data Functions'''''''''''''''
'''''''''''''''''''''''''''''''''''''''''
Sub SaveARE()

End Sub
Sub SavePE()

End Sub
Sub SaveAE()

End Sub
Sub SaveTME()

End Sub
Sub SaveTTAE()

End Sub
Sub SaveECA()

End Sub
