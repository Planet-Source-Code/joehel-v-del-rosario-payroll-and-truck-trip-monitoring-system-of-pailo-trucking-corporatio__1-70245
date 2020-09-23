VERSION 5.00
Begin VB.Form extra 
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12810
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   12810
   StartUpPosition =   3  'Windows Default
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
      Begin MOVERS.JOELine JOELine1 
         Height          =   30
         Left            =   15
         TabIndex        =   2
         Top             =   300
         Width           =   15180
         _ExtentX        =   26776
         _ExtentY        =   53
      End
      Begin MOVERS.JOeMenu JOeMenu2 
         Height          =   315
         Left            =   1260
         TabIndex        =   3
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
         ForeColor       =   0
         Caption         =   "Payroll"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Payroll"
         BackColorHover  =   12648447
      End
      Begin MOVERS.JOeMenu JOeMenu3 
         Height          =   315
         Left            =   2550
         TabIndex        =   4
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
         ForeColor       =   0
         Caption         =   "Reports"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Reports"
         BackColorHover  =   12648447
      End
      Begin MOVERS.JOeMenu JOeMenu4 
         Height          =   315
         Left            =   3825
         TabIndex        =   5
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
         ForeColor       =   0
         Caption         =   "Settings"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Settings"
         BackColorHover  =   12648447
      End
      Begin MOVERS.JOeMenu JOeMenu5 
         Height          =   315
         Left            =   5115
         TabIndex        =   6
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
         ForeColor       =   0
         Caption         =   "Help"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Help"
         BackColorHover  =   12648447
      End
      Begin MOVERS.JOeMenu JOeMenu1 
         Height          =   315
         Left            =   -15
         TabIndex        =   7
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
         ForeColor       =   0
         Caption         =   "Transactions"
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Transactions"
         BackColorHover  =   12648447
      End
   End
   Begin MOVERS.LynxGrid3 LynxGrid32 
      Height          =   4965
      Left            =   3570
      TabIndex        =   0
      Top             =   840
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   8758
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
      ColumnSort      =   -1  'True
      Striped         =   -1  'True
      SBackColor2     =   16777215
   End
   Begin MOVERS.JOESBtop JOESBtop1 
      Height          =   945
      Left            =   0
      TabIndex        =   8
      Top             =   615
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   1667
      Begin VB.Image Image4 
         Height          =   525
         Left            =   105
         Picture         =   "extra.frx":0000
         Stretch         =   -1  'True
         Top             =   45
         Width           =   2535
      End
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   0
      Picture         =   "extra.frx":511B
      Stretch         =   -1  'True
      Top             =   945
      Width           =   19995
   End
End
Attribute VB_Name = "extra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
