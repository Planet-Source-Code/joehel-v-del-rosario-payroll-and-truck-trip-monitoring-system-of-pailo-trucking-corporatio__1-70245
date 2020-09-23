VERSION 5.00
Begin VB.Form FormPayPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preview Payroll"
   ClientHeight    =   6585
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   11430
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Salary :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5685
      TabIndex        =   37
      Top             =   4200
      Width           =   1890
   End
   Begin VB.Label ts 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   8280
      TabIndex        =   36
      Top             =   4245
      Width           =   2880
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Deductions:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5670
      TabIndex        =   35
      Top             =   4695
      Width           =   2670
   End
   Begin VB.Label td 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   8280
      TabIndex        =   34
      Top             =   4680
      Width           =   2880
   End
   Begin VB.Label Label33 
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
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   5685
      TabIndex        =   33
      Top             =   5430
      Width           =   2370
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
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   8280
      TabIndex        =   32
      Top             =   5475
      Width           =   2880
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000000&
      Height          =   2295
      Left            =   120
      Top             =   3915
      Width           =   5400
   End
   Begin VB.Label ors 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2730
      TabIndex        =   31
      Top             =   5775
      Width           =   1410
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   255
      TabIndex        =   30
      Top             =   5775
      Width           =   2325
   End
   Begin VB.Label ad 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2745
      TabIndex        =   29
      Top             =   5415
      Width           =   1410
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Advances :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   255
      TabIndex        =   28
      Top             =   5415
      Width           =   2325
   End
   Begin VB.Label lo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2745
      TabIndex        =   27
      Top             =   5070
      Width           =   1410
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   270
      TabIndex        =   26
      Top             =   5070
      Width           =   2325
   End
   Begin VB.Label md 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2730
      TabIndex        =   25
      Top             =   4725
      Width           =   1410
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   255
      TabIndex        =   24
      Top             =   4725
      Width           =   2325
   End
   Begin VB.Label ss 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2715
      TabIndex        =   23
      Top             =   4365
      Width           =   1410
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SSS Premium :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   22
      Top             =   4365
      Width           =   2325
   End
   Begin VB.Label wh 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2715
      TabIndex        =   21
      Top             =   4020
      Width           =   1410
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   270
      TabIndex        =   20
      Top             =   4020
      Width           =   2325
   End
   Begin VB.Label ec 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   8415
      TabIndex        =   19
      Top             =   3435
      Width           =   1410
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "E- Cola:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   5940
      TabIndex        =   18
      Top             =   3435
      Width           =   2325
   End
   Begin VB.Label hp 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   8400
      TabIndex        =   17
      Top             =   3090
      Width           =   1410
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
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
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   5925
      TabIndex        =   16
      Top             =   3090
      Width           =   2325
   End
   Begin VB.Label op 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   8385
      TabIndex        =   15
      Top             =   2730
      Width           =   1410
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "OT Pay:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   5910
      TabIndex        =   14
      Top             =   2730
      Width           =   2325
   End
   Begin VB.Label dp 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   8385
      TabIndex        =   13
      Top             =   2385
      Width           =   1410
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Days Pay:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   5910
      TabIndex        =   12
      Top             =   2385
      Width           =   2325
   End
   Begin VB.Label nh 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2700
      TabIndex        =   11
      Top             =   3090
      Width           =   1410
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Holidays:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   225
      TabIndex        =   10
      Top             =   3090
      Width           =   2325
   End
   Begin VB.Label no 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2685
      TabIndex        =   9
      Top             =   2730
      Width           =   1410
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Number of OT Hours:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   210
      TabIndex        =   8
      Top             =   2730
      Width           =   2325
   End
   Begin VB.Label nd 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2685
      TabIndex        =   7
      Top             =   2385
      Width           =   1410
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Days Worked:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   210
      TabIndex        =   6
      Top             =   2385
      Width           =   2325
   End
   Begin VB.Label cd 
      BackStyle       =   0  'Transparent
      Caption         =   "Salary Coverage from"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   225
      TabIndex        =   5
      Top             =   1845
      Width           =   10920
   End
   Begin VB.Label d 
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
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   7860
      TabIndex        =   4
      Top             =   1290
      Width           =   3330
   End
   Begin VB.Label n 
      BackStyle       =   0  'Transparent
      Caption         =   "Del Rosario, Joehel V."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   210
      TabIndex        =   3
      Top             =   1275
      Width           =   7035
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
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   195
      TabIndex        =   2
      Top             =   720
      Width           =   2370
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
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   210
      TabIndex        =   1
      Top             =   465
      Width           =   5685
   End
   Begin VB.Label po 
      BackStyle       =   0  'Transparent
      Caption         =   "PERSONELS OF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   240
      TabIndex        =   0
      Top             =   195
      Width           =   6810
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      Height          =   2205
      Left            =   135
      Top             =   1725
      Width           =   11145
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      Height          =   1140
      Left            =   135
      Top             =   1125
      Width           =   11145
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   6105
      Left            =   135
      Top             =   105
      Width           =   11145
   End
   Begin VB.Menu mnuF 
      Caption         =   "File"
      Begin VB.Menu MnuP 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuC 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "FormPayPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><>'
'Programmed By: Joehel V. del Rosario    '
'822 MOVERS(PAILO) Trucking Corporation  '
'Date: June 1, 2007                      '
'<><><><><><><><><><><><><><><><><><><><>'

Private Sub Form_Load()
    Me.Top = 1455
    Me.Left = 3060
End Sub

Private Sub mnuC_Click()
    Unload Me
End Sub

Private Sub MnuP_Click()
    Me.PrintForm
End Sub
