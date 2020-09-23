VERSION 5.00
Begin VB.Form FormTRIPprint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Preview [Trip]"
   ClientHeight    =   8535
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   11355
   Icon            =   "FormTRIPprint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text28 
      Appearance      =   0  'Flat
      ForeColor       =   &H00404040&
      Height          =   1770
      Left            =   285
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   5040
      Width           =   2130
   End
   Begin VB.Line Line2 
      X1              =   8355
      X2              =   10350
      Y1              =   6285
      Y2              =   6285
   End
   Begin VB.Line Line1 
      X1              =   8370
      X2              =   10365
      Y1              =   5685
      Y2              =   5685
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "SHORT/OVER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   6225
      TabIndex        =   46
      Top             =   6330
      Width           =   2355
   End
   Begin VB.Label SO 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   8520
      TabIndex        =   45
      Top             =   6345
      Width           =   1425
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "CASH RETURN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   6225
      TabIndex        =   44
      Top             =   6015
      Width           =   2355
   End
   Begin VB.Label CR 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   8520
      TabIndex        =   43
      Top             =   6030
      Width           =   1425
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "BALANCE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   6210
      TabIndex        =   42
      Top             =   5715
      Width           =   2355
   End
   Begin VB.Label BAL 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   8520
      TabIndex        =   41
      Top             =   5730
      Width           =   1425
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL TRIP EXPENCE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   6165
      TabIndex        =   40
      Top             =   5385
      Width           =   2265
   End
   Begin VB.Label TE 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   8520
      TabIndex        =   39
      Top             =   5385
      Width           =   1425
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "CASH ALLOWANCE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   6150
      TabIndex        =   38
      Top             =   5055
      Width           =   1995
   End
   Begin VB.Label TCA 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   8520
      TabIndex        =   37
      Top             =   5055
      Width           =   1425
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL CHARGES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   2580
      TabIndex        =   36
      Top             =   6510
      Width           =   1800
   End
   Begin VB.Label TOC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   4065
      TabIndex        =   35
      Top             =   6495
      Width           =   1425
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "LTO/TMG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   2595
      TabIndex        =   34
      Top             =   6270
      Width           =   1425
   End
   Begin VB.Label LTO 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   4065
      TabIndex        =   33
      Top             =   6270
      Width           =   1425
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "XEROX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   2610
      TabIndex        =   32
      Top             =   6000
      Width           =   1425
   End
   Begin VB.Label X 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   4065
      TabIndex        =   31
      Top             =   6000
      Width           =   1425
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "TOLL FEE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   2610
      TabIndex        =   30
      Top             =   5685
      Width           =   1425
   End
   Begin VB.Label TF 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   4065
      TabIndex        =   29
      Top             =   5685
      Width           =   1425
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "MEAL ALLOW."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   2595
      TabIndex        =   28
      Top             =   5370
      Width           =   1425
   End
   Begin VB.Label MA 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   4050
      TabIndex        =   27
      Top             =   5370
      Width           =   1425
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "GAS AND  OIL :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   2580
      TabIndex        =   26
      Top             =   5040
      Width           =   1425
   End
   Begin VB.Label GO 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   4035
      TabIndex        =   25
      Top             =   5040
      Width           =   1425
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CHARGES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   810
      TabIndex        =   24
      Top             =   4800
      Width           =   1185
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      Height          =   2085
      Left            =   225
      Top             =   4770
      Width           =   5805
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   2085
      Left            =   225
      Top             =   4770
      Width           =   11025
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "NUMBER OF CASES :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   240
      TabIndex        =   22
      Top             =   4245
      Width           =   2235
   End
   Begin VB.Label NC 
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAAAAA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   2340
      TabIndex        =   21
      Top             =   4230
      Width           =   2910
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMERS & ADDRESS :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   225
      TabIndex        =   20
      Top             =   3780
      Width           =   2610
   End
   Begin VB.Label CA 
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAAAAA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   2880
      TabIndex        =   19
      Top             =   3795
      Width           =   8220
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "DELIVERY TYPE :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   6195
      TabIndex        =   18
      Top             =   3300
      Width           =   1740
   End
   Begin VB.Label DT 
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAAAAA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   8025
      TabIndex        =   17
      Top             =   3300
      Width           =   2910
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "POINT OF ORIGIN :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   255
      TabIndex        =   16
      Top             =   3300
      Width           =   1740
   End
   Begin VB.Label PO 
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAAAAA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   2085
      TabIndex        =   15
      Top             =   3300
      Width           =   2910
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   7215
      TabIndex        =   14
      Top             =   1005
      Width           =   1245
   End
   Begin VB.Label DC 
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAAAAA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   8340
      TabIndex        =   13
      Top             =   1005
      Width           =   2955
   End
   Begin VB.Label H5 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   1485
      TabIndex        =   12
      Top             =   2895
      Width           =   3540
   End
   Begin VB.Label H4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1470
      TabIndex        =   11
      Top             =   2565
      Width           =   3540
   End
   Begin VB.Label H3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   1455
      TabIndex        =   10
      Top             =   2250
      Width           =   3540
   End
   Begin VB.Label H2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   1455
      TabIndex        =   9
      Top             =   1935
      Width           =   3540
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "HELPERS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   255
      TabIndex        =   8
      Top             =   1620
      Width           =   1245
   End
   Begin VB.Label H1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   1455
      TabIndex        =   7
      Top             =   1635
      Width           =   3540
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DRIVER :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   1290
      Width           =   1245
   End
   Begin VB.Label DN 
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAAAAA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   1440
      TabIndex        =   5
      Top             =   1305
      Width           =   3540
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   1380
      Left            =   195
      Top             =   3210
      Width           =   11055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "PLATE NO. :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   975
      Width           =   1245
   End
   Begin VB.Label LP 
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAAAAA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   300
      Left            =   1440
      TabIndex        =   3
      Top             =   990
      Width           =   3540
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MOVERS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   25
      TabIndex        =   2
      Top             =   45
      Width           =   11115
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "San Ildefonso Alaminos, Laguna"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   25
      TabIndex        =   1
      Top             =   285
      Width           =   11130
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TRIP TICKET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Left            =   25
      TabIndex        =   0
      Top             =   540
      Width           =   11145
   End
   Begin VB.Menu mnuF 
      Caption         =   "File"
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "FormTRIPprint"
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

    Call LOADALLTXTdata
End Sub
Sub LOADALLTXTdata()
        Label1.Caption = Trim(frmTrips.Combo2.Text)
        LP.Caption = Trim(frmTrips.Combo4.Text)
        DN.Caption = Trim(frmTrips.Text3.Text)
        H1.Caption = Trim(frmTrips.Text1.Text)
        H2.Caption = Trim(frmTrips.Text2.Text)
        H3.Caption = Trim(frmTrips.Text10.Text)
        H4.Caption = Trim(frmTrips.Text4.Text)
        H5.Caption = Trim(frmTrips.Text11.Text)
        
        DC.Caption = Trim(frmTrips.Combo3.Text)
        po.Caption = Trim(frmTrips.Combo5.Text)
        DT.Caption = Trim(frmTrips.Combo10.Text)
        CA.Caption = Trim(frmTrips.Text9.Text)
        NC.Caption = Trim(frmTrips.Text12.Text)
        TCA.Caption = Trim(frmTrips.Text13.Text)
        CR.Caption = Trim(frmTrips.Text5.Text)
        
        
        
        
        GO.Caption = Trim(frmTrips.Text15.Text)
        MA.Caption = Trim(frmTrips.Text16.Text)
        TF.Caption = Trim(frmTrips.Text17.Text)
        X.Caption = Trim(frmTrips.Text18.Text)
        LTO.Caption = Trim(frmTrips.Text19.Text)
        TOC.Caption = Trim(frmTrips.Text20.Text)
        Text28.Text = Trim(frmTrips.Text28.Text)
        
        TE.Caption = Val(Val(GO) + Val(MA) + Val(TF) + Val(X) + Val(LTO) + Val(TOC))
        
        BAL.Caption = Val(TCA.Caption) - Val(TE.Caption)
        SO.Caption = Val(BAL.Caption) - Val(CR.Caption)
        
End Sub

Private Sub mnnuClose_Click()
    Unload Me
End Sub

Private Sub MnuPrint_Click()
    Me.PrintForm
End Sub
