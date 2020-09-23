VERSION 5.00
Begin VB.Form FormPreviewPayroll 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Preview [Drivers and Employees Payroll]"
   ClientHeight    =   8280
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11430
   FontTransparent =   0   'False
   Icon            =   "FormPreviewPayroll.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   552
   ScaleMode       =   0  'User
   ScaleWidth      =   762
   Begin VB.Label NoteJoe 
      BackStyle       =   0  'Transparent
      Caption         =   "PAALALA: Kung may mga katanungan kayo tungkol sa inyong mga PAYROLL e-text lamang kay JOEHEL. CP #:09233292934, 09215059922"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   300
      TabIndex        =   26
      Top             =   7620
      Width           =   11400
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   975
      Left            =   6360
      Top             =   4485
      Width           =   4830
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME :"
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
      Height          =   300
      Left            =   285
      TabIndex        =   25
      Top             =   1035
      Width           =   825
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "JOEHEL V. DEL ROSARIO"
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
      Height          =   300
      Left            =   1110
      TabIndex        =   24
      Top             =   1035
      Width           =   5940
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE :"
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
      Height          =   300
      Left            =   7020
      TabIndex        =   23
      Top             =   1020
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "JUNE 24-30, 2007"
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
      Height          =   300
      Left            =   7785
      TabIndex        =   22
      Top             =   1035
      Width           =   3765
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
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
      Height          =   300
      Left            =   4950
      TabIndex        =   21
      Top             =   5160
      Width           =   1170
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "OTHERS"
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
      Height          =   300
      Left            =   3285
      TabIndex        =   20
      Top             =   5160
      Width           =   1200
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
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
      Height          =   300
      Left            =   4845
      TabIndex        =   19
      Top             =   4860
      Width           =   1260
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SHORTAGE  "
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
      Height          =   300
      Left            =   3420
      TabIndex        =   18
      Top             =   4845
      Width           =   1230
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
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
      Height          =   300
      Left            =   4920
      TabIndex        =   17
      Top             =   4545
      Width           =   1185
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ADVANCES "
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
      Height          =   300
      Left            =   3285
      TabIndex        =   16
      Top             =   4545
      Width           =   1275
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
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
      Height          =   300
      Left            =   1845
      TabIndex        =   15
      Top             =   5145
      Width           =   1050
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "P-HEALTH :  "
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
      Height          =   300
      Left            =   375
      TabIndex        =   14
      Top             =   5145
      Width           =   1110
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
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
      Height          =   300
      Left            =   1800
      TabIndex        =   13
      Top             =   4830
      Width           =   1095
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SSS "
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
      Height          =   300
      Left            =   330
      TabIndex        =   12
      Top             =   4830
      Width           =   1155
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   ".00"
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
      Height          =   300
      Left            =   1815
      TabIndex        =   11
      Top             =   4530
      Width           =   1080
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00000000&
      Height          =   975
      Left            =   285
      Top             =   4485
      Width           =   3015
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LOANS "
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
      Height          =   300
      Left            =   345
      TabIndex        =   10
      Top             =   4530
      Width           =   1140
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Deductions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   300
      TabIndex        =   9
      Top             =   4230
      Width           =   1260
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00000000&
      Height          =   975
      Left            =   285
      Top             =   4485
      Width           =   6090
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "000,000.00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   8880
      TabIndex        =   8
      Top             =   5115
      Width           =   1920
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL NET PAY "
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
      Height          =   300
      Left            =   6690
      TabIndex        =   7
      Top             =   5175
      Width           =   2010
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "000,000.00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   8940
      TabIndex        =   6
      Top             =   4785
      Width           =   1875
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LESS DEDUCTIONS "
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
      Height          =   300
      Left            =   6480
      TabIndex        =   5
      Top             =   4845
      Width           =   2220
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "000,000.00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   8925
      TabIndex        =   4
      Top             =   4500
      Width           =   1875
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL :"
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
      Height          =   300
      Left            =   7935
      TabIndex        =   3
      Top             =   4590
      Width           =   765
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SALARY OF EMPLOYEE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   210
      TabIndex        =   2
      Top             =   570
      Width           =   10890
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   210
      TabIndex        =   1
      Top             =   360
      Width           =   10890
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Truckers"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   210
      TabIndex        =   0
      Top             =   75
      Width           =   10890
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu MnuPrint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close Preview"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "FormPreviewPayroll"
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
    
    frmEmployeePatroll.listEntrieS1.BackColorSel = vbWhite
    frmEmployeePatroll.listEntrieS1.GridColor = vbWhite
    frmEmployeePatroll.listEntrieS1.BackColorEdit = vbWhite
    frmEmployeePatroll.listEntrieS1.BackColor = vbWhite
    SetParent frmEmployeePatroll.Picture3.hwnd, FormPreviewPayroll.hwnd
    
    frmEmployeePatroll.listEntries.BackColorSel = vbWhite
    frmEmployeePatroll.listEntries.GridColor = vbWhite
    frmEmployeePatroll.listEntries.BackColorEdit = vbWhite
    frmEmployeePatroll.listEntries.BackColor = vbWhite
    SetParent frmEmployeePatroll.Picture2.hwnd, FormPreviewPayroll.hwnd
    
    
    frmEmployeePatroll.Picture3.Top = 367
    frmEmployeePatroll.Picture3.Left = 162
    frmEmployeePatroll.Picture2.Top = 95
    frmEmployeePatroll.Picture2.Left = 17
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMainForm.RemoveChild Me.Name
    
    
    frmEmployeePatroll.listEntrieS1.BackColorSel = &H80C0FF
    frmEmployeePatroll.listEntrieS1.GridColor = &HA9EEFF
    frmEmployeePatroll.listEntrieS1.BackColorEdit = &HC0FFFF
    frmEmployeePatroll.listEntrieS1.BackColor = vbWhite
    SetParent frmEmployeePatroll.Picture3.hwnd, frmEmployeePatroll.Picture1.hwnd
    
    frmEmployeePatroll.listEntries.BackColorSel = &H80C0FF
    frmEmployeePatroll.listEntries.GridColor = &HA9EEFF
    frmEmployeePatroll.listEntrieS1.BackColorEdit = vbWhite
    frmEmployeePatroll.listEntrieS1.BackColor = vbWhite
    SetParent frmEmployeePatroll.Picture2.hwnd, frmEmployeePatroll.hwnd
    
    
    
    frmEmployeePatroll.Picture3.Top = 480
    frmEmployeePatroll.Picture3.Left = 41
    frmEmployeePatroll.Picture2.Top = 202
    frmEmployeePatroll.Picture2.Left = 44
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub MnuPrint_Click()
    'Printer.TwipsPerPixelX
    Me.PrintForm
    On Error Resume Next
    'Paid Payroll status
    OpenPBDataBase ("Payrolls")
    'Set PRFile = PDbase.OpenRecordset("SELECT * FROM Payrolls WHERE ecode Like '" & Trim(FormPayroll.Text1.Text) & "' and Coverdate Like '" & Trim(Label7.Caption) & "' ")
    With PRFile
      .MoveFirst
       Do While Not .EOF
        If Trim(![ECOde]) = Trim(FormPayroll.Text1.Text) And Trim(![coverdate]) = Trim(Label7.Caption) Then
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
    'Set PRFile = PDbase.OpenRecordset("SELECT * FROM DeductionsInfo WHERE Dname Like '" & Trim(FormPayroll.Text1.Text) & "' and Status Like '" & "1" & "' ")
                                     
    With PRFile
       .MoveFirst
           Do While Not .EOF
            If Trim(![DName]) = Trim(FormPayroll.Text1.Text) And Trim(!Status) = "1" Then
                .Edit
                    ![Status] = "0"
                .Update
            End If
             .MoveNext
           Loop
        .Close
    End With
   


End Sub
