VERSION 5.00
Begin VB.Form FormCoveredDate 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3015
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6555
   ControlBox      =   0   'False
   Icon            =   "FormCoveredDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   2880
      Left            =   75
      ScaleHeight     =   2820
      ScaleWidth      =   6375
      TabIndex        =   15
      Top             =   60
      Width           =   6435
      Begin MOVERS.CandyButton CandyButton2 
         Height          =   495
         Left            =   3585
         TabIndex        =   21
         Top             =   2295
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Okay"
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
      Begin MOVERS.CandyButton CandyButton3 
         Height          =   480
         Left            =   4845
         TabIndex        =   22
         Top             =   2295
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   847
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Enter New Date"
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
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   6150
         TabIndex        =   18
         Top             =   -15
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "NO DATE SET"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   480
         Left            =   60
         TabIndex        =   17
         Top             =   990
         Width           =   6270
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Date Coverage for Salary"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   480
         Left            =   45
         TabIndex        =   16
         Top             =   120
         Width           =   6225
      End
   End
   Begin VB.ComboBox Combo7 
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
      ItemData        =   "FormCoveredDate.frx":6852
      Left            =   5040
      List            =   "FormCoveredDate.frx":6854
      TabIndex        =   13
      Top             =   1815
      Width           =   1380
   End
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
      ItemData        =   "FormCoveredDate.frx":6856
      Left            =   3465
      List            =   "FormCoveredDate.frx":68B7
      TabIndex        =   11
      Top             =   1830
      Width           =   795
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
      ItemData        =   "FormCoveredDate.frx":692E
      Left            =   1110
      List            =   "FormCoveredDate.frx":6956
      TabIndex        =   8
      Top             =   1875
      Width           =   1545
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
      ItemData        =   "FormCoveredDate.frx":69BC
      Left            =   5025
      List            =   "FormCoveredDate.frx":69BE
      TabIndex        =   6
      Top             =   960
      Width           =   1380
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
      ItemData        =   "FormCoveredDate.frx":69C0
      Left            =   3450
      List            =   "FormCoveredDate.frx":6A21
      TabIndex        =   4
      Top             =   975
      Width           =   795
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
      ItemData        =   "FormCoveredDate.frx":6A98
      Left            =   1095
      List            =   "FormCoveredDate.frx":6AC0
      TabIndex        =   0
      Top             =   1005
      Width           =   1545
   End
   Begin MOVERS.CandyButton CandyButton1 
      Height          =   495
      Left            =   2700
      TabIndex        =   20
      Top             =   2385
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Okay"
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
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   6330
      TabIndex        =   19
      Top             =   15
      Width           =   255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Year :"
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
      Left            =   4425
      TabIndex        =   14
      Top             =   1860
      Width           =   765
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Day :"
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
      Left            =   2835
      TabIndex        =   12
      Top             =   1860
      Width           =   765
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "TO:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   150
      TabIndex        =   10
      Top             =   1575
      Width           =   945
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Month :"
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
      Left            =   420
      TabIndex        =   9
      Top             =   1890
      Width           =   765
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Year :"
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
      Left            =   4410
      TabIndex        =   7
      Top             =   1005
      Width           =   765
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Day :"
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
      Left            =   2820
      TabIndex        =   5
      Top             =   1005
      Width           =   765
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "FROM:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   660
      Width           =   945
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Covered Date for Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   90
      TabIndex        =   2
      Top             =   75
      Width           =   3810
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Month :"
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
      Left            =   405
      TabIndex        =   1
      Top             =   1035
      Width           =   765
   End
End
Attribute VB_Name = "FormCoveredDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><>'
'Programmed By: Joehel V. del Rosario    '
'822 MOVERS(PAILO) Trucking Corporation  '
'Date: June 1, 2007                      '
'<><><><><><><><><><><><><><><><><><><><>'
Dim DC As String

Private Sub CandyButton1_Click()
    If Combo2.Text = Combo5.Text And Combo4.Text = Combo7.Text Then
        DC = Combo2.Text & " " & Combo3.Text & "-" & Combo6.Text & ", " & Combo4.Text
    ElseIf Combo2.Text <> Combo5.Text And Combo4.Text = Combo7.Text Then
        DC = Combo2.Text & " " & Combo3.Text & "-" & Combo5.Text & " " & Combo6.Text & ", " & Combo4.Text
    Else
        DC = Combo2.Text & " " & Combo3.Text & ", " & Combo4.Text & "-" & Combo5.Text & " " & Combo6.Text & ", " & Combo7.Text
    End If
    'MsgBox DC
    Call saveDateC
    Label4.Caption = Trim(DC)
    Picture1.Visible = True
End Sub

Private Sub CandyButton2_Click()
    'FormMain.Show
    'FormMain.Label8.Caption = Label4.Caption
    Unload FormCoveredDate
    MDIMainForm.OpenDATECOver
End Sub

Private Sub CandyButton3_Click()
    Picture1.Visible = False
End Sub

Private Sub Combo2_Click()
    Combo5.Text = Combo2.Text
End Sub
Private Sub Combo3_Click()
    If Val(Combo3.Text) <= 25 Then
        Combo6.Text = Val(Combo3.Text) + 6
    Else
        Combo6.Text = Val(Val(Combo3.Text) + 7) - 31
    End If
End Sub
Private Sub Combo4_Click()
    Combo7.Text = Combo4.Text
End Sub

Private Sub Form_Load()
    Dim a As Double
    'If Format(Date, "MM/DD/YYYY") = "09/25/2007" _
    'Or Format(Date, "MM/DD/YYYY") = "09/26/2007" _
    'Or Format(Date, "MM/DD/YYYY") = "09/27/2007" _
    'Or Format(Date, "MM/DD/YYYY") = "09/28/2007" _
    'Or Format(Date, "MM/DD/YYYY") = "09/29/2007" _
    'Or Format(Date, "MM/DD/YYYY") = "09/30/2007" Then
    '    MsgBox "PROGRAM IS LOCK.... Please contact the system programmer", vbCritical, "LOCK"
    '    FormSEARCHtrip.Show
    '    Unload Me
    '    Exit Sub
    'End If
    
    Call OpenDATECOver
    
    For a = 2006 To 2050
        Combo4.AddItem a
        Combo7.AddItem a
    Next a
    
    
    
End Sub
Sub OpenDATECOver()
On Error Resume Next
If Label4.Caption <> "" Then
    OpenPBDataBase ("DateCover")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM DateCover WHERE Status Like '" & "1" & "' ")
    With PRFile
        If Not .EOF Then
            Label4.Caption = ![CoveredDate]
        Else
            Label4.Caption = "Date NOT SET"
        End If
   End With
 End If
End Sub
Sub saveDateC()
On Error Resume Next
    OpenPBDataBase ("DateCover")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM DateCover WHERE Status Like '" & "1" & "' ")
    With PRFile
        If Not .EOF Then
        .MoveFirst
            Do While Not .EOF
              If ![Status] = "1" Then
                .Edit
                    ![Status] = "0"
                .Update
              End If
              If ![CoveredDate] = Trim(DC) Then
                Exit Sub
              End If
               .MoveNext
            Loop
            GoTo JMP
        Else
JMP:
            .AddNew
                ![CoveredDate] = Trim(DC)
                ![Status] = "1"
            .Update
            
        End If
   End With
End Sub

Private Sub Label12_Click()
    Unload Me
End Sub

Private Sub Label13_Click()
    Unload Me
End Sub
