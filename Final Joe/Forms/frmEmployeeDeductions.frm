VERSION 5.00
Begin VB.Form frmEmployeeDeductions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Employee Deduction"
   ClientHeight    =   9135
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   12120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   609
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   808
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   660
      Left            =   2595
      MaxLength       =   13
      TabIndex        =   25
      Text            =   "0000000000000"
      Top             =   2085
      Width           =   3645
   End
   Begin VB.ComboBox Combo4 
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
      ItemData        =   "frmEmployeeDeductions.frx":0000
      Left            =   2565
      List            =   "frmEmployeeDeductions.frx":0002
      TabIndex        =   23
      Top             =   840
      Width           =   3705
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Paid Deductions"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6975
      TabIndex        =   22
      Top             =   7020
      Width           =   2535
   End
   Begin VB.OptionButton Option9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View NOT-Paid Deductions"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6975
      TabIndex        =   21
      Top             =   6615
      Value           =   -1  'True
      Width           =   2685
   End
   Begin VB.TextBox Text10 
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
      Height          =   285
      Left            =   420
      TabIndex        =   20
      Top             =   6840
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.OptionButton Option8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Others"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   6435
      Width           =   1260
   End
   Begin VB.OptionButton Option7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&SSS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   5715
      Width           =   1185
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&P- Health"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   7
      Top             =   6075
      Width           =   1290
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Loans"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   5370
      Width           =   1230
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Advance"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   4680
      Width           =   1290
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "S&hortage"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   5025
      Width           =   1290
   End
   Begin VB.TextBox Text6 
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
      Height          =   315
      Left            =   2595
      TabIndex        =   11
      Top             =   3540
      Width           =   2205
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   7665
      TabIndex        =   9
      Text            =   ".00"
      Top             =   4110
      Width           =   2205
   End
   Begin VB.TextBox Text7 
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
      Height          =   315
      Left            =   7665
      TabIndex        =   14
      Top             =   3525
      Width           =   2205
   End
   Begin VB.TextBox Text12 
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
      Height          =   315
      Left            =   2580
      TabIndex        =   2
      Top             =   4110
      Width           =   2205
   End
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
      Height          =   315
      Left            =   2580
      TabIndex        =   1
      Top             =   3000
      Width           =   7290
   End
   Begin MOVERS.JOETitleBar JOETitleBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   661
      Caption         =   "Employee Deduction"
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
   Begin MOVERS.LynxGrid3 listEntries 
      Height          =   1665
      Left            =   2550
      TabIndex        =   19
      Top             =   4680
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2937
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
      ColumnSort      =   -1  'True
      Striped         =   -1  'True
      SBackColor1     =   16056319
      SBackColor2     =   14940667
   End
   Begin MOVERS.CandyButton ButSave 
      Height          =   615
      Left            =   9990
      TabIndex        =   10
      Top             =   6720
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   1085
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
      Picture         =   "frmEmployeeDeductions.frx":0004
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
      BackColor       =   &H00808080&
      BorderColor     =   &H00808080&
      Height          =   2490
      Left            =   345
      Top             =   4665
      Width           =   1785
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
      Height          =   270
      Left            =   1020
      TabIndex        =   24
      Top             =   870
      Width           =   1950
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Amount :"
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
      Left            =   6090
      TabIndex        =   18
      Top             =   4125
      Width           =   1575
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Plate Number:"
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
      Left            =   6240
      TabIndex        =   17
      Top             =   3555
      Width           =   1425
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
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
      Height          =   270
      Left            =   705
      TabIndex        =   16
      Top             =   3525
      Width           =   1845
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date :"
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
      Left            =   1410
      TabIndex        =   15
      Top             =   4125
      Width           =   1140
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2145
      Left            =   9960
      Picture         =   "frmEmployeeDeductions.frx":077E
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
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
      Left            =   1845
      TabIndex        =   13
      Top             =   3000
      Width           =   945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Code :"
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
      TabIndex        =   12
      Top             =   2280
      Width           =   1800
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      Height          =   6990
      Left            =   120
      Top             =   600
      Width           =   11895
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   -105
      Picture         =   "frmEmployeeDeductions.frx":1BB8
      Stretch         =   -1  'True
      Top             =   7740
      Width           =   12360
   End
   Begin VB.Menu Mnudel 
      Caption         =   "Del"
      Visible         =   0   'False
      Begin VB.Menu MnuRem 
         Caption         =   "Remove Deduction"
      End
      Begin VB.Menu mnuRef 
         Caption         =   "Resresh List"
      End
   End
End
Attribute VB_Name = "frmEmployeeDeductions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><>'
'Programmed By: Joehel V. del Rosario    '
'822 MOVERS(PAILO) Trucking Corporation  '
'Date: June 1, 2007                      '
'<><><><><><><><><><><><><><><><><><><><>'

Dim TypesD As String

Private Sub Form_Activate()
    Text1.SetFocus
    SENDtxt Text1
    MDIMainForm.ActivateChild Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MDIMainForm.RemoveChild Me.Name
End Sub

Private Sub Form_Load()
    With listEntries
        .Redraw = False
        .AddColumn "           Date ", 100   '0
        .AddColumn "           Amount", 150   '1
        .AddColumn "             Type of Deductions", 233   '2
        '.ImageList = ilList
        .Redraw = True
        .Refresh
    End With
    
    Text12.Text = Format(Now, "MM/DD/YYYY")
    Call LOadcombo4
End Sub
Sub LOadcombo4()
    On Error Resume Next
    Combo4.Clear
    OpenPBDataBase ("DateCover")
    With PRFile
      .MoveFirst
        Do While Not .EOF
            Combo4.AddItem ![CoveredDate]
            
            If ![Status] = "1" Then
                Combo4.Text = ![CoveredDate]
            End If
            .MoveNext
        Loop
   End With
End Sub

Private Sub listEntries_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu Me.Mnudel
    End If
End Sub

Private Sub MnuRem_Click()
    Dim DD As String
    Dim DA As String
    Dim DT As String
    On Error Resume Next
    DD = Trim(listEntries.CellText(listEntries.Row, 0))
    DA = Trim(listEntries.CellText(listEntries.Row, 1))
    DT = Trim(listEntries.CellText(listEntries.Row, 2))
    
        OpenPBDataBase ("DeductionsInfo")
        Set PRFile = PDbase.OpenRecordset("SELECT * FROM DeductionsInfo WHERE DName Like '" & Trim(Text2.Text) & "'  and DateCover Like '" & Trim(Combo4.Text) & "' and  DType Like '" & Trim(DT) & _
         "' and  DDate Like '" & Trim(DD) & "' and  DAmount Like '" & Trim(DA) & "' ")
        With PRFile
            If Not .EOF Then
                .Delete
            End If
            .Close
        End With
    listEntries.RemoveItem (listEntries.Row)
End Sub

Private Sub Option1_Click()
    Text10.Visible = False
    TypesD = "Shortage"
    Text8.SetFocus
    SENDtxt Text8
End Sub

Private Sub Option3_Click()
    Text10.Visible = False
    TypesD = "Advance"
    Text8.SetFocus
    SENDtxt Text8
End Sub

Private Sub Option4_Click()
    Call OpenDeducTionINFO
End Sub

Private Sub Option5_Click()
    Text10.Visible = False
    TypesD = "Loans"
    Text8.SetFocus
    SENDtxt Text8
End Sub

Private Sub Option6_Click()
    Text10.Visible = False
    TypesD = "P-Health"
    Text8.SetFocus
    SENDtxt Text8
End Sub

Private Sub Option7_Click()
    Text10.Visible = False
    TypesD = "SSS"
    Text8.SetFocus
    SENDtxt Text8
End Sub

Private Sub Option8_Click()
If Option8.Value = True Then
    Text10.Visible = True
    Text10.SetFocus
    SENDtxt Text10
End If
End Sub

Private Sub Option9_Click()
    Call OpenDeducTionINFO
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

Private Sub Text1_Change()
    Call AutoTXTcomplete(MDIMainForm.List2, Text1)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error Resume Next
        'Open Form Info
            OpenPBDataBase ("EmployeeInfo")
            Set PRFile = PDbase.OpenRecordset("SELECT * FROM EmployeeInfo WHERE ECode Like '" & Text1.Text & "' ")
            With PRFile
                If Not .EOF Then
                    Text2.Text = ![Ename]
                    Text6.Text = ![EOccupation]
                    Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & Trim(Text1.Text) & ".pic")
               End If
               .Close
            End With
            'Open truk Plate
            OpenPBDataBase ("TruckPersonel")
            Set PRFile = PDbase.OpenRecordset("SELECT * FROM TruckPersonel WHERE Driver Like '" & Trim(Text2.Text) & "' or Helper1 Like '" & Trim(Text2.Text) & _
                                            "' or Helper2 Like '" & Trim(Text2.Text) & "' or Helper3 Like '" & Trim(Text2.Text) & "' or Helper4 Like '" & Trim(Text2.Text) & _
                                            "' or Helper5 Like '" & Trim(Text2.Text) & "'")
            With PRFile
                If Not .EOF Then
                    Text7.Text = ![PlateNumber]
                End If
               .Close
            End With
            'Open DeductionsInfo
            Call OpenDeducTionINFO
            
            
            Text12.SetFocus
            SENDtxt Text12
    End If
End Sub

Private Sub Text10_GotFocus()
    Text10.BackColor = &HC0FFFF
    Text10.FontBold = True
End Sub
Private Sub Text10_LostFocus()
    Text10.BackColor = &HFFFFFF
    Text10.FontBold = False
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TypesD = Trim(Text10.Text)
        Text8.SetFocus
        SENDtxt Text8
        'Text10.Text = ""
        'Text10.Visible = False
    End If
End Sub

Private Sub Text12_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Text2.SetFocus
        SENDtxt Text2
    End If
End Sub

Private Sub Text2_Change()
    Call AutoTXTcomplete(MDIMainForm.List1, Text2)
End Sub

Private Sub text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
        DelKey = True
        Exit Sub
    End If
End Sub
Sub OpenMe()
    Call Text2_KeyPress(13)
End Sub
Sub AddMEdeductions(SHORTamt As Double)
    Call Text2_KeyPress(13)
    Text12.Text = Format(Date, "MM/DD/YYYY")
    'Option8.Value = True
    'Text10.Visible = True
    TypesD = "Short on last Salary"
    'Call Text10_KeyPress(13)
    Text8.Text = Val(SHORTamt)
    'Call Text8_KeyPress(13)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
    'Open Form Info
            OpenPBDataBase ("EmployeeInfo")
            Set PRFile = PDbase.OpenRecordset("SELECT * FROM EmployeeInfo WHERE EName Like '" & Trim(Text2.Text) & "' ")
            With PRFile
                If Not .EOF Then
                    Text1.Text = ![Ecode]
                    Text6.Text = ![EOccupation]
                    Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & "NULL.pic")
                    Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & Trim(Text1.Text) & ".pic")
                    
                End If
               .Close
            End With
            'Open truk Plate
            OpenPBDataBase ("TruckPersonel")
            Set PRFile = PDbase.OpenRecordset("SELECT * FROM TruckPersonel WHERE Driver Like '" & Trim(Text2.Text) & "' or Helper1 Like '" & Trim(Text2.Text) & _
                                            "' or Helper2 Like '" & Trim(Text2.Text) & "' or Helper3 Like '" & Trim(Text2.Text) & "' or Helper4 Like '" & Trim(Text2.Text) & _
                                            "' or Helper5 Like '" & Trim(Text2.Text) & "'")
            With PRFile
                If Not .EOF Then
                    Text7.Text = ![PlateNumber]
                End If
               .Close
            End With
            'Open DeductionsInfo
            Call OpenDeducTionINFO
            
            
            Text12.SetFocus
            SENDtxt Text12
    End If
End Sub
Sub OpenDeducTionINFO()
Dim sts As Integer
Dim DD As String
Dim DA As String
Dim DT As String
Dim iL As Long
On Error Resume Next
If Option9.Value = True Then
    sts = 1
Else
    sts = 0
End If

    'clear list
    listEntries.Redraw = False
    listEntries.Clear
            OpenPBDataBase ("DeductionsInfo")
            Set PRFile = PDbase.OpenRecordset("SELECT * FROM DeductionsInfo WHERE DName Like '" & Trim(Text2.Text) & "' and Status Like '" & Val(sts) & "' and DateCover Like '" & Trim(Combo4.Text) & "' ")
            With PRFile
              .MoveFirst
               Do While Not .EOF
                    If Not .EOF Then
                        DD = Trim(![ddate])
                        DA = Trim(![DAmount])
                        DT = Trim(![DType])
                        
                        With listEntries
                            
                            .AddItem (Trim(DD)), 0
                            '.ItemImage(iL) =
                            .CellAlignment(iL, 0) = lgAlignCenterCenter
                            .CellAlignment(iL, 1) = lgAlignRightCenter
                            .CellAlignment(iL, 2) = lgAlignLeftCenter
                            .CellText(iL, 1) = Trim(DA)
                            .CellText(iL, 2) = Trim(DT)
                            iL = iL + 1
                        End With
                    
                    End If
               .MoveNext
               Loop
                .Close
            End With
    listEntries.Redraw = True
    listEntries.Refresh
End Sub
Private Sub ButSave_Click()
    Call SAVEDEductions
End Sub

Sub SAVEDEductions()
    If Text12.Text <> "" Then
    Option9.Value = True
        'Save Deductions
        OpenPBDataBase ("DeductionsInfo")
        Set PRFile = PDbase.OpenRecordset("SELECT * FROM DeductionsInfo WHERE DName Like '" & Trim(Text2.Text) & "' and  DDate Like '" & Trim(Text12.Text) & _
                                        "' and  DType Like '" & Trim(TypesD) & "'  and  DateCover Like '" & Trim(Combo4.Text) & "' ")
        With PRFile
            If Not .EOF Then
                .Edit
                    ![DAmount] = ![DAmount] + Val(Text8.Text)
                .Update
            Else
                .AddNew
                    ![DName] = Trim(Text2.Text)
                    ![ddate] = Trim(Text12.Text)
                    ![DType] = TypesD
                    ![DAmount] = Val(Text8.Text)
                    ![Status] = "1"
                    ![DateCover] = Trim(Combo4.Text)
                .Update
            End If
            .Close
        End With
        Text12.SetFocus
        SENDtxt Text12
        
        
        Call OpenDeducTionINFO
        Text10.Visible = False
        
        Option1.Value = False
        Option3.Value = False
        Option5.Value = False
        Option6.Value = False
        Option7.Value = False
        Option8.Value = False
        
        Call Option9_Click
  End If
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    MDIMainForm.ShowSideTab
'End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call SAVEDEductions
    End If
End Sub
