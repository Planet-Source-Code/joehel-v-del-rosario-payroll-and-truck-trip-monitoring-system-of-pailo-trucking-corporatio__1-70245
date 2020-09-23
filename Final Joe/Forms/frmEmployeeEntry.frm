VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEmployeeEntry 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Employee Entry"
   ClientHeight    =   9465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   631
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   813
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cd 
      Left            =   3255
      Top             =   7290
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   5040
      TabIndex        =   28
      Top             =   8685
      Width           =   915
   End
   Begin MOVERS.CandyButton ButSave 
      Height          =   465
      Left            =   5820
      TabIndex        =   11
      Top             =   7230
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
      Picture         =   "frmEmployeeEntry.frx":0000
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
   Begin MOVERS.CandyButton ButIDGen 
      Height          =   675
      Left            =   6225
      TabIndex        =   27
      Top             =   1650
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1191
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "..."
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
      Left            =   2475
      MaxLength       =   13
      TabIndex        =   0
      Text            =   "0000000000000"
      Top             =   1665
      Width           =   3645
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
      Left            =   2490
      TabIndex        =   1
      Top             =   2820
      Width           =   7350
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
      ItemData        =   "frmEmployeeEntry.frx":077A
      Left            =   8085
      List            =   "frmEmployeeEntry.frx":078D
      TabIndex        =   4
      Top             =   4020
      Width           =   1785
   End
   Begin VB.TextBox Text9 
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
      Left            =   2490
      TabIndex        =   7
      Top             =   5310
      Width           =   1785
   End
   Begin VB.TextBox Text5 
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
      Left            =   2490
      TabIndex        =   3
      Top             =   4035
      Width           =   1785
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
      Height          =   360
      ItemData        =   "frmEmployeeEntry.frx":07BE
      Left            =   8085
      List            =   "frmEmployeeEntry.frx":07C8
      TabIndex        =   6
      Top             =   4695
      Width           =   1785
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
      ItemData        =   "frmEmployeeEntry.frx":07DA
      Left            =   8085
      List            =   "frmEmployeeEntry.frx":07F6
      TabIndex        =   8
      Top             =   5355
      Width           =   1785
   End
   Begin VB.TextBox Text4 
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
      Left            =   2490
      TabIndex        =   5
      Top             =   4665
      Width           =   1785
   End
   Begin VB.TextBox Text3 
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
      Left            =   2490
      TabIndex        =   2
      Top             =   3405
      Width           =   7365
   End
   Begin VB.TextBox Text11 
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
      Left            =   2490
      TabIndex        =   10
      Top             =   5940
      Width           =   1785
   End
   Begin MOVERS.JOETitleBar JOETitleBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   661
      Caption         =   "Employee Registration"
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
   Begin MOVERS.CandyButton ButDelete 
      Height          =   465
      Left            =   7485
      TabIndex        =   12
      Top             =   7245
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
      Picture         =   "frmEmployeeEntry.frx":0848
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
      Left            =   9105
      TabIndex        =   13
      Top             =   7245
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
      Picture         =   "frmEmployeeEntry.frx":0FC2
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
      Left            =   10725
      TabIndex        =   14
      Top             =   7245
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
      Picture         =   "frmEmployeeEntry.frx":173C
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      Height          =   855
      Left            =   180
      Top             =   7050
      Width           =   11985
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      Height          =   6600
      Left            =   180
      Top             =   465
      Width           =   11985
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2205
      Left            =   10110
      Picture         =   "frmEmployeeEntry.frx":1EB6
      Stretch         =   -1  'True
      ToolTipText     =   "Click to ADD or CHANGE picture"
      Top             =   585
      Width           =   1995
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   -180
      Picture         =   "frmEmployeeEntry.frx":32F0
      Stretch         =   -1  'True
      Top             =   8070
      Width           =   12765
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Civil Status :"
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
      Left            =   6180
      TabIndex        =   25
      Top             =   4080
      Width           =   1845
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
      Left            =   795
      TabIndex        =   24
      Top             =   1860
      Width           =   1800
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(MM-DD-YY)"
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
      Left            =   4320
      TabIndex        =   23
      Top             =   5340
      Width           =   1245
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date Employed :"
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
      Left            =   615
      TabIndex        =   22
      Top             =   5340
      Width           =   1845
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(MM-DD-YY)"
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
      Left            =   4320
      TabIndex        =   21
      Top             =   4065
      Width           =   1245
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Age :"
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
      Left            =   615
      TabIndex        =   20
      Top             =   4710
      Width           =   1845
   End
   Begin VB.Label Label6 
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
      Left            =   6195
      TabIndex        =   19
      Top             =   5385
      Width           =   1845
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Gender :"
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
      Left            =   6180
      TabIndex        =   18
      Top             =   4725
      Width           =   1845
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Birth Date :"
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
      Left            =   615
      TabIndex        =   17
      Top             =   4065
      Width           =   1845
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
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
      Left            =   1590
      TabIndex        =   16
      Top             =   3420
      Width           =   945
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
      Left            =   1755
      TabIndex        =   15
      Top             =   2850
      Width           =   945
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number :"
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
      Left            =   615
      TabIndex        =   9
      Top             =   5925
      Width           =   1845
   End
End
Attribute VB_Name = "frmEmployeeEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><>'
'Programmed By: Joehel V. del Rosario    '
'822 MOVERS(PAILO) Trucking Corporation  '
'Date: June 1, 2007                      '
'<><><><><><><><><><><><><><><><><><><><>'

Private Sub Command1_Click()
    Dim xxx As Double
    Dim TMPnum As Double
    OpenPBDataBase ("EmployeeInfo")
    TMPnum = 1
    With PRFile
     .MoveFirst
        Do While Not .EOF
            If Not .EOF Then
                .Edit
                    'TMPnum = Format(TMPnum, "00000")
                    ![ECOde] = Format(Now, "MMDDYYYY") & Format(TMPnum, "00000")
                .Update
            End If
            TMPnum = TMPnum + 1
            .MoveNext
        Loop
    End With
End Sub
Private Sub Form_Activate()
    MDIMainForm.ActivateChild Me
    Text1.SetFocus
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MDIMainForm.RemoveChild Me.Name
End Sub

Private Sub ButDelete_Click()
If Text1.Text <> "" And Text2.Text <> "" Then
    OpenPBDataBase ("EmployeeInfo")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM EmployeeInfo WHERE ECode Like '" & Text1.Text & "' and  EName Like '" & Text2.Text & "' ")
    With PRFile
        If Not .EOF Then
            If MsgBox("Want ot delete this Record?", vbYesNo + vbInformation, "Confirm") = vbYes Then
              .Delete
                Text1.Text = "00000"
                Text2.Text = ""
                Text3.Text = ""
                Text5.Text = ""
                Text4.Text = ""
                Text9.Text = ""
                Combo2.Text = ""
                Combo3.Text = ""
                Combo1.Text = ""
            End If
                Text1.SetFocus
                SendKeys "{HOME}+{END}"
        Else
                MsgBox "Employee Code number " & Text1.Text & " not found.", vbCritical, "Record not Found"
                Text1.SetFocus
                SendKeys "{HOME}+{END}"
        End If
        .Close
    End With
    Text1.Text = "0000"
    Text1.Locked = False
 End If
End Sub
Private Sub ButIDGen_Click()
On Error Resume Next
Dim TMPCode As String
Dim Cnumber As Double
    OpenPBDataBase ("EmployeeInfo")
JOEhel:
        Cnumber = Val(Cnumber) + 1
        TMPCode = Format(Now, "MMDDYYYY") & Format(Val(Cnumber), "0000") 'Format(Now, "MMDDYYYY") & Format(TMPnum, "00000")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM EmployeeInfo WHERE ECode Like '" & Trim(TMPCode) & " ' ")
    With PRFile
        If Not .EOF Then
            GoTo JOEhel
        Else
            Text1.Text = TMPCode
            Text1.Text = Format(Text1.Text, "0000")
        End If
        .Close
    End With
    Text2.SetFocus
    SendKeys "{HOME}+{END}"
End Sub

Private Sub ButNew_Click()
    Text1.Text = "0000"
    Text2.Text = ""
    Text3.Text = ""
    Text5.Text = ""
    Text4.Text = ""
    Text9.Text = ""
    Combo2.Text = ""
    Combo3.Text = ""
    Combo1.Text = ""
    Call ButIDGen_Click
    Text1.Locked = True
    Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & "NULL.pic")
    Text2.SetFocus
    SendKeys "{HOME}+{END}"
End Sub

Private Sub ButSave_Click()
If Text2.Text <> "" Then
    OpenPBDataBase ("EmployeeInfo")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM EmployeeInfo WHERE ECode Like '" & Trim(Text1.Text) & "' ") ' and  EName Like '" & Trim(Text2.Text) & "'
    With PRFile
        If Not .EOF Then
            .Edit
                '![Ecode] = Trim(Text1.Text)
                ![Ename] = Trim(Text2.Text)
                ![ContactNumber] = Trim(Text11.Text)
                ![EAddress] = Trim(Text3.Text)
                ![EBdate] = Trim(Text5.Text)
                ![EAge] = Trim(Text4.Text)
                ![EDEmployed] = Trim(Text9.Text)
                ![EOccupation] = Trim(Combo2.Text)
                ![EGender] = Trim(Combo3.Text)
                ![ECivilStatus] = Trim(Combo1.Text)
                SavePicture Image1.Picture, App.Path & "\Database\PICTURES\" & Trim(Text1.Text) & ".pic"
                'If Option1.Value = True Then
                '    ![Status] = "1"
                'Else
                '    ![Status] = "0"
                'End If
            .Update
        Else
            .AddNew
                ![ContactNumber] = Trim(Text11.Text)
                ![ECOde] = Trim(Text1.Text)
                ![Ename] = Trim(Text2.Text)
                ![EAddress] = Trim(Text3.Text)
                ![EBdate] = Trim(Text5.Text)
                ![EAge] = Trim(Text4.Text)
                ![EDEmployed] = Trim(Text9.Text)
                ![EOccupation] = Trim(Combo2.Text)
                ![EGender] = Trim(Combo3.Text)
                ![ECivilStatus] = Trim(Combo1.Text)
                
                SavePicture Image1.Picture, App.Path & "\Database\PICTURES\" & Trim(Text1.Text) & ".pic"
                
                'If Option1.Value = True Then
                '    ![Status] = "1"
                'Else
                '    ![Status] = "0"
                'End If
            .Update
        End If
        .Close
    End With
        frmManageEmployee.LSTREF
    Text1.SetFocus
    SendKeys "{HOME}+{END}"
 End If
End Sub

Private Sub ButSearch_Click()
    Text1.Locked = False
    Text2.Text = ""
    Text3.Text = ""
    Text5.Text = ""
    Text4.Text = ""
    Text9.Text = ""
    Combo2.Text = ""
    Combo3.Text = ""
    Combo1.Text = ""
    Call Text1_KeyPress(13)
    Text1.SetFocus
    SendKeys "{HOME}+{END}"
End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    MDIMainForm.ShowSideTab
'End Sub

Private Sub Image1_Click()
    cd.CancelError = True
    On Error GoTo JmPs
    cd.ShowOpen
    If cd.FileName <> "" Then
        Image1.Picture = LoadPicture(cd.FileName)
    Else
        Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & "NULL.pic")
    End If
    'MsgBsox CD.FileName
JmPs:
End Sub
Sub SearchKey()
On Error Resume Next
    OpenPBDataBase ("EmployeeInfo")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM EmployeeInfo WHERE ECode Like '" & Text1.Text & "' ")
    With PRFile
        If Not .EOF Then
                'Text1.Text = ![Ecode]
                Text2.Text = ![Ename]
                Text3.Text = ![EAddress]
                Text5.Text = ![EBdate]
                Text4.Text = ![EAge]
                Text9.Text = ![EDEmployed]
                Combo2.Text = ![EOccupation]
                Combo3.Text = ![EGender]
                Combo1.Text = ![ECivilStatus]
                Text11.Text = ![ContactNumber]
                Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & Trim(Text1.Text) & ".pic")
                Text2.SetFocus
                SendKeys "{HOME}+{END}"
        Else
                If MsgBox("Employee Code number " & Text1.Text & " not found.. Add new record?", vbYesNo + vbInformation, "Search not Found") = vbYes Then
                    Text2.SetFocus
                    SendKeys "{HOME}+{END}"
                Else
                    Text1.SetFocus
                    SendKeys "{HOME}+{END}"
                End If
        End If
        .Close
    End With
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = 8 Then
        DelKey = True
        Exit Sub
    End If
    
        If KeyCode = vbKeyF12 Then
        Call ButIDGen_Click
    End If


End Sub

Private Sub Text1_Change()
    Call AutoTXTcomplete(MDIMainForm.List2, Text1)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        Call SearchKey
    End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call ButSave_Click
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
    Text2_KeyPress (13)
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error Resume Next
        OpenPBDataBase ("EmployeeInfo")
        Set PRFile = PDbase.OpenRecordset("SELECT * FROM EmployeeInfo WHERE EName Like '" & Trim(Text2.Text) & "' ")
        With PRFile
            If Not .EOF Then
                Text1.Text = ![ECOde]
                'Text2.Text = ![ENAMe]
                Text3.Text = ![EAddress]
                Text5.Text = ![EBdate]
                Text4.Text = ![EAge]
                Text9.Text = ![EDEmployed]
                Combo2.Text = ![EOccupation]
                Combo3.Text = ![EGender]
                Combo1.Text = ![ECivilStatus]
                Text11.Text = ![ContactNumber]
                Text1.SetFocus
                SendKeys "{HOME}+{END}"
                Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & Trim(Text1.Text) & ".pic")
            Else
                Text3.SetFocus
                Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & "NULL.pic")
            End If
        End With
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text5.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo1.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo3.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text4.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Combo2.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub Combo3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text9.SetFocus
        SendKeys "{HOME}+{END}"
    End If
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Text11.SetFocus
    End If
End Sub
Private Sub Text1_GotFocus()
    Text1.BackColor = &H0&
    Text1.FontBold = True
End Sub
Private Sub Text1_LostFocus()
    Text1.BackColor = &H0&
    Text1.FontBold = False
End Sub
''---------------------------------------
Private Sub Text2_GotFocus()
    Text2.BackColor = &HC0FFFF
    Text2.FontBold = True
End Sub
Private Sub Text2_LostFocus()
    Text2.BackColor = &HFFFFFF
    Text2.FontBold = False
End Sub
''---------------------------------------
Private Sub Text3_GotFocus()
    Text3.BackColor = &HC0FFFF
    Text3.FontBold = True
End Sub
Private Sub Text3_LostFocus()
    Text3.BackColor = &HFFFFFF
    Text3.FontBold = False
End Sub
''---------------------------------------
Private Sub Text4_GotFocus()
    Text4.BackColor = &HC0FFFF
    Text4.FontBold = True
End Sub
Private Sub Text4_LostFocus()
    Text4.BackColor = &HFFFFFF
    Text4.FontBold = False
End Sub
''---------------------------------------
Private Sub Text5_GotFocus()
    Text5.BackColor = &HC0FFFF
    Text5.FontBold = True
End Sub
Private Sub Text5_LostFocus()
    Text5.BackColor = &HFFFFFF
    Text5.FontBold = False
End Sub
''---------------------------------------
Private Sub Text6_GotFocus()
    Text6.BackColor = &HC0FFFF
    Text6.FontBold = True
End Sub
Private Sub Text6_LostFocus()
    Text6.BackColor = &HFFFFFF
    Text6.FontBold = False
End Sub
''---------------------------------------
Private Sub Text7_GotFocus()
    Text7.BackColor = &HC0FFFF
    Text7.FontBold = True
End Sub
Private Sub Text7_LostFocus()
    Text7.BackColor = &HFFFFFF
    Text7.FontBold = False
End Sub
''---------------------------------------
Private Sub Text8_GotFocus()
    Text8.BackColor = &HC0FFFF
    Text8.FontBold = True
End Sub
Private Sub Text8_LostFocus()
    Text8.BackColor = &HFFFFFF
    Text8.FontBold = False
End Sub
''---------------------------------------
Private Sub Text9_GotFocus()
    Text9.BackColor = &HC0FFFF
    Text9.FontBold = True
End Sub
Private Sub Text9_LostFocus()
    Text9.BackColor = &HFFFFFF
    Text9.FontBold = False
End Sub
''---------------------------------------
Private Sub Text10_GotFocus()
    Text10.BackColor = &HC0FFFF
    Text10.FontBold = True
End Sub
Private Sub Text10_LostFocus()
    Text10.BackColor = &HFFFFFF
    Text10.FontBold = False
End Sub
''---------------------------------------
Private Sub Text11_GotFocus()
    Text11.BackColor = &HC0FFFF
    Text11.FontBold = True
End Sub
Private Sub Text11_LostFocus()
    Text11.BackColor = &HFFFFFF
    Text11.FontBold = False
End Sub
''---------------------------------------
Private Sub Text12_GotFocus()
    Text12.BackColor = &HC0FFFF
    Text12.FontBold = True
End Sub
Private Sub Text12_LostFocus()
    Text12.BackColor = &HFFFFFF
    Text12.FontBold = False
End Sub
