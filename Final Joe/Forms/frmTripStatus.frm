VERSION 5.00
Begin VB.Form frmTripStatus 
   BorderStyle     =   0  'None
   Caption         =   "Trip Status"
   ClientHeight    =   9405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13755
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9405
   ScaleWidth      =   13755
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check2 
      Caption         =   "From Billing Entry"
      Height          =   465
      Left            =   10665
      TabIndex        =   26
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "From Trip Entry"
      Height          =   540
      Left            =   10665
      TabIndex        =   25
      Top             =   1755
      Value           =   1  'Checked
      Width           =   1500
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   45
      ScaleHeight     =   6495
      ScaleWidth      =   10620
      TabIndex        =   23
      Top             =   1590
      Width           =   10620
      Begin MOVERS.LynxGrid3 listEntries 
         Height          =   6495
         Left            =   -15
         TabIndex        =   24
         Top             =   15
         Width           =   10080
         _ExtentX        =   17780
         _ExtentY        =   11456
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
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   450
      Left            =   3690
      TabIndex        =   21
      Top             =   9615
      Width           =   795
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   3075
      TabIndex        =   20
      Top             =   9495
      Width           =   2130
   End
   Begin VB.ComboBox Combo7 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmTripStatus.frx":0000
      Left            =   6450
      List            =   "frmTripStatus.frx":0002
      TabIndex        =   8
      Top             =   1035
      Width           =   1380
   End
   Begin VB.ComboBox Combo6 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmTripStatus.frx":0004
      Left            =   4785
      List            =   "frmTripStatus.frx":0065
      TabIndex        =   7
      Top             =   1035
      Width           =   795
   End
   Begin VB.ComboBox Combo5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmTripStatus.frx":00E5
      Left            =   2265
      List            =   "frmTripStatus.frx":010D
      TabIndex        =   6
      Top             =   1050
      Width           =   1545
   End
   Begin VB.ComboBox Combo4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmTripStatus.frx":0141
      Left            =   6450
      List            =   "frmTripStatus.frx":0143
      TabIndex        =   5
      Top             =   615
      Width           =   1380
   End
   Begin VB.ComboBox Combo3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmTripStatus.frx":0145
      Left            =   4785
      List            =   "frmTripStatus.frx":01A6
      TabIndex        =   4
      Top             =   630
      Width           =   795
   End
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmTripStatus.frx":0226
      Left            =   2265
      List            =   "frmTripStatus.frx":024E
      TabIndex        =   3
      Top             =   630
      Width           =   1545
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      Caption         =   "All Trucks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   8235
      TabIndex        =   2
      Top             =   570
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      Caption         =   "Enter Plate No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   8235
      TabIndex        =   1
      Top             =   990
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmTripStatus.frx":0282
      Left            =   10080
      List            =   "frmTripStatus.frx":0284
      TabIndex        =   0
      Top             =   1035
      Visible         =   0   'False
      Width           =   2040
   End
   Begin MOVERS.JOETitleBar JOETitleBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   661
      Caption         =   "Trip Status"
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
   Begin MOVERS.CandyButton ButPrev 
      Height          =   360
      Left            =   10800
      TabIndex        =   18
      Top             =   8175
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
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
      Picture         =   "frmTripStatus.frx":0286
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
      Height          =   360
      Left            =   9180
      TabIndex        =   19
      Top             =   8175
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   635
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
      Picture         =   "frmTripStatus.frx":0A00
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
      Height          =   1020
      Left            =   10695
      TabIndex        =   29
      Top             =   6990
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1799
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Make EXCEL copy"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Picture         =   "frmTripStatus.frx":117A
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
      Height          =   375
      Left            =   7635
      TabIndex        =   30
      Top             =   8175
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
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
      Picture         =   "frmTripStatus.frx":18F4
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
   Begin VB.Label NumberTR 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   10800
      TabIndex        =   28
      Top             =   4530
      Width           =   1245
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Number of TRIPS"
      Height          =   510
      Left            =   10665
      TabIndex        =   27
      Top             =   4170
      Width           =   1485
   End
   Begin VB.Label lblRecSum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Records"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   75
      TabIndex        =   22
      Top             =   8295
      Width           =   720
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   6585
      Left            =   15
      Top             =   1545
      Width           =   12240
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   7665
      Left            =   15
      Top             =   465
      Width           =   12240
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Year :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5880
      TabIndex        =   16
      Top             =   1035
      Width           =   765
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Day :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4275
      TabIndex        =   15
      Top             =   1050
      Width           =   765
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "TO:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   105
      TabIndex        =   14
      Top             =   1050
      Width           =   945
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Month :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1545
      TabIndex        =   13
      Top             =   1050
      Width           =   765
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Year :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5880
      TabIndex        =   12
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Day :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4275
      TabIndex        =   11
      Top             =   615
      Width           =   765
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "FROM:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   10
      Top             =   615
      Width           =   945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Month :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1545
      TabIndex        =   9
      Top             =   630
      Width           =   765
   End
   Begin VB.Image Image2 
      Height          =   465
      Left            =   -15
      Picture         =   "frmTripStatus.frx":206E
      Stretch         =   -1  'True
      Top             =   8145
      Width           =   12720
   End
   Begin VB.Menu menume 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mnuEdte 
         Caption         =   "Edit Trip Date"
      End
      Begin VB.Menu mnuEtEX 
         Caption         =   "Edit Trip Expence"
      End
      Begin VB.Menu mnuDelTRec 
         Caption         =   "Delete Trip Record"
      End
      Begin VB.Menu mnuRef 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "frmTripStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><>'
'Programmed By: Joehel V. del Rosario    '
'822 MOVERS(PAILO) Trucking Corporation  '
'Date: June 1, 2007                      '
'<><><><><><><><><><><><><><><><><><><><>'

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, ByVal lpParameters _
        As String, ByVal lpDirectory As String, ByVal nShowCmd _
        As Long) As Long

Private Sub ButDelete_Click()
Call DELrecsme
End Sub

Private Sub ButPrev_Click()
Dim DC As String
    'frmPreviewTripExpenses
    If Combo2.Text = Combo5.Text And Combo4.Text = Combo7.Text Then
        DC = Combo2.Text & "-" & Combo3.Text & "-" & Combo6.Text & ", " & Combo4.Text
    ElseIf Combo2.Text <> Combo5.Text And Combo4.Text = Combo7.Text Then
        DC = Combo2.Text & "-" & Combo3.Text & "-" & Combo5.Text & " " & Combo6.Text & ", " & Combo4.Text
    Else
        DC = Combo2.Text & "-" & Combo3.Text & ", " & Combo4.Text & "-" & Combo5.Text & "-" & Combo6.Text & ", " & Combo7.Text
    End If
    
    TMPHTxt = Trim(DC) & " --- Total TRIPS: " & NumberTR.Caption
    MDIMainForm.AddChild frmPreviewTripExpenses, True
End Sub

Private Sub ButSearch_Click()
If Combo2.Text <> "" And Combo3.Text <> "" And Combo4.Text <> "" Then
    Call Command1_Click
    Call OPENTrucExpense
End If
End Sub
Sub OPENTrucExpense()
Dim X1 As Integer
Dim trap1 As Integer

NumberTR.Caption = ""

Dim TMPdate As String
Dim TMPplate As String

Dim TMPca As String
Dim TMPgo As String
Dim TMPma As String
Dim TMPtf As String
Dim TMPl As String
Dim TMPlto As String
Dim TMPo As String
Dim TmpTotal As Double
Dim TMPChange As Double


Dim GTMPca As Double
Dim GTMPgo As Double
Dim GTMPma As Double
Dim GTMPtf As Double
Dim GTMPl As Double
Dim GTMPlto As Double
Dim GTMPo As Double
Dim GTmpTotal As Double
Dim GTMPChange As Double

Dim iL As Long

On Error Resume Next

    listEntries.Redraw = False
    listEntries.Clear




    With listEntries
        .Redraw = False
        .Cols = 0
        .Clear
        .AddColumn "        Date", 70   '0
        .AddColumn "    Plate", 60  '1
        .AddColumn " Gate Pass", 100   '2
        .AddColumn "      Custumer", 260   '3
        .AddColumn "    Cases", 60   '4
        .AddColumn " Status", 120    '5
        .Redraw = True
        .Refresh
    End With



'=====================================================================================
If Check1.Value = 1 Then

Dim TMPCus As String
Dim TMPCases As String
Dim TMPtripA As String
Dim TMPTotalE As Double
Dim TMPEchange As Double
Dim TMPsal As Double

Dim TMPDsal As Double
Dim TMPHsal As Double

Dim GTMPtripA As Double
Dim GTMPTotalE As Double
Dim GTMPEchange As Double
Dim GTMPsal As Double

        For X1 = 0 To List1.ListCount - 1
            List1.ListIndex = X1
            
            OpenPBDataBase ("TripInfo")
            If Option1.Value = True Then
                Set PRFile = PDbase.OpenRecordset("SELECT * FROM TripInfo WHERE Tripdate Like '" & Trim(List1.Text) & "' ")
            Else
                Set PRFile = PDbase.OpenRecordset("SELECT * FROM TripInfo WHERE Tripdate Like '" & Trim(List1.Text) & "' and  TruckNumber Like '" & Trim(Combo1.Text) & "' ")
            End If
            
            With PRFile
             .MoveFirst
               Do While Not .EOF
                If Not .EOF Then
                            If Trim(![Particulars]) <> "EMPTY" Then
                                    TMPplate = ![truckNumber]
                                    TMPCus = ![Particulars]
                                    TMPCases = Trim(![Cases])
                                    
                                    TMPtripA = ![ECOde]
                                    
                                    
                                    If trap1 = 0 Then
                                        trap1 = 1
                                        listEntries.AddItem (Trim(List1.Text)), 0
                                    Else
                                        'GRID.Col = 0
                                        listEntries.AddItem (""), 0
                                    End If
                                    '.AddItem (Trim(PD)), 0
                                    listEntries.CellAlignment(iL, 0) = lgAlignLeftCenter
                                    listEntries.CellAlignment(iL, 1) = lgAlignLeftCenter
                                    listEntries.CellAlignment(iL, 2) = lgAlignCenterCenter
                                    listEntries.CellAlignment(iL, 3) = lgAlignLeftCenter
                                    listEntries.CellAlignment(iL, 4) = lgAlignRightCenter
                                    listEntries.CellAlignment(iL, 5) = lgAlignCenterCenter
                                    
                                    listEntries.CellText(iL, 1) = Trim(TMPplate)
                                    listEntries.CellText(iL, 2) = Trim(TMPtripA)
                                    listEntries.CellText(iL, 3) = Trim(TMPCus)
                                    listEntries.CellText(iL, 4) = Trim(TMPCases)
                                    iL = iL + 1
                                    
                                    TMPplate = ""
                                    TMPCus = ""
                                    TMPCases = ""
                                    TMPtripA = ""
                            End If

                End If
                .MoveNext
              Loop
              .Close
           End With
           
           
           trap1 = 0
           
        Next X1
        
        
        
            If listEntries.RowCount >= 22 Then
                listEntries.Width = 10420
            Else
                listEntries.Width = 10125 '10470
            End If

    
    listEntries.AddItem (""), 0

    NumberTR.Caption = listEntries.RowCount - 3
    
    listEntries.Redraw = True
    listEntries.Refresh
    
    RefreshRecSum

    'check trip status
    For iL = 0 To listEntries.RowCount - 2
        If Trim(listEntries.CellText(iL, 2)) <> "" Then
            OpenPBDataBase ("Billings")
            Set PRFile = PDbase.OpenRecordset("SELECT * FROM Billings WHERE Phealth Like '" & Trim(listEntries.CellText(iL, 2)) & "' ")
            With PRFile
                If Not .EOF Then
                    listEntries.CellText(iL, 5) = "Billed in " & ![Status]
                    listEntries.CellForeColor(iL, 5) = vbBlue
                    listEntries.CellBackColor(iL, 5) = vbGreen
                Else
                    listEntries.CellText(iL, 5) = "Not yet Billed"
                    listEntries.CellForeColor(iL, 5) = vbYellow
                    listEntries.CellBackColor(iL, 5) = vbRed
                End If
                .Close
            End With
        Else
            listEntries.CellText(iL, 5) = "No Gate Pass"
            listEntries.CellForeColor(iL, 5) = vbBlue
            listEntries.CellBackColor(iL, 5) = vbYellow
        End If
    Next iL

''=======================================================================================================
Else
        For X1 = 0 To List1.ListCount - 1
            List1.ListIndex = X1
            
            OpenPBDataBase ("Billings")
            If Option1.Value = True Then
                Set PRFile = PDbase.OpenRecordset("SELECT * FROM Billings WHERE Loans Like '" & Trim(List1.Text) & "' ")
            Else
                Set PRFile = PDbase.OpenRecordset("SELECT * FROM Billings WHERE Loans Like '" & Trim(List1.Text) & "' and  SSS Like '" & Trim(Combo1.Text) & "' ")
            End If
            
            With PRFile
             .MoveFirst
               Do While Not .EOF
                If Not .EOF Then
                                    TMPplate = ![SSS]
                                    TMPCus = ![Advances]
                                    TMPCases = Trim(![ECOde])
                                    
                                    TMPtripA = ![Phealth]
                                    
                                    
                                    If trap1 = 0 Then
                                        trap1 = 1
                                        listEntries.AddItem (Trim(List1.Text)), 0
                                    Else
                                        'GRID.Col = 0
                                        listEntries.AddItem (""), 0
                                    End If
                                    '.AddItem (Trim(PD)), 0
                                    listEntries.CellAlignment(iL, 0) = lgAlignLeftCenter
                                    listEntries.CellAlignment(iL, 1) = lgAlignLeftCenter
                                    listEntries.CellAlignment(iL, 2) = lgAlignCenterCenter
                                    listEntries.CellAlignment(iL, 3) = lgAlignLeftCenter
                                    listEntries.CellAlignment(iL, 4) = lgAlignRightCenter
                                    listEntries.CellAlignment(iL, 5) = lgAlignCenterCenter
                                    
                                    listEntries.CellText(iL, 1) = Trim(TMPplate)
                                    listEntries.CellText(iL, 2) = Trim(TMPtripA)
                                    listEntries.CellText(iL, 3) = Trim(TMPCus)
                                    listEntries.CellText(iL, 4) = Trim(TMPCases)
                                    iL = iL + 1
                                    
                                    TMPplate = ""
                                    TMPCus = ""
                                    TMPCases = ""
                                    TMPtripA = ""
                End If
                .MoveNext
              Loop
              .Close
           End With
           
           
           trap1 = 0
           
        Next X1
        
        
        
            If listEntries.RowCount >= 22 Then
                listEntries.Width = 10420
            Else
                listEntries.Width = 10125 '10470
            End If

    
    listEntries.AddItem (""), 0

    NumberTR.Caption = listEntries.RowCount - 3
    
    listEntries.Redraw = True
    listEntries.Refresh
    
    RefreshRecSum

    'check trip status
    For iL = 0 To listEntries.RowCount - 2
        If Trim(listEntries.CellText(iL, 2)) <> "" Then
            OpenPBDataBase ("TripInfo")
            Set PRFile = PDbase.OpenRecordset("SELECT * FROM TripInfo WHERE Ecode Like '" & Trim(listEntries.CellText(iL, 2)) & "' ")
            With PRFile
                If Not .EOF Then
                    listEntries.CellText(iL, 5) = "Trip Found"
                    listEntries.CellForeColor(iL, 5) = vbBlue
                    listEntries.CellBackColor(iL, 5) = vbGreen
                Else
                    listEntries.CellText(iL, 5) = "Trip Not Found"
                    listEntries.CellForeColor(iL, 5) = vbYellow
                    listEntries.CellBackColor(iL, 5) = vbRed
                End If
                .Close
            End With
        Else
            listEntries.CellText(iL, 5) = "No Gate Pass"
            listEntries.CellForeColor(iL, 5) = vbBlue
            listEntries.CellBackColor(iL, 5) = vbYellow
        End If
    Next iL


End If

End Sub
Private Sub RefreshRecSum()
    lblRecSum.Caption = "Record " & listEntries.Row + 1 & " of " & listEntries.RowCount
    'refresh picture
End Sub

Private Sub CandyButton1_Click()
Dim i As Long
Dim n As Long
On Error Resume Next
Set objExcel = GetObject(, "Excel.Application")
If Err.Number Then
   Err.Clear
   Set objExcel = CreateObject("Excel.Application")
   If Err.Number Then
      MsgBox "Can't open Excel."
   End If
End If
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add
'AppActivate "ListGrid To Excel"
For i = 0 To listEntries.RowCount
    For n = 0 To listEntries.Cols
        If i <= 0 Then
            objWorkbook.ActiveSheet.Cells(i, 3).ColumnWidth = 30
            objWorkbook.ActiveSheet.Cells(i + 1, n + 1).Value = listEntries.ColHeading(n)
        Else
            objWorkbook.ActiveSheet.Cells(i, 3).ColumnWidth = 30
            objWorkbook.ActiveSheet.Cells(i + 1, n + 1).Value = listEntries.CellText(i - 1, n)
        End If
    Next
Next
End Sub

Private Sub Check1_Click()
   If Check1.Value = 1 Then
     Check2.Value = 0
   End If
End Sub

Private Sub Check2_Click()
   If Check2.Value = 1 Then
    Check1.Value = 0
   End If
    
End Sub

Private Sub Combo1_Click()
    Call ButSearch_Click
End Sub

Private Sub Combo2_Click()
    Combo5.Text = Combo2.Text
End Sub

Private Sub Combo3_Click()
    If Val(Combo3.Text) <= 25 Then
        Combo6.Text = Val(Combo3.Text) + 6
        
    Else
        Combo6.Text = Val(Val(Combo3.Text) + 7) - 31
        Combo5.Text = Val(Combo2.Text + 1)
    End If
         If Combo6.Text <= 9 Then
            Combo6.Text = "0" & Combo6.Text
         Else
            Combo6.Text = Combo6.Text
         End If
End Sub

Private Sub Combo4_Click()
    Combo7.Text = Combo4.Text
End Sub

Private Sub Command1_Click()
    Dim dd1 As Integer
    Dim dd2 As Integer
    Dim dd3 As Integer
    Dim dd4 As Integer
    Dim trap1 As Integer
    dd1 = Val(Combo3.Text)
    dd2 = Val(Combo6.Text)
    List1.Clear
xxx:
    If dd1 >= 32 Then
        dd1 = 1
        trap1 = 1
    End If
    
    If dd1 = dd2 Then
       If trap1 = 0 Then
         If dd1 <= 9 Then
            List1.AddItem Trim(Combo2.Text) & "/0" & dd1 & "/" & Trim(Combo4.Text)
         Else
            List1.AddItem Trim(Combo2.Text) & "/" & dd1 & "/" & Trim(Combo4.Text)
         End If
       ElseIf trap1 = 1 Then
         If dd1 <= 9 Then
            List1.AddItem Trim(Combo5.Text) & "/0" & dd1 & "/" & Trim(Combo4.Text)
         Else
            List1.AddItem Trim(Combo5.Text) & "/" & dd1 & "/" & Trim(Combo4.Text)
         End If
       End If
        Exit Sub
    Else
       If trap1 = 0 Then
         If dd1 <= 9 Then
            List1.AddItem Trim(Combo2.Text) & "/0" & dd1 & "/" & Trim(Combo4.Text)
         Else
            List1.AddItem Trim(Combo2.Text) & "/" & dd1 & "/" & Trim(Combo4.Text)
         End If
       ElseIf trap1 = 1 Then
         If dd1 <= 9 Then
            List1.AddItem Trim(Combo5.Text) & "/0" & dd1 & "/" & Trim(Combo4.Text)
         Else
            List1.AddItem Trim(Combo5.Text) & "/" & dd1 & "/" & Trim(Combo4.Text)
         End If
       End If
        dd1 = dd1 + 1
        GoTo xxx
    End If
    

End Sub

Private Sub Form_Load()
   Dim a As Double
    
    
    For a = 2006 To 2050
        Combo4.AddItem a
        Combo7.AddItem a
    Next a
End Sub
Private Sub Form_Activate()
    MDIMainForm.ActivateChild Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MDIMainForm.RemoveChild Me.Name
End Sub

Private Sub listEntries_Click()
    RefreshRecSum
End Sub

Private Sub listEntries_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call DELrecsme
    End If
End Sub
Sub DELrecsme()
Dim TMPDte As String '0
        Dim TMPPte As String '1
        Dim TMPCus As String '2
        Dim TMPCases As String '3
        Dim TMPtripA As Double '4
        Dim TMPTotalE As Double '5
        Dim TMPEchange As Double '6
        Dim STRDte As String
        
           On Error Resume Next
            
            If Trim(listEntries.CellText(listEntries.Row, 0)) = "" Then
                STRDte = InputBox("Enter the the exact date to be deleted.  Format: MM/DD/YYYY Example: 11/22/2007 ", "Enter Date")
                TMPDte = Trim(STRDte)
            Else
                TMPDte = Trim(listEntries.CellText(listEntries.Row, 0))
            End If
            
            TMPPte = Trim(listEntries.CellText(listEntries.Row, 1))
            TMPCus = Trim(listEntries.CellText(listEntries.Row, 2))
            TMPCases = Trim(listEntries.CellText(listEntries.Row, 3))
            TMPtripA = Trim(Format(listEntries.CellText(listEntries.Row, 4), "###"))
            TMPTotalE = Trim(Format(listEntries.CellText(listEntries.Row, 5), "###"))
        
        'MsgBox TMPDte & "-" & TMPCus & "-" & TMPCases & "-" & TMPtripA '& "-" & TMPTotalE
        
            OpenPBDataBase ("TripInfo")
            Set PRFile = PDbase.OpenRecordset("SELECT * FROM TripInfo WHERE Tripdate Like '" & Trim(TMPDte) & "' and TruckNumber LIKE '" & Trim(TMPPte) & "' and Particulars LIKE '" & Trim(TMPCus) & "' and Cases LIKE '" & Trim(TMPCases) & "' and TripAmount LIKE '" & Trim(TMPtripA) & "' ") 'and TripConsume LIKE '" & Trim(TMPTotalE) & "' ")
            
            With PRFile
                If Not .EOF Then
                   If MsgBox("Delete this Record?", vbYesNo + vbInformation, "Delete") = vbYes Then
                    .Delete
                    listEntries.RemoveItem (listEntries.Row)
                   Else
                    .Close
                    Exit Sub
                   End If
                Else
                    MsgBox "Record Not Found!", vbExclamation, "Not Found"
                End If
                .Close
            End With
End Sub
Sub EditDATEme()
        Dim TMPDte2 As String
        Dim TMPDte As String '0
        Dim TMPPte As String '1
        Dim TMPCus As String '2
        Dim TMPCases As String '3
        Dim TMPtripA As Double '4
        Dim TMPTotalE As Double '5
        Dim TMPEchange As Double '6
        Dim STRDte As String
        Dim STRDte2 As String
        
           On Error Resume Next
            
            If Trim(listEntries.CellText(listEntries.Row, 0)) = "" Then
                STRDte = InputBox("Enter the the exact date to Edit.  Format: MM/DD/YYYY Example: 11/22/2007 ", "Enter Date")
                TMPDte = Trim(STRDte)
            Else
                TMPDte = Trim(listEntries.CellText(listEntries.Row, 0))
            End If
            
            TMPPte = Trim(listEntries.CellText(listEntries.Row, 1))
            TMPCus = Trim(listEntries.CellText(listEntries.Row, 2))
            TMPCases = Trim(listEntries.CellText(listEntries.Row, 3))
            TMPtripA = Trim(Format(listEntries.CellText(listEntries.Row, 4), "###"))
            TMPTotalE = Trim(Format(listEntries.CellText(listEntries.Row, 5), "###"))
        
        'MsgBox TMPDte & "-" & TMPCus & "-" & TMPCases & "-" & TMPtripA '& "-" & TMPTotalE
        
            OpenPBDataBase ("TripInfo")
            Set PRFile = PDbase.OpenRecordset("SELECT * FROM TripInfo WHERE Tripdate Like '" & Trim(TMPDte) & "' and TruckNumber LIKE '" & Trim(TMPPte) & "' and Particulars LIKE '" & Trim(TMPCus) & "' and Cases LIKE '" & Trim(TMPCases) & "' and TripAmount LIKE '" & Trim(TMPtripA) & "' ") 'and TripConsume LIKE '" & Trim(TMPTotalE) & "' ")
            
            With PRFile
                If Not .EOF Then
                   If MsgBox("Edit this Record?", vbYesNo + vbInformation, "Edit") = vbYes Then
                        STRDte2 = InputBox("Enter the the date.  Format: MM/DD/YYYY Example: 11/22/2007 ", "Enter Date")
                        TMPDte2 = Trim(STRDte2)
                        
                        .Edit
                            ![Tripdate] = Trim(TMPDte2)
                        .Update
                        
                        listEntries.RemoveItem (listEntries.Row)
                   Else
                    .Close
                    Exit Sub
                   End If
                Else
                    MsgBox "Record Not Found!", vbExclamation, "Not Found"
                End If
                .Close
            End With

End Sub
Private Sub listEntries_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu Me.menume
    End If
End Sub

Private Sub listEntries_RowColChanged()
    RefreshRecSum
End Sub

Private Sub mnuDelTRec_Click()
    Call DELrecsme
End Sub

Private Sub mnuEdte_Click()
    Call EditDATEme
    Call ButSearch_Click
End Sub

Private Sub mnuRef_Click()
    Call ButSearch_Click
End Sub

Private Sub Option1_Click()
    Combo1.Visible = False
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then
        Combo1.Visible = True
        Call LOADPlateNumbers
    End If
End Sub
Sub LOADPlateNumbers()
On Error Resume Next
Combo1.Clear
    OpenPBDataBase ("TruckPersonel")
    With PRFile
      .MoveFirst
        Do While Not .EOF
            Combo1.AddItem ![PlateNumber]
            .MoveNext
        Loop
      .Close
    End With
End Sub
