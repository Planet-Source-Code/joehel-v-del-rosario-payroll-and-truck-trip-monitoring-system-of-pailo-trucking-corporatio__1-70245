VERSION 5.00
Begin VB.Form frmPayrollReport 
   BorderStyle     =   0  'None
   Caption         =   "Payroll Reports"
   ClientHeight    =   9270
   ClientLeft      =   150
   ClientTop       =   75
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Breakdown"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2385
      Left            =   9840
      TabIndex        =   12
      Top             =   5625
      Width           =   2295
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   1050
         TabIndex        =   22
         Top             =   1575
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "SHORT            :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   60
         TabIndex        =   21
         Top             =   1575
         Width           =   1320
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   1050
         TabIndex        =   20
         Top             =   1215
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "SUB TOTAL    :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   45
         TabIndex        =   19
         Top             =   1215
         Width           =   1260
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "SALARY         :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   45
         TabIndex        =   18
         Top             =   465
         Width           =   1200
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "DEDUCTIONS :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   45
         TabIndex        =   17
         Top             =   810
         Width           =   1215
      End
      Begin VB.Label TS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1050
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label TD 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1050
         TabIndex        =   15
         Top             =   810
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL             :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   45
         TabIndex        =   14
         Top             =   1995
         Width           =   1350
      End
      Begin VB.Label NT 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   1050
         TabIndex        =   13
         Top             =   2010
         Width           =   1215
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   -15
         X2              =   2400
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   -15
         X2              =   2385
         Y1              =   1125
         Y2              =   1125
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Other Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      Left            =   9840
      TabIndex        =   9
      Top             =   3150
      Width           =   2310
      Begin MOVERS.CandyButton ButPrev 
         Height          =   435
         Left            =   120
         TabIndex        =   10
         Top             =   1605
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   767
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "    Print Preview"
         IconHighLiteColor=   0
         CaptionHighLiteColor=   0
         Picture         =   "frmPayrollReport.frx":0000
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
         Left            =   120
         TabIndex        =   11
         Top             =   420
         Width           =   2100
         _ExtentX        =   3704
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
         Picture         =   "frmPayrollReport.frx":077A
         Checked         =   0   'False
         ColorButtonHover=   16760976
         ColorButtonUp   =   15309136
         ColorButtonDown =   15309136
         BorderBrightness=   0
         ColorBright     =   16772528
         DisplayHand     =   0   'False
         ColorScheme     =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "View Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2070
      Left            =   9840
      TabIndex        =   5
      Top             =   900
      Width           =   2310
      Begin VB.OptionButton Option1 
         Caption         =   "View all Records"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   45
         TabIndex        =   8
         Top             =   210
         Value           =   -1  'True
         Width           =   1680
      End
      Begin VB.OptionButton Option2 
         Caption         =   "View with Salary Only"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   45
         TabIndex        =   7
         Top             =   765
         Width           =   2205
      End
      Begin VB.OptionButton Option3 
         Caption         =   "View Short Salary Only"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   45
         TabIndex        =   6
         Top             =   1350
         Width           =   2190
      End
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
      ItemData        =   "frmPayrollReport.frx":0EF4
      Left            =   1620
      List            =   "frmPayrollReport.frx":0EF6
      TabIndex        =   3
      Top             =   510
      Width           =   8025
   End
   Begin MOVERS.JOETitleBar JOETitleBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   661
      Caption         =   "Payroll Reports"
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
      Height          =   7020
      Left            =   105
      TabIndex        =   1
      Top             =   990
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   12383
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
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      Height          =   7665
      Left            =   9720
      Top             =   405
      Width           =   2520
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   7665
      Left            =   30
      Top             =   405
      Width           =   9705
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Covered :"
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
      Height          =   270
      Left            =   120
      TabIndex        =   4
      Top             =   525
      Width           =   1950
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
      Left            =   45
      TabIndex        =   2
      Top             =   8175
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   -15
      Picture         =   "frmPayrollReport.frx":0EF8
      Stretch         =   -1  'True
      Top             =   8130
      Width           =   12360
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuVP 
         Caption         =   "View Payroll"
      End
      Begin VB.Menu mnuVD 
         Caption         =   "View Deductions"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSD 
         Caption         =   "Add Short to Deductions"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmPayrollReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><>'
'Programmed By: Joehel V. del Rosario    '
'822 MOVERS(PAILO) Trucking Corporation  '
'Date: June 1, 2007                      '
'<><><><><><><><><><><><><><><><><><><><>'


Dim AllView As Boolean
Private Sub LoadEntries()
    Dim iL As Long
    Dim ec As String
    Dim EN As String
    Dim EO As String
    Dim eB As String
    Dim EA As String
    
    'clear list
    
    MDIMainForm.MousePointer = vbHourglass
    
    listEntries.Redraw = False
    listEntries.Clear
    
    On Error Resume Next
    
    'Open for personnels ID and name
    OpenPBDataBase ("EmployeeInfo")
    With PRFile
    .MoveFirst
      Do While Not .EOF
        If Not .EOF Then
            'ec = Trim(![ECOde])
            EN = Trim(![Ename])
            
        With listEntries
            .AddItem (Trim(ec)), 0 'Trim(EC)
            .CellFontBold(iL, 1) = True
            .CellText(iL, 1) = Trim(EN)
            iL = iL + 1
        End With
        
        End If
       .MoveNext
      Loop
      .Close
    End With
     
    listEntries.Redraw = True
    listEntries.Refresh
    RefreshRecSum
    Call ViewALLRecords
    
   
End Sub
Private Sub RefreshRecSum()
    EPRepR = listEntries.Row
    lblRecSum.Caption = "Record " & listEntries.Row + 1 & " of " & listEntries.RowCount
End Sub

Private Sub ButPrev_Click()
    frmPayrollReport.MousePointer = vbHourglass
    Call ADDrowTotal
    frmPayrollReport.MousePointer = vbHourglass
    MDIMainForm.AddChild frmPreviewPayroll, True
    
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
            objWorkbook.ActiveSheet.Cells(i, 2).ColumnWidth = 30
            objWorkbook.ActiveSheet.Cells(i + 1, n + 1).Value = listEntries.ColHeading(n)
        Else
            objWorkbook.ActiveSheet.Cells(i, 2).ColumnWidth = 30
            objWorkbook.ActiveSheet.Cells(i + 1, n + 1).Value = listEntries.CellText(i - 1, n)
        End If
    Next
Next
End Sub

Private Sub Combo4_Click()
   Call LoadEntries
   If Option1.Value = True Then
        Call ViewALLrecordsME
   ElseIf Option2.Value = True Then
        Call ViewWSalary
   ElseIf Option3.Value = True Then
        Call ViewShSalary
   End If
End Sub
Sub ViewALLrecordsME()
    Dim r As Long
r = 0
BackME:
        With listEntries
            If Trim(.CellText(r, 2)) = ".00" And Trim(.CellText(r, 3)) = ".00" And Trim(.CellText(r, 4)) = ".00" Then
                .RemoveItem (r)
                r = r - 1
            End If
        End With
    r = r + 1
    If r <> listEntries.RowCount Then
        GoTo BackME
    End If
    listEntries.Refresh
    Call RefreshRecSum
     MDIMainForm.MousePointer = vbNormal
End Sub
Sub ViewWSalary()
    Dim r As Long
r = 0
BackME:
        With listEntries
            If Val(Format(.CellText(r, 4), "###.00")) <= 0 Then
                .RemoveItem (r)
                r = r - 1
            End If
        End With
    r = r + 1
    If r <> listEntries.RowCount Then
        GoTo BackME
    End If
    listEntries.Refresh
    Call RefreshRecSum
     MDIMainForm.MousePointer = vbNormal
End Sub
Sub ViewShSalary()
    Dim r As Long
r = 0
BackME:
        With listEntries
            If Val(Format(.CellText(r, 4), "###.00")) >= Val(0) Then
                .RemoveItem (r)
                r = r - 1
            ElseIf Val(Format(.CellText(r, 4), ".00")) = ".00" Then
                .RemoveItem (r)
                r = r - 1
            End If
        End With
    r = r + 1
    If r <> listEntries.RowCount Then
        GoTo BackME
    End If
    
    listEntries.Refresh
    Call RefreshRecSum
     MDIMainForm.MousePointer = vbNormal
End Sub
Sub ViewALLRecords()
    Dim TMPname As String
    On Error Resume Next
    Dim RRrs As Long
    listEntries.Refresh
    
    listEntries.Redraw = False
  RRrs = 0
BackMejoe:
            'Open For Payrolls salary
            TMPname = Trim(listEntries.CellText(RRrs, 1))
            OpenPBDataBase ("Payrolls")
            Set PRFile = PDbase.OpenRecordset("SELECT * FROM Payrolls WHERE ecode Like '" & Trim(TMPname) & "' and Coverdate Like '" & Trim(Combo4.Text) & "' ")
            With PRFile
              Dim AMt As Double
              .MoveFirst
              AMt = 0
               Do While Not .EOF
                If Not .EOF Then
                    AMt = Val(AMt) + Val(![Amount])
                End If
                .MoveNext
               Loop
               listEntries.CellAlignment(RRrs, 2) = lgAlignRightCenter
               listEntries.CellText(RRrs, 2) = Format(Val(AMt), "###,###.00")
                .Close
            End With
            
            'Open For Deductions
            OpenPBDataBase ("DeductionsInfo")
            Set PRFile = PDbase.OpenRecordset("SELECT * FROM DeductionsInfo WHERE Dname Like '" & Trim(TMPname) & "' and DateCover LIKE '" & Trim(Combo4.Text) & "' ")
            With PRFile
              Dim AMts As Double
              .MoveFirst
              AMts = 0
               Do While Not .EOF
                If Not .EOF Then
                     AMts = Val(AMts) + Val(![DAmount])
                End If
                .MoveNext
               Loop
               listEntries.CellAlignment(RRrs, 3) = lgAlignRightCenter
               listEntries.CellText(RRrs, 3) = Format(Val(AMts), "###,###.00")
                .Close
            End With
            listEntries.CellAlignment(RRrs, 4) = lgAlignRightCenter
            listEntries.CellText(RRrs, 4) = Format(Round(Val(AMt) - Val(AMts), 2), "###,###.00")
            
    RRrs = RRrs + 1
    If RRrs <> listEntries.RowCount Then
        GoTo BackMejoe
    End If
    
    listEntries.Redraw = True
    listEntries.Refresh
    
    Call TotalSalary
    
End Sub
Sub TotalSalary()
Dim TSal As Double
Dim TDec As Double
Dim NSal As Double
Dim TSho As Double
Dim RRrs As Long
Dim Li As Long
RRrs = 0
BackMejoe:
    
    TSal = Val(TSal) + Val(Format(listEntries.CellText(RRrs, 2), "###.00"))
    TDec = Val(TDec) + Val(Format(listEntries.CellText(RRrs, 3), "###.00"))
    
    If Val(Format(listEntries.CellText(RRrs, 4), "###.00")) >= 1 Then
        
    Else
        TSho = Val(TSho) + Val(Format(listEntries.CellText(RRrs, 4), "###.00"))
    End If
    
    RRrs = RRrs + 1
    If RRrs <> listEntries.RowCount Then
        GoTo BackMejoe
    End If
    
    
    TS.Caption = Format(Val(TSal), "###,###.00")
    TD.Caption = Format(Val(TDec), "###,###.00")
    
    Label4.Caption = Format(Val(TSal) - Val(TDec), "###,###.00")
    
    TSho = Val(TSho * -2) / 2
    Label7.Caption = Format(Val(TSho), "###,###.00")
    
    NT.Caption = Format(Val(TSal - TDec) + Val(TSho), "###,###.00")
    
End Sub
Sub ADDrowTotal()
With listEntries
     .Redraw = False
     .Refresh
        .AddItem (""), 0
        .AddItem (""), 0
        .CellFontBold(.RowCount - 1, 0) = True
        .CellFontBold(.RowCount - 1, 1) = True
        '.CellText(.RowCount - 1, 0) = "822 PAYROLL for"
        .CellText(.RowCount - 1, 1) = "822 PAYROLL for " & Trim(Combo4.Text) '& " YEAR " & Format(Now, "YYYY")
        .AddItem (""), 0
        .CellFontBold(.RowCount - 1, 1) = True
        .CellText(.RowCount - 1, 1) = "Salary"
        .CellText(.RowCount - 1, 2) = Trim(TS.Caption)
        .AddItem (""), 0
        .CellFontBold(.RowCount - 1, 1) = True
        .CellText(.RowCount - 1, 1) = "Deductions"
        .CellText(.RowCount - 1, 2) = Trim(TD.Caption)
        .AddItem (""), 0
        .CellFontBold(.RowCount - 1, 1) = True
        .CellText(.RowCount - 1, 1) = "Subtotal"
        .CellText(.RowCount - 1, 2) = Trim(Label4.Caption)
        .AddItem (""), 0
        .CellFontBold(.RowCount - 1, 1) = True
        .CellText(.RowCount - 1, 1) = "Short"
        .CellText(.RowCount - 1, 2) = Trim(Label7.Caption)
        .AddItem (""), 0
        .CellFontBold(.RowCount - 1, 1) = True
        .CellText(.RowCount - 1, 1) = "Grand Total"
        .CellText(.RowCount - 1, 2) = Trim(NT.Caption)
        .AddItem (""), 0
     .Redraw = True
     .Refresh
    End With
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    
    'set list columns
    With listEntries
        .Redraw = False
        .AddColumn "Employee Signature", 120   '0
        .AddColumn "                         Name", 250   '1
        .AddColumn "      Salary", 80   '2
        .AddColumn "     Deductions", 80   '3
        .AddColumn "     Net Salary", 80   '4
        
        .Redraw = True
        .Refresh
    End With
    
    Call LOadcombo4
    Call LoadEntries
    
    Call Combo4_Click
End Sub
Private Sub Form_Activate()
    'MDIMainForm.JST(2).Expanded = True
    MDIMainForm.ActivateChild Me
    EPReports = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'MDIMainForm.Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & "NULL.pic")
    'MDIMainForm.JST(2).Expanded = False
    MDIMainForm.RemoveChild Me.Name
    EPReports = False
End Sub
Private Sub listEntries_Click()
    Dim TMPid As Double
    Dim TMPid2 As Integer
    
        If Trim(listEntries.CellText(listEntries.Row, 4)) >= 1 Then
            mnuSD.Enabled = False
        Else
            mnuSD.Enabled = True
        End If
        
        If Trim(listEntries.CellText(listEntries.Row, 3)) >= 1 Then
            mnuVD.Enabled = True
        Else
            mnuVD.Enabled = False
        End If
    
    
    TMPid = Val(listEntries.CellText(listEntries.Row, 0))
    On Error GoTo FFF
       MDIMainForm.Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & Trim(Format(TMPid, "0000000000000")) & ".pic")
    Exit Sub
FFF:
    MDIMainForm.Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & "NULL.pic")
End Sub
Private Sub listEntries_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu Me.MnuFile
    End If
End Sub
Private Sub listEntries_RowColChanged()
    RefreshRecSum
    Call listEntries_Click
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
Private Sub mnuSD_Click()
Dim ShrtAmt As Double
    frmEmployeeDeductions.Text2.Text = Trim(listEntries.CellText(listEntries.Row, 1))
    
    ShrtAmt = Val(Format(listEntries.CellText(listEntries.Row, 4), "###.00"))
    
    ShrtAmt = Val(ShrtAmt * -2) / 2
    
    frmEmployeeDeductions.AddMEdeductions (ShrtAmt)
    'frmEmployeeDeductions.Show 1
    
    'frmEmployeePatroll.OpenMe
    MDIMainForm.AddChild frmEmployeeDeductions, False
    frmEmployeeDeductions.Picture = MDIMainForm.ACPRibbon1.LoadBackground
    frmEmployeeDeductions.BackColor = MDIMainForm.ACPRibbon1.BackColor
    frmEmployeeDeductions.Option1.BackColor = MDIMainForm.ACPRibbon1.BackColor
    frmEmployeeDeductions.Option3.BackColor = MDIMainForm.ACPRibbon1.BackColor
    frmEmployeeDeductions.Option4.BackColor = MDIMainForm.ACPRibbon1.BackColor
    frmEmployeeDeductions.Option5.BackColor = MDIMainForm.ACPRibbon1.BackColor
    frmEmployeeDeductions.Option6.BackColor = MDIMainForm.ACPRibbon1.BackColor
    frmEmployeeDeductions.Option7.BackColor = MDIMainForm.ACPRibbon1.BackColor
    frmEmployeeDeductions.Option8.BackColor = MDIMainForm.ACPRibbon1.BackColor
    frmEmployeeDeductions.Option9.BackColor = MDIMainForm.ACPRibbon1.BackColor
End Sub
Private Sub mnuVD_Click()
    frmEmployeeDeductions.Text2.Text = Trim(listEntries.CellText(listEntries.Row, 1))
    frmEmployeeDeductions.OpenMe
    frmEmployeeDeductions.Show 1
End Sub
Private Sub mnuVP_Click()
    frmEmployeePatroll.Text1.Text = Trim(listEntries.CellText(listEntries.Row, 1))
    frmEmployeePatroll.OpenMe
    MDIMainForm.AddChild frmEmployeePatroll, False
End Sub

Private Sub Option1_Click()
    Call Combo4_Click
End Sub
Private Sub Option2_Click()
    Call Combo4_Click
End Sub
Private Sub Option3_Click()
    Call Combo4_Click
End Sub
Public Function Form_CanManageEmployee() As Boolean
        Form_CanManageEmployee = True
End Function
