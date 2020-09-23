VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageEmployee 
   BorderStyle     =   0  'None
   Caption         =   "Manage Employee"
   ClientHeight    =   9180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   612
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   822
   ShowInTaskbar   =   0   'False
   Begin MOVERS.LynxGrid3 listEntries 
      Height          =   7785
      Left            =   60
      TabIndex        =   1
      Top             =   390
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   13732
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
   Begin MOVERS.JOETitleBar JOETitleBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12330
      _ExtentX        =   21749
      _ExtentY        =   661
      Caption         =   "Manage Employee Records"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      ShadowColor     =   12632064
      BorderColor     =   49344
      BackColor       =   12648447
   End
   Begin MSComctlLib.ImageList ilList 
      Left            =   1470
      Top             =   8610
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageEmployee.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
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
      Left            =   105
      TabIndex        =   2
      Top             =   8280
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   345
      Left            =   15
      Picture         =   "frmManageEmployee.frx":059A
      Stretch         =   -1  'True
      Top             =   8220
      Width           =   12360
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuRefList 
         Caption         =   "Refresh List"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVinfo 
         Caption         =   "View Employee Info"
      End
      Begin VB.Menu mnuVEdec 
         Caption         =   "View Employee Deductions"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVEP 
         Caption         =   "View Employee Payroll"
      End
   End
End
Attribute VB_Name = "frmManageEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><>'
'Programmed By: Joehel V. del Rosario    '
'822 MOVERS(PAILO) Trucking Corporation  '
'Date: June 1, 2007                      '
'<><><><><><><><><><><><><><><><><><><><>'

Option Explicit

Public Function Form_CanManageEmployee() As Boolean
    'If GetTxtVal(b8DPStudent.BoundData) > 0 Then
        Form_CanManageEmployee = True
    'End If
End Function

Private Sub Form_Load()
    MDIMainForm.MousePointer = vbHourglass
    'set list columns
    With listEntries
        .Redraw = False
        .AddColumn "  Employee Code", 90   '0
        .AddColumn "                        Name", 230   '1
        .AddColumn "      Position", 80   '2
        .AddColumn "     Contact No.", 113   '3
        .AddColumn "                                      Address", 270   '4
        
        .Redraw = True
        .Refresh
    End With
    
    Call LoadEntries
    
End Sub
Private Sub Form_Activate()
    MDIMainForm.JST(2).Expanded = True
    MDIMainForm.ActivateChild Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MDIMainForm.Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & "NULL.pic")
    MDIMainForm.JST(2).Expanded = False
    MDIMainForm.RemoveChild Me.Name
End Sub
'----------------------------------------------------------
' Record Procedures
'----------------------------------------------------------

Private Sub LoadEntries()
    Dim iL As Long
    Dim ec As String
    Dim EN As String
    Dim EO As String
    Dim eB As String
    Dim EA As String
    
    
    'set app mouse icon
    'mdiMain.Form_StartBussy
    
    'clear list
    listEntries.Redraw = False
    listEntries.Clear
    
    On Error Resume Next
    
    'for database function
    OpenPBDataBase ("EmployeeInfo")
    With PRFile
    .MoveFirst
      Do While Not .EOF
        If Not .EOF Then
            ec = Trim(![ECOde])
            EN = Trim(![Ename])
            EO = Trim(![EOccupation])
            eB = Trim(![ContactNumber])
            EA = Trim(![EAddress])
            
        With listEntries
            .AddItem (Trim(ec)), 0
            .CellFontBold(iL, 1) = True
            .CellText(iL, 1) = Trim(EN)
            .CellText(iL, 2) = Trim(EO)
            .CellText(iL, 3) = Trim(eB)
            .CellText(iL, 4) = Trim(EA)
            iL = iL + 1
        End With
        
        End If
       .MoveNext
      Loop
    End With
    
    'Set vRS = Nothing
    listEntries.Redraw = True
    listEntries.Refresh
    'refresh rec sum
    RefreshRecSum
    'refresh recopt buttons
    'mdiMain.ActivateChild Me
    'restore mouse pointer
    'mdiMain.Form_EndBussy
    MDIMainForm.MousePointer = vbNormal
End Sub
Private Sub listEntries_Click()
    Dim TMPid As Double
    Dim TMPid2 As Integer
    
    TMPid = Val(listEntries.CellText(listEntries.Row, 0))
    'TMPid =
    On Error GoTo FFF
       MDIMainForm.Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & Trim(Format(TMPid, "0000000000000")) & ".pic")
    Exit Sub
FFF:
    MDIMainForm.Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & "NULL.pic")
End Sub

Private Sub listEntries_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu Me.mnufile
    End If
End Sub

Private Sub listEntries_RowColChanged()
    RefreshRecSum
    Call listEntries_Click
End Sub
Private Sub RefreshRecSum()
    lblRecSum.Caption = "Record " & listEntries.Row + 1 & " of " & listEntries.RowCount
    'refresh picture
    'mdiMain.Form_ShowStudentDetail
End Sub

Private Sub mnuRefList_Click()
    Me.Refresh
    Call LoadEntries
End Sub
Sub LSTREF()
    Call mnuRefList_Click
End Sub
Private Sub mnuVEdec_Click()
    frmEmployeeDeductions.Text2.Text = Trim(listEntries.CellText(listEntries.Row, 1))
    frmEmployeeDeductions.OpenMe
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

Private Sub mnuVEP_Click()
    frmEmployeePatroll.Text2.Text = Trim(listEntries.CellText(listEntries.Row, 0))
    frmEmployeePatroll.OpenMe
    MDIMainForm.AddChild frmEmployeePatroll, False
    'frmEmployeeDeductions.Show 1
End Sub

Private Sub mnuVinfo_Click()
    frmEmployeeEntry.Text2.Text = Trim(listEntries.CellText(listEntries.Row, 1))
    frmEmployeeEntry.OpenMe
    MDIMainForm.AddChild frmEmployeeEntry, False
    'Call JST_CaptionClick(m_TabFilterDate)
    frmEmployeeEntry.Picture = MDIMainForm.ACPRibbon1.LoadBackground
    frmEmployeeEntry.BackColor = MDIMainForm.ACPRibbon1.BackColor
    
End Sub
