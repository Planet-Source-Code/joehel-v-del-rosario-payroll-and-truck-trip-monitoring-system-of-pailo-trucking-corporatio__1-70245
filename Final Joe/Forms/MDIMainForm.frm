VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIMainForm 
   BackColor       =   &H8000000F&
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   -77490
   ClientWidth     =   15120
   Icon            =   "MDIMainForm.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3105
      Top             =   2400
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   3105
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3105
      Top             =   1470
   End
   Begin VB.PictureBox bgHeader 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   0
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1008
      TabIndex        =   0
      Top             =   0
      Width           =   15120
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "822 Movers(PAILO) Payroll and Billing System ™"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C25418&
         Height          =   360
         Left            =   105
         TabIndex        =   21
         Top             =   75
         Width           =   8505
      End
      Begin VB.Image Image3 
         Height          =   465
         Left            =   -105
         Picture         =   "MDIMainForm.frx":000C
         Stretch         =   -1  'True
         Top             =   15
         Width           =   15675
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ystem"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7965
         TabIndex        =   17
         Top             =   570
         Width           =   6915
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   -15
         Picture         =   "MDIMainForm.frx":79FC
         Stretch         =   -1  'True
         Top             =   465
         Width           =   15390
      End
   End
   Begin MOVERS.JOEClientWin JOEClientWin1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   1
      Top             =   10515
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   820
   End
   Begin MOVERS.JoeSBCenter JoeSBCenter1 
      Align           =   3  'Align Left
      Height          =   9525
      Left            =   0
      TabIndex        =   2
      Top             =   990
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   16801
      Begin VB.PictureBox Picture3 
         Height          =   360
         Left            =   165
         ScaleHeight     =   300
         ScaleWidth      =   2430
         TabIndex        =   12
         Top             =   10555
         Width           =   2490
         Begin VB.ListBox List4 
            Height          =   255
            Left            =   30
            TabIndex        =   13
            Top             =   900
            Width           =   1275
         End
         Begin VB.ListBox List3 
            Height          =   255
            Left            =   30
            TabIndex        =   14
            Top             =   615
            Width           =   1275
         End
         Begin VB.ListBox List2 
            Height          =   255
            Left            =   30
            TabIndex        =   15
            Top             =   360
            Width           =   1275
         End
         Begin VB.ListBox List1 
            Height          =   255
            Left            =   30
            TabIndex        =   16
            Top             =   60
            Width           =   1275
         End
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1515
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   10395
         Width           =   1140
      End
      Begin MOVERS.JOESideTab JST 
         Height          =   3015
         Index           =   0
         Left            =   60
         TabIndex        =   3
         Top             =   645
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   5318
         Caption         =   "Menu"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   9
         ForeColor       =   12735512
         BorderColor     =   16777215
         Begin MSComctlLib.ListView listQL 
            Height          =   2670
            Left            =   0
            TabIndex        =   4
            Top             =   345
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   4710
            View            =   2
            Arrange         =   1
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   -1  'True
            OLEDropMode     =   1
            _Version        =   393217
            Icons           =   "ilQL"
            SmallIcons      =   "ilQL"
            ForeColor       =   12735512
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OLEDropMode     =   1
            NumItems        =   0
         End
         Begin MSComctlLib.ImageList ilQL 
            Left            =   1020
            Top             =   630
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   4
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "MDIMainForm.frx":10C5E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "MDIMainForm.frx":10D71
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "MDIMainForm.frx":114EB
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "MDIMainForm.frx":11C65
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin MOVERS.JOESideTab JST 
         Height          =   1950
         Index           =   1
         Left            =   60
         TabIndex        =   5
         Top             =   3675
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   3440
         Caption         =   "Information"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   9
         ForeColor       =   12735512
         BorderColor     =   16777215
         Begin VB.Label animates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   $"MDIMainForm.frx":123DF
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   3810
            Index           =   0
            Left            =   180
            TabIndex        =   19
            Top             =   -2025
            Width           =   2520
         End
         Begin VB.Label animates 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   $"MDIMainForm.frx":1257F
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   3810
            Index           =   1
            Left            =   180
            TabIndex        =   18
            Top             =   1770
            Width           =   2520
         End
      End
      Begin MOVERS.JOESideTab JST 
         Height          =   3105
         Index           =   2
         Left            =   60
         TabIndex        =   6
         Top             =   5640
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   5477
         Caption         =   "Employee's Picture"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   9
         ForeColor       =   12735512
         BorderColor     =   16777215
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            ForeColor       =   &H80000008&
            Height          =   2760
            Left            =   0
            ScaleHeight     =   2730
            ScaleWidth      =   2880
            TabIndex        =   7
            Top             =   345
            Width           =   2910
            Begin VB.PictureBox Picture1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   2130
               Left            =   540
               ScaleHeight     =   2100
               ScaleWidth      =   1755
               TabIndex        =   20
               Top             =   300
               Width           =   1785
               Begin VB.Image Image1 
                  Height          =   2070
                  Left            =   30
                  Stretch         =   -1  'True
                  Top             =   15
                  Width           =   1710
               End
            End
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Today is "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C25418&
         Height          =   195
         Left            =   105
         TabIndex        =   11
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome,"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C25418&
         Height          =   195
         Left            =   105
         TabIndex        =   10
         Top             =   90
         Width           =   705
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C25418&
         Height          =   195
         Left            =   810
         TabIndex        =   9
         Top             =   360
         Width           =   180
      End
      Begin VB.Label lblCurrentUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Not Log-in"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C25418&
         Height          =   195
         Left            =   840
         TabIndex        =   8
         Top             =   105
         Width           =   855
      End
   End
   Begin VB.Menu mnuF 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuER 
         Caption         =   "Employee Registration"
      End
      Begin VB.Menu mnuED 
         Caption         =   "Employee Deductions"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTE 
         Caption         =   "Trip Entry"
      End
      Begin VB.Menu mnuTS 
         Caption         =   "Trip Status"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuLout 
         Caption         =   "Log-Out"
      End
      Begin VB.Menu sdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuC 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuP 
      Caption         =   "Payroll"
      Visible         =   0   'False
      Begin VB.Menu mnuDHP 
         Caption         =   "Drivers and Helpers Personnel"
      End
      Begin VB.Menu mnuOP 
         Caption         =   "Office Personnel"
      End
   End
   Begin VB.Menu mnuR 
      Caption         =   "Reports"
      Visible         =   0   'False
      Begin VB.Menu mnuERep 
         Caption         =   "Employee Report"
      End
      Begin VB.Menu mnuTrep 
         Caption         =   "Trip Reports"
      End
      Begin VB.Menu mnuPrep 
         Caption         =   "Payroll Report"
      End
   End
   Begin VB.Menu mnuS 
      Caption         =   "Settings"
      Visible         =   0   'False
      Begin VB.Menu mnuTP 
         Caption         =   "Truck Personnels"
      End
      Begin VB.Menu mnuSW 
         Caption         =   "Salaries and Wages"
      End
      Begin VB.Menu mnuCL 
         Caption         =   "Customers List"
      End
   End
   Begin VB.Menu MnuH 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu mnuAA 
         Caption         =   "About the Author"
      End
   End
End
Attribute VB_Name = "MDIMainForm"
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

Private Const m_TabShowQuickLaunch = 0
Private Const m_TabSearch = 1
Private Const m_TabFilterDate = 2
Public Function ShowForm()
    
 

End Function

Private Sub JOESBtop1_Resize()
    'JoeSBCenter1.Width = JOESBtop1.Width * Screen.TwipsPerPixelX
    'frmWelcome.JOEres
End Sub

Private Sub JOESBtop1_SizeChange(ByVal newSizeState As eSizeState)
    If newSizeState = ssContracted Then
        JOEClientWin1.SBWidth = JoeSBCenter1.Width / Screen.TwipsPerPixelX
        JoeSBCenter1.Visible = True
        frmWelcome.JOEres 3280
    Else
        JOEClientWin1.SBWidth = 0
        JoeSBCenter1.Visible = False
        frmWelcome.JOEres 180
        
    End If
    'call mdi resize to resize all opened child forms
End Sub

Private Sub jst_BeforeExpand(Index As Integer)
    'resize contained controlsbeofre expanding
    Select Case Index
        Case m_TabShowQuickLaunch
            listQL.Move 90, listQL.Top, JST(Index).Width - 150

        Case m_TabSearch 'search
            'resize
            'Picture1.Move 90, Picture1.Top, JST(Index).Width - 150
        Case m_TabFilterDate 'filter date
            If Form_CanManageEmployee = False Then
                MsgBox "There is no Employee's Picture to display.", vbInformation
                JST(Index).Expanded = False
            Else
            Picture2.Move 90, Picture2.Top, JST(Index).Width - 150
            End If
    End Select

End Sub

Private Sub JST_CaptionClick(Index As Integer)
    Select Case Index
        Case m_TabShowQuickLaunch
            JST(m_TabShowQuickLaunch).Height = 3015

        Case m_TabSearch 'search
            JST(m_TabSearch).Height = 1950
        Case m_TabFilterDate 'filter date
            JST(m_TabFilterDate).Height = 3660
    End Select
    
End Sub

Private Sub jst_CompleteExpand(Index As Integer)
    Dim i As Integer
    
    For i = 0 To 2
        If Index <> i Then
            If JST(i).AutoContract = True Then
                JST(i).Expanded = False
            End If
        End If
    Next

End Sub

Private Sub jst_Resize(Index As Integer)
    
    Dim i As Integer
    
    For i = 1 To 2
        JST(i).Move JST(i).Left, (JST(i - 1).Top + JST(i - 1).Height) '- 200
    Next
    
    If JST(Index).Expanded = True Then
        Select Case Index
            Case m_TabShowQuickLaunch
                listQL.Move 90, listQL.Top, JST(Index).Width - 150

            Case m_TabSearch 'search
                'resize
                'Picture1.Move 90, Picture1.Top, JST(Index).Width - 150
            Case m_TabFilterDate 'filter date
                Picture2.Move 90, Picture2.Top, JST(Index).Width - 150
        End Select
    End If

End Sub
Private Sub JOEClientWin1_CloseClick(ByVal sFormName As String, ByVal Index As Integer)
    'close form
    Dim frm As Form
    
    On Error GoTo RAE
    
    For Each frm In Forms
        If LCase(Trim(frm.Name)) = LCase(Trim(sFormName)) Then
            Unload frm
            Exit For
        End If
    Next
    
RAE:
    Set frm = Nothing

End Sub

Private Sub JOEClientWin1_FormTabClick(ByVal sFormName As String, ByVal Index As Integer)
    modFuncChild.ActivateMDIChildForm sFormName
End Sub

Public Function Form_ShowQuickLaunch()

    'expand side bar
    'If JOESBtop1.SizeState <> ssContracted Then
    '    JOESBtop1.SizeState = ssContracted
    'End If

    'expand search tab
    If JST(m_TabShowQuickLaunch).Expanded = False Then
        JST(m_TabShowQuickLaunch).Expanded = True
    End If
    
    On Error Resume Next
    JST(m_TabShowQuickLaunch).SetFocus
    'HLTxt txtSearchWhat
    Err.Clear
    
End Function

Public Function Form_ShowSearch()

    'expand side bar
    'If JOESBtop1.SizeState <> ssContracted Then
    '    JOESBtop1.SizeState = ssContracted
    'End If

    'expand search tab
    If JST(m_TabSearch).Expanded = False Then
        JST(m_TabSearch).Expanded = True
    End If
    
    On Error Resume Next
    JST(m_TabSearch).SetFocus
    'HLTxt txtSearchWhat
    Err.Clear
    
End Function


Public Function Form_ShowDateFilter()

    'expand side bar
    'If JOESBtop1.SizeState <> ssContracted Then
    '    JOESBtop1.SizeState = ssContracted
    'End If

    'expand search tab
    If JST(m_TabFilterDate).Expanded = False Then
        JST(m_TabFilterDate).Expanded = True
    End If
    
    On Error Resume Next
    JST(m_TabFilterDate).SetFocus
    'b8DateP.SetFocus
    Err.Clear
    
End Function


Private Sub Label5_Click()
    FormCoveredDate.Show
End Sub

Private Sub listQL_Click()
    Dim selItemKey As String
    
    On Error GoTo RAE
    
    selItemKey = listQL.SelectedItem.Key
    
    Select Case selItemKey
        Case "EmR" 'Employee Registration"
            'Call JOESBtop1_SizeChange(ssExpanded)
            'MDIMainForm.AddChild frmDailyRecord, True
            frmEmployeeEntry.Show 1
        Case "EmD" 'Employee Deductions"
            'Call JOESBtop1_SizeChange(ssExpanded)
            frmEmployeeDeductions.Show 1
        Case "TrE" 'Trip Entry"
            MDIMainForm.AddChild frmTrips, False
        Case "EmPay" 'Payrolls"
            MDIMainForm.AddChild frmEmployeePatroll, False
            Call JST_CaptionClick(m_TabFilterDate)
        Case "PReP" 'Payroll Reports"
            MDIMainForm.AddChild frmPayrollReport, False
            Call JST_CaptionClick(m_TabFilterDate)
        Case "Rank" 'Manage Employee List"
            MDIMainForm.AddChild frmManageEmployee, False
            Call JST_CaptionClick(m_TabFilterDate)
        Case "TrERep"
            MDIMainForm.AddChild frmTripReport, False
        Case "Sett"
            MDIMainForm.AddChild FormSettings, False
        Case "About"
            MDIMainForm.AddChild FormPAY, False
    End Select
  
RAE:

End Sub
Sub ShowSideTab()
    Call JOESBtop1_SizeChange(ssContracted)
End Sub
Private Sub McP_Click()
    End
End Sub

Private Sub listQL_KeyPress(KeyAscii As Integer)
    'Call listQL_DblClick
End Sub

Private Sub MDIForm_Load()
    'JOELine1.Width = Me.Width + 200
    'JOELine2.Width = Me.Width + 200
    
    'Set JOeMenu1.Menu = Me.mnuF
    'Set JOeMenu2.Menu = Me.mnuP
    'Set JOeMenu3.Menu = Me.mnuR
    'Set JOeMenu4.Menu = Me.mnuS
    'Set JOeMenu5.Menu = Me.MnuH
    
    'add Quick Launch Items
    With listQL.ListItems
        .Add , "EmR", "Employee Registration", 2, 2
        .Add , "EmD", "Employee Deductions", 2, 2
        .Add , "TrE", "Trip Entry", 2, 2
        .Add , "TrERep", "Trip Expense Reports", 2, 2
        .Add , "EmPay", "Payrolls", 2, 2
        .Add , "PReP", "Payroll Reports", 2, 2
        .Add , "Rank", "Manage Employee Records", 2, 2
        .Add , "Sett", "System Settings", 2, 2
        .Add , "About", "About the Author", 2, 2
    End With
    
    MDIMainForm.AddChild frmWelcome, False
    
    JOEClientWin1.SBWidth = JoeSBCenter1.Width / Screen.TwipsPerPixelX
    Call LoadLISTname
    Call OpenDATECOver
    frmLogin.Show 1
End Sub
Sub OpenDATECOver()
On Error Resume Next
If Label5.Caption <> "" Then
    OpenPBDataBase ("DateCover")
    Set PRFile = PDbase.OpenRecordset("SELECT * FROM DateCover WHERE Status Like '" & "1" & "' ")
    With PRFile
        If Not .EOF Then
            Label5.Caption = "Covered date for Salary : " & ![CoveredDate]
        Else
            Label5.Caption = "Date NOT SET"
        End If
   End With
 End If
End Sub
Sub LoadLISTname()
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    OpenPBDataBase ("EmployeeInfo")
    With PRFile
        If .RecordCount > 1 Then
            Do Until .EOF
                List1.AddItem ![Ename]
                List2.AddItem ![ECOde]
                If ![EOccupation] = "Helper" Then
                    List3.AddItem ![Ename]
                Else
                    List4.AddItem ![Ename]
                End If
            .MoveNext
            Loop
        End If
        .Close
    End With
End Sub
Private Sub mnuC_Click()
    End
End Sub

Private Sub MnuLout_Click()
    Me.lblCurrentUser.Caption = "Log-Out"
    frmLogin.Show 1
End Sub

Private Sub Timer1_Timer()
    lblDate.Caption = FormatDateTime(Now, vbGeneralDate)
End Sub

' MDI Form procedures
'-----------------------------------------------------------
Private Sub MDIForm_Resize()
        
    Dim frm As Form
    
    
    On Error Resume Next
    
    
    'resize childs
    If GetActiveChildCount > 0 Then
        For Each frm In Forms
        If frm.Name <> Me.Name Then
            If frm.MDIChild = True Then
                If frm.Name = Me.ActiveForm.Name Then
                    ResizeMdiChildForm frm
                Else
                    frm.Visible = False
                End If
            End If
        End If
        
        Next
        
    End If
    
    Set frm = Nothing
End Sub

'Get Opened MDI Child Forms Count
Public Function GetActiveChildCount() As Integer
    
    Dim frm As Form
    Dim iCount As Integer
    
    iCount = 0
    
    For Each frm In Forms
        If frm.Name <> Me.Name Then
            If frm.MDIChild = True Then
                iCount = iCount + 1
            End If
        End If
    Next
    
    GetActiveChildCount = iCount
    Set frm = Nothing
    
End Function
'-----------------------------------------------------------
' >> End MDI Form procedures
'------------------------------------------------------------

'------------------------------------------------------------
' Parent To Child procedures
'------------------------------------------------------------

Public Sub AddChild(ByRef CFrm As Form, Optional CloseButton As Boolean = True)

    'load form
    modFuncChild.LoadForm CFrm, CloseButton
    
End Sub



Public Sub ActivateChild(ByRef CFrm As Form)
    'activate form
    Me.JOEClientWin1.SetActiveWindow CFrm.Name
    Form_CanManageEmployee
End Sub
Public Function Form_CanManageEmployee() As Boolean
    
    Dim bReturn As Boolean
    
    On Error Resume Next
    
    bReturn = False
    bReturn = Me.ActiveForm.Form_CanManageEmployee
            
    If bReturn = False Then
        JST(m_TabFilterDate).Expanded = False
    End If
    
    Form_CanManageEmployee = bReturn
    
    Err.Clear
    
End Function

Public Sub RemoveChild(ByVal sFormName As String)
    'remove form
     Me.JOEClientWin1.RemoveChildWindow sFormName
End Sub
Private Sub Timer2_Timer()
    If animates(0).Top = -2055 Then
        animates(1).Top = 1710
    End If
    
    If animates(1).Top = -2055 Then
        animates(0).Top = 1710
    End If
    
    animates(0).Top = animates(0).Top - 5
    animates(1).Top = animates(1).Top - 5
End Sub

Private Sub Timer3_Timer()
    Label5.Visible = Not Label5.Visible
End Sub

Public Sub AFForm_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 83 And Shift = 4 Then
        
    ElseIf KeyCode = 82 And Shift = 4 Then
        
    ElseIf KeyCode = 77 And Shift = 4 Then
        
    ElseIf KeyCode = 84 And Shift = 4 Then
        
    ElseIf KeyCode = 72 And Shift = 4 Then
        
    ElseIf KeyCode = 81 And Shift = 2 Then

    ElseIf KeyCode = 68 And Shift = 2 Then

    End If
    
    'MsgBox KeyCode & " - " & Shift
End Sub

