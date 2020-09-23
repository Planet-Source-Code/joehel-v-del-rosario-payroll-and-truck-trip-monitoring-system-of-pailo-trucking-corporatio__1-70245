VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPrint 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printing Preview"
   ClientHeight    =   14145
   ClientLeft      =   165
   ClientTop       =   -3285
   ClientWidth     =   10755
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   14145
   ScaleMode       =   0  'User
   ScaleWidth      =   10473.33
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   15360
      Left            =   15
      ScaleHeight     =   15360
      ScaleWidth      =   11640
      TabIndex        =   2
      Top             =   30
      Width           =   11640
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2265
         Left            =   45
         ScaleHeight     =   2265
         ScaleWidth      =   10335
         TabIndex        =   24
         Top             =   9225
         Width           =   10335
      End
      Begin MSComDlg.CommonDialog cmd 
         Left            =   1215
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00808080&
         Height          =   2190
         Left            =   315
         Top             =   11520
         Width           =   4950
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         Height          =   2190
         Left            =   315
         Top             =   11520
         Width           =   9840
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Trip Expense report From to"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   810
         TabIndex        =   23
         Top             =   11235
         Width           =   9120
      End
      Begin VB.Line Line1 
         X1              =   8340
         X2              =   10050
         Y1              =   12855
         Y2              =   12855
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cash Return :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   6225
         TabIndex        =   22
         Top             =   12855
         Width           =   1905
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   ".00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   9690
         TabIndex        =   21
         Top             =   12900
         Width           =   300
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   ".00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   9690
         TabIndex        =   20
         Top             =   12525
         Width           =   300
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Expense:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   6225
         TabIndex        =   19
         Top             =   12495
         Width           =   1455
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   ".00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   1875
         TabIndex        =   18
         Top             =   13065
         Width           =   1755
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   ".00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   1830
         TabIndex        =   17
         Top             =   13380
         Width           =   1800
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   ".00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   2085
         TabIndex        =   16
         Top             =   11745
         Width           =   1545
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   ".00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   2040
         TabIndex        =   15
         Top             =   12075
         Width           =   1590
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   ".00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   2010
         TabIndex        =   14
         Top             =   12735
         Width           =   1620
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   ".00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   2070
         TabIndex        =   13
         Top             =   12405
         Width           =   1545
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   ".00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   270
         Left            =   9675
         TabIndex        =   12
         Top             =   12180
         Width           =   300
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Total LTO/TMG :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   375
         TabIndex        =   11
         Top             =   13005
         Width           =   1440
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Others :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   375
         TabIndex        =   10
         Top             =   13320
         Width           =   1200
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Gas/Oil:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   375
         TabIndex        =   9
         Top             =   11715
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Meal Allow.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   375
         TabIndex        =   8
         Top             =   12015
         Width           =   1560
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Load :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   375
         TabIndex        =   7
         Top             =   12675
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Tool Fee:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Left            =   375
         TabIndex        =   6
         Top             =   12345
         Width           =   1290
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   9855
         TabIndex        =   5
         Top             =   60
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   9855
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cash Advance:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   6195
         TabIndex        =   3
         Top             =   12120
         Width           =   2025
      End
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   13860
      TabIndex        =   0
      Top             =   120
      Width           =   1290
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   30
      ScaleHeight     =   465
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   105
      Width           =   300
   End
   Begin VB.Line Line4 
      X1              =   3856.289
      X2              =   7478.864
      Y1              =   3420
      Y2              =   3420
   End
   Begin VB.Line Line2 
      X1              =   3856.289
      X2              =   7478.864
      Y1              =   3375
      Y2              =   3375
   End
   Begin VB.Line Line3 
      X1              =   3856.289
      X2              =   7478.864
      Y1              =   3330
      Y2              =   3330
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuprintdata 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnupagesetup 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu mnudefaultprinter 
         Caption         =   "Default Printer"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Print big form screen reolution declare
Private Const twipFactor = 1440
Private Const WM_PAINT = &HF
Private Const WM_PRINT = &H317
Private Const PRF_CLIENT = &H4&    ' Draw the window's client area.
Private Const PRF_CHILDREN = &H10& ' Draw all visible child windows.
Private Const PRF_OWNED = &H20&    ' Draw all owned windows.
Private Declare Function SendMessage Lib "user32" Alias _
   "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim DC As String
Private Sub Form_Load()
Me.Top = -100
If FormS = 1 Then
    Label25.FontBold = True
    Label25.Caption = "822 MOVERS PAYROL ----------------------- DATE COVERED: " & FormREPORT.Combo4.Text
    
    Label7.Caption = FormREPORT.Label2.Caption
    Label6.Caption = FormREPORT.Label5.Caption
    Label4.Caption = FormREPORT.Label9.Caption
    
    Label14.Caption = FormREPORT.ts.Caption
    Label13.Caption = FormREPORT.td.Caption
    Label11.Caption = FormREPORT.nt.Caption
    
    
    Line1.Visible = False
    Label3.Visible = False
    Label5.Visible = False
    Label8.Visible = False
    Label9.Visible = False
    Label10.Visible = False
    Label12.Visible = False
    Label15.Visible = False
    Label16.Visible = False
    Label17.Visible = False
    Label18.Visible = False
    Label21.Visible = False
    Label22.Visible = False
    Picture3.Visible = True
ElseIf FormS = 2 Then
Picture3.Visible = False
With FormTripL
    If .Combo2.Text = .Combo5.Text And .Combo4.Text = .Combo7.Text Then
        DC = .Combo2.Text & " " & .Combo3.Text & "-" & .Combo6.Text & ", " & .Combo4.Text
    ElseIf Combo2.Text <> Combo5.Text And Combo4.Text = Combo7.Text Then
        DC = .Combo2.Text & " " & .Combo3.Text & "-" & .Combo5.Text & " " & .Combo6.Text & ", " & .Combo4.Text
    Else
        DC = .Combo2.Text & " " & .Combo3.Text & ", " & .Combo4.Text & "-" & .Combo5.Text & " " & .Combo6.Text & ", " & .Combo7.Text
    End If

    
    If .Option1.Value = True Then
        Label25.Caption = "Trip Expense Report of " & DC
    Else
        Label25.Caption = .Combo1.Text & " " & "Trip Expense Report of " & DC
    End If
    
    .GRID.Row = .GRID.Rows - 1
    .GRID.Col = 2
        Label10.Caption = Format(.GRID.Text, "###,###.00")
    .GRID.Col = 3
        Label14.Caption = Format(.GRID.Text, "###,###.00")
    .GRID.Col = 4
        Label13.Caption = Format(.GRID.Text, "###,###.00")
    .GRID.Col = 5
        Label11.Caption = Format(.GRID.Text, "###,###.00")
    .GRID.Col = 6
        Label12.Caption = Format(.GRID.Text, "###,###.00")
    .GRID.Col = 7
        Label16.Caption = Format(.GRID.Text, "###,###.00")
    .GRID.Col = 8
        Label15.Caption = Format(.GRID.Text, "###,###.00")
    .GRID.Col = 9
        Label18.Caption = Format(.GRID.Text, "###,###.00")
    .GRID.Col = 10
        Label21.Caption = Format(.GRID.Text, "###,###.00")
End With
End If
On Error Resume Next
Call fixprintarea
'frmscroll.Show



End Sub

Sub fixprintarea()
On Error Resume Next
Picture1.BackColor = vbWhite
Me.BackColor = vbWhite

Dim sWide As Single, sTall As Single
   Dim rv As Long

   Me.ScaleMode = vbTwips
   sWide = 8.5
   sTall = 13      ' or 14
   Me.Width = twipFactor * sWide
   Me.Height = twipFactor * sTall
   With Picture1
      .Top = 0
      .Left = 0
      .Width = twipFactor * sWide
      .Height = twipFactor * sTall
   End With
   With Picture2
      .Top = 0
      .Left = 0
      .Width = twipFactor * sWide
      .Height = twipFactor * sTall
   End With
   With Label1
      .Left = Me.Width / 2
      .Top = 0
   End With
   With Label2
      .Top = (twipFactor * sTall) - .Height * 2
      .Left = Me.Width / 2
   End With
   Me.Visible = True
   DoEvents

   Picture1.SetFocus
   Picture2.AutoRedraw = True
   rv = SendMessage(Picture1.hwnd, WM_PAINT, Picture2.hdc, 0)
   rv = SendMessage(Picture1.hwnd, WM_PRINT, Picture2.hdc, _
   PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
   Picture2.Picture = Picture2.Image
   Picture2.AutoRedraw = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call mnuexit_Click
End Sub

Private Sub mnudefaultprinter_Click()
frmprnt.Label1.Caption = (Printer.DeviceName)

frmprnt.Show
End Sub

Private Sub mnuexit_Click()
On Error Resume Next
If FormS = 1 Then
    SetParent FormREPORT.lvData.hwnd, FormREPORT.hwnd
    FormREPORT.lvData.Top = 1170
    FormREPORT.lvData.Left = 180
    FormREPORT.lvData.Height = 5565
    FormREPORT.lvData.Width = 11115
    FormREPORT.Show
    Unload Me
ElseIf FormS = 2 Then
    SetParent FormTripL.GRID.hwnd, FormTripL.hwnd
    FormTripL.GRID.Height = 6045
    FormTripL.GRID.Width = 11145
    FormTripL.GRID.Top = 1845
    FormTripL.GRID.Left = 390
    FormTripL.Show
    Unload Me
End If
End Sub

Private Sub mnupagesetup_Click()
Dim BeginPage, EndPage, NumCopies, Orientation, i
   ' Set Cancel to True.
   cmd.CancelError = True
   On Error GoTo ErrHandler
   ' Display the Print dialog box.
   cmd.ShowPrinter
   ' Get user-selected values from the dialog box.
   BeginPage = cmd.FromPage
   EndPage = cmd.ToPage
   NumCopies = cmd.Copies
   Orientation = cmd.Orientation
   For i = 1 To NumCopies
   
   frmscroll.Hide
   Printer.PrintQuality = vbPRPQLow
   Printer.Print ""
   Printer.PaintPicture Picture2.Picture, 0, 0
   
   frmPrint.PrintForm
   Printer.EndDoc
   ' Put code here to send data to your printer.
   Next
   Exit Sub

ErrHandler:
Exit Sub

End Sub

Private Sub mnuprintdata_Click()

On Error Resume Next

'frmscroll.Hide
Printer.PrintQuality = vbPRPQLow

  Printer.Print ""
  Printer.PaintPicture Picture2.Picture, 0, 0
  frmPrint.PrintForm
  
  Printer.EndDoc

End Sub

