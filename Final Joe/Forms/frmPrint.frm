VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrint 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printing Preview"
   ClientHeight    =   14145
   ClientLeft      =   165
   ClientTop       =   -3285
   ClientWidth     =   10470
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   14145
   ScaleMode       =   0  'User
   ScaleWidth      =   10195.79
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   15360
      Left            =   60
      ScaleHeight     =   15360
      ScaleWidth      =   11640
      TabIndex        =   2
      Top             =   75
      Width           =   11640
      Begin MSComDlg.CommonDialog cmd 
         Left            =   3855
         Top             =   780
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   9855
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   9855
         TabIndex        =   4
         Top             =   60
         Width           =   1335
      End
      Begin VB.Label lblTittle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Tittle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4785
         TabIndex        =   3
         Top             =   75
         Width           =   420
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
      Left            =   210
      ScaleHeight     =   465
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   405
      Width           =   300
   End
   Begin VB.Line Line4 
      X1              =   3856.287
      X2              =   7478.86
      Y1              =   3420
      Y2              =   3420
   End
   Begin VB.Line Line2 
      X1              =   3856.287
      X2              =   7478.86
      Y1              =   3375
      Y2              =   3375
   End
   Begin VB.Line Line3 
      X1              =   3856.287
      X2              =   7478.86
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
'<><><><><><><><><><><><><><><><><><><><>'
'Programmed By: Joehel V. del Rosario    '
'822 MOVERS(PAILO) Trucking Corporation  '
'Date: June 1, 2007                      '
'<><><><><><><><><><><><><><><><><><><><>'

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

    frmTripReport.listEntries.BackColorSel = vbWhite
    frmTripReport.listEntries.GridColor = &HC0C0C0
    frmTripReport.listEntries.BackColorEdit = vbWhite
    frmTripReport.listEntries.BackColor = vbWhite
    frmTripReport.Picture1.Top = 120
    frmTripReport.Picture1.Left = 20
    frmTripReport.listEntries.Height = 15000
    frmTripReport.Picture1.Height = 15010
    SetParent frmTripReport.Picture1.hwnd, frmPrint.Picture1.hwnd

On Error Resume Next
Call fixprintarea
'frmscroll.Show

Call fixprintarea

Call fixprintarea

End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmTripReport.listEntries.BackColorSel = &H80C0FF
    frmTripReport.listEntries.GridColor = &HA9EEFF
    frmTripReport.listEntries.Height = 6165
    frmTripReport.Picture1.Height = 6167
    'frmTripReport.listEntrieS1.BackColor = vbWhite
    SetParent frmTripReport.Picture1.hwnd, frmTripReport.hwnd
    
    frmTripReport.Picture1.Top = 1590
    frmTripReport.Picture1.Left = 60
    Call mnuexit_Click
End Sub

Sub fixprintarea()
On Error Resume Next
frmTripReport.Picture1.BackColor = vbWhite
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

   frmTripReport.Picture1.SetFocus
   Picture2.AutoRedraw = True
   rv = SendMessage(frmTripReport.Picture1.hwnd, WM_PAINT, Picture2.hDc, 0)
   rv = SendMessage(frmTripReport.Picture1.hwnd, WM_PRINT, Picture2.hDc, _
   PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
   Picture2.Picture = Picture2.Image
   Picture2.AutoRedraw = False

End Sub



Private Sub mnudefaultprinter_Click()
frmprnt.Label1.Caption = (Printer.DeviceName)

frmprnt.Show
End Sub

Private Sub mnuexit_Click()
On Error Resume Next
    Unload Me
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
   
   'frmscroll.Hide
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

