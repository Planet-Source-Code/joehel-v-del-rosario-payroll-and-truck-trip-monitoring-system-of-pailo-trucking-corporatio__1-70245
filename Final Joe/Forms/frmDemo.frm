VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo with MSFlexGrid"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin MOVERS.LynxGrid3 listEntries 
      Height          =   3810
      Left            =   90
      TabIndex        =   8
      Top             =   435
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   6720
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SBackColor1     =   0
      SBackColor2     =   0
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print the grid on the printer..."
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CheckBox chkIcons 
      Caption         =   "With &Icons"
      Height          =   255
      Left            =   8400
      TabIndex        =   7
      Top             =   3840
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox chkColWidth 
      Caption         =   "Resize Col &widths to fill page"
      Height          =   195
      Left            =   5760
      TabIndex        =   6
      Top             =   3840
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.PictureBox picScroll 
      Height          =   3375
      Left            =   5595
      ScaleHeight     =   3315
      ScaleWidth      =   4635
      TabIndex        =   4
      Top             =   420
      Width           =   4695
      Begin VB.PictureBox picTarget 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2655
         Left            =   -15
         ScaleHeight     =   2625
         ScaleWidth      =   3825
         TabIndex        =   5
         Top             =   15
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh PictureBox"
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   4080
      Width           =   1695
   End
   Begin MSForms.ScrollBar hscScroll 
      Height          =   240
      Left            =   3030
      TabIndex        =   10
      Top             =   5850
      Width           =   1245
      Size            =   "2196;423"
      Orientation     =   1
   End
   Begin MSForms.ScrollBar vscScroll 
      Height          =   2805
      Left            =   6090
      TabIndex        =   9
      Top             =   4860
      Width           =   315
      Size            =   "556;4948"
   End
   Begin ComctlLib.ImageList imlImages 
      Left            =   4800
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   32896
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmDemo.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "PictureBox as target:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   1
      Top             =   0
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   5520
      X2              =   5520
      Y1              =   4560
      Y2              =   0
   End
   Begin VB.Label Label1 
      Caption         =   "MSFlexGrid as source:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'The dimensions of the DIN A4 paper size in Twips:
Const A4Height = 16840, A4Width = 11907

'To get the scroll width:
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CYHSCROLL = 3
Private Const SM_CXVSCROLL = 2

'Declared Private WithEvents to get NewPage event:
Private WithEvents cTP As clsTablePrint
Attribute cTP.VB_VarHelpID = -1
Private Sub FillListView()
    Dim iL As Long
    Dim EC As String
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
            EC = Trim(![ECOde])
            EN = Trim(![Ename])
            EO = Trim(![EOccupation])
            eB = Trim(![ContactNumber])
            EA = Trim(![EAddress])
            
        With listEntries
            .AddItem (Trim(EC)), 0
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
    'RefreshRecSum
    'refresh recopt buttons
    'mdiMain.ActivateChild Me
    'restore mouse pointer
    'mdiMain.Form_EndBussy
    'MDIMainForm.MousePointer = vbNormal
End Sub

Private Sub InitializePictureBox()
    Dim sngVSCWidth As Single, sngHSCHeight As Single
    'Set the size to the DIN A4 width:
    picTarget.Width = A4Width
    picTarget.Height = A4Height
    'Resize the scrollbars:
    sngVSCWidth = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
    sngHSCHeight = GetSystemMetrics(SM_CYHSCROLL) * Screen.TwipsPerPixelY
    'hscScroll.Move 0, picScroll.ScaleHeight - sngHSCHeight, picScroll.ScaleWidth - sngVSCWidth, sngHSCHeight
    'vscScroll.Move picScroll.ScaleWidth - sngVSCWidth, 0, sngVSCWidth, picScroll.ScaleHeight
    
    SetScrollBars
End Sub

Private Sub SetScrollBars()
    hscScroll.Max = (picTarget.Width - picScroll.ScaleWidth + vscScroll.Width) / 120 + 1
    vscScroll.Max = (picTarget.Height - picScroll.ScaleHeight + hscScroll.Height) / 120 + 1
End Sub


Private Sub chkColWidth_Click()
    cmdRefresh_Click
End Sub

Private Sub chkIcons_Click()
    cmdRefresh_Click
End Sub

Private Sub cmdPrint_Click()
    
    If MsgBox("The application will now print the grid on the default printer (Show a print dialog here later !).", vbInformation + vbOKCancel, "Print") = vbCancel Then Exit Sub
    
    'Simply initialize the printer:
    Printer.Print
    
    'Read the FlexGrid:
    'Set the wanted width of the table to -1 to get the exact widths of the FlexGrid,
    ' to ScaleWidth - [the left and right margins] to get a fitting table !
    ImportListView cTP, fxgSrc, IIf((chkColWidth.Value = vbChecked), Printer.ScaleWidth - 2 * 567, -1)
    
    'Set margins (not needed, but looks better !):
    cTP.MarginBottom = 567 '567 equals to 1 cm
    cTP.MarginLeft = 567
    cTP.MarginTop = 567
    
    'Class begins drawing at CurrentY !
    Printer.CurrentY = cTP.MarginTop
    
    'Finally draw the Grid !
    cTP.DrawTable Printer
    'Done with drawing !
    
    'Say VB it should finally send it:
    Printer.EndDoc
End Sub

Private Sub cmdRefresh_Click()
    
    'Read the ListView:
    'Set the wanted width of the table to -1 to get the exact widths of the FlexGrid,
    ' to ScaleWidth - [the left and right margins] to get a fitting table !
    ImportListView cTP, listEntries, IIf((chkColWidth.Value = vbChecked), picTarget.ScaleWidth - 2 * 567, -1), chkIcons.Value
    
    'Here you can set RowHeightMin and HeaderRowMinHeight if the rows are too small:
'    cTP.RowHeightMin = 180
'    cTP.HeaderRowHeightMin = cTP.RowHeightMin
    
    
    'Set margins (not needed, but looks better !):
    cTP.MarginBottom = 567 '567 equals to 1 cm
    cTP.MarginLeft = 567
    cTP.MarginTop = 567
    
    'Clear the box:
    picTarget.Cls
    
    'Class begins drawing at CurrentY !
    picTarget.CurrentY = cTP.MarginTop
    
    'Finally draw the Grid !
    cTP.DrawTable picTarget
    'Done with drawing !
End Sub

Private Sub cTP_NewPage(objOutput As Object, TopMarginAlreadySet As Boolean, bCancel As Boolean, ByVal lLastPrintedRow As Long)
    
    'The class wants a new page, look what to do
    If TypeOf objOutput Is Printer Then
        Printer.NewPage
    Else 'We are printing on the PictureBox !
        objOutput.CurrentY = objOutput.ScaleHeight
        'Simply increase the height of the PicBox here
        ' (very simple, but looks bad in "real" applications)
        objOutput.Height = objOutput.Height + A4Height
        'Draw a line to show the new page:
        objOutput.Line (0, objOutput.CurrentY)-(objOutput.ScaleWidth, objOutput.CurrentY), &H808080
        
        'Set the CurrentY to the position the class should continie with drawing and...
        objOutput.CurrentY = objOutput.CurrentY + cTP.MarginTop
        '... tell it to do so:
        TopMarginAlreadySet = True
        
        'Set the ScrollBar Max properties:
        SetScrollBars
    End If
End Sub

Private Sub Form_Load()
    InitializePictureBox
    
    
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
    
    
    FillListView
    Set cTP = New clsTablePrint
    cmdRefresh_Click
End Sub


Private Sub hscScroll_Change()
    picTarget.Left = -hscScroll.Value * 120
End Sub

Private Sub hscScroll_Scroll()
    hscScroll_Change
End Sub


Private Sub listEntries_Click()
    MsgBox listEntries.ColHeading(0) '  .CellText(0, 0)
End Sub

Private Sub vscScroll_Change()
        'vscScroll.Max = 7000
    picTarget.Top = -vscScroll.Value * 120
End Sub


Private Sub vscScroll_Scroll()
    vscScroll_Change
End Sub


