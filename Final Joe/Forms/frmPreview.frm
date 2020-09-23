VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPreview 
   BorderStyle     =   0  'None
   Caption         =   "Print Preview"
   ClientHeight    =   9285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12360
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   5190
      TabIndex        =   6
      Top             =   8700
      Width           =   1230
   End
   Begin VB.PictureBox picScroll 
      Height          =   8160
      Left            =   60
      ScaleHeight     =   8100
      ScaleWidth      =   11895
      TabIndex        =   2
      Top             =   90
      Width           =   11955
      Begin VB.PictureBox picTarget 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   2655
         Left            =   165
         ScaleHeight     =   2625
         ScaleWidth      =   3825
         TabIndex        =   3
         Top             =   195
         Width           =   3855
      End
   End
   Begin VB.CheckBox chkColWidth 
      Caption         =   "Resize Col &widths to fill page"
      Height          =   195
      Left            =   3765
      TabIndex        =   1
      Top             =   9450
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox chkIcons 
      Caption         =   "With &Icons"
      Height          =   255
      Left            =   6405
      TabIndex        =   0
      Top             =   9450
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin MSForms.ScrollBar vscScroll 
      Height          =   8250
      Left            =   12015
      TabIndex        =   5
      Top             =   45
      Width           =   225
      Size            =   "397;14552"
   End
   Begin MSForms.ScrollBar hscScroll 
      Height          =   210
      Left            =   60
      TabIndex        =   4
      Top             =   8265
      Width           =   12150
      Size            =   "21431;370"
      Orientation     =   1
   End
End
Attribute VB_Name = "frmPreview"
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
    'cmdRefresh_Click
End Sub

Private Sub chkIcons_Click()
    'cmdRefresh_Click
End Sub
Sub PrintPL(LiR As LynxGrid3)
    If MsgBox("The application will now print the grid on the default printer (Show a print dialog here later !).", vbInformation + vbOKCancel, "Print") = vbCancel Then Exit Sub
    
    'Simply initialize the printer:
    Printer.Print
    
    'Read the FlexGrid:
    'Set the wanted width of the table to -1 to get the exact widths of the FlexGrid,
    ' to ScaleWidth - [the left and right margins] to get a fitting table !
    ImportListView cTP, LiR, IIf((chkColWidth.Value = vbChecked), Printer.ScaleWidth - 2 * 567, -1)
    
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

Private Sub Command1_Click()
    PrintPL frmPayrollReport.listEntries
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
    
    
    Set cTP = New clsTablePrint
    TransferDataList frmPayrollReport.listEntries
    'cmdRefresh_Click
End Sub
Private Sub Form_Activate()
    'MDIMainForm.JST(2).Expanded = True
    MDIMainForm.ActivateChild Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'MDIMainForm.Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & "NULL.pic")
    'MDIMainForm.JST(2).Expanded = False
    MDIMainForm.RemoveChild Me.Name
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


Sub TransferDataList(LiR As LynxGrid3)
    'Read the ListView:
    'Set the wanted width of the table to -1 to get the exact widths of the FlexGrid,
    ' to ScaleWidth - [the left and right margins] to get a fitting table !
    ImportListView cTP, LiR, IIf((chkColWidth.Value = vbChecked), picTarget.ScaleWidth - 2 * 567, -1), chkIcons.Value
    
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
