VERSION 5.00
Begin VB.Form frmWelcome 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Welcome"
   ClientHeight    =   9105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13935
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   607
   ScaleMode       =   0  'User
   ScaleWidth      =   929
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmWelcome"
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
Private Sub Form_Activate()
    MDIMainForm.ActivateChild Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'pass keyinfo to mdiMain
    MDIMainForm.AFForm_KeyDown KeyCode, Shift
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
End Sub

Private Sub Form_Resize()
    'Image5.Width = Me.Width + 10
    'Shape1.Width = Me.Width + 10
    'Shape1.Height = Me.Height + 10
    'Image2.Width = Me.Width + 10
End Sub
Sub JOEres(cC As Integer)
    'Image5.Width = MDIMainForm.Width + 10
    'Shape1.Width = MDIMainForm.Width + 10
    'Shape1.Height = MDIMainForm.Height + 10
    'Image2.Width = MDIMainForm.Width + 10
    'Me.Width = MDIMainForm.Width - 9000
    'Me.Height = MDIMainForm.Height - 2300
End Sub

