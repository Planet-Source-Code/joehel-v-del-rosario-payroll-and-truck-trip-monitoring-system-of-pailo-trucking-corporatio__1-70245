VERSION 5.00
Begin VB.Form frmBilling 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Daily Record Entry"
   ClientHeight    =   9645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15315
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   15315
   ShowInTaskbar   =   0   'False
   Begin MOVERS.JOETitleBar JOETitleBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   661
      Caption         =   "Billing Entry"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ShadowColor     =   49152
      BorderColor     =   0
      BackColor       =   16384
   End
   Begin MOVERS.LynxGrid3 LynxGrid32 
      Height          =   6645
      Left            =   60
      TabIndex        =   1
      Top             =   2280
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   11721
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      BackColorBkg    =   16777215
      BackColorSel    =   8438015
      GridColor       =   11136767
      FocusRectColor  =   33023
      ColumnSort      =   -1  'True
      Striped         =   -1  'True
      SBackColor2     =   16777215
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   -255
      Picture         =   "frmBilling.frx":0000
      Stretch         =   -1  'True
      Top             =   9090
      Width           =   15795
   End
End
Attribute VB_Name = "frmBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<><><><><><><><><><><><><><><><><><><><>'
'Programmed By: Joehel V. del Rosario    '
'822 MOVERS(PAILO) Trucking Corporation  '
'Date: June 1, 2007                      '
'<><><><><><><><><><><><><><><><><><><><>'

Private Sub Form_Activate()
    'MDIMainForm.JST(2).Expanded = True
    MDIMainForm.ActivateChild Me
    Me.Width = MDIMainForm.Width - 200
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'MDIMainForm.Image1.Picture = LoadPicture(App.Path & "\Database\PICTURES\" & "NULL.pic")
    'MDIMainForm.JST(2).Expanded = False
    MDIMainForm.RemoveChild Me.Name
End Sub

