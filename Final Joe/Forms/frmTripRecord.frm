VERSION 5.00
Begin VB.Form frmTripRecord 
   BorderStyle     =   0  'None
   Caption         =   "Trip Record"
   ClientHeight    =   9645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   ScaleHeight     =   9645
   ScaleWidth      =   15315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MOVERS.LynxGrid3 listEntries 
      Height          =   8400
      Left            =   1260
      TabIndex        =   0
      Top             =   510
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   14817
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
      BackColorSel    =   8438015
      GridColor       =   11136767
      FocusRectColor  =   33023
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
      TabIndex        =   1
      Top             =   0
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   661
      Caption         =   "Trip Record"
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
   Begin VB.Image Image2 
      Height          =   435
      Left            =   0
      Picture         =   "frmTripRecord.frx":0000
      Stretch         =   -1  'True
      Top             =   9090
      Width           =   15795
   End
End
Attribute VB_Name = "frmTripRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
