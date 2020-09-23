Attribute VB_Name = "Functions"
'for database
Global PDbase As Database
Global PRFile As Recordset
'Global FormS As Integer
Global PFOrms As Integer
Global EPReports As Boolean
Global EPRepR As Long


Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByRef lParam As Any _
) As Long
Const LB_FINDSTRING = &H18F
Global DelKey As Boolean
Global bNoClick As Boolean
Sub AutoTXTcomplete(LST As ListBox, TXT As TextBox)
Dim strt As Long, nIndex As Long
Dim nLen As Long, sText As String
Const LB_GETTEXTLEN As Long = &H18A
Const LB_GETTEXT As Long = &H189
Static blnBusy As Boolean

If blnBusy Then
   Exit Sub
End If
     
     bNoClick = True
     blnBusy = True
    
    'Retrieve the item's listindex
    LST.ListIndex = SendMessage(LST.hwnd, LB_FINDSTRING, -1, ByVal CStr(TXT.Text))
    
    If Not DelKey Then


    If LST.ListIndex <> -1 Then
        strt = Len(TXT.Text)
        TXT.Text = LST.List(LST.ListIndex)
        TXT.SelStart = strt
        TXT.SelLength = Len(TXT.Text) - strt
    Else
    
    End If
    End If
       DelKey = False
       blnBusy = False
       bNoClick = False
End Sub
Sub OpenPBDataBase(a As String)
    'for network database
    'Dim DBPath As String
    'DBPath = "\\192.168.0.23\i\Final Joe\Database\PayrollBilling1.mdb"
    'Set PDbase = OpenDatabase(DBPath)
    
    'for server database
    Set PDbase = OpenDatabase(App.Path & "\Database\PayrollBilling1.mdb")
    Set PRFile = PDbase.OpenRecordset(a)
End Sub
Sub PBLOAD(Pbar As ProgressBar)
    Static a As Integer
    Dim b As Double
        Pbar.Value = 0
            Pbar.Max = 1000
                Pbar.Visible = True
                    For b = 0 To 999
                        Pbar.Value = Pbar.Value + 1
                    Next b
                Pbar.Value = 0
        Pbar.Visible = False
End Sub
Sub SENDtxt(TXT As TextBox)
    'SendKeys "{HOME}"
    TXT.SelStart = 0
    TXT.SelLength = Len(TXT.Text)
End Sub
