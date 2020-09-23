Attribute VB_Name = "modFuncChild"
Option Explicit

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

' Rectangle
Private Type RECT
   Left As Long     ' Left of the rectangle
   Top As Long      ' Top of the rectangle
   Right As Long    ' Right of the rectangle
   Bottom As Long   ' Bottom of the rectangle
End Type


Public Sub LoadForm(ByRef CFrm As Form, Optional CloseButton As Boolean = True)
    
    Dim r As RECT
    
    CFrm.Visible = False
    CFrm.WindowState = vbNormal

    GetClientRect MDIMainForm.hwnd, r
    
    'set client size
    'right
    If MDIMainForm.JoeSBCenter1.Visible = True Then
        r.Right = r.Right - (MDIMainForm.JoeSBCenter1.Width / Screen.TwipsPerPixelX)
    End If
    'bottom
    'r.Bottom = r.Bottom - ((MDIMainForm.JOEClientWin1.Height / Screen.TwipsPerPixelY) + (MDIMainForm.bgHeader.Height / Screen.TwipsPerPixelY)) - r.Top
    
    MDIMainForm.JOEClientWin1.LoadChildWindow MDIMainForm.hwnd, CFrm.hwnd, CFrm.Name, CFrm.Caption, r.Top, r.Left, r.Right, r.Bottom, CloseButton

    CFrm.Visible = True
    CFrm.Show
    CFrm.SetFocus
    
    ResizeMdiChildForm CFrm
End Sub

Public Sub ResizeMdiChildForm(ByRef CFrm As Form)

    Dim r As RECT
    
    GetClientRect MDIMainForm.hwnd, r
    
    'set client size
    'right
    If MDIMainForm.JoeSBCenter1.Visible = True Then
        r.Right = r.Right - (MDIMainForm.JoeSBCenter1.Width / Screen.TwipsPerPixelX) - 5
    End If
    'bottom
    'r.Bottom = r.Bottom - ((MDIMainForm.JOEClientWin1.Height / Screen.TwipsPerPixelY) + (MDIMainForm.bgHeader.Height / Screen.TwipsPerPixelY)) - r.Top - 5

    MDIMainForm.JOEClientWin1.ResizeClientWin CFrm.hwnd, r.Top, r.Left, r.Right, r.Bottom

End Sub

Public Sub ActivateMDIChildForm(ByVal sFormName As String)
    
    Dim frm As Form
    
    For Each frm In Forms
        If frm.Name = sFormName Then
            'activate form
            ResizeMdiChildForm frm
            frm.Visible = True
            frm.Show
            frm.SetFocus
            'set tab active window
            MDIMainForm.JOEClientWin1.SetActiveWindow sFormName
            Exit For
        End If
    Next
    
    
    Set frm = Nothing
End Sub

