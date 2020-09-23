Attribute VB_Name = "mdlImportLVTR"
Option Explicit


Global TMPHTxt As String

Private fntOld As StdFont
'ImportDBGrid:
' This Sub reads the DBGrid specified by dbGrd into clsTP.
' rstData has to be set to the recordset dbGrd gets its data from (it seems to be impossible to get DataSource at runtime !???)
' (e.g. if it's bound to Data1, rstData should be Data1.Recordset)
Sub ImportListView(TMPHeadertxt As String, clsTP As clsTablePrint, lv As LynxGrid3, Optional ByVal sngDesiredWidth As Single = -1, Optional ByVal bWithIcons As Boolean = True)
    Dim lRow As Long, lCol As Long
    Dim sngFXGGesWidth As Single
    
    'Call Class_Initialize
    
    clsTP.Rows = Val(lv.RowCount)
    clsTP.Cols = Val(lv.Cols + 1)
    clsTP.HeaderRows = 2
    
    clsTP.HasFooter = False
    clsTP.LineThickness = 1 'frmDemo.GR.GridLineWidth  '2 'LV.GridLineWidth
    'Use double line width
    clsTP.HeaderLineThickness = 2 * clsTP.LineThickness

    'Set the row height
    clsTP.RowHeightMin = lv.RowHeightMin '  flxGrd.RowHeightMin
    clsTP.FooterRowHeightMin = clsTP.RowHeightMin
    clsTP.HeaderRowHeightMin = clsTP.RowHeightMin
    
    'Use some reasonable default values:
    clsTP.CellXOffset = 60
    clsTP.CellYOffset = 30
    clsTP.CenterMergedHeader = False
    clsTP.ResizeCellsToPicHeight = True
    clsTP.PrintHeaderOnEveryPage = True
    
    
    'clsTP.HasFooter = True
    'clsTP.FooterText(1) = "sdsdfsdfsdf"
    'clsTP.TextMatrix
    'Set fntOld = New StdFont
    With lv
        sngFXGGesWidth = 0
        For lRow = 0 To 1
            For lCol = 0 To .Cols '- 1
                 .Col = lCol
                 .Row = lRow '+ .FixedRows
                'Set clsTP.HeaderFont(lRow, lCol) = lv.Font   'GetGridFont(LV)
                If (lRow = 0) Then
                    If .CellAlignment(lRow, lCol) = lgAlignLeftCenter Then '  (lCol) '.CellAlignment
                        clsTP.ColAlignment(lCol) = eLeft
                    ElseIf .CellAlignment(lRow, lCol) = lgAlignRightCenter Then
                        clsTP.ColAlignment(lCol) = eRight
                    ElseIf .CellAlignment(lRow, lCol) = lgAlignCenterCenter Then
                        clsTP.ColAlignment(lCol) = eCenter
                    End If
                    sngFXGGesWidth = sngFXGGesWidth + .ColWidth(lCol)
                End If
                
                If lRow = 0 Then
                    clsTP.HeaderText(0, 2) = TMPHeadertxt '"" '
                    clsTP.HeaderText(0, .Cols - 1) = "Page No:"
                    'clsTP.HeaderText(0, .Cols - 1) = PgNum
                ElseIf lRow = 1 Then
                    clsTP.HeaderText(1, lCol) = .ColHeading(lCol)
                End If
                'clsTP.ColHeading (lCol) ' "JOehel" '.CellText(lRow, lCol) 'Text
            Next
           'clsTP.MergeHeaderRow(lRow) = frmDemo.GR.MergeRow(lRow)   'True    '.r me .MergeRow(lRow)
        Next
        For lCol = 0 To .Cols '- 1
            For lRow = 0 To .RowCount - 1 '- 1
                .Col = lCol
                .Row = lRow  '.FixedRows
                
                'Set clsTP.FontMatrix(lRow, lCol) = lv.Font
                
                clsTP.TextMatrix(lRow, lCol) = .CellText(lRow, lCol)
                
                
            Next
            If sngDesiredWidth > 0 Then
                clsTP.ColWidth(lCol) = (.ColWidth(lCol) / sngFXGGesWidth) * sngDesiredWidth
            Else
                clsTP.ColWidth(lCol) = .ColWidth(lCol)
            End If
            'clsTP.MergeCol(lCol) = frmDemo.GR.MergeCol(lCol)
            'clsTP.MergeHeaderCol(lCol) = frmDemo.GR.MergeCol(lCol)
        Next
         
    End With
End Sub

'Helper Function for ImportFlexGrid()
Private Function GetGridFont(flxGrd As LynxGrid3) As StdFont
    Dim bDiff As Boolean
    
    'If fntOld Is Nothing Then bDiff = True: GoTo DiffCheck
    'Font styles:
    'bDiff = bDiff Or (flxGrd.CellFontBold <> fntOld.Bold) Or _
            (flxGrd.CellFontItalic <> fntOld.Italic) Or (flxGrd.CellFontUnderline <> fntOld.Underline) 'Or _
            '(flxGrd.CellFontStrikeThrough <> fntOld.Strikethrough)
    'Name:
    'bDiff = bDiff Or (flxGrd.Font <> fntOld.Name)
    'Size:
    'bDiff = bDiff Or (flxGrd.fon CellFontSize <> fntOld.Size)
DiffCheck:
    If bDiff Then
        Set fntOld = New StdFont
        'fntOld.Name = flxGrd.CellFontName
        'fntOld.Size = flxGrd.CellFontSize
        'fntOld.Bold = flxGrd.CellFontBold
        'fntOld.Italic = flxGrd.CellFontItalic
        'fntOld.Underline = flxGrd.CellFontUnderline
        'fntOld.Strikethrough = flxGrd.CellFontStrikeThrough
    End If
    Set GetGridFont = fntOld
End Function


