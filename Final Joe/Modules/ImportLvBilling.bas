Attribute VB_Name = "ImportLvBilling"
Option Explicit



Private BfntOld As StdFont
'ImportDBGrid:
' This Sub reads the DBGrid specified by dbGrd into clsTP.
' rstData has to be set to the recordset dbGrd gets its data from (it seems to be impossible to get DataSource at runtime !???)
' (e.g. if it's bound to Data1, rstData should be Data1.Recordset)
Sub BImportListView(BclsTP As clsTablePrint, Blv As LynxGrid3, Optional ByVal sngDesiredWidth As Single = -1, Optional ByVal bWithIcons As Boolean = True)
    Dim BlRow As Long, BlCol As Long
    Dim BsngFXGGesWidth As Single
    
    BclsTP.Rows = Val(Blv.RowCount)
    BclsTP.Cols = Val(Blv.Cols + 1)
    BclsTP.HeaderRows = 1
    BclsTP.HasFooter = False
    BclsTP.LineThickness = 1 'frmDemo.GR.GridLineWidth  '2 'LV.GridLineWidth
    'Use double line width
    BclsTP.HeaderLineThickness = 2 * BclsTP.LineThickness

    'Set the row height
    BclsTP.RowHeightMin = Blv.RowHeightMin '  flxGrd.RowHeightMin
    BclsTP.FooterRowHeightMin = BclsTP.RowHeightMin
    BclsTP.HeaderRowHeightMin = BclsTP.RowHeightMin
    
    'Use some reasonable default values:
    BclsTP.CellXOffset = 60
    BclsTP.CellYOffset = 30
    BclsTP.CenterMergedHeader = False
    BclsTP.ResizeCellsToPicHeight = True
    BclsTP.PrintHeaderOnEveryPage = True
    
    
    'clsTP.HasFooter = True
    'clsTP.FooterText(1) = "sdsdfsdfsdf"
    'clsTP.TextMatrix
    'Set fntOld = New StdFont
    With Blv
        sngFXGGesWidth = 0
        For BlRow = 0 To 1 - 1
            For lCol = 0 To .Cols '- 1
                 .Col = BlCol
                 .Row = BlRow '+ .FixedRows
                'Set clsTP.HeaderFont(lRow, lCol) = lv.Font   'GetGridFont(LV)
                If (BlRow = 0) Then
                    If .CellAlignment(BlRow, BlCol) = lgAlignLeftCenter Then '  (lCol) '.CellAlignment
                        clsTP.ColAlignment(BlCol) = eLeft
                    ElseIf .CellAlignment(BlRow, BlCol) = lgAlignRightCenter Then
                        clsTP.ColAlignment(BlCol) = eRight
                    ElseIf .CellAlignment(BlRow, BlCol) = lgAlignCenterCenter Then
                        clsTP.ColAlignment(BlCol) = eCenter
                    End If
                    sngFXGGesWidth = sngFXGGesWidth + .ColWidth(BlCol)
                End If
                clsTP.HeaderText(BlRow, BlCol) = .ColHeading(BlCol) ' "JOehel" '.CellText(lRow, lCol) 'Text
            Next
           'clsTP.MergeHeaderRow(lRow) = frmDemo.GR.MergeRow(lRow)   'True    '.r me .MergeRow(lRow)
        Next
        For BlCol = 0 To .Cols '- 1
            For BlRow = 0 To .RowCount - 1 - 1
                .Col = BlCol
                .Row = BlRow + 1 '.FixedRows
                Set BclsTP.FontMatrix(BlRow, BlCol) = lv.Font
                'If Not (.ItemImage Is Nothing) Then
                '    If .CellPicture.handle <> 0 Then
                '        Set clsTP.PictureMatrix(lRow, lCol) = .CellPicture
                '    End If
                'End If
                BclsTP.TextMatrix(BlRow, BlCol) = .CellText(BlRow, BlCol) '.Text
                'clsTP.HeaderText(0, 0) = "JOehel"
                If (BlCol = 0) Then
                '    clsTP.MergeRow(lRow) = frmDemo.GR.MergeRow(lRow)
                End If
            Next
            If sngDesiredWidth > 0 Then
                BclsTP.ColWidth(BlCol) = (.ColWidth(BlCol) / sngFXGGesWidth) * sngDesiredWidth
            Else
                BclsTP.ColWidth(BlCol) = .ColWidth(BlCol)
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




