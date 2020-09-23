Attribute VB_Name = "mdlImportLVPR"
Option Explicit



Private PRfntOld As StdFont
'ImportDBGrid:
' This Sub reads the DBGrid specified by dbGrd into clsTP.
' rstData has to be set to the recordset dbGrd gets its data from (it seems to be impossible to get DataSource at runtime !???)
' (e.g. if it's bound to Data1, rstData should be Data1.Recordset)
Sub PRImportListView(PRclsTP As clsTablePrint, PRlv As LynxGrid3, Optional ByVal PRsngDesiredWidth As Single = -1, Optional ByVal PRbWithIcons As Boolean = True)
    Dim PRlRow As Long, PRlCol As Long
    Dim PRsngFXGGesWidth As Single
    
    PRclsTP.Rows = Val(PRlv.RowCount)
    PRclsTP.Cols = Val(PRlv.Cols + 1)
    PRclsTP.HeaderRows = 1
    PRclsTP.HasFooter = False
    
    PRclsTP.LineThickness = 1 'frmDemo.GR.GridLineWidth  '2 'LV.GridLineWidth
    'Use double line width
    PRclsTP.HeaderLineThickness = 2 * PRclsTP.LineThickness

    'Set the row height
    PRclsTP.RowHeightMin = PRlv.RowHeightMin '  flxGrd.RowHeightMin
    PRclsTP.FooterRowHeightMin = PRclsTP.RowHeightMin
    PRclsTP.HeaderRowHeightMin = PRclsTP.RowHeightMin
    
    'Use some reasonable default values:
    PRclsTP.CellXOffset = 60
    PRclsTP.CellYOffset = 30
    PRclsTP.CenterMergedHeader = False
    PRclsTP.ResizeCellsToPicHeight = True
    PRclsTP.PrintHeaderOnEveryPage = True
    
    'Set fntOld = New StdFont
    With PRlv
        PRsngFXGGesWidth = 0
        For PRlRow = 0 To 1 - 1
            For PRlCol = 0 To .Cols '- 1
                 .Col = PRlCol
                 .Row = PRlRow '+ .FixedRows
                'Set clsTP.HeaderFont(lRow, lCol) = lv.Font   'GetGridFont(LV)
                If (PRlRow = 0) Then
                    If .CellAlignment(PRlRow, PRlCol) = lgAlignLeftCenter Then '  (lCol) '.CellAlignment
                        PRclsTP.ColAlignment(PRlCol) = eLeft
                    ElseIf .CellAlignment(PRlRow, PRlCol) = lgAlignRightCenter Then
                        PRclsTP.ColAlignment(PRlCol) = eRight
                    ElseIf .CellAlignment(PRlRow, PRlCol) = lgAlignCenterCenter Then
                        PRclsTP.ColAlignment(PRlCol) = eCenter
                    End If
                    PRsngFXGGesWidth = PRsngFXGGesWidth + .ColWidth(PRlCol)
                End If
                PRclsTP.HeaderText(PRlRow, PRlCol) = .ColHeading(PRlCol) ' "JOehel" '.CellText(lRow, lCol) 'Text
            Next
           'clsTP.MergeHeaderRow(lRow) = frmDemo.GR.MergeRow(lRow)   'True    '.r me .MergeRow(lRow)
        Next
        For PRlCol = 0 To .Cols '- 1
            For PRlRow = 0 To .RowCount - 1 - 1
                .Col = PRlCol
                .Row = PRlRow + 1 '.FixedRows
                Set PRclsTP.FontMatrix(PRlRow, PRlCol) = PRlv.Font
                'If Not (.ItemImage Is Nothing) Then
                '    If .CellPicture.handle <> 0 Then
                '        Set clsTP.PictureMatrix(lRow, lCol) = .CellPicture
                '    End If
                'End If
                PRclsTP.TextMatrix(PRlRow, PRlCol) = .CellText(PRlRow, PRlCol) '.Text
                'clsTP.HeaderText(0, 0) = "JOehel"
                If (PRlCol = 0) Then
                '    clsTP.MergeRow(lRow) = frmDemo.GR.MergeRow(lRow)
                End If
            Next
            If PRsngDesiredWidth > 0 Then
                PRclsTP.ColWidth(PRlCol) = (.ColWidth(PRlCol) / PRsngFXGGesWidth) * PRsngDesiredWidth
            Else
                PRclsTP.ColWidth(PRlCol) = .ColWidth(PRlCol)
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
        Set PRfntOld = New StdFont
        'fntOld.Name = flxGrd.CellFontName
        'fntOld.Size = flxGrd.CellFontSize
        'fntOld.Bold = flxGrd.CellFontBold
        'fntOld.Italic = flxGrd.CellFontItalic
        'fntOld.Underline = flxGrd.CellFontUnderline
        'fntOld.Strikethrough = flxGrd.CellFontStrikeThrough
    End If
    Set GetGridFont = PRfntOld
End Function


