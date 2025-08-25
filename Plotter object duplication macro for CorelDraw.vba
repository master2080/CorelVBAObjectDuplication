Sub Duplicate()
    Dim horizontal_gap_MM As Double
    Dim vertical_gap_MM As Double
    Dim page_border_right_MM As Double
    Dim page_border_bottom_MM As Double
    Dim maxObjectsBeforeBitmap As Double
    ' =======================
    ' CONFIGURATION SECTION
    ' =======================
    horizontal_gap_MM = 5           ' Horizontal gap between items in mm
    vertical_gap_MM = 5             ' Vertical gap between items in mm
    page_border_left_MM = 13        ' Position in X to move the element to before duplicating
    page_border_right_MM = 13       ' Gap from right edge of page in mm
    page_border_top_MM = 20         ' Position in Y to move the element to before duplicating
    page_border_bottom_MM = 11      ' Gap from bottom edge of page in mm
    maxObjectsBeforeBitmap = 100    ' Past this amount turn the objects(other than the outline) into a bitmap
    ' =======================
    ' END CONFIGURATION
    ' =======================
    
    Dim doc As Document
    Dim pg As Page
    Dim sr As ShapeRange, srWithoutMagenta As New ShapeRange
    Dim grp As Shape
    Dim shp As Shape, oc As Outline, col As Color
    Dim bmp As Shape
    Dim startX As Double, startY As Double
    Dim rightLimit As Double, bottomLimit As Double
    Dim gapH As Double, gapV As Double
    Dim minRightGap As Double, minBottomGap As Double
    
    Set doc = ActiveDocument
    Set pg = doc.ActivePage
    
    ' Get selection
    Set sr = ActiveSelectionRange
    If sr.Count = 0 Then
        MsgBox "Please select one or more objects first."
        Exit Sub
    End If
        
    If sr.Count > maxObjectsBeforeBitmap Then ' Too many objects, convert all of them to a bitmap
        For Each shp In sr
            Dim isMagenta As Boolean
            isMagenta = False
                If Not shp.Outline Is Nothing Then
                    If shp.Outline.Width > 0 Then
                        Set oc = shp.Outline
                        Set col = oc.Color
                        If col.Type = cdrColorCMYK Then
                            If col.CMYKCyan = 0 And col.CMYKMagenta = 100 And col.CMYKYellow = 0 And col.CMYKBlack = 0 Then
                                isMagenta = True
                            End If
                        End If
                    End If
                End If
            If Not isMagenta Then
                srWithoutMagenta.Add shp
            End If
        Next shp
        
        ' Parameters: Image type, Dithered?, Transparent?, Resolution dpi, Anti aliasing type[cdrAntiAliasingType], Use color profile(icc?), AlwaysOverprintBlack, OverprintBlackLimit
        Set bmp = srWithoutMagenta.ConvertToBitmapEx(cdrCMYKColorImage, False, True, 600, cdrNoAntiAliasing, True, False, 0)
        srWithoutMagenta.Delete ' Delete the now obsolete elements that were converted into a bitmap
        bmp.AddToSelection ' The selection will now contain the magenta outline + the bitmap
        Set sr = ActiveSelectionRange ' Set it to the active selection again after these updates, should have magenta outline+bmp
    End If
    
    ' Temporarily group selection so we can treat it as one unit
    Set grp = sr.Group
    
    ' Convert mm to doc units
    gapH = doc.ToUnits(horizontal_gap_MM, cdrMillimeter)
    gapV = doc.ToUnits(vertical_gap_MM, cdrMillimeter)
    minRightGap = doc.ToUnits(page_border_right_MM, cdrMillimeter)
    minBottomGap = doc.ToUnits(page_border_bottom_MM, cdrMillimeter)
    page_border_left_MM = doc.ToUnits(page_border_left_MM, cdrMillimeter)
    page_border_top_MM = doc.ToUnits(page_border_top_MM, cdrMillimeter)
    
    ' Page limits
    rightLimit = pg.SizeWidth - minRightGap
    bottomLimit = minBottomGap
    
    Dim count0 As Long, count90 As Long
    
    grp.RotationAngle = 0
    count0 = CountFit(grp, pg, gapH, gapV, page_border_left_MM, page_border_top_MM, rightLimit, bottomLimit)
    
    ' Test with rotation 90
    grp.RotationAngle = 90
    count90 = CountFit(grp, pg, gapH, gapV, page_border_left_MM, page_border_top_MM, rightLimit, bottomLimit)
    
    If count90 > count0 Then
        grp.RotationAngle = 90
    Else
        grp.RotationAngle = 0
    End If
    
    ' Move element to a top left position based on requirements
    grp.LeftX = page_border_left_MM
    grp.TopY = (pg.TopY - page_border_top_MM)
    
    ' Starting position
    startX = grp.LeftX
    startY = grp.TopY
    
    ' Duplicate horizontally
    Dim rowShapes As New ShapeRange
    rowShapes.Add grp
    
    Dim x As Double
    x = startX + grp.SizeWidth + gapH
    Do While (x + grp.SizeWidth) <= rightLimit
        Dim newGrp As Shape
        Set newGrp = grp.Duplicate
        newGrp.LeftX = x
        newGrp.TopY = startY
        rowShapes.Add newGrp
        x = x + grp.SizeWidth + gapH
    Loop
    
    ' Group the row
    Dim rowGroup As Shape
    Set rowGroup = rowShapes.Group
    
    ' Duplicate vertically
    Dim currentY As Double
    currentY = startY - rowGroup.SizeHeight - gapV
    
    Dim rowCopies As New ShapeRange
    rowCopies.Add rowGroup
    
    Do While (currentY - rowGroup.SizeHeight) >= bottomLimit
        Dim newRow As Shape
        Set newRow = rowGroup.Duplicate
        newRow.TopY = currentY
        newRow.LeftX = startX
        rowCopies.Add newRow
        currentY = currentY - rowGroup.SizeHeight - gapV
    Loop
    
    ' Ungroup all rows (including the first one)

    For Each shp In rowCopies
        If shp.Type = cdrGroupShape Then
            shp.UngroupAll
        End If
    Next shp
    
    ' Find all magenta elements + group them
    Dim magentaShapes As New ShapeRange
    For Each shp In pg.Shapes
        If Not shp.Outline Is Nothing Then
            If shp.Outline.Width > 0 Then
                Dim c As Color
                Set c = shp.Outline.Color
                If c.Type = cdrColorCMYK Then ' Only run over CMYK objects
                    If c.CMYKCyan = 0 And c.CMYKMagenta = 100 And c.CMYKYellow = 0 And c.CMYKBlack = 0 Then ' Magenta
                        magentaShapes.Add shp
                    End If
                End If
            End If
        End If
    Next shp
    
    If magentaShapes.Count > 0 Then
        magentaShapes.Group
    End If
    
    ' Unselect whatever is selected
    doc.ClearSelection
    
    ' Center everything on the page(group+ H center) and move it to the bottom of the page( minus the bottom gap), then ungroup
    pg.Shapes.All.CreateSelection
    Dim allGroup As Shape
    Set allGroup = ActiveSelection.Group
    allGroup.AlignToPageCenter cdrAlignHCenter
    allGroup.BottomY = minBottomGap
    allGroup.Ungroup
    
    ' Undo optimization + refresh the screen for changes
    
End Sub

Private Function CountFit(ByVal grp As Shape, ByVal pg As Page, _
                        ByVal gapH As Double, ByVal gapV As Double, _
                        ByVal page_border_left_MM As Double, ByVal page_border_top_MM As Double, _
                        ByVal rightLimit As Double, ByVal bottomLimit As Double) As Long
    Dim startX As Double, startY As Double
    Dim x As Double, y As Double
    Dim cols As Long, rows As Long
    
    ' Move to top-left starting position
    grp.LeftX = page_border_left_MM
    grp.TopY = (pg.TopY - page_border_top_MM)
    
    startX = grp.LeftX
    startY = grp.TopY
    
    ' How many fit horizontally?
    cols = 1
    x = startX + grp.SizeWidth + gapH
    Do While (x + grp.SizeWidth) <= rightLimit
        cols = cols + 1
        x = x + grp.SizeWidth + gapH
    Loop
    
    ' How many fit vertically?
    rows = 1
    y = startY - grp.SizeHeight - gapV
    Do While (y - grp.SizeHeight) >= bottomLimit
        rows = rows + 1
        y = y - grp.SizeHeight - gapV
    Loop
    
    CountFit = cols * rows
End Function





