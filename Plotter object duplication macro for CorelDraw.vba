Public Sub RunDuplicate( _
    ByVal horizontal_gap_MM As Double, _
    ByVal vertical_gap_MM As Double, _
    ByVal page_border_left_MM As Double, _
    ByVal page_border_right_MM As Double, _
    ByVal page_border_top_MM As Double, _
    ByVal page_border_bottom_MM As Double, _
    ByVal maxObjectsBeforeBitmap As Double, _
    ByVal marker_distance_X_MM As Double, _
    ByVal marker_distance_Y_MM As Double, _
    ByVal marker_size_MM As Double, _
    ByVal marker_count As Double)
    
    Dim doc As Document
    Dim pg As Page
    Dim sr As ShapeRange, srWithoutMagenta As New ShapeRange
    Dim grp As Shape
    Dim shp As Shape, oc As Outline, col As Color
    Dim bmp As Shape
    Dim startX As Double, startY As Double
    Dim rightLimit As Double, bottomLimit As Double, page_border_left As Double, page_border_top As Double
    Dim gapH As Double, gapV As Double
    Dim minRightGap As Double, minBottomGap As Double
    Dim marker_distance_X As Double, marker_distance_Y As Double, marker_size As Double
    
    Set doc = ActiveDocument
    Set pg = doc.ActivePage
    
    If marker_count < 4 Then
        MsgBox "Marker count should be 4 or more"
        Exit Sub
    End If

    If Not marker_count Mod 2 = 0 Then
        MsgBox "Marker amount should be even(4,6,8...)"
        Exit Sub
    End If

    ' Get selection
    Set sr = ActiveSelectionRange
    If sr.Count = 0 Then
        MsgBox "Please select one or more objects first."
        Exit Sub
    End If
        
    If sr.Count > maxObjectsBeforeBitmap Then ' Too many objects, convert all of them(except the magenta outline) to a bitmap
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
    
    ' Convert mm to doc units, prevents wrong units being used in code
    gapH = doc.ToUnits(horizontal_gap_MM, cdrMillimeter)
    gapV = doc.ToUnits(vertical_gap_MM, cdrMillimeter)
    minRightGap = doc.ToUnits(page_border_right_MM, cdrMillimeter)
    minBottomGap = doc.ToUnits(page_border_bottom_MM, cdrMillimeter)
    page_border_left = doc.ToUnits(page_border_left_MM, cdrMillimeter)
    page_border_top = doc.ToUnits(page_border_top_MM, cdrMillimeter)
    marker_distance_X = doc.ToUnits(marker_distance_X_MM, cdrMillimeter)
    marker_distance_Y = doc.ToUnits(marker_distance_Y_MM, cdrMillimeter)
    marker_size = doc.ToUnits(marker_size_MM, cdrMillimeter)
    
    ' Page limits
    rightLimit = pg.SizeWidth - minRightGap
    bottomLimit = minBottomGap
    
    Dim count0 As Long, count90 As Long
    
    grp.RotationAngle = 0
    count0 = CountFit(grp, pg, gapH, gapV, page_border_left, page_border_top, rightLimit, bottomLimit)
    
    ' Test with rotation 90
    grp.RotationAngle = 90
    count90 = CountFit(grp, pg, gapH, gapV, page_border_left, page_border_top, rightLimit, bottomLimit)
    
    If count90 > count0 Then
        grp.RotationAngle = 90
    Else
        grp.RotationAngle = 0
    End If
    
    ' Move element to a top left position based on requirements
    grp.LeftX = page_border_left
    grp.TopY = (pg.TopY - page_border_top)
    
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
    
    Dim magentaGroup As Shape
    If magentaShapes.Count > 0 Then
        Set magentaGroup = magentaShapes.Group
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
    
    ' Add OPOS markers based on the magenta group specifically, if it exists
    If Not magentaGroup Is Nothing Then
        Dim halfSize As Double, rows As Double
        Dim rect As Shape, dup As Shape
        Dim xLeft As Double, xRight As Double
        Dim stepY As Double
        Dim coords As Collection

        halfSize = marker_size / 2 ' SetPosition relies on the center point for movement
        rows = marker_count / 2 ' Amount of markers vertically, i.e 8 means 4 rows of vertical markers(2 corners and 2 middle ones)
        xLeft = magentaGroup.LeftX - marker_distance_X - halfSize ' Center X position of the left column
        xRight = magentaGroup.RightX + marker_distance_X + halfSize ' Center X position of the right column

        ' Get total distance between markers(their centers), figure out where to place each marker(with an equal distance)
        stepY = (magentaGroup.TopY + marker_distance_Y + halfSize) - (magentaGroup.BottomY - marker_distance_Y - halfSize)
        stepY = stepY / (rows - 1)

        
        Set coords = New Collection
        Dim i As Double
        For i = 0 To rows - 1
            Dim yPos As Double
            yPos = (magentaGroup.TopY + marker_distance_Y + halfSize) - (i * stepY)
            coords.Add Array(xLeft, yPos)
            coords.Add Array(xRight, yPos)
        Next i
        
        ' Base rectangle used for all
        Set rect = pg.ActiveLayer.CreateRectangle2(0, 0, marker_size, marker_size) ' X, Y, Width, Height
        rect.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
        rect.Outline.SetNoOutline

        For i = 1 To coords.Count
            Set dup = rect.Duplicate
            dup.SetPosition coords(i)(0), coords(i)(1)
        Next i

        ' Delete the base rectangle
        rect.Delete

    End If
End Sub

Public Sub Duplicate()
    ' RunDuplicate 5, 5, 13, 13, 20, 11, 100, 4, 4, 3, 4
    NaklForm.Show
End Sub


Private Function CountFit(ByVal grp As Shape, ByVal pg As Page, _
                        ByVal gapH As Double, ByVal gapV As Double, _
                        ByVal page_border_left As Double, ByVal page_border_top As Double, _
                        ByVal rightLimit As Double, ByVal bottomLimit As Double) As Long
    Dim startX As Double, startY As Double
    Dim x As Double, y As Double
    Dim cols As Long, rows As Long
    
    ' Move to top-left starting position
    grp.LeftX = page_border_left
    grp.TopY = (pg.TopY - page_border_top)
    
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



