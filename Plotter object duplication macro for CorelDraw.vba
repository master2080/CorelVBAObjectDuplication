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
    ByVal isSplitMode As Boolean, _
    ByVal gap_split_MM As Double, _
    ByVal gap_distance_MM As Double)
   
    Dim doc As Document
    Dim pg As Page
    Dim sr As ShapeRange, newSr As New ShapeRange
    Dim totalObjects As Long
    Dim grp As Shape
    Dim shp As Shape, oc As Outline, col As Color
    Dim bmp As Shape
    Dim rightLimit As Double, bottomLimit As Double, leftLimit As Double, topLimit As Double
    Dim gapH As Double, gapV As Double
    Dim gapSplit As Double, gapDistance As Double
    Dim minRightGap As Double, bottomLimitOrigin As Double
    Dim marker_distance_X As Double, marker_distance_Y As Double, marker_size As Double
    Dim markerAmountThreshold_CM As Double
    Dim markerAmountThreshold As Double
    Dim marker_count As Double
    Dim refPoint As cdrReferencePoint
    Dim i As Double
    dim splitMode As Boolean ' Splitting the objects apart with a gap after every #mm 
   
    Set doc = ActiveDocument
    Set pg = doc.ActivePage
    refPoint = doc.ReferencePoint ' Store current reference point for restoration later
    ' Get selection
    Set sr = ActiveSelectionRange
   
    If sr.Count < 2 Then
        MsgBox "Please select two or more objects first."
        Exit Sub
    End If
    markerAmountThreshold_CM = 23 ' For every additional X cm of a page, add 2 more markers, with an always minimum of 4
    markerAmountThreshold = doc.ToUnits(markerAmountThreshold_CM, cdrCentimeter)
    marker_count = 2 + (2 * Int(pg.SizeHeight / markerAmountThreshold))
   
    totalObjects = CountShapes(sr)
   
    Dim magShp As Shape
   
    ' find the magenta object
    For Each shp In sr
        If Not shp.Outline Is Nothing Then
            If shp.Outline.Width > 0 Then
                Set col = shp.Outline.Color
                If col.Type = cdrColorCMYK Then
                    If col.CMYKCyan = 0 And col.CMYKMagenta = 100 And col.CMYKYellow = 0 And col.CMYKBlack = 0 Then
                        Set magShp = shp
                        Exit For
                    End If
                End If
            End If
        End If
    Next shp
       
    If totalObjects > maxObjectsBeforeBitmap Then ' Too many objects, convert all of them(except the magenta outline) to a bitmap
        ' turn sr into bitmap
        Set newSr = CreateShapeRange
        newSr.Add magShp
        sr.Remove sr.IndexOf(magShp)
        ' Parameters: Image type, Dithered?, Transparent?, Resolution dpi, Anti aliasing type[cdrAntiAliasingType], Use color profile(icc?), AlwaysOverprintBlack, OverprintBlackLimit
        Set bmp = sr.ConvertToBitmapEx(cdrCMYKColorImage, False, True, 600, cdrNoAntiAliasing, True, False, 0)
        sr.Delete ' Delete the now obsolete elements that were converted into a bitmap
        newSr.Add bmp
        newSr.AddToSelection
        Set sr = newSr 
    End If
   
    ' Temporarily group selection so we can treat it as one unit
    Set grp = sr.Group
   
    ' Convert mm to doc units, prevents wrong units being used in code
    gapH = doc.ToUnits(horizontal_gap_MM, cdrMillimeter)
    gapV = doc.ToUnits(vertical_gap_MM, cdrMillimeter)
    gapSplit = doc.ToUnits(gap_split_MM, cdrMillimeter)
    minRightGap = doc.ToUnits(page_border_right_MM, cdrMillimeter)
    bottomLimitOrigin = doc.ToUnits(page_border_bottom_MM, cdrMillimeter)
    leftLimit = doc.ToUnits(page_border_left_MM, cdrMillimeter)
    topLimit = doc.ToUnits(page_border_top_MM, cdrMillimeter)
    marker_distance_X = doc.ToUnits(marker_distance_X_MM, cdrMillimeter)
    marker_distance_Y = doc.ToUnits(marker_distance_Y_MM, cdrMillimeter)
    marker_size = doc.ToUnits(marker_size_MM, cdrMillimeter)
    splitMode = IsSplitMode
    gapDistance = doc.ToUnits(gap_distance_MM, cdrMillimeter)
   
    ' Page limits(including markers, markers reduce the workable area)
    leftLimit = leftLimit + marker_distance_X + marker_size
    rightLimit = pg.SizeWidth - (minRightGap + marker_distance_X + marker_size)
    topLimit = pg.TopY - (topLimit + marker_distance_Y + marker_size)
    bottomLimit = bottomLimitOrigin + marker_distance_Y + marker_size
   
    Dim magWidth As Double, magHeight As Double
    Dim magOffsetX As Double, magOffsetY As Double

    magWidth = magShp.SizeWidth
    magHeight = magShp.SizeHeight
    ' Will be used to move the object group to the right position as it needs to be based on the magenta, even if the graphic is larger
    magOffsetX = Abs(magShp.LeftX - grp.LeftX)
    magOffsetY = Abs(grp.TopY - magShp.TopY)

    ' Move element to the bottom left position based on requirements and magenta outline
    grp.LeftX = leftLimit - magOffsetX
    grp.BottomY = bottomLimit - magOffsetY
      
    Dim count0hz As Long, count0vr As Long, count90hz As Long, count90vr As Long
    Dim count0 As Long, count90 As Long
    Dim horizontalCount As Long, verticalCount As Long
   
    ' magShp.RotationAngle = 0
    count0hz = CountFit(magShp.SizeWidth, gapV, leftLimit, rightLimit)
    count0vr = CountFit(magShp.SizeHeight, gapH, topLimit, bottomLimit)
    count0 = count0hz * count0vr
    
    ' Test with rotation 90 - swap widths and heights around
    ' magShp.RotationAngle = 90
    count90hz = CountFit(magShp.SizeHeight, gapV, leftLimit, rightLimit)
    count90vr = CountFit(magShp.SizeWidth, gapH, topLimit, bottomLimit)
    count90 = count90hz * count90vr
   
    If count90 > count0 Then
        grp.RotationAngle = 90
        magShp.RotationAngle = 90
        ' recalculate the magenta offsets due to the new rotation
        magWidth = magShp.SizeWidth
        magHeight = magShp.SizeHeight
        magOffsetX = Abs(magShp.LeftX - grp.LeftX)
        magOffsetY = Abs(grp.TopY - magShp.TopY)
        
        horizontalCount = count90hz
        verticalCount = count90vr
    Else
        grp.RotationAngle = 0
        magShp.RotationAngle = 0
        horizontalCount = count0hz
        verticalCount = count0vr
    End If
   
    ' Duplicate horizontally
    Dim rowShapes As New ShapeRange
    rowShapes.Add grp
   
    Dim nextX As Double
    nextX = grp.LeftX + grp.SizeWidth + gapH

    for X = 2 To horizontalCount
        Dim newGrp As Shape
        Set newGrp = grp.Duplicate

        newGrp.LeftX = nextX
        newGrp.BottomY = grp.BottomY
        rowShapes.Add newGrp

        nextX = nextX + grp.SizeWidth + gapH
    Next X 

    ' Group the row
    Dim rowGroup As Shape
    Set rowGroup = rowShapes.Group

    Dim newRow As Shape

    ' Duplicate vertically
    If splitMode Then
        ' Figure out how many fit within a split block, if splitting is set to true
        Dim nextY As Double, numInBlock As Long, blockCount As Long
        nextY = grp.bottomY + magHeight + gapV
        numInBlock = CountFit(rowGroup.SizeHeight, gapV, 0, gapDistance)
        ' Most of the time there is going to be spare room, shrink block size to fit new size
        If (((rowGroup.SizeHeight + gapV) * numInBlock) - gapV) < gapDistance Then
            gapDistance = ((rowGroup.SizeHeight + gapV) * numInBlock) - gapV
        End If
        blockCount = CountFit(gapDistance, gapSplit, bottomLimit, topLimit)

        Dim blockGroupItems As New ShapeRange
        blockGroupItems.Add rowGroup

        For X = 2 To numInBlock
            Set newRow = rowGroup.Duplicate

            newRow.BottomY = nextY
            newRow.LeftX = grp.LeftX

            blockGroupItems.Add newRow

            nextY = nextY + magHeight + gapV
        Next X

        Dim blockGroup As Shape
        Set blockGroup = blockGroupItems.Group
        Dim blocks As New ShapeRange
        
        blocks.Add blockGroup

        nextY = blockGroup.bottomY + blockGroup.SizeHeight + gapSplit

        For X = 2 To blockCount
            Dim newBlock As Shape
            Set newBlock = blockGroup.Duplicate

            newBlock.bottomY = nextY
            newBlock.LeftX = grp.LeftX

            blocks.Add newBlock

            nextY = nextY + blockGroup.SizeHeight + gapSplit
        Next X

        ' Ungroup everything
        Blocks.UngroupAll
        
    Else 
        ' fill out the entire working area
        Dim rowCopies As New ShapeRange
        rowCopies.Add rowGroup
    
        Dim currentY As Double
        nextY = grp.bottomY + magHeight + gapV

        Do While (nextY + magHeight) <= topLimit
            Set newRow = rowGroup.Duplicate
            newRow.BottomY = nextY
            newRow.LeftX = grp.LeftX
            rowCopies.Add newRow
            nextY = nextY + magHeight + gapV
        Loop

        ' Ungroup all rows
        For Each shp In rowCopies
            If shp.Type = cdrGroupShape Then
                shp.UngroupAll
            End If
        Next shp
    End If
           
    ' Find all magenta elements + group them
    Dim magentaShapes As New ShapeRange
    For Each shp In pg.shapes
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
   
   
    doc.ReferencePoint = cdrCenter ' Set ref point to center for marker placement
   
    ' Add OPOS markers based on the magenta group specifically, if it exists
    If Not magentaGroup Is Nothing Then
        Dim halfSize As Double, rows As Double
        Dim rect As Shape, dup As Shape
        Dim xLeft As Double, xRight As Double
        Dim stepY As Double
        Dim coords As Collection
        halfSize = marker_size / 2 ' SetPosition relies on the center point for placement, based on doc.ReferencePoint
        rows = marker_count / 2 ' Amount of markers vertically, i.e 8 means 4 rows of vertical markers(2 corners and 2 middle ones)
        xLeft = magentaGroup.LeftX - marker_distance_X - halfSize ' Center X position of the left column
        xRight = magentaGroup.RightX + marker_distance_X + halfSize ' Center X position of the right column
        ' Get total distance between markers(their centers), figure out where to place each marker(with an equal distance)
        stepY = (magentaGroup.TopY + marker_distance_Y + halfSize) - (magentaGroup.BottomY - marker_distance_Y - halfSize)
        stepY = stepY / (rows - 1)
       
        Set coords = New Collection
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
    ' Center everything on the page(group+ H center) and move it to the bottom of the page( minus the bottom gap), then ungroup
    pg.shapes.All.CreateSelection
    Dim allGroup As Shape
    Set allGroup = ActiveSelection.Group
    allGroup.AlignToPageCenter cdrAlignHCenter
    allGroup.BottomY = bottomLimitOrigin - (marker_distance_Y + marker_size) ' The actual bottom limit from the entire thing, markers included
    allGroup.Ungroup
   
    ' Restore reference point
    doc.ReferencePoint = refPoint
   
End Sub
Public Sub Duplicate()
    NaklForm.Show
End Sub

Private Function CountFit(ByVal size As Double, ByVal gap As Double, ByVal fromLimit As Double, ByVal toLimit As Double) As Long
    Dim totalLength As Double
    totalLength = Abs(toLimit - fromLimit)
    If (size + gap) <= 0 Or totalLength < size Then
        CountFit = 0
        Exit Function
    End If
    ' return the truncated number
    CountFit = Fix((totalLength + gap) / (size + gap))
End Function

Function CountShapes(shapes As ShapeRange) As Long
    Dim s As Shape
    Dim total As Long
    total = 0
   
    For Each s In shapes
        If s.Type = cdrGroupShape Then
            total = total + CountShapes(s.shapes.All)
        Else
            total = total + 1
        End If
    Next s
   
    CountShapes = total
   
End Function 