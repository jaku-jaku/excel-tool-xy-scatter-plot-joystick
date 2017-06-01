VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Chart1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim temp As Double
Dim isZoomIn As Boolean
Dim isToggleTheore As Boolean
Dim isToggleViz As Boolean
Private Const RATE = 0.2
Private Const GRIDNUM = 10
Private Const interMarkSize = 3
Private Const beginIndTheore = 3
Private Const EndIndTheore = 18
Private Const beginIndViz = 19
Private Const EndIndViz = 34


Private Sub Chart_Activate()
    IntersectionCalc beginIndTheore, EndIndTheore, "IntersectionsTheore"
    IntersectionCalc beginIndViz, EndIndViz, "IntersectionsViztool"
    CreateIt
End Sub

Private Sub Chart_BeforeDoubleClick(ByVal ElementID As Long, ByVal Arg1 As Long, ByVal Arg2 As Long, Cancel As Boolean)
 Cancel = True
End Sub

Private Sub Chart_Deactivate()
    KillIt
End Sub


Private Sub Chart_BeforeRightClick(Cancel As Boolean)
    Cancel = True
    If isZoomIn Then
        zoomInToggle
    End If
End Sub


Private Sub Chart_MouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
    With ActiveChart
      Dim temp As Double
      Dim length As Double
      Dim X_axis_max As Double, X_axis_min As Double, Y_axis_max As Double, Y_axis_min As Double
      Dim t As Double, l As Double, w As Double, h As Double
      
      If isZoomIn = True Then
       ' force 100% zoom and maximize active Excel window
            If ActiveWindow.Zoom <> 100 Then ActiveWindow.Zoom = 100
            Application.WindowState = xlMaximized
            
            With .PlotArea
                w = .Width
                h = .Height
                t = .Top
                l = .Left
            End With
            
            'If X > l And X < (l + w) And Y > t And Y < (t + h) Then
            Dim xmax As Double, xmin As Double, ymax As Double, ymin As Double, mjunit As Double
                With .Axes(xlCategory, xlPrimary)
                    length = .MaximumScale - .MinimumScale
                    temp = (x - l) / w
                    length = length * RATE
                    xmax = .MaximumScale - length * (1 - temp)
                    xmin = .MinimumScale + length * temp
                    mjunit = (xmax - xmin) / GRIDNUM
                    .MajorUnit = mjunit
                    .MaximumScale = xmax
                    .MinimumScale = xmin
                End With
                With .Axes(xlValue, xlPrimary)
                    length = .MaximumScale - .MinimumScale
                    temp = (y - t) / h
                    length = length * RATE
                    ymax = .MaximumScale - length * (1 - temp)
                    ymin = .MinimumScale + length * temp
                    .MajorUnit = mjunit
                    .MaximumScale = ymax
                    .MinimumScale = ymin
                End With
                
            'adjustMAXMIN xmax, xmin, ymax, ymin
            'End If
            
      End If
    End With
  
    'MsgBox "Button = " & Button & Chr$(13) & _
       "Shift = " & Shift & Chr$(13) & _
       "X = " & X & " Y = " & Y & _
       "zoom= " & isZoomIn & Chr$(13) & _
       "L = " & l & Chr$(13) & _
       "T = " & t & Chr$(13) & _
       "w = " & w & Chr$(13) & _
       "h = " & h & Chr$(13) & _
       "iL = " & il & Chr$(13) & _
       "iT = " & it & Chr$(13) & _
       "iW = " & iw & Chr$(13) & _
       "iH = " & ih & Chr$(13)
       


End Sub

'Create btns & Destroy when visit other sheet
Sub CreateIt()
    Dim shp As Shape
    Rem: Comment OUT To disable selection chart
    'Set shp = Me.Shapes.AddShape(msoShapeRectangle, 0, 0, ActiveChart.ChartArea.Width, ActiveChart.ChartArea.Height)
    'With shp
        '.Fill.Visible = msoTrue
        '.Fill.Solid
        '.Fill.Transparency = 1
        '.Fill.ForeColor.RGB = RGB(255, 255, 255)
        '.Line.ForeColor.RGB = RGB(131, 195, 15)
        '.Line.Weight = 2
        '.Line.DashStyle = msoLineSolid
        '.Line.Style = msoLineSingle
        '.Line.Visible = msoTrue
        '.Name = "bkground"
    'End With
    
    Set shp = Me.Shapes.AddShape(msoShapeRightArrow, ActiveChart.ChartArea.Width - 50, ActiveChart.ChartArea.Top + 30, 20, 20)
    With shp
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.Transparency = 0.5
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(131, 195, 15)
        .Line.Weight = 2
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.Visible = msoTrue
        .Name = "Right"
    End With
    shp.OnAction = "Chart1.moveRight"
    
    Set shp = Me.Shapes.AddShape(msoShapeLeftArrow, ActiveChart.ChartArea.Width - 90, ActiveChart.ChartArea.Top + 30, 20, 20)
    With shp
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.Transparency = 0.5
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(131, 195, 15)
        .Line.Weight = 2
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.Visible = msoTrue
        .Name = "Left"
    End With
    shp.OnAction = "Chart1.moveLeft"
    Set shp = Me.Shapes.AddShape(msoShapeDownArrow, ActiveChart.ChartArea.Width - 70, ActiveChart.ChartArea.Top + 50, 20, 20)
    With shp
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.Transparency = 0.5
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(131, 195, 15)
        .Line.Weight = 2
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.Visible = msoTrue
        .Name = "Down"
    End With
     shp.OnAction = "Chart1.moveDown"
    Set shp = Me.Shapes.AddShape(msoShapeUpArrow, ActiveChart.ChartArea.Width - 70, ActiveChart.ChartArea.Top + 10, 20, 20)
    With shp
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.Transparency = 0.5
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(131, 195, 15)
        .Line.Weight = 2
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.Visible = msoTrue
        .Name = "Up"
    End With
     shp.OnAction = "Chart1.moveUp"

    Set shp = Me.Shapes.AddShape(msoShapeMathPlus, 10, 10, 40, 40)
    With shp
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.Transparency = 0.4
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.ForeColor.RGB = RGB(131, 195, 15)
        .Line.Weight = 2
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.Visible = msoTrue
        .Name = "zoomBTN"
    End With
    shp.OnAction = "Chart1.zoomInToggle"
    
    Set shp = Me.Shapes.AddShape(msoShapeOval, 20, 60, 20, 20)
    With shp
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.Transparency = 0.4
        .Fill.ForeColor.RGB = RGB(150, 180, 40)
        .Line.ForeColor.RGB = RGB(131, 195, 15)
        .Line.Weight = 2
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.Visible = msoTrue
        .Name = "resetBTN"
    End With
    shp.OnAction = "Chart1.reset"
    
    Set shp = Me.Shapes.AddShape(msoShapeRectangle, 20, 100, 20, 20)
    With shp
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.Transparency = 0.4
        .Fill.ForeColor.RGB = RGB(150, 180, 40)
        .Line.ForeColor.RGB = RGB(131, 195, 15)
        .Line.Weight = 2
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.Visible = msoTrue
        isToggleTheore = False
        .Name = "toggleTheoreBTN"
        toggleTheore
        .TextFrame.Characters.Text = "T"
    End With
    shp.OnAction = "Chart1.toggleTheore"
    
    Set shp = Me.Shapes.AddShape(msoShapeRectangle, 20, 140, 20, 20)
    With shp
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.Transparency = 0.4
        .Fill.ForeColor.RGB = RGB(150, 180, 40)
        .Line.ForeColor.RGB = RGB(131, 195, 15)
        .Line.Weight = 2
        .Line.DashStyle = msoLineSolid
        .Line.Style = msoLineSingle
        .Line.Visible = msoTrue
        isToggleViz = False
        .Name = "toggleVizBTN"
        toggleViz
        .TextFrame.Characters.Text = "V"
    End With
    shp.OnAction = "Chart1.toggleViz"
End Sub
 
Sub KillIt()
    Dim shp As Shape
    For Each shp In Me.Shapes
        shp.Delete
    Next
End Sub

'BTN functions
Sub zoomInToggle()
    With ActiveChart
      With .Shapes("zoomBTN")
        If isZoomIn = True Then
            isZoomIn = False
            .Fill.ForeColor.RGB = RGB(255, 255, 255)
        Else
            isZoomIn = True
            .Fill.ForeColor.RGB = RGB(131, 195, 15)
        End If
      End With
    End With
End Sub

Sub moveRight()
    If isZoomIn = True Then
        zoomInToggle
    End If
    With ActiveChart
        With .Axes(xlCategory, xlPrimary)
            .MaximumScale = .MaximumScale - .MajorUnit / 2
            .MinimumScale = .MinimumScale - .MajorUnit / 2
        End With
    End With
End Sub
Sub moveLeft()
    If isZoomIn = True Then
        zoomInToggle
    End If
    With ActiveChart
        With .Axes(xlCategory, xlPrimary)
            .MaximumScale = .MaximumScale + .MajorUnit / 2
            .MinimumScale = .MinimumScale + .MajorUnit / 2
        End With
    End With
End Sub

Sub moveUp()
    If isZoomIn = True Then
        zoomInToggle
    End If
    With ActiveChart
        With .Axes(xlValue, xlPrimary)
            .MaximumScale = .MaximumScale - .MajorUnit / 2
            .MinimumScale = .MinimumScale - .MajorUnit / 2
        End With
    End With
End Sub
Sub moveDown()
    If isZoomIn = True Then
        zoomInToggle
    End If
    With ActiveChart
        With .Axes(xlValue, xlPrimary)
            .MaximumScale = .MaximumScale + .MajorUnit / 2
            .MinimumScale = .MinimumScale + .MajorUnit / 2
        End With
    End With
End Sub

Sub reset()
Dim xmax As Double, xmin As Double, ymax As Double, ymin As Double

    With ActiveChart
        isZoomIn = False
        .Shapes("zoomBTN").Fill.ForeColor.RGB = RGB(255, 255, 255)
        
        With .Axes(xlCategory)
            '.MinimumScaleIsAuto = False
            '.MaximumScaleIsAuto = False
            '.MajorUnitIsAuto = False
            .MinimumScaleIsAuto = True
            .MaximumScaleIsAuto = True
            .MajorUnitIsAuto = True
            xmax = .MaximumScale
            xmin = .MinimumScale
        End With
        With .Axes(xlValue)
            .MinimumScaleIsAuto = True
            .MaximumScaleIsAuto = True
            .MajorUnitIsAuto = True
            ymax = .MaximumScale
            ymin = .MinimumScale
        End With
        
        adjustMAXMIN xmax, xmin, ymax, ymin
    End With
End Sub

Sub toggleTheore()
    With ActiveChart
        With .Shapes("toggleTheoreBTN")
          If isToggleTheore = True Then
              isToggleTheore = False
              .Fill.ForeColor.RGB = RGB(255, 255, 255)
          Else
              isToggleTheore = True
              .Fill.ForeColor.RGB = RGB(131, 195, 15)
          End If
        End With
      
        If isToggleTheore = False Then
            .SeriesCollection("IntersectionsTheore").MarkerStyle = xlMarkerStyleNone
        Else
            .SeriesCollection("IntersectionsTheore").MarkerStyle = xlMarkerStyleCircle
        End If
    End With
    setSeriesVisibility beginIndTheore, EndIndTheore, isToggleTheore
End Sub

Sub toggleViz()
    With ActiveChart
        With .Shapes("toggleVizBTN")
          If isToggleViz = True Then
              isToggleViz = False
              .Fill.ForeColor.RGB = RGB(255, 255, 255)
          Else
              isToggleViz = True
              .Fill.ForeColor.RGB = RGB(131, 195, 15)
          End If
        End With
      
        If isToggleViz = False Then
            .SeriesCollection("IntersectionsViztool").MarkerStyle = xlMarkerStyleNone
        Else
            .SeriesCollection("IntersectionsViztool").MarkerStyle = xlMarkerStyleCircle
        End If
    End With
    setSeriesVisibility beginIndViz, EndIndViz, isToggleViz
End Sub

Private Sub setSeriesVisibility(ByVal startInd As Integer, ByVal endInd As Integer, ByVal State As Boolean)
    Dim ser As Series, i As Integer
    
    With ActiveChart
        For i = startInd To endInd
            Set ser = .SeriesCollection(i)
            With ser
                .Format.Line.Visible = State
            End With
        Next i
    End With
End Sub

Private Sub adjustMAXMIN(ByVal xmax As Double, ByVal xmin As Double, ByVal ymax As Double, ByVal ymin As Double)
    Dim rw As Double, rh As Double, dw As Double, dh As Double

    With ActiveChart
        rw = .PlotArea.InsideWidth
        rh = .PlotArea.InsideHeight
        dw = xmax - xmin
        rw = rw / dw
            dh = ymax - ymin
            rh = rh / dh
        
        If rw > rh Then
            .Axes(xlCategory).MinimumScale = xmin
            .Axes(xlCategory).MaximumScale = xmin + (xmax - xmin) * rw / rh
        Else
            .Axes(xlValue).MinimumScale = ymin
            .Axes(xlValue).MaximumScale = ymin + (ymax - ymin) * rh / rw
        End If
    End With
End Sub



Private Sub Chart_MouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
        Dim t As Double, l As Double, w As Double, h As Double

            With ActiveChart.PlotArea
                w = .InsideWidth + .Width
                h = .InsideHeight + .Height
                t = .InsideTop + .Top
                l = .InsideLeft + .Left
            End With
            
    
    'Application.StatusBar = "X: " & x & ", Y: " & y & ", L: " & l & ", T: " & t & ", W: " & w & ", H: " & h
    
End Sub

Sub IntersectionCalc(ByVal BeginInd As Integer, ByVal endInd As Integer, ByVal seriesName As String)
    Dim ser As Series
    Dim linex As Variant, liney As Variant
    Dim cnt As Integer, tick As Integer, cnt2 As Integer
    Dim Px As Double, Py As Double
    Dim Name As String
    Dim index As Integer
    Dim isExist As Boolean
            
    linex = Array(0, 0, 0, 0, 0)
    liney = Array(0, 0, 0, 0, 0)
    tick = 0
    'Set chart and series initialization
    Name = seriesName
    isExist = False
    For index = 1 To ActiveChart.SeriesCollection.Count
        If ActiveChart.SeriesCollection(index).Name = Name Then
            isExist = True
            Exit For
        End If
    Next
                            
    If isExist Then
        Set ser = ActiveChart.SeriesCollection(index)
    Else
        Set ser = ActiveChart.SeriesCollection.NewSeries
    End If
    
    With ser
        .XValues = 0
        .Values = 0
        .Name = Name
        .Format.Line.Visible = msoFalse
        .MarkerSize = interMarkSize
        .MarkerBackgroundColor = RGB(60, 40, 20) 'marker background
    End With
    
    'Find intersections
    For cnt = BeginInd To (endInd - 1)
        For cnt2 = (cnt + 1) To endInd
            'If ActiveChart.SeriesCollection.Count < 255 Then
                'Obtain line vectors
                Dim isNotAvailable As Boolean
                isNotAvailable = False
                Set ser = ActiveChart.SeriesCollection(cnt)
                linex(1) = ser.XValues(1)
                linex(2) = ser.XValues(2)
                liney(1) = ser.Values(1)
                liney(2) = ser.Values(2)
                Set ser = ActiveChart.SeriesCollection(cnt2)
                linex(3) = ser.XValues(1)
                linex(4) = ser.XValues(2)
                liney(3) = ser.Values(1)
                liney(4) = ser.Values(2)
                
                If (linex(1) = 4096 And linex(2) = 4096) Or (linex(3) = 4096 And linex(4) = 4096) Then
                        isNotAvailable = True
                End If
                
                If isNotAvailable = False Then
                        'Calculation
                        Px = ((linex(1) * liney(2) - liney(1) * linex(2)) * (linex(3) - linex(4)) - (linex(1) - linex(2)) * (linex(3) * liney(4) - liney(3) * linex(4))) / ((linex(1) - linex(2)) * (liney(3) - liney(4)) - (linex(3) - linex(4)) * (liney(1) - liney(2)))
                        Py = ((linex(1) * liney(2) - liney(1) * linex(2)) * (liney(3) - liney(4)) - (liney(1) - liney(2)) * (linex(3) * liney(4) - liney(3) * linex(4))) / ((linex(1) - linex(2)) * (liney(3) - liney(4)) - (linex(3) - linex(4)) * (liney(1) - liney(2)))
                        
                        'Plot on Graph
                        Set ser = ActiveChart.SeriesCollection(index)
                        'Add if intersecting within the frame
                        If Px >= 0 And Px <= 4096 And Py <= 2322 And Py >= 0 Then
                            Dim arr() As Variant
                            With ser
                                arr = .XValues
                                ReDim Preserve arr(1 To UBound(arr) + 1) As Variant
                                arr(UBound(arr)) = Px
                                .XValues = arr
                                                
                                arr = .Values
                                ReDim Preserve arr(1 To UBound(arr) + 1) As Variant
                                arr(UBound(arr)) = Py
                                .Values = arr
                            End With
                            'tick = tick + 1
                        End If
                End If
        Next
    Next
    
    'Set ser = ActiveChart.SeriesCollection(3)
    'Xvs = ser.XValues
    'Yvs = ser.Values
    
        
'Application.StatusBar = "X: " & Xvs(1) & ", Y: " & Yvs(1) & "X2: " & Xvs(2) & ", Y2: " & Yvs(2) & ", Size: " & Size
End Sub









