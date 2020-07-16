Attribute VB_Name = "LCARgraphics"
Option Explicit '?curdir(

Public dest As Form, Rotate As Boolean, Buffer As PictureBox

Public Const LCAR_Black As Long = vbBlack 'RGB(0, 0, 0)
Public Const LCAR_DarkOrange As Long = 27607   'RGB(215, 107, 0)
Public Const LCAR_Orange As Long = 39421 ' rgb(253,153,0)  33023 'RGB(255, 128, 0)
Public Const LCAR_LightOrange As Long = 33023 '65535 'RGB(255, 255, 0)
'Public Const LCAR_DarkPurple As Long = 8388736 'rgb(128,0,128)
Public Const LCAR_Purple As Long = 16711935 'rgb(255,0,255)
Public Const LCAR_LightPurple As Long = 13408716 ' rgb(204,153,204)
Public Const LCAR_LightBlue As Long = 13408665 'rgb(153,153,204)
Public Const LCAR_Red As Long = 6710988 'rgb(204,102,102)
Public Const LCAR_Yellow As Long = 10079487 'rgb(255,204,153)
Public Const LCAR_DarkBlue As Long = 16751001 'rgb(153,153,255)
Public Const LCAR_DarkYellow As Long = 6724095 'rgb(255,153,102)
Public Const LCAR_DarkPurple As Long = 10053324 'rgb(204,102,153)
Public Const LCAR_White As Long = vbWhite

Public Declare Function GetPixel Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixelV Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Byte

Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Const LOGPIXELSX = 88 ' Logical pixels/inch in X
    Private Const LOGPIXELSY = 90 ' Logical pixels/inch in Y
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Const LF_FACESIZE = 32

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFacename As String * 33
End Type
Private Const FW_BOLD As Long = 700, FW_NORMAL As Long = 400

Private Declare Function CreateFontIndirect Lib "GDI32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Public isRotated As Boolean, hPrevFont As Long, hFont As Long, issetup As Boolean, F As LOGFONT, FontName As String
Dim Buttonmode As Boolean

Public Type LCARColor
    Name As String
    Color As OLE_COLOR
    Blink As OLE_COLOR
    
    'Gradient IDs
    Gradient As Long
    BlinkColor As Long
    Nemesis As Long
End Type
Public ColorList() As LCARColor, ColorCount As Long

Public Type AApoint
    DirX As Long
    DirY As Long
    
    Pixels As Long
End Type

Public Type AAcache
    L As Double
    R As Double
    Radius As Long
    Q As Long
    
    Points As Long
    Grid() As AApoint
End Type
Public AAlist() As AAcache, AAcount As Long

Public LeftAA1 As Long, RightAA1 As Long, LeftAA2 As Long, RightAA2 As Long
Public LeftAB1 As Long, RightAB1 As Long, LeftAB2 As Long, RightAB2 As Long 'Rotated

Public Const HalfPi As Double = PI * 0.5

Public Sub AALCAR(ByVal X As Long, ByVal Y As Long, ColorID As Long, Blink As Boolean, Optional RightSide As Boolean)
    Const ItemHeight As Long = 21
    If Not AntiAliasing Or ColorID = 0 Then Exit Sub
    
    If Rotate Then
        RotateXY X, Y
        If RightSide Then
            DrawAA X + 9, Y + 10, RightAB1, ColorID, Blink
            DrawAA X - 1, Y, RightAB2, ColorID, Blink
        Else
            DrawAA X + ItemHeight, Y - 10, LeftAB1, ColorID, Blink
            DrawAA X + 11, Y - 22, LeftAB2, ColorID, Blink
        End If
    Else
    
        If RightSide Then
            DrawAA X - 11, Y + 7, RightAA1, ColorID, Blink
            DrawAAline X + 10, Y + 7, 0, -1, 3, ColorID, Blink, False, -1, 0
        
            DrawAA X + 3, Y - 1, RightAA2, ColorID, Blink
            DrawAAline X + 3, Y + ItemHeight - 1, 1, 0, 3, ColorID, Blink, False, 0, -1
        Else
            DrawAA X + 7, Y + ItemHeight, LeftAA1, ColorID, Blink
            DrawAAline X + 7, Y, -1, 0, 3, ColorID, Blink, False, 0, 1
            
            DrawAA X + 21, Y + ItemHeight - 8, LeftAA2, ColorID, Blink
            DrawAAline X, Y + 13, 0, 1, 3, ColorID, Blink, False, 1, 0
        End If
    
    End If
End Sub

Public Sub SetupLCARAA()
    Const ItemHeight As Long = 21
    '2 1 2 TURN 1 2  2
    '2 2 TURN 1 2 1 2
    
    LeftAA1 = AddAA(HalfPi, PI, ItemHeight)
        AddAApoint LeftAA1, -1, 0, 2
        AddAApoint LeftAA1, -1, 0, 1
        AddAApoint LeftAA1, -1, 0, 2
        AddAApoint LeftAA1, 0, 1, 1
        AddAApoint LeftAA1, 0, 1, 2
        AddAApoint LeftAA1, 0, 1, 2
        'AddAApoint LeftAA1, 0, 1, 2
        
    LeftAA2 = AddAA(PI, PI + HalfPi, ItemHeight)
        AddAApoint LeftAA2, 0, 1, 2
        AddAApoint LeftAA2, 0, 1, 2
        AddAApoint LeftAA2, 1, 0, 1
        AddAApoint LeftAA2, 1, 0, 2
        AddAApoint LeftAA2, 1, 0, 1
        AddAApoint LeftAA2, 1, 0, 2
    
    LeftAB1 = AddAA(PI, PI + HalfPi, ItemHeight + 1)
        AddAApoint LeftAB1, 0, 1, 2
        AddAApoint LeftAB1, 0, 1, 2
        AddAApoint LeftAB1, 0, 1, 1
        AddAApoint LeftAB1, 0, 1, 2
        AddAApoint LeftAB1, 0, 1, 1
        AddAApoint LeftAB1, 1, 0, 2
        AddAApoint LeftAB1, 1, 0, 2
        
    LeftAB2 = AddAA(PI + HalfPi, 0, ItemHeight + 1)
        AddAApoint LeftAB2, 1, 0, 2
        AddAApoint LeftAB2, 1, 0, 2
        AddAApoint LeftAB2, 1, 0, 2
        AddAApoint LeftAB2, 1, 0, 1
        AddAApoint LeftAB2, 0, -1, 2
        AddAApoint LeftAB2, 0, -1, 1
        AddAApoint LeftAB2, 0, -1, 2
    
    RightAA1 = AddAA(0, HalfPi, ItemHeight)
        AddAApoint RightAA1, 0, -1, 2
        AddAApoint RightAA1, 0, -1, 2
        AddAApoint RightAA1, 0, -1, 1
        AddAApoint RightAA1, -1, 0, 1
        AddAApoint RightAA1, -1, 0, 2
        AddAApoint RightAA1, -1, 0, 2
    
    RightAA2 = AddAA(PI + HalfPi, 0, ItemHeight)
        AddAApoint RightAA2, 1, 0, 2
        AddAApoint RightAA2, 1, 0, 2
        AddAApoint RightAA2, 1, 0, 1
        AddAApoint RightAA2, 0, -1, 1
        AddAApoint RightAA2, 0, -1, 2
        AddAApoint RightAA2, 0, -1, 2
    
    RightAB1 = AddAA(HalfPi, PI, ItemHeight + 1)
        AddAApoint RightAB1, -1, 0, 2
        AddAApoint RightAB1, -1, 0, 2
        AddAApoint RightAB1, -1, 0, 2
        AddAApoint RightAB1, -1, 0, 1
        AddAApoint RightAB1, -1, 0, 1
        AddAApoint RightAB1, -1, 0, 1
        AddAApoint RightAB1, 0, 1, 5
        
   RightAB2 = AddAA(0, HalfPi, ItemHeight + 1)
        AddAApoint RightAB2, 0, -1, 1
        AddAApoint RightAB2, 0, -1, 6
        AddAApoint RightAB2, 0, -1, 1
        AddAApoint RightAB2, 0, -1, 1
        AddAApoint RightAB2, 0, -1, 1
        AddAApoint RightAB2, -1, 0, 2
        AddAApoint RightAB2, -1, 0, 2
End Sub

Public Function AddAA(ByVal L As Double, ByVal R As Double, Radius As Long) As Long
    Dim Quad1 As Long, Quad2 As Long, DeltaQuad As Long, Direction As Boolean, temp As Long
    
    AddAA = AAcount
    AAcount = AAcount + 1
    ReDim Preserve AAlist(AAcount)

    Quad1 = L / HalfPi
    Quad2 = R / HalfPi
    Direction = Quad2 > Quad1
    If Not Direction And Not Quad2 = 0 Then
        temp = Quad1
        Quad1 = Quad2
        Quad2 = temp
    End If
    DeltaQuad = Quad2 - Quad1

    With AAlist(AAcount - 1)
        .L = L
        .R = R
        .Radius = Radius
        .Q = Quad1
    End With
End Function

Public Function SampleAA(X As Long, Y As Long, ByVal L As Double, ByVal R As Double, Radius As Long, Optional OutsideEdge As Boolean) As Long
    Dim temp As Long, temp2 As Long, X2 As Long, Y2 As Long, Direction As Boolean, DirX As Long, DirY As Long, ScanMethod As Boolean
    Dim Quad1 As Long, Quad2 As Long, DeltaQuad As Long, DirX2 As Long, DirY2 As Long, Point As Long, temp3 As Long, Points As Long
    SampleAA = -1
    
    If AAcount = 0 Then SetupLCARAA
    
    For temp = 0 To AAcount - 1
        With AAlist(temp)
            If .L = L Then
                If .R = R Then
                    If .Radius = Radius Then
                        SampleAA = temp
                        Exit Function
                    End If
                End If
            End If
        End With
    Next
    
    Quad1 = L / HalfPi
    Quad2 = R / HalfPi
    Direction = Quad2 > Quad1
    If Not Direction And Not (Quad1 = 3 And Quad2 = 0) Then
        temp = Quad1
        Quad1 = Quad2
        Quad2 = temp
    End If
    DeltaQuad = Quad2 - Quad1
    
    'If DeltaQuad <> 1 Then Exit Function
    
    If Not OutsideEdge Then Exit Function 'FIX THIS
    
    SampleAA = AddAA(L, R, Radius)
    'SampleAA = AAcount
    'AAcount = AAcount + 1
    'ReDim Preserve AAlist(AAcount)

    'With AAlist(AAcount - 1)
    '    .L = L
    '    .R = R
    '    .Radius = Radius
    '    .Q = Quad1
    'End With
    
    If OutsideEdge Then
    
        'Debug.Print Quad1 & " " & Quad2 & " " & Radius
        
        X2 = X
        Y2 = Y
        CalcQuad Quad1, X, Y, Radius, X2, Y2, DirX, DirY, DirX2, DirY2
        
        temp2 = ScanXY(X2, Y2, DirX, DirY)
        'temp3 = ScanXY(X2 + DirX2, Y2 + DirY2, DirX, DirY)
        'temp3 = temp3 - temp2
        AddAApoint AAcount - 1, DirX, DirY, temp2
        'AAlist(AAcount - 1).Points = 1
        'ReDim Preserve AAlist(AAcount - 1).Grid(1)
        'With AAlist(AAcount - 1).Grid(0)
        '    .DirX = DirX
        '    .DirY = DirY
        '    .Pixels = temp2
        'End With
            
        For temp = 0 To Radius
            Point = Point + 1
            X2 = X2 + DirX2
            Y2 = Y2 + DirY2
            If DirX <> 0 Then
                X2 = X2 + (DirX * temp2)
            Else
                Y2 = Y2 + (DirY * temp2)
            End If
            
            If ScanMethod Then Points = Points + 1 Else Points = Points + temp2
            If Points = Radius Then Exit For
            
            temp3 = ScanXY(X2, Y2, DirX, DirY, ScanMethod)
                        
            If temp3 = 0 Then
                'Debug.Print "TURNING POINT"
                ScanMethod = Not ScanMethod ' True
                
                X2 = X2 - DirX2
                Y2 = Y2 - DirY2
            
                Quad1 = Quad1 + 1
                CalcQuad Quad1, 0, 0, 0, 0, 0, DirX, DirY, DirX2, DirY2, ScanMethod
                       
                       
                
                temp3 = ScanXY(X2, Y2, DirX, DirY, True)
            End If
            
            temp2 = temp3
            
            AddAApoint AAcount - 1, DirX, DirY, temp2
            
            'Debug.Print " Point: " & temp & " Pixels: " & temp2 & " DirX: " & DirX & " DirY: " & DirY '& " DirX2: " & DirX2 & " DirY2: " & DirY2
            'AAlist(AAcount - 1).Points = AAlist(AAcount - 1).Points + 1
            'ReDim Preserve AAlist(AAcount - 1).Grid(AAlist(AAcount - 1).Points)
            'With AAlist(AAcount - 1).Grid(AAlist(AAcount - 1).Points - 1)
            '    .DirX = DirX
            '    .DirY = DirY
            '    .Pixels = temp2
            'End With
            
        Next
        
    End If
        'MsgBox temp2
End Function


Private Sub AddAApoint(AAIndex As Long, DirX As Long, DirY As Long, Pixels As Long)
    AAlist(AAIndex).Points = AAlist(AAIndex).Points + 1
    ReDim Preserve AAlist(AAIndex).Grid(AAlist(AAIndex).Points)
    With AAlist(AAIndex).Grid(AAlist(AAIndex).Points - 1)
        .DirX = DirX
        .DirY = DirY
        .Pixels = Pixels
    End With
End Sub



Public Function ScanXY(ByVal X As Long, ByVal Y As Long, Optional DirX As Long, Optional DirY As Long, Optional ScanMethod As Boolean) As Long
    Dim temp As Long, Color As Long, Pixels As Long
    
    Color = GetPixel(dest.hdc, X, Y)
    'SetPixelV dest.hDC, X, Y, vbGreen
    If ScanMethod Then
        Pixels = 1
        Do While Color = vbBlack 'Or Pixels = 0
            Pixels = Pixels + 1
            X = X + DirX
            Y = Y + DirY
            Color = GetPixel(dest.hdc, X, Y)
            'SetPixelV dest.hDC, X, Y, vbRed
        Loop
    Else
        Do While Color <> vbBlack 'Or Pixels = 0
            Pixels = Pixels + 1
            X = X + DirX
            Y = Y + DirY
            Color = GetPixel(dest.hdc, X, Y)
            'SetPixelV dest.hDC, X, Y, vbRed
        Loop
    End If
    
    'dest.Refresh
    ScanXY = Pixels
End Function






Public Sub CalcQuad(Quad As Long, X As Long, Y As Long, Radius As Long, X2 As Long, Y2 As Long, DirX As Long, DirY As Long, DirX2 As Long, DirY2 As Long, Optional Reverse As Boolean)
    DirX = 0
    DirY = 0
    DirX2 = 0
    DirY2 = 0
    X2 = X
    Y2 = Y
    
    Select Case Quad
            Case 0, 4
                Quad = 0
                X2 = X + Radius - 1
                If Reverse Then DirX2 = 1 Else DirX2 = -1
                DirY = -1
            Case 1
                Y2 = Y - Radius + 1
                DirX = -1
                If Reverse Then DirY2 = -1 Else DirY2 = 1
            Case 2
                X2 = X - Radius + 1
                If Reverse Then DirX2 = -1 Else DirX2 = 1
                DirY = 1
            Case 3
                Y2 = Y + Radius - 1
                DirX = 1
                If Reverse Then DirY2 = 1 Else DirY2 = -1
        End Select
End Sub





Public Function DrawAA(ByVal X As Long, ByVal Y As Long, Index As Long, ByVal ColorID As Long, Blink As Boolean) As Boolean
    Dim temp As Long, temp2 As Long, X2 As Long, Y2 As Long, DirX As Long, DirY As Long, DirX2 As Long, DirY2 As Long, Quad1 As Long
    Dim step As Long, StepInc As Long, Color As Long, Reverse As Boolean, Steps As Long
    If ColorID = 0 Then Exit Function
   ' Exit Function
    
    With AAlist(Index)
        CalcQuad .Q, X, Y, .Radius, X2, Y2, DirX, DirY, DirX2, DirY2
        Quad1 = .Q
        
        Y = Y2 + (DirY * .Grid(0).Pixels)
        X = X2 + (DirX * .Grid(0).Pixels)
        
        For temp = 0 To .Points - 2

            If .Grid(temp + 1).DirX <> DirX Or .Grid(temp + 1).DirY <> DirY Then
                Quad1 = Quad1 + 1
                Reverse = True
                CalcQuad Quad1, 0, 0, 0, 0, 0, DirX, DirY, DirX2, DirY2, Reverse
            End If

            Steps = .Grid(temp + 1).Pixels
            If Steps = 0 Then Exit Function
            
            StepInc = (GradientSteps / Steps)
            If Reverse Then
                step = (GradientSteps - 1) - StepInc  '- (StepInc * .Grid(temp + 1).Pixels)
                StepInc = -StepInc / 2
            Else
                step = StepInc - 1
                StepInc = StepInc / 2
            End If

            X2 = X
            Y2 = Y
            
            
            If Steps = 1 Then
                step = GradientSteps / 2
                If Blink Then
                    Color = ColorList(ColorID).BlinkColor
                Else
                    Color = ColorList(ColorID).Gradient
                End If
                Color = GradientList(Color).Cache(step)
                SetPixelV dest.hdc, X2, Y2, Color
                
                Y2 = Y2 + DirY
                X2 = X2 + DirX
            Else
                For temp2 = 1 To Steps
                    If Blink Then
                        Color = ColorList(ColorID).BlinkColor
                    Else
                        Color = ColorList(ColorID).Gradient
                    End If
                    If step < 0 Then step = 0
                    If step > GradientSteps Then step = GradientSteps
                    Color = GradientList(Color).Cache(step)
                
                    'Color = vbRed
                
                
                    SetPixelV dest.hdc, X2, Y2, Color
                    step = step + StepInc
                
                    Y2 = Y2 + DirY
                    X2 = X2 + DirX
                Next
            End If
            
            Y = Y + DirY2 + (.Grid(temp + 1).Pixels * DirY)
            X = X + DirX2 + (.Grid(temp + 1).Pixels * DirX)
        Next
    End With
    DrawAA = True
End Function

Public Sub DrawAAline(ByVal X As Long, ByVal Y As Long, ByVal DirX As Long, ByVal DirY As Long, Pixels As Long, ByVal ColorID As Long, Blink As Boolean, Reverse As Boolean, ByVal DirXr As Long, ByVal DirYr As Long)
    Dim temp As Long, Inc As Long, step As Long, Color As Long
    
    If Not AntiAliasing Then Exit Sub
    
    If Rotate Then
        RotateXY X, Y
        Y = Y - 1
        'X = X - 1
        DirX = DirXr
        DirY = DirYr
    End If
    
    If Pixels = 1 Then
        step = GradientSteps / 2
    Else
        Inc = GradientSteps / Pixels
        step = Inc
    End If
    
    If Blink Then
        ColorID = ColorList(ColorID).BlinkColor
    Else
        ColorID = ColorList(ColorID).Gradient
    End If
        
    For temp = 1 To Pixels
        If step > GradientSteps Then step = GradientSteps
        Color = GradientList(ColorID).Cache(step)
        SetPixelV dest.hdc, X, Y, Color
        X = X + DirX
        Y = Y + DirY
        step = step + Inc
    Next
    
End Sub

'Draing LCAR elements
Public Sub DrawLCARButton(X As Long, Y As Long, Width As Long, Height As Long, Text As String, Optional EdgeColor As OLE_COLOR = LCAR_DarkOrange, Optional FillColor As OLE_COLOR = LCAR_Orange, Optional LeftSideWidth As Long, Optional RightSideWidth As Long, Optional WhiteSpace As Long = 4, Optional TextAlign As Long = 4, Optional TextSize As Single, Optional TextColor As OLE_COLOR = vbBlack, Optional ColorID As Long)
    Dim temp As Long, temp2 As Long, Unit As Long, Start As Long, tX As Long, tY As Long, Color As OLE_COLOR, Blink As Boolean
    If LeftSideWidth > 0 Or RightSideWidth > 0 Then
        If Height Mod 2 = 0 Then Height = Height + 1 'must be an odd number
    End If
    Unit = Height / 2
    Buttonmode = True
    
    If LeftSideWidth > 0 Or RightSideWidth > 0 Then
        ColorID = LCAR_ColorIDfromColor(EdgeColor)
        Blink = EdgeColor = ColorList(ColorID).Blink
    End If
    
    If LeftSideWidth > 0 Then
        If LeftSideWidth < Height Then LeftSideWidth = Height
        DrawSquare Unit - 2 + X, Y, LeftSideWidth - Unit, Height, EdgeColor, FillColor
        DrawSemiCircle Unit + X, Unit + Y, Unit - 1, 90, 180, -1, FillColor ', , , , , , ColorID
        DrawSemiCircle Unit + X - 1, Unit + Y, Unit - 1, 90, 180, EdgeColor, -1 ', , , , , , ColorID
        DrawLine Unit + X - 1, Y, 5, 1, EdgeColor
        
        AALCAR X, Y, ColorID, Blink
        
        temp = LeftSideWidth + WhiteSpace
    End If
    
    If RightSideWidth > 0 Then
        If RightSideWidth < Height Then RightSideWidth = Height
        Start = (X + Width) - RightSideWidth
        DrawSquare Start - 1, Y, RightSideWidth - Unit, Height, EdgeColor, FillColor
        DrawSemiCircle Unit + Start, Unit + Y, Unit - 1, 270, -180, -1, FillColor ', , , , , , ColorID
        DrawSemiCircle Unit + Start, Unit + Y, Unit, 270, -180, EdgeColor, -1 ', , , , , , ColorID
        DrawLine Unit + Start - 2, Y + 1, 1, Height - 1, FillColor
        DrawLine Unit + Start - 2, Y, 5, 1, EdgeColor
        
        AALCAR Unit + Start, Y, ColorID, Blink, True
        
        If WhiteSpace = 0 Then
            temp2 = RightSideWidth
        Else
            temp2 = WhiteSpace + RightSideWidth + 2
        End If
    End If
    
    DrawSquare temp + X, Y, Width - temp - temp2, Height, EdgeColor, FillColor
    If Len(Text) > 0 Then
        If TextSize > 0 Then
            SwitchToUnRotated
            dest.Font.Size = TextSize
        End If
        Select Case TextAlign
            Case 1, 2, 3: tY = Y  'top row
            Case 4, 5, 6: tY = Y + (Unit - dest.TextHeight(Text) / 2)  'middle row
            Case 7, 8, 9: tY = Y + Height - dest.TextHeight(Text) 'bottom row
        End Select
        Select Case TextAlign
            Case 1, 4, 7: tX = temp + X + 3 ' left column
            Case 2, 5, 8: tX = X + ((Width - temp - temp2) / 2) - (dest.TextWidth(Text) / 2) + temp 'middle column
            Case 3, 6, 9: tX = X + Width - temp2 - dest.TextWidth(Text) - 2 'right column
        End Select
        'If TextAlign = 5 And Text = UCase("This operation will cause damage to the file system") Then MsgBox "HI"
        
        Color = TextColor 'vbBlack
        If EdgeColor = vbBlack And Not RedAlert Then Color = LCAR_Orange
        
        DrawText tX, tY, Text, Color, TextSize
    End If
    
    Buttonmode = False
End Sub

Public Sub DrawLCARelbow(X As Long, Y As Long, Width As Long, Height As Long, BarWidth As Long, BarHeight As Long, Optional Radius As Long, Optional Align As Long, Optional EdgeColor As OLE_COLOR = LCAR_DarkOrange, Optional FillColor As OLE_COLOR = LCAR_Orange, Optional Text As String, Optional TextAlign As Long = 4, Optional ColorID As Long)
    Dim Aspect As Double, temp As Long, temp2 As Long, Blink As Boolean
    Const AspectMode As Boolean = True
    
    If AspectMode Then
        temp2 = BarWidth / 2
    Else
        Aspect = BarHeight / BarWidth
        If Rotate Then Aspect = BarWidth / BarHeight
    End If
    
    Blink = EdgeColor = ColorList(Color).Blink
    If Radius = 0 Then Radius = 10
    temp = Radius ' * Aspect
    EdgeColor = FillColor
    
    Select Case Align
                '_
        Case 0 '|  top left
            If AspectMode Then 'new aspect ratio (1.0)
                DrawAAline X + temp2 - 2, Y, -1, 0, 12, Color, Blink, False, 0, 1
                
                DrawAAline X + BarWidth, Y + BarHeight + Radius, 0, 1, 12, Color, Blink, False, 1, 0
                DrawAAline X + BarWidth + Radius, Y + BarHeight, 1, 0, 12, Color, Blink, False, 0, -1
                
                DrawAAline X + BarWidth + 1, Y + BarHeight + Radius - 4, 0, 1, 8, Color, Blink, False, 1, 0
                DrawAAline X + BarWidth + 2, Y + BarHeight + Radius - 5, 0, 1, 6, Color, Blink, False, 1, 0
                DrawAAline X + BarWidth + 3, Y + BarHeight + Radius - 6, 0, 1, 4, Color, Blink, False, 1, 0
                DrawAAline X + BarWidth + 4, Y + BarHeight + Radius - 7, 0, 1, 4, Color, Blink, False, 1, 0
 
                DrawAAline X + BarWidth + Radius - 4, Y + BarHeight + 1, 1, 0, 8, Color, Blink, False, 0, -1
                DrawAAline X + BarWidth + Radius - 5, Y + BarHeight + 2, 1, 0, 6, Color, Blink, False, 0, -1
                DrawAAline X + BarWidth + Radius - 6, Y + BarHeight + 3, 1, 0, 4, Color, Blink, False, 0, -1
                DrawAAline X + BarWidth + Radius - 7, Y + BarHeight + 4, 1, 0, 4, Color, Blink, False, 0, -1
                
                DrawSemiCircle X + temp2 - 1, Y + temp2, temp2, 90, 90, -1, FillColor, , , , , , ColorID  'outside corner
                DrawSquare X + temp2 - 1, Y, Width - temp2, BarHeight, EdgeColor, FillColor
                DrawSquare X + temp2 - 1, Y + BarHeight, temp2 + 1, temp2 - BarHeight + 1, EdgeColor, FillColor
                DrawSquare X, Y + temp2, BarWidth, Height - temp2, EdgeColor, FillColor
                
                DrawPixel X + BarWidth + 3, Y + BarHeight + 3, EdgeColor
            Else 'Old aspect ratio (1.5)
                DrawSquare X + BarWidth, Y, Width - BarWidth, BarHeight, EdgeColor, FillColor
                DrawSquare X, Y + BarHeight - 1, BarWidth, Height - BarHeight + 1, EdgeColor, FillColor
                DrawSemiCircle X + BarWidth, Y + BarHeight + 1, BarWidth - 1, 90, 90, -1, FillColor, , Aspect
            End If
            DrawSemiCircle X + BarWidth + Radius, Y + BarHeight + temp, Radius * 2, 90, 90, -1, FillColor, 1, 1, Radius + 2, , , ColorID, False 'inside corner
            Select Case TextAlign
                Case 1: DrawText X + 3, Y + Height - dest.Font.Size - 4, Text, vbBlack 'left column
                Case 2: DrawText X + BarWidth / 2 - dest.TextWidth(Text) / 2 + 2, Y + Height - dest.Font.Size - 4, Text, vbBlack  'middle column
                Case 3: DrawText X + BarWidth - dest.TextWidth(Text) - 3, Y + Height - dest.Font.Size - 4, Text, vbBlack  'right column
                Case 4: DrawText X + BarWidth, Y, Text, vbBlack    'bar
            End Select
               '_
        Case 1 ' | top right
            If AspectMode Then 'new aspect ratio (1.0)
                DrawAAline X + Width - temp2, Y, 1, 0, 12, Color, Blink, False, 0, -1
                DrawAAline X + Width - 1, Y + Height - 10, 0, -1, 12, Color, Blink, False, -1, 0
                
                DrawAAline X + Width - BarWidth - 10, Y + BarHeight, -1, 0, 12, Color, Blink, False, 0, 1
                DrawAAline X + Width - BarWidth - 7, Y + BarHeight + 1, -1, 0, 8, Color, Blink, False, 0, 1
                DrawAAline X + Width - BarWidth - 5, Y + BarHeight + 2, -1, 0, 6, Color, Blink, False, 0, 1
                DrawAAline X + Width - BarWidth - 4, Y + BarHeight + 3, -1, 0, 4, Color, Blink, False, 0, 1
                DrawAAline X + Width - BarWidth - 3, Y + BarHeight + 4, -1, 0, 4, Color, Blink, False, 0, 1
                
                DrawAAline X + Width - BarWidth - 1, Y + BarHeight + 10, 0, 1, 12, Color, Blink, False, 1, 0
                DrawAAline X + Width - BarWidth - 2, Y + BarHeight + 6, 0, 1, 8, Color, Blink, False, 1, 0
                DrawAAline X + Width - BarWidth - 3, Y + BarHeight + 5, 0, 1, 6, Color, Blink, False, 1, 0
                DrawAAline X + Width - BarWidth - 4, Y + BarHeight + 4, 0, 1, 4, Color, Blink, False, 1, 0
                DrawAAline X + Width - BarWidth - 5, Y + BarHeight + 3, 0, 1, 4, Color, Blink, False, 1, 0

                DrawSemiCircle X + Width - temp2 - 1, Y + temp2, temp2, 0, -90, -1, FillColor, , , , , , ColorID 'outside corner
                DrawSquare X, Y, Width - temp2, BarHeight, EdgeColor, FillColor
                DrawSquare X + Width - BarWidth, Y + BarHeight, temp2 + 1, temp2 - BarHeight + 1, EdgeColor, FillColor
                DrawSquare X + Width - BarWidth, Y + temp2, BarWidth, Height - temp2, EdgeColor, FillColor
                
                'If Rotate And AntiAliasing Then
                '    DrawPixel X + Width - BarWidth - 3, Y + BarHeight + 3, EdgeColor
                '    DrawPixel X + Width - BarWidth - 6, Y + BarHeight + 1, EdgeColor
                'End If
            Else 'Old aspect ratio (1.5)
                DrawSquare X, Y, Width - BarWidth + 1, BarHeight, EdgeColor, FillColor
                DrawSquare X + Width - BarWidth, Y + BarHeight, BarWidth, Height - BarHeight, EdgeColor, FillColor
                DrawSemiCircle X + Width - BarWidth - 2, Y + BarHeight, BarWidth, 0, -90, EdgeColor, FillColor, , Aspect
            End If
            DrawSemiCircle X + Width - BarWidth - Radius - 1, Y + BarHeight + temp, Radius * 2, 0, -90, -1, FillColor, 1, 1, Radius + 2, , , ColorID, False 'inside corner
            
            Select Case TextAlign
                'Case 1: DrawText X + Width - BarWidth + 3, Y, Text, vbBlack 'left column
                Case 2: DrawText X + Width - (BarWidth / 2) - (dest.TextWidth(Text) / 2) + 2, Y + Height - dest.Font.Size - 4, Text, vbBlack 'middle column
                'Case 3: DrawText X + Width - Dest.TextWidth(Text) - 3, Y, Text, vbBlack    'right column
                'Case 4: DrawText X + BarWidth, Y + Height - BarHeight, Text, vbBlack 'bar
            End Select
            
        Case 2 '|_ bottom left
            If AspectMode Then 'new aspect ratio (1.0)
                DrawAAline X, Y + 10, 0, 1, 12, Color, Blink, False, 1, 0
                DrawAAline X + BarWidth, Y + Height - BarHeight - Radius, 0, -1, 12, Color, Blink, False, -1, 0
                DrawAAline X + BarWidth + Radius, Y + Height - BarHeight - 1, 1, 0, 12, Color, Blink, False, 0, -1
                DrawAAline X + BarWidth + 1, Y + Height - BarHeight - Radius + 3, 0, -1, 8, Color, Blink, False, -1, 0
                DrawAAline X + BarWidth + Radius - 4, Y + Height - BarHeight - 2, 1, 0, 8, Color, Blink, False, 0, -1
                DrawAAline X + BarWidth + 2, Y + Height - BarHeight - Radius + 4, 0, -1, 6, Color, Blink, False, -1, 0
                DrawAAline X + BarWidth + Radius - 5, Y + Height - BarHeight - 3, 1, 0, 6, Color, Blink, False, 0, -1
                DrawAAline X + BarWidth + 3, Y + Height - BarHeight - Radius + 5, 0, -1, 4, Color, Blink, False, -1, 0
                DrawAAline X + BarWidth + Radius - 6, Y + Height - BarHeight - 4, 1, 0, 4, Color, Blink, False, 0, -1
                DrawAAline X + BarWidth + 4, Y + Height - BarHeight - Radius + 6, 0, -1, 4, Color, Blink, False, -1, 0
                DrawAAline X + BarWidth + Radius - 7, Y + Height - BarHeight - 5, 1, 0, 4, Color, Blink, False, 0, -1
                
                DrawSemiCircle X + temp2, Y + Height - temp2, temp2, 180, 90, -1, FillColor, , , , , , ColorID  'outside corner
                DrawSquare X, Y, BarWidth, Height - temp2, EdgeColor, FillColor
                DrawSquare X + temp2 - 1, Y + Height - BarHeight, Width - temp2, BarHeight, EdgeColor, FillColor
                DrawSquare X + temp2 - 1, Y + Height - temp2, temp2 + 1, temp2 - BarHeight + 1, EdgeColor, FillColor

               
            Else 'Old aspect ratio (1.5)
                DrawSquare X + BarWidth, Y + Height - BarHeight, Width - BarWidth, BarHeight, EdgeColor, FillColor
                DrawSquare X, Y, BarWidth, Height - BarHeight, EdgeColor, FillColor
                DrawSemiCircle X + BarWidth + 1, Y + Height - BarHeight - 1, BarWidth, 180, 90, -1, FillColor, , Aspect
            End If
            DrawSemiCircle X + BarWidth + Radius, Y + Height - BarHeight - temp - 1, Radius * 2, 180, 90, -1, FillColor, 1, 1, Radius + 2, , , ColorID, False 'inside corner
            Select Case TextAlign
                Case 1: DrawText X + 3, Y, Text, vbBlack  'left column
                Case 2: DrawText X + BarWidth / 2 - dest.TextWidth(Text) / 2 + 2, Y, Text, vbBlack   'middle column
                Case 3: DrawText X + BarWidth - dest.TextWidth(Text) - 3, Y, Text, vbBlack 'right column
                Case 4: DrawText X + BarWidth, Y + Height - dest.Font.Size - 4, Text, vbBlack    'bar
            End Select
            
        Case 3 '_| bottom right
            If AspectMode Then 'new aspect ratio (1.0)
                DrawSemiCircle X + Width - temp2, Y + Height - temp2, temp2, 270, 90, -1, FillColor, , , , , , ColorID 'outside corner
                DrawSquare X + Width - BarWidth, Y, BarWidth, Height - temp2, EdgeColor, FillColor
                DrawSquare X, Y + Height - BarHeight, Width - temp2, BarHeight, EdgeColor, FillColor
                DrawSquare X + Width - BarWidth, Y + Height - temp2, temp2 + 1, temp2 - BarHeight + 1, EdgeColor, FillColor
                
                DrawAAline X + Width - BarWidth - 10, Y + Height - BarHeight - 1, -1, 0, 12, Color, Blink, False, 0, 1
                DrawAAline X + Width - BarWidth - 6, Y + Height - BarHeight - 2, -1, 0, 8, Color, Blink, False, 0, 1
                DrawAAline X + Width - BarWidth - 4, Y + Height - BarHeight - 3, -1, 0, 6, Color, Blink, False, 0, 1
                DrawAAline X + Width - BarWidth - 4, Y + Height - BarHeight - 4, -1, 0, 4, Color, Blink, False, 0, 1
                DrawAAline X + Width - BarWidth - 3, Y + Height - BarHeight - 5, -1, 0, 4, Color, Blink, False, 0, 1
                
                DrawAAline X + Width - BarWidth - 1, Y + Height - BarHeight - 10, 0, -1, 12, Color, Blink, False, -1, 0
                DrawAAline X + Width - BarWidth - 2, Y + Height - BarHeight - 7, 0, -1, 8, Color, Blink, False, -1, 0
                DrawAAline X + Width - BarWidth - 3, Y + Height - BarHeight - 5, 0, -1, 6, Color, Blink, False, -1, 0
                DrawAAline X + Width - BarWidth - 4, Y + Height - BarHeight - 4, 0, -1, 4, Color, Blink, False, -1, 0
                DrawAAline X + Width - BarWidth - 5, Y + Height - BarHeight - 3, 0, -1, 4, Color, Blink, False, -1, 0
                
                If Rotate Then DrawPixel X + Width + 1, Y + Height - temp2 - 1, vbBlack
            Else 'Old aspect ratio (1.5)
                DrawSquare X, Y + Height - BarHeight, Width - BarWidth + 1, BarHeight, EdgeColor, FillColor
                DrawSquare X + Width - BarWidth, Y, BarWidth, Height - BarHeight, EdgeColor, FillColor
                DrawSemiCircle X + Width - BarWidth - 2, Y + Height - BarHeight - 1, BarWidth, 270, 90, EdgeColor, FillColor, , Aspect
            End If
            DrawSemiCircle X + Width - BarWidth - Radius, Y + Height - BarHeight - temp, Radius * 2, 270, 90, -1, FillColor, 1, 1, Radius + 1, , , ColorID, False 'inside corner
            'DrawSemiCircle X + BarWidth - Radius - 2, Y + Height - BarHeight - temp - 1, Radius * 2, 270, 90, EdgeColor, -1, 1, Aspect, Radius + 2, True
            
            Select Case TextAlign
                Case 1: DrawText X + Width - BarWidth + 3, Y, Text, vbBlack 'left column
                Case 2: DrawText X + Width - (BarWidth / 2) - dest.TextWidth(Text) / 2 + 2, Y, Text, vbBlack    'middle column
                Case 3: DrawText X + Width - dest.TextWidth(Text) - 3, Y, Text, vbBlack    'right column
                Case 4: DrawText X + BarWidth, Y + Height - BarHeight, Text, vbBlack 'bar
            End Select
    End Select
End Sub











'Drawing primitives + Rotation
Public Function DestHeight()
    If Rotate Then DestHeight = dest.ScaleWidth Else DestHeight = dest.ScaleHeight
End Function
Public Function DestWidth()
    If Rotate Then DestWidth = dest.ScaleHeight Else DestWidth = dest.ScaleWidth
End Function

Private Sub DrawPixel(ByVal X As Long, ByVal Y As Long, Color As Long)
    If Rotate Then RotateXY X, Y
    SetPixelV dest.hdc, X, Y, Color
End Sub
Public Sub RotateXY(ByRef X As Long, ByRef Y As Long)
    Dim temp As Long
    temp = X
    X = Y
    Y = dest.ScaleHeight - temp
End Sub
Public Sub DrawSemiCircle(ByVal X As Long, ByVal Y As Long, Radius As Long, Angle As Long, Width As Long, Optional EdgeColor As Long = vbBlack, Optional FillColor As Long = vbGreen, Optional DrawWidth As Long = 1, Optional Factor As Double = 1, Optional Start As Long = 1, Optional Edge As Boolean, Optional Steps As Long = 2, Optional ColorID As Long = -1, Optional OutsideEdge As Boolean = True)
    'Const Rot As Double = 1.5707963267949
    Dim pdegree As Double, L As Double, R As Double, A As Double, temp As Long, Blink As Boolean
    Dim oldStyle As Long, oldColor As Long
    If Rotate Then
        A = 90
        RotateXY X, Y
        Y = Y - 1
    End If
    L = DegreesToRadians(Angle + A)
    
    If Width < 0 Then ' And Width > -90 Then
        If Buttonmode Then
            R = DegreesToRadians(Angle + Width + A)
        Else
            R = DegreesToRadians(Angle - Width + A)
        End If
    Else
        R = DegreesToRadians(Angle + Width + A)
    End If
    
    If L < 0 Then L = L + 2 * PI
    
    'If L < 2 * PI And (R >= 2 * PI Or R <= 0) Then
    '    R = (2 * PI) - 1
    'End If
    
    'If R = 0 Then
        'L = L + PI2
        'R = R + PI2
    'End If
    
    If FillColor <> -1 Then
        If Width = 360 And Angle = 0 And Start = 0 Then
            oldStyle = dest.FillStyle
            oldColor = dest.FillColor
            dest.FillStyle = vbSolid
            dest.FillColor = FillColor
            dest.Circle (X, Y), Radius, FillColor
            dest.FillStyle = oldStyle
            dest.FillColor = oldColor
        Else
            dest.DrawWidth = 2
            If EdgeColor = -1 Then temp = Radius Else temp = Radius - 1
            If Steps = 1 Then Start = Start + 1 ': temp = temp - 1
            For pdegree = Start To temp Step Steps
                dest.Circle (X, Y), pdegree, FillColor, L, R, Factor
                If pdegree > 0 Then dest.Circle (X, Y), pdegree - 1, FillColor, L, R, Factor
            Next
        End If
    End If
    
    If EdgeColor <> -1 Then
        dest.DrawWidth = DrawWidth
        dest.Circle (X, Y), IIf(Edge, Start, Radius), EdgeColor, L, R, Factor
    End If
    
    If AntiAliasing And ColorID > -1 Then
        temp = SampleAA(X, Y, L, R, Radius, OutsideEdge)
        Blink = EdgeColor = ColorList(Color).Blink
        If temp > -1 Then
            DrawAA X, Y, temp, Color, Blink
        End If
    End If
End Sub
Public Sub DrawSquare(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Color As OLE_COLOR, Optional FillColor As OLE_COLOR = -1)
    'Dim EdgePen As Long, FillBrush As Long, temp As Long
    'EdgePen = CreatePen(PS_SOLID, 1, Color)
    'DeleteObject SelectObject(Dest.hdc, EdgePen)
    'FillBrush = CreateSolidBrush(FillColor)
    'DeleteObject SelectObject(Dest.hdc, FillBrush)
    'If Rotate Then
    '    temp = Dest.ScaleHeight - X - Width + 1
    '    RectangleX Dest.hdc, Y, temp, Y + Height - 1, temp + Width - 1
    'Else
    '    RectangleX Dest.hdc, X, Y, X + Width - 1, Y + Height - 1
    'End If
    
    Dim temp As Long
    dest.DrawWidth = 1
    'If FillColor > -1 Then
        If FillColor > -1 Then dest.FillColor = FillColor
    '    Dest.FillStyle = vbSolid
    'Else
    '    Dest.FillStyle = 1
    'End If
    
    If Rotate Then
        'temp = Dest.ScaleHeight - X
        'Dest.Line (Y, temp)-(Y + Height - 1, temp - Width - 1), Color, B
        
        temp = dest.ScaleHeight - X - Width
        dest.Line (Y, temp)-(Y + Height - 1, temp + Width - 1), Color, B
    Else
        dest.Line (X, Y)-(X + Width - 1, Y + Height - 1), Color, B
    End If
End Sub
Public Sub DrawLine(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Color As OLE_COLOR)
    Dim temp As Long
    If Rotate Then
        temp = dest.ScaleHeight - X
        dest.Line (Y, temp)-(Y + Height - 1, temp - Width + 1), Color
    Else
        dest.Line (X, Y)-(X + Width - 1, Y + Height - 1), Color
    End If
End Sub





'FONT
Public Sub DrawText(ByVal X As Long, ByVal Y As Long, Text As String, Optional Color As OLE_COLOR, Optional Size As Single)
    Dim tempstr() As String, temp As Long, oldsize As Single
    
    If Size > 0 Then
        If Size <> dest.Font.Size Then
            SwitchToUnRotated
            oldsize = dest.Font.Size
            dest.Font.Size = Size
        End If
    End If
    
    If Rotate Then
        SwitchToRotated
        Y = Y - 1
        RotateXY X, Y
    Else
        SwitchToUnRotated
    End If
    dest.ForeColor = Color
    
    'If InStr(Text, vbNewLine) > 0 Then Text = Replace(Text, vbNewLine, " ")
    
    If InStr(Text, vbNewLine) = 0 Then
        dest.CurrentX = X
        dest.CurrentY = Y
        dest.Print Text
    Else
        tempstr = Split(Text, vbNewLine)
        For temp = 0 To UBound(tempstr)
            dest.CurrentX = X
            dest.CurrentY = Y
            dest.Print tempstr(temp)
            
            If Rotate Then
                X = X + dest.TextHeight(tempstr(temp))
            Else
                Y = Y + dest.TextHeight(tempstr(temp))
            End If
        Next
    End If
    
    If oldsize > 0 Then
        SwitchToUnRotated
        dest.Font.Size = oldsize
    End If
End Sub

Public Sub SwitchToRotated()
    If Not isRotated Then
        F.lfEscapement = 900 'rotation angle, in tenths
        F.lfFacename = dest.Font.Name + Chr$(0)
        'F.lfHeight = (Dest.Font.Size * -20) / Screen.TwipsPerPixelY
        F.lfHeight = -MulDiv((dest.Font.Size), (GetDeviceCaps(dest.hdc, LOGPIXELSY)), 72)
        F.lfWeight = IIf(dest.Font.Bold, FW_BOLD, FW_NORMAL)
        F.lfCharSet = dest.Font.Charset
        hFont = CreateFontIndirect(F)
        hPrevFont = SelectObject(dest.hdc, hFont)
        isRotated = True
    End If
End Sub
Public Sub SwitchToUnRotated()
    If isRotated Then
        DeleteObject hFont
        hFont = SelectObject(dest.hdc, hPrevFont)
        isRotated = False
    End If
End Sub

Public Sub DrawLCARButton3D(X As Long, Y As Long, Width As Long, Height As Long, Text As String, GID As Long, Optional LeftSideWidth As Long, Optional RightSideWidth As Long, Optional WhiteSpace As Long = 4, Optional TextAlign As Long = 4, Optional Border As Long = 5)
    Dim temp As Long, temp2 As Long, Unit As Long, Start As Long, tX As Long, tY As Long, Color As OLE_COLOR, GID2 As OLE_COLOR
    'If LeftSideWidth > 0 Or RightSideWidth > 0 Then
        'If Height Mod 2 = 0 Then Height = Height + 1 'must be an even number
    'End If
    Unit = Height / 2
    Buttonmode = True
    
    GID2 = CacheGradient(GradientList(GID).LightColor, GradientList(GID).DarkColor)
    
    With GradientList(GID)
    
    If LeftSideWidth > 0 Then
        If LeftSideWidth < Height Then LeftSideWidth = Height

        'DrawCurvedGradientSquare Dest.hdc, .StartColor, .EndColor, X + Border, Y + Border, LeftSideWidth - Border * 2, Height - Border * 2
        'DrawGradientLine Dest.hdc, X + Unit - Border, Y + Border, LeftSideWidth - Border - 2, GID2, False, Height - Border * 2, True
        
        DrawRhomboid .LightColor, X + Unit, Y, LeftSideWidth - Border - Unit, Border, False, True, True, False
        DrawRhomboid .DarkColor, X + Unit, Y + Height - 1 - Border, LeftSideWidth - Unit, Border, False, True, False, False
        DrawRhomboid .DarkColor, X + LeftSideWidth - Border - 1, Y, Border, Height + 1, True, True, False, True
        
        DrawGradientSemiCircle dest.hdc, GID2, False, X + Unit, Y + Unit - 1, Unit - Border + 1, Unit - 1, 1, 90, 180
        
        temp = LeftSideWidth + WhiteSpace
    End If
    
    If RightSideWidth > 0 Then
        If RightSideWidth < Height Then RightSideWidth = Height
        Start = (X + Width) - RightSideWidth
        
        'DrawCurvedGradientSquare Dest.hdc, .StartColor, .EndColor, X + Border + Width - RightSideWidth - 1, Y + Border, LeftSideWidth - Border * 2, Height - Border * 2
        DrawRhomboid .LightColor, X + Width - RightSideWidth - 1, Y, RightSideWidth - Unit + Border, Border, True, False, True, False
        DrawRhomboid .DarkColor, X + Width - RightSideWidth - 2, Y + Height - 1 - Border, RightSideWidth - Unit + Border, Border, True, False, False, False
        DrawRhomboid .LightColor, X + Width - RightSideWidth - 1, Y, Border, Height - Border, True, True, True, True
        'DrawGradientSemiCircle Dest.hdc, GID2, False, X + Width - Unit, Y + Unit - 1, Unit - Border + 1, Unit - 1, 1, 90, -181
        DrawGradientSemiCircle dest.hdc, GID2, False, X + Width - Unit, Y + Unit - 1, Unit - Border + 1, Unit - 1, 1, 90, -180
        
        temp2 = WhiteSpace + RightSideWidth + 2
    End If
    
    'DrawSquare temp + X, Y, Width - temp - temp2, Height, EdgeColor, FillColor
    DrawCurvedGradientSquare dest.hdc, .StartColor, .EndColor, temp + X, Y + Border, Width - temp - temp2, Height - Border * 2
        
    DrawRhomboid .LightColor, temp + X + 1, Y, Width - temp - temp2 - Border, Border, True, True, True, False
    DrawRhomboid .LightColor, temp + X + 1, Y, Border, Height - Border, True, True, True, True
    DrawRhomboid .DarkColor, temp + X + 1, Y + Height - Border - 1, Width - temp - temp2, Border, True, True, False, False
    DrawRhomboid .DarkColor, temp + X + Width - Border - temp - temp2, Y, Border, Height + 1, True, True, False, True
    
    If Len(Text) > 0 Then
        Select Case TextAlign
            Case 1, 2, 3: tY = Y  'top row
            Case 4, 5, 6: tY = Y + (Unit - dest.TextHeight(Text) / 2)  'middle row
            Case 7, 8, 9: tY = Y + Height - dest.TextHeight(Text) 'bottom row
        End Select
        Select Case TextAlign
            Case 1, 4, 7: tX = temp + X + 3 ' left column
            Case 2, 5, 8: tX = X + ((Width - temp - temp2) / 2) - (dest.TextWidth(Text) / 2) + temp 'middle column
            Case 3, 6, 9: tX = X + Width - temp2 - dest.TextWidth(Text) - 2 'right column
        End Select
        
        Color = vbBlack
        DrawText tX, tY, Text, vbRed ' Color
    End If
    
    End With
    
    Buttonmode = False
End Sub

Private Function AddColor(Name As String, Color As OLE_COLOR) As Long
    AddColor = ColorCount
    ColorCount = ColorCount + 1
    ReDim Preserve ColorList(ColorCount)
    With ColorList(ColorCount - 1)
        .Name = Name
        
        .Color = Color
        .Blink = AlterBrightness(Color, Brightness)
        
        .Gradient = CacheGradient(Color, vbBlack)
        .BlinkColor = CacheGradient(.Blink, vbBlack)
        .Nemesis = CacheGradient(Color, .Blink)
    End With
End Function
Public Sub SetupLCARcolors()
    If ColorCount = 0 Then
        AddColor "Black", LCAR_Black
        AddColor "White", LCAR_White
        
        AddColor "Red", LCAR_Red
        AddColor "Dark Orange", LCAR_DarkOrange
        AddColor "Orange", LCAR_Orange
        AddColor "Light Orange", LCAR_LightOrange
        AddColor "Dark Yellow", LCAR_DarkYellow
        AddColor "Yellow", LCAR_Yellow
        AddColor "Dark Blue", LCAR_DarkBlue
        AddColor "Light Blue", LCAR_LightBlue
        AddColor "Dark Purple", LCAR_DarkPurple
        AddColor "Purple", LCAR_Purple
        AddColor "Light Purple", LCAR_LightPurple
        
        AddColor "Legacy Yellow", cLCAR_Yellow
        AddColor "Legacy Green", cLCAR_Green
        AddColor "Legacy Light Blue", cLCAR_LightBlue
        AddColor "Legacy Blue", cLCAR_Blue
    End If
End Sub
Public Sub AddColorsToList(ID As Long, Optional Selected As Long = -1)
    Dim temp As Long
    SetupLCARcolors
    LCAR_ClearList ID
    For temp = 0 To ColorCount - 1
        LCAR_AddListItem ID, ColorList(temp).Name, ColorList(temp).Color, , , , , Selected = temp
    Next
End Sub
Public Sub AddNumbersToList(ID As Long, ByVal Start As Single, Increment As Single, Finish As Single, Optional Selected As Single)
    Dim temp As Long, count As Long, temp2 As Single
    count = ((Finish - Start) / Increment)
    For temp = 0 To count
        temp2 = Start + (Increment * temp)
        LCAR_AddListItem ID, CStr(temp2), , , , , , Selected = Start
    Next
End Sub
Public Function LCAR_ColorIDfromColor(Color As OLE_COLOR) As Long
    Dim temp As Long
    SetupLCARcolors
    LCAR_ColorIDfromColor = -1
    For temp = 0 To ColorCount - 1
        If ColorList(temp).Color = Color Or ColorList(temp).Blink = Color Then
            LCAR_ColorIDfromColor = temp
            Exit For
        End If
    Next
End Function
Public Function LCAR_ColorIDfromName(Name As String) As Long
    Dim temp As Long
    SetupLCARcolors
    LCAR_ColorIDfromName = -1
    For temp = 0 To ColorCount - 1
        If StrComp(ColorList(temp).Name, Name, vbTextCompare) = 0 Then
            LCAR_ColorIDfromName = temp
            Exit For
        End If
    Next
End Function
