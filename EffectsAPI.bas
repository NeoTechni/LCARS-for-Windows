Attribute VB_Name = "EffectsAPI"
Option Explicit
    
Public Const HighliteColor As Long = LCAR_LightBlue
Public Const btntxtbox As String = "btntxtbox"

'keyboard
Public Text As String, SelStart As Long, SelLength As Long, Symbols As Boolean, Shift As Boolean, Caps As Boolean, TextID As Long, DefaultText As String, OldText As String
Public Operation As String

'sensor grid
Public Cx As Double, Cy As Double, pX As Double, pY As Double

Public Function isShift() As Boolean
    isShift = Shift
End Function

Public Function CursorIsVisible() As Boolean
    Static oldPosition As Long, oldWidth As Long
    CursorIsVisible = (Second(Now) Mod 2 = 0) Or (oldPosition <> SelStart) Or (oldWidth <> SelLength)
    oldPosition = SelStart
    oldWidth = SelLength
End Function

Private Sub Emergency()
    TextID = LCAR_FindLCAR(btntxtbox)
End Sub

Public Function CursorPosition(Optional Position As Long = -1) As Long
    Dim oldsize As Long, Start As Long
    oldsize = dest.Font.Size
    Emergency
    With LCAR_ButtonList(TextID)
        dest.Font.Size = .TextSize
        If Position = -1 Then
            Start = dest.TextWidth(GetSelText)
        Else
            Start = .X + 4
            If Start < 0 Then Start = DestWidth - .X
            If Position > 0 Then Start = Start + dest.TextWidth(Mid(Text, 1, Position)) ': Debug.Print Mid(Text, 1, Position)
        End If
        CursorPosition = Start
    End With
    dest.Font.Size = oldsize
End Function

Public Sub DrawCursor()
    Dim X As Long, Width As Long, Y As Long
    Emergency
    X = CursorPosition(SelStart)
    If SelLength <> 0 Then
        Width = CursorPosition
        If SelLength < 0 Then Width = -Width
    End If
    With LCAR_ButtonList(TextID)
        If Width <> 0 Then
            If Width > 0 Then
                DrawSquare X, Y + 2, Width, .Height - 1, HighliteColor, HighliteColor
                DrawText X - 1, Y + 2, GetSelText, LCAR_Orange, .TextSize
            Else
                DrawSquare X + Width, Y + 2, Abs(Width) + 1, .Height - 1, HighliteColor, HighliteColor
                DrawText X + Width, Y + 2, GetSelText, LCAR_Orange, .TextSize
            End If
        End If
        
        If CursorIsVisible Then
            DrawSquare X - 1, Y + 2, 3, .Height - 1, LCAR_Orange, LCAR_Orange
            
            'DrawLine X, .Y, 1, .Height, LCAR_Orange
        Else
            LCAR_ButtonList(TextID).IsClean = False
        End If
    End With
End Sub

Public Function KeyboardIsVisible() As Boolean
    KeyboardIsVisible = LCARlists(2).Visible
End Function

Public Sub ShowKeyboard(Default As String, Optional Op As String)
    Emergency
    Operation = Op
    HideAllGroups 5
    HideAllLists 2
    LCAR_ButtonList(preview).Visible = False
    DefaultText = Default
    Text = Default
    SelStart = 0
    SelLength = Len(Default)
    LCAR_SetText "btntxtbox", Default, , , , , , , , False
End Sub

Public Function HideKeyboard() As String
    HideAllGroups 3
    HideAllLists 0
    LCAR_ButtonList(preview).Visible = True
    
    HideKeyboard = Text
End Function

Public Sub DrawEffects()
    If KeyboardIsVisible Then
        IncrementGrid 0.005 '1
        DrawGridAuto
    End If
End Sub

Public Function ProcessKey(Key As String)
    Static LCARid As Long, SymbolID As Long
    Emergency
    If Len(Key) > 0 Then
        If Len(Key) = 1 Then
            SetSelText Key
            If Shift Then
                Shift = False
                If LCARid = 0 Then LCARid = LCAR_FindLCAR("frmbottom", , 6)
                LCAR_Blink LCARid, False
            End If
            If Symbols Then
                Symbols = False
                If SymbolID = 0 Then SymbolID = LCAR_FindLCAR("frmbottom", , 4)
                LCAR_Blink SymbolID, False
            End If
        Else
            Select Case LCase(Key)
                Case "left"
                    If Shift Then
                        SelLength = SelLength - 1
                        If SelLength < 0 Then
                            If SelStart + SelLength < 0 Then SelLength = SelLength + 1
                        End If
                    Else
                        SelLength = 0
                        SelStart = SelStart - 1
                        If SelStart < 0 Then ProcessKey "end"
                    End If
                    
                Case "right"
                    If Shift Then
                        SelLength = SelLength + 1
                        If SelLength > 0 Then
                            If SelStart + SelLength > Len(Text) Then SelLength = SelLength - 1
                        End If
                    Else
                        SelLength = 0
                        SelStart = SelStart + 1
                        If SelStart > Len(Text) Then ProcessKey "home"
                    End If
                    
                Case "home"
                    If Shift Then SelLength = SelStart Else SelLength = 0
                    SelStart = 0
                    
                Case "end"
                    If Shift Then SelLength = -(Len(Text) - SelStart) Else SelLength = 0
                    SelStart = Len(Text)
                    
                Case "delete"
                    If Len(Text) > 0 Then
                        If SelLength = 0 Then
                            SelLength = 1
                        End If
                        SetSelText Empty
                    End If
                    
                Case "backspace"
                    If Len(Text) > 0 Then
                        If SelLength = 0 Then
                            If SelStart > 0 Then
                                SelStart = SelStart - 1
                                SelLength = 1
                            End If
                        End If
                        SetSelText Empty
                    End If
                    
                Case "space"
                    ProcessKey " "
                    
                Case "shift"
                    LCARid = LCAR_FindLCAR("frmbottom", , 6)
                    LCAR_Blink LCARid, Not LCAR_isBlinking(LCARid)
                    Shift = LCAR_isBlinking(LCARid)
                    
            End Select
        End If
    End If
    
    LCAR_SetText LCAR_ButtonList(TextID).Name, Text, , , , , , , , False
End Function

Public Function GetSelText() As String
    If Abs(SelLength) = Len(Text) Then
        GetSelText = Text
    Else
        If SelLength > 0 Then
            GetSelText = Mid(Text, SelStart + 1, SelLength)
        ElseIf SelLength < 0 Then
            GetSelText = Mid(Text, SelStart + SelLength + 1, Abs(SelLength))
        End If
    End If
End Function

Public Sub SetSelText(Key As String)
    Dim LSide As Long, RSide As Long, L As String, R As String
    If SelLength > 0 Then
        LSide = SelStart
        RSide = Len(Text) - SelStart - SelLength
        SelStart = SelStart + Len(Key)
    ElseIf SelLength < 0 Then
        LSide = SelStart + SelLength
        RSide = Len(Text) - SelStart
        SelStart = LSide + Len(Key)
    Else
        LSide = SelStart
        RSide = Len(Text) - SelStart
        SelStart = SelStart + Len(Key)
    End If
    If LSide > 0 Then L = Left(Text, LSide)
    If RSide > 0 Then R = Right(Text, RSide)
    Text = L + Key + R
    SelLength = 0
End Sub

Public Function GetKey(Index As Long) As String
    Dim tempstr As String
    If Symbols Then tempstr = LCARlists(2).ListItems(Index).Side
    If Len(tempstr) = 0 Then
        tempstr = LCARlists(2).ListItems(Index).Text
        If (Shift And Caps) Or ((Not Shift) And (Not Caps)) Then
            tempstr = LCase(tempstr)
        'ElseIf (Shift And Not Caps) Or (Not Shift And Caps) Then
        '    tempstr = UCase(tempstr)
        End If
    End If
    GetKey = tempstr
End Function
Public Function vKey2String(vKey As Long, Optional Default As String = "Press Any Key") As String
    Select Case vKey
        
        Case 8: vKey2String = "Backspace"
       
        Case 13: vKey2String = "Enter"

        Case 16: vKey2String = "Shift"
        Case 17: vKey2String = "Ctrl"
        Case 18: vKey2String = "Alt"
        Case 19: vKey2String = "Pause"
        Case 20: vKey2String = "Caps Lock"
        
        Case 27: vKey2String = "Escape"
        
        Case 32: vKey2String = "Space"
        Case 33: vKey2String = "Page Up"
        Case 34: vKey2String = "Page Down"
        Case 35: vKey2String = "End"
        Case 36: vKey2String = "Home"
        Case 37: vKey2String = "Left"
        Case 38: vKey2String = "Up"
        Case 39: vKey2String = "Right"
        Case 40: vKey2String = "Down"
        
        Case 45: vKey2String = "Insert"
        Case 46: vKey2String = "Delete"
        
        Case 48 To 57: vKey2String = Chr(vKey) '0-9
        
        Case 65 To 90: vKey2String = LCase(Chr(vKey)) 'a-z
                
        Case 91: vKey2String = "Start"
        
        Case 93: vKey2String = "Menu"
        
        Case 96 To 105: vKey2String = Chr(vKey - 48)
        Case 106: vKey2String = "*"
        Case 107: vKey2String = "+"
        
        Case 109, 189: vKey2String = "-" '
        Case 110, 190: vKey2String = "."
        Case 111, 191: vKey2String = "/"
        Case 220: vKey2String = "\"
        'Case 186: vKey2String = ";"
        
        'Case 112 To 123: vKey2String = "F" & vKey - 111
        
        Case 144: vKey2String = "Num Lock"
        Case 145: vKey2String = "Scroll Lock"
         
        Case 192: vKey2String = "`"
        
        Case Else: vKey2String = Default: MsgBox vKey
    End Select
End Function


Public Sub SetupEffects()
    Cx = 0.5
    Cy = 0.5
    pX = 0.5
    pY = 0.5
End Sub

Public Sub DrawGridAuto()
    Dim Width As Long, Height As Long
    Width = DestWidth - 223
    Height = DestHeight - 338
    DrawSensorGrid 110, 88, Width, Height, Width * Cx, Height * Cy
End Sub

Public Sub IncrementGrid(Optional Speed As Double = 0.05)
    If Cx < pX Then
        Cx = Cx + Speed
        If Cx > pX Then Cx = pX
    ElseIf Cx > pX Then
        Cx = Cx - Speed
        If Cx < pX Then Cx = pX
    End If
    If Cy < pY Then
        Cy = Cy + Speed
        If Cy > pY Then Cy = pY
    ElseIf Cy > pY Then
        Cy = Cy - Speed
        If Cy < pY Then Cy = pY
    End If
    If Cx = pX And Cy = pY Then
        Randomize Timer
        pX = Rnd
        pY = Rnd
    End If
End Sub

Public Sub DrawSensorGrid(X As Long, Y As Long, Width As Long, Height As Long, oX As Long, oY As Long, Optional StartSize As Double = 0.1, Optional Factor As Double = 0.95, Optional Border As Long = 2)
    Dim Cx As Double, cWidth As Double, temp As Long
    Static WasVisible As Boolean
    'Units = 2 ^ Lines
    
    DrawSquare X, Y, Width, Height, vbBlack, vbBlack
    DrawSquare X + Border, Y + Border, Width - Border * 2, Height - Border * 2, vbWhite, IIf(RedAlert, LCAR_Red, LCAR_DarkBlue)
        
    cWidth = StartSize * oX
    Cx = oX + X
    
    dest.DrawWidth = 3
    DrawLine Cx, Y + 1, 1, Height - 2, vbWhite
    DrawLine X + 1, oY + Y, Width - 2, 1, vbWhite
    dest.DrawWidth = 1
    
    temp = X + Border
    Do While Cx > temp
        cWidth = cWidth * Factor
        Cx = Cx - cWidth
        If Cx > X Then DrawLine Cx, Y, 1, Height, vbWhite
        If cWidth < 2 Then Cx = 0
    Loop
    
    cWidth = StartSize * (Width - oX)
    Cx = oX + X
    temp = X + Width - Border
    Do While Cx < temp
        cWidth = cWidth * Factor
        Cx = Cx + cWidth
        If Cx < temp Then DrawLine Cx, Y, 1, Height, vbWhite
        If cWidth < 2 Then Cx = X + Width
    Loop
    
    cWidth = StartSize * oY
    Cx = oY + Y
    temp = Y + Border
    Do While Cx > temp
        cWidth = cWidth * Factor
        Cx = Cx - cWidth
        If Cx > temp Then DrawLine X, Cx, Width, 1, vbWhite
        If cWidth < 2 Then Cx = 0
    Loop
    
    cWidth = StartSize * (Height - oY)
    Cx = oY + Y
    temp = Y + Height - Border
    Do While Cx < temp
        cWidth = cWidth * Factor
        Cx = Cx + cWidth
        If Cx < temp Then DrawLine X, Cx, Width, 1, vbWhite
        If cWidth < 2 Then Cx = temp
    Loop
    
    'If Not CursorIsVisible Then
    DrawCursor ' Else LCAR_ButtonList(TextID).IsClean = False
    'If WasVisible <> CursorIsVisible Then
    '    WasVisible = Not WasVisible
    '    DrawCursor
    'End If
End Sub
