Attribute VB_Name = "ProgramAPI"
Option Explicit

Public Const NoneSelected As String = "There are no selected items"
Public Enum OpToDo
    DoNothing
    DoDelete
End Enum
Public DoOp As OpToDo
Public Bookmarks() As String, BookmarkCount As Long, NeedsSaving As Boolean, preview As Long

Public Function sec2time(ByVal whattime As Long) As String
    On Error Resume Next
    If InStr(whattime, ".") > 0 Then whattime = Left(whattime, ".") - 1
    Const time_min As Long = 60, time_hour As Long = 3600
    Dim time_hours As Byte, time_minutes As Byte, time_seconds As Byte

    time_hours = whattime \ time_hour
    whattime = whattime Mod time_hour
    time_minutes = whattime \ time_min
    whattime = whattime Mod time_min
    time_seconds = whattime

    'If time_hours = 0 Then
    '    sec2time = Format(time_minutes, "#0") & ":" & Format(time_seconds, "00")
    'Else
        sec2time = Format(time_hours, "#0:") & Format(time_minutes, "00") & ":" & Format(time_seconds, "00")
    'End If
End Function

Public Function IsInIDE() As Boolean
    IsInIDE = App.LogMode = 0
End Function

Public Sub ResizeLCARs()
    Dim temp As Long, Width As Long, Top As Long
    
    'If GroupList(5).Visible Then
        Width = DestWidth / 2 - 130
    
        temp = LCAR_FindLCAR("frmbottom", , 1) 'LEFT
        LCAR_ButtonList(temp).Width = Width + 1
        
        temp = LCAR_FindLCAR("frmbottom", , 7) 'DELETE
        LCAR_ButtonList(temp).Width = Width + 1
        
        temp = LCAR_FindLCAR("frmbottom", , 2) 'RIGHT
        With LCAR_ButtonList(temp)
            .Width = Width
            .X = DestWidth / 2 + 1
        End With
        
        temp = LCAR_FindLCAR("frmbottom", , 8) 'RIGHT
        With LCAR_ButtonList(temp)
            .Width = Width
            .X = DestWidth / 2 + 1
        End With
        
        Width = (DestHeight - 410) / 2
        'Temp = LCAR_FindLCAR("frmbottom", , 4) 'SYMBOLS
        'LCAR_ButtonList(Temp).Height = Width
        'Top = LCAR_ButtonList(Temp).Y
        
        'Temp = LCAR_FindLCAR("frmbottom", , 5) 'CAPS
        'With LCAR_ButtonList(Temp)
        '    .Y = Top + Width + 2
        '    .Height = Width - 1
        'End With
    'End If
    
    IsClean = False

End Sub
Public Sub HideAllGroups(Optional Except As Long = -1)
    LCARlists(0).Visible = False
    LCARlists(2).Visible = False
    
    Dim temp As Long
    'GroupList(5).Visible = True
    For temp = 3 To GroupCount - 1
        GroupList(temp).Visible = (temp = Except)
    Next
    
    ResizeLCARs
    IsClean = False
End Sub

Public Sub HideGroup(ID As Long, Optional Visible As Boolean)
    GroupList(ID).Visible = Visible
    IsClean = False
End Sub

Public Sub HideAllLists(Optional Except As Long = -1)
    Dim temp As Long
    If KeyboardIsVisible Then HideKeyboard
    For temp = 0 To LCARListCount - 1
        LCARlists(temp).Visible = (Except = temp)
    Next
    IsClean = False
End Sub
Public Sub RefreshPreview(Optional ForceText As String)
    Dim tempstr As String, tempstr2 As String, temp3 As Long
                With LCAR_ButtonList(preview)
                    .IsClean = False
                    '.Visible = True
                    If Len(ForceText) > 0 Then
                        .Text = ForceText
                    Else
                    
                    Select Case LCARlists(0).SelectedItems
                        Case 0: .Text = NoneSelected
                        Case 1
                            If LCARlists(0).SelectedItem > -1 And LCARlists(0).SelectedItem < LCARlists(0).ListCount Then
                                tempstr = LCARlists(0).ListItems(LCARlists(0).SelectedItem).Text
                                tempstr2 = UCase(LCARlists(0).ListItems(LCARlists(0).SelectedItem).Side)
                                temp3 = LCARlists(0).TotalSize
                                'If Len(tempstr2) > 0 Then
                                    If Len(tempstr2) = Empty Then
                                        .Text = tempstr & vbNewLine & "File Folder"
                                    Else
                                        tempstr = tempstr & "." & LCase(tempstr2)
                                        .Text = tempstr & vbNewLine & FileTypeName(tempstr2, , "The ""*"" extention has no association at this time") & vbNewLine & "This file occupies " & SizeToText(temp3, " Quads", " KiloQuads", " MegaQuads", " GigaQuads", 2)
                                    End If
                                'Else
                                '    .Text = tempstr
                                'End If
                                
                            End If
                        Case Else
                            .Text = "There are " & LCARlists(0).SelectedItems & " selected items occupying a total of " & SizeToText(LCARlists(0).TotalSize, " Quads", " KiloQuads", " MegaQuads", " GigaQuads", 2)
                    End Select
                    
                    End If
                End With
                LCAR_DrawLCARs
End Sub

Public Sub API_ShowOpenWith(ListID As Long)
    Dim temp As Long, count As Long, List() As String, Program As String, ProgramName As String ', Extention As String
    Dim temp2 As Long, count2 As Long, List2() As String
    'LCAR_ClearList ListID
    'Extention = GetExtention(Filename)
    'count = EnumOpenWith(Extention, List)
    count = EnumPrograms(List)
    For temp = 0 To count - 1
        count2 = EnumVerbs(List(temp), List2, False)
        For temp2 = 0 To count2 - 1
            Program = GetVerbPath(List(temp), List2(temp2))
            LCAR_AddListItem ListID, EXEname(Program), , , , Program, , , List2(temp2)
        Next
    Next
    'LCARlists(ListID).Visible = True
    'IsClean = False
End Sub

Public Sub API_LoadBookmarks()
    Dim temp As Long, count As Long
    count = Val(GetSetting("LCAR", "Bookmarks", "Count", "0"))
    For temp = 0 To count - 1
        API_BookmarkFolder GetSetting("LCAR", "Bookmarks", CStr(temp)), False
    Next
End Sub

Public Function API_IsBookmarked(Path As String) As Boolean
    API_IsBookmarked = API_FindBookmark(Path) > -1
End Function

Public Function API_FindBookmark(Path As String) As Long
    Dim temp As Long
    If Right(Path, 1) = "\" Then Path = Left(Path, Len(Path) - 1)
    
    API_FindBookmark = -1
    For temp = 0 To BookmarkCount - 1
        If StrComp(Path, Bookmarks(temp), vbTextCompare) = 0 Then
            API_FindBookmark = temp
        End If
    Next
End Function

Public Function API_BookmarkFolder(Path As String, Optional Save As Boolean = True) As Boolean
    If Right(Path, 1) = "\" Then Path = Left(Path, Len(Path) - 1)
        
    If StrComp(Path, ShellFolder("Desktop")) = 0 Then Exit Function
    If StrComp(Path, ShellFolder) = 0 Then Exit Function
    If StrComp(Path, ShellFolder("My Music")) = 0 Then Exit Function
    If StrComp(Path, ShellFolder("My Pictures")) = 0 Then Exit Function
    If StrComp(Path, ShellFolder("My Video")) = 0 Then Exit Function
    If Len(Path) <= 3 Then Exit Function
    
    If API_FindBookmark(Path) = -1 And InStr(Path, "\") > 0 And direxists(Path) Then
        API_BookmarkFolder = True
        BookmarkCount = BookmarkCount + 1
        If Save Then
            SaveSetting "LCAR", "Bookmarks", "Count", Str(BookmarkCount)
            SaveSetting "LCAR", "Bookmarks", CStr(BookmarkCount - 1), Path
        End If
        ReDim Preserve Bookmarks(BookmarkCount)
        Bookmarks(BookmarkCount - 1) = Path
    End If
End Function

Public Function API_DeleteBookmark(Path As String) As Boolean
    Dim temp As Long, Bookmark As Long
    Bookmark = API_FindBookmark(Path)
    If Bookmark > -1 Then
        For temp = Bookmark To BookmarkCount - 2
            Bookmarks(temp) = Bookmarks(temp + 1)
            SaveSetting "LCAR", "Bookmarks", CStr(temp), Bookmarks(temp)
        Next
        BookmarkCount = BookmarkCount - 1
        If BookmarkCount = 0 Then
            ReDim Bookmarks(BookmarkCount)
        Else
            ReDim Preserve Bookmarks(BookmarkCount)
        End If
        SaveSetting "LCAR", "Bookmarks", "Count", Str(BookmarkCount)
    End If
End Function

Public Sub API_ListBookmarks(ListID As Long)
    Dim temp As Long
    For temp = 0 To BookmarkCount - 1
        LCAR_AddFolder ListID, Bookmarks(temp), "Delete"
    Next
End Sub

Public Function IsFontInstalled(Name As String) As Boolean
    Dim temp As Long
    For temp = 0 To Screen.FontCount - 1
        If StrComp(Name, Screen.Fonts(temp), vbTextCompare) = 0 Then
            IsFontInstalled = True
            Exit For
        End If
    Next
End Function
