' draw arrow
Option Explicit

Dim linemode(1, 1, 1, 1)
linemode(0, 0, 0, 0) = ""
linemode(0, 0, 0, 1) = ""
linemode(0, 0, 1, 0) = ""
linemode(0, 0, 1, 1) = "„Ÿ"
linemode(0, 1, 0, 0) = ""
linemode(0, 1, 0, 1) = "„¡"
linemode(0, 1, 1, 0) = "„¢"
linemode(0, 1, 1, 1) = "„¦"
linemode(1, 0, 0, 0) = ""
linemode(1, 0, 0, 1) = "„¤"
linemode(1, 0, 1, 0) = "„£"
linemode(1, 0, 1, 1) = "„¨"
linemode(1, 1, 0, 0) = "„ "
linemode(1, 1, 0, 1) = "„¥"
linemode(1, 1, 1, 0) = "„§"
linemode(1, 1, 1, 1) = "„©"

Dim top_joint, bottom_joint, left_joint, right_joint
top_joint    = Array("„Ÿ", "„¡", "„¢", "„¦", "„¤", "„£", "„¨", "„ ", "„¥", "„§", "„©", "„ª", "„¬", "„­", "„±", "„¯", "„®", "„³", "„«", "„°", "„²", "„´")
bottom_joint = Array("„Ÿ", "„¡", "„¢", "„¦", "„¤", "„£", "„¨", "„ ", "„¥", "„§", "„©", "„ª", "„¬", "„­", "„±", "„¯", "„®", "„³", "„«", "„°", "„²", "„´")
left_joint   = Array("„Ÿ", "„¡", "„¢", "„¦", "„¤", "„£", "„¨", "„ ", "„¥", "„§", "„©", "„ª", "„¬", "„­", "„±", "„¯", "„®", "„³", "„«", "„°", "„²", "„´")
right_joint  = Array("„Ÿ", "„¡", "„¢", "„¦", "„¤", "„£", "„¨", "„ ", "„¥", "„§", "„©", "„ª", "„¬", "„­", "„±", "„¯", "„®", "„³", "„«", "„°", "„²", "„´")

Call DrawLine("Top")

Sub DrawLine(direct)
        Dim ln
        Dim defchar
        defchar = "„ "

        Select Case direct
        Case "Bottom": ln = CStr(linemode(IsStrMatch(GetTop, top_joint), 1, IsStrMatch(GetLeft, left_joint), IsStrMatch(GetRight, right_joint)))
        Case "Left":   ln = CStr(linemode(IsStrMatch(GetTop, top_joint), IsStrMatch(GetBottom, bottom_joint), 1, IsStrMatch(GetRight, right_joint)))
        Case "Right":  ln = CStr(linemode(IsStrMatch(GetTop, top_joint), IsStrMatch(GetBottom, bottom_joint), IsStrMatch(GetLeft, left_joint), 1))
        Case "Top":    ln = CStr(linemode(1, IsStrMatch(GetBottom, bottom_joint), IsStrMatch(GetLeft, left_joint), IsStrMatch(GetRight, right_joint)))
        Case Else:     ln = CStr(linemode(IsStrMatch(GetTop, top_joint), 1, IsStrMatch(GetLeft, left_joint), IsStrMatch(GetRight, right_joint)))
        End Select

        If ln = "" Then ln = defchar

        Call InsertText(ln)

        Select Case direct
        Case "Bottom": If Not MoveBottom Then Exit Sub
        Case "Left":   If Not MoveLeft   Then Exit Sub
        Case "Right":  If Not MoveRight  Then Exit Sub
        Case "Top":    If Not MoveTop    Then Exit Sub
        Case Else:     If Not MoveBottom Then Exit Sub
        End Select

        Call InsertText(CStr(linemode(IsStrMatch(GetTop, top_joint), IsStrMatch(GetBottom, bottom_joint), IsStrMatch(GetLeft, left_joint), IsStrMatch(GetRight, right_joint))))
        CancelMode
End Sub

Function GetCur()
        GetCur = ""
        Dim x, s
        x = CLng(ExpandParameter("$x"))
        s = GetLineStr(CLng(0))
        If CLng(x) - 1 >= ByteLen(s) Then Exit Function
        GetCur = Mid(s, x, 1)
End Function

Function GetTop:    GetTop = GetTopOrBottom("Top"):    End Function
Function GetBottom: GetBottom = GetTopOrBottom("Bottom"): End Function

Function GetTopOrBottom(direct)
        GetTopOrBottom = ""
        Dim x, y
        x = CLng(ExpandParameter("$x"))
        y = CLng(ExpandParameter("$y"))

        If direct = "Top" Then
            If y <= 1 Then Exit Function
            GetTopOrBottom = Mid(GetLineStr(y - 2), x, 1)
        Else
            If y >= GetLineCount(0) Then Exit Function
            GetTopOrBottom = Mid(GetLineStr(y), x, 1)
        End If
End Function

Function GetLeft()
        GetLeft = ""
        Dim x, y
        x = CLng(ExpandParameter("$x"))
        y = CLng(ExpandParameter("$y"))
        If x <= 1 Then Exit Function
        GetLeft = Mid(GetLineStr(y - 1), x - 1, 1)
End Function

Function GetRight()
        GetRight = ""
        Dim x, y, s
        x = CLng(ExpandParameter("$x"))
        y = CLng(ExpandParameter("$y"))
        s = GetLineStr(y - 1)
        If x >= Len(s) Then Exit Function
        GetRight = Mid(s, x + 1, 1)
End Function

Function IsStrMatch(s, arr)
        IsStrMatch = 0
        Dim ar
        For Each ar In arr
                If s = ar Then
                        IsStrMatch = 1
                        Exit Function
                End If
        Next
End Function

Sub InsertText(ByVal c)
        If c = "" Then Exit Sub
        Dim isrep
        If IsStrMatch(GetCur, Array(" ", "„Ÿ", "„¡", "„¢", "„¦", "„¤", "„£", "„¨", "„ ", "„¥", "„§", "„©", "„ª", "„¬", "„­", "„±", "„¯", "„®", "„³", "„«", "„°", "„²", "„´")) = 1 Then
                isrep = True
        Else
                isrep = False
        End If
        If isrep Then
                BeginSelect
                MoveRight
        End If
        Call InsText(CStr(c))
        MoveLeft
End Sub

Function MoveTop()
        MoveTop = False
        Dim x, y
        x = CLng(ExpandParameter("$x"))
        y = CLng(ExpandParameter("$y"))
        If y = 1 Then Exit Function
        Dim byteLenCur, sp, s
        byteLenCur = ByteLen(Mid(GetLineStr(CLng(y)), 1, x))
        sp = 0
        s = GetLineStr(CLng(y - 1))
        If ByteLen(s) < byteLenCur Then sp = byteLenCur - ByteLen(s) - 2
        Editor.Up
        MoveTop = True
        If sp > 0 Then Call InsText(CStr(Space(sp)))
End Function

Function MoveBottom()
        MoveBottom = True
        Dim x, y, byteLenCur, sp
        x = CLng(ExpandParameter("$x"))
        y = CLng(ExpandParameter("$y"))
        byteLenCur = ByteLen(Mid(GetLineStr(CLng(y)), 1, x))
        If IsFinalLine(y) Then
                GoLineEnd
                InsertCR
                sp = byteLenCur - ByteLen(Mid(GetLineStr(CLng(CLng(ExpandParameter("$y")))), 1, CLng(ExpandParameter("$x")))) - 2
        Else
                Editor.Down
                Dim s: s = GetLineStr(CLng(ExpandParameter("$y")))
                If ByteLen(s) < byteLenCur Then sp = byteLenCur - ByteLen(s) - 2 Else sp = 0
        End If
        If sp > 0 Then Call InsText(CStr(Space(sp)))
End Function

Function MoveLeft()
        MoveLeft = False
        If CLng(ExpandParameter("$x")) = 1 Then Exit Function
        Editor.Left
        If GetCur = " " And CLng(ExpandParameter("$x")) > 1 Then
                Editor.Left
                If GetCur <> " " Then Editor.Right
        End If
        MoveLeft = True
End Function

Function MoveRight()
        MoveRight = True
        Dim cur: cur = GetCur
        If cur = "" Then Exit Function
        Editor.Right
        If cur = " " And GetCur = " " Then Editor.Right
End Function

Function ByteLen(ByVal s)
        ByteLen = 0
        Dim i, c
        For i = 1 to Len(s)
                c = Mid(s, i, 1)
                If c = vbCr Or c = vbLf Then Exit For
                ByteLen = ByteLen + ByteSize(c)
        Next
End Function

Function ByteSize(ByVal c)
        ByteSize = 0
        If Len(c) = 0 Then Exit Function
        If (Asc(c) >= 1) And (Asc(c) <= 255) Then ByteSize = 1 Else ByteSize = 2
End Function

Function ByteMid(ByVal s, ByVal index, ByVal length)
        ByteMid = ""
        Dim i, bidx, c
        bidx = 0
        For i = 1 To Len(s)
                c = Mid(s, i, 1)
                bidx = bidx + ByteSize(c)
                If bidx >= index Then
                        ByteMid = ByteMid & c
                        length = length - 1
                        If length = 0 Then Exit For
                End If
        Next
End Function

Function IsFinalLine(lineno)
        IsFinalLine = (lineno = GetLineCount(CLng(0)))
End Function

Function InsertCR()
        Call Char(CLng(13))
End Function
