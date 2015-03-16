' draw arrow
Option Explicit

Dim linemode(1, 1, 1, 1) ' 上, 下, 左, 右
linemode(0, 0, 0, 0) = ""
linemode(0, 0, 0, 1) = ""
linemode(0, 0, 1, 0) = ""
linemode(0, 0, 1, 1) = "━"
linemode(0, 1, 0, 0) = ""
linemode(0, 1, 0, 1) = "┏"
linemode(0, 1, 1, 0) = "┓"
linemode(0, 1, 1, 1) = "┳"
linemode(1, 0, 0, 0) = ""
linemode(1, 0, 0, 1) = "┗"
linemode(1, 0, 1, 0) = "┛"
linemode(1, 0, 1, 1) = "┻"
linemode(1, 1, 0, 0) = "┃"
linemode(1, 1, 0, 1) = "┣"
linemode(1, 1, 1, 0) = "┫"
linemode(1, 1, 1, 1) = "╋"

Dim top_joint, bottom_joint, left_joint, right_joint
top_joint    = Array("┃", "╋", "┏", "┓", "┣", "┳", "┫")
bottom_joint = Array("┃", "╋", "┛", "┗", "┣", "┫", "┻")
left_joint   = Array("━", "╋", "┏", "┗", "┣", "┳", "┻")
right_joint  = Array("━", "╋", "┛", "┓", "┳", "┫", "┫")

Call DrawLine("Bottom")

Sub DrawLine(direct)
	Dim ln
	Dim defchar
	Select Case direct
	Case "Bottom": ln = CStr(linemode( _
			IsStrMatch(GetTop,    top_joint), _
			1, _
			IsStrMatch(GetLeft,   left_joint), _
			IsStrMatch(GetRight,  right_joint))) ' 下方向強制連結
			defchar = "┃"
	Case "Left": ln = CStr(linemode( _
			IsStrMatch(GetTop,    top_joint), _
			IsStrMatch(GetBottom, bottom_joint), _
			1, _
			IsStrMatch(GetRight,  right_joint))) ' 左方向強制連結
			defchar = "━"
	Case "Right": ln = CStr(linemode( _
			IsStrMatch(GetTop,    top_joint), _
			IsStrMatch(GetBottom, bottom_joint), _
			IsStrMatch(GetLeft,   left_joint), _
			1)) ' 右方向強制連結
			defchar = "━"
	Case "Top": ln = CStr(linemode( _
			1, _
			IsStrMatch(GetBottom, bottom_joint), _
			IsStrMatch(GetLeft,   left_joint), _
			IsStrMatch(GetRight,  right_joint))) ' 上方向強制連結
			defchar = "┃"
	Case Else: ln = CStr(linemode( _
			IsStrMatch(GetTop,    top_joint), _
			1, _
			IsStrMatch(GetLeft,   left_joint), _
			IsStrMatch(GetRight,  right_joint))) ' 下方向強制連結
			defchar = "┃"
	End Select
	
	If ln = "" Then ln = defchar
	
	Call InsertText(ln)

	Select Case direct
	Case "Bottom": If Not MoveBottom Then Exit Sub
	Case "Left": If Not MoveLeft Then Exit Sub
	Case "Right": If Not MoveRight Then Exit Sub
	Case "Top": If Not MoveTop Then Exit Sub
	Case Else: If Not MoveBottom Then Exit Sub
	End Select

	Call InsertText( _
		CStr(linemode( _
			IsStrMatch(GetTop,    top_joint), _
			IsStrMatch(GetBottom, bottom_joint), _
			IsStrMatch(GetLeft,   left_joint), _
			IsStrMatch(GetRight,  right_joint))))
	'To exit selecting mode	
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

Function GetTop
	GetTop = GetTopOrBottom("Top")
End Function

Function GetBottom
	GetBottom = GetTopOrBottom("Bottom")
End Function

Function GetTopOrBottom(direct)
	GetTopOrBottom = ""
	
	Dim x, y
	x = CLng(ExpandParameter("$x"))
	y = CLng(ExpandParameter("$y"))
	
	Select Case direct
	Case "Top": If y = 1 Then Exit Function
	Case "Bottom": If IsFinalLine(y) Then Exit Function
	Case Else: If y = 1 Then Exit Function
	End Select
	
	Dim s
	Select Case direct
	Case "Top": s = GetLineStr(CLng(y - 1))
	Case "Bottom": s = GetLineStr(CLng(y + 1))
	Case Else: s = GetLineStr(CLng(y - 1))
	End Select

	Dim sostr, sobln
	sostr = Mid(GetLineStr(CLng(0)), 1, x)
	sobln = ByteLen(sostr)

	Dim ln,bln
	ln = Len(s)
	bln = ByteLen(s)

	If bln < sobln Then Exit Function
	GetTopOrBottom = ByteMid(s, sobln, 1)
End Function

Function GetLeft()
	GetLeft = ""
	
	Dim x, s
	x = CLng(ExpandParameter("$x"))
	s = GetLineStr(CLng(0))
	
	If CLng(x) = 1 Then Exit Function
	
	GetLeft = Mid(s, (x - 1), 1)
End Function

Function GetRight()
	GetRight = ""
	
	Dim x, s, ln
	x  = CLng(ExpandParameter("$x"))
	s  = GetLineStr(CLng(0))
	ln = Len(s)
	
	If CLng(x) - 1 >= ln Then Exit Function
	If CLng(x + 1) - 1 >= ln Then Exit Function
	
	GetRight = Mid(s, (x + 1), 1)
End Function

' いずれかにマッチしていれば 1, 非マッチなら 0 を返す
Function IsStrMatch(s, arr)
	IsStrMatch = 0
	
	If Not IsArray(arr) Then Exit Function
	
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
	
	Dim ismulti
	ismulti = False
	
	Dim isrep
	If IsStrMatch(GetCur, Array(" ", "━", "┃", "╋", "┏", "┓", "┛", "┗", "┣", "┳", "┫", "┻")) = 1 Then
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
	
	Dim strToCur
	strToCur = Mid(GetLineStr(CLng(y)), 1, x)
	Dim byteLenCur
	byteLenCur = ByteLen(strToCur)
	Dim charByte
	charByte = 2

	Dim sp, s
	sp = 0
	s = GetLineStr(CLng(y - 1))
	If ByteLen(s) < byteLenCur Then sp = byteLenCur - ByteLen(s) - charByte
	Editor.Up
	MoveTop = True
	
	If sp = 0 Then Exit Function
	
	Dim i, spcs
	spcs = ""
	For i = 1 To sp
		spcs = spcs & " "
	Next
	Call InsText(CStr(spcs))
End Function

Function MoveBottom()
	MoveBottom = True
	
	Dim x, y
	x = CLng(ExpandParameter("$x"))
	y = CLng(ExpandParameter("$y"))

	Dim strToCur
	strToCur = Mid(GetLineStr(CLng(y)), 1, x)
	Dim byteLenCur
	byteLenCur = ByteLen(strToCur)
	Dim charByte
	charByte = 2
	
	Dim sp
	If IsFinalLine(y) Then
		GoLineEnd
		InsertCR
		Dim xNew, yNew
		xNew = CLng(ExpandParameter("$x"))
		yNew = CLng(ExpandParameter("$y"))
		Dim strToCurNew
		strToCurNew =  Mid(GetLineStr(CLng(yNew)), 1, xNew)
		sp = byteLenCur - ByteLen(strToCurNew) - charByte
	Else
		Dim s
		s = GetLineStr(CLng(y + 1))
		If ByteLen(s) < byteLenCur Then sp = byteLenCur - ByteLen(s) - charByte
		Editor.Down
	End If
	
	If sp = 0 Then Exit Function
	
	Dim i, spcs
	spcs = ""
	For i = 1 To sp
		spcs = spcs & " "
	Next
	Call InsText(CStr(spcs))
End Function

Function MoveLeft()
	MoveLeft = False
	
	If CLng(ExpandParameter("$x")) = 1 Then Exit Function
	Editor.Left
	
	If GetCur = " " Then
		If CLng(ExpandParameter("$x")) > 1 Then
			Editor.Left
			If GetCur <> " " Then Editor.Right
		End If
	End If
	
	MoveLeft = True
End Function

Function MoveRight()
	MoveRight = True
	
	Dim cur
	cur = GetCur
	If cur = "" Then Exit Function
	Editor.Right
	
	If cur = " " And GetCur = " " Then Editor.Right
End Function

Function ByteLen(ByVal s)
	ByteLen = 0
	Dim i
	For i = 1 to Len(s)
		Dim c
		c = Mid(s, i, 1)
		If c = vbCr Or c = vbLf Then Exit For
		ByteLen = ByteLen + ByteSize(c)
	Next
End Function

Function ByteSize(ByVal c)
	ByteSize = 0
	If Len(c) = 0 Then Exit Function
	
	'半角文字は1バイトとして扱う
	If (Asc(c) >= 1) And (Asc(c) <= 255) Then
		ByteSize = 1
	Else
		ByteSize = 2
    	End If
End Function

' 文字列sの indexバイト目から length文字取得
Function ByteMid(ByVal s, ByVal index, ByVal length)
	ByteMid = ""
	
	Dim i, bidx
	bidx = 0
	For i = 1 To Len(s)
		Dim c
		c    = Mid(s, i, 1)
		bidx = bidx + ByteSize(c)
		If bidx >= index Then
			If length > 0 Then ByteMid = ByteMid & c
			length = length - 1
			If length = 0 Then Exit For
		End If
	Next
End Function

Function IsFinalLine(lineno)
	IsFinalLine = False
	If lineno = GetLineCount(CLng(0)) Then IsFinalLine = True
End Function

Function InsertCR()
	Call Char(CLng(13))
End Function
