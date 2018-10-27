Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Declare Function CreateWindowExW Lib "@user32" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function LoadLibrary Lib "@kernel32"  Alias "LoadLibraryW" (ByVal lpLibFileName As String) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SCI_Send Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wp As Any, ByVal lp As Any) As Long

<#Module=mdlSci>

Const GWL_STYLE = (-16)
Const WS_EX_CLIENTEDGE = &H200
Const WS_CHILD = &H40000000
Const WS_VISIBLE = &H10000000
Const WM_NOTIFY = &H4E
Const SWP_NOZORDER = &H4
Const SW_SHOW = 5

Const SCI_GETTEXT = 2182
Const SCI_SETTEXT = 2181
Const SCI_GETLENGTH = 2006
Const SCI_GETCURLINE = 2027
Const SCI_STYLECLEARALL = 2050
Const SCI_STYLERESETDEFAULT = 2058
Const SCI_SETSELBACK = 2068
Const SCI_SETKEYWORDS = 4005
Const SCI_SETLEXER = 4001
Const SCI_SETPROPERTY = 4004
Const SCI_SETWORDCHARS = 2077
Const SCI_LINEFROMPOSITION = 2166
Const SCI_TOGGLEFOLD = 2231
Const SCI_GETCURRENTPOS = 2008
Const SCI_GETCHARAT = 2007
Const SCI_WORDSTARTPOSITION = 2266
Const SCI_WORDENDPOSITION = 2267
Const SCI_GETTEXTRANGE = 2162 
Const SCI_AUTOCSHOW = 2100
Const SCI_COLOURISE = 4003
Const SCI_GETLINECOUNT = 2154
Const SCI_GETFOLDLEVEL = 2223
Const SCI_GETFOLDEXPANDED = 2230      '
Const SCI_STYLESETFORE = 2051
Const SCI_STYLESETBACK = 2052
Const SCI_STYLESETBOLD = 2053
Const SCI_STYLESETITALIC = 2054
Const SCI_STYLESETSIZE = 2055
Const SCI_STYLESETFONT = 2056
Const SCI_STYLESETEOLFILLED = 2057
Const SCI_STYLESETUNDERLINE = 2059
Const SCI_STYLESETCASE = 2060
Const SCI_STYLESETCHARACTERSET = 2066
Const SCI_STYLESETVISIBLE = 2074
Const SCI_FOLDLEVELNUMBERMASK = &HFFF
Const SCI_FOLDLEVELHEADERFLAG = &H2000
Const SCI_CP_UTF8 = 65001
Const SCI_CP_DBCS = 1
Const SCI_SETCODEPAGE = 2037
Const SCI_FOLDLEVELBASE = &H400
Const SCN_MARGINCLICK = 2010
Const SCN_CHARADDED = 2001

Type CharacterRange
    cpMin As Long
    cpMax As Long
End Type

Type TextRange
    CharRange As CharacterRange
    lpstrText As String
End Type

Type NMHDR
    hwndFrom As Long
    idFrom As Long
    code As Long
End Type

Type SCNotification
    NotifyHeader As NMHDR
    position As Long
    ch As Long
    modifiers As Long
    modificationType As Long
    Text As Long
    Length As Long
    linesAdded As Long
    Message As Long
    wParam As Long
    lParam As Long
    line As Long
    foldLevelNow As Long
    foldLevelPrev As Long
    margin As Long
    listType As Long
    X As Long
    Y As Long
	token As Long
	annotationLinesAdded As Long
	updated As Long
End Type


'-----------------------------------------
Sub SCI_SetStyle(hWnd, sty, param)
	sys.prm param, vbBlack, vbWhite, "Courier New", 10, False, False, False, False, True, 0, 204

	if param(0) <> -1 then SCI_Send hWnd, SCI_STYLESETFORE, sty, param(0)
	if param(1) <> -1 then SCI_Send hWnd, SCI_STYLESETBACK, sty, param(1)

	SCI_Send hWnd, SCI_STYLESETFONT, sty, param(2)	
	SCI_Send hWnd, SCI_STYLESETSIZE, sty, param(3)
	SCI_Send hWnd, SCI_STYLESETBOLD, sty, param(4)
	SCI_Send hWnd, SCI_STYLESETITALIC, sty, param(5)
	SCI_Send hWnd, SCI_STYLESETUNDERLINE, sty, param(6)
	SCI_Send hWnd, SCI_STYLESETEOLFILLED, sty, param(7)
	SCI_Send hWnd, SCI_STYLESETVISIBLE, sty, param(8)
	SCI_Send hWnd, SCI_STYLESETCASE, sty, CLng(param(9))
	SCI_Send hWnd, SCI_STYLESETCHARACTERSET, sty, CLng(param(10))
End Sub



Class Scintilla
	Private Notif, Range, hLib, Parent, SubClass, lTop, lLeft, lWidth, lHeight
	Public hWnd, Name
	
	'---------------------------------------
	Private Sub Class_Initialize()
		Dim Notif As New SCNotification
		Dim Range As New TextRange
		Set SubClass = Sys.NewSubClass
		lLeft = 0 : lTop = 0 : lWidth = 100 : lHeight = 100
	End Sub
    
	'---------------------------------------
	Private Sub Class_Terminate()
		If hWnd Then
			SubClass.HookClear Me, Parent.hWnd
			DestroyWindow hWnd
		End If
		If hLib Then FreeLibrary hLib
		Set SubClass = Nothing
		Set Parent = Nothing
	End Sub
	
	'---------------------------------------
	Sub Class_New(frm, m_name)
		'On Error Resume Next
		
		Set Parent = frm
		Name = m_name
        
        frm.caption = "Scintilla - " + sys.ext.versiondll("SciLexer.DLL")
        
		hLib = LoadLibrary("SciLexer.dll")
		if hLib Then hWnd = CreateWindowExW(WS_EX_CLIENTEDGE, "Scintilla", m_name, WS_CHILD Or WS_VISIBLE, lLeft, lTop, lWidth, lHeight, Parent.hWnd, 0, sys.App.hInstance, 0)

		If hWnd Then
			SubClass.HookSet Me, Parent.hWnd	
			SubClass.Msg(WM_NOTIFY) = 1
		End If
	End Sub
	
	'---------------------------------------
	Property Let Left(value): lLeft = value: Resize: End Property
	Property Get Left(): Left = lLeft: End Property
	Property Let Top(value): lTop = value: Resize: End Property
	Property Get Top(): Top = lTop: End Property
	Property Let Width(value): lWidth = value: Resize: End Property
	Property Get Width(): Width = lWidth: End Property
	Property Let Height(value): lHeight = value: Resize: End Property
	Property Get Height(): Height = lHeight: End Property
	Property Get Visible(): Visible = WindowLong(hWnd, WS_VISIBLE): End Property
	Property Let Visible(value): Call WindowLong(hWnd, WS_VISIBLE, value): End Property
	
	'---------------------------------------
	Private Sub Refresh()
		ShowWindow hWnd, SW_SHOW
	End Sub

	'---------------------------------------
	Private Sub Resize()
		If hWnd <> 0 Then Call SetWindowPos(hWnd, 0, lLeft, lTop, lWidth, lHeight, SWP_NOZORDER)
	End Sub

	'---------------------------------------
	Sub Move(vLeft, vTop, vWidth, vHeight)
		If IsNumeric(vTop) Then lTop = vTop
		If IsNumeric(vLeft) Then lLeft = vLeft
		If IsNumeric(vWidth) Then lWidth = vWidth
		If IsNumeric(vHeight) Then lHeight = vHeight
		Resize
	End Sub

	'---------------------------------------
	Property Get Text()
		Dim lenTxt, txt
		lenTxt = SCI_Send(hWnd, SCI_GETLENGTH, 0, 0) + 1
		txt = String(lenTxt,0)
		SCI_Send hWnd, SCI_GETTEXT, lenTxt, StrPtr(txt)
		txt = strconv(txt, vbunicode)
		Text = sys.str.Left(txt, lenTxt - 1)
	End Property

	'---------------------------------------
	Property Let Text(txt)
        'SCI_Send hWnd, SCI_SETCODEPAGE, SCI_CP_UTF8, 0
		SCI_Send hWnd, SCI_SETTEXT, 0, txt
	End Property

	'---------------------------------------
	Property Let Syntax(value)
		Dim bMulti, old, v, t
		
		SCI_Send hWnd, SCI_STYLERESETDEFAULT, 0, 0
		SCI_Send hWnd, SCI_STYLECLEARALL, 0, 0
		
		With Sys.Host.Script
			For Each t In Split(value, vbCrLf)
				For Each v In Split(t, ";")
					v = LTrim(v)
					bMulti = False
					Select Case sys.str.Left(v, 1)
						Case "&"	:	bMulti = True	:	v = Mid(v, 2)
						Case "~"	:	bMulti = True	:	v = Mid(v, 2) + vbCrLf
						Case "$"	:	v = "SCI_SetStyle " & hWnd & ", " & Mid(v, 2)
						Case "@"	:	v = "SCI_Send " & hWnd & ", " & Mid(v, 2)
					End Select
					If bMulti = False Then .Execute old : old = ""
					old = old + v
				Next
			Next
			.Execute old
		End With
	End Property
	
	'-----------------------------------------
	Property Let Expand(value)
		Dim maxLine, level, eStat
		
		value = abs(value)
		
		SCI_Send hWnd, SCI_COLOURISE, 0, -1
		
		maxLine = SCI_Send(hWnd, SCI_GETLINECOUNT, 0, 0)
		If maxLine = 0 Then Exit Property
		
		For a = 1 To maxLine - 1
			level = SCI_Send(hWnd, SCI_GETFOLDLEVEL, a, 0)
			If (level Xor SCI_FOLDLEVELBASE Xor SCI_FOLDLEVELHEADERFLAG) < SCI_FOLDLEVELNUMBERMASK Then
				eStat = SCI_Send(hWnd, SCI_GETFOLDEXPANDED, a, 0)
				If eStat <> value Then SCI_Send hWnd, SCI_TOGGLEFOLD, a, 0
			End If
		Next
	End Property

	'-----------------------------------------
	Function WindowProc(bHandled, uhwnd, umsg, wparam, lparam, dwRefData)
		Dim nLine, sym, pos, lenTxt, curWord, txt

		Select Case umsg
			Case WM_NOTIFY
				CopyMemory Notif, lParam, LenType(Notif)
				
				if Notif.NotifyHeader.hwndFrom <> hWnd Then exit function

				Select Case Notif.NotifyHeader.code
					Case SCN_MARGINCLICK
						nLine = SCI_Send(hWnd, SCI_LINEFROMPOSITION, Notif.position, 0)
						SCI_Send hWnd, SCI_TOGGLEFOLD, nLine, 0
						
					Case SCN_CHARADDED
						pos = SCI_Send(hWnd, SCI_GETCURRENTPOS, 0, 0) - 1
						sym = SCI_Send(hWnd, SCI_GETCHARAT, pos, 0)
						If sym < 0 Then sym = 256 + sym
						
						select case sym
						  case 46
							Range.CharRange.cpMin = SCI_Send(hWnd, SCI_WORDSTARTPOSITION, pos, 1)
							Range.CharRange.cpMax = SCI_Send(hWnd, SCI_WORDENDPOSITION, pos, 1)
							Range.lpstrText = String(256, 0)
							lenTxt = SCI_Send(hWnd, SCI_GETTEXTRANGE, 0, Range)
							curWord = StrConv(Range.lpstrText, vbUnicode)
							curWord = sys.str.Left(curWord, lenTxt)
							
							Select Case LCase(curWord)
								Case "form."
									SCI_Send hWnd, SCI_AUTOCSHOW, 0, "Show Visible"
							End Select                
						end select
				End Select
		End Select
	End Function
End Class
<#Module>
