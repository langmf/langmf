<#--debug="0">

Declare Function GetModuleFileNameExW Lib "psapi" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As Long, ByVal nSize As Long) As Long
Declare Function GetCurrentProcess Lib "kernel32" () As Long


<#Module = modDebug>
Type RTC_BUFFER
    Data() As Byte
End Type

Type RTC_CONTEXT
    Prog        As Long
    Syms        As Long
    Mods        As Long
    Glob        As Long
    Vars        As Long
    State(31)   As Long
    Buffer(3)   As RTC_BUFFER
End Type

'---------------------------------------
Function Debug(obj)
	Dim ctx as RTC_CONTEXT(obj.ptr)
	Dim c(8) as RTC_CONTEXT(obj.ptr)
	Dim b
	
	sys.conv.str2file join(mods(ctx), vbcrlf), "dll.txt"
	
	b = ctx.buffer(0).data   :   sys.conv.buf2file b, "syms.txt"   :   vtype(b) = 0
	b = ctx.buffer(1).data   :   sys.conv.buf2file b, "glob.txt"   :   vtype(b) = 0
	b = ctx.buffer(2).data   :   sys.conv.buf2file b, "vars.txt"   :   vtype(b) = 0
	b = ctx.buffer(3).data   :   sys.conv.buf2file b, "mods.txt"   :   vtype(b) = 0
	
	'b = ctx.buffer(0).data   :   msgbox typename(b) & vbcrlf & byrefvar(b, 0),,ubound(b)   :   vtype(b) = 0
End Function

'---------------------------------------
Function Mods(ctx)
    Dim a, b, i, m, sz, rc, p, txt, v()

	b = ctx.Buffer(3).Data   :   sz = UBound(b) + 1   :   vtype(b) = 0   :   sz = sz / 4 - 1   :   i = -1   :   If sz <= 0 Then Exit Function
    
    ReDim v(sz):    p = GetCurrentProcess():	txt = String(1000, 0)
    
    For a = 0 To sz
        m = sys.var.MemLong(ctx.Mods + a * 4):
		If m Then rc = GetModuleFileNameExW(p, m, StrPtr(txt), Len(txt)):     i = i + 1:     v(i) = Format(hex(m), "00000000") & vbTab & Left(txt, rc)
    Next
    
    ReDim Preserve v(i):       Mods = v
End Function
<#Module>