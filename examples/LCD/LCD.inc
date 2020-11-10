<#Module=mdlLCD>
Class LCD
    Private m_Formats, m_Value, m_Minus, m_Text, m_Fon
    Public Obj, Style, Styles

    '---------------------------------------
    Public Sub Refresh()
        Dim c, cx, cy

        If Obj Is Nothing Then Exit Sub
        
        Obj.FontSize = Obj.Height / 20
        
        cx = Obj.TextWidth("8") * 0.3
        cy = Obj.ScaleHeight / 2 - Obj.TextHeight("0") / 2 + Obj.ScaleHeight * 0.025
        br = ((Obj.Width - Obj.ScaleWidth) \ vbPY) * vbPY
        
        Obj.Width = Obj.TextWidth(m_Minus & m_Formats) + (cx + br) * 2

        c = GetStyle(Style)
        
        Obj.Cls
        Obj.BackColor = c(0)
        Obj.ForeColor = c(1)  :  Obj.CurrentX = cx  :  Obj.CurrentY = cy  :  GPrint Obj, m_Fon
        Obj.ForeColor = c(2)  :  Obj.CurrentX = cx  :  Obj.CurrentY = cy  :  GPrint Obj, m_Text
    End Sub
    
        '---------------------------------------
    Public Property Get Minus()
        Minus = m_Minus
    End Property

    '---------------------------------------
    Public Property Let Minus(newValue)
        m_Minus = newValue
        Formats = m_Formats
        Value = m_Value
    End Property
    
    '---------------------------------------
    Public Property Get Formats()
        Formats = m_Formats
    End Property

    '---------------------------------------
    Public Property Let Formats(ByVal newValue)
        Dim a, sym
        
        m_Formats = newValue
        
        On Error Resume Next

        For a = 1 To Len(newValue)
            sym = Mid(newValue, a, 1)
            If sym <> "." Then sys.str.Mid(newValue, a) = "8"
        Next

        m_Fon = m_Minus & newValue
        Value = m_Value
    End Property

    '---------------------------------------
    Public Property Get Value()
        Value = m_Value
    End Property

    '---------------------------------------
    Public Property Let Value(ByVal newValue)
        Dim a, sym, rc, tmp, sz
        
        On Error Resume Next
        
        m_Value = newValue
        
        tmp = Format(newValue, m_Formats)
        newValue = CDbl(tmp)
        tmp = Replace(tmp, ",", ".")
        If newValue >= 0 Then tmp = string(Len(m_Minus), " ") + tmp

        rc = InStr(1, tmp, ".")
        If rc = 0 Then rc = Len(tmp) + 1

        For a = 1 + Len(m_Minus) To rc - 2
            sym = Mid(tmp, a, 1)
            If sym = "0" Then sys.str.Mid(tmp, a) = " " Else Exit For
        Next

        rc = InStr(1, m_Formats, ".") - 1
        If rc < 0 Then rc = Len(m_Formats)
        sz = Len(CStr(Fix(newValue)))
        If (newValue >= 0 And sz > rc) Or (newValue < 0 And sz > (rc + 1)) Or (newValue < 0 And Len(m_Minus) = 0) Then
            tmp = m_Minus & m_Formats
            For a = 1 + Len(m_Minus) To Len(tmp)
                sym = Mid(tmp, a, 1)
                If sym <> "." Then sys.str.Mid(tmp, a) = "-"
            Next
        End If

        m_Text = tmp

        Refresh
    End Property

    '---------------------------------------
    Private Function GetStyle(nameStyle)
        On Error Resume Next
        GetStyle = Styles(nameStyle)
        If err Then GetStyle = Array(&HC3D2C7, &HBFCDC0, 0)         'Default
    End Function

    '---------------------------------------
    Private Sub Class_Initialize()
        m_Minus = "-"  :  Formats = "00.00"  :  Value = 0
        Set Styles = Sys.NewCol
        With Styles
            .Add Array(0, RGB(0, 40, 0), RGB(0, 210, 0)), "Green"
            .Add Array(0, RGB(50, 0, 0), RGB(210, 0, 0)), "Red"
            .Add Array(RGB(62, 162, 255), RGB(54, 154, 255), 0), "Blue"
        End With
    End Sub

    '---------------------------------------
    Public Sub Class_New(pic)
		Set Obj = pic
	End Sub
End Class
<#Module>
