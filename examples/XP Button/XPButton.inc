Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

<#Module=mdlXPButton>

Type POINT_API
    x As Long
    Y As Long
End Type


Class XPButton
	Private XPB_Form, Clip, Pos, IsLoad
	Public Caption
	
	'---------------------------------------
	Private Sub Class_Initialize()
		Dim Pos as New POINT_API
	End Sub

	'---------------------------------------
	Private Sub Class_Terminate()
		Unload XPB_Form
	End Sub
	
	'---------------------------------------
	Sub Class_New(frm, m_name)
		'On Error Resume Next
		
		Set Clip = sys.gdi.clip.createnew
		Clip.picture = ResToPic("XPButton")
		Clip.cols = 5		
		
		Set XPB_Form = Sys.NewForm
		
		With XPB_Form
			Set .codeobject = Me
			
			Set .parent = frm
			.child = true

			.caption = m_name
			.backcolor = qbcolor(7)
			.style.titlebar = false
			.autoredraw = true
			.fontsize = 10
			.fontbold = true
			
			.add "timer",1
			.timer(1).interval = 10
			.timer(1).enabled = false
			
			.show
		End With
		
		Caption = "OK"
	End Sub
	
	'---------------------------------------
	Sub Move(left, top, width, height)
	  XPB_Form.move left, top, width, height
	End Sub
	
	'---------------------------------------
	Sub statevalue_pic()
	    make_xpbutton 0
	End Sub
	
	'---------------------------------------
	Sub make_xpbutton(z)
	    Dim brx, bry, bw, bh, txt
	
		with XPB_form
		
		.scalemode = 3
		
	    brx = .ScaleWidth - 3 'right x
	    bry = .ScaleHeight - 3 'right y
	    bw = .ScaleWidth - 6 'border width - corners width
	    bh = .ScaleHeight - 6 'border height - corners height
		
	    .PaintPicture Clip.GraphicCell(z), 0, 0, 3, 3, 0, 0, 3, 3
	    .PaintPicture Clip.GraphicCell(z), brx, 0, 3, 3, 15, 0, 3, 3
	    .PaintPicture Clip.GraphicCell(z), brx, bry, 3, 3, 15, 18, 3, 3
	    .PaintPicture Clip.GraphicCell(z), 0, bry, 3, 3, 0, 18, 3, 3
	    .PaintPicture Clip.GraphicCell(z), 3, 0, bw, 3, 3, 0, 12, 3
	    .PaintPicture Clip.GraphicCell(z), brx, 3, 3, bh, 15, 3, 3, 15
	    .PaintPicture Clip.GraphicCell(z), 0, 3, 3, bh, 0, 3, 3, 15
	    .PaintPicture Clip.GraphicCell(z), 3, bry, bw, 3, 3, 18, 12, 3
	    .PaintPicture Clip.GraphicCell(z), 3, 3, bw, bh, 3, 3, 12, 15
		.scalemode = 1
	
	    GPrint XPB_form, Caption, .scalewidth / 2 - .textwidth(Caption) / 2, .scaleheight / 2 - .textheight(Caption) / 2
	    
		end with
	End Sub
	
	Sub form_MouseDown(Button, Shift, x, Y)
	    make_xpbutton 1
	    XPB_form.parent.SetFocus
	    XPB_form.parent.Events XPB_form.caption + "_click"
	End Sub
	
	Sub form_MouseMove(Button, Shift, x, Y)
	    XPB_form.Timer(1).Enabled = True
	    If x >= 0 And Y >= 0 And x <= XPB_form.ScaleWidth And Y <= XPB_form.ScaleHeight Then
	       if button=0 then make_xpbutton 3 else make_xpbutton 1
	    End If
	End Sub
	
	Sub form_MouseUp(Button, Shift, x, Y)
	    statevalue_pic
	End Sub
	
	Sub Form_Activate()
		if IsLoad = false then statevalue_pic
		IsLoad = True
	End Sub
	
	Sub Timer1_Timer()
	    GetCursorPos pos
		
		With XPB_form
			ScreenToClient .hwnd, pos
			
			.ScaleMode = 3
		
			If pos.x < .ScaleLeft Or pos.Y < .ScaleTop Or _
			   (pos.x > (.ScaleLeft + .ScaleWidth)) Or _
			   (pos.Y > (.ScaleTop + .ScaleHeight)) Then
		
			   .Timer(1).Enabled = False
			   statevalue_pic
			End If
			
			.ScaleMode = 1
		End With
	End Sub
End Class

<#Module>

<#res id="XPButton" mode=base64 #>
R0lGODdhWgAVAPcAAAA8dBVMfeWXAPiyL/iyMPizMPi0Nfm0Nfi1Nfm3Ovi3O/m3O/m4O4OEb4SG
ce21R++1Sfm6Qfq6Qfm6Qvm9SPm9Sfq9Sfq+SfrAUPrBUPrAUfrDWPrEWPrEWfrHYfvHYfvIYPvK
avzKafvPcvzOcfzOcvzSefzRevzSeidNhSJVg0ZtjkhukElwkkpxk011mE93m1Z7mVd8mVV9ok19
wV6CnmWGn2GEoGWGoGqLo2mC7o2eoZ2orKCrsaOutKm1vbG3trG3t7a9vf3Vgf3Vgv3Yif3bj/7d
lf7fmv3gm8nHutDLv+7crYmt5Iut5Iuv5Y2v5Y6u5Y6v5pGx5pGx55Sz6JS06JW06Ji46Zi46pm4
6py7652866C+7KG+7bzFx7LBy7/Hyr/O3KXB7aTC7aXD7qnF76jG8K3H8K3I8LHL8rHM8rXO87TO
87fQ9LjR9brT9brT9rzU9tHMwdHMwNXPw9XQxNbQxdbQxt3XxdfTytbUydfUytfVytfUy9jUy9jW
ytnWzdjWzNvXzdvXztrXz9rYz9zYz8TO0sDP3crT2tva0NzZ0Nzb0d3a0dza0t3b0t3b09/c0t7c
09/d1d7d1N/e1t/f1uHcyuDe1OHe1eDe1+Df1+Lf1v/wz+Dg1+/o0uDg2OHg2OHh2OHg2eLh2eLg
2OLi2eHg2uLh2uLh2+Li2uPi2uPi2+Pj2+Ti2uXi2+Tj2+Pj3OTj3OTj3eXj3Obj3eTk3eXk3eXl
3efk3eXk3uXl3ubl3+bm3+jl3+zq2erp3vjt09Xf6M7n/+bl4Obm4Obm4enm4Orn4Ono4ero4erp
4+vo4u7t4urq5Ovq5Ovr5ezr5Ozr5ezr5u3q5u3s5e3s5+7t5+/u6O/v6fDv6fDu6vDw6vDw6/Hw
6/X06vDw7PHx7PHx7fLy7fLx7vPx7vLy7vPz7vLy7/Pz7/T07/T08PX18fX18vb08vb28vb28/b2
9Pf39fn48vj49vj49/r6+Pr6+fv7+vz8+/z8/P39/P7+/f7+/v///ywAAAAAWgAVAAAI/wDz2FAB
oKDBgwUDyMAkkCDChwox/VNCsaLFi0r+NXwIceFGjgcj2piR6J/JkyhNDpuBw8YKHkvozJlJs2aQ
FjiUAAoGrqfPnz2dAVJiYwcoT0iTKkUqbEfLFMSiSp06NQUOFWL8ad3KdWs/MSpU9MDzh9Chs2jR
EmL0Q4USnkDj9gymRAUTE0Xy6t2b1wSTsHICCx5MWA5Bfvz26VvMuLG+fO8K/nFkahauy5gv15pl
GYASuaDBKQGAxIRix45NJCkIJ85p1IxbF8RHu7bt2vfccStoaNMtX8aCCx9u7Jjn0HJHHzFxu/k9
E0cKunnT/Pa9N24K2qvHvbt37vPMVf8rWGjTrF3o06vf1avYceRARxsx8b3+PBNGCrZhU//7vP0F
ySNPPAQWaGA87nzTTEGBUNIKZxBGOAstu/DyHnw+jVaECQd26A5eBamhRocHuqPGGgURCM+KLLYI
TzvdLMOggxJKSKGFn2GYIQBEmODij+2YMERBaaDxo4vtoIFGQfC4086TUEbZDjvbJFOQIJWs4sqW
XHYpyy25XKjjaCigIOWZ7JhwQkFmnHGmlOyYYUZBT7Jj5514sqNONshINkkprAQq6KAP1iImhqOR
QEKejKozQgkFjUEGo3mqU0YZBdmpzqacdqrOOdj8UpAfkZCSyqmopqrKg4fCN5oIIXj/Kus5IYhQ
UBddyOrpOV50UdA666Aj7LDEokNOqAXxAYkopTTr7LOptBJLq8iN9gEIxWZLjgcfFLQFF9kWS84W
WxQkLDnopqsuOeJYo0uyyz4rbynRTpujjqIBsEEH6/YrDgcdFJQFFv2uK44WWBR0zjnjNOzww+OE
M40tBenxCCekZKzxxqisSm1oo2GQAcQkh6NBBgVZUQXJEIdzxRUFoSvOzDTXLI430hgKgB6NcCLK
z0AHbSorH4M2WgUX2Ky0NxZQUBAVUyhtszdQFzSzN1hnrbU32kADS0F7LHLJJ2SXbXYop6xSdHIA
SDDB1nBrE0EEBUkRBdxbayMFFAV5/9PN34AH/nc2y7xSUB+NjG324qNouXZcozGwgOCUZ6NAAgU5
4QTlgmfjxBMFURNNNNNQY/rppkejzCuZ8GZJKqysIvvssj/YSy+PxwfAAQiQjjrqqiNgQEFNNOH7
76krUzwAAYTRSSaadCL99NJrIskgQgQQwBe+MPPMM8yEL/4z0FBzDSIBKOEMvs4oEQAEBBQwQAEF
xE9//QTE/4D2xffv///FC0AMXOCDOxjwgAi0Qx2AwIIaxOAFYAhHOSZIQQqeIx2KgEENdAIX5AhF
CTFwgABGSMISlrABDkwBDXTAwha6sIUpqAEwcuAQkCTkBvSYYQ1tGAAcTgQjQKTIPw50aEOD9DCH
NCziDekREAA7
<#res#>
