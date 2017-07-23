
%macro movab 2
  push dword %2
  pop dword %1
%endmacro

%define ptrData	[ebp + 8]
%define Width	[ebp + 12]	
%define Height	[ebp + 16]
%define X1	[ebp + 20]	
%define Y1	[ebp + 24]
%define X2	[ebp + 28]	
%define Y2	[ebp + 32]
%define Mode	[ebp + 36]

%define X	[ebp - 4]	
%define Y	[ebp - 8]


[bits 32]

	push ebp
	mov ebp,esp
	sub esp,8
	push edi
	push esi
	push ebx

Func_0:
	cmp byte Mode,0
	jnz Func_1
	Call Gray_Image
	jmp GETOUT

Func_1:
	cmp byte Mode,1
	jnz GETOUT
	Call Red_Image

GETOUT:
	mov eax, 1
	pop ebx
	pop esi
	pop edi
	mov esp,ebp
	pop ebp
	ret 32


;############################################################
Gray_Image: 		
	mov edi, ptrData

	movab Y,Y2
.LOOPY
	movab X,X2
.LOOPX
	Call GetPoint

	mov bl, Byte[eax]
	mov Byte[eax+1], bl
	mov Byte[eax+2], bl

	dec dword X
	mov ebx, X1
	cmp dword X,ebx
	ja .LOOPX
	
	dec dword Y
	mov ebx, Y1
	cmp dword Y,ebx
	ja .LOOPY
RET

;############################################################
Red_Image: 		
	mov edi, ptrData

	movab Y,Y2
.LOOPY
	movab X,X2
.LOOPX
	Call GetPoint

	;mov Byte[eax+0], 0	;B
	;mov Byte[eax+1], 0	;G
	;mov Byte[eax+2], 0	;R
	;mov Byte[eax+3], 0	;A

	and dword [eax], 0FF0000h

	dec dword X
	mov ebx, X1
	cmp dword X,ebx
	ja .LOOPX
	
	dec dword Y
	mov ebx, Y1
	cmp dword Y,ebx
	ja .LOOPY
RET

;############################################################
GetPoint:
	;offset = edi + 4 * ((Height - y) * Width + (x-1))
	mov eax,Height
	sub eax,Y
	mul dword Width
	add eax,X
	dec eax
	shl eax,2		
	add eax,edi
RET
