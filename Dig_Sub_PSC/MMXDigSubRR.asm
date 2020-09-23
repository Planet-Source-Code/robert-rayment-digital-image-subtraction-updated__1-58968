; MMXDigSubRR.asm  by  Robert Rayment 15/2/05
; NB Assumes MMX present. cpuid can be used to
; to test for MMX if wanted.

; FlatAssembler syntax

macro movab %1,%2
 {
    push dword %2
    pop dword %1
 }

format binary
Use32

PICW	  equ [ebp-4]
PICH	  equ [ebp-8]
ptrARR0   equ [ebp-12]
ptrARR1   equ [ebp-16]
ptrARRRes equ [ebp-20]
SUBMODE   equ [ebp-24]
BGL	  equ [ebp-28]
WF	  equ [ebp-32]
INVERT	  equ [ebp-36]
lo32	  equ [ebp-40]
hi32	  equ [ebp-44]

    emms
    push ebp
    mov ebp,esp
    sub esp,44	    ; RR To match lo stack value
    push edi
    push esi
    push ebx

    ; Copy structure
    mov ebx,[ebp+8]
    movab PICW,     [ebx]
    movab PICH,     [ebx+4]
    movab ptrARR0,  [ebx+8]
    movab ptrARR1,  [ebx+12]
    movab ptrARRRes,[ebx+16]
    movab SUBMODE,  [ebx+20]
    movab BGL,	    [ebx+24]
    movab WF,	    [ebx+28]
    movab INVERT,   [ebx+32]

    ; For mul by -1 or 255 - val
    xor eax,eax
    mov eax,0FFFFFFFFh
    mov lo32,eax
    mov hi32,eax
    movq mm3,hi32

    ; Place WF Weighting
    mov eax,WF
    mov edx,eax
    shl eax,16
    add eax,edx
    movd mm4,eax
    movq mm5,mm4
    punpckldq mm4,mm5	; mm4 = WF,WF,WF,WF in words
    ; Place BGL Base GreyLevel
    mov eax,BGL
    mov edx,eax
    shl eax,16
    add eax,edx
    movd mm5,eax
    movq mm6,mm5
    punpckldq mm5,mm6	; mm5 = BGL,BGL,BGL,BGL in words

    pxor mm7,mm7	; mm7 = 0

    mov esi,ptrARR0
    mov edi,ptrARR1

; RR Can use a jump table here but this is easier to follow
    mov eax,SUBMODE
    cmp eax,0
    jne Test1
    Call near MODE0
    jmp near GETOUT
Test1:
    cmp eax,1
    jne Test2
    Call near MODE1
    jmp near GETOUT
Test2:
    cmp eax,2
    jne Test3
    Call near MODE2
    jmp near GETOUT
Test3:

GETOUT:
    emms
    pop ebx
    pop esi
    pop edi
    mov esp,ebp
    pop ebp
    ret 16

;###################################

MODE0:
;        If (chkInvert.Value = 1) = False Then   ' INVERT
;             LngToRGB ARR0(ix, iy), R0, G0, B0
;             LngToRGB ARR1(ix, iy), R1, G1, B1
;        Else
;             LngToRGB ARR0(ix, iy), R1, G1, B1
;             LngToRGB ARR1(ix, iy), R0, G0, B0
;        End If

;   CulR = (BaseGreyLevel + (CLng(R0) - R1) * DiffWeight)
;   CulG = (BaseGreyLevel + (CLng(G0) - G1) * DiffWeight)
;   CulB = (BaseGreyLevel + (CLng(B0) - B1) * DiffWeight)

    mov eax,INVERT
    cmp eax,1
    jnz ModeBranch
	; Swap array pointers
	push esi
	push edi
	pop esi
	pop edi
ModeBranch:

    mov eax,ptrARRRes

    mov edx,4

    mov ebx,PICH
L0:
    mov ecx,PICW	; Num 4 byte chunks, 1 Long/Pixel at a time
L1:
    movd mm1,[esi]	; mm1 = 0 0 0 0 A B R G ARR0 ' ABGR 1 pixel in lo word
    add esi,edx 	; esi = esi + 4
    punpcklbw mm1,mm7	; mm1 = A  B  G  R  in 4 words  mm7 = 0

    movd mm2,[edi]	; mm2 = 0 0 0 0 A B R G ARR1 ' ABGR 1 pixel in lo word
    add edi,edx 	; edi = edi + 4
    punpcklbw mm2,mm7	; mm2 = A  B  G  R  ARR1 in words  mm7 = 0

    psubsw mm1,mm2	; mm1 = +/- (ARR0 - ARR1)

    pmullw mm1,mm4	; mm1 = +/- (ABGR0 - ABRG1) * DFW
    paddsw mm1,mm5	; mm1 = +/- (BGL + (ABGR0 - ABRG1) * DFW)

    packuswb mm1,mm7	; mm1 = 0 0 0 0 ABGR  in lo word. Saturates to 0 - 255   mm7 = 0

    movd [eax],mm1	; mm1 to ARRRes
    add eax,edx 	; eax = ptrARRRes = ptrARRRes + 4

    dec ecx
    jnz L1

    dec ebx
    jnz L0

RET
;-----------------------------------------------

MODE1:

;        If (chkInvert.Value = 1) = False Then
;             CulR = (BaseGreyLevel - (R0 Xor R1) * DiffWeight)
;             CulG = (BaseGreyLevel - (G0 Xor G1) * DiffWeight)
;             CulB = (BaseGreyLevel - (B0 Xor B1) * DiffWeight)
;        Else
;             CulR = (BaseGreyLevel + (R0 Xor R1) * DiffWeight)
;             CulG = (BaseGreyLevel + (G0 Xor G1) * DiffWeight)
;             CulB = (BaseGreyLevel + (B0 Xor B1) * DiffWeight)
;        End If


    mov eax,INVERT
    cmp eax,1
    jne Leavemm3
    pmullw mm3,mm3	; mm3 = mm3 x mm3  mm3 = +1
 Leavemm3:		; else mm3 = -1

    mov eax,ptrARRRes

    mov edx,4

    mov ebx,PICH
L2:
    mov ecx,PICW	; Num 4 byte chunks, 1 Long/Pixel at a time
L3:
    movd mm1,[esi]	; mm1 = 0 0 0 0 A B R G ARR0 ' ABGR 1 pixel in lo word
    add esi,edx 	; esi = esi + 4
    punpcklbw mm1,mm7	; mm1 = A  B  G  R  in 4 words  mm7 = 0

    movd mm2,[edi]	; mm2 = 0 0 0 0 A B R G ARR1 ' ABGR 1 pixel in lo word
    add edi,edx 	; edi = edi + 4
    punpcklbw mm2,mm7	; mm2 = A  B  G  R  ARR1 in words  mm7 = 0

    pxor mm1,mm2	; mm1 = +/- (ARR0 Xor ARR1)

    pmullw mm1,mm4	; mm1 = +/- (ABGR0 Xor ABRG1) * DFW

    pmullw mm1,mm3	; mm1 = mm3 x mm1  mm3 = +/-1

    paddsw mm1,mm5	; mm1 = +/- (BGL +/- (ABGR0 - ABRG1) * DFW)

    packuswb mm1,mm7	; mm1 = 0 0 0 0 ABGR  in lo word. Saturates to 0 - 255  mm7 = 0

    movd [eax],mm1	; mm1 to ARRRes
    add eax,edx 	; eax = ptrARRRes = ptrARRRes + 4

    dec ecx
    jnz L3

    dec ebx
    jnz L2

RET
;-----------------------------------------------
MODE2:




RET
;####################################
