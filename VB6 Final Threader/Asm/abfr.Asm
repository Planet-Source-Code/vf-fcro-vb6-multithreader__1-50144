   .586                      ; force 32 bit code
      .model flat, stdcall      ; memory model & calling convention
      option casemap :none      ; case sensitive
      .k3d
include abfr.inc



ThreadX PROTO :DWORD
Search2 PROTO :DWORD,:DWORD,:DWORD,:DWORD
wProc PROTO :DWORD,:DWORD,:DWORD,:DWORD
InitializeComm PROTO :DWORD
UnInitializeComm PROTO :DWORD
GetItem PROTO :DWORD,:DWORD,:DWORD,:DWORD
SetItem PROTO :DWORD,:DWORD,:DWORD,:DWORD
DelegateCaller PROTO :DWORD,:DWORD,:DWORD
DialogProc proto :DWORD,:DWORD,:DWORD,:DWORD
.data

ObjReplace dd InitCaller,CloseCaller,CreateNewThread,AddThread,RemoveThread,CheckThreadExist,CallThread,EnterSync,LeaveSync,WaitLocalObject,SignalLocalObject,TranslateArg,GetThreadPar,AttachCallArgs,DetachCallArgs,RemoveCallArgs,FastEvent,EnterPump,ExitPump,UpdateVM,About

slot dd 0

WindowString db "VB6MTX",0
HandleString db "HANDLE",0
MemCallString db  "CALLMEM",0
CSECTION db "CRITICAL",0
CSECTION2 db "INCRITICAL",0

PVAdr dd 0
CLAdr dd 0
SearchK db 056h,042h,035h,021h,01ch,023h,02ah
MSVBVM db "msvbvm60.dll",0
MSVBVMH dd 0

THMAIN db "ThunRTMain",0

Knl db "Kernel32.dll",0
ExitPr db "ExitProcess",0
ExitTh db "ExitThread",0

HInst dd 0

MAINOBJ dd 0

Globs db 24 dup (0)
ARSync db 24 dup (0)


Version9237 dd 084e9c085h,090fffd22h
Version8877 dd 06de9c085h,090fffd1fh
Version8176 dd 001e9c085h,090ffff22h
Version8268 dd 058e9c085h,090fffad0h
Version8495 dd 0fde9c085h,090fffd28h
Version8964 dd 0a9e9c085h,090fffd2ch

ExThread dd 0,0


.code

DllEntry proc hInstance:DWORD, reason:DWORD, reserved1:DWORD
sub esp,4
Push DWORD PTR [ExitThread+2]
pop eax
push DWORD PTR [eax]
pop ExThread

invoke GetModuleHandle,addr MSVBVM
mov MSVBVMH,eax
mov DWORD PTR [ebp-4],eax
invoke GetProcAddress,eax,addr THMAIN
mov CLAdr,eax

invoke DisableThreadLibraryCalls,hInstance

push hInstance
pop HInst

mov  eax,1
ret
DllEntry Endp


InitThreadCall proc Param1:DWORD
invoke Search2,Param1,32768,offset SearchK,7
add eax,Param1
dec eax
mov PVAdr,eax
ret
InitThreadCall endp

;Param1:Stack Size
;Param2:ThreadId

;Param3=Message
;Param4=VarArg
;Param5=ThreadPriority
CreateMT proc  StackSize:DWORD,ThreadId:DWORD,ThreadPriority:DWORD,Reason:DWORD,Message:DWORD,Args:DWORD,ThreadType:DWORD

sub esp,24
	invoke CreateEvent,0,0,0,0
	mov DWORd PTR [ebp-4],eax


	invoke TlsGetValue,0
	mov DWORD PTR [ebp-24],eax
	invoke TlsGetValue,3
	mov DWORD PTR [ebp-20],eax
	invoke TlsGetValue,4
	mov DWORD PTR [ebp-16],eax

	push Args
	push Message
	push Reason
	call DuplicateVariant

	push eax ;TransArgs
	pop DWORD PTR [ebp-8]

	push ThreadType
	pop DWORD PTR [ebp-12]
	

	lea eax,[ebp-24]	;Predaj na strukturu
	invoke CreateThread,0,StackSize,ThreadX,eax,0,ThreadId
	push eax
	invoke SetThreadPriority,eax,ThreadPriority
	invoke WaitForSingleObject,DWORD PTR [ebp-4],0ffffffffh
	invoke CloseHandle,DWORD PTR [ebp-4]
	pop eax	


	ret

CreateMT endp

ThreadX proc Param1:DWORD
assume FS: nothing
push SEHproc
push FS:[0]
mov FS:[0],esp

push SEHproc
call SetUnhandledExceptionFilter


mov esi,Param1
invoke TlsSetValue,020h,DWORD PTR  [esi+16]

push slot
call FindFree
mov ebx,eax

invoke GetCurrentThreadId
mov DWORD ptr [ebx],eax

sub esp,16
push esp
push eax
call dwtoa
mov edi,esp

invoke GlobalAlloc,040h,24
mov DWORD PTR [ebx+4],eax
mov ebx,eax

invoke CreateWindowEx,0,offset WindowString,edi,0,0,0,0,0,0,0,0,0
add esp,16

mov DWORD PTR [ebx],eax

mov edi,DWORD PTR [esi+12] ;EDI=TypeCalling
invoke SetTimer,eax,0,0,0
invoke CoInitialize,0
invoke OleInitialize,0

test edi,edi
je Nastavi
invoke TlsSetValue,0,DWORD PTR [esi]
invoke TlsSetValue,3,DWORD PTR [esi+4]
invoke TlsSetValue,4,DWORD PTR [esi+8]
invoke SetEvent,DWORD PTR [esi+20]
mov ecx,PVAdr
add ecx,02ch
Call DWORD PTR [ecx]
ret


Nastavi:
invoke SetEvent,DWORD PTR [esi+20]
push PVAdr
call CLAdr
ret
ThreadX endp

;pExcept:dword,pFrame:dword,pContext:dword,pDispatch:dword
SEHproc proc 
; mov eax,pContext
; add eax,0b8h ;'POSTAVI NA EIP Izlaz!
; mov ebx,ExitIt
; mov DWORD PTR [eax],ebx
; mov eax,0
sub esp,4
invoke GetCurrentThreadId
mov esi,eax
lea eax,[ebp-4]
push eax
push 1
push 0
push 0
push -1
push -1
push esi
push MAINOBJ
call CallThread

push esi
push MAINOBJ
call RemoveThread
invoke ExitThread,0
ret 
SEHproc endp

Search2 proc PARAM1:DWORD,PARAM2:DWORD,PARAM3:DWORD,PARAM4:DWORD
push esi
push edi
push ebx
mov ecx,PARAM1
add ecx,PARAM2
xor edx,edx ;STRING COUNTER
cld
mov esi,PARAM1
mov edi,PARAM3
LPS:
mov al,BYTE PTR [esi+edx]
cmp al,BYTE PTR [edi+edx]
jne Dalje
inc edx
jmp Dalje2
Dalje:
inc esi
xor edx,edx
Dalje2:
cmp edx,PARAM4 ;DUZINA STRINGA ISTA SA STRING COUNTEROM?
jne Dalje3
mov eax,esi
sub eax,PARAM1
inc eax
pop ebx
pop edi
pop esi
ret
Dalje3:
cmp esi,ecx
jl LPS
mov eax,-1
pop ebx
pop edi
pop esi
ret
Search2 endp

IatPatch proc ModulePatch:DWORD,ModuleImport:DWORD,ModuleFunction:DWORD,NewAddress:DWORD
sub esp,4
push esi
push edi
push ebx

invoke GetModuleHandle,ModulePatch
test eax,eax
je Kraj

mov DWORD PTR [ebp-4],eax
mov esi,eax
mov eax,DWORD PTR [esi+60]
add esi,eax ;ECX=NT HEADER

cmp WORD PTR [esi],04550h
jne Kraj
lea esi,[esi+128] ;ECX=Data Directory 1 {Import Table}
mov eax,DWORD PTR [esi]

test eax,eax
je Kraj ;Nema Import Tablu!

add eax,DWORD PTR [ebp-4]
mov esi,eax

Again:
cmp DWORD PTR [esi+16],0 ;Dali je zadnji??
je Kraj 

mov eax,DWORD PTR [esi+12]
add eax,DWORD PTR [ebp-4]

invoke lstrcmpi,eax,ModuleImport ;PRONADJI KERNEL32

push esi ;SAVE

test eax,eax
jne NijeIsti

;NASAO:
mov edi,DWORD PTR [esi]
add edi,DWORD PTR [ebp-4] ;Pocetak imena

xor esi,esi
Lps:
mov eax,DWORD PTR [edi+esi*4]
test eax,eax
je NijeIsti ;Ako je zadnji...
cmp eax,080000000h
jb Ime ;ako je ime
and eax,0FFFFh
;HINT=eax

inc esi
jmp Lps
Ime:
add eax,DWORD PTR [ebp-4]
add eax,2 ;Uzmi ime {eax}

invoke lstrcmpi,eax,ModuleFunction ;PRONADJI ExitProcess
test eax,eax
jne Nije

pop eax ;Uzmi Pointer na IMAGE IMPORT DIR
mov eax,DWORD PTR [eax+16]
lea eax,[eax+esi*4] ;EAX=IAT pozicija
add eax,DWORD PTR [ebp-4]
lea ecx,NewAddress

invoke WriteProcessMemory,-1,eax,ecx,4,0
;Gotovo!
pop ebx
pop edi
pop esi
ret

Nije:
inc esi
jmp Lps

NijeIsti:
pop esi
add esi,20
jmp Again


Gotovo:
pop ebx
pop edi
pop esi
ret

;Nije Exekutivan header
Kraj:
or eax,0ffffffffh
pop ebx
pop edi
pop esi
 	ret
IatPatch endp


InitializeComm proc Param1:DWORD
sub esp,40
push esi
push edi

mov edi,esp
xor eax,eax
mov ecx,10
cld
rep stosd ;Clear

mov DWORD PTR [ebp-40],04000h
mov DWORD PTR [ebp-4],OFFSET WindowString
mov DWORD PTR [ebp-36],WProc

push Param1
pop DWORD PTR [ebp-24]


lea eax,[ebp-40]
invoke RegisterClass,eax

pop edi
pop esi
 	ret
InitializeComm endp

UnInitializeComm proc Param1:DWORD
invoke UnregisterClass,OFFSET WindowString,Param1
ret
UnInitializeComm endp


WProc proc hWnd:DWORD,uMsg:DWORD,wParam:DWORD,lParam:DWORD
	.if uMsg==0113H
	invoke SleepEx,0,1
	xor eax,eax
	ret
	.elseif uMsg==0c000h
	push lParam
	call wParam
	ret
	.elseif uMsg==0c001h
	push lParam
	Call wParam
	xor eax,eax
	ret
	.endif
	invoke DefWindowProc,hWnd,uMsg,wParam,lParam
	ret
WProc endp


DuplicateVariant proc 
	invoke GlobalAlloc,GMEM_ZEROINIT,24
	
	push eax
	push DWORD PTR [esp+8]
	pop DWORD PTR [eax]
	push DWORD PTR [esp+12]
	pop DWORD PTR [eax+4]
	cmp DWORD PTR [esp+16],0
	je Dalje
	lea eax,[eax+8]
	invoke VariantCopy,eax,DWORD PTR [esp+16]

Dalje:	
	pop eax
	ret 12
DuplicateVariant endp

MoveVariant proc
	invoke GlobalAlloc,GMEM_ZEROINIT,24
	push eax
	push DWORD PTR [esp+8]
	pop DWORD PTR [eax]
	push DWORD PTR [esp+12]
	pop DWORD PTR [eax+4]
	cmp DWORD PTR [esp+16],0
	je Dalje

	xor ecx,ecx
	lea eax,[eax+8]
	mov edx,DWORD PTR [esp+16]
Lps:
	push DWORD PTR [edx+ecx*4]
	mov DWORD PTR [edx+ecx*4],0
	pop DWORD PTR [eax+ecx*4]
	inc ecx
	cmp ecx,4	
	jne Lps
Dalje:
	pop eax
	ret 12
MoveVariant endp



;Param1:DWORD,Param2:DWORd,Param3:DWORD,Param4:DWORD
TranslateArguments proc public uses esi edi 
	mov esi,DWORD PTR [esp+12]
	mov ecx,DWORD PTR [esp+16]
	push DWORD PTR [esi]
	pop DWORD PTR [ecx]
	mov ecx,DWORD PTR [esp+20]
	push DWORD PTR [esi+4]
	pop DWORd PTR [ecx]
	
	cmp DWORD PTR [esi+8],0
	je Van
	lea esi,[esi+8]
	mov edi,DWORD PTR [esp+24]
	cld
	mov ecx,4
	rep movsd
Van:
	invoke GlobalFree,DWORD PTR [esp+12]
	ret 16
TranslateArguments endp


Delegator proc  Param1:DWORD,Param2:DWORD,Param3:DWORD

sub esp,20

push esi
push edi
push ebx 

mov esi,Param1
mov esi,DWORD PTR [esi] ;OBJECT
mov DWORD PTR [ebp-8],esi
mov esi,DWORD PTR [esi] ;VTABLE

mov eax,Param2
lea esi,[esi+eax*4]
add esi,24 ;EAX=FIRST ADR
mov esi,DWORD PTR [esi]

mov eax,Param3
mov ecx,3
mul ecx
add eax,36 ;EAX=space

mov edi,eax
mov DWORD PTR [ebp-20],edi

invoke GlobalAlloc,0,eax



mov edi,DWORD PTR [ebp-20]

mov DWORD PTR [ebp-12],eax
add edi,eax
mov DWORD PTR [ebp-16],edi
mov DWORD PTR [edi-4],esi


mov edi,eax
mov DWORD PTR [edi],068EC8B55h

mov ecx,3
mov eax,Param3
mul ecx
lea eax,[eax+edi+28]
mov DWORD PTR [edi+4],eax
mov DWORD PTR [ebp-4],eax ;PRESERVE VIZNUM

lea edi,[edi+8]

mov ebx,Param3
shl ebx,2
add ebx,4 ;EBX=STACK P



mov eax,1
cmp Param3,0
je Van

Lps:
mov WORD PTR [edi],075FFh
add edi,2
mov BYTE PTR [edi],bl
inc edi
sub bl,4
inc eax
cmp eax,Param3
jbe Lps

Van:

mov BYTE PTR [edi],068h
push DWORD PTR [ebp-8]
pop DWORD PTR [edi+1]
mov WORD PTR [edi+5],015ffh
add edi,7

mov eax,DWORD PTR [ebp-16]
sub eax,4
mov DWORD PTR [edi],eax

mov byte PTR [edi+4],0a1h
push DWORD PTR [ebp-4]
pop DWORD PTR [edi+5]
mov WORD PTR [edi+9],0C2C9h
mov eax,Param3
shl eax,2
mov DWORD PTR [edi+11],eax

mov eax,DWORD PTR [ebp-12]
pop ebx
pop edi
pop esi


ret
Delegator endp


;Param1=OBJECT
GatherObject proc  Param1:DWORD
	
	push esi
	push edi
	push ebx
	
	mov eax,Param1
	mov eax,DWORD PTR [eax] ;CInstance
	mov eax,DWORD PTR [eax] ;VTABLE
	lea eax,[eax+28] ;FIRST FUNCTION
	

	invoke WriteProcessMemory,-1,eax,OFFSET ObjReplace,84,0
	
	pop ebx
	pop edi
	pop esi
	
	ret

GatherObject endp


InitCaller proc public uses esi edi ebx Me:DWORD,Hinstance:DWORD
mov esi,Me

mov edi,DWORD PTR [esi]
push esi
call DWORD PTR [edi+4] ;add ref

lea eax,[esi+76]
cmp WORD PTR [eax],0
Jne Van

mov MAINOBJ,esi

push eax
push ExThread
push offset ExitPr
push offset Knl
push offset MSVBVM
call IatPatch

push Hinstance
call InitializeComm
push Hinstance
call InitThreadCall

;Initialize Thread slots
push 04000h
Call InitPage
mov slot,eax

push eax
call FindFree
mov ebx,eax

mov esi,Me

invoke GetCurrentThreadId
mov DWORD ptr [ebx],eax
mov DWORD PTR [esi+92],eax

sub esp,16
push esp
push eax
call dwtoa

invoke GlobalAlloc,040h,24
mov DWORD PTR [ebx+4],eax
mov ebx,eax

mov eax,esp
invoke CreateWindowEx,0,offset WindowString,eax,0,0,0,0,0,0,0,0,0
add esp,16

mov DWORD PTR [esi+88],eax
mov DWORD PTR [ebx],eax
invoke SetTimer,eax,0,0,0
invoke InitializeCriticalSection,offset Globs
invoke InitializeCriticalSection,offset ARSync

pop eax
mov WORD PTR [eax],-1

Van:
mov edi,DWORD PTR [esi]
push esi
call DWORD PTR [edi+8] ;release

xor eax,eax
ret
InitCaller endp



CloseCaller proc public uses esi edi ebx Me:DWORD,Hinstance:DWORD

	mov esi,Me
	
	mov edi,DWORD PTR [esi]
	push esi
	call DWORD PTR [edi+4] ;add ref
	
	cmp WORD PTR [esi+76],0
	je Van

	push Hinstance
	call UnInitializeComm

	push slot
	call RemovePage

	invoke DeleteCriticalSection,offset Globs
	invoke DeleteCriticalSection,offset ARSync
	
Van:	
	push esi
	call DWORD PTR [edi+8] ;release
	xor eax,eax
	ret
CloseCaller endp

CreateNewThread proc  Me:DWORD,ThreadId:DWORD,StackSize:DWORD,ThreadPriority:DWORD,Reason:DWORD,Message:DWORD,Args:DWORD,ThreadType:DWORD,Rtn:DWORD

push esi
push edi
push ebx

mov esi,Me
mov edi,DWORD PTR [esi]
push esi
call DWORD PTR [edi+4] ;add ref


push ThreadType
push Args
push Message
push Reason
push ThreadPriority
push ThreadId
push StackSize
call CreateMT

mov ecx,Rtn
mov DWORD PTR [ecx],eax
push esi
call DWORD PTR [edi+8] ;release

pop ebx
pop edi
pop esi

xor eax,eax
ret
CreateNewThread endp


AddThread proc Me:DWORD,CallerInterface:DWORD
push esi
push edi
push ebx

mov esi,Me
mov edi,DWORD PTR [esi]
push esi
call DWORD PTR [edi+4] ;add ref


;invoke EnterCriticalSection,offset ARSync

invoke GetCurrentThreadId

push eax
push slot
invoke EnterCriticalSection,offset ARSync
call FindItem
cmp eax,0ffffffffh
je Van
mov eax,DWORD PTR [eax+4]
cmp DWORD PTR [eax+4],1
je Van ;Ako je vec Dodan...van..

invoke DelegateCaller,CallerInterface,DWORD PTR [esi+52],Me

mov BYTE PTR [esi+52],1

;SEND ADD NOTIFY

push eax
push DWORD PTR [esi+96]
push DWORD PTR [esi+84]
invoke LeaveCriticalSection,offset ARSync
call QueueUserAPC



GoRelease:
push esi
call DWORD PTR [edi+8] ;release

pop ebx
pop edi
pop esi
xor eax,eax
ret

Van:
invoke LeaveCriticalSection,offset ARSync
jmp GoRelease

AddThread endp





RemoveThread proc Me:DWORD,ThreadId:DWORD
	push esi
	push edi
	push ebx
	mov esi,Me
	mov edi,DWORD PTR [esi]
	push esi
	call DWORD PTR [edi+4] ;add ref

	mov ecx,ThreadId
	cmp ecx,0
	je TerMain
	cmp DWORD PTR [esi+92],ecx
	jne NijeMainThread
	
TerMain:
	invoke ExitProcess,0

NijeMainThread:

	;invoke EnterCriticalSection,offset ARSync

	push ThreadId
	push slot
	invoke EnterCriticalSection,offset ARSync
	call FindItem
	cmp eax,0ffffffffh
	je Van
	mov ebx,DWORD PTR [eax+4]
	push ebx
	mov DWORD PTR [eax],0
	mov DWORD PTR [eax+4],0

	invoke IsWindow,DWORD PTR [ebx]
	test eax,eax
	je NemaProzora
	invoke DestroyWindow,DWORD PTR [ebx]
NemaProzora:
	cmp DWORD PTR [ebx+12],0
	je Nema1
	
	invoke GetCurrentThreadId
	cmp eax,ThreadId
	je Isti
	invoke TerminateThread,DWORD PTR [ebx+12],0	
Isti:
	invoke CloseHandle,DWORD PTR [ebx+12]
Nema1:
	cmp DWORD PTR [ebx+16],0
	je Nema2
	invoke GlobalFree,DWORD PTR [ebx+16]
Nema2:	
	invoke CloseHandle,DWORD PTR [ebx+8]
	Call GlobalFree
	
	
	;SEND REMOVE NOTIFY!
	push ThreadId
	Or DWORD PTR [esp],080000000h
	push DWORD PTR [esi+96]
	push DWORD PTR [esi+84]
	invoke LeaveCriticalSection,offset ARSync
	call QueueUserAPC

	
GoRelease:	
	push esi
	call DWORD PTR [edi+8] ;release

	pop ebx
	pop edi
	pop esi
	xor eax,eax
	ret

	Van:
	invoke LeaveCriticalSection,offset ARSync
	jmp GoRelease


RemoveThread endp

CheckThreadExist proc Me:DWORD,ThreadId:DWORD,Rtn:DWORD
	push esi
	push edi
	push ebx
	mov esi,Me
	mov edi,DWORD PTR [esi]
	push esi
	call DWORD PTR [edi+4] ;add ref

	CMP ThreadId,0
	je Van
	
	sub esp,16
	push esp
	push ThreadId
	call dwtoa
	push esp
	push offset WindowString
	invoke EnterCriticalSection,offset ARSync
	call FindWindow
	add esp,16
	test eax,eax
	jne Postoji	
	
	push ThreadId
	push slot
	call FindItem
	cmp eax,0ffffffffh
	je NoStruct
	mov ebx,DWORD PTR [eax+4]
	push ebx
	mov DWORD PTR [eax],0
	mov DWORD PTR [eax+4],0

	cmp DWORD PTR [ebx+12],0
	je Nema1

	invoke CloseHandle,DWORD PTR [ebx+12]
Nema1:
	cmp DWORD PTR [ebx+16],0
	je Nema2
	invoke GlobalFree,DWORD PTR [ebx+16]
Nema2:	
	invoke CloseHandle,DWORD PTR [ebx+8]
	Call GlobalFree
	
	push ThreadId
	Or DWORD PTR [esp],080000000h
	push DWORD PTR [esi+96]
	push DWORD PTR [esi+84]
	invoke LeaveCriticalSection,offset ARSync
	call QueueUserAPC
	mov ecx,Rtn
	mov DWORD PTR [ecx],0
Van:	
	push esi
	call DWORD PTR [edi+8] ;release

	pop ebx
	pop edi
	pop esi
	xor eax,eax
	ret

Postoji:
	invoke LeaveCriticalSection,offset ARSync
	mov ecx,Rtn
	mov DWORD PTR [ecx],-1
	jmp Van
NoStruct:
 	invoke LeaveCriticalSection,offset ARSync
	mov ecx,Rtn
	mov DWORD PTR [ecx],0
	jmp Van
CheckThreadExist endp

CallThread proc Me:DWORD,ThreadId:DWORD,Reason:DWORD,Message:DWORD,Args:DWORD,CopyMove:DWORD,CallingType:DWORD,IsValidCall:DWORD,Rtn:DWORD
sub esp,4
	push esi
	push edi
	push ebx
	mov esi,Me
	mov edi,DWORD PTR [esi]
	push esi
	call DWORD PTR [edi+4] ;add ref

;invoke EnterCriticalSection,offset ARSync

cmp ThreadId,0
jne ToId
push DWORD PTR [esi+92]
jmp ToMain

ToId:
push ThreadId
ToMain:
push slot
invoke EnterCriticalSection,offset ARSync
call FindItem
cmp eax,0FFFFFFFFh
je NotValid
mov ebx,DWORD PTR [eax+4]


	push Args
	push Message
	push Reason

	cmp CopyMove,1
	je MoveVR
	call DuplicateVariant
	jmp Proceed
MoveVR:
	call MoveVariant
Proceed:
	.if CallingType==0
	push eax
	PUSH DWORD PTR [ebx+12]
	push DWORD PTR [ebx+16]
	invoke LeaveCriticalSection,offset ARSync
	call QueueUserAPC
	mov ecx,IsValidCall
	mov DWORD PTR [ecx],eax
	
	.elseif CallingType==1
	push eax
	push DWORD PTR [ebx+16]
	push 0c000h
	push DWORD PTR [ebx]
	invoke LeaveCriticalSection,offset ARSync
	call SendMessage
	
	mov ecx,Rtn
	mov DWORD PTR [ecx],eax
	mov ecx,IsValidCall
	mov DWORD PTR [ecx],-1
	
	.elseif CallingType==2
	push eax
	push DWORD PTR [ebx+16]
	push 0c000h
	push DWORD PTR [ebx]
	invoke LeaveCriticalSection,offset ARSync
	call PostMessage
	mov ecx,IsValidCall
	mov DWORD PTR [ecx],eax
	
	.elseif CallingType==3
	lea ecx,[ebp-4]
	push ecx
	push eax
	mov ecx,DWORD PTR [ebx+20]
	push ecx
	mov ebx,DWORD PTR [ecx]
	invoke LeaveCriticalSection,offset ARSync

	call DWORD PTR [ebx+28]
	test eax,eax
	jne NotValid
	mov ecx,Rtn
	mov eax,DWORD PTR [ebp-4]
	mov DWORD PTR [ecx],eax
	mov ecx,IsValidCall
	mov DWORD PTR [ecx],-1
	.endif

Van:
	push esi
	call DWORD PTR [edi+8] ;release
	
	pop ebx
	pop edi
	pop esi
	xor eax,eax
	ret

NotValid:
invoke LeaveCriticalSection,offset ARSync
	mov ecx,IsValidCall
	mov DWORD PTR [ecx],0
	jmp Van
CallThread endp


EnterSync proc Me:DWORD,WaitOwner:DWORD,Rtn:DWORD

push esi
push edi
push ebx

;mov esi,Me
;mov edi,DWORD PTR [esi]
;push esi
;call DWORD PTR [edi+4] ;addref
	

 cmp WaitOwner,0
 jne Dalje
 invoke EnterCriticalSection,offset Globs
 jmp Van

Dalje:	
invoke TryEnterCriticalSection,offset Globs
mov ecx,Rtn
mov DWORD PTR [ecx],eax

Van:	
;	push esi
;	call DWORD PTR [edi+8] ;release

pop ebx
pop edi
pop esi
xor eax,eax
	
	ret
EnterSync endp

LeaveSync proc Me:DWORD

push esi
push edi
push ebx

;	mov esi,Me
;	mov edi,DWORD PTR [esi]
;	push esi
;	call DWORD PTR [edi+4] ;addref

 invoke LeaveCriticalSection,offset Globs


;	push esi
;	call DWORD PTR [edi+8] ;release
	pop ebx
	pop edi
	pop esi
	xor eax,eax

	ret
LeaveSync endp

WaitLocalObject proc Me:DWORD,ThreadId:DWORD,WaitTime:DWORD,Rtn:DWORD
sub esp,8

push esi
push edi
push ebx

mov esi,Me
mov edi,DWORD PTR [esi]
push esi
call DWORD PTR [edi+4] ;addref
	
;invoke EnterCriticalSection,offset ARSync	
cmp ThreadId,0
jne ToId
push DWORD PTR [esi+92]
jmp ToMain

ToId:
push ThreadId
ToMain:
push slot
invoke EnterCriticalSection,offset ARSync	
call FindItem
cmp eax,0FFFFFFFFh
je NotValid
mov ebx,DWORD PTR [eax+4]	
push DWORD PTR [ebx+8]
pop DWORD PTR [ebp-8]
push DWORD PTR [ebx+12]
pop DWORD PTR [ebp-4]
invoke LeaveCriticalSection,offset ARSync
lea eax,[ebp-8]
invoke WaitForMultipleObjects,2,eax,0,WaitTime
cmp eax,1
jne	Dalje
mov eax,0ffffffffh
Dalje:
mov ecx,Rtn
mov DWORD PTR [ecx],eax

ExitIt:	
	push esi
	call DWORD PTR [edi+8] ;release
	pop ebx
	pop edi
	pop esi
	xor eax,eax
	ret

NotValid:
invoke LeaveCriticalSection,offset ARSync
mov ecx,Rtn
mov DWORD PTR [ecx],0ffffffffh ;return Invalid
jmp ExitIt

WaitLocalObject endp



SignalLocalObject proc Me:DWORD,ThreadId:DWORD,Rtn:DWORD
push esi
push edi
push ebx

mov esi,Me
mov edi,DWORD PTR [esi]
push esi
call DWORD PTR [edi+4] ;addref	

;invoke EnterCriticalSection,offset ARSync	
cmp ThreadId,0
jne ToId
push DWORD PTR [esi+92]
jmp ToMain
ToId:
push ThreadId
ToMain:
push slot
invoke EnterCriticalSection,offset ARSync
call FindItem
cmp eax,0FFFFFFFFh
je NotValid
mov ebx,DWORD PTR [eax+4]
mov ebx,DWORD PTR [ebx+8]
invoke LeaveCriticalSection,offset ARSync
invoke SetEvent,ebx
mov ecx,Rtn
mov DWORD PTR [ecx],eax

ExitIt:	
	push esi
	call DWORD PTR [edi+8] ;release
	pop ebx
	pop edi
	pop esi
	xor eax,eax
	ret
	
NotValid:
invoke LeaveCriticalSection,offset ARSync
mov ecx,Rtn
mov DWORD PTR [ecx],0ffffffffh ;return Invalid
jmp ExitIt

SignalLocalObject endp






TranslateArg proc Me:DWORD,CallArgs:DWORD,Reason:DWORD,Message:DWORD,Arguments:DWORD
	push esi
	push edi
	push ebx

	mov esi,Me
	mov edi,DWORD PTR [esi]
	push esi
	call DWORD PTR [edi+4] ;addref
	
	push Arguments
	push Message
	push Reason
	push CallArgs
	call TranslateArguments	
	
	push esi
	call DWORD PTR [edi+8] ;release
	pop ebx
	pop edi
	pop esi
	xor eax,eax
	ret

TranslateArg endp

GetThreadPar proc Me:DWORD,Reason:DWORD,Message:DWORD,Arguments:DWORD
	push esi
	push edi
	push ebx

	mov esi,Me
	mov edi,DWORD PTR [esi]
	push esi
	call DWORD PTR [edi+4] ;addref


	invoke TlsGetValue,020h
	test eax,eax
	je Van
	
	push Arguments
	push Message
	push Reason
	push eax
	call TranslateArguments	

	invoke TlsSetValue,020h,0	
	
	Van:
	push esi
	call DWORD PTR [edi+8] ;release
	pop ebx
	pop edi
	pop esi
	xor eax,eax
	ret

GetThreadPar endp

AttachCallArgs proc Me:DWORD,ThreadId:DWORD,Index:DWORD,Reason:DWORD,Message:DWORD,Args:DWORD,CopyMove:DWORD,PreviousCA:DWORD,AttachExist:DWORD,Rtn:DWORD
sub esp,4
push esi
push edi
push ebx

mov esi,Me
mov edi,DWORD PTR [esi]
push esi
call DWORD PTR [edi+4] ;addref	

;invoke EnterCriticalSection,offset ARSync	
cmp ThreadId,0
jne ToId
push DWORD PTR [esi+92]
jmp ToMain
ToId:
push ThreadId
ToMain:
push slot
invoke EnterCriticalSection,offset ARSync
call FindItem
cmp eax,0FFFFFFFFh
je NotValid
mov ebx,DWORD PTR [eax+4]


sub esp,16
push esp
push Index
call dwtoa

invoke GetProp,DWORD PTR [ebx],esp
test eax,eax
je Nema

cmp AttachExist,0
jne AttachIt
mov ecx,PreviousCA
mov DWORD PTR [ecx],-1
jmp PrijeExit

AttachIt:
mov ecx,PreviousCA
mov DWORD PTR [ecx],eax
Nema:

push Args
push Message
push Reason

cmp CopyMove,1
je MoveVR
call DuplicateVariant
jmp Proceed
MoveVR:
call MoveVariant
Proceed:	
push eax
lea eax,[esp+4]
push eax
push DWORD PTR [ebx]
Call SetProp	


PrijeExit:

invoke LeaveCriticalSection,offset ARSync	

add esp,16
mov ecx,Rtn
mov DWORD PTR [ecx],0

ExitIt:	
	push esi
	call DWORD PTR [edi+8] ;release
	pop ebx
	pop edi
	pop esi
	xor eax,eax
	ret

NotValid:
invoke LeaveCriticalSection,offset ARSync
mov ecx,Rtn
mov DWORD PTR [ecx],0FFFFFFFFh ;return Invalid
jmp ExitIt
AttachCallArgs endp

DetachCallArgs proc Me:DWORD,ThreadId:DWORD,Index:DWORD,Reason:DWORD,Message:DWORD,Args:DWORD,DetachType:DWORD,Rtn:DWORD
push esi
push edi
push ebx

mov esi,Me
mov edi,DWORD PTR [esi]
push esi
call DWORD PTR [edi+4] ;addref	

;invoke EnterCriticalSection,offset ARSync	
cmp ThreadId,0
jne ToId
push DWORD PTR [esi+92]
jmp ToMain
ToId:
push ThreadId
ToMain:
push slot
invoke EnterCriticalSection,offset ARSync
call FindItem
cmp eax,0FFFFFFFFh
je NotValid
mov ebx,DWORD PTR [eax+4]

sub esp,16
push esp
push Index
call dwtoa
push esp
push DWORD PTR [ebx]
Call GetProp
test eax,eax
jne Ima
add esp,16
jmp NotValid

Ima:


cmp DetachType,0
je NeSkidaj
push Args
push Message
push Reason
push eax
call TranslateArguments	
invoke RemoveProp,DWORD PTR [ebx],esp
jmp Izlaz
NeSkidaj:
mov ecx,Reason
push DWORD PTR [eax]
pop DWORD PTR [ecx]
mov ecx,Message
push DWORD PTR [eax+4]
pop DWORD PTR [ecx]
lea ecx,[eax+8]
invoke VariantCopy,Args,ecx


Izlaz:
invoke LeaveCriticalSection,offset ARSync

add esp,16
mov ecx,Rtn
mov DWORD PTR [ecx],0	


ExitIt:	
	push esi
	call DWORD PTR [edi+8] ;release
	pop ebx
	pop edi
	pop esi
	xor eax,eax
	ret

NotValid:
invoke LeaveCriticalSection,offset ARSync
mov ecx,Rtn
mov DWORD PTR [ecx],0FFFFFFFFh ;return Invalid
jmp ExitIt
DetachCallArgs endp

RemoveCallArgs proc Me:DWORD,ThreadId:DWORD,Index:DWORD,Rtn:DWORD

push esi
push edi
push ebx

mov esi,Me
mov edi,DWORD PTR [esi]
push esi
call DWORD PTR [edi+4] ;addref	

;invoke EnterCriticalSection,offset ARSync	
cmp ThreadId,0
jne ToId
push DWORD PTR [esi+92]
jmp ToMain
ToId:
push ThreadId
ToMain:
push slot
invoke EnterCriticalSection,offset ARSync	
call FindItem
cmp eax,0FFFFFFFFh
je NotValid
mov ebx,DWORD PTR [eax+4]

sub esp,16
push esp
push Index
call dwtoa
push esp
push DWORD PTR [ebx]
Call GetProp
test eax,eax
jne Ima
add esp,16
jmp NotValid

Ima:
push esp
push DWORD PTR [ebx]
mov ebx,eax
call RemoveProp

invoke LeaveCriticalSection,offset ARSync
lea eax,[ebx+8]
invoke VariantClear,eax
invoke GlobalFree,ebx
add esp,16
mov ecx,Rtn
mov DWORD PTR [ecx],0


ExitIt:	
	push esi
	call DWORD PTR [edi+8] ;release
	pop ebx
	pop edi
	pop esi
	xor eax,eax
	ret

NotValid:
invoke LeaveCriticalSection,offset ARSync
mov ecx,Rtn
mov DWORD PTR [ecx],0FFFFFFFFh ;return Invalid
jmp ExitIt
RemoveCallArgs endp




FastEvent proc public uses esi edi ebx
mov esi,DWORD PTR [esp+20]
sub esp,28
lea ebx,[esp]
invoke PeekMessage,ebx,esi,0,0,3
test eax,eax
je NoWait
invoke TranslateMessage,ebx
invoke DispatchMessage,ebx
NoWait:
add esp,28
ret 8
FastEvent endp


EnterPump proc public uses esi edi ebx
sub esp,28
lea ebx,[esp]
MPump:
invoke GetMessage,ebx,0,0,0
test eax,eax
je NoWait
invoke TranslateMessage,ebx
invoke DispatchMessage,ebx
jmp MPump
NoWait:
add esp,28
ret 4
EnterPump endp

ExitPump proc public uses esi edi ebx
invoke PostQuitMessage,0	
ret 4
ExitPump endp


UpdateVM proc Me:DWORD,Rtn:DWORD
push esi
push edi
push ebx	

;Check Version 9237
mov esi,MSVBVMH
add esi,03bf72h
cmp DWORD PTR [esi],840fc085h
je Druga1
jmp Dalje1
Druga1:
cmp DWORD PTR [esi+4],0fffd2283h
jne Dalje1
mov edi,offset Version9237
jmp Ima

Nema:
mov ecx,Rtn
mov DWORD PTR [ecx],0
pop ebx
pop edi
pop esi
ret

Ima:
mov ecx,Rtn
mov DWORD PTR [ecx],-1
invoke WriteProcessMemory,-1,esi,edi,8,0
pop ebx
pop edi
pop esi
ret


Dalje1:
;Check Version 8877
mov esi,MSVBVMH
add esi,03c299h
cmp DWORD PTR [esi],840fc085h
je Druga2
jmp Dalje2
Druga2:
cmp DWORD PTR [esi+4],0fffd1f6ch
jne Dalje2
mov edi,offset Version8877
jmp Ima


Dalje2:
;Check Version 8176
mov esi,MSVBVMH
add esi,077f6fh
cmp DWORD PTR [esi],840fc085h
je Druga3
jmp Dalje3
Druga3:
cmp DWORD PTR [esi+4],0ffff2200h
jne Dalje3
mov edi,offset Version8176
jmp Ima

Dalje3:
;Check Version 8268
mov esi,MSVBVMH
add esi,077498h
cmp DWORD PTR [esi],840fc085h
je Druga4
jmp Dalje4
Druga4:
cmp DWORD PTR [esi+4],0fffad057h
jne Dalje4
mov edi,offset Version8268
jmp Ima

Dalje4:
;Check Version 8495
mov esi,MSVBVMH
add esi,03b886h
cmp DWORD PTR [esi],840fc085h
je Druga5
jmp Dalje5
Druga5:
cmp DWORD PTR [esi+4],0fffd28fch
jne Dalje5
mov edi,offset Version8495
jmp Ima

Dalje5:
;Check Version 8964
mov esi,MSVBVMH
add esi,03b55dh
cmp DWORD PTR [esi],840fc085h
je Druga6
jmp Nema
Druga6:
cmp DWORD PTR [esi+4],0fffd2ca8h
jne Nema
mov edi,offset Version8964
jmp Ima


UpdateVM endp




About proc Me:DWORD
	mov esi,Me
	mov edi,DWORD PTR [esi]
	push esi
	call DWORD PTR [edi+4]
	
	invoke DialogBoxParam,HInst,101,0,DialogProc,0

	push esi
	call DWORD PTR [edi+8] ;release
	xor eax,eax
	ret

About endp


DialogProc proc Hwnd:DWORD,uMsg:DWORD,wParam:DWORD,lParam:DWORD
.if uMsg==WM_COMMAND
mov		eax,wParam
and		eax,0FFFFh
cmp eax,1
jne Dalje
invoke DestroyWindow,Hwnd
Dalje:
.endif
xor eax,eax
ret
DialogProc endp


DelegateCaller proc CallerInterface:DWORD,IsMainThread:DWORD,Obj:DWORD
sub esp,8

push esi
push edi
push ebx

push 1
push 1
push CallerInterface
Call Delegator
mov DWORD PTR [ebp-4],eax

;OPENTHREAD...
invoke GetCurrentThreadId
push eax
push eax

invoke GetCurrentThread
lea ecx,[ebp-8]
invoke DuplicateHandle,-1,eax,-1,ecx,01F03FFh,1,0

push slot
call FindItem
mov ebx,DWORD PTR [eax+4]

mov DWORD PTR [ebx+4],1

push DWORD PTR [ebp-8]
pop DWORD ptr [ebx+12]

push dword ptr [ebp-4]
pop DWORD ptr [ebx+16]

invoke CreateEvent,0,0,0,0
mov DWORD ptr [ebx+8],eax

mov eax,CallerInterface
mov eax,DWORD PTR [eax]
mov DWORD PTR [ebx+20],eax


cmp IsMainThread,0
jne Dalje
mov esi,Obj
push DWORD PTR [ebp-8]
pop DWORD PTR [esi+96]
push 1
push 2
push CallerInterface
call Delegator
mov DWORD PTR [esi+84],eax
Dalje:


pop eax
pop ebx
pop edi
pop esi
ret
DelegateCaller endp


InitPage proc public uses esi edi ebx
mov eax,DWORD PTR [esp+010h]
inc eax
shl eax,3
invoke VirtualAlloc,0,eax,MEM_COMMIT,PAGE_READWRITE
push DWORD PTR [esp+010h]
pop DWORD PTR [eax]
ret 4
InitPage endp


RemovePage proc public uses esi edi ebx
mov eax,DWORD PTR [esp+010h]
mov ecx,DWORD ptr [eax]
inc ecx
shl ecx,3
invoke VirtualFree,eax,ecx,MEM_DECOMMIT
ret 4
RemovePage endp


;PageHandle,return=memptr
FindFree proc
mov eax,DWORD PTR [esp+4]
xor ecx,ecx
Ponovo:
cmp ecx,DWORD PTR [eax]
jg NotValid
inc ecx
cmp DWORD PTR [eax+ecx*8],0
jne Ponovo
lea eax,DWORD PTR [eax+ecx*8]

ret 4
NotValid:
or eax,-1
ret 4

FindFree endp

;PageHandle,Index
FindItem proc
mov eax,DWORD PTR [esp+4]
mov edx,DWORD PTR [esp+8]
xor ecx,ecx
Ponovo:
cmp ecx,DWORD PTR [eax]
jg NotValid
inc ecx
cmp DWORD PTR [eax+ecx*8],edx
jne Ponovo
lea eax,DWORD PTR [eax+ecx*8]
ret 8
NotValid:
or eax,-1
ret 8
FindItem endp


;dwValue:DWORD, lpBuffer:DWORD
dwtoa proc 

	push esi
	push edi
	push ebx
	
    mov eax, DWORD PTR [esp+16]
    mov edi, DWORD PTR [esp+20]

 ;   or eax,eax
 ;   jnz sign
  
  	test eax,eax
  	jne sign
    
  zero:
    mov word ptr [edi],30h
	pop ebx
	pop edi
	pop esi
    ret 8
    
  sign:
    jns pos
    mov byte ptr [edi],'-'
    neg eax
    inc edi

  pos:      
    mov ecx,429496730
    mov esi, edi

    .while (eax > 0)
      mov ebx,eax
      mul ecx
      mov eax,edx
      lea edx,[edx*4+edx]
      add edx,edx
      sub ebx,edx
      add bl,'0'
      mov byte ptr [edi],bl
      inc edi
    .endw

    mov byte ptr [edi], 0  

	pop ebx
	pop edi
	pop esi
    ret 8

dwtoa endp


end DllEntry
