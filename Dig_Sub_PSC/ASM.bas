Attribute VB_Name = "ASM"
'Module1.bas

Option Explicit


Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
(ByVal lpMCode As Long, _
ByVal Long1 As Long, ByVal Long2 As Single, _
ByVal Long3 As Single, ByVal Long4 As Long) As Long

Public MMXCode() As Byte     'Array to hold machine code

'MCode Structure
Public Type MCodeStruc
   PICW As Long
   PICH As Long
   PtrARR0 As Long
   PtrARR1 As Long
   PtrARRRES As Long
   SUBMODE As Long      'Subtraction Mode (Kind of Subtraction 0=normal(Mode 1), 1=xor(Mode 2) , 2=.....(Mode 3) ,.....)
   BGL As Long          'Base Grey Level
   WF As Long           'Weighting Factor
   INVERT As Long
End Type
Public MCode As MCodeStruc
'-------------------------------------


Public Sub ASM_DigSub()
Dim ptrStruc As Long
Dim ptrMC As Long
Dim res As Long
Dim Index As Long
   ' Fill MCodeStruc
   MCode.PICW = PICW(0)
   MCode.PICH = PICH(0)
   MCode.PtrARR0 = VarPtr(ARR0(1, 1))
   MCode.PtrARR1 = VarPtr(ARR1(1, 1))
   MCode.PtrARRRES = VarPtr(ARRRes(1, 1))
   MCode.SUBMODE = TheMode
   MCode.BGL = GreyLevel
   MCode.WF = Weighting
   MCode.INVERT = InvertYN
   ptrStruc = VarPtr(MCode.PICW)
   ptrMC = VarPtr(MMXCode(1))
   res = CallWindowProc(ptrMC, ptrStruc, 0&, 0&, 0&)
End Sub

Public Sub Loadmcode(InFile$, MCCode() As Byte)
Dim MCSize As Long
   'Load machine code into MCCode() byte array
   On Error GoTo InFileErr
   If Dir$(InFile$) = "" Then
      MsgBox (InFile$ & " missing")
      DoEvents
      Unload Form1
      End
   End If
   Open InFile$ For Binary As #1
   MCSize = LOF(1)
   If MCSize = 0 Then
InFileErr:
      MsgBox (InFile$ & " missing")
      DoEvents
      Unload Form1
      End
   End If
   ReDim MCCode(1 To MCSize)
   Get #1, , MCCode
   Close #1
   On Error GoTo 0
End Sub


