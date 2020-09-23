VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Digital Subtraction"
   ClientHeight    =   9105
   ClientLeft      =   165
   ClientTop       =   -825
   ClientWidth     =   14745
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   607
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   983
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PIC_Save 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   14100
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   26
      Top             =   495
      Width           =   510
   End
   Begin VB.HScrollBar HSGreyLevel 
      Height          =   255
      Left            =   6075
      TabIndex        =   19
      Top             =   915
      Width           =   7320
   End
   Begin VB.HScrollBar HSWeighting 
      Height          =   270
      Left            =   6090
      TabIndex        =   18
      Top             =   525
      Width           =   7290
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run ASM"
      Height          =   240
      Index           =   1
      Left            =   9420
      TabIndex        =   13
      Top             =   105
      Width           =   1080
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run BASIC"
      Height          =   240
      Index           =   0
      Left            =   8160
      TabIndex        =   12
      Top             =   105
      Width           =   1080
   End
   Begin VB.CheckBox chkInvert 
      Caption         =   "Invert"
      Height          =   225
      Left            =   7140
      TabIndex        =   11
      Top             =   105
      Width           =   780
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Mode 1"
      Height          =   255
      Index           =   1
      Left            =   6045
      TabIndex        =   10
      Top             =   105
      Width           =   885
   End
   Begin VB.OptionButton optMode 
      Caption         =   "Mode 0"
      Height          =   255
      Index           =   0
      Left            =   5055
      TabIndex        =   9
      Top             =   90
      Width           =   885
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Pic 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   540
      TabIndex        =   8
      Top             =   4740
      Width           =   1080
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Pic 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   555
      TabIndex        =   7
      Top             =   90
      Width           =   1080
   End
   Begin VB.HScrollBar HSRes 
      Height          =   330
      Left            =   5055
      TabIndex        =   6
      Top             =   8625
      Width           =   3270
   End
   Begin VB.VScrollBar VSRes 
      Height          =   4650
      Left            =   4590
      TabIndex        =   5
      Top             =   1365
      Width           =   345
   End
   Begin VB.PictureBox PICRes 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   5055
      ScaleHeight     =   478
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   618
      TabIndex        =   4
      Top             =   1380
      Width           =   9300
   End
   Begin VB.HScrollBar HS 
      Height          =   315
      Left            =   585
      TabIndex        =   3
      Top             =   4320
      Width           =   3795
   End
   Begin VB.VScrollBar VS 
      Height          =   3885
      Left            =   165
      TabIndex        =   2
      Top             =   405
      Width           =   330
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3840
      Index           =   1
      Left            =   555
      ScaleHeight     =   254
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   254
      TabIndex        =   1
      Top             =   5100
      Width           =   3840
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3840
      Index           =   0
      Left            =   555
      ScaleHeight     =   254
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   254
      TabIndex        =   0
      Top             =   420
      Width           =   3840
   End
   Begin VB.Label LabNWH 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   1725
      TabIndex        =   25
      Top             =   4770
      Width           =   60
   End
   Begin VB.Label LabNWH 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   1770
      TabIndex        =   24
      Top             =   120
      Width           =   60
   End
   Begin VB.Label LabInvert 
      Height          =   180
      Left            =   7905
      TabIndex        =   23
      Top             =   105
      Width           =   225
   End
   Begin VB.Label LabMode 
      Height          =   180
      Left            =   6945
      TabIndex        =   22
      Top             =   105
      Width           =   150
   End
   Begin VB.Label LabGreyLevel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "G"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   13500
      TabIndex        =   21
      Top             =   900
      Width           =   480
   End
   Begin VB.Label LabWeighting 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "W"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   13500
      TabIndex        =   20
      Top             =   540
      Width           =   480
   End
   Begin VB.Label LabG 
      Caption         =   "GreyLevel"
      Height          =   195
      Left            =   5085
      TabIndex        =   17
      Top             =   930
      Width           =   780
   End
   Begin VB.Label LabW 
      Caption         =   "Weighting"
      Height          =   225
      Left            =   5070
      TabIndex        =   16
      Top             =   540
      Width           =   840
   End
   Begin VB.Label LabRes 
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13125
      TabIndex        =   15
      Top             =   90
      Width           =   1005
   End
   Begin VB.Label LabTime 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   10740
      TabIndex        =   14
      Top             =   105
      Width           =   375
   End
   Begin VB.Menu mnuF 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Save Result"
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Exit"
         Index           =   1
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Info"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Digital Subtraction by  Robert Rayment

' 19/02/05
' Faster Basic
' ASM action on GreyLevel & Weighting Scrollbars

' Formulae:
' Mode 0:  GreyLevel +/- (RGB0 - RGB1) * Weighting
' Mode 1:  GreyLevel -/+ (RGB0 XOR RGB1) * Weighting
'          Invert swaps sign


' Main.frm Form1
Option Explicit

Dim aRun As Boolean

Dim tmHowLong As CTimingPC
Dim CommonDialog1 As OSDialog

' Variables:
' PicBox, Array,    Width,   Height,  Loaded boolean
' PIC(0), ARR0(),   PICW(0), PICH(0), aPIC(0),        picture 0
' PIC(1), ARR1(),   PICW(1), PICH(1), aPIC(1),        picture 1
' PICRes, ARRREs()  resulting picture

' Weighting & GreyLevel & InvertYN

' Public Const PICWOrg As Long = 256
' Public Const PICHOrg As Long = 256
' Public Const PICResWOrg As Long = 620
' Public Const PICResHOrg As Long = 480


Private Sub Form_Load()
'Public Const FormWOrg As Long = 14865
'Public Const FormHOrg As Long = 9975

   PIC_Save.Visible = False
   
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY

   If Screen.Width \ STX < 1024 Then
      MsgBox "Minimum screen resolution must be >= 1024 x 768", vbCritical, "DigSub"
      Unload Me
      End
   End If
   
   Me.Width = FormWOrg
   Me.Height = FormHOrg
   PICResWDef = 620
   PICResHDef = 480
   PICResWOrg = PICResWDef
   PICResHOrg = PICResHDef
   
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   CurrPath$ = PathSpec$

   PositionControls
   ReDim aPIC(2)
   ReDim PICW(2), PICH(2)
   ReDim aPIC(2)
   ' Start values
   aPIC(0) = False
   aPIC(1) = False
   TheMode = 0
   InvertYN = 0
   optMode(0).Value = True
   chkInvert.Value = Unchecked
   cmdRun(0).Enabled = False
   cmdRun(1).Enabled = False
   mnuFile(0).Enabled = False
   aRun = False
   
   Loadmcode PathSpec$ & "MMXDigSubRR.bin", MMXCode()

End Sub

Private Sub cmdLoad_Click(Index As Integer)
Dim Title$, Filt$, InDir$
Dim FIndex As Long

Dim iBPP As Integer
   aRES = False
   aRun = False
   aPIC(Index) = False
   mnuFile(0).Enabled = False
   
   'Filt$ = "Pics bmp,jpg,gif|*.bmp;*.jpg;*.gif"
   Filt$ = "BMP(*.bmp)|*.bmp|JPEG(*.jpg)|*.jpg|GIF(*.gif)|*.gif"

   FileSpec$ = ""
   Title$ = "Load PIC" & Str$(Index)
   InDir$ = CurrPath$ 'Pathspec$
   Set CommonDialog1 = New OSDialog
   CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd, FIndex
   Set CommonDialog1 = Nothing
   If Len(FileSpec$) > 0 Then
      CurrPath$ = FileSpec$
      
      PIC(Index).Picture = LoadPicture(FileSpec$)
      If FIndex = 1 Then Mul = 1 Else Mul = -1
      GetObjectAPI PIC(Index), Len(PICWH), PICWH
      iBPP = PICWH.bmBitsPixel      ' 24 bpp
      PICW(Index) = PICWH.bmWidth
      PICH(Index) = PICWH.bmHeight
      
      If Index = 0 Then
         ReDim ARR0(1 To PICW(0), 1 To PICH(0))
         GETLONGS PIC(0).Picture, ARR0(), PICW(Index), PICH(Index)
         aPIC(0) = True
      Else
         ReDim ARR1(1 To PICW(1), 1 To PICH(1))
         GETLONGS PIC(1).Picture, ARR1(), PICW(Index), PICH(Index)
         aPIC(1) = True
      End If
      DoEvents
      LabNWH(Index) = " " & FName$(FileSpec$) & Str$(PICW(Index)) & " x" & Str$(PICH(Index)) & " "
      PICRes.SetFocus
   End If

   If aPIC(0) And aPIC(1) Then
      If (PICW(0) <> PICW(1)) Or _
         (PICH(0) <> PICH(1)) Then
         MsgBox "Make sure pictures are the same size", vbCritical, " "
         cmdRun(0).Enabled = False
         cmdRun(1).Enabled = False
         Exit Sub
      End If
      Mul = -1
      ReDim ARRRes(1 To PICW(0), 1 To PICH(0))
      aSBarsActive = False
      SetScrollBars
      aSBarsActive = True
      cmdRun(0).Enabled = True
      cmdRun(1).Enabled = True
   End If
End Sub


Private Sub cmdRun_Click(Index As Integer)
'Public Const PICResWOrg As Long = 620
'Public Const PICResHOrg As Long = 480
   aSBarsActive = False
   ' To Set PICRes size as image
   If PICW(0) <= PICRes.Width Then
      PICRes.Width = PICW(0)
   Else
      If PICW(0) <= PICResWOrg Then
         PICRes.Width = PICW(0)
      Else
         PICRes.Width = PICResWOrg
      End If
   End If
   If PICH(0) <= PICRes.Height Then
      PICRes.Height = PICH(0)
   Else
      If PICH(0) <= PICResHOrg Then
         PICRes.Height = PICH(0)
      Else
         PICRes.Height = PICResHOrg
      End If
   End If
   PositionPICResScrollBars
   SetResScrollBars
   aSBarsActive = True
      
   Set tmHowLong = New CTimingPC
   
   Select Case Index
   Case 0   ' BASIC
      tmHowLong.Reset
      PICRes.Picture = LoadPicture  ' To Flash
      PICRes.Refresh
      RunBASIC
      LabTime = " BASIC timing = " & Format(tmHowLong.Elapsed, "0000") & " msec "
      aRES = True
      DisplayArray PICRes, PICRes.Width, PICRes.Height, ARRRes(), 0, 0, -1
      mnuFile(0).Enabled = True
   Case 1   ' ASM
      tmHowLong.Reset
      PICRes.Picture = LoadPicture  ' To Flash
      PICRes.Refresh
      ASM_DigSub
      LabTime = " ASM timing = " & Format(tmHowLong.Elapsed, "0000") & " msec "
      aRES = True
      DisplayArray PICRes, PICRes.Width, PICRes.Height, ARRRes(), 0, 0, -1
      mnuFile(0).Enabled = True
   End Select
   
   Set tmHowLong = Nothing
   aRun = True
End Sub



'#### SCROLLING ####

Private Sub SetScrollBars()
' VS
' HS
   aSBarsActive = False
   HS.Width = PICWOrg
   VS.Height = PICHOrg
   If PICW(0) > PICWOrg Then
      HS.Visible = True
      HS.Min = 0
      HS.Max = PICW(0) - PICWOrg - 1
      sorcX = HS.Min
      HS.Value = sorcX
   Else
      sorcX = HS.Min
      HS.Value = sorcX
      HS.Visible = False
   End If
   If PICH(0) > PICHOrg Then
      VS.Visible = True
      VS.Max = 1
      VS.Min = (PICH(0) - PICHOrg) ' - 1)
      sorcY = VS.Min
      VS.Value = sorcY
   Else
      VS.Max = 1
      VS.Min = (PICH(0) - PICHOrg) ' - 1)
      sorcY = VS.Min
      VS.Value = sorcY
      VS.Visible = False
   End If
   
   SetResScrollBars
   aSBarsActive = True
End Sub

Private Sub SetResScrollBars()

   If PICW(0) > PICRes.Width Then
      HSRes.Width = PICRes.Width
      HSRes.Visible = True
      HSRes.Min = 0
      HSRes.Max = PICW(0) - PICRes.Width - 1
      sorcXRes = HSRes.Min
      HSRes.Value = sorcXRes
   Else
      sorcXRes = HSRes.Min
      HSRes.Value = sorcXRes
      HSRes.Visible = False
   End If
   If PICH(0) > PICRes.Height Then
      VSRes.Height = PICRes.Height
      VSRes.Visible = True
      VSRes.Max = 1
      VSRes.Min = (PICH(0) - PICRes.Height) ' - 1)
      sorcYRes = VSRes.Min
      VSRes.Value = sorcYRes
   Else
      VSRes.Max = 1
      VSRes.Min = (PICH(0) - PICRes.Height) ' - 1)
      sorcYRes = VSRes.Min
      VSRes.Value = sorcYRes
      VSRes.Visible = False
   End If
End Sub

Private Sub HS_Change()
   Call HS_Scroll
End Sub

Private Sub HS_Scroll()
   If aSBarsActive Then
      sorcX = HS.Value
      sorcY = VS.Value
      DisplayArray PIC(0), PICWOrg, PICHOrg, ARR0(), sorcX, sorcY, -1 'Mul
      DisplayArray PIC(1), PICWOrg, PICHOrg, ARR1(), sorcX, sorcY, -1 'Mul
   End If
End Sub

Private Sub VS_Change()
   Call VS_Scroll
End Sub

Private Sub VS_Scroll()
   If aSBarsActive Then
      sorcX = HS.Value
      sorcY = VS.Value
      DisplayArray PIC(0), PICWOrg, PICHOrg, ARR0(), sorcX, sorcY, -1 'Mul
      DisplayArray PIC(1), PICWOrg, PICHOrg, ARR1(), sorcX, sorcY, -1 'Mul
   End If
End Sub

Private Sub HSRes_Change()
   Call HSRes_Scroll
End Sub

Private Sub HSRes_Scroll()
   If aSBarsActive And aRES Then
      sorcXRes = HSRes.Value
      sorcYRes = VSRes.Value
      DisplayArray PICRes, PICRes.Width, PICRes.Height, ARRRes(), sorcXRes, sorcYRes, -1 'Mul
   End If
End Sub

Private Sub VSRes_Change()
   Call VSRes_Scroll
End Sub

Private Sub VSRes_Scroll()
   If aSBarsActive And aRES Then
      sorcXRes = HSRes.Value
      sorcYRes = VSRes.Value
      DisplayArray PICRes, PICResWOrg, PICResHOrg, ARRRes(), sorcXRes, sorcYRes, -1 'Mul
   End If
End Sub


Private Sub HSWeighting_Change()
   Call HSWeighting_Scroll
End Sub

Private Sub HSWeighting_Scroll()
   Weighting = HSWeighting.Value
   LabWeighting = Weighting
   If aSBarsActive And aRun Then
      ASM_DigSub
      DisplayArray PICRes, PICRes.Width, PICRes.Height, ARRRes(), sorcXRes, sorcYRes, -1
   End If
End Sub

Private Sub HSGreyLevel_Change()
   Call HSGreyLevel_Scroll
End Sub

Private Sub HSGreyLevel_Scroll()
   GreyLevel = HSGreyLevel.Value
   LabGreyLevel = GreyLevel
   If aSBarsActive And aRun Then
      ASM_DigSub
      DisplayArray PICRes, PICRes.Width, PICRes.Height, ARRRes(), sorcXRes, sorcYRes, -1
   End If
End Sub

'#### END SCROLLING ####


Private Sub mnuHelp_Click()
   MsgBox "Digital Subtraction" & vbCr _
        & "by  Robert Rayment 2005" & vbCr & vbCr _
        & "NB needs screen res >= 1024 x 768." & vbCr & vbCr _
        & "NB pictures need to be the same size." & vbCr & vbCr _
        & "Mode 0:  GreyLevel +/- (RGB0 - RGB1) * Weighting" & vbCr _
        & "Mode 1:  GreyLevel -/+ (RGB0 XOR RGB1) * Weighting " & vbCr & vbCr _
        & "   Invert swaps sign" & vbCr & vbCr _
        & "ASM is MMX machine code." & vbCr & vbCr _
        , vbInformation, "Info"
        
End Sub


Private Sub optMode_Click(Index As Integer)
   TheMode = Index
End Sub


Private Sub chkInvert_Click()
   InvertYN = 1 - InvertYN
End Sub


Private Sub PositionControls()
'Public Const PICWOrg As Long = 256
'Public Const PICHOrg As Long = 256
'Public Const PICResWOrg As Long = 620
'Public Const PICResHOrg As Long = 480
'Public Const FormWOrg As Long = 14865
'Public Const FormHOrg As Long = 9975

Dim k As Long
   GetExtras Me.BorderStyle
   ' IN:  BStyle = Me.BorderStyle
   ' OUT: Public ExtraBorder, ExtraHeight

   For k = 0 To 1
   With PIC(k)
      .Width = PICWOrg
      .Height = PICHOrg
   End With
   Next k
   PIC(1).Left = PIC(0).Left
   With PICRes
      .Width = PICResWOrg
      .Height = PICResHOrg
   End With
   With VS
      .Top = PIC(0).Top
      .Height = PIC(0).Height
      .Left = PIC(0).Left - .Width - 2
      .TabStop = False
   End With
   With HS
      .Top = PIC(0).Top + PIC(0).Height + 2
      .Width = PIC(0).Width
      .Left = PIC(0).Left
      .TabStop = False
   End With
   
   cmdLoad(1).Left = cmdLoad(0).Left
   optMode(1).Top = optMode(0).Top
   chkInvert.Top = optMode(0).Top
   cmdRun(0).Top = optMode(0).Top
   cmdRun(1).Top = optMode(0).Top
   LabTime.Top = optMode(0).Top
   
   With HSWeighting
      .Min = 1
      .Max = 32
      .TabStop = False
      .Value = .Min
      Weighting = .Min
      LabWeighting = Weighting
   End With
   With HSGreyLevel
      .Min = 0
      .Max = 255
      .TabStop = False
      .Value = 128
      GreyLevel = .Value
      LabGreyLevel = GreyLevel
   End With
End Sub

Private Sub PositionPICResScrollBars()
   With VSRes
      .Top = PICRes.Top
      .Height = PICRes.Height
      .Left = PICRes.Left - .Width - 2
      .TabStop = False
   End With
   With HSRes
      .Top = PICRes.Top + PICRes.Height + 2
      .Width = PICRes.Width
      .Left = PICRes.Left
      .TabStop = False
   End With
End Sub

Private Sub Form_Resize()
'   PICResWDef = 620
'   PICResHDef = 480
'   PICResWOrg = PICResWDef
'   PICResHOrg = PICResHDef
'Public Const FormWOrg As Long = 14865
'Public Const FormHOrg As Long = 9975
   If WindowState = vbMinimized Then Exit Sub
   
   If Me.Width <= FormWOrg Or Me.Height <= FormHOrg Then
      Me.Width = FormWOrg
      Me.Height = FormHOrg + ExtraHeight
      PICResWOrg = PICResWDef
      PICResHOrg = PICResHDef
      PICRes.Width = PICResWOrg
      PICRes.Height = PICResHOrg
   ElseIf Me.Width > FormWOrg Or Me.Height > FormHOrg Then
      PICRes.Width = Me.Width / STX - PICRes.Left - 30
      PICRes.Height = Me.Height / STY - PICRes.Top - 100
      PICResWOrg = PICRes.Width
      PICResHOrg = PICRes.Height
   End If
   If aRES Then
      cmdRun_Click 1
   Else
      PositionPICResScrollBars
   End If
End Sub

Public Sub DisplayArray(p As PictureBox, DW As Long, DH As Long, ARR() As Long, _
   Optional ByVal sorcX As Long = 0, Optional ByVal sorcY As Long = 0, Optional ByVal Mul As Long = 1)
' DisplayArray PIC, DW, DH, ARR(),sorcX, sorcY, Mul
' Public ARR0(), ARR1(), ARRres()
' DW & DH = Display W & H : -
' Public Const PICWOrg As Long = 256
' Public Const PICHOrg As Long = 256
' Public Const PICResWOrg As Long = 620
' Public Const PICResHOrg As Long = 480
' sorcX & sorcY Scrollbar values

' Image in ARR()
Dim UX As Long, UY As Long
Dim BS As BITMAPINFO
UX = UBound(ARR(), 1)
UY = UBound(ARR(), 2)
   With BS.bmi
      .biSize = 40
      .biwidth = UX
      .biheight = UY * Mul ' if -1 required
      .biPlanes = 1
      .biBitCount = 32    ' Sets up 32-bit colors
      .biCompression = 0
      .biSizeImage = UX * UY
      .biXPelsPerMeter = 0
      .biYPelsPerMeter = 0
      .biClrUsed = 0
      .biClrImportant = 0
   End With
   
   p.Picture = LoadPicture
      
   SetStretchBltMode p.hdc, 3  ' NB Of Dest picbox
   If StretchDIBits(p.hdc, 0, 0, DW, DH, sorcX, sorcY, _
      DW, DH, ARR(1, 1), BS, DIB_PAL_COLORS, vbSrcCopy) = 0 Then
         MsgBox "StretchDIBits Error", vbCritical, " "
   End If
   p.Refresh
End Sub


Private Sub RunBASIC()
' Formulae:
' Mode 0:  GreyLevel +/- (RGB0 - RGB1) * Weighting
' Mode 1:  GreyLevel -/+ (RGB0 XOR RGB1) * Weighting
'          Invert swaps sign
Dim ix As Long, iy As Long, iyy As Long
Dim R0 As Byte, G0 As Byte, B0 As Byte
Dim R1 As Byte, G1 As Byte, B1 As Byte
Dim CulR As Long, CulG As Long, CulB As Long
Dim DiffMul As Long

   Select Case TheMode
   Case 0
      DiffMul = 1
      If InvertYN = 0 Then DiffMul = -1
      For iy = 1 To PICH(0)
         For ix = 1 To PICW(0)
            'LngToRGB ARR0(ix, iy), R0, G0, B0
CulR = ARR0(ix, iy)
R0 = (CulR And &HFF&)
G0 = (CulR And &HFF00&) / &H100&
B0 = (CulR And &HFF0000) / &H10000
            
            'LngToRGB ARR1(ix, iy), R1, G1, B1
CulR = ARR1(ix, iy)
R1 = (CulR And &HFF&)
G1 = (CulR And &HFF00&) / &H100&
B1 = (CulR And &HFF0000) / &H10000
            
            CulR = (GreyLevel - (1& * R0 - R1) * Weighting * DiffMul)
            CulG = (GreyLevel - (1& * G0 - G1) * Weighting * DiffMul)
            CulB = (GreyLevel - (1& * B0 - B1) * Weighting * DiffMul)
           
            If CulR < 0 Then CulR = 0
            If CulG < 0 Then CulG = 0
            If CulB < 0 Then CulB = 0
               
            ARRRes(ix, iy) = RGB(CulR, CulG, CulB)
         Next ix
      Next iy
   
   Case 1
      DiffMul = -1
      If InvertYN = 0 Then DiffMul = 1
      For iy = 1 To PICH(0)
         For ix = 1 To PICW(0)
            'LngToRGB ARR0(ix, iy), R0, G0, B0
CulR = ARR0(ix, iy)
R0 = (CulR And &HFF&)
G0 = (CulR And &HFF00&) / &H100&
B0 = (CulR And &HFF0000) / &H10000
            'LngToRGB ARR1(ix, iy), R1, G1, B1
CulR = ARR1(ix, iy)
R1 = (CulR And &HFF&)
G1 = (CulR And &HFF00&) / &H100&
B1 = (CulR And &HFF0000) / &H10000
            CulR = (GreyLevel - (R0 Xor R1) * Weighting * DiffMul)
            CulG = (GreyLevel - (G0 Xor G1) * Weighting * DiffMul)
            CulB = (GreyLevel - (B0 Xor B1) * Weighting * DiffMul)
           
            If CulR < 0 Then CulR = 0
            If CulG < 0 Then CulG = 0
            If CulB < 0 Then CulB = 0
               
            ARRRes(ix, iy) = RGB(CulR, CulG, CulB)
         Next ix
      Next iy
   End Select
End Sub

Private Sub mnuFile_Click(Index As Integer)
Dim Title$, Filt$, InDir$
Dim FIndex As Long
   Select Case Index
   Case 0   ' Save Result As 24bpp BMP
      Filt$ = "BMP(*.bmp)|*.bmp"
      FileSpec$ = ""
      Title$ = "Save Result 24bpp BMP"
      InDir$ = CurrPath$ 'Pathspec$
      Set CommonDialog1 = New OSDialog
      CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd, FIndex
      Set CommonDialog1 = Nothing
      If Len(FileSpec$) > 0 Then
         With PIC_Save
            .Width = PICW(0)
            .Height = PICH(0)
         End With
         DisplayArray PIC_Save, PICW(0), PICH(0), ARRRes(), 0, 0, -1 'Mul
         SavePicture PIC_Save.Image, FileSpec$
         PIC_Save.Picture = LoadPicture
         With PIC_Save
            .Width = 6
            .Height = 6
         End With
         
      End If
   Case 1   ' Exit
      Form_Unload 1
      End
   End Select
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Unload Me
   End
End Sub


