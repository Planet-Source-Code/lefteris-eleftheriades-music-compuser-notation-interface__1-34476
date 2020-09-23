VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Music Maker"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Load"
      Height          =   495
      Left            =   5940
      TabIndex        =   41
      Top             =   660
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      Height          =   495
      Left            =   5940
      TabIndex        =   40
      Top             =   150
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   3210
      TabIndex        =   37
      Top             =   -60
      Width           =   945
      Begin VB.OptionButton Option3 
         Height          =   345
         Index           =   1
         Left            =   30
         Picture         =   "Form1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   120
         Value           =   -1  'True
         Width           =   435
      End
      Begin VB.OptionButton Option3 
         Height          =   345
         Index           =   0
         Left            =   480
         Picture         =   "Form1.frx":047A
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   120
         Width           =   435
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1275
      Index           =   1
      Left            =   5190
      Picture         =   "Form1.frx":078C
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   36
      Top             =   450
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Picture10 
      Height          =   2775
      Left            =   0
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   443
      TabIndex        =   34
      Top             =   1860
      Width           =   6705
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   0
         ScaleHeight     =   177
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   447
         TabIndex        =   35
         Top             =   30
         Width           =   6705
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "7"
      Enabled         =   0   'False
      Height          =   345
      Index           =   6
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3420
      Width           =   345
   End
   Begin VB.OptionButton Option2 
      Caption         =   "6"
      Enabled         =   0   'False
      Height          =   345
      Index           =   5
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3090
      Width           =   345
   End
   Begin VB.OptionButton Option2 
      Caption         =   "5"
      Height          =   345
      Index           =   4
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2760
      Width           =   345
   End
   Begin VB.OptionButton Option2 
      Caption         =   "4"
      Height          =   345
      Index           =   3
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2430
      Width           =   345
   End
   Begin VB.OptionButton Option2 
      Caption         =   "3"
      Height          =   345
      Index           =   2
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2100
      Width           =   345
   End
   Begin VB.OptionButton Option2 
      Caption         =   "2"
      Height          =   345
      Index           =   1
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1770
      Width           =   345
   End
   Begin VB.OptionButton Option2 
      Caption         =   "1"
      Height          =   345
      Index           =   0
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1440
      Width           =   345
   End
   Begin VB.CommandButton Command2 
      Caption         =   "b"
      Height          =   375
      Left            =   2760
      TabIndex        =   26
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "#"
      Height          =   375
      Left            =   2400
      TabIndex        =   25
      Top             =   0
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   6
      Left            =   390
      TabIndex        =   24
      Text            =   " Violin"
      Top             =   1530
      Width           =   1995
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   5
      Left            =   2970
      TabIndex        =   23
      Text            =   "Rock Guitar"
      Top             =   1170
      Width           =   1995
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   4
      Left            =   390
      TabIndex        =   22
      Text            =   "Sax"
      Top             =   1170
      Width           =   1995
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   3
      Left            =   2970
      TabIndex        =   21
      Text            =   "Guitar"
      Top             =   810
      Width           =   1995
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      Left            =   390
      TabIndex        =   20
      Text            =   "Drums"
      Top             =   810
      Width           =   1995
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      Left            =   2970
      TabIndex        =   19
      Text            =   "Acoustic Piano"
      Top             =   450
      Width           =   1995
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   390
      TabIndex        =   18
      Text            =   "Acoustic Piano"
      Top             =   450
      Width           =   1995
   End
   Begin VB.CheckBox Check1 
      Caption         =   "7"
      Height          =   375
      Index           =   6
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1470
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "6"
      Height          =   375
      Index           =   5
      Left            =   2550
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1140
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "5"
      Height          =   375
      Index           =   4
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1110
      Value           =   1  'Checked
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "4"
      Height          =   375
      Index           =   3
      Left            =   2550
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   780
      Value           =   1  'Checked
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "3"
      Height          =   375
      Index           =   2
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   750
      Value           =   1  'Checked
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "2"
      Height          =   375
      Index           =   1
      Left            =   2550
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   420
      Value           =   1  'Checked
      Width           =   375
   End
   Begin VB.CheckBox Check1 
      Caption         =   "1"
      Height          =   375
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   390
      Value           =   1  'Checked
      Width           =   375
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   195
      Left            =   30
      Max             =   100
      TabIndex        =   10
      Top             =   4650
      Width           =   6705
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   2355
      TabIndex        =   1
      Top             =   0
      Width           =   2355
      Begin VB.OptionButton Option1 
         Height          =   375
         Index           =   5
         Left            =   1470
         Picture         =   "Form1.frx":092A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Height          =   375
         Index           =   4
         Left            =   1080
         Picture         =   "Form1.frx":0AB0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Height          =   375
         Index           =   3
         Left            =   750
         Picture         =   "Form1.frx":0DC2
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Height          =   375
         Index           =   2
         Left            =   360
         Picture         =   "Form1.frx":10D4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Height          =   375
         Index           =   1
         Left            =   0
         Picture         =   "Form1.frx":1576
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   375
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   390
         Left            =   2160
         Max             =   20
         Min             =   5
         TabIndex        =   4
         Top             =   0
         Value           =   10
         Width           =   195
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H8000000E&
         Height          =   375
         Left            =   1830
         ScaleHeight     =   315
         ScaleWidth      =   255
         TabIndex        =   2
         Top             =   0
         Width           =   315
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   30
            TabIndex        =   3
            Text            =   "10"
            Top             =   60
            Width           =   225
         End
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   2670
      Picture         =   "Form1.frx":1A18
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   0
      Top             =   1500
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const LR_LOADFROMFILE = &H10
Private Const IMAGE_BITMAP = 0
Private Const IMAGE_ICON = 1
Private Const IMAGE_CURSOR = 2
Private Const IMAGE_ENHMETAFILE = 3

Const SRCCOPY = &HCC0020
Const SRCAND = &H8800C6
Const SRCINVERT = &H660046
'                  Line(Y), Coloum(X)
Dim PentagramLines(1 To 40, 1 To 150) As Integer
Dim Selection As Integer
Dim MouseBlink As Boolean
Dim OriginalPentagramDesignerCursor&


Dim curDevice
Dim hmidi
Dim rc
Dim channel
Dim volume
Dim IsDrum(7) As Boolean, DrumNum%(7)
Dim CCount As Long

Private Sub Check1_Click(Index As Integer)
 Option2(Index).Enabled = Check1(Index).Value
 Picture1.SetFocus
End Sub

Private Sub Combo1_Click(Index As Integer)
Dim Instrument%

Instrument = Combo1(Index).ListIndex

'channel = CurrentTrack
ChangeInstrument Instrument

'play short sample of instrument
    
    StartNote 67
    DelayTimer 250
    StopNote 67
End Sub

Private Sub Command3_Click()
  Dim OutputS$
  For LineY& = 1 To 30
     For ColoumX = 1 To 150
       A = Chr(PentagramLines(LineY&, ColoumX))
       OutputS = OutputS & A
     Next ColoumX
  Next LineY
  
  Open "C:\Windows\Desktop\MusicSav.txt" For Binary As #1
       Put #1, , OutputS
  Close #1
End Sub

Private Sub Command4_Click()
  Dim InputS$
  Open "C:\Windows\Desktop\MusicSav.txt" For Binary As #1
       InputS = String(4500, 0)
       Get #1, , InputS
  Close #1
  MsgBox InputS$
End Sub

Private Sub Form_Load()
Dim LineY&, ColoumX&
Dim A&
  Selection = 3
  OriginalPentagramDesignerCursor& = SetObjectCursor(App.Path & "\Animated.ani", Picture1.hwnd)
  DisplayPentagramNotes
   
   
   curDevice = 0
   rc = midiOutClose(hmidi)
   rc = midiOutOpen(hmidi, curDevice, 0, 0, 0)
   channel = 0
   ' Set volume range
   volume = 127
   For A = 100 To 227 'load up instrument names
     Combo1(0).AddItem A - 100 & " " & LoadResString(A)
   Next A
   For A = 35 To 81 'load up drum names
     Combo1(0).AddItem A + 93 & " " & LoadResString(A)
   Next A
   Combo1(0).ListIndex = 0
End Sub

Sub DisplayPentagramNotes()
  Dim LineY&, ColoumX&
  'Picture1.Cls
  Pentagram 70
  For LineY& = 1 To 30
     For ColoumX = 1 To 50
       Draw_Note ColoumX * 16 + 30, (LineY& * 5) + 13, PentagramLines(LineY&, ColoumX + HScroll1.Value)
     Next ColoumX
  Next LineY
  Picture1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
'unSetObjectCursor Me.hwnd, OriginalPentagramDesignerCursor&
rc = midiOutClose(hmidi)
unSetObjectCursor Picture1.hwnd, OriginalPentagramDesignerCursor&
End Sub

Private Sub HScroll1_Change()
'Picture1.Left = 0 - (HScroll1.Value * 2)
DisplayPentagramNotes
End Sub

Private Sub HScroll1_Scroll()
'Picture1.Left = 0 - (HScroll1.Value * 2)
DisplayPentagramNotes
End Sub

Private Sub Image1_Click()
End Sub

Private Sub Option1_Click(Index As Integer)
 Picture1.SetFocus
 Selection = Index
End Sub

Private Sub Option2_Click(Index As Integer)
 Picture1.SetFocus
End Sub

Sub Draw_Note(X&, Y&, ID%)
 Select Case ID%
  Case 1:  BitBlt Picture1.hdc, X& - 7, Y& - 18, 20, 30, Picture2(0).hdc, 66, 0, SRCAND
  Case 2:  BitBlt Picture1.hdc, X& - 5, Y& - 18, 20, 30, Picture2(0).hdc, 0, 0, SRCAND
  Case 3:  BitBlt Picture1.hdc, X& - 5, Y& - 18, 14, 30, Picture2(0).hdc, 23, 0, SRCAND
  Case 4:  BitBlt Picture1.hdc, X& - 7, Y& - 18, 14, 30, Picture2(0).hdc, 36, 0, SRCAND
  Case 5:  BitBlt Picture1.hdc, X& - 7, Y& - 18, 14, 30, Picture2(0).hdc, 52, 0, SRCAND
 End Select
End Sub

Sub Pentagram(Y&)
Dim Distance&
 Distance& = Val(Text1.Text)
 Picture1.Cls
 BitBlt Picture1.hdc, 5, Y - 20, 32, 85, Picture2(1).hdc, 0, 0, SRCAND
 Picture1.Line (0, Y&)-(Picture1.ScaleWidth, Y&)
 Picture1.Line (0, Y& + Distance&)-(Picture1.ScaleWidth, Y& + Distance&)
 Picture1.Line (0, Y& + Distance& * 2)-(Picture1.ScaleWidth, Y& + Distance& * 2)
 Picture1.Line (0, Y& + Distance& * 3)-(Picture1.ScaleWidth, Y& + Distance& * 3)
 Picture1.Line (0, Y& + Distance& * 4)-(Picture1.ScaleWidth, Y& + Distance& * 4)
 Picture1.Refresh
End Sub

Private Sub Option3_Click(Index As Integer)
If Index = 0 Then
   Call SetObjectCursor(App.Path & "\pencil.cur", Picture1.hwnd)
Else
   Call SetObjectCursor(App.Path & "\Animated.ani", Picture1.hwnd)
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim LineY&, ColoumX&
On Error Resume Next
LineY& = Round((Y - 13) / 5)
ColoumX& = Round((X - 30) / 16)
If PentagramLines(LineY&, ColoumX& + HScroll1.Value) = 0 Then
   PentagramLines(LineY&, ColoumX& + HScroll1.Value) = Selection
Else
   PentagramLines(LineY&, ColoumX& + HScroll1.Value) = 0
End If
DisplayPentagramNotes
    'MsgBox 46 + LineY&
    StartNote 46 + LineY&
    DelayTimer 110 * Selection
    StopNote 46 + LineY&
End Sub

Private Sub Text1_Change()
 Pentagram 30
End Sub

Sub DelayTimer(Milisecs As Long)
Dim Tick
    Tick = GetTickCount
    Do While Tick + Milisecs > GetTickCount
        DoEvents
    Loop
End Sub
Private Sub StartNote(Index As Integer)
Dim Flip
Dim TempChannel
Dim midimsg
If IsDrum(channel) Then
    Flip = DrumNum(channel)
    TempChannel = 9
Else
    Flip = 127 - Index 'notes recorded on grid are 127 - midi number
    TempChannel = channel
End If

midimsg = &H90 + ((Flip) * &H100) + (volume * &H10000) + TempChannel
midiOutShortMsg hmidi, midimsg
End Sub

Private Sub StopNote(Index As Integer)
Dim Flip
Dim TempChannel
Dim midimsg
If IsDrum(channel) Then
    Flip = DrumNum(channel)
    TempChannel = 9
Else
    Flip = 127 - Index 'notes recorded on grid are 127 - midi number
    TempChannel = channel
End If
   
midimsg = &H80 + ((Flip) * &H100) + TempChannel
midiOutShortMsg hmidi, midimsg
   
End Sub


Private Sub ChangeInstrument(Inst As Integer)

If Inst < 128 Then
    'melody instrument
    midiOutShortMsg hmidi, &HB0 + channel
    midiOutShortMsg hmidi, 32 * &H100 + &HB0 + channel
    midiOutShortMsg hmidi, Inst * &H100 + &HC0 + channel
    IsDrum(channel) = False
Else
    'percussion instrument
    IsDrum(channel) = True
    DrumNum(channel) = Inst - 93
End If
End Sub
