VERSION 5.00
Begin VB.Form fHelp 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Help"
   ClientHeight    =   4320
   ClientLeft      =   1800
   ClientTop       =   3600
   ClientWidth     =   6075
   ForeColor       =   &H00E0E0E0&
   Icon            =   "fHelp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton btShutup 
      Caption         =   "&Shut up"
      Height          =   345
      Left            =   3405
      TabIndex        =   2
      Top             =   3870
      Width           =   1125
   End
   Begin VB.CommandButton btSpeak 
      Height          =   345
      Left            =   1200
      TabIndex        =   1
      Top             =   3870
      Width           =   1125
   End
   Begin VB.TextBox txtHelp 
      BackColor       =   &H80000018&
      ForeColor       =   &H80000017&
      Height          =   3600
      Left            =   68
      MousePointer    =   1  'Pfeil
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "fHelp.frx":000C
      Top             =   75
      Width           =   5910
   End
End
Attribute VB_Name = "fHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z 'we're 32 bit

Private Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long

Private MyWidth             As Long
Private HeightDiff          As Long
Private Speaking            As Boolean
Private Paused              As Boolean
Private WithEvents Voice    As SpVoice
Attribute Voice.VB_VarHelpID = -1

Private Sub btShutup_Click()

    txtHelp.SetFocus
    If Speaking And Not Paused Then
        Paused = True
        Voice.Pause
        btSpeak.Caption = "&Resume"
    End If

End Sub

Private Sub btSpeak_Click()

    txtHelp.SetFocus
    Paused = False
    If Speaking Then
        Voice.Resume
      Else 'SPEAKING = FALSE/0
        Speaking = True
        Voice.Speak txtHelp, SVSFlagsAsync
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
      Case vbKeyEscape
        Unload Me
      Case Asc("s"), Asc("S")
        btShutup_Click
      Case Asc("r"), Asc("R")
        btSpeak_Click
    End Select

End Sub

Private Sub Form_Load()

    MyWidth = Width
    HeightDiff = Height - txtHelp.Height
    Set Voice = New SpVoice
    Voice_EndStream 0, 0

End Sub

Private Sub Form_Resize()

    If Width <> MyWidth Then
        ReleaseCapture
        Width = MyWidth
    End If
    On Error Resume Next
        txtHelp.Height = Height - HeightDiff
    On Error GoTo 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Voice.Pause
    Set Voice = Nothing
    Set fHelp = Nothing

End Sub

Private Sub txtHelp_GotFocus()

    HideCaret txtHelp.hWnd

End Sub

Private Sub txtHelp_KeyPress(KeyAscii As Integer)

    KeyAscii = 0

End Sub

Private Sub Voice_EndStream(ByVal StreamNumber As Long, ByVal StreamPosition As Variant)

    btSpeak.Caption = "&Read this"
    Speaking = False
    On Error Resume Next
        btSpeak.SetFocus
    On Error GoTo 0

End Sub

Private Sub Voice_Word(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal CharacterPosition As Long, ByVal Length As Long)

    With txtHelp
        .SelStart = CharacterPosition
        .SelLength = Length
    End With 'TXTHELP

End Sub

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 10  Code: 100  Total: 110 Lines
':) CommentOnly: 2 (1,8%)  Commented: 3 (2,7%)  Filled: 77 (70%)  Empty: 33 (30%)  Max Logic Depth: 2
