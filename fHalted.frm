VERSION 5.00
Begin VB.Form fHalted 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   ClientHeight    =   1695
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   15  'Größenänderung alle
   ScaleHeight     =   1695
   ScaleWidth      =   2625
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ckHaltOnOff 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Halt Mode On/Off"
      Height          =   195
      Left            =   480
      MousePointer    =   1  'Pfeil
      TabIndex        =   2
      Top             =   660
      Value           =   1  'Aktiviert
      Width           =   1590
   End
   Begin VB.CommandButton btResume 
      Caption         =   "&Resume"
      Height          =   435
      Left            =   120
      MousePointer    =   1  'Pfeil
      TabIndex        =   0
      Top             =   1065
      Width           =   885
   End
   Begin VB.Shape Brdr 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Height          =   1650
      Left            =   30
      Top             =   30
      Width           =   2580
   End
   Begin VB.Image imgExcla 
      Height          =   480
      Left            =   30
      Picture         =   "fHalted.frx":0000
      Top             =   465
      Width           =   480
   End
   Begin VB.Label lbErrText 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   300
      Left            =   480
      MousePointer    =   3  'I-Cursor
      TabIndex        =   1
      Top             =   225
      Width           =   165
   End
End
Attribute VB_Name = "fHalted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btResume_Click()

    Hide

End Sub

Private Sub ckHaltOnOff_Click()

    StopRequested = (ckHaltOnOff = vbChecked)
    If ckHaltOnOff = vbChecked Then
        VBInstance.ActiveCodePane.Show 'reset focus on active code pane
    End If

End Sub

Private Sub Form_Activate()

  'adjust size to fit text

    With lbErrText
        Width = (.Width + .Left + .Left)
    End With 'LBERRTEXT
    With btResume
        .Left = (Width - .Width) / 2
    End With 'BTRESUME
    With Brdr
        .Height = Height - (.BorderWidth + 1) * Screen.TwipsPerPixelY
        .Width = Width - (.BorderWidth + 1) * Screen.TwipsPerPixelX
    End With 'BRDR

End Sub

Private Sub Form_Load()

    SetWindowPos hWnd, SWP_TOPMOST, 0, 0, 0, 0, SWP_COMBINED
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 3

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'grab form to move

    ReleaseCapture 'release the Mouse
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0& 'non-client area button down (in caption)
    ckHaltOnOff_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)

    SetWindowPos fPreparing.hWnd, SWP_TOPMOST, 0, 0, 0, 0, SWP_COMBINED

End Sub

Private Sub lbErrText_Click()

    Form_MouseDown 0, 0, 0, 0

End Sub

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 1  Code: 64  Total: 65 Lines
':) CommentOnly: 4 (6,2%)  Commented: 6 (9,2%)  Filled: 41 (63,1%)  Empty: 24 (36,9%)  Max Logic Depth: 2
