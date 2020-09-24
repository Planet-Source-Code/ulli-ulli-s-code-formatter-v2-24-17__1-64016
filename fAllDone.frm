VERSION 5.00
Begin VB.Form fAllDone 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4800
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox ckPrint 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00800080&
      Height          =   390
      Left            =   255
      TabIndex        =   4
      ToolTipText     =   "Delete printed pages"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton btClose 
      BackColor       =   &H0000C0C0&
      Caption         =   "Close"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1830
      MaskColor       =   &H00000000&
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   1035
      Width           =   1140
   End
   Begin VB.Timer tmrTgl 
      Interval        =   1333
      Left            =   4185
      Top             =   195
   End
   Begin VB.Image imgCWFl 
      Height          =   660
      Index           =   0
      Left            =   240
      Picture         =   "fAllDone.frx":0000
      Top             =   255
      Width           =   585
   End
   Begin VB.Image imgCWFl 
      Appearance      =   0  '2D
      Height          =   825
      Index           =   1
      Left            =   3615
      MouseIcon       =   "fAllDone.frx":14E2
      MousePointer    =   99  'Benutzerdefiniert
      Picture         =   "fAllDone.frx":17EC
      Top             =   645
      Width           =   1020
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "© 2000/2006 UMGEDV"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   150
      Index           =   2
      Left            =   3375
      TabIndex        =   3
      Top             =   1485
      Width           =   1350
   End
   Begin VB.Shape sh 
      BackColor       =   &H00000000&
      BorderColor     =   &H0000C0C0&
      BorderStyle     =   2  'Strich
      BorderWidth     =   3
      Height          =   1695
      Left            =   15
      Top             =   30
      Width           =   4770
   End
   Begin VB.Label lbl 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "All requested Components have been processed."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   525
      Index           =   0
      Left            =   1005
      TabIndex        =   1
      Top             =   225
      Width           =   2745
   End
   Begin VB.Label lbl 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "All requested Components have been processed."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   1
      Left            =   990
      TabIndex        =   2
      Top             =   210
      Width           =   2745
   End
End
Attribute VB_Name = "fAllDone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z 'we're 32 bit

Private Tooltips    As Collection
Private Toggle      As Boolean

Private Sub btClose_Click()

    Hide

End Sub

Private Sub btClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MPIcon.SetPointerIcon RightHand

End Sub

Private Sub Form_Load()

    SetButtonForeColor btClose, &HF8F8&, AlignThreeD
    tmrTgl_Timer
    btClose.ToolTipText = "Good bye " & UserName
    Set Tooltips = CreateTooltips(Me)
    lbl(2) = Copyright

End Sub

Private Sub Form_Paint()

    If PrintLineLen Then
        ckPrint.Visible = (Printer.CurrentY > 0)
        ckPrint.Caption = "Kill Print Doc " & Format$(Printer.Page) & OneOrMany(" Page", Printer.Page)
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    UnsetButtonForeColor btClose
    KillDoc = (ckPrint = vbChecked)
    Set Tooltips = Nothing

End Sub

Private Sub tmrTgl_Timer()

    Toggle = Not Toggle
    imgCWFl(1).ToolTipText = IIf(Toggle, "Superior Code Contest Winner.", "www.Planet-Source-Code.com")

End Sub

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 5  Code: 49  Total: 54 Lines
':) CommentOnly: 2 (3,7%)  Commented: 1 (1,9%)  Filled: 34 (63%)  Empty: 20 (37%)  Max Logic Depth: 2
