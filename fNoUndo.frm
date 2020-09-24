VERSION 5.00
Begin VB.Form fNoUndo 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Festes Werkzeugfenster
   ClientHeight    =   1950
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3780
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox btOK 
      BackColor       =   &H00C0C000&
      Caption         =   "OK"
      Height          =   480
      Left            =   1478
      Style           =   1  'Grafisch
      TabIndex        =   4
      Top             =   1320
      Width           =   825
   End
   Begin VB.CommandButton btOKAb 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Index           =   1
      Left            =   2085
      Picture         =   "fNoUndo.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Cancel this run"
      Top             =   1305
      Width           =   1065
   End
   Begin VB.CommandButton btOKAb 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Index           =   0
      Left            =   630
      Picture         =   "fNoUndo.frx":0222
      Style           =   1  'Grafisch
      TabIndex        =   1
      ToolTipText     =   "Continue with next component"
      Top             =   1305
      Width           =   1065
   End
   Begin VB.Label lbName 
      BackColor       =   &H00808000&
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
      ForeColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   15
      Width           =   3795
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   165
      Picture         =   "fNoUndo.frx":0444
      Top             =   555
      Width           =   480
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "There was no previous action which could be undone."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   615
      Width           =   2535
   End
End
Attribute VB_Name = "fNoUndo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z 'we're 32 bits
Private Tooltips As Collection

Private Sub btOK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    btOK = vbUnchecked
    btOkAb_Click 1

End Sub

Private Sub btOkAb_Click(Index As Integer)

    Tag = Index
    Hide

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Select Case Chr$(KeyAscii)
      Case "C", "c", Chr$(vbKeyReturn)
        btOkAb_Click 0
      Case "A", "a", Chr$(vbKeyEscape)
        btOkAb_Click 1
    End Select

End Sub

Private Sub Form_Load()

    Set Tooltips = CreateTooltips(Me)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set Tooltips = Nothing

End Sub

Public Property Let IsLastPanel(TruthVal As Boolean)

    btOKAb(0).Visible = Not TruthVal
    btOKAb(1).Visible = Not TruthVal
    btOK.Visible = TruthVal

End Property

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 3  Code: 48  Total: 51 Lines
':) CommentOnly: 2 (3,9%)  Commented: 1 (2%)  Filled: 32 (62,7%)  Empty: 19 (37,3%)  Max Logic Depth: 2
