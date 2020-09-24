VERSION 5.00
Begin VB.Form fProgress 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3900
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   Icon            =   "fProgress.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   92
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picMenu 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      Height          =   240
      Left            =   2385
      Picture         =   "fProgress.frx":000C
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   4
      Top             =   315
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Frame fr 
      BackColor       =   &H00E0E0E0&
      Height          =   870
      Index           =   0
      Left            =   1110
      TabIndex        =   0
      Top             =   165
      Width           =   2670
      Begin VB.PictureBox picProgress 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Ausgefüllt
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   195
         ScaleHeight     =   13
         ScaleMode       =   0  'Benutzerdefiniert
         ScaleWidth      =   100
         TabIndex        =   3
         Top             =   480
         Width           =   2310
      End
      Begin VB.Label lbl 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Progress"
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
         Height          =   240
         Index           =   1
         Left            =   195
         TabIndex        =   2
         Top             =   135
         Width           =   960
      End
      Begin VB.Label lblXofY 
         Alignment       =   1  'Rechts
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1395
         TabIndex        =   6
         Top             =   165
         Width           =   1110
      End
      Begin VB.Label lbl 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Progress"
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
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   120
         Width           =   960
      End
      Begin VB.Label lblXofYShdw 
         Alignment       =   1  'Rechts
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1380
         TabIndex        =   7
         Top             =   150
         Width           =   1110
      End
   End
   Begin VB.Label lbCreating 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Creating Structure..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   2040
      TabIndex        =   9
      Top             =   1095
      Width           =   1740
   End
   Begin VB.Label lbPrinting 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Printing..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   2940
      TabIndex        =   8
      Top             =   1095
      Width           =   840
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "© 2000/2006  UMGEDV"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1140
      Width           =   1380
   End
   Begin VB.Image img 
      BorderStyle     =   1  'Fest Einfach
      Height          =   765
      Left            =   120
      Picture         =   "fProgress.frx":034E
      Top             =   240
      Width           =   825
   End
End
Attribute VB_Name = "fProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z 'we're 32 bit

Private Sub Form_Load()

    SetWindowPos hWnd, SWP_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE

End Sub

Private Sub Form_Paint()

    lbPrinting.Visible = PrintLineLen
    lbCreating.Visible = StrucRequested
    lbl(2) = Copyright

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = (UnloadMode <> vbFormCode)

End Sub

Private Sub lblXofY_Change()

    lblXofYShdw = lblXofY

End Sub

Public Property Let Percent(nuPercent As Long)

    picProgress.Line (0, picProgress.ScaleHeight / 2)-(nuPercent, picProgress.ScaleHeight), IIf(StrucRequested, &HC00000, IIf(PrintLineLen, &HC000C0, vbRed)), BF

End Property

Public Property Let Total(nuPercent As Long)

    picProgress.Line (0, 0)-(nuPercent, picProgress.ScaleHeight / 2 - 1), &HC000&, BF

End Property

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 2  Code: 41  Total: 43 Lines
':) CommentOnly: 2 (4,7%)  Commented: 1 (2,3%)  Filled: 24 (55,8%)  Empty: 19 (44,2%)  Max Logic Depth: 1
