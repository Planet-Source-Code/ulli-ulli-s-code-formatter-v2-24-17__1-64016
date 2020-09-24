VERSION 5.00
Begin VB.Form fPreparing 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   525
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2100
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   Icon            =   "fPreparing.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   35
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer tmrAnimate 
      Interval        =   222
      Left            =   1650
      Top             =   15
   End
   Begin VB.Image imgFormat 
      Height          =   480
      Left            =   45
      Picture         =   "fPreparing.frx":000C
      Top             =   30
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSort 
      Height          =   360
      Left            =   180
      Picture         =   "fPreparing.frx":0316
      Top             =   90
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgBroom 
      Height          =   360
      Left            =   150
      Picture         =   "fPreparing.frx":07D8
      Top             =   45
      Width           =   225
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   630
      TabIndex        =   0
      Top             =   165
      Width           =   75
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   615
      TabIndex        =   1
      Top             =   150
      Width           =   75
   End
End
Attribute VB_Name = "fPreparing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z 'we're 32 bit

Private Sub Form_Load()

    SetWindowPos hWnd, SWP_TOPMOST, 0, 0, 0, 0, SWP_COMBINED

End Sub

Private Sub lbl_Change(Index As Integer)

    lbl(1) = lbl(0)

End Sub

Private Sub tmrAnimate_Timer()

    imgBroom.Move imgBroom.Left Xor 3, imgBroom.Top Xor 1

End Sub

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 2  Code: 21  Total: 23 Lines
':) CommentOnly: 2 (8,7%)  Commented: 1 (4,3%)  Filled: 13 (56,5%)  Empty: 10 (43,5%)  Max Logic Depth: 1
