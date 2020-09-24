VERSION 5.00
Begin VB.Form fSearch 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Search"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4620
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox ckSelected 
      BackColor       =   &H00E0E0E0&
      Caption         =   "In &Selection"
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   990
      TabIndex        =   6
      Top             =   1275
      Width           =   1155
   End
   Begin VB.CommandButton btCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   3360
      TabIndex        =   5
      Top             =   1080
      Width           =   1005
   End
   Begin VB.CommandButton btSrch 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   3360
      TabIndex        =   4
      Top             =   630
      Width           =   1005
   End
   Begin VB.CheckBox ckCase 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Match Case"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   990
      TabIndex        =   3
      Top             =   945
      Width           =   1215
   End
   Begin VB.CheckBox ckWhole 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Whole Word"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   990
      TabIndex        =   2
      Top             =   630
      Width           =   1215
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   975
      TabIndex        =   1
      Top             =   165
      Width           =   3390
   End
   Begin VB.Image imgSrch 
      BorderStyle     =   1  'Fest Einfach
      Height          =   330
      Left            =   225
      Top             =   780
      Width           =   330
   End
   Begin VB.Label lbl 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Find:"
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   345
   End
End
Attribute VB_Name = "fSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z 'we're 32 bit

Private Sub btCancel_Click()

    Tag = "C"
    Hide

End Sub

Private Sub btSrch_Click()

    Tag = "O"
    Hide

End Sub

Private Sub Form_Load()

    imgSrch.Picture = fCopy.imgList.ListImages(7).Picture

End Sub

Private Sub txtSearch_Change()

    btSrch.Enabled = Len(txtSearch)

End Sub

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 2  Code: 29  Total: 31 Lines
':) CommentOnly: 2 (6,5%)  Commented: 1 (3,2%)  Filled: 18 (58,1%)  Empty: 13 (41,9%)  Max Logic Depth: 1
