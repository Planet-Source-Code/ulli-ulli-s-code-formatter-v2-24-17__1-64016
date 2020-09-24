VERSION 5.00
Begin VB.Form fUndoList 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Undo List"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3915
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox lstUndone 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808000&
      Height          =   1425
      Left            =   375
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Undo List"
      Top             =   750
      Width           =   3180
   End
   Begin VB.CommandButton btCloseUndo 
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
      Left            =   1387
      MaskColor       =   &H00000000&
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   2415
      Width           =   1140
   End
   Begin VB.Image imgCWFl 
      Height          =   660
      Left            =   180
      Picture         =   "fUndoList.frx":0000
      Top             =   165
      Width           =   585
   End
   Begin VB.Label lb 
      BackStyle       =   0  'Transparent
      Caption         =   "These modules were affected by the Undo function:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   975
      TabIndex        =   2
      Top             =   150
      Width           =   2580
   End
End
Attribute VB_Name = "fUndoList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z 'we're 32 bits

Private Tooltips As Collection

Private Sub btCloseUndo_Click()

    Unload Me

End Sub

Private Sub Form_Activate()

    With lstUndone
        .AddItem "[ " & .ListCount & OneOrMany(" Module", .ListCount) & " ]"
    End With 'LSTUNDONE

End Sub

Private Sub Form_Load()

    SetButtonForeColor btCloseUndo, &HF8F8&, AlignThreeD
    btCloseUndo.ToolTipText = "Good bye " & UserName
    Set Tooltips = CreateTooltips(Me)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    UnsetButtonForeColor btCloseUndo
    Set Tooltips = Nothing

End Sub

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 4  Code: 32  Total: 36 Lines
':) CommentOnly: 2 (5,6%)  Commented: 2 (5,6%)  Filled: 22 (61,1%)  Empty: 14 (38,9%)  Max Logic Depth: 2
