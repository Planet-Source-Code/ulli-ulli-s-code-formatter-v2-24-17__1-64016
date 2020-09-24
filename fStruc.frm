VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fStruc 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   ForeColor       =   &H00E0E0E0&
   Icon            =   "fStruc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   373
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   431
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton btUnload 
      Caption         =   "&Unload"
      Height          =   330
      Left            =   5325
      TabIndex        =   4
      Top             =   90
      Width           =   1050
   End
   Begin VB.CommandButton btExpNext 
      Caption         =   "&Expand this"
      Height          =   330
      Left            =   1185
      TabIndex        =   3
      Top             =   90
      Width           =   1050
   End
   Begin MSComctlLib.TreeView tvwStruc 
      Height          =   4875
      Left            =   60
      TabIndex        =   2
      Top             =   540
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   8599
      _Version        =   393217
      Indentation     =   503
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton btExpCol 
      Caption         =   "&Collapse all"
      Height          =   330
      Index           =   1
      Left            =   2310
      TabIndex        =   1
      Top             =   90
      Width           =   1050
   End
   Begin VB.CommandButton btExpCol 
      Caption         =   "E&xpand all"
      Height          =   330
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   90
      Width           =   1050
   End
   Begin VB.Label lblFoot 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Click on text to position code"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   165
      Left            =   90
      TabIndex        =   5
      Top             =   5415
      Width           =   1815
   End
End
Attribute VB_Name = "fStruc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z 'we're 32 bit

Private MarginX, MarginY
Private Tooltips As Collection

Private Sub btExpCol_Click(Index As Integer)

  Dim i

    With tvwStruc
        MPIcon = vbHourglass
        .Visible = False
        For i = 2 To .Nodes.Count
            .Nodes(i).Expanded = (Index = 0)
        Next i
        .Nodes(1).Expanded = True
        .Visible = True
        .Nodes(1).EnsureVisible
        .SetFocus
    End With 'TVWSTRUC
    MPIcon = vbDefault

End Sub

Private Sub btExpNext_Click()

  Dim Node As MSComctlLib.Node

    With tvwStruc
        If .SelectedItem Is Nothing Then
            MessageBeep vbQuestion
          ElseIf .SelectedItem.Children = 0 Then 'NOT .SELECTEDITEM...
            MessageBeep vbQuestion
          Else 'NOT .SELECTEDITEM.CHILDREN...
            .Visible = False
            .SelectedItem.Child.FirstSibling.EnsureVisible
            Set Node = .SelectedItem.Child.LastSibling
            Do
                Node.Selected = (Node.Children <> 0)
                Node.Expanded = True
                Set Node = .Nodes(Node.Index).Previous
            Loop Until Node Is Nothing
            .Visible = True
        End If
        .SetFocus
    End With 'TVWSTRUC

End Sub

Private Sub btUnload_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    MarginX = ScaleWidth - tvwStruc.Width
    MarginY = ScaleHeight - tvwStruc.Height
    Set Tooltips = CreateTooltips(Me)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = vbFormCode Then
        fSummary.btViewStruc.Enabled = False
      Else 'NOT UNLOADMODE...
        Cancel = True
        Hide
    End If

End Sub

Private Sub Form_Resize()

    On Error Resume Next
        tvwStruc.Width = ScaleWidth - MarginX
        tvwStruc.Height = ScaleHeight - MarginY
        btUnload.Left = ScaleWidth - MarginX - btUnload.Width + 2
        lblFoot.Move 6, ScaleHeight - 12
    On Error GoTo 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set Tooltips = Nothing

End Sub

Private Sub tvwStruc_NodeClick(ByVal Node As MSComctlLib.Node)

  'stolen from Roger

    VBInstance.ActiveCodePane.TopLine = Val(Node.Tag)

End Sub

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 5  Code: 97  Total: 102 Lines
':) CommentOnly: 3 (2,9%)  Commented: 6 (5,9%)  Filled: 73 (71,6%)  Empty: 29 (28,4%)  Max Logic Depth: 4
