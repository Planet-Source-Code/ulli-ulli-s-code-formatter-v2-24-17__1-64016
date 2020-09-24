VERSION 5.00
Begin VB.Form fStrFncts 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "String Functions returning a Variant"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4365
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ComboBox cboSel 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00404040&
      Height          =   315
      ItemData        =   "fStrFncts.frx":0000
      Left            =   2400
      List            =   "fStrFncts.frx":000D
      Style           =   2  'Dropdown-Liste
      TabIndex        =   1
      Top             =   120
      Width           =   1905
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   765
      Top             =   60
   End
   Begin VB.CommandButton btClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   420
      Left            =   1770
      TabIndex        =   3
      ToolTipText     =   "Close"
      Top             =   1905
      Width           =   825
   End
   Begin VB.ListBox lstStrFncts 
      BackColor       =   &H00F0FFFF&
      ForeColor       =   &H00006000&
      Height          =   1185
      ItemData        =   "fStrFncts.frx":0035
      Left            =   2400
      List            =   "fStrFncts.frx":0093
      Style           =   1  'Kontrollkästchen
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Select Functions"
      Top             =   555
      Width           =   1905
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Select"
      Height          =   195
      Index           =   1
      Left            =   1875
      TabIndex        =   0
      Top             =   180
      Width           =   450
   End
   Begin VB.Label lblAni 
      BackColor       =   &H00E0E0E0&
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080E0&
      Height          =   285
      Left            =   510
      TabIndex        =   6
      Top             =   150
      Width           =   135
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Rechts
      BackColor       =   &H00E0E0E0&
      Caption         =   "Chr"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000E0&
      Height          =   285
      Index           =   2
      Left            =   90
      TabIndex        =   5
      Top             =   150
      Width           =   420
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00F0FFFF&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   $"fStrFncts.frx":0175
      ForeColor       =   &H00007000&
      Height          =   1185
      Index           =   0
      Left            =   75
      TabIndex        =   4
      Top             =   555
      Width           =   2265
   End
End
Attribute VB_Name = "fStrFncts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z

Private Tooltips            As Collection
Private m, n
Private TopIndex
Private Internal            As Boolean
Private Const myBackColor   As Long = &HE0E0E0
Private Const ColorStep     As Long = &H80800
Private Const EndColor      As Long = &HE0 'this is the color we get when &H80800 is repetitively subtracted from MyBackColor
Private Const TiShort       As Long = 111
Private Const TiLong        As Long = 1111

Public Property Get AnySelected() As Boolean

    AnySelected = (lstStrFncts.SelCount <> 0)

End Property

Private Sub btClose_Click()

    tmr.Enabled = False
    Hide

End Sub

Private Sub cbosel_Change()

    cboSel_DropDown

End Sub

Private Sub cboSel_Click()

    If Not Internal Then
        Internal = True
        cboSel_DropDown
        With lstStrFncts
            Select Case cboSel.ListIndex
              Case 0, 1
                TopIndex = SendMessage(.hWnd, LB_GETTOPINDEX, 0&, ByVal 0&)
                For n = 0 To .ListCount - 1
                    .Selected(n) = (cboSel.ListIndex = 0)
                Next n
                .ListIndex = -1
                SendMessage .hWnd, LB_SETTOPINDEX, TopIndex, ByVal 0&
              Case 2
                TopIndex = SendMessage(.hWnd, LB_GETTOPINDEX, 0&, ByVal 0&)
                For n = 0 To .ListCount - 1
                    .Selected(n) = (Right$(.List(n), 1) = Spce)
                Next n
                .ListIndex = -1
                SendMessage .hWnd, LB_SETTOPINDEX, TopIndex, ByVal 0&
            End Select
            StringCboListIndex = .ListIndex
        End With 'LSTSTRFNCTS
        Internal = False
    End If
    tmr_Timer
    On Error Resume Next 'this can happen on load when we can't set the focus
        btClose.SetFocus
    On Error GoTo 0

End Sub

Private Sub cboSel_DropDown()

    If cboSel.ListCount = 4 Then
        cboSel.RemoveItem (3)
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    With cboSel
        Select Case KeyCode
          Case vbKeyF5
            .ListIndex = 0
          Case vbKeyF6
            .ListIndex = 1
          Case vbKeyF7
            .ListIndex = 2
        End Select
    End With 'CBOSEL

End Sub

Private Sub Form_Load()

    ReDim Preserve bStrFncts(0 To lstStrFncts.ListCount - 1)
    BackColor = myBackColor
    cboSel.ListIndex = StringCboListIndex
    Set Tooltips = CreateTooltips(Me)

End Sub

Private Sub Form_Paint()

    lblAni.ForeColor = myBackColor
    tmr_Timer

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set Tooltips = Nothing

End Sub

Public Sub LoadStrFncts()

    Internal = True
    With lstStrFncts
        TopIndex = SendMessage(.hWnd, LB_GETTOPINDEX, 0&, ByVal 0&)
        For n = 0 To .ListCount - 1
            .Selected(n) = bStrFncts(n)
        Next n
        .ListIndex = -1
        SendMessage .hWnd, LB_SETTOPINDEX, TopIndex, ByVal 0&
    End With 'LSTSTRFNCTS
    Internal = False

End Sub

Private Sub lstStrFncts_ItemCheck(Item As Integer)

    bStrFncts(Item) = lstStrFncts.Selected(Item)
    If Not Internal Then
        With cboSel
            Internal = True
            Select Case lstStrFncts.SelCount
              Case lstStrFncts.ListCount
                If .ListCount = 4 Then
                    .RemoveItem 3
                End If
                .ListIndex = 0 'all
              Case 0
                If .ListCount = 4 Then
                    .RemoveItem 3
                End If
                .ListIndex = 1 'none
              Case Else
                For m = 0 To lstStrFncts.ListCount - 1
                    If lstStrFncts.Selected(m) <> (Right$(lstStrFncts.List(m), 1) = Spce) Then
                        Exit For 'loop varying m
                    End If
                Next m
                If m = lstStrFncts.ListCount Then
                    If .ListCount = 4 Then
                        .RemoveItem 3
                    End If
                    .ListIndex = 2 'default
                  Else 'NOT M...
                    If .ListCount = 3 Then
                        .AddItem "Some"
                        .ListIndex = 3
                    End If
                End If
            End Select
            Internal = False
        End With 'CBOSEL
    End If

End Sub

Private Sub lstStrFncts_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    tmr_Timer

End Sub

Private Sub tmr_Timer()

    With lblAni
        If AnySelected Then
            tmr.Interval = TiShort
            tmr.Enabled = True
            If .ForeColor = EndColor Then
                .ForeColor = myBackColor
              Else 'NOT .FORECOLOR...
                .ForeColor = .ForeColor - ColorStep
                If .ForeColor = EndColor Then
                    tmr.Interval = TiLong
                End If
            End If
          Else 'ANYSELECTED = FALSE/0
            tmr.Enabled = False
            .ForeColor = myBackColor
        End If
    End With 'LBLANI

End Sub

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 12  Code: 184  Total: 196 Lines
':) CommentOnly: 2 (1%)  Commented: 13 (6,6%)  Filled: 155 (79,1%)  Empty: 41 (20,9%)  Max Logic Depth: 6
