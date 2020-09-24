VERSION 5.00
Begin VB.Form fSummary 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "S"
   ClientHeight    =   2040
   ClientLeft      =   3645
   ClientTop       =   4635
   ClientWidth     =   5115
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   Icon            =   "fSummary.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox ckPrint 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00800080&
      Height          =   435
      Left            =   540
      TabIndex        =   5
      ToolTipText     =   "Delete printed pages"
      Top             =   1425
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton btStop 
      BackColor       =   &H008080FF&
      Cancel          =   -1  'True
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      Style           =   1  'Grafisch
      TabIndex        =   4
      ToolTipText     =   "Skip remaning Panels"
      Top             =   1425
      Width           =   465
   End
   Begin VB.CommandButton btViewStruc 
      Height          =   435
      Left            =   3495
      Picture         =   "fSummary.frx":000C
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Show Module Structure"
      Top             =   1425
      UseMaskColor    =   -1  'True
      Width           =   1230
   End
   Begin VB.CommandButton btOK 
      Default         =   -1  'True
      Height          =   435
      Left            =   1942
      Picture         =   "fSummary.frx":02F2
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Close Window"
      Top             =   1425
      UseMaskColor    =   -1  'True
      Width           =   1230
   End
   Begin VB.Image imgSmiley 
      Height          =   480
      Left            =   240
      Picture         =   "fSummary.frx":0478
      Top             =   405
      Width           =   480
   End
   Begin VB.Label lblComplaints 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "   "
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1065
      TabIndex        =   1
      ToolTipText     =   "Scan Results"
      Top             =   600
      Width           =   135
   End
   Begin VB.Label lblSummary 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   1065
      TabIndex        =   0
      Top             =   225
      Width           =   75
   End
   Begin VB.Image imgSerious 
      Height          =   480
      Left            =   240
      Picture         =   "fSummary.frx":10BA
      Top             =   645
      Width           =   480
   End
End
Attribute VB_Name = "fSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z 'we're 32 bit

Private Tooltips       As Collection
Private IsVirgin       As Boolean
Public ForCompiling    As Boolean

Private Sub btOK_Click()

    Unload fStruc
    Hide

End Sub

Private Sub btStop_Click()

    BreakLoop = True
    btOK_Click

End Sub

Private Sub btViewStruc_Click()

    Unload fProgress
    fStruc.Show vbModal

End Sub

Private Sub ckPrint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    btOK.SetFocus

End Sub

Private Sub Form_Activate()

    If IsVirgin Then
        MessageBeep IIf(imgSmiley.Visible, vbInformation, vbCritical)
        Sleep 10
        IsVirgin = False
    End If

End Sub

Private Sub Form_Load()

    IsVirgin = True
    Caption = AppDetails
    SetButtonForeColor btStop, vbYellow, AlignThreeD
    On Error Resume Next
        ckPrint.Caption = "Kill Print Doc " & Format$(Printer.Page) & OneOrMany(" Page", Printer.Page)
    On Error GoTo 0
    Set Tooltips = CreateTooltips(Me)

End Sub

Private Sub Form_Paint()

    If PrintLineLen Then
        ckPrint.Visible = (Printer.CurrentY > 0) And Not ForCompiling
      Else 'PRINTLINELEN = FALSE/0
        ckPrint.Visible = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    UnsetButtonForeColor btStop
    If Not ForCompiling Then
        KillDoc = (ckPrint = vbChecked)
    End If
    Set Tooltips = Nothing

End Sub

Private Sub lblComplaints_Change()

  'resize and move form and reposition controls

    Width = lblComplaints.Left + lblComplaints.Width + ScaleX(30, vbPixels, vbTwips)
    btOK.Move (Width - btOK.Width) * IIf((StrucRequested Or PrintLineLen) And Not ForCompiling, 0.25, 0.5), lblComplaints.Top + lblComplaints.Height + 225
    btViewStruc.Move (Width - btViewStruc.Width) * 0.75, btOK.Top
    btViewStruc.Enabled = True
    btViewStruc.Visible = StrucRequested
    btStop.Move Width - btStop.Width - ScaleX(16, vbPixels, vbTwips), btOK.Top
    ckPrint.Move btViewStruc.Left, btViewStruc.Top
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2, Width, btOK.Top + btOK.Height + ScaleX(30, vbPixels, vbTwips)

End Sub

Public Property Let Serious(TruthVal As Boolean)

    imgSmiley.Visible = Not TruthVal
    imgSerious.Visible = TruthVal

End Property

Public Property Let StopButtonVisible(TruthVal As Boolean)

    btStop.Visible = TruthVal
    If TruthVal Then
        btOK.ToolTipText = "Continue with next Pane"
        UpdateTooltip Tooltips, btOK
    End If

End Property

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 6  Code: 104  Total: 110 Lines
':) CommentOnly: 3 (2,7%)  Commented: 2 (1,8%)  Filled: 74 (67,3%)  Empty: 36 (32,7%)  Max Logic Depth: 2
