VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fPrintFont 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "IDE-Font not printable..."
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   ForeColor       =   &H00E0E0E0&
   Icon            =   "fPrintFont.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox chkSave 
      Alignment       =   1  'Rechts ausgerichtet
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Save selected font properties in registry . . ."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   525
      TabIndex        =   4
      ToolTipText     =   "Registry-Key is HKEY_CURRENT_USER\Software\VB and VBA Program Settings\Code Formatter\Print"
      Top             =   1725
      Value           =   1  'Aktiviert
      Width           =   3345
   End
   Begin VB.CommandButton btCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2272
      TabIndex        =   3
      Top             =   1125
      Width           =   1605
   End
   Begin VB.CommandButton btSelect 
      Caption         =   "S&elect Printer-Font"
      Default         =   -1  'True
      Height          =   495
      Left            =   577
      TabIndex        =   2
      Top             =   1125
      Width           =   1605
   End
   Begin MSComDlg.CommonDialog cdlFont 
      Left            =   1980
      Top             =   585
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbComplain 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   645
      Left            =   112
      TabIndex        =   1
      Top             =   375
      Width           =   4230
   End
   Begin VB.Label lbUser 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   4230
   End
End
Attribute VB_Name = "fPrintFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z 'we're 32 bits

Private Tooltips    As Collection

Private Sub btCancel_Click()

    Hide

End Sub

Private Sub btSelect_Click()

    On Error Resume Next
        With cdlFont
            .Min = 6
            .Max = 14
            .Flags = cdlCFPrinterFonts Or _
                     cdlCFFixedPitchOnly Or _
                     cdlCFScalableOnly Or _
                     cdlCFLimitSize
            .ShowFont
        End With 'CDLFONT
    On Error GoTo 0
    Hide

End Sub

Private Sub Form_Load()

    lbUser = "Sorry " & UserName & ","
    lbComplain = "your printer " & Printer.DeviceName & " on " & Printer.Port & " does not support your IDE-font " & MyFontName & ", or " & MyFontName & " is not a fixed font."
    Set Tooltips = CreateTooltips(Me)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Cancel = (UnloadMode = vbFormControlMenu)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set Tooltips = Nothing

End Sub

Public Property Get PrintFontName() As String

    PrintFontName = Trim$(cdlFont.Fontname)

End Property

Public Property Get PrintFontSize() As Long

    PrintFontSize = cdlFont.Fontsize

End Property

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 4  Code: 58  Total: 62 Lines
':) CommentOnly: 2 (3,2%)  Commented: 2 (3,2%)  Filled: 39 (62,9%)  Empty: 23 (37,1%)  Max Logic Depth: 2
