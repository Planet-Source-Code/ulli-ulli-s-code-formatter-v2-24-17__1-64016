VERSION 5.00
Begin VB.Form fQuestion 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   7860
   ClientLeft      =   4095
   ClientTop       =   4755
   ClientWidth     =   4095
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   Icon            =   "fQuestion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   524
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picContainer 
      BackColor       =   &H00FFFFF0&
      ForeColor       =   &H00000000&
      Height          =   3840
      Left            =   2220
      ScaleHeight     =   3780
      ScaleWidth      =   1530
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1410
      Width           =   1590
      Begin VB.CheckBox ckEmptyLines 
         BackColor       =   &H00FFFFF0&
         Caption         =   "C&ompress"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   30
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "Remove all empty lines"
         Top             =   3555
         Width           =   1260
      End
      Begin VB.CheckBox ckLinNum 
         BackColor       =   &H00FFFFF0&
         Caption         =   "Remo&ve LN's"
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   30
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Remove Line Numbers"
         Top             =   3285
         Width           =   1260
      End
      Begin VB.CheckBox ckCall 
         BackColor       =   &H00FFFFF0&
         Caption         =   "Ca&ll Statements"
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   30
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "Remove Call verbs"
         Top             =   2475
         Width           =   1410
      End
      Begin VB.CheckBox ckIfExp 
         BackColor       =   &H00FFFFF0&
         Caption         =   "If E&xpansion"
         ForeColor       =   &H00808000&
         Height          =   210
         Left            =   30
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Expand single line  If's (right click for silent expansion)"
         Top             =   1140
         Width           =   1185
      End
      Begin VB.CheckBox ckEnum 
         BackColor       =   &H00FFFFF0&
         Caption         =   "E&num Case"
         ForeColor       =   &H00008080&
         Height          =   210
         Left            =   30
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "Insert code to preserve Enum member capitalization"
         Top             =   3015
         Width           =   1125
      End
      Begin VB.CheckBox ckStop 
         BackColor       =   &H00FFFFF0&
         Caption         =   "Hal&t on Mark"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   30
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Stop scan at mark insertion"
         Top             =   2745
         Width           =   1230
      End
      Begin VB.CheckBox ckTypeSuff 
         BackColor       =   &H00FFFFF0&
         Caption         =   "Type S&uffixes"
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   30
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Substitute type suffixes by As [type]"
         Top             =   2205
         Width           =   1290
      End
      Begin VB.CheckBox ckMark 
         BackColor       =   &H00FFFFF0&
         Caption         =   "Insert Co&mments"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   30
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Insert comments to mark offending lines"
         Top             =   328
         Value           =   1  'Aktiviert
         Width           =   1485
      End
      Begin VB.CheckBox ckStruc 
         BackColor       =   &H00FFFFF0&
         Caption         =   "Cr&eate Structure"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   30
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Create code structure tree"
         Top             =   1665
         Width           =   1470
      End
      Begin VB.CheckBox ckSumma 
         BackColor       =   &H00FFFFF0&
         Caption         =   "Insert Summar&y"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   30
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Insert summary into code"
         Top             =   60
         Value           =   1  'Aktiviert
         Width           =   1395
      End
      Begin VB.CheckBox ckSep 
         BackColor       =   &H00FFFFF0&
         Caption         =   "&Break Multiple"
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   30
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "...into one statement per line"
         Top             =   596
         Value           =   1  'Aktiviert
         Width           =   1320
      End
      Begin VB.CheckBox ckSort 
         BackColor       =   &H00FFFFF0&
         Caption         =   "Sort Mo&dules"
         ForeColor       =   &H00C000C0&
         Height          =   210
         Left            =   30
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Sort Subs and Functions alphabetically"
         Top             =   1935
         Width           =   1245
      End
      Begin VB.CheckBox ckStrFncts 
         BackColor       =   &H00FFFFF0&
         Caption         =   "String &Functions"
         ForeColor       =   &H00808000&
         Height          =   210
         Left            =   30
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Modifies variant string functions; {LR} click for selection"
         Top             =   1395
         Width           =   1455
      End
      Begin VB.CheckBox ckHalfIndent 
         BackColor       =   &H00FFFFF0&
         Caption         =   "Half &Indent"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   30
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Enable half indenting"
         Top             =   864
         Value           =   1  'Aktiviert
         Width           =   1095
      End
   End
   Begin VB.CheckBox ckCompile 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Compile project after formatting"
      ForeColor       =   &H00006000&
      Height          =   225
      Left            =   165
      TabIndex        =   37
      Top             =   6885
      Width           =   2505
   End
   Begin VB.CommandButton btShowAll 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Sho&w All"
      Height          =   285
      Left            =   870
      Style           =   1  'Grafisch
      TabIndex        =   36
      ToolTipText     =   "Open all code panels"
      Top             =   2265
      Width           =   1245
   End
   Begin VB.CommandButton btCoAb 
      Cancel          =   -1  'True
      Height          =   465
      Index           =   2
      Left            =   1500
      Picture         =   "fQuestion.frx":030A
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Undo last formatting"
      Top             =   7200
      UseMaskColor    =   -1  'True
      Width           =   1065
   End
   Begin VB.CheckBox ckWinXPLook 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Create Win &XP Look"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   225
      Left            =   960
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Create manifest file and insert approriate code"
      Top             =   4680
      Width           =   2070
   End
   Begin VB.CheckBox ckBook 
      Alignment       =   1  'Rechts ausgerichtet
      BackColor       =   &H00E0E0E0&
      Caption         =   " Boo&k"
      Enabled         =   0   'False
      ForeColor       =   &H00800080&
      Height          =   225
      Left            =   3000
      TabIndex        =   23
      ToolTipText     =   "Print alternate pages"
      Top             =   6270
      Width           =   735
   End
   Begin VB.CheckBox ckPrint 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Left            =   960
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Print Source File"
      Top             =   5520
      Width           =   720
   End
   Begin VB.ComboBox cboStopWhen 
      ForeColor       =   &H00C00000&
      Height          =   315
      ItemData        =   "fQuestion.frx":0490
      Left            =   870
      List            =   "fQuestion.frx":049D
      Style           =   2  'Dropdown-Liste
      TabIndex        =   6
      ToolTipText     =   "Select Condition"
      Top             =   1815
      Width           =   1260
   End
   Begin VB.CommandButton btSendMeMail 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      Style           =   1  'Grafisch
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "Send Mail to Author"
      Top             =   0
      Width           =   285
   End
   Begin VB.CommandButton btHelp 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3690
      Style           =   1  'Grafisch
      TabIndex        =   26
      TabStop         =   0   'False
      ToolTipText     =   "Show Help"
      Top             =   0
      Width           =   195
   End
   Begin VB.CommandButton btAbout 
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3900
      Style           =   1  'Grafisch
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Show About Box"
      Top             =   0
      Width           =   195
   End
   Begin VB.CheckBox ckAll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "De/&Select all"
      ForeColor       =   &H000000D0&
      Height          =   225
      Left            =   915
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Select/Unselect all Components"
      Top             =   2640
      Width           =   1245
   End
   Begin VB.ListBox lstModNames 
      BackColor       =   &H00E0F8FF&
      ForeColor       =   &H000000C0&
      Height          =   1185
      Left            =   870
      Style           =   1  'Kontrollkästchen
      TabIndex        =   16
      ToolTipText     =   "Select Components"
      Top             =   3045
      Width           =   2925
   End
   Begin VB.CommandButton btCoAb 
      Height          =   465
      Index           =   1
      Left            =   2850
      Picture         =   "fQuestion.frx":04BE
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Exit Code Formatter"
      Top             =   7200
      UseMaskColor    =   -1  'True
      Width           =   1065
   End
   Begin VB.CommandButton btCoAb 
      Height          =   465
      Index           =   0
      Left            =   150
      Picture         =   "fQuestion.frx":06E0
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Format selected Components"
      Top             =   7200
      UseMaskColor    =   -1  'True
      Width           =   1065
   End
   Begin VB.OptionButton opPor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Portr&ait"
      Enabled         =   0   'False
      ForeColor       =   &H00800080&
      Height          =   225
      Left            =   945
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Long edge of paper vertical"
      Top             =   5790
      Width           =   840
   End
   Begin VB.OptionButton opLand 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Landsc&ape"
      Enabled         =   0   'False
      ForeColor       =   &H00800080&
      Height          =   225
      Left            =   945
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Short edge of paper vertical"
      Top             =   6030
      Width           =   1125
   End
   Begin VB.CheckBox ckColor 
      Alignment       =   1  'Rechts ausgerichtet
      BackColor       =   &H00E0E0E0&
      Caption         =   " in &Color"
      Enabled         =   0   'False
      ForeColor       =   &H00800080&
      Height          =   225
      Left            =   2850
      TabIndex        =   21
      ToolTipText     =   "Print in Color"
      Top             =   5730
      Width           =   885
   End
   Begin VB.CheckBox ckStat 
      Alignment       =   1  'Rechts ausgerichtet
      BackColor       =   &H00E0E0E0&
      Caption         =   " Stati&onary"
      Enabled         =   0   'False
      ForeColor       =   &H00800080&
      Height          =   225
      Left            =   2670
      TabIndex        =   22
      ToolTipText     =   "Create preprinted stationary"
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Image imgComp 
      Height          =   225
      Left            =   2625
      Picture         =   "fQuestion.frx":0902
      Top             =   6885
      Width           =   1275
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "(Windows XP required)"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   165
      Index           =   6
      Left            =   1245
      TabIndex        =   35
      Top             =   4890
      Width           =   1395
   End
   Begin VB.Shape sh 
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   855
      Top             =   4590
      Width           =   2940
   End
   Begin VB.Label lbFont 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   225
      Left            =   3705
      TabIndex        =   33
      ToolTipText     =   "Print Font Sample"
      Top             =   5235
      Width           =   105
   End
   Begin VB.Label lblPages 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   165
      Left            =   1845
      TabIndex        =   32
      ToolTipText     =   "Rough estimate"
      Top             =   5535
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Label lbl 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Printing: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   5
      Left            =   345
      TabIndex        =   31
      ToolTipText     =   "Select Options below"
      Top             =   5205
      Width           =   990
   End
   Begin VB.Line ln 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Index           =   6
      X1              =   13
      X2              =   264
      Y1              =   357
      Y2              =   357
   End
   Begin VB.Line ln 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   5
      X1              =   11
      X2              =   262
      Y1              =   355
      Y2              =   355
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   2
      Left            =   195
      Picture         =   "fQuestion.frx":1844
      Top             =   5535
      Width           =   480
   End
   Begin VB.Label lbLL 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
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
      Height          =   225
      Left            =   2085
      TabIndex        =   30
      Top             =   6270
      Width           =   360
   End
   Begin VB.Image imLand 
      Height          =   180
      Left            =   2130
      Picture         =   "fQuestion.frx":2486
      ToolTipText     =   "Landscape"
      Top             =   5910
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imPor 
      Height          =   255
      Left            =   2160
      Picture         =   "fQuestion.frx":2738
      ToolTipText     =   "Portrait"
      Top             =   5865
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Shape sh 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FFFFFF&
      Height          =   945
      Index           =   0
      Left            =   855
      Top             =   5595
      Width           =   2940
   End
   Begin VB.Shape sh 
      BackColor       =   &H00000000&
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      Height          =   945
      Index           =   1
      Left            =   870
      Top             =   5610
      Width           =   2940
   End
   Begin VB.Line ln 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Index           =   4
      X1              =   11
      X2              =   262
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Line ln 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   3
      X1              =   9
      X2              =   260
      Y1              =   448
      Y2              =   448
   End
   Begin VB.Label lbl 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "S&how Summary Box:"
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   2
      Left            =   885
      TabIndex        =   5
      ToolTipText     =   "Show Summary Box"
      Top             =   1380
      Width           =   1005
   End
   Begin VB.Label lblNumSel 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0F8FF&
      BorderStyle     =   1  'Fest Einfach
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   870
      TabIndex        =   17
      Top             =   4260
      Width           =   2925
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   1
      Left            =   195
      Picture         =   "fQuestion.frx":29DE
      ToolTipText     =   "Attention"
      Top             =   345
      Width           =   480
   End
   Begin VB.Line ln 
      BorderColor     =   &H80000002&
      Index           =   0
      X1              =   1
      X2              =   276
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblBar 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   210
      Left            =   15
      TabIndex        =   29
      Top             =   15
      Width           =   4140
   End
   Begin VB.Label lbl 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "ò"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   1
      Left            =   960
      TabIndex        =   28
      Top             =   2850
      Width           =   135
   End
   Begin VB.Label lbl 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Your  Code  must be syntactically correct and  may also  have Line Numbers. Lines above 800 chars will not be formattet."
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
      Height          =   780
      Index           =   0
      Left            =   930
      MouseIcon       =   "fQuestion.frx":3620
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   1
      Top             =   315
      Width           =   2850
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   195
      Picture         =   "fQuestion.frx":392A
      ToolTipText     =   "Select Options"
      Top             =   1515
      Width           =   480
   End
   Begin VB.Label lbl 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Options: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   3
      Left            =   345
      TabIndex        =   4
      ToolTipText     =   "Select Options below"
      Top             =   1110
      Width           =   990
   End
   Begin VB.Line ln 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   2
      X1              =   11
      X2              =   262
      Y1              =   82
      Y2              =   82
   End
   Begin VB.Line ln 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Index           =   1
      X1              =   13
      X2              =   264
      Y1              =   83
      Y2              =   84
   End
   Begin VB.Label lbl 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Line Length will be             Chars"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00606060&
      Height          =   165
      Index           =   4
      Left            =   960
      TabIndex        =   24
      Top             =   6300
      Width           =   1860
   End
   Begin VB.Shape sh 
      BackColor       =   &H00808080&
      BorderColor     =   &H00808080&
      Height          =   495
      Index           =   3
      Left            =   870
      Top             =   4605
      Width           =   2940
   End
End
Attribute VB_Name = "fQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z 'we're 32 bit

'This shows how to fake a Caption Bar

'Hi resolution timer
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private HiResTimerPresent       As Boolean
Private CPUTicksNow             As Currency
Private CPUTicksDone            As Currency
Private WaitCycles              As Currency

Public Enum CoAb 'Indexes of btCoAb
    Continue = 0
    Abandon = 1
    Undo = 2
End Enum
#If False Then
Private Continue, Abandon, Undo
#End If

Public Reply            As CoAb
Private CodeModules     As VBIDE.CodePanes 'the open codepanes
Private Tooltips        As Collection 'keeping references to all tooltip class instances, killed on fade out
Private NumToText       As Variant
Private TopIndex
Private m
Private n
Private NumLines
Private NumPages
Private LinesPerPage
Private Words()         As String
Private MeActive        As Boolean
Private Internal        As Boolean
Private NeedsRepaint    As Boolean
Private KeyIsDown       As Boolean
Private PrevSep         As Integer

Private Const CollapsedHeight   As Long = 96 'picContainer
Private Const ExpandedHeight    As Long = 256

Private Sub ActivateBar()

    lblBar.BackColor = vbActiveTitleBar
    ln(0).BorderColor = vbActiveTitleBar
    lblBar.ForeColor = vbActiveTitleBarText
    MeActive = True

End Sub

Private Sub btAbout_Click()

    DeactivateBar

    With frmAbout
        .Theme = 20
        .AppIcon(&HC8C8C8) = fQuestion.Icon
        .Title(vbRed) = "Ulli's Code Formatter"
        .Version(vbYellow) = "Version " & App.Major & "." & App.Minor & "." & App.Revision
        .Copyright(vbYellow) = App.LegalCopyright
        .Otherstuff1(vbCyan) = "The Ultimate Code Formatter"
        .Otherstuff2(vbWhite) = "Features Code Formatter, Module Sorting, Splitting of Compound Lines and Single Line If's, and much more (See Help) "
        .Show vbModal, Me
    End With 'FRMABOUT

    ActivateBar

End Sub

Private Sub btAbout_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If MeActive Then
        MPIcon.SetPointerIcon RightHand
        btAbout.BackColor = vbCyan
        btHelp.BackColor = vbButtonFace
        NeedsRepaint = True
    End If

End Sub

Private Sub btCoAb_Click(Index As Integer)

    StrucRequested = (ckStruc = vbChecked)
    TypeSuffRequested = (ckTypeSuff = vbChecked)
    StrFnctsRequested = (ckStrFncts = vbChecked)
    StopRequested = (ckStop = vbChecked)
    PauseAfterScan = cboStopWhen.ListIndex
    InsertComments = (ckSumma.Value = vbChecked)
    If Index = Continue Then 'Continue button
        StoreSettings ckStrFncts.Value, CInt(PauseAfterScan), ckSumma.Value, ckMark.Value, ckSep.Value Or PrevSep, ckHalfIndent.Value
      Else 'NOT INDEX...
        ckCompile = vbUnchecked 'no compilation after Undo or Cancel
        ckPrint = vbUnchecked 'no printing either
    End If
    PrevSep = vbUnchecked
    SortRequested = (ckSort.Value = vbChecked)
    ckSort = vbUnchecked
    XPLookRequested = (ckWinXPLook.Value = vbChecked) And 3 'bit 1 = decl req, 2 = call req, 4 = in curr proc
    ckWinXPLook.Value = vbUnchecked
    Reply = Index
    Hide
    DoEvents

End Sub

Private Sub btCoAb_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    MPIcon.SetPointerIcon RightHand

End Sub

Private Sub btHelp_Click()

    DeactivateBar
    fHelp.Show vbModal    'fHelp unloads itself
    ActivateBar

End Sub

Private Sub btHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If MeActive Then
        MPIcon.SetPointerIcon RightHand
        btHelp.BackColor = vbCyan
        btAbout.BackColor = vbButtonFace
        NeedsRepaint = True
    End If

End Sub

Private Sub btSendMeMail_Click()

    SendMeMail hWnd, AppDetails

End Sub

Private Sub btSendMeMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If MeActive Then
        MPIcon.SetPointerIcon RightHand
        btSendMeMail.BackColor = vbCyan
        NeedsRepaint = True
    End If

End Sub

Private Sub btShowAll_Click()

    ResetFocus
    MPIcon = vbHourglass
    With VBInstance
        For Each Proj In .VBProjects
            With Proj
                For Each Compo In .VBComponents
                    With Compo
                        If .Type <> vbext_ct_ResFile And .Type <> vbext_ct_RelatedDocument Then
                            .CodeModule.CodePane.Show
                            DoEvents
                        End If
                    End With 'COMPO
                Next Compo
            End With 'PROJ
        Next Proj
        btShowAll.Enabled = (GetModuleCount <> .CodePanes.Count)
        lstModNames.Clear
        LoadListbox
        ckAll = vbUnchecked
    End With 'VBINSTANCE
    MPIcon = vbDefault

End Sub

Private Sub btShowAll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MPIcon.SetPointerIcon RightHand

End Sub

Private Sub cboStopWhen_Click()

    If cboStopWhen.ListIndex <> Always Then
        Internal = True
        ckStruc = vbUnchecked
        Internal = False
    End If
    If Not KeyIsDown Then
        ResetFocus
    End If

End Sub

Private Sub ckAll_Click()

  Dim li

    With lstModNames
        TopIndex = SendMessage(.hWnd, LB_GETTOPINDEX, 0&, ByVal 0&)
        For li = 0 To .ListCount - 1
            .Selected(li) = (ckAll = vbChecked)
        Next li
        SendMessage .hWnd, LB_SETTOPINDEX, TopIndex, ByVal 0&
    End With 'LSTMODNAMES
    If ckAll = vbChecked Then
        cboStopWhen.ListIndex = IfNecessary
      Else 'NOT CKALL...
        cboStopWhen.ListIndex = PaS
    End If
    ResetFocus
    UpdatePages (ckPrint = vbChecked)

End Sub

Private Sub ckBook_Click()

    ResetFocus
    BookRequested = (ckBook = vbChecked)

End Sub

Private Sub ckCall_Click()

    ResetFocus

End Sub

Private Sub ckCall_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picContainer_MouseMove 0, 0, 0, 0

End Sub

Private Sub ckColor_Click()

    ResetFocus
    ColorRequested = (ckColor = vbChecked)

End Sub

Private Sub ckCompile_Click()

    ResetFocus

End Sub

Private Sub ckEmptyLines_Click()

    ResetFocus

End Sub

Private Sub ckEnum_Click()

    ResetFocus

End Sub

Private Sub ckEnum_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picContainer_MouseMove 0, 0, 0, 0

End Sub

Private Sub ckHalfIndent_Click()

    ResetFocus

End Sub

Private Sub ckHalfIndent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picContainer_MouseMove 0, 0, 0, 0

End Sub

Private Sub ckIfExp_Click()

    If ckIfExp = vbChecked Then
        PrevSep = ckSep 'save state for next time
        ckSep = vbUnchecked
      Else 'NOT CKIFEXP...
        ckSep = PrevSep
    End If
    ResetFocus

End Sub

Private Sub ckIfExp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picContainer_MouseMove 0, 0, 0, 0

End Sub

Private Sub ckIfExp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        ckIfExp = IIf(ckIfExp = vbChecked, vbUnchecked, vbChecked)
        ckIfExp.Tag = "1"
      Else 'NOT BUTTON...
        ckIfExp.Tag = "0"
    End If

End Sub

Private Sub ckMark_Click()

    If ckMark = vbUnchecked Then
        ckStop = vbUnchecked
        ckStop.Enabled = False
        ckSumma = vbUnchecked
      Else 'NOT CKMARK...
        ckStop.Enabled = True
    End If
    ResetFocus

End Sub

Private Sub ckMark_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picContainer_MouseMove 0, 0, 0, 0

End Sub

Private Sub ckPrint_Click()

    opPor.Enabled = (ckPrint = vbChecked)
    opLand.Enabled = (ckPrint = vbChecked)
    ckColor.Enabled = (ckPrint = vbChecked)
    ckStat.Enabled = (ckPrint = vbChecked)
    ckBook.Enabled = (ckPrint = vbChecked)
    ckColor = IIf(ColorRequested, vbChecked, vbUnchecked)
    ckStat = IIf(WithStationary, vbChecked, vbUnchecked)
    ckBook = IIf(BookRequested, vbChecked, vbUnchecked)
    opPor = (Printer.Orientation = vbPRORPortrait)
    opLand = (Printer.Orientation = vbPRORLandscape)
    If ckPrint = vbChecked And NumSelected Then
        ckStruc = vbUnchecked
        btCoAb(Continue).ToolTipText = "Format and Print selected Components"
        UpdateTooltip Tooltips, btCoAb(Continue)
      Else 'NOT CKPRINT...
        PrintLineLen = 0
        lbLL = NullStr
        UpdatePages False
        opPor = False
        opLand = False
        imPor.Visible = False
        imLand.Visible = False
        PBOdd.Top = 0
        PBEven.Top = 0
        btCoAb(Continue).ToolTipText = "Format selected Components"
        UpdateTooltip Tooltips, btCoAb(Continue)
        ckPrint = vbUnchecked 'this causes a recursion if ckPrint was vbChecked before
    End If
    ResetFocus

End Sub

Private Sub ckSep_Click()

    ResetFocus
    If ckSep = vbChecked Then
        ckIfExp = vbUnchecked
    End If

End Sub

Private Sub ckSep_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picContainer_MouseMove 0, 0, 0, 0

End Sub

Private Sub ckSort_Click()

    ResetFocus

End Sub

Private Sub ckSort_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picContainer_MouseMove 0, 0, 0, 0

End Sub

Private Sub ckStat_Click()

    ResetFocus
    WithStationary = (ckStat = vbChecked)

End Sub

Private Sub ckStop_Click()

    ResetFocus

End Sub

Private Sub ckStrFncts_Click()

    If fStrFncts.AnySelected = False Then
        ckStrFncts = vbUnchecked
    End If
    ResetFocus

End Sub

Private Sub ckStrFncts_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then
        fStrFncts.Show vbModal
        ckStrFncts = IIf(fStrFncts.AnySelected, vbChecked, vbUnchecked)
    End If

End Sub

Private Sub ckStrFncts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picContainer_MouseMove 0, 0, 0, 0

End Sub

Private Sub ckStruc_Click()

    If NumSelected = 0 Or NumSelected > 1 Then
        ckStruc = vbUnchecked
    End If
    If ckStruc = vbChecked Then
        cboStopWhen.ListIndex = Always
        ckPrint = vbUnchecked
        btCoAb(Continue).ToolTipText = "Format selected Components and Create Structure"
        UpdateTooltip Tooltips, btCoAb(Continue)
      Else 'NOT CKSTRUC...
        If Not Internal Then
            cboStopWhen.ListIndex = PaS
        End If
        btCoAb(Continue).ToolTipText = "Format selected Components"
        UpdateTooltip Tooltips, btCoAb(Continue)
    End If
    ResetFocus

End Sub

Private Sub ckStruc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picContainer_MouseMove 0, 0, 0, 0

End Sub

Private Sub ckSumma_Click()

    If ckSumma = vbChecked Then
        ckMark = vbChecked
    End If
    ResetFocus

End Sub

Private Sub ckSumma_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picContainer_MouseMove 0, 0, 0, 0

End Sub

Private Sub ckTypeSuff_Click()

    ResetFocus

End Sub

Private Sub ckTypeSuff_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picContainer_MouseMove 0, 0, 0, 0

End Sub

Private Sub ckWinXPLook_Click()

    If ckWinXPLook = vbChecked Then
        ckAll = vbChecked
        ckAll_Click
        ckWinXPLook = vbChecked
    End If

End Sub

Private Sub CollapseOpts()

    If Not KeyIsDown Then
        With picContainer
            For n = .Height To CollapsedHeight Step -1
                .Height = n
                lblNumSel.Refresh
                Wait
            Next n
        End With 'PICCONTAINER
    End If

End Sub

Private Sub DeactivateBar()

    lblBar.BackColor = vbInactiveTitleBar
    ln(0).BorderColor = vbInactiveTitleBar
    lblBar.ForeColor = vbInactiveTitleBarText
    btAbout.BackColor = vbButtonFace
    btHelp.BackColor = vbButtonFace
    btSendMeMail.BackColor = vbButtonFace
    ResetFocus
    MeActive = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyMenu And Not KeyIsDown Then 'Alt Key
        picContainer_MouseMove 0, 0, 0, 0
    End If
    KeyIsDown = True

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Select Case Chr$(KeyAscii)
      Case "C", "c", Chr$(vbKeyReturn)
        btCoAb_Click Continue
      Case "A", "a", Chr$(vbKeyEscape)
        btCoAb_Click Abandon
      Case "U", "u"
        btCoAb_Click Undo
      Case "M", "m"
        btSendMeMail_Click
      Case "I", "i"
        btAbout_Click
      Case "?"
        btHelp_Click
      Case Else
        MessageBeep vbCritical
    End Select

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyIsDown = False
    If KeyCode = vbKeyMenu Then 'Alt Key
        CollapseOpts
    End If

End Sub

Private Sub Form_Load()

  Dim p1 As Long
  Dim p2 As Long

    AddShadow Me
    lblBar = Space$(6) & AppDetails
    Copyright = App.LegalCopyright
    p1 = InStr(Copyright, Spce)
    p2 = InStr(p1 + 1, Copyright, Spce)
    p2 = InStr(p2 + 1, Copyright, Spce)
    p2 = InStr(p2 + 1, Copyright, Spce)
    Copyright = Mid$(Copyright, p1 + 1, p2 - p1 - 1)
    ActivateBar
    HiResTimerPresent = False
    On Error Resume Next
        HiResTimerPresent = CBool(QueryPerformanceFrequency(WaitCycles))
        WaitCycles = WaitCycles / 666
    On Error GoTo 0
    With ckStrFncts
        .Value = StF
        If MouseButtonsSwapped Then
            .ToolTipText = Replace$(.ToolTipText, "{LR}", "left")
          Else 'MOUSEBUTTONSSWAPPED = FALSE/0
            .ToolTipText = Replace$(.ToolTipText, "{LR}", "right")
        End If
    End With 'CKSTRFNCTS
    fStrFncts.LoadStrFncts
    Unload fStrFncts
    StringCboListIndex = 2
    cboStopWhen.ListIndex = PaS
    ckSumma.Value = Isc
    ckMark.Value = InM
    ckSep.Value = Sep
    ckHalfIndent.Value = HfI
    ckPrint.Enabled = PrintingOK
    If PrintingOK Then
        imgIcon(2).ToolTipText = Printer.DeviceName & " on " & Printer.Port
        sh(1).FillStyle = vbFSTransparent
        On Error Resume Next
            lbFont.Fontname = MyFontName
        On Error GoTo 0
        lbFont = Spce & MyFontName & Spce
        lbFont.Visible = True
      Else 'PRINTINGOK = FALSE/0
        sh(1).FillStyle = vbUpwardDiagonal
        lbFont.Visible = False
        imgIcon(2).ToolTipText = "No suitable " & IIf(HasPrinter, "font", "printer") & " available"
    End If
    NumToText = Array("No", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten")
    Set CodeModules = VBInstance.CodePanes
    picContainer.Height = CollapsedHeight
    Set Tooltips = CreateTooltips(Me)

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblBar_MouseDown Button, 0, 0, 0

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If NeedsRepaint Then
        btAbout.BackColor = vbButtonFace
        btHelp.BackColor = vbButtonFace
        btSendMeMail.BackColor = vbButtonFace
    End If
    CollapseOpts

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set CodeModules = Nothing
    Set Tooltips = Nothing

End Sub

Private Sub imLand_Click()

    opPor = True

End Sub

Private Sub imPor_Click()

    opLand = True

End Sub

Private Sub lblBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'Fake to move the window

    If Button = vbLeftButton Then
        ReleaseCapture 'release the Mouse
        SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0& 'non-client area button down (in caption)
    End If

End Sub

Private Sub lblBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseMove 0, 0, 0, 0

End Sub

Public Sub LoadListbox()

    With lstModNames
        m = 0
        n = 0
        For Each Pane In VBInstance.CodePanes
            Words = Split(Pane.Window.Caption)
            Inc n
            If Words(1) = "-" Then
                .AddItem n & ": " & vbTab & Words(2) & IIf(UBound(Words) < 4, NullStr, " {wp}") 'write protected
              Else 'NOT WORDS(1)...
                .AddItem n & ": " & vbTab & Words(0) & IIf(UBound(Words) < 2, NullStr, " {wp}")
            End If
            .ItemData(.NewIndex) = (InStr(.List(.NewIndex), "{wp}") = 0) 'not write protected
            If Pane Is VBInstance.ActiveCodePane Then
                m = .NewIndex
            End If
        Next Pane
        NumPanels = n
        .Selected(m) = .ItemData(m)
        'remove selection bar in listbox
        .ListIndex = -1
        'make selected item visible
        SendMessage .hWnd, LB_SETTOPINDEX, m, ByVal 0&
    End With 'LSTMODNAMES

End Sub

Private Sub lstModNames_Click()

  Dim NumText       As String

    With lstModNames
        .ListIndex = -1
        SendMessage .hWnd, LB_SETTOPINDEX, TopIndex, ByVal 0&
        NumSelected = .SelCount
        If NumSelected <= UBound(NumToText) Then
            NumText = NumToText(NumSelected)
          Else 'NOT NUMSELECTED...
            NumText = CStr(NumSelected)
        End If
        If NumSelected < 2 Then
            ckStruc.Enabled = (NumSelected = 1)
            btCoAb(0).Enabled = NumSelected
            btCoAb(2).Enabled = NumSelected
            cboStopWhen.ListIndex = PaS
          Else 'NOT NUMSELECTED...
            ckStruc = vbUnchecked
            ckStruc.Enabled = False
            cboStopWhen.ListIndex = IfNecessary
        End If
        Select Case NumSelected
          Case .ListCount
            ckAll = vbChecked
          Case 0
            ckAll = vbUnchecked
            ckWinXPLook = vbUnchecked
            ckPrint = vbUnchecked
          Case Else
            ckWinXPLook = vbUnchecked
        End Select
        If NumSelected = .ListCount Then
            lblNumSel = "All (" & NumText & OneOrMany(" Pane", NumSelected) & ") selected"
          Else 'NOT NUMSELECTED...
            lblNumSel = NumText & OneOrMany(" Pane", NumSelected) & " selected (of " & .ListCount & ")"
        End If
    End With 'LSTMODNAMES

End Sub

Private Sub lstModNames_ItemCheck(Item As Integer)

    lstModNames.Selected(Item) = lstModNames.ItemData(Item) And lstModNames.Selected(Item)

End Sub

Private Sub lstModNames_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    TopIndex = SendMessage(lstModNames.hWnd, LB_GETTOPINDEX, 0&, ByVal 0&)

End Sub

Private Sub lstModNames_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    CollapseOpts

End Sub

Private Sub lstModNames_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    UpdatePages (ckPrint = vbChecked)

End Sub

Private Sub opLand_Click()

  'landscape

    With Printer
        .Orientation = vbPRORLandscape
        .DrawWidth = 3
        imPor.Visible = False
        imLand.Visible = True
        PrintLineLen = Int(.ScaleWidth / PrintCharWidth) - LnLen - 1
        lbLL = PrintLineLen
        PBOdd.Left = .DrawWidth * .TwipsPerPixelX
        PBEven.Left = PBOdd.Left
        PBOdd.Right = .ScaleWidth - .DrawWidth * .TwipsPerPixelX
        PBEven.Right = PBOdd.Right
        PBOdd.Top = .ScaleHeight * 0.1
        PBEven.Top = .DrawWidth * .TwipsPerPixelY
        PBOdd.Bottom = .ScaleHeight - .DrawWidth * .TwipsPerPixelY
        PBEven.Bottom = .ScaleHeight * 0.9
        PBOdd.PunchX = .Width / 2
        PBEven.PunchX = PBOdd.PunchX
        PBOdd.PunchY = .ScaleTop
        PBEven.PunchY = .ScaleHeight - LenPunchMark * .TwipsPerPixelY
        UpdatePages True
    End With 'PRINTER

End Sub

Private Sub opLand_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetFocus

End Sub

Private Sub opPor_Click()

  'portrait

    With Printer
        .Orientation = vbPRORPortrait
        .DrawWidth = 3
        imLand.Visible = False
        imPor.Visible = True
        PBOdd.Left = .ScaleWidth * 0.1
        PBEven.Left = .DrawWidth * .TwipsPerPixelX
        PBOdd.Right = .ScaleWidth - .DrawWidth * .TwipsPerPixelX
        PBEven.Right = .ScaleWidth * 0.9
        PBOdd.Top = .DrawWidth * .TwipsPerPixelY
        PBEven.Top = PBOdd.Top
        PBOdd.Bottom = .ScaleHeight - .DrawWidth * .TwipsPerPixelY
        PBEven.Bottom = PBOdd.Bottom
        PBOdd.PunchX = .ScaleLeft
        PBEven.PunchX = .ScaleWidth - LenPunchMark * .TwipsPerPixelX
        PBOdd.PunchY = .Height / 2
        PBEven.PunchY = PBOdd.PunchY
        PrintLineLen = Int(PBEven.Right / PrintCharWidth) - LnLen - 1
        lbLL = PrintLineLen
        UpdatePages True
    End With 'PRINTER

End Sub

Private Sub opPor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ResetFocus

End Sub

Private Sub picContainer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With picContainer
        For n = .Height To ExpandedHeight
            .Height = n
            .Refresh
            Wait
        Next n
    End With 'PICCONTAINER

End Sub

Private Sub ResetFocus()

    On Error Resume Next
        btCoAb(Continue).SetFocus
    On Error GoTo 0

End Sub

Private Sub UpdatePages(Show As Boolean)

  'Estimate number of print pages (the factors are empirical values)

    If HasPrinter Then
        NumPages = 0
        LinesPerPage = Printer.ScaleHeight / PrintLineHeight * IIf(opLand, 0.75, 0.83)
        With lstModNames
            For n = 0 To .ListCount - 1
                If .Selected(n) Then
                    NumLines = CodeModules(n + 1).CodeModule.CountOfLines
                    Inc NumPages, 2
                    If NumLines > LinesPerPage Then
                        Inc NumPages, Int(NumLines / LinesPerPage)
                    End If
                End If
            Next n
        End With 'LSTMODNAMES
        lblPages.Visible = NumPages And Show
        lblPages = "About " & Int(NumPages * 1.2) + 1 & " Pages"
    End If

End Sub

Private Sub Wait()

  'high resolution timing function

    If HiResTimerPresent Then
        QueryPerformanceCounter CPUTicksDone
        CPUTicksDone = CPUTicksDone + WaitCycles
        Do
            QueryPerformanceCounter CPUTicksNow
        Loop Until CPUTicksNow > CPUTicksDone
    End If

End Sub

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 41  Code: 841  Total: 882 Lines
':) CommentOnly: 11 (1,2%)  Commented: 44 (5%)  Filled: 650 (73,7%)  Empty: 232 (26,3%)  Max Logic Depth: 7
