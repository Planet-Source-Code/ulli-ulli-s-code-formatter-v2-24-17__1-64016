VERSION 5.00
Begin VB.Form fSplash 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4620
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   Icon            =   "fSplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Formatter Add-In..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1230
      TabIndex        =   0
      Top             =   450
      Width           =   2820
   End
   Begin VB.Image imgUMG 
      BorderStyle     =   1  'Fest Einfach
      Height          =   765
      Left            =   195
      Top             =   188
      Width           =   825
   End
End
Attribute VB_Name = "fSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form has no code

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:26)  Decl: 4  Code: 0  Total: 4 Lines
':) CommentOnly: 3 (75%)  Commented: 0 (0%)  Filled: 3 (75%)  Empty: 1 (25%)  Max Logic Depth: 0
