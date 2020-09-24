VERSION 5.00
Begin VB.UserControl ctlFader 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   630
   ForwardFocus    =   -1  'True
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   330
   ScaleWidth      =   630
   ToolboxBitmap   =   "ctlFader.ctx":0000
   Windowless      =   -1  'True
   Begin VB.Label lbName 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Name"
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
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   555
   End
End
Attribute VB_Name = "ctlFader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'This code is based on a submission to PSC by Ed Preston

Public Enum FadingSpeed
    FadeSlow = 1
    FadeMedium = 2
    FadeFast = 4
    FadeVeryFast = 8
End Enum
#If False Then
Private FadeSlow, FadeMedium, FadeFast, FadeVeryFast
#End If

'Properties
Private Const pnEnabled         As String = "Enabled"
Private Const pnFadeIn          As String = "FadeInSpeed"
Private Const pnFadeOut         As String = "FadeOutSpeed"
Private Const pnOpacity         As String = "Opacity"
Private myEnabled               As Boolean
Private myFadeInSpeed           As FadingSpeed
Private myFadeOutSpeed          As FadingSpeed
Private myOpacity               As Long

'Private variables
Private Alpha                   As Long
Private hWndParent              As Long
Private Internal                As Boolean

'Events
Public Event FadeInReady()
Public Event FadeOutReady()

'Win API
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Function OSVersion Lib "kernel32" Alias "GetVersion" () As Long
Private Const RequiredVersion   As Long = 5

'Win Consts
Private Const WS_EX_LAYERED     As Long = &H80000
Private Const GWL_EXSTYLE       As Long = -20
Private Const LWA_ALPHA         As Long = 2

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Sets/returns whether the Control is operable."

    Enabled = myEnabled

End Property

Public Property Let Enabled(ByVal nwEnabled As Boolean)

    myEnabled = (nwEnabled <> False) And WindowsIsSuitable
    PropertyChanged pnEnabled

End Property

Public Sub FadeIn(Optional ByVal Speed As FadingSpeed = 0)

    If myEnabled Then
        If Speed = 0 Then
            Speed = myFadeInSpeed
        End If
        Do
            DoEvents
            SetLayeredWindowAttributes hWndParent, 0, Alpha, LWA_ALPHA
            Sleep 1
            Alpha = Alpha + Speed
        Loop Until Alpha > (myOpacity / 100) * 255
        Alpha = Alpha - Speed
        If Not Internal Then
            RaiseEvent FadeInReady
        End If
      Else 'MYENABLED = FALSE/0
        If Not Internal Then
            SetLayeredWindowAttributes hWndParent, 0, 255, LWA_ALPHA
        End If
    End If

End Sub

Public Property Get FadeInSpeed() As FadingSpeed

    FadeInSpeed = myFadeInSpeed

End Property

Public Property Let FadeInSpeed(ByVal nwFadeInSpeed As FadingSpeed)

    Select Case nwFadeInSpeed
      Case FadeVeryFast, FadeFast, FadeMedium, FadeSlow
        myFadeInSpeed = nwFadeInSpeed
        PropertyChanged pnFadeIn
      Case Else 'NOT NWFADEINSPEED...
        Err.Raise 380
    End Select

End Property

Public Sub FadeOut(Optional ByVal Speed As FadingSpeed = 0)

    If myEnabled Then
        If Speed = 0 Then
            Speed = myFadeOutSpeed
        End If
        Do
            DoEvents
            SetLayeredWindowAttributes hWndParent, 0, Alpha, LWA_ALPHA
            Sleep 1
            Alpha = Alpha - Speed
        Loop Until Alpha < IIf(Internal, (myOpacity / 100) * 255, 0)
        Alpha = Alpha + Speed
        If Not Internal Then
            RaiseEvent FadeOutReady
        End If
      Else 'MYENABLED = FALSE/0
        If Not Internal Then
            SetLayeredWindowAttributes hWndParent, 0, 0, LWA_ALPHA
        End If
    End If

End Sub

Public Property Get FadeOutSpeed() As FadingSpeed

    FadeOutSpeed = myFadeOutSpeed

End Property

Public Property Let FadeOutSpeed(ByVal nwFadeOutSpeed As FadingSpeed)

    Select Case nwFadeOutSpeed
      Case FadeVeryFast, FadeFast, FadeMedium, FadeSlow
        myFadeOutSpeed = nwFadeOutSpeed
        PropertyChanged pnFadeOut
      Case Else 'NOT NWFADEOUTSPEED...
        Err.Raise 380
    End Select

End Property

Public Property Let Opacity(ByVal nwOpacity As Long)
Attribute Opacity.VB_Description = "Percent value of opacity."

  Dim PreviousOpacity   As Long

    PreviousOpacity = myOpacity
    If nwOpacity >= 25 And nwOpacity <= 100 Then
        myOpacity = nwOpacity
        PropertyChanged pnOpacity
        If Ambient.UserMode Then
            Internal = True
            If myOpacity > PreviousOpacity Then
                FadeIn
                RaiseEvent FadeInReady
              ElseIf myOpacity < PreviousOpacity Then 'NOT MYOPACITY...
                FadeOut
                RaiseEvent FadeOutReady
              Else 'NOT MYOPACITY...
                RaiseEvent FadeOutReady
            End If
            Internal = False 'NOT RIGHT$(STMP,...
        End If
      Else 'NOT NWOPACITY...
        Err.Raise 380
    End If

End Property

Public Property Get Opacity() As Long

    Opacity = myOpacity

End Property

Private Sub UserControl_InitProperties()

    myFadeInSpeed = FadeMedium
    myFadeOutSpeed = FadeMedium
    myEnabled = WindowsIsSuitable
    myOpacity = 100

End Sub

Private Sub UserControl_Paint()

    lbName = Ambient.DisplayName
    UserControl_Resize

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        myEnabled = .ReadProperty(pnEnabled, True) And WindowsIsSuitable
        myFadeInSpeed = .ReadProperty(pnFadeIn, FadeMedium)
        myFadeOutSpeed = .ReadProperty(pnFadeOut, FadeMedium)
        myOpacity = .ReadProperty(pnOpacity, 100)
    End With 'PROPBAG

    If Ambient.UserMode Then
        hWndParent = Parent.hWnd
        If WindowsIsSuitable Then
            SetWindowLong hWndParent, GWL_EXSTYLE, GetWindowLong(hWndParent, GWL_EXSTYLE) Or WS_EX_LAYERED
            SetLayeredWindowAttributes hWndParent, 0, 0, LWA_ALPHA
        End If
        Alpha = 1
    End If

End Sub

Private Sub UserControl_Resize()

    Size lbName.Width, lbName.Height

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty pnEnabled, myEnabled, WindowsIsSuitable
        .WriteProperty pnFadeIn, myFadeInSpeed, FadeMedium
        .WriteProperty pnFadeOut, myFadeOutSpeed, FadeMedium
        .WriteProperty pnOpacity, myOpacity, 100
    End With 'PROPBAG

End Sub

Private Function WindowsIsSuitable() As Boolean

    WindowsIsSuitable = ((OSVersion And &HFF&) >= RequiredVersion)

    'uncoment next line for experiments with other Windows'es
    'WindowsIsSuitable = True

End Function

':) Ulli's VB Code Formatter V2.24.11 (2008-Apr-11 10:25)  Decl: 45  Code: 195  Total: 240 Lines
':) CommentOnly: 10 (4,2%)  Commented: 10 (4,2%)  Filled: 180 (75%)  Empty: 60 (25%)  Max Logic Depth: 4
